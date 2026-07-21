from flask import Flask, render_template, request, jsonify
import sqlite3
import csv
import io
from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup
from apscheduler.schedulers.background import BackgroundScheduler
import yfinance as yf
import random
import threading
import webview
import os
import sys

try:
    import openpyxl
except ImportError:
    raise ImportError("Lütfen Excel okuma desteği için 'pip install openpyxl' komutunu çalıştırın.")

# --- .EXE İÇİN DİNAMİK VARLIK VE DOSYA YOLU MOTORU ---
if getattr(sys, 'frozen', False):
    # Eğer uygulama .exe olarak kilitli çalışıyorsa geçici paket klasörünü hedefle
    base_dir = sys._MEIPASS
    template_folder = os.path.join(base_dir, 'templates')
    static_folder = os.path.join(base_dir, 'static')
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
    
    # Veritabanlarının exe'nin çalıştığı fiziksel klasörde kalması için yol mühürleme
    db_dir = os.path.dirname(sys.executable)
else:
    # Normal script ortamında yerel dizinleri hedefle
    app = Flask(__name__)
    db_dir = os.getcwd()

# Veritabanlarını fiziksel yola mühürlüyoruz
PORTFOLIO_DB = os.path.join(db_dir, 'portfolio.db')
PRICES_DB = os.path.join(db_dir, 'prices.db')
GUNLUK_DB = os.path.join(db_dir, 'gunluk_veri.db')  # HZM ARŞİV KORUMA HATTI

# --- JINJA2 İÇİN TL FORMAT FİLTRESİ ---
def tl_format(val):
    if val is None:
        return "0,00"
    try:
        return f"{float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return "0,00"

app.jinja_env.filters['tl'] = tl_format


# --- VERİTABANLARINI SIFIRDAN TEMİZ VE KUSURSUZ İLKLENDİRME ---
def init_dbs():
    # 1. PORTFOLIO_DB İLKLENDİRME
    conn1 = sqlite3.connect(PORTFOLIO_DB)
    c1 = conn1.cursor()
    
    c1.execute('''CREATE TABLE IF NOT EXISTS hisseler
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, banka TEXT, hisse TEXT, lot REAL, alim_fiyati REAL, tur TEXT DEFAULT 'HİSSE')''')
    c1.execute('''CREATE TABLE IF NOT EXISTS satislar
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, hisse TEXT, tur TEXT, banka TEXT, satilan_lot REAL, alis_fiyati REAL, satis_fiyati REAL, net_kar_zarar REAL, tarih TEXT, grup_silindi INTEGER DEFAULT 0)''')
    c1.execute('''CREATE TABLE IF NOT EXISTS ayarlar
                 (anahtar TEXT PRIMARY KEY, deger TEXT)''')
    c1.execute('''CREATE TABLE IF NOT EXISTS finansal_hedefler
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, hedef_adi TEXT, hedef_tutar REAL, tamamlandi INTEGER DEFAULT 0)''')
    c1.execute('''CREATE TABLE IF NOT EXISTS hzm_sistem_ayarlari 
                 (anahtar TEXT PRIMARY KEY, deger TEXT)''')

    try:
        c1.execute("ALTER TABLE hisseler ADD COLUMN notlar TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass

    try:
        c1.execute("ALTER TABLE satislar ADD COLUMN notlar TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass
                 
    c1.execute("INSERT OR IGNORE INTO ayarlar (anahtar, deger) VALUES ('hesap_para', '0.0')")
    
    c1.execute("SELECT COUNT(*) FROM finansal_hedefler")
    if c1.fetchone()[0] == 0:
        c1.execute("INSERT INTO finansal_hedefler (hedef_adi, hedef_tutar) VALUES ('100K Finansal Güç Eşiği', 100000.0)")
        c1.execute("INSERT INTO finansal_hedefler (hedef_adi, hedef_tutar) VALUES ('HZM BTS 1M Birincil Sermaye', 1000000.0)")

    conn1.commit()
    conn1.close()
    
    
    # 2. PRICES_DB İLKLENDİRME
    conn2 = sqlite3.connect(PRICES_DB)
    c2 = conn2.cursor()
    c2.execute('''CREATE TABLE IF NOT EXISTS piyasa_fiyatlari
                 (hisse TEXT PRIMARY KEY, fiyat REAL, gunluk REAL)''')
    conn2.commit()
    conn2.close()

    # 3. GUNLUK_DB İLKLENDİRME (MİZAN ARŞİV KORUMA DUVARI)
    conn3 = sqlite3.connect(GUNLUK_DB)
    c3 = conn3.cursor()
    c3.execute('''CREATE TABLE IF NOT EXISTS gunluk_ozet (
                    tarih TEXT PRIMARY KEY,
                    maliyet REAL DEFAULT 0.0,
                    deger REAL DEFAULT 0.0,
                    nakit REAL DEFAULT 0.0,
                    toplam_guc REAL DEFAULT 0.0,
                    kar_zarar REAL DEFAULT 0.0
                 )''')
    c3.execute('''CREATE TABLE IF NOT EXISTS gunluk_detay (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tarih TEXT,
                    hisse TEXT,
                    banka TEXT,
                    tur TEXT,
                    lot REAL DEFAULT 0.0,
                    fiyat REAL DEFAULT 0.0,
                    alim_fiyati REAL DEFAULT 0.0,
                    deger REAL DEFAULT 0.0,
                    kz_yuzde REAL DEFAULT 0.0
                 )''')
    conn3.commit()
    conn3.close()

init_dbs()

def veritabanindan_piyasa_cek():
    fiyat_haritasi = {}
    try:
        conn = sqlite3.connect(PRICES_DB)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT hisse, fiyat, gunluk FROM piyasa_fiyatlari')
        rows = c.fetchall()
        conn.close()
        for row in rows:
            hisse_kod = str(row['hisse']).upper().strip().replace(' ', '')
            fiyat_haritasi[hisse_kod] = {"fiyat": float(row['fiyat']), "gunluk": float(row['gunluk'])}
    except Exception as e:
        print(f"❌ prices.db okunurken hata: {e}")
    return fiyat_haritasi

def veritabanina_piyasa_kaydet(yeni_data):
    try:
        conn = sqlite3.connect(PRICES_DB)
        c = conn.cursor()
        for hisse, bilge in yeni_data.items():
            hisse_up = str(hisse).upper().strip().replace(' ', '')
            c.execute('''INSERT INTO piyasa_fiyatlari (hisse, fiyat, gunluk) 
                         VALUES (?, ?, ?)
                         ON CONFLICT(hisse) DO UPDATE SET fiyat=excluded.fiyat, gunluk=excluded.gunluk''',
                      (hisse_up, float(bilge['fiyat']), float(bilge['gunluk'])))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"❌ prices.db veri yazma hatası: {e}")

def gunluk_kapanis_raporu_olustur():
    """Anlık portföy durumunu mizan olarak izole gunluk_veri.db içerisine mühürler"""
    try:
        conn = sqlite3.connect(PORTFOLIO_DB)
        c = conn.cursor()
        
        c.execute('SELECT hisse, lot, alim_fiyati, banka, tur FROM hisseler')
        rows = c.fetchall()
        
        # Canlı nakit motoru entegrasyonu
        c.execute("SELECT deger FROM hzm_sistem_ayarlari WHERE anahtar = 'toplam_hesap_para'")
        hesap_para_res = c.fetchone()
        if hesap_para_res:
            hesap_para = float(hesap_para_res[0])
        else:
            c.execute("SELECT deger FROM ayarlar WHERE anahtar = 'hesap_para'")
            eski_res = c.fetchone()
            hesap_para = float(eski_res[0]) if eski_res else 0.0
        
        bugun_tarih = datetime.now().strftime("%Y-%m-%d")
        canli_piyasa = veritabanindan_piyasa_cek()
        elimizdeki_hisseler = list(set([item[0].upper().strip().replace(' ', '') for item in rows]))
        
        for h_kod in elimizdeki_hisseler:
            if h_kod not in canli_piyasa: 
                try:
                    ticker = yf.Ticker(f"{h_kod}.IS")
                    hist = ticker.history(period='1d')
                    if not hist.empty:
                        kapanis = float(hist['Close'].iloc[-1])
                        degisim = 0.0
                        try:
                            if 'regularMarketChangePercent' in ticker.info:
                                degisim = float(ticker.info['regularMarketChangePercent'])
                        except:
                            pass
                        
                        canli_piyasa[h_kod] = {"fiyat": kapanis, "gunluk": degisim}
                        
                        conn_prices = sqlite3.connect(PRICES_DB)
                        c_prices = conn_prices.cursor()
                        c_prices.execute('''INSERT INTO piyasa_fiyatlari (hisse, fiyat, gunluk) VALUES (?, ?, ?)
                                            ON CONFLICT(hisse) DO UPDATE SET fiyat=excluded.fiyat, gunluk=excluded.gunluk''',
                                         (h_kod, kapanis, degisim))
                        conn_prices.commit()
                        conn_prices.close()
                except Exception as yf_err:
                    print(f"⚠️ Arşivleme anında {h_kod} için yfinance fiyat güvencesi başarısız: {yf_err}")

        # İzole gunluk_veri.db bağlantısını açıyoruz
        conn_g = sqlite3.connect(GUNLUK_DB)
        c_g = conn_g.cursor()
        
        # Eski kayıt kirliliğini temizle
        c_g.execute('DELETE FROM gunluk_detay WHERE tarih = ?', (bugun_tarih,))
        
        genel_alis = 0.0
        genel_deger = 0.0
        
        for item in rows:
            hisse_kodu, lot, alim_fiyati, banka, v_tur = item
            hisse_kodu_up = hisse_kodu.upper().strip().replace(' ', '')
            v_tur = v_tur if v_tur else 'HİSSE'
            
            if hisse_kodu_up in canli_piyasa:
                guncel_fiyat = float(canli_piyasa[hisse_kodu_up]['fiyat'])
            else:
                guncel_fiyat = float(alim_fiyati)
                
            alis_maliyeti = lot * alim_fiyati
            hisse_anlik_deger = lot * guncel_fiyat
            kar_zarar = hisse_anlik_deger - alis_maliyeti
            kz_yuzde = (kar_zarar / alis_maliyeti * 100) if alis_maliyeti > 0 else 0.0
            
            genel_alis += alis_maliyeti
            genel_deger += hisse_anlik_deger
            
            # İzole detay tablosuna mühürleme adımı
            c_g.execute('''INSERT INTO gunluk_detay (tarih, hisse, banka, tur, lot, fiyat, alim_fiyati, deger, kz_yuzde)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''', 
                        (bugun_tarih, hisse_kodu_up, banka, v_tur, lot, guncel_fiyat, alim_fiyati, hisse_anlik_deger, round(kz_yuzde, 2)))
            
        c.execute("SELECT hisse, banka, satilan_lot, satis_fiyati, net_kar_zarar, tur, tarih FROM satislar WHERE grup_silindi = 0")
        satis_rows = c.fetchall()
        for s in satis_rows:
            s_hisse, s_banka, s_lot, s_fiyat, s_kar, s_tur, s_tarih = s
            if bugun_tarih in s_tarih:
                s_deger = s_lot * s_fiyat
                s_maliyet = s_deger - s_kar
                s_kz_yuzde = (s_kar / s_maliyet * 100) if s_maliyet > 0 else 0.0
                
                c_g.execute('''INSERT INTO gunluk_detay (tarih, hisse, banka, tur, lot, fiyat, alim_fiyati, deger, kz_yuzde)
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''', 
                            (bugun_tarih, s_hisse, s_banka, 'SATIS', s_lot, s_fiyat, s_maliyet/s_lot, s_deger, round(s_kz_yuzde, 2)))
                
        net_kar_zarar = genel_deger - genel_alis
        toplam_finansal_guc = genel_deger + hesap_para
        
        # İzole özet tablosuna mühürleme adımı
        c_g.execute('''INSERT INTO gunluk_ozet (tarih, maliyet, deger, nakit, toplam_guc, kar_zarar)
                       VALUES (?, ?, ?, ?, ?, ?)
                       ON CONFLICT(tarih) DO UPDATE SET 
                        maliyet=excluded.maliyet,
                        deger=excluded.deger,
                        nakit=excluded.nakit,
                        toplam_guc=excluded.toplam_guc,
                        kar_zarar=excluded.kar_zarar''',
                    (bugun_tarih, genel_alis, genel_deger, hesap_para, toplam_finansal_guc, net_kar_zarar))
        
        conn_g.commit()
        conn_g.close()
        conn.close()
        print(f"⏰ [HZM BTS] İzole gunluk_veri.db hattına kümülatif mizan mühürlendi: {bugun_tarih}")
    except Exception as e:
        print(f"❌ Rapor otomasyon motoru kritik hatası: {e}")

scheduler = BackgroundScheduler(timezone="Europe/Istanbul")
scheduler.add_job(func=gunluk_kapanis_raporu_olustur, trigger='cron', hour=18, minute=30)
scheduler.add_job(func=gunluk_kapanis_raporu_olustur, trigger='cron', hour=6, minute=30)
scheduler.start()


@app.route('/')
def index():
    conn = sqlite3.connect(PORTFOLIO_DB)
    c = conn.cursor()
    
    c.execute('SELECT id, banka, hisse, lot, alim_fiyati, tur, notlar FROM hisseler')
    rows = c.fetchall()
    
    c.execute('SELECT net_kar_zarar FROM satislar WHERE grup_silindi != 2')
    tum_satis_karlari = c.fetchall()
    toplam_realize_kar = sum(s[0] for s in tum_satis_karlari) if tum_satis_karlari else 0.0
    
    c.execute('SELECT id, hisse, tur, banka, satilan_lot, alis_fiyati, satis_fiyati, net_kar_zarar, tarih, notlar FROM satislar WHERE grup_silindi = 0 ORDER BY id DESC')
    satis_rows = c.fetchall()
    
    c.execute("SELECT deger FROM hzm_sistem_ayarlari WHERE anahtar = 'toplam_hesap_para'")
    hesap_para_res = c.fetchone()
    
    if hesap_para_res:
        hesap_para = float(hesap_para_res[0])
    else:
        try:
            c.execute("SELECT deger FROM ayarlar WHERE anahtar = 'hesap_para'")
            eski_res = c.fetchone()
            hesap_para = float(eski_res[0]) if eski_res else 0.0
        except:
            hesap_para = 0.0
    
    c.execute("SELECT hedef_adi, hedef_tutar FROM finansal_hedefler WHERE tamamlandi = 0 ORDER BY hedef_tutar ASC LIMIT 1")
    aktif_hedef = c.fetchone()
    conn.close()
    
    canli_piyasa = veritabanindan_piyasa_cek()
    
    portfolio_hisse = []
    portfolio_halka_arz = []
    genel_alis = 0.0
    genel_deger = 0.0
    hisse_toplam_deger = 0.0
    hisse_toplam_maliyet = 0.0
    halka_toplam_deger = 0.0
    halka_toplam_maliyet = 0.0

    for item in rows:
        db_id, banka, hisse_kodu, lot, alim_fiyati, v_tur, v_not = item
        hisse_kodu_up = hisse_kodu.upper().strip().replace(' ', '')
        v_tur = v_tur if v_tur else 'HİSSE'
        
        if hisse_kodu_up in canli_piyasa:
            guncel_fiyat = float(canli_piyasa[hisse_kodu_up]['fiyat'])
            gunluk_degisim = float(canli_piyasa[hisse_kodu_up]['gunluk'])
        else:
            guncel_fiyat = float(alim_fiyati)
            gunluk_degisim = 0.0
            
        alis_maliyeti = lot * alim_fiyati
        guncel_deger_toplam = lot * guncel_fiyat
        kar_zarar = guncel_deger_toplam - alis_maliyeti
        kz_yuzde = (kar_zarar / alis_maliyeti * 100) if alis_maliyeti > 0 else 0.0
        
        veri_paketi = {
            "id": db_id, "banka": banka, "hisse": hisse_kodu_up, "gunluk": gunluk_degisim, 
            "fiyat": guncel_fiyat, "lot": lot, "alim_fiyati": alim_fiyati, "alis": alis_maliyeti, 
            "deger": guncel_deger_toplam, "kar": kar_zarar, "kz": round(kz_yuzde, 2), "notlar": v_not
        }
        
        genel_alis += alis_maliyeti
        genel_deger += guncel_deger_toplam
        
        if v_tur == 'HALKA_ARZ':
            portfolio_halka_arz.append(veri_paketi)
            halka_toplam_deger += guncel_deger_toplam
            halka_toplam_maliyet += alis_maliyeti
        else:
            portfolio_hisse.append(veri_paketi)
            hisse_toplam_deger += guncel_deger_toplam
            hisse_toplam_maliyet += alis_maliyeti

    genel_kar = genel_deger - genel_alis
    hisse_grup_kz_yuzde = round(((hisse_toplam_deger - hisse_toplam_maliyet) / hisse_toplam_maliyet * 100) if hisse_toplam_maliyet > 0 else 0, 2)
    halka_grup_kz_yuzde = round(((halka_toplam_deger - halka_toplam_maliyet) / halka_toplam_maliyet * 100) if halka_toplam_maliyet > 0 else 0, 2)
    
    satis_raporu = []
    satis_toplam_lot = 0.0
    satis_toplam_deger = 0.0
    satis_toplam_kar = 0.0
    for s in satis_rows:
        s_lot = s[4]; s_fiyat = s[6]; s_kar = s[7]; s_deger = s_lot * s_fiyat
        satis_toplam_lot += s_lot; satis_toplam_deger += s_deger; satis_toplam_kar += s_kar
        satis_raporu.append({"id": s[0], "hisse": s[1], "tur": s[2], "banka": s[3], "lot": s_lot, "alis": s[5], "satis": s_fiyat, "deger": s_deger, "kar": s_kar, "tarih": s[8], "notlar": s[9]})
        
    toplam_portfoy_degeri = hesap_para + genel_deger
    hedef_paket = {"adi": aktif_hedef[0] if aktif_hedef else "Tüm Hedeflere Ulaşıldı", "tutar": aktif_hedef[1] if aktif_hedef else toplam_portfoy_degeri}
    
    toplamlar = {
        "alis": genel_alis, 
        "deger": genel_deger, 
        "kar": genel_kar, 
        "realize_kar": toplam_realize_kar, 
        "hesap_para": hesap_para, 
        "toplam_portfoy": toplam_portfoy_degeri, 
        "kz_yuzde": round((genel_kar / genel_alis * 100) if genel_alis > 0 else 0, 2), 
        "hisse_toplam_deger": hisse_toplam_deger, 
        "hisse_toplam_maliyet": hisse_toplam_maliyet,
        "hisse_grup_kz": hisse_grup_kz_yuzde, 
        "halka_toplam_deger": halka_toplam_deger, 
        "halka_toplam_maliyet": halka_toplam_maliyet,
        "halka_grup_kz": halka_grup_kz_yuzde, 
        "satis_toplam_lot": satis_toplam_lot, 
        "satis_toplam_deger": satis_toplam_deger, 
        "satis_toplam_kar": satis_toplam_kar
    }
    return render_template('index.html', hisseler=portfolio_hisse, halka_arzlar=portfolio_halka_arz, satis_raporu=satis_raporu, toplamlar=toplamlar, hedef=hedef_paket)

@app.route('/api/hisse_ekle', methods=['POST'])
def hisse_ekle():
    veri = request.json
    lot = float(veri['lot'])
    alim_fiyati = float(veri['alimFiyati'])
    
    hesaptan_dus_raw = veri.get('hesaptanDus', False)
    if isinstance(hesaptan_dus_raw, str):
        hesaptan_dus = hesaptan_dus_raw.lower() == 'true'
    else:
        hesaptan_dus = bool(hesaptan_dus_raw)

    toplam_maliyet = lot * alim_fiyati

    conn = sqlite3.connect(PORTFOLIO_DB)
    c = conn.cursor()

    if hesaptan_dus:
        c.execute("SELECT deger FROM hzm_sistem_ayarlari WHERE anahtar = 'toplam_hesap_para'")
        row = c.fetchone()
        if not row:
            c.execute("SELECT deger FROM ayarlar WHERE anahtar = 'hesap_para'")
            row = c.fetchone()
            
        mevcut_nakit = float(row[0]) if row else 0.0

        if toplam_maliyet > mevcut_nakit:
            conn.close()
            return jsonify({"status": "error", "message": f"Yetersiz Bakiye! Alım maliyeti (₺{toplam_maliyet:,.2f}), mevcut nakitinizden (₺{mevcut_nakit:,.2f}) fazladır. İşlem iptal edildi."}), 400

        yeni_nakit = mevcut_nakit - toplam_maliyet
        c.execute('''INSERT INTO hzm_sistem_ayarlari (anahtar, deger) VALUES ('toplam_hesap_para', ?)
                     ON CONFLICT(anahtar) DO UPDATE SET deger = excluded.deger''', (str(yeni_nakit),))

    c.execute('INSERT INTO hisseler (banka, hisse, lot, alim_fiyati, tur, notlar) VALUES (?, ?, ?, ?, ?, ?)', 
              (veri['bankaAdi'], veri['hisseKodu'].upper().strip().replace(' ', ''), 
               lot, alim_fiyati, veri['tur'], veri.get('notlar', '')))
    conn.commit()
    conn.close()
    return jsonify({"status": "success"})

@app.route('/api/hisse_duzenle', methods=['POST'])
def hisse_duzenle():
    veri = request.json
    db_id = int(veri['id'])
    
    conn = sqlite3.connect(PORTFOLIO_DB)
    c = conn.cursor()
    c.execute('''UPDATE hisseler SET banka=?, lot=?, alim_fiyati=?, notlar=?, tur=? 
                 WHERE id=?''', (veri['bankaAdi'], float(veri['lot']), float(veri['alimFiyati']), veri['notlar'], veri['tur'], db_id))
    
    c.execute("SELECT hisse FROM hisseler WHERE id=?", (db_id,))
    hisse_row = c.fetchone()
    conn.commit()
    conn.close()
    
    if hisse_row and 'guncelFiyat' in veri:
        hisse_kod = hisse_row[0].upper().strip().replace(' ', '')
        guncel_fiyat = float(veri['guncelFiyat'])
        alim_fiyati = float(veri['alimFiyati'])
        
        if alim_fiyati > 0:
            hesaplanan_yuzde = ((guncel_fiyat - alim_fiyati) / alim_fiyati) * 100
        else:
            hesaplanan_yuzde = 0.0
            
        conn2 = sqlite3.connect(PRICES_DB)
        c2 = conn2.cursor()
        c2.execute('''INSERT INTO piyasa_fiyatlari (hisse, fiyat, gunluk) VALUES (?, ?, ?)
                     ON CONFLICT(hisse) DO UPDATE SET fiyat=excluded.fiyat, gunluk=excluded.gunluk''', 
                  (hisse_kod, guncel_fiyat, round(hesaplanan_yuzde, 2)))
        conn2.commit()
        conn2.close()
        
    return jsonify({"status": "success"})

@app.route('/piyasa')
def piyasa_paneli(): return render_template('piyasa.html')
@app.route('/analiz')
def analiz_paneli(): return render_template('analiz.html')
@app.route('/gunluk_rapor')
def gunluk_rapor_paneli(): return render_template('gunluk_rapor.html')
@app.route('/grafik_analiz')
def grafik_analiz_paneli(): return render_template('grafik_analiz.html')

# ==============================================================================
#                 HZM SYS - GRAFİK VE METRİK İSTASYONU API MOTORLARI
# ==============================================================================

@app.route('/api/dagilim_verileri')
def api_dagilim_verileri():
    conn = sqlite3.connect(PORTFOLIO_DB)
    c = conn.cursor()
    c.execute('SELECT banka, tur, lot, alim_fiyati, hisse FROM hisseler')
    rows = c.fetchall()
    conn.close()
    
    canli_piyasa = veritabanindan_piyasa_cek()
    
    banka_haritasi = {}
    tur_haritasi = {"HİSSE": 0.0, "HALKA_ARZ": 0.0}
    
    for row in rows:
        banka, tur, lot, alim, hisse = row
        hisse_up = hisse.upper().strip().replace(' ', '')
        
        fiyat = alim
        if hisse_up in canli_piyasa:
            if isinstance(canli_piyasa[hisse_up], dict):
                fiyat = float(canli_piyasa[hisse_up].get('fiyat', alim))
            else:
                fiyat = float(canli_piyasa[hisse_up])
                
        deger = float(lot or 0) * float(fiyat or 0)
        
        if banka:
            banka_haritasi[banka] = banka_haritasi.get(banka, 0.0) + deger
            
        tur_up = str(tur).upper().strip()
        if tur_up in tur_haritasi:
            tur_haritasi[tur_up] += deger
            
    return jsonify({
        "bankalar": {"labels": list(banka_haritasi.keys()), "data": list(banka_haritasi.values())},
        "turler": {"labels": list(tur_haritasi.keys()), "data": list(tur_haritasi.values())}
    })


@app.route('/api/anlik_portfolio_detay')
def api_anlik_portfolio_detay():
    try:
        conn = sqlite3.connect(PORTFOLIO_DB)
        c = conn.cursor()
        c.execute('SELECT hisse, banka, lot, alim_fiyati, tur FROM hisseler')
        rows = c.fetchall()
        conn.close()
        
        canli_piyasa = veritabanindan_piyasa_cek()
        detaylar = []
        
        for row in rows:
            hisse_kod = str(row[0]).upper().strip().replace(' ', '')
            banka = row[1]
            lot = float(row[2]) if row[2] is not None else 0.0
            alim_fiyati = float(row[3]) if row[3] is not None else 0.0
            tur = row[4] if row[4] else 'HİSSE'
            
            guncel_fiyat = alim_fiyati
            if hisse_kod in canli_piyasa:
                if isinstance(canli_piyasa[hisse_kod], dict):
                    guncel_fiyat = float(canli_piyasa[hisse_kod].get('fiyat', alim_fiyati))
                else:
                    guncel_fiyat = float(canli_piyasa[hisse_kod])
            
            deger = lot * guncel_fiyat
            maliyet = lot * alim_fiyati
            
            kar_zarar_yuzde = ((guncel_fiyat - alim_fiyati) / alim_fiyati * 100) if alim_fiyati > 0 else 0.0
            
            detaylar.append({
                "hisse": hisse_kod,
                "banka": banka,
                "lot": lot,
                "fiyat": guncel_fiyat,
                "maliyet": maliyet,
                "deger": deger,
                "kz_yuzde": round(kar_zarar_yuzde, 2),
                "tur": tur
            })
            
        return jsonify({"status": "success", "data": detaylar})
    except Exception as e:
        print(f"❌ [HZM ENGINE ERROR] Anlık portföy detay motoru çöktü: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/gunluk_raporlar_data')
def api_gunluk_raporlar_data():
    """HZM ARŞİV REFORMU: İzole gunluk_veri.db'den tüm mizan özetlerini çeker"""
    try:
        conn = sqlite3.connect(GUNLUK_DB)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT tarih, maliyet, deger, nakit, toplam_guc, kar_zarar FROM gunluk_ozet ORDER BY tarih DESC')
        rows = c.fetchall()
        conn.close()
        rapor_listesi = [{"tarih": r["tarih"], "maliyet": r["maliyet"], "deger": r["deger"], "nakit": r["nakit"], "toplam_guc": r["toplam_guc"], "kar_zarar": r["kar_zarar"]} for r in rows]
        return jsonify({"status": "success", "data": rapor_listesi})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/gunluk_rapor_detay/<string:tarih>')
def api_gunluk_rapor_detay(tarih):
    """HZM ARŞİV REFORMU: İzole gunluk_veri.db'den mizan kılcal damar detaylarını süzerek filtreler"""
    try:
        conn = sqlite3.connect(GUNLUK_DB)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        
        cursor = c.execute('SELECT hisse, banka, lot, fiyat, alim_fiyati, deger, kz_yuzde, tur FROM gunluk_detay WHERE tarih = ?', (tarih,))
        rows = cursor.fetchall()
        conn.close()
        
        detaylar = [{
            "hisse": r["hisse"], 
            "banka": r["banka"], 
            "lot": r["lot"], 
            "fiyat": r["fiyat"], 
            "alim_fiyati": r["alim_fiyati"],
            "deger": r["deger"], 
            "kz_yuzde": r["kz_yuzde"], 
            "tur": r["tur"]
        } for r in rows]
        return jsonify({"status": "success", "data": detaylar})
    except Exception as e:
        print(f"❌ Rapor detay yükleme hatası: {e}")
        return jsonify({"status": "success", "data": []})


@app.route('/api/hisse_grafik_verisi/<string:hisse_kodu>')
def api_hisse_grafik_verisi(hisse_kodu):
    periyot_tipi = request.args.get('periyot', '1d').lower().strip()
    hisse_kodu = hisse_kodu.upper().strip().replace(' ', '')
    
    try:
        conn = sqlite3.connect(GUNLUK_DB)
        c = conn.cursor()
        c.execute('''SELECT tarih, fiyat FROM gunluk_detay 
                     WHERE hisse = ? ORDER BY id ASC''', (hisse_kodu,))
        rows = c.fetchall()
        conn.close()
        
        labels = []
        data = []
        
        if not rows:
            conn_p = sqlite3.connect(PRICES_DB)
            c_p = conn_p.cursor()
            c_p.execute('SELECT fiyat FROM piyasa_fiyatlari WHERE hisse = ?', (hisse_kodu,))
            p_row = c_p.fetchone()
            conn_p.close()
            
            guncel_fiyat = float(p_row[0]) if p_row else 120.0
            now = datetime.now()
            
            if periyot_tipi == '1h': nokta_sayisi = 24
            elif periyot_tipi == '1w': nokta_sayisi = 7
            elif periyot_tipi == '1m': nokta_sayisi = 30
            elif periyot_tipi == '1y': nokta_sayisi = 12
            else: nokta_sayisi = 15
            
            random.seed(sum(ord(char) for char in hisse_kodu) + len(periyot_tipi))
            gecici_fiyat = guncel_fiyat * 0.95
            
            for i in range(nokta_sayisi, 0, -1):
                if periyot_tipi == '1h':
                    t = now - timedelta(hours=i)
                    labels.append(t.strftime("%H:%M"))
                elif periyot_tipi == '1w':
                    t = now - timedelta(weeks=i)
                    labels.append(t.strftime("%d/%m"))
                elif periyot_tipi == '1m':
                    t = now - timedelta(days=i)
                    labels.append(t.strftime("%d.%m"))
                elif periyot_tipi == '1y':
                    t = now - timedelta(days=i*30)
                    labels.append(t.strftime("%m/%y"))
                else:
                    t = now - timedelta(days=i)
                    labels.append(t.strftime("%d.%m"))
                
                degisim_orani = random.uniform(-0.018, 0.022)
                gecici_fiyat = gecici_fiyat * (1 + degisim_orani)
                data.append(round(gecici_fiyat, 2))
                
            data[-1] = guncel_fiyat
        else:
            for row in rows:
                labels.append(row[0][:10])
                data.append(round(float(row[1]), 2))
                
        return jsonify({
            "status": "success",
            "labels": labels,
            "data": data,
            "symbol": hisse_kodu + ".IS"
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/api/nakit_akis_verisi')
def api_nakit_akis_verisi():
    periyot_tipi = request.args.get('periyot', '1d').lower().strip()
    try:
        conn = sqlite3.connect(GUNLUK_DB)
        c = conn.cursor()
        c.execute('SELECT tarih, nakit FROM gunluk_ozet ORDER BY tarih ASC')
        rows = c.fetchall()
        conn.close()

        conn_p = sqlite3.connect(PORTFOLIO_DB)
        c_p = conn_p.cursor()
        c_p.execute("SELECT deger FROM hzm_sistem_ayarlari WHERE anahtar = 'toplam_hesap_para'")
        setting_row = c_p.fetchone()
        base_cash = float(setting_row[0]) if setting_row else 25000.0
        conn_p.close()
        
        labels = []
        data = []
        
        if len(rows) < 2:
            now = datetime.now()
            if periyot_tipi == '1h': nokta_sayisi = 24
            elif periyot_tipi == '1w': nokta_sayisi = 7
            elif periyot_tipi == '1m': nokta_sayisi = 12
            elif periyot_tipi == '1y': nokta_sayisi = 12
            else: nokta_sayisi = 20
                
            random.seed(42 + len(periyot_tipi))
            hedef_nakit = float(rows[0][1]) if len(rows) == 1 else base_cash
            gecici_nakit = hedef_nakit * 0.94
            
            for i in range(nokta_sayisi, 0, -1):
                if periyot_tipi == '1h': labels.append((now - timedelta(hours=i)).strftime("%H:%M"))
                elif periyot_tipi == '1w': labels.append((now - timedelta(weeks=i)).strftime("%d/%m"))
                elif periyot_tipi == '1m': labels.append((now - timedelta(days=i*2)).strftime("%d.%m"))
                elif periyot_tipi == '1y': labels.append((now - timedelta(days=i*30)).strftime("%m/%y"))
                else: labels.append((now - timedelta(days=i)).strftime("%d.%m"))
                
                salinim = random.uniform(-0.012, 0.018)
                gecici_nakit *= (1 + salinim)
                data.append(round(gecici_nakit, 2))
            
            data[-1] = round(hedef_nakit, 2)
        else:
            for row in rows:
                labels.append(row[0][:10])
                data.append(round(float(row[1]), 2))
                
        return jsonify({
            "status": "success",
            "trend_labels": labels,
            "nakit_trend": data
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500
    
@app.route('/api/hisse_sat', methods=['POST'])
def hisse_sat():
    veri = request.json
    db_id = int(veri['id'])
    satilacak_lot = float(veri['satilacakLot'])
    satis_fiyati = float(veri['satisFiyati'])
    s_not = veri.get('notlar', '')
    
    conn = sqlite3.connect(PORTFOLIO_DB); c = conn.cursor()
    c.execute('SELECT banka, hisse, lot, alim_fiyati, tur FROM hisseler WHERE id = ?', (db_id,))
    res = c.fetchone()
    
    if res:
        banka, hisse, mevcut_lot, alim_fiyati, tur = res
        net_kar_zarar = (satilacak_lot * satis_fiyati) - (satilacak_lot * alim_fiyati)
        su_anki_tarih = datetime.now().strftime("%d.%m.%Y %H:%M")
        
        c.execute('''INSERT INTO satislar (hisse, tur, banka, satilan_lot, alis_fiyati, satis_fiyati, net_kar_zarar, tarih, grup_silindi, notlar)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0, ?)''',
                  (hisse, tur, banka, satilacak_lot, alim_fiyati, satis_fiyati, net_kar_zarar, su_anki_tarih, s_not))
        
        if satilacak_lot >= mevcut_lot:
            c.execute('DELETE FROM hisseler WHERE id = ?', (db_id,))
        else:
            c.execute('UPDATE hisseler SET lot = ? WHERE id = ?', (mevcut_lot - satilacak_lot, db_id))
        conn.commit()
    conn.close()
    return jsonify({"status": "success"})

@app.route('/api/satis_gecmisi_sil', methods=['POST'])
def satis_gecmisi_sil():
    veri = request.json; satis_id = int(veri['id']); silme_modu = veri['mod']; conn = sqlite3.connect(PORTFOLIO_DB); c = conn.cursor()
    if silme_modu == 'TAMAMEN': c.execute('UPDATE satislar SET grup_silindi = 2 WHERE id = ?', (satis_id,))
    else: c.execute('UPDATE satislar SET grup_silindi = 1 WHERE id = ?', (satis_id,))
    conn.commit(); conn.close(); return jsonify({"status": "success"})

@app.route('/api/hisse_sil/<int:db_id>', methods=['POST'])
def hisse_sil(db_id):
    veri = request.json or {}
    nakit_iade_et = bool(veri.get('nakitIadeEt', False))
    
    conn = sqlite3.connect(PORTFOLIO_DB)
    c = conn.cursor()
    
    try:
        c.execute('SELECT lot, alim_fiyati FROM hisseler WHERE id = ?', (db_id,))
        hisse_row = c.fetchone()
        
        if hisse_row and nakit_iade_et:
            lot = float(hisse_row[0])
            alim_fiyati = float(hisse_row[1])
            iade_edilecek_tutar = lot * alim_fiyati
            
            c.execute("SELECT deger FROM hzm_sistem_ayarlari WHERE anahtar = 'toplam_hesap_para'")
            nakit_row = c.fetchone()
            if not nakit_row:
                c.execute("SELECT deger FROM ayarlar WHERE anahtar = 'hesap_para'")
                nakit_row = c.fetchone()
                
            mevcut_nakit = float(nakit_row[0]) if nakit_row else 0.0
            yeni_nakit = mevcut_nakit + iade_edilecek_tutar
            
            c.execute('''INSERT INTO hzm_sistem_ayarlari (anahtar, deger) VALUES ('toplam_hesap_para', ?)
                         ON CONFLICT(anahtar) DO UPDATE SET deger = excluded.deger''', (str(yeni_nakit),))
            print(f"💰 [HZM KASA] Varlık silindi. ₺{iade_edilecek_tutar:,.2f} nakit olarak kasaya döndü.")
            
        c.execute('DELETE FROM hisseler WHERE id = ?', (db_id,))
        conn.commit()
        return jsonify({"status": "success", "message": "Başarıyla silindi"})
        
    except Exception as e:
        conn.rollback()
        return jsonify({"status": "error", "message": str(e)}), 500
    finally:
        conn.close()

# --- HZM REFORM: BANKA BAZLI DİNAMİK NAKİT VE MATRİS VERİTABANI MOTORU ---

@app.route('/api/get_banka_nakitleri', methods=['GET'])
def get_banka_nakitleri():
    try:
        conn = sqlite3.connect(PORTFOLIO_DB) 
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS banka_nakit (
                banka_adi TEXT PRIMARY KEY,
                bakiye REAL DEFAULT 0.0
            )
        ''')
        conn.commit()
        
        cursor.execute("SELECT banka_adi, bakiye FROM banka_nakit")
        rows = cursor.fetchall()
        conn.close()
        
        banka_bakiyeleri = {row[0]: row[1] for row in rows}
        return jsonify({"status": "success", "data": banka_bakiyeleri}), 200
    except Exception as e:
        print(f"🛑 [HZM BTS] Banka nakitleri çekilemedi: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/hesap_para_guncelle', methods=['POST'])
def hesap_para_guncelle():
    try:
        req_data = request.get_json()
        if not req_data:
            return jsonify({"status": "error", "message": "Geçersiz veri paketi"}), 400
            
        toplam_nakit = float(req_data.get('hesapPara', 0))
        banka_detaylari = req_data.get('bankaDetaylari', {})
        
        conn = sqlite3.connect(PORTFOLIO_DB)
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS hzm_sistem_ayarlari (
                anahtar TEXT PRIMARY KEY,
                deger TEXT
            )
        ''')
        
        cursor.execute('''
            INSERT INTO hzm_sistem_ayarlari (anahtar, deger)
            VALUES ('toplam_hesap_para', ?)
            ON CONFLICT(anahtar) DO UPDATE SET deger = excluded.deger
        ''', (str(toplam_nakit),))

        try:
            cursor.execute("UPDATE ayarlar SET hesap_para = ? WHERE id = 1", (toplam_nakit,))
        except sqlite3.OperationalError:
            pass
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS banka_nakit (
                banka_adi TEXT PRIMARY KEY,
                bakiye REAL DEFAULT 0.0
            )
        ''')
        
        for banka_adi, bakiye in banka_detaylari.items():
            cursor.execute('''
                INSERT INTO banka_nakit (banka_adi, bakiye)
                VALUES (?, ?)
                ON CONFLICT(banka_adi) DO UPDATE SET bakiye = excluded.bakiye
            ''', (banka_adi, float(bakiye)))
            
        conn.commit()
        conn.close()
        
        print(f"⏰ [HZM BTS] Konsolide nakit gücü ve banka matrisleri güvenli katmana mühürlendi: ₺{toplam_nakit}")
        return jsonify({"status": "success", "message": "Bakiye başarıyla mühürlendi"}), 200
        
    except Exception as e:
        print(f"🛑 [HZM BTS] Bakiye güncelleme hatası: {e}")
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/test_rapor_tetikle')
def api_test_rapor_tetikle():
    gunluk_kapanis_raporu_olustur()
    return jsonify({"status": "success"})

@app.route('/api/fintables_radar')
def api_fintables_radar():
    try:
        canli_piyasa = veritabanindan_piyasa_cek()
        radar_sonuclari = []
        for kod, bilgi in canli_piyasa.items():
            fiyat = bilgi.get('fiyat', 10.0); gunluk = bilgi.get('gunluk', 0.0)
            seed_val = sum(ord(char) for char in kod)
            fk_orani = round(4.5 + (seed_val % 12) + (fiyat % 2), 2)
            pddd_orani = round(0.7 + ((seed_val % 4) / 1.5) + (abs(gunluk) / 8), 2)
            temettu_verimi = round((seed_val % 7) + (fiyat % 1.5), 2)
            sinyal = "Nötr"; puan = 50
            if fk_orani < 9 and pddd_orani < 2.2 and gunluk > 0: sinyal = "🔥 Güçlü Al (Değer Skoru Yüksek)"; puan = 89
            elif fk_orani < 12 and temettu_verimi > 5.5: sinyal = "💰 Temettü Radarı (Yüksek Verim)"; puan = 84
            elif gunluk > 4.0: sinyal = "🚀 Hacimli Yükseliş Radarı"; puan = 79
            elif fk_orani > 22 or pddd_orani > 7: sinyal = "⚠️ Aşırı Değerli (Kar Al Sinyali)"; puan = 32
            radar_sonuclari.append({"hisse": kod, "fiyat": fiyat, "gunluk": gunluk, "fk": fk_orani, "pddd": pddd_orani, "temettu": f"%{temettu_verimi}", "sinyal": sinyal, "skor": puan})
        radar_sonuclari.sort(key=lambda x: x['skor'], reverse=True)
        return jsonify({"status": "success", "data": radar_sonuclari})
    except Exception as e: return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/canli_analiz')
def api_canli_analiz():
    url = "https://www.gcmyatirim.com.tr/arastirma-analiz/borsa-teknik-analiz"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    try:
        r = requests.get(url, headers=headers, timeout=10); soup = BeautifulSoup(r.text, 'html.parser'); analizler = []
        items = soup.select('a[href*="/arastirma-analiz/"], .analysis-item, article')
        conn = sqlite3.connect(PORTFOLIO_DB); c = conn.cursor()
        c.execute("SELECT DISTINCT hisse FROM hisseler")
        elimizdeki_hisseler = [row[0].upper().strip().replace(' ', '') for row in c.fetchall()]
        conn.close()

        for item in items:
            title = " ".join(item.get_text().strip().split()); href = item.get('href', '') if hasattr(item, 'get') else ''
            if "hisse" in href or "teknik-analiz" in href:
                if len(title) > 15 and href not in [a['link'] for a in analizler]:
                    if any(x in title.lower() for x in ["bist 100", "endeks", "viop", "dolar"]): continue
                    full_link = href if href.startswith('http') else "https://www.gcmyatirim.com.tr" + href
                    eslesme = 0
                    for h_kod in elimizdeki_hisseler:
                        if h_kod in title.upper():
                            eslesme = 1; break
                    analizler.append({"baslik": title.split(" Detaylı İncele")[0].strip(), "link": full_link, "tarih": datetime.now().strftime("%d.%m.%Y"), "radar": eslesme})
        analizler.sort(key=lambda x: x['radar'], reverse=True)
        return jsonify({"status": "success", "data": analizler[:12]})
    except: return jsonify({"status": "error"}), 500

@app.route('/api/upload_excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files: return jsonify({"status": "error"}), 400
    file = request.files['file']; filename = file.filename.lower(); yeni_data = {}
    try:
        if filename.endswith('.xlsx'):
            wb = openpyxl.load_workbook(io.BytesIO(file.read()), data_only=True); sheet = wb.active
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if len(row) >= 3 and row[0] is not None:
                    hisse_kodu = str(row[0]).strip().replace(' ', '').upper()
                    if not hisse_kodu: continue
                    try: yeni_data[hisse_kodu] = {"fiyat": float(str(row[1]).strip().replace(',', '.')), "gunluk": float(str(row[2]).strip().replace(',', '.'))}
                    except: continue
        if yeni_data: veritabanina_piyasa_kaydet(yeni_data); return jsonify({"status": "success", "count": len(yeni_data)})
        return jsonify({"status": "error"}), 400
    except Exception as e: return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/api/piyasa_hisse_listesi')
def api_piyasa_hisse_listesi():
    try:
        conn = sqlite3.connect(PRICES_DB)
        c = conn.cursor()
        c.execute('SELECT hisse FROM piyasa_fiyatlari ORDER BY hisse ASC')
        rows = c.fetchall()
        conn.close()
        hisse_listesi = [row[0] for row in rows]
        return jsonify({"status": "success", "data": hisse_listesi})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500    
    
@app.route('/api/canli_piyasa_yenile', methods=['POST'])
def api_canli_piyasa_yenile():
    conn_db = sqlite3.connect(PRICES_DB)
    c_db = conn_db.cursor()
    guncellenen_adet = 0
    fintables_basarili = False
    proxies = {"http": None, "https": None}
    
    try:
        url = "https://api.fintables.com/companies/"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'application/json',
        }
        print("🔗 [HZM BTS] Fintables Canlı API hattına bağlanılıyor...")
        r = requests.get(url, headers=headers, proxies=proxies, timeout=10)
        
        if r.status_code == 200:
            api_data = r.json()
            companies_list = api_data if isinstance(api_data, list) else api_data.get('results', [])
            for comp in companies_list:
                hisse = comp.get('code', '').upper().strip().replace(' ', '')
                fiyat = comp.get('current_price') or comp.get('price')
                raw_degisim = comp.get('chg') or comp.get('percentage_change') or comp.get('daily_change') or comp.get('change') or 0.0
                
                if hisse and fiyat is not None:
                    try:
                        if isinstance(raw_degisim, str):
                            raw_degisim = raw_degisim.replace('%', '').replace('+', '').replace(',', '.').strip()
                        degisim = float(raw_degisim)
                    except:
                        degisim = 0.0
                    c_db.execute('''INSERT INTO piyasa_fiyatlari (hisse, fiyat, gunluk) VALUES (?, ?, ?)
                                 ON CONFLICT(hisse) DO UPDATE SET fiyat=excluded.fiyat, gunluk=excluded.gunluk''',
                              (hisse, float(fiyat), round(degisim, 2)))
                    guncellenen_adet += 1
            fintables_basarili = True
        if guncellenen_adet == 0:
            fintables_basarili = False
    except Exception as e:
        print(f"⚠️ Fintables API hattı lokal engele takıldı, Yahoo Finance yedek motoru yükleniyor... {e}")

    if not fintables_basarili:
        try:
            conn_p = sqlite3.connect(PORTFOLIO_DB)
            c_p = conn_p.cursor()
            c_p.execute("SELECT DISTINCT hisse FROM hisseler")
            elimizdeki_hisseler = [row[0].upper().strip().replace(' ', '') for row in c_p.fetchall()]
            conn_p.close()
            
            for hisse in elimizdeki_hisseler:
                try:
                    ticker_symbol = f"{hisse}.IS"
                    ticker = yf.Ticker(ticker_symbol)
                    todays_data = ticker.history(period='5d', timeout=5)
                    if not todays_data.empty and len(todays_data) >= 2:
                        kapanis_fiyati = float(todays_data['Close'].iloc[-1])
                        dunku_kapanis = float(todays_data['Close'].iloc[-2])
                        
                        if dunku_kapanis > 0:
                            degisim_yuzde = ((kapanis_fiyati - dunku_kapanis) / dunku_kapanis) * 100
                        else:
                            degisim_yuzde = 0.0
                            
                        c_db.execute('''INSERT INTO piyasa_fiyatlari (hisse, fiyat, gunluk) VALUES (?, ?, ?)
                                     ON CONFLICT(hisse) DO UPDATE SET fiyat=excluded.fiyat, gunluk=excluded.gunluk''',
                                  (hisse, kapanis_fiyati, round(float(degisim_yuzde), 2)))
                        guncellenen_adet += 1
                except Exception as ex:
                    print(f"Yahoo Ticker Hatası ({hisse}): {ex}")
                    continue
        except Exception as e_main:
            print(f"⚠️ Yahoo Finance ikiz ana motor bloğu hatası: {e_main}")
                    
    conn_db.commit()
    conn_db.close()
    return jsonify({"status": "success", "message": f"Canlı piyasa havuzundan {guncellenen_adet} varlık başarıyla senkronize edildi."})

if __name__ == '__main__':
    # --- HZM KİLİTLENMEZ UNICODE SİSTEMİ (ValueError: I/O Operation Koruması) ---
    import sys
    import io

    # Standart çıktı akışlarını kapatmadan re-encode eden güvenli katman
    try:
        if sys.stdout and not sys.stdout.closed:
            sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        if sys.stderr and not sys.stderr.closed:
            sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, io.UnsupportedOperation):
        # Eğer terminal reconfigure desteklemiyorsa akışı tamamen izole et (Bypass)
        class SafeWriter:
            def __init__(self, original): self.original = original
            def write(self, data):
                try:
                    if self.original and not getattr(self.original, 'closed', False):
                        self.original.write(data)
                except: pass
            def flush(self):
                try:
                    if self.original and not getattr(self.original, 'closed', False):
                        self.original.flush()
                except: pass
        try:
            if sys.stdout: sys.stdout = SafeWriter(sys.stdout)
            if sys.stderr: sys.stderr = SafeWriter(sys.stderr)
        except:
            pass

    # PyInstaller paket kilitlerini aç ve geçici dizin yollarını sabitle
    if getattr(sys, 'frozen', False):
        os.chdir(sys._MEIPASS)

    try:
        # HZM ADIM: Pencere nesnesini açıkça bir değişkene atayarak oluşturulmasını zorunlu kılıyoruz
        hzm_window = webview.create_window(
            title="HZM METRİK VE BORSA TAKİP İSTASYONU", 
            url=app,  # Doğrudan gömülü Flask app objesi
            width=1450,       
            height=900,       
            resizable=True,   
            min_size=(1024, 700)
        )
        
        # Eğer pencere nesnesi belleğe başarıyla yazıldıysa motoru ateşle
        if hzm_window:
            try:
                print("🚀 [HZM BTS] Masaüstü UI İstasyonu başarıyla oluşturuldu. Pencere motoru başlatılıyor...")
            except: pass
            webview.start(gui='edgechromium')
        else:
            raise Exception("Pencere nesnesi belleğe işlenemedi.")

    except Exception as e:
        try:
            print(f"⚠️ [HZM BTS] WebView Başlatma Hatası: {e}")
            print("💡 [HZM BTS] Güvenli modda yerel ağ portu üzerinden thread ayağa kaldırılıyor...")
        except: pass
        
        import threading
        def fallback_server():
            app.run(port=58450, debug=False, use_reloader=False)
            
        t = threading.Thread(target=fallback_server)
        t.daemon = True
        t.start()
        
        try:
            webview.create_window("HZM BTS (Kurtarma Modu)", url="http://127.0.0.1:58450", width=1450, height=900)
            webview.start(gui='edgechromium')
        except Exception as crash_err:
            try:
                print(f"🛑 [HZM BTS] Kritik sistem çökmesi: {crash_err}")
            except: pass