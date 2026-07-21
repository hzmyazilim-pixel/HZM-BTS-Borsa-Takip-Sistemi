import sqlite3

def hzm_bts_veritabani_onar():
    print("⚡ HZM BTS Veritabanı Onarım İstasyonu Başlatıldı...")
    
    try:
        conn = sqlite3.connect('portfolio.db')
        c = conn.cursor()
        
        # 1. Hisseler tablosunu kontrol et ve notlar sütununu ekle
        try:
            c.execute("ALTER TABLE hisseler ADD COLUMN notlar TEXT DEFAULT ''")
            print("✅ 'hisseler' tablosuna 'notlar' sütunu güvenle eklendi.")
        except sqlite3.OperationalError as e:
            if "duplicate column name" in str(e).lower() or "already exists" in str(e).lower():
                print("ℹ️ 'hisseler' tablosunda 'notlar' sütunu zaten mevcut, geçildi.")
            else:
                print(f"❌ Hisseler tablosu onarılırken hata: {e}")

        # 2. Satislar tablosunu kontrol et ve notlar sütununu ekle
        try:
            c.execute("ALTER TABLE satislar ADD COLUMN notlar TEXT DEFAULT ''")
            print("✅ 'satislar' tablosuna 'notlar' sütunu güvenle eklendi.")
        except sqlite3.OperationalError as e:
            if "duplicate column name" in str(e).lower() or "already exists" in str(e).lower():
                print("ℹ️ 'satislar' tablosunda 'notlar' sütunu zaten mevcut, geçildi.")
            else:
                print(f"❌ Satislar tablosu onarılırken hata: {e}")
                
        # 3. Finansal Hedefler tablosunun varlığını garanti altına al
        c.execute('''CREATE TABLE IF NOT EXISTS finansal_hedefler
                     (id INTEGER PRIMARY KEY AUTOINCREMENT, hedef_adi TEXT, hedef_tutar REAL, tamamlandi INTEGER DEFAULT 0)''')
        
        c.execute("SELECT COUNT(*) FROM finansal_hedefler")
        if c.fetchone()[0] == 0:
            c.execute("INSERT INTO finansal_hedefler (hedef_adi, hedef_tutar) VALUES ('100K Finansal Güç Eşiği', 100000.0)")
            c.execute("INSERT INTO finansal_hedefler (hedef_adi, hedef_tutar) VALUES ('HZM BTS 1M Birincil Sermaye', 1000000.0)")
            print("✅ Varsayılan HZM BTS sistem hedefleri eklendi.")

        conn.commit()
        conn.close()
        print("\n🚀 [HZM BTS] Onarım tamamlandı! Tüm eski verileriniz korundu. Şimdi 'app.py'yi çalıştırabilirsiniz.")
        
    except Exception as ex:
        print(f"❌ Kritik Genel Hata: {ex}")

if __name__ == '__main__':
    hzm_bts_veritabani_onar()