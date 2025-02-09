from bs4 import BeautifulSoup

def extract_company_names():
    try:
        # Netherland.txt dosyasını oku
        with open('Hollanda.txt', 'r', encoding='utf-8') as file:
            content = file.read()

        # HTML içeriğini parse et
        soup = BeautifulSoup(content, 'html.parser')

        # Tüm şirket isimlerini bul (<th> etiketleri içinde)
        company_names = [th.text.strip() for th in soup.find_all('th')]

        # Şirket isimlerini yeni dosyaya kaydet
        with open('Holland.txt', 'w', encoding='utf-8') as output_file:
            for company in company_names:
                output_file.write(company + '\n')

        print(f"✅ Toplam {len(company_names)} şirket ismi başarıyla ayıklandı ve kaydedildi.")
        print("📄 Dosya adı: Holland.txt")

    except FileNotFoundError:
        print("❌ Hata: Hollanda.txt dosyası bulunamadı!")
    except Exception as e:
        print(f"❌ Hata oluştu: {str(e)}")

if __name__ == "__main__":
    extract_company_names() 