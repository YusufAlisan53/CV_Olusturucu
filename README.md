# Jake's Resume - AI Destekli CV Oluşturucu (Word + PDF)

Bu proje, **Jake's Resume** tasarımını birebir aynı tipografi, kenar boşluğu, small-caps başlıklar, yatay çizgiler ve tıklanabilir linklerle önce **Word (`.docx`)** olarak oluşturan ve ardından otomatik olarak **PDF (`.pdf`)** formatına dönüştüren bir CV şablonudur.

## Nasıl Kullanılır?

### 1. AI Asistan ile Konuşarak (Önerilen)
Bu klasörü açıp asistanla doğrudan sohbet etmeniz yeterlidir:
- *"Merhaba, benim için CV oluşturur musun?"* diyerek adım adım soruları yanıtlayabilirsiniz.
- Veya okulunuzu, stajlarınızı, projelerinizi ve bildiğiniz teknolojileri serbestçe yazıp *"Bunlarla bana CV hazırla"* diyebilirsiniz.
- Asistan anlattıklarınızı profesyonel CV maddelerine dönüştürür, `cv_data.json` dosyasını günceller ve `generate_cv.py` kodunu çalıştırarak `.docx` + `.pdf` dosyalarınızı hazırlar.

### 2. Коdla / JSON ile Manuel Kullanım
1. [cv_data.json](file:///c:/Users/yusuf/Desktop/cv/cv_data.json) dosyasındaki alanları kendi bilgilerinizle düzenleyin.
2. Aşağıdaki komutla Word ve PDF çıktısını üretin:
   ```powershell
   & "C:\Users\yusuf\AppData\Local\Python\pythoncore-3.14-64\python.exe" generate_cv.py --data cv_data.json
   ```
   Farklı bir dosya adıyla kaydetmek için:
   ```powershell
   & "C:\Users\yusuf\AppData\Local\Python\pythoncore-3.14-64\python.exe" generate_cv.py --data cv_data.json --output Ad_Soyad_CV
   ```

## Dosya Yapısı
- [generate_cv.py](file:///c:/Users/yusuf/Desktop/cv/generate_cv.py): `python-docx` ile Word belgesini oluşturan ve Microsoft Word üzerinden PDF'e çeviren ana Python kodu.
- [cv_data.json](file:///c:/Users/yusuf/Desktop/cv/cv_data.json): CV içeriğini tutan şablon veri dosyası (varsayılan olarak Yusuf Talha Alisan CV verilerini içerir).
- [.agent/rules/cv-olusturucu.md](file:///c:/Users/yusuf/Desktop/cv/.agent/rules/cv-olusturucu.md): AI asistanın seninle veya başka kullanıcılarla konuşarak CV hazırlamasını sağlayan otomatik kural dosyası.
