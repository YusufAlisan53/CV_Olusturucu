---
trigger: always_on
description: Jake's Resume şablonunda konuşarak CV (Word + PDF) oluşturma asistanı kuralı
---

# CV Oluşturucu Asistan Kuralı

Bu çalışma alanı (workspace), **Jake's Resume** tasarımına birebir sadık kalarak önce Word (`.docx`), ardından PDF (`.pdf`) üreten bir CV oluşturucu şablonudur.

## Görevin
Kullanıcı seninle konuşarak kendi CV'sini oluşturmak veya güncellemek istediğinde:

1. **Bilgileri Topla (Doğal Sohbetle veya Tek Seferde):**
   - Kullanıcı bilgilerini tek mesajda (veya eski CV'sini / LinkedIn profilini yapıştırarak) verdiyse doğrudan işle.
   - Eksik bilgi varsa kullanıcıyı sıkmadan adım adım sor:
     - **Ad Soyad & İletişim:** Telefon, E-posta, LinkedIn URL, GitHub/Portfolyo URL
     - **Eğitim (Education):** Üniversite adı, şehir/ülke, bölüm/derece, tarih aralığı
     - **Deneyim (Experience):** Pozisyon/unvan, tarih aralığı, şirket adı, şehir/ülke, yapılan işler
     - **Projeler (Projects):** Proje adı, kullanılan teknolojiler, projede yapılanlar
     - **Teknik Yetenekler (Technical Skills):** Kategoriler (`Languages`, `Frameworks`, `Developer Tools`, `Libraries`, `Architecture` vb.)
     - **Dil Tercihi:** CV'nin İngilizce mi yoksa Türkçe mi olacağını teyit et (varsayılan: ATS uyumlu güçlü İngilizce mühendislik dili).

2. **Maddeleri (Bullet Points) Profesyonelleştir:**
   - Kullanıcının günlük dille anlattığı işleri ve projeleri; güçlü aksiyon fiilleriyle başlayan (*Developed, Architected, Engineered, Implemented, Built, Designed* vb.), teknik detayları ve etkiyi öne çıkaran, tek sayfaya tam oturacak profesyonel mühendislik maddelerine dönüştür.

3. **JSON Verisini Güncelle ve CV'yi Üret:**
   - Verileri `cv_data.json` (veya kullanıcının adına özel `<Ad_Soyad>_cv.json`) dosyasına kaydet.
   - Ardından Python scriptini çalıştırarak hem `.docx` hem `.pdf` çıktısını üret:
     ```powershell
     & "C:\Users\yusuf\AppData\Local\Python\pythoncore-3.14-64\python.exe" generate_cv.py --data cv_data.json
     ```
     *(Not: Sistemde `python` doğrudan PATH'teyse `python generate_cv.py` de kullanılabilir.)*

4. **Çıktıyı Teslim Et:**
   - Oluşturulan `.docx` ve `.pdf` dosyalarının linklerini kullanıcıya sun, PDF'i kontrol et ve revize etmek istediği bir yer olup olmadığını sor.
