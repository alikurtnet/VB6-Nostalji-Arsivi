<h1 align="center">🏛️ VB6 Nostalji Arşivi</h1>
<blockquote align="center"><strong>30 Yıllık Mühendislik ve Yazılım Serüveni</strong></blockquote>

<div align="center">
    <img src="https://img.shields.io/badge/Language-VB6_/_VBA-blue" alt="Language">
    <img src="https://img.shields.io/badge/Database-MS_Access-red" alt="Database">
    <img src="https://img.shields.io/badge/Status-Active_Archive-brightgreen" alt="Status">
</div>

<p align="center">
Bu depo, teknik eğitimci ve yazılımcı kimliğimle geliştirdiğim profesyonel araçların koleksiyonudur. 15 Mart 2026 itibarıyla tüm çalışma ortamları ve kaynak kodları geleceğe miras olarak dökümante edilmiştir.
</p>

<hr>

<h2>🛡️ Güvenlik ve Kurulum Standartları / Security & Installation Standards</h2>

<table bgcolor="#fff3cd">
    <tr>
        <td>
            ⚠️ <b>ÖNEMLİ (Teknoloji ve Güvenlik Notu):</b> 
            <br><br>
            <b>[TR]</b> Bu arşivdeki uygulamalar <b>Visual Basic 6 (VB6)</b> tabanlıdır. Günümüz modern antivirüs motorları, VB6'nın kullandığı ActiveX/DLL kütüphane yapılarını "eski teknoloji" kategorisinde değerlendirmekte ve bazen bu dosyalara (False Positive) hatalı bir ön kabulle şüpheli etiketi yapıştırabilmektedir. Bu durum, tamamen yazılımın yaşı ve kütüphane kayıt yöntemleriyle (Self-Registration) ilgilidir.
            <br><br>
            Şeffaflık ve güven için:
            <ul>
                <li>Modern sistemlerde stabil çalışma ve DLL çakışmalarını önlemek için <b>Setup (EXE)</b> paketleri tercih edilmiştir.</li>
                <li>Her yayının altında tam <b>Kaynak Kodları (Source Code)</b> açıkça sunulmuştur.</li>
                <li>Uygulamaları <b>Yönetici Olarak</b> çalıştırmanız, eski kütüphanelerin Windows çekirdeğine güvenli bir şekilde tanıtılması için önerilir.</li>
            </ul>
            <hr>
            <b>[EN]</b> The applications in this archive are <b>Visual Basic 6 (VB6)</b> based. Modern antivirus engines often flag VB6-specific ActiveX/DLL structures as suspicious (False Positive) due to the legacy nature of the technology and its self-registration methods.
            <br><br>
            For transparency and security:
            <ul>
                <li><b>Setup (MSI/EXE)</b> packages are provided to ensure library registration and system stability.</li>
                <li>Full <b>Source Code</b> is included with every release.</li>
                <li>Running the apps as <b>Administrator</b> is recommended for proper registration of legacy components.</li>
            </ul>
        </td>
    </tr>
</table>

<hr>
<details open>
<summary><h3 id="calc">🧮 1. ms_Calc - Mühendislik ve Matematik Çözümleyici ↕️ </h3></summary>

> **Nostalji Serisi: No 1**

<p>Matematiksel ifadeleri "eval" mantığıyla çözümleyen mS_Calc, birim çevrimlerinden karmaşık mekanik hesaplamalara kadar geniş bir yelpazede hizmet veren profesyonel bir yardımcı araçtır. Mühendislik hesaplamalarını sesli geri bildirimle birleştiren bu modül, teknik dökümantasyon hassasiyetinde sonuçlar üretir.</p>

<div align="center">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/calc/4_Calc_Full_Page.gif" alt="mS_Calc Önizleme" width="100%">
</div>

#### ✨ Kapsamlı Hesaplama Yetenekleri:
* **Mekanik ve Talaşlı İmalat:** Düz ve Helis dişli çark eleman hesapları, Talaşlı imalatta Kesme Hızı ve İlerleme miktarı analizleri.
* **Malzeme ve Geometri:** İçi dolu/boş malzemelerin ağırlık hesaplamaları, Koniklik ve Eğim tayini.
* **İleri Matematik:** Birinci dereceden iki bilinmeyenli ve ikinci dereceden bir bilinmeyenli denklem çözümleri; Permütasyon, Kombinasyon ve Olasılık hesapları.
* **Birim ve Tarih:** Kapsamlı birim çevrimleri; iki tarih arası fark bulma veya belirtilen tarihe gün ekleme/çıkartma gibi dinamik tarih işlemleri.
* **Seslendirme Desteği:** Hesaplama sonuçlarını yazı formatından sesli ifadeye dönüştürerek kullanıcıya raporlama özelliği.
* **Arayüz:** Tahoma fontu ile yenilenmiş, yüksek çözünürlük (DPI) uyumlu modern görünüm.

#### 🛠️ İndirme ve Kaynak Kod:
| Dosya / Bilgi | Açıklama | Bağlantı |
| :--- | :--- | :--- |
| 💿 **Hesap-Makinesi-Calc (EXE)** | Windows Installer Paketi | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-msCalc/Hesap-Makinesi-Calc.exe) |
| 🛡️ **Güvenlik** | VirusTotal Tarama Raporu | [Görüntüle](https://www.virustotal.com/gui/file/19aba22263cd63c2719a98547bfbef3ba3e6a42534e5f4a142cb4bd509cea460/detection) |
| 📂 **Kaynak Kod** | VB6 Proje Dosyaları (Zip) | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-msCalc/Eval-Calculator-Full_Workspace_v1.0.zip) |

</details>

<details open>
<summary><h3 id="explorer">📂 2. mS_Explorer - Dosya Yönetim ve Sistem Merkezi ↕️ </h3></summary>

> **Nostalji Serisi: No 2**

<p>Windows Explorer'a güçlü bir alternatif olarak geliştirilen mS_Explorer; dosya indeksleme, gelişmiş arama ve profesyonel yedekleme araçlarını tek bir merkezde toplar. 58 alt klasör ve 1100'den fazla dosyadan oluşan bu devasa çalışma alanı, sadece bir dosya yöneticisi değil, aynı zamanda kapsamlı bir sistem bakım kitidir.</p>

<div align="center">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/explorer/1_mS_Explorer.png" alt="mS_Explorer Arayüz" width="49%">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/explorer/2_RoboCopy.png" alt="mS_RoboCopy Modülü" width="49%">
</div>

#### ✨ Öne Çıkan Özellikler:
* **Akıllı Dosya Yönetimi:** Windows Gezgini mantığında hızlı erişim, kategorize edilmiş dosya indeksleme ve gelişmiş arama motoru.
* **mS_RoboCopy Arayüzü:** Karmaşık RoboCopy komutlarını görselleştiren, güvenli ve hızlı veri yedekleme modülü.
* **Sistem Bakım Araçları:** Kayıt Defteri (Registry) düzenleyici, sistem kilitlerini açma ve ActiveX/DLL kütüphane yönetim yardımcıları.
* **Kapsamlı Altyapı:** Onlarca modül ve yüzlerce formdan oluşan, VB6'nın sınırlarını zorlayan modüler mimari.
* **Güvenli Dağıtım:** Gerekli tüm sistem bileşenlerinin hatasız kaydedilmesi için profesyonel kurulum paketi.

#### 🚀 mS_Explorer v2a (Öncü / Gelişmiş Sürüm) Yenilikleri:
* **Özel Filtreleme Sistemi:** İçerik listeleme bölümüne, verilere çok daha hızlı ulaşmanızı sağlayacak **Özel ComboBox filtreleme özelliği** eklendi.
* **Gelişmiş Dosya Doğrulama:** Toplu entegrasyon ve veri bütünlüğü takipleri için **Toplu SHA (Hash) Tarama modülü** sisteme dahil edildi.
* **Görsel İyileştirmeler:** Uygulama içi ikon setleri optimize edilerek modern ve daha net bir arayüz görünümü sağlandı.
* *Not: Bu öncü sürüm, itibar (reputation) süreci tamamlandıktan sonra ilerleyen aylarda doğrudan ana mS_Explorer.exe dosyasının yerini alacaktır.*

#### 🛠️ İndirme ve Kaynak Kod:
| Dosya / Bilgi | Açıklama | Bağlantı |
| :--- | :--- | :--- |
| 📦 **mS-Explorer-Kur (EXE)** | **[Kararlı Sürüm]** Orijinal Windows Kurulum Paketi | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Explorer/mS-Explorer-Kur.exe) |
| 🛡️ **Güvenlik (v1.0.0)** | VirusTotal Tarama Raporu | [Görüntüle](https://www.virustotal.com/gui/file/a180d4966ad544f3cca5b2fca0c7d82fdcd6d3e57b889ba216b6c00189d2aabf/detection) |
| 🚀 **mS_Explorer_v2a (EXE)** | **[Öncü Sürüm]** En yeni özellikleri içeren test derlemesi *(AV İtibar Analizindedir)* | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Explorer/mS_Explorer_v2a.exe) |
| 📂 **Kaynak Kod** | Tam Çalışma Ortamı (Zip) | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Explorer/mS_Explorer_Full_Workspace_v1.0.zip) |

> 💡 **v2a Güvenlik Notu:** Bu yeni derleme, antivirüs motorlarının bulut/itibar (cloud reputation) davranışlarını gözlemlemek amacıyla TotalVirus gibi platformlara hemen yansıtılmamıştır. Geçmiş kararlı sürümlerimizin güvenilirliği ve AV devlerinin (Microsoft Defender vb.) onayları ışığında dosyamız tamamen güvenlidir. İlk etapta (yeni derleme olmasından ötürü) en fazla 1-2 yerel AV motorunun hatalı (False-Positive) uyarı verebileceği öngörülmektedir, bu durum tamamen geçicidir.

</details>

<details open>
<summary><h3 id="game">🎮 3. ms_Game - Düşün, Oyna, Öğren: Eğitici Oyunlar ↕️ </h3></summary>

> **Nostalji Serisi: No 3**

<p>VB6 ile geliştirilmiş bu koleksiyon; zekâ, hafıza ve hızlı düşünme yetisini geliştirmeye odaklanan modüler bir oyun arşividir. Algoritma mantığını nostaljik bir arayüzle sunan bu paket, hem eğlendirir hem de eğitir.</p>

#### 📸 Oyun Arayüzleri
<div align="center">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/game/00_Game.png" width="24%">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/game/01_Tahmin.png" width="24%">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/game/02_Bil.png" width="24%">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/game/03_AdamAsmaca.png" width="24%">
    <br><br>
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/game/04_Puzzle.png" width="24%">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/game/05_Tenis.png" width="24%">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/game/06_Zingir.png" width="24%">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/game/07_Math.png" width="24%">
    <br><br>
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/game/08_UcTas.png" width="24%">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/game/09_Hafiza.png" width="24%">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/game/10_ResimKelime.png" width="24%">
</div>

#### ✨ Arşivdeki Eğitici Modüller:
* **Zekâ ve Kelime Dağarcığı:**
    * **Adam Asmaca:** Türkçe/İngilizce sözlük desteği ve harf seslendirme özelliği.
    * **4 Resim 1 Kelime:** Görsel ve kavramsal bağ kurma, dil geliştirme.
    * **Zincirleme Harfler:** Komşu harflerle kelime türetme and puan katlama (Zekâ & Şans).
* **Matematik ve Mantık:**
    * **Hızlı Matematik:** Zamanla yarışarak doğru kavrama ve işlem yetisi kazanma.
    * **Sudoku & Puzzle:** Klasik mantık yürütme ve dikkat geliştirme bulmacaları.
    * **Tuttuğum Sayıyı Bil:** Bilgi, hafıza ve stratejik tahmin yürütme.
* **Hız ve Nostalji:**
    * **Rally & mini Ralliciler:** Hızlı refleks yönetimi.
    * **Tenis:** Görsel efektlerle desteklenmiş klasik arcade deneyimi.
    * **Üçtaş:** Strateji odaklı geleneksel zekâ oyunu.

#### 🛠️ İndirme ve Kaynak Kod:
| Dosya / Bilgi | Açıklama | Bağlantı |
| :--- | :--- | :--- |
| 💿 **Game-Game (EXE)** | Oyun Kurulum Paketi | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Game/Game-SetUp.exe) |
| 🛡️ **Güvenlik** | VirusTotal Tarama Raporu | [Görüntüle](https://www.virustotal.com/gui/file/81bfea350be1e342cba033b72e6a8538da1bf6601fdf10c54d61ddd16ab18a51/detection) |
| 📂 **Kaynak Kod** | Tüm Oyun Kaynakları (Zip) | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Game/mS-Game-Full_Workspace_v1.0.zip) |

</details>

<details open>
<summary><h3 id="robocopy">🚀 4. mS_RoboCopy - Gelişmiş Veri Yedekleme ve Senkronizasyon Arayüzü ↕️ </h3></summary>

> **Nostalji Serisi: No 4**

<p>Windows'un güçlü komut satırı aracı RoboCopy'yi tamamen görselleştiren, sade ve kararlı bir sistem yardımcı aracıdır. Karmaşık parametreleri tek bir tıklamaya indirgeyen bu modül, sistem yöneticileri, teknik eğiticiler ve verilerini gamsızca yedeklemek isteyenler için tasarlanmıştır.</p>

<div align="center">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/RoboCopy/mS_RoboCopy.png" alt="mS_RoboCopy Bağımsız Arayüz" width="60%">
</div>

#### ✨ Öne Çıkan Özellikler:
* **Hızlı Görev Yönetimi:** Kaynak ve hedon klasör tanımlamalarını hafızada tutarak tek tuşla senkronizasyon sağlama.
* **Antivirüs Dostu Çekirdek:** Windows'un alt kabuk ve kayıt mekanizmalarıyla tam uyumlu çalışan, tarama filtrelerine takılmayan şeffaf altyapı.
* **Yalın Tasarım:** Gereksiz hiçbir görsel yük barındırmayan, doğrudan performansa ve amaca odaklı VB6 arabirimi.
* **Eğitim Odaklı Açık Kaynak:** Kodların sadeleştirilmiş ve budanmış mimarisi sayesinde, kütüphane yönetimini anlamak isteyen öğrenciler için kusursuz bir mehaz (referans).

#### 🛠️ İndirme ve Kaynak Kod:
| Dosya / Bilgi | Açıklama | Bağlantı |
| :--- | :--- | :--- |
| 💿 **mS-RoboCopy-SetUp (EXE)** | Kurulum ve Entegrasyon Paketi | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0/mS-RoboCopy-SetUp.exe) 
| 🛡️ **Güvenlik** | VirusTotal Tarama Raporu | [Görüntüle](https://www.virustotal.com/gui/file/9fbc7210477c7eb629c158df4ef30db038bfbf1a93e68309dd6f20a45751adab/detection) |
| 📂 **Kaynak Kod** | Açık Kaynak Kod Dünyası (Zip) | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0/RoboCopy-OpenSource-v1.0.0.zip) |

</details>

<hr>

<h2>📜 Geliştirici Notu / Developer Notes</h2>

<p>
<b>[TR]</b> Bu arşiv, forum kültüründen ve yardımlaşma ruhundan beslenerek bugünlere gelmiştir. Her bir satır kodda bir teknik çözüm arayışı ve mühendislik emeği vardır. Bu kaynak kodlar; güncel teknolojilerle uygulama geliştirmek isteyenler için bir ufuk açıcı ve teknik bir <b>mehaz (referans)</b> olması düşüncesiyle paylaşılmıştır.
<br><br>
<b>[EN]</b> This archive has evolved through the spirit of collaboration and forum culture. Every line of code represents an engineering effort and a search for technical solutions. These source codes are shared with the intent of serving as an <b>inspiring reference (resource)</b> for those aiming to develop applications with modern technologies.
</p>

<blockquote>
  <p><b>Açık Kaynak Katkısı Hakkında:</b> Bu arşivdeki bazı uygulamaların ve oyunların temel iskeleti, internet üzerinde paylaşılan değerli açık kaynak kodlara dayanmaktadır. Söz konusu temeller; tarafımdan geliştirilen ilave kodlar, yeni fonksiyonlar ve özgün görsel arayüzlerle zenginleştirilerek profesyonel bir yapıya kavuşturulmuştur. Bilgi paylaşımına katkıda bulunan tüm küresel geliştiricilere teşekkür ederim.</p>
  
  <p><b>Open Source Credits:</b> The core frameworks of some applications and games in this archive are based on valuable open-source code shared globally. These foundations have been enhanced and professionalized through additional coding, new functionalities, and custom visual interfaces developed by me. I would like to express my gratitude to all developers worldwide for their contributions to the knowledge-sharing community.</p>
</blockquote>

<div align="center">
    <sub>© 2026 alikurtnet (Ali Kurt). Teknik Eğitimci & Yazılım Geliştirici.</sub>
</div>
