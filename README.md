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
            <b>[TR]</b> Bu arşivdeki uygulamalar (Python tabanlı Şifreleme Aracı hariç) <b>Visual Basic 6 (VB6)</b> tabanlıdır. Günümüz modern antivirüs motorları, VB6'nın kullandığı ActiveX/DLL kütüphane yapılarını "eski teknoloji" kategorisinde değerlendirmekte ve bazen bu dosyalara (False Positive) hatalı bir ön kabulle şüpheli etiketi yapıştırabilmektedir. Bu durum, tamamen yazılımın yaşı ve kütüphane kayıt yöntemleriyle (Self-Registration) ilgilidir.
            <br><br>
            Şeffaflık ve güven için:
            <ul>
                <li>Modern sistemlerde stabil çalışma ve DLL çakışmalarını önlemek için <b>Setup (EXE)</b> paketleri tercih edilmiştir.</li>
                <li>Her yayının altında tam <b>Kaynak Kodları (Source Code)</b> açıkça sunulmuştur.</li>
                <li>VB6 tabanlı uygulamalar açılışta <b>yönetici yetkisi</b> ister (Windows UAC onay penceresi görünür); bu, uygulamaların VB6 çekirdeğinden gelen bir özelliğidir.</li>
                <li>Uygulamalar dijital imzalı olmadığından Windows "Bilinmeyen yayımcı" uyarısı gösterebilir.</li>
                <li>VB6 çalışma zamanı bileşenleri (OCX/DLL), kurulum sırasında sistem klasörüne kopyalanıp kaydedilir ve kaldırma sırasında silinmez.</li>
            </ul>
            <hr>
            <b>[EN]</b> The applications in this archive (except the Python-based Encryption Tool) are <b>Visual Basic 6 (VB6)</b> based. Modern antivirus engines often flag VB6-specific ActiveX/DLL structures as suspicious (False Positive) due to the legacy nature of the technology and its self-registration methods.
            <br><br>
            For transparency and security:
            <ul>
                <li><b>Setup (EXE)</b> packages are provided to ensure library registration and system stability.</li>
                <li>Full <b>Source Code</b> is included with every release.</li>
                <li>The VB6-based applications request <b>administrator privileges</b> at startup (a Windows UAC prompt appears); this is a characteristic of the VB6 core executables.</li>
                <li>The applications are not digitally signed, so Windows may show an "Unknown publisher" warning.</li>
                <li>VB6 runtime components (OCX/DLL) are copied to the system folder and registered during setup; they are not removed during uninstall.</li>
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
* **RoboCopy Dahil:** Kurulum paketi mS-RoboCopy'yi içerir; RoboCopy sağ tık menüsü kurulumda isteğe bağlı bir bileşendir.
* **Sistem Bakım Araçları:** Kayıt Defteri (Registry) düzenleyici, sistem kilitlerini açma ve ActiveX/DLL kütüphane yönetim yardımcıları.
* **Kapsamlı Altyapı:** Onlarca modül ve yüzlerce formdan oluşan, VB6'nın sınırlarını zorlayan modüler mimari.
* **Hazır Kurulum Paketi:** Gerekli tüm sistem bileşenlerinin hatasız kaydedilmesi için hazırlanmış profesyonel kurulum paketi.

#### 🚀 mS_Explorer v2a (Öncü / Gelişmiş Sürüm - Beta) Yenilikleri:
* **Özel Filtreleme Sistemi:** İçerik listeleme bölümüne, verilere çok daha hızlı ulaşmanızı sağlayacak **Özel ComboBox filtreleme özelliği** eklendi.
* **Gelişmiş Dosya Doğrulama:** Toplu entegrasyon ve veri bütünlüğü takipleri için **Toplu SHA (Hash) Tarama modülü** sisteme dahil edildi.
* **Görsel İyileştirmeler:** Uygulama içi ikon setleri optimize edilerek modern ve daha net bir arayüz görünümü sağlandı.
* *Not: Bu öncü sürüm, itibar (reputation) süreci tamamlandıktan sonra ilerleyen aylarda doğrudan ana mS_Explorer.exe dosyasının yerini alacaktır.*

#### 🛠️ İndirme ve Kaynak Kod:
| Dosya / Bilgi | Açıklama | Bağlantı |
| :--- | :--- | :--- |
| 📦 **mS-Explorer-Kur (EXE)** | **Windows Kurulum Paketi (Kararlı Sürüm)** | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Explorer/mS-Explorer-Kur.exe) |
| 🛡️ **VirusTotal Raporu (Kararlı)** | VirusTotal Tarama Raporu | [Görüntüle](https://www.virustotal.com/gui/file/902c24163e2736b3246db4e7989cf300416874714f903b922fd138086f1adbb3/detection) |
| 🛡️ **Microsoft Raporu** | Kararlı sürüm için Microsoft Defender tarama raporu (PDF) | [Görüntüle-İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Explorer/microsoft-defender-report-ms-explorer-2026-08-EN.pdf) |
| 📂 **Kaynak Kod** | Tam Çalışma Ortamı (Zip) | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Explorer/mS_Explorer_Full_Workspace_v1.0.zip) |
| 📦 **mS-Explorer-Kur-Beta (EXE)** | Windows Kurulum Paketi (Beta, daha yeni, henüz Defender raporu yok) | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Explorer/mS-Explorer-Kur-Beta.exe) |

> 💡 **Güvenlik Notu:** Kararlı sürüm için Microsoft Defender tarama raporu yukarıda PDF olarak sunulmuştur. Sertifikasız, VB6 tabanlı freeware yazılımlarda bazı antivirüs motorları yanlış pozitif (False-Positive) verebilir; güncel sonuç için VirusTotal bağlantısına bakınız. İndirdiğiniz dosyanın SHA-256 değeri VirusTotal sayfasındakiyle aynıysa aynı dosyadır: `902c24163e2736b3246db4e7989cf300416874714f903b922fd138086f1adbb3`

<hr>

<h4 id="sifreleme">🔑 mS Şifreleme ve Dosya Araçları (Sifreleme-SHA-FileList-SetUp.exe) ⚠️</h4>
<p>Python tabanlı Şifreleme Aracı (.msfr ilişkilendirmesi) ile VB6 tabanlı kardeş modüllerden oluşan bağımsız kurulum paketidir: SHA-256 Dosya Doğrulama, SHA-256 Klasör Toplu Tarama (dizin içerisindeki tüm dosyaların benzersiz SHA-256 değerlerini topluca hesaplar) ve Esnek Dosya/Klasör İçerik Listeleme (File List Joker). Araçlar Gezgin sağ tık menüsünden kullanılabilir; sağ tık modülleri kurulum sırasında seçilebilir.</p>

<div align="center">
    <img src="https://raw.githubusercontent.com/alikurtnet/VB6-Nostalji-Arsivi/main/images/explorer/Sifreleme-SHA-FileList.png" alt="Sifreleme SHA FileList Arayüzü" width="60%">
</div>

| Dosya / Bilgi | Açıklama | Bağlantı |
| :--- | :--- | :--- |
| 📂 **Kaynak Kod** | Şifreleme Aracı (Python) Kaynak Kodu (Zip) | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Explorer/SifrelemeApp-Source-v1.0.zip) |
| ⚡ **Sifreleme-SHA-FileList (EXE)** | Windows Kurulum Paketi (Kararlı Sürüm) | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Explorer/Sifreleme-SHA-FileList-SetUp.exe) |
| 🛡️ **VirusTotal Raporu (Kararlı)** | VirusTotal Tarama Raporu | [Görüntüle](https://www.virustotal.com/gui/file/0cccef603066222ef6f5984467aafc3bc101d2793408fa18da6602ced888bd25/detection) |
| 🛡️ **Microsoft Raporu** | Microsoft Defender tarama raporu (PDF, 2026-08) | [Görüntüle-İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Explorer/microsoft-defender-report-ms-sifrele-2026-08.pdf) |
| 📦 **Sifreleme-SHA-FileList-Beta (EXE)** | Windows Kurulum Paketi (Beta, daha yeni sürüm) | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Explorer/Sifreleme-SHA-FileList-SetUp-Beta.exe) |

> ℹ️ Bu paket, mS-Explorer ile aynı kurulum dizinini paylaşır; iki paket birbirinden bağımsız kurulup kaldırılabilir, ortak dosyalar diğer paket kuruluyken silinmez.

> ⚠️ **Güvenlik Notu:** Bu modül, yakın zamanda yapılan derleme güncellemesi nedeniyle
> bazı bulut/itibar tabanlı AV motorları tarafından yanlış pozitif (False-Positive)
> olarak işaretlenebilmektedir. Yukarıdaki **VirusTotal Tarama Raporu**'nda 68 motordan
> 2'si uyarı vermiş, 66'sı vermemiştir; güncel sonuç için bağlantıya bakınız.
>
> Windows Defender uyarı gösterirse, dışlama eklemeden önce indirdiğiniz dosyanın
> SHA-256 değerinin VirusTotal sayfasındakiyle aynı olduğunu doğrulayın:
> `0cccef603066222ef6f5984467aafc3bc101d2793408fa18da6602ced888bd25`
> Ardından **Virüs ve tehdit koruması → Ayarları yönet → Dışlamalar** yolunu izleyerek
> ilgili dizini/dosyayı güvenilir listesine ekleyebilirsiniz.
>
> İleriki aşamalarda ilgili motorlara resmi temizlik (false-positive) başvurusu
> yapılması planlanmaktadır. Bu süreçle ilgili örnek olarak: arşivdeki
> **Game-SetUp.exe** için daha önce Microsoft'tan temiz rapor alınmış olup
> ilgili PDF aşağıdaki oyun bölümünde paylaşılmıştır; benzer şekilde
> mS_Explorer'ın eski sürümü **mS-Explorer-Kur-Clean.exe** için de VirusTotal'da
> 0/67 oranında temiz bir tarama sonucu elde edilmişti (SHA-256:
> `22477bb52a0594af4da6445e7d8b8d60e9b296a82f34e572d14c34d9fa20d67a`) — ancak bu
> dosya o dönemde Microsoft'un kendi kayıt sistemine resmi olarak işlenememişti.
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
    * **Zincirleme Harfler:** Komşu harflerle kelime türetme ve puan katlama (Zekâ & Şans).
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
| 💿 **Game-SetUp (EXE)** | Oyun Kurulum Paketi | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Game/Game-SetUp.exe) |
| 🛡️ **Güvenlik** | VirusTotal Tarama Raporu | [Görüntüle](https://www.virustotal.com/gui/file/168d3f1b65c406a93595960e591bdeafd6b72aa1b0e6c3583ffa9de80c0008dc/detection) |
| 🛡️ **Microsoft Raporu** | Microsoft Defender Temiz Raporu (pdf) | [Görüntüle-İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0-Game/microsoft-defender-clean-report-EN.pdf) |
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
* **Hızlı Görev Yönetimi:** Kaynak ve hedef klasör tanımlamalarını hafızada tutarak tek tuşla senkronizasyon sağlama.
* **Şeffaf Altyapı:** Windows'un alt kabuk ve kayıt mekanizmalarıyla uyumlu çalışan, sade ve açık kaynak kodlu yapı.
* **Gezgin Entegrasyonu:** Klasör, sürücü ve klasör arka planı sağ tık menülerine "RoboCopy: … Yedekle (Kaynak)" komutu eklenir; komut uygulamayı seçilen yol kaynak olarak ayarlanmış şekilde açar, kopyalamayı kendiliğinden başlatmaz.
* **Yalın Tasarım:** Gereksiz hiçbir görsel yük barındırmayan, doğrudan performansa ve amaca odaklı VB6 arabirimi.
* **Eğitim Odaklı Açık Kaynak:** Kodların sadeleştirilmiş ve budanmış mimarisi sayesinde, kütüphane yönetimini anlamak isteyen öğrenciler için kusursuz bir mehaz (referans).

#### 🛠️ İndirme ve Kaynak Kod:
| Dosya / Bilgi | Açıklama | Bağlantı |
| :--- | :--- | :--- |
| 💿 **mS-RoboCopy-SetUp (EXE)** | **Windows Kurulum Paketi (Kararlı Sürüm), bağımsız kurulur; Gezgin sağ tık menülü | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0/mS-RoboCopy-SetUp.exe) |
| 🛡️ **VirusTotal Raporu (Kararlı)** | VirusTotal Tarama Sonucu (0/71, Eylül 2026) | [Görüntüle](https://www.virustotal.com/gui/file/598e3210fdeffbfc854f014b8cd51376cef61ef1bff04e470fd70c2a907a3990/detection) |
| 📂 **Kaynak Kod** | Açık Kaynak Kod Dünyası (Zip) | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0/RoboCopy-OpenSource-v1.0.0.zip) |
| 📦 **mS-RoboCopy-SetUp-Beta (EXE)** | Windows Kurulum Paketi (Beta, daha yeni sürüm; VirusTotal raporu henüz yok) | [İndir](https://github.com/alikurtnet/VB6-Nostalji-Arsivi/releases/download/v1.0.0/mS-RoboCopy-SetUp-Beta.exe) |

> 🔎 Bu sürüm için VirusTotal taramasında 71 güvenlik motorunun hiçbiri dosyayı işaretlememiştir. Sonuç, taramanın yapıldığı tarihe aittir; güncel durum için bağlantıya bakınız.

> ℹ️ mS-Explorer paketi RoboCopy'yi zaten içerir; bu bağımsız paket yalnızca RoboCopy isteyenler içindir. İki paket ayrı dizinlere kurulur ve birbirinden bağımsız çalışır. mS-RoboCopy kuruluysa mS-Explorer kurulumu RoboCopy sağ tık menüsünü otomatik devre dışı bırakır; mS-Explorer kuruluyken bağımsız RoboCopy kurulursa kurulum sizi uyarır, devam ederseniz Gezgin menüsünde RoboCopy komutu iki kez görünür.

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
