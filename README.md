# Excel Unprotector

**Coder:** SerdarMSC — [github.com/SerdarMSC](https://github.com/SerdarMSC/) · serdarmsc@gmail.com

Excel `.xlsx` / `.xlsm` / `.xltx` / `.xltm` dosyalarındaki **çalışma kitabı (yapı) koruması**
ve **sayfa korumalarını** kaldıran, C# (.NET Framework 4.8, WinForms) ile yazılmış bir masaüstü
aracı. .NET Framework 4.8 bilinçli olarak seçildi: Windows 10/11'e zaten gömülü geldiğinden,
çalıştırmak için ekstra bir "runtime" kurulumu gerekmez — tek bir küçük `.exe` yeterlidir.

Bu araç, Excel dosyalarının aslında bir ZIP paketi (OOXML) olduğu ve korumaların paket içindeki
XML parçalarında (`xl/workbook.xml`, `xl/worksheets/sheetN.xml`) saklandığı gerçeğine dayanır.
Şifre ne olursa olsun — çünkü Excel şifreyi doğrudan saklamaz, yalnızca bir **hash** değeri
saklar ve "korumalı mı değil mi" bayrağına bakar — ilgili etiketleri XML'den silmek korumayı
kaldırmaya yeter. Dosyanın geri kalanı (formüller, biçimlendirme, VBA makroları, ActiveX
kontrolleri, resimler, grafikler, vs.) bayt bayt aynen korunur; sadece koruma etiketleri
değiştirilir.

## Neyi kaldırır, neyi kaldırmaz

| Koruma türü | Durum |
|---|---|
| Çalışma kitabı yapı koruması (`workbookProtection`, "Yapıyı Koru") | ✅ Kaldırılır |
| Sayfa koruması (`sheetProtection`, "Sayfayı Koru") — şifreli veya şifresiz | ✅ Kaldırılır |
| Korumalı hücre aralıkları (`protectedRange` / `protectedRanges`) | ✅ Kaldırılır |
| VBA proje şifresi ("Görüntülemeyi Koru" / VBA'yı şifreleme) | ❌ Kaldırılmaz — bu, `vbaProject.bin` içindeki ayrı bir OLE bileşik dosya biçiminde saklanır, XML değil. Araç bu dosyayı algılar ve kullanıcıyı uyarır, ancak dokunmaz. |
| Açma/Değiştirme şifresi (tüm dosyayı şifreleyen "Open Password") | ❌ Kaldırılmaz — bu durumda dosya standart bir ZIP olarak bile açılamaz (tamamen şifrelenmiş bir OLE konteynerdir), farklı bir yaklaşım gerekir. |

## Proje yapısı

```
ExcelUnprotector/
├── .github/
│   └── workflows/
│       └── build.yml           # GitHub Actions: windows-latest'te derler, .exe yayınlar
├── ExcelUnprotector.sln
└── ExcelUnprotector/
    ├── ExcelUnprotector.csproj
    ├── Program.cs               # Giriş noktası (GUI veya CLI modu)
    ├── MainForm.cs              # WinForms arayüzü (sürükle-bırak, dosya listesi, günlük)
    ├── ProtectionRemover.cs     # Asıl iş mantığı (ZIP/XML işleme) — bağımsız, test edilebilir
    └── ConsoleRunner.cs         # Komut satırı / toplu iş modu
```

Harici bir NuGet paketine ihtiyaç yoktur — yalnızca .NET Framework'ün kendi
`System.IO.Compression` ve `System.Text.RegularExpressions` kütüphaneleri kullanılır.

## GitHub Actions ile otomatik derleme

Proje, `.github/workflows/build.yml` içinde hazır bir GitHub Actions iş akışıyla gelir. Windows
gerektiren bir WinForms uygulaması olduğu için iş akışı `windows-latest` runner'ında çalışır ve
kendi bilgisayarınızda .NET SDK kurulu olmasına gerek kalmadan projeyi derler.

**Kurulum:**

1. Bu klasörün tamamını (kök dizinde `.github/` dahil) bir GitHub deposuna gönderin (push edin).
2. GitHub'da depo → **Actions** sekmesine gidin. "Build Excel Unprotector" iş akışı, `main` (veya
   `master`) dalına her push ve her pull request'te otomatik çalışır; isterseniz **"Run workflow"**
   ile elle de tetikleyebilirsiniz.

**Derlenen .exe'yi indirme:**

- Her çalıştırmanın sonunda, o çalıştırmanın sayfasındaki **Artifacts** bölümünden
  `ExcelUnprotector-net48` adlı .zip'i indirebilirsiniz; içinde tek başına, ekstra kurulum
  gerektirmeden çalışan `ExcelUnprotector.exe` bulunur (birkaç yüz KB).
- Kalıcı bir sürüm yayınlamak isterseniz bir etiket (tag) push edin:
  ```bash
  git tag v1.0.0
  git push origin v1.0.0
  ```
  Bu durumda iş akışı `.exe`'yi doğrudan bir **GitHub Release**'e ekler; böylece indirme linki
  depo → **Releases** sayfasında kalıcı olarak durur.

**İş akışının yaptıkları:** `dotnet restore` → `dotnet build -c Release` → derlenen
`ExcelUnprotector.exe`'yi artifact/Release olarak yükler. .NET Framework derlemeleri tamamen
framework-dependent (işletim sistemine gömülü GAC derlemelerini kullanır) olduğundan ayrı bir
"publish" veya self-contained adımına gerek yoktur — `dotnet build`'in ürettiği `.exe` doğrudan
dağıtılabilir durumdadır.

> Durum rozetini kendi reponuza eklemek isterseniz, README'nizin en üstüne şunu koyup
> `OWNER/REPO` kısmını kendi kullanıcı adınız/depo adınızla değiştirin:
> `![Build](https://github.com/OWNER/REPO/actions/workflows/build.yml/badge.svg)`

## Derleme ve çalıştırma (yerelde)

Gereksinim: Windows üzerinde **.NET SDK** (proje `net48`'i hedeflese de, derlemek için modern
`dotnet` CLI aracı kullanılır — SDK bunu bulur ve .NET Framework 4.8 targeting pack'iyle derler)
ya da Visual Studio 2022 ("Windows Forms uygulama geliştirme" iş yükü ile; .NET Framework 4.8
hedef paketinin de kurulu olduğundan emin olun).

**Visual Studio ile:** `ExcelUnprotector.sln` dosyasını açın, F5 ile çalıştırın.

**Komut satırından:**

```bash
cd ExcelUnprotector
dotnet build -c Release
```

Derlenen `.exe`, `ExcelUnprotector/bin/Release/net48/ExcelUnprotector.exe` yolunda oluşur ve
doğrudan çalıştırılabilir — ayrı bir "publish" adımına gerek yoktur.

> Farklı bir .NET Framework sürümü kullanmak isterseniz `ExcelUnprotector.csproj` içindeki
> `<TargetFramework>net48</TargetFramework>` satırını örn. `net472` olarak değiştirebilirsiniz.
> Modern .NET (Core) 5/6/8'e dönmek isterseniz aynı satırı `net8.0-windows` yapıp
> `<Reference Include="System.IO.Compression" />` satırını kaldırmanız yeterli (o sürümlerde bu
> kütüphane zaten örtük olarak referanslanır) — ancak bu durumda hedef makinede .NET Desktop
> Runtime kurulu olması gerekir.

## Kullanım

**Arayüz (GUI) modu:** Uygulamayı argümansız çalıştırın. "Dosya Ekle…" ile dosya seçin ya da
pencereye sürükleyip bırakın, ardından **"Korumaları Kaldır"** düğmesine basın. Her dosya için,
aynı klasörde `<ad>_unprotected<uzantı>` adıyla yeni bir dosya oluşturulur; kaynak dosyaya asla
yazılmaz.

**Komut satırı modu:** Dosya yolu argüman olarak verildiğinde GUI açılmaz, sonuçlar konsola
yazılır:

```bash
ExcelUnprotector.exe "C:\rapor\bütçe.xlsm" "C:\rapor\model.xlsx"
```

Bu, dosyaları doğrudan `.exe` simgesinin üzerine sürükleyip bırakarak da tetiklenebilir.

## Nasıl çalışıyor (özet)

1. Kaynak dosya bir ZIP arşivi olarak açılır (değiştirilmez, salt okunur).
2. Yeni bir ZIP arşivi oluşturulur; kaynaktaki her giriş sırayla işlenir:
   - `xl/workbook.xml` ise, içindeki `<workbookProtection .../>` etiketi metin düzeyinde silinir.
   - `xl/worksheets/sheetN.xml` ise, `<sheetProtection .../>` ve `<protectedRange(s)>` etiketleri
     silinir.
   - Diğer tüm girişler (formüller, stiller, VBA, medya, ilişkiler, vb.) baytları değişmeden
     doğrudan kopyalanır.
3. Sonuç, geçici bir dosyaya yazılıp başarıyla tamamlandığında hedef dosya adına taşınır — yarım
   kalmış/bozuk bir çıktı dosyası bırakılmaz.

## Sınırlamalar

- `.xlsb` (ikili Excel biçimi) desteklenmez; bu format XML değil ikili kayıtlar kullanır.
- VBA proje şifresi ve dosya açma şifresi yukarıda açıklandığı gibi kapsam dışıdır.
- Bu araç yalnızca *kilit/koruma bayraklarını* kaldırır; şifreyi "kırmaz" veya çözmez — zaten
  Excel'in koruma mekanizması şifreyi saklamadığı için buna gerek de yoktur.

## Hakkında

**Coder:** SerdarMSC
**GitHub:** [https://github.com/SerdarMSC/](https://github.com/SerdarMSC/)
**E-posta:** serdarmsc@gmail.com

Uygulama içinde de bu bilgilere **"Hakkında"** düğmesinden ulaşılabilir.
