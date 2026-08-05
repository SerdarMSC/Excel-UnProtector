# Excel Unprotector

**Coder:** SerdarMSC — [github.com/SerdarMSC](https://github.com/SerdarMSC/) · serdarmsc@gmail.com

Excel `.xlsx` / `.xlsm` / `.xltx` / `.xltm` dosyalarındaki **çalışma kitabı (yapı) koruması**
ve **sayfa korumalarını** kaldıran, C# (.NET 8, WinForms) ile yazılmış bir masaüstü aracı.

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

Harici bir NuGet paketine ihtiyaç yoktur — yalnızca .NET'in kendi `System.IO.Compression` ve
`System.Text.RegularExpressions` kütüphaneleri kullanılır.

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

Her çalıştırma **iki farklı çıktı** üretir; ihtiyacınıza göre birini seçin:

| Artifact / Release dosyası | Boyut | Gereksinim |
|---|---|---|
| `ExcelUnprotector-win-x64-slim` (`ExcelUnprotector-slim.exe`) | ~birkaç yüz KB | Hedef makinede **[.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)** kurulu olmalı (yoksa Windows ilk açılışta otomatik kurulum sayfasına yönlendirir) |
| `ExcelUnprotector-win-x64-portable` (`ExcelUnprotector-portable.exe`) | ~60–100 MB | Hiçbir kurulum gerektirmez, tek başına çalışır (.NET çalışma zamanı exe'nin içine gömülü) |

Çoğu güncel Windows 10/11 makinesinde .NET Desktop Runtime zaten kurulu ya da Windows Update
üzerinden kolayca kurulabilir olduğundan, günlük kullanım için **`slim`** sürümünü önerir;
başka bir bilgisayara USB ile taşımak gibi "kur gerektirmesin" senaryolarında **`portable`**
sürümünü kullanın.

- Her çalıştırmanın sonunda, o çalıştırmanın sayfasındaki **Artifacts** bölümünden ikisini de
  indirebilirsiniz.
- Kalıcı bir sürüm yayınlamak isterseniz bir etiket (tag) push edin:
  ```bash
  git tag v1.0.0
  git push origin v1.0.0
  ```
  Bu durumda iş akışı her iki `.exe`'yi de doğrudan bir **GitHub Release**'e ekler; böylece
  indirme linki depo → **Releases** sayfasında kalıcı olarak durur.

**İş akışının yaptıkları:** `dotnet restore` → `dotnet build -c Release` (derleme hatalarını
erken yakalamak için) → iki ayrı `dotnet publish` çağrısı ile hem küçük (framework-dependent)
hem de taşınabilir (self-contained + sıkıştırılmış tek dosya) `.exe` üretir.

> Durum rozetini kendi reponuza eklemek isterseniz, README'nizin en üstüne şunu koyup
> `OWNER/REPO` kısmını kendi kullanıcı adınız/depo adınızla değiştirin:
> `![Build](https://github.com/OWNER/REPO/actions/workflows/build.yml/badge.svg)`

## Derleme ve çalıştırma (yerelde)

Gereksinim: Windows üzerinde **.NET 8 SDK** (ya da Visual Studio 2022, "Windows Forms uygulama
geliştirme" iş yükü ile).

**Visual Studio ile:** `ExcelUnprotector.sln` dosyasını açın, F5 ile çalıştırın.

**Komut satırından:**

```bash
cd ExcelUnprotector
dotnet build
dotnet run --project ExcelUnprotector
```

Yayınlanabilir tek dosya `.exe` üretmek için:

```bash
dotnet publish ExcelUnprotector -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

> Farklı bir .NET sürümü kullanmak isterseniz `ExcelUnprotector.csproj` içindeki
> `<TargetFramework>net8.0-windows</TargetFramework>` satırını örn. `net6.0-windows` ya da
> (Windows Forms destekleyen) `net472` olarak değiştirebilirsiniz.

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
