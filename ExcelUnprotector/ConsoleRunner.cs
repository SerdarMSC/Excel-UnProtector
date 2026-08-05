namespace ExcelUnprotector;

/// <summary>
/// Komut satiri modu: "ExcelUnprotector.exe dosya1.xlsm dosya2.xlsx ..." seklinde,
/// ya da bir veya birden fazla dosyayi .exe uzerine surukle-birak yaparak kullanilabilir.
/// Her dosya icin "<ad>_unprotected<uzanti>" ciktisi olusturulur, kaynaklara dokunulmaz.
/// </summary>
internal static class ConsoleRunner
{
    public static void Run(string[] paths)
    {
        AttachConsoleIfNeeded();

        Console.WriteLine("Excel Unprotector - toplu islem modu");
        Console.WriteLine(new string('-', 60));

        var successCount = 0;
        var failCount = 0;

        foreach (var path in paths)
        {
            Console.WriteLine($"Isleniyor: {path}");
            var result = ProtectionRemover.RemoveProtections(path);

            if (!result.Success)
            {
                failCount++;
                Console.WriteLine($"  HATA: {result.ErrorMessage}");
                continue;
            }

            successCount++;

            if (!result.AnyProtectionFound)
            {
                Console.WriteLine("  Bu dosyada kaldirilacak bir koruma bulunamadi (zaten korumasiz olabilir).");
            }
            else
            {
                if (result.WorkbookProtectionRemoved)
                    Console.WriteLine("  - Calisma kitabi (yapi) korumasi kaldirildi.");

                foreach (var sheet in result.WorksheetsUnprotected)
                    Console.WriteLine($"  - Sayfa korumasi kaldirildi: {sheet}");
            }

            if (result.VbaProjectPresent)
                Console.WriteLine("  Not: Dosyada VBA projesi var; VBA proje (kod goruntuleme) sifresi bu araçla kaldirilmaz.");

            Console.WriteLine($"  Cikti: {result.OutputPath}");
        }

        Console.WriteLine(new string('-', 60));
        Console.WriteLine($"Tamamlandi. Basarili: {successCount}, Hatali: {failCount}");
    }

    /// <summary>
    /// WinExe cikti tipiyle derlenen bir uygulama komut satirindan calistirildiginda konsola
    /// otomatik baglanmayabilir; bu, Windows'ta mevcut bir konsol penceresine ciktiyi gorunur
    /// kilmaya calisir. Konsol yoksa (orn. dogrudan cift tiklanip dosya suruklenmisse) sessizce
    /// gecilir, cunku zaten yeni bir pencere acilmasi beklenmez.
    /// </summary>
    private static void AttachConsoleIfNeeded()
    {
        try
        {
            NativeMethods.AttachConsole(NativeMethods.ATTACH_PARENT_PROCESS);
        }
        catch
        {
            // Platform desteklemiyorsa (orn. gelecekte baska bir OS hedeflenirse) sessizce yut.
        }
    }
}

internal static class NativeMethods
{
    internal const int ATTACH_PARENT_PROCESS = -1;

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    internal static extern bool AttachConsole(int dwProcessId);
}
