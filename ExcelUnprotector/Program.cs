using System;
using System.Windows.Forms;

namespace ExcelUnprotector;

internal static class Program
{
    /// <summary>
    /// Uygulamanin ana giris noktasi.
    /// Komut satiri argumani olarak dosya yolu(lari) verilirse konsol modunda calisir
    /// (orn. dosyalari Explorer'dan surukleyip .exe uzerine birakmak, ya da toplu betikten
    /// cagirmak icin). Argument verilmezse splash screen'i gosterip normal WinForms arayuzu acilir.
    /// </summary>
    [STAThread]
    private static void Main(string[] args)
    {
        if (args.Length > 0)
        {
            ConsoleRunner.Run(args);
            return;
        }

        ApplicationConfiguration();

        // Splash screen'i goster ve uygulamayi baslat
        using var splash = new SplashScreen();
        splash.ShowWithDelay(2500);

        Application.Run(new MainForm());
    }

    private static void ApplicationConfiguration()
    {
        // Not: Application.SetHighDpiMode(...) .NET Core WinForms'a (net5.0+) ozgudur ve
        // klasik .NET Framework'te (net48) mevcut degildir. Framework'te yuksek DPI
        // farkindaligi bunun yerine App.config icindeki
        // <System.Windows.Forms.ApplicationConfigurationSection> ile ayarlanir
        // (bkz. App.config).
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
    }
}
