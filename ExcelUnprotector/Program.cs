namespace ExcelUnprotector;

internal static class Program
{
    /// <summary>
    /// Uygulamanin ana giris noktasi.
    /// Komut satiri argumani olarak dosya yolu(lari) verilirse konsol modunda calisir
    /// (orn. dosyalari Explorer'dan surukleyip .exe uzerine birakmak, ya da toplu betikten
    /// cagirmak icin). Argument verilmezse normal WinForms arayuzu acilir.
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
        Application.Run(new MainForm());
    }

    private static void ApplicationConfiguration()
    {
        Application.SetHighDpiMode(HighDpiMode.SystemAware);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
    }
}
