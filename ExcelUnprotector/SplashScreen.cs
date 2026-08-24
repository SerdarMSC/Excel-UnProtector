using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelUnprotector;

/// <summary>
/// Uygulama başlangıcında gösterilen splash screen.
/// "Unlock Excel Security Icon.png" görselini ve uygulama bilgisini gösterir.
/// </summary>
public class SplashScreen : Form
{
    private readonly PictureBox _pictureBox = new();
    private readonly Label _titleLabel = new();
    private readonly Label _descriptionLabel = new();
    private readonly Label _versionLabel = new();
    private readonly ProgressBar _progressBar = new();

    public SplashScreen()
    {
        FormBorderStyle = FormBorderStyle.None;
        BackColor = Color.White;
        Size = new Size(500, 350);
        StartPosition = FormStartPosition.CenterScreen;
        ShowInTaskbar = false;
        TopMost = true;
        DoubleBuffered = true;

        BuildLayout();
        LoadIcon();
    }

    private void BuildLayout()
    {
        // Icon/Image
        _pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
        _pictureBox.Location = new Point(50, 30);
        _pictureBox.Size = new Size(120, 120);
        _pictureBox.BackColor = Color.White;

        // Başlık
        _titleLabel.Text = "Excel Unprotector";
        _titleLabel.Font = new Font("Segoe UI", 18F, FontStyle.Bold);
        _titleLabel.AutoSize = true;
        _titleLabel.Location = new Point(190, 40);
        _titleLabel.ForeColor = Color.FromArgb(0, 120, 212);

        // Açıklama
        _descriptionLabel.Text = "Excel dosyalarındaki korumalar\nkaldıran masaüstü aracı";
        _descriptionLabel.Font = new Font("Segoe UI", 10F);
        _descriptionLabel.AutoSize = true;
        _descriptionLabel.Location = new Point(190, 75);
        _descriptionLabel.ForeColor = SystemColors.ControlText;

        // Sürüm
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        _versionLabel.Text = $"Sürüm {version?.ToString(3) ?? "1.0.0"}";
        _versionLabel.Font = new Font("Segoe UI", 9F);
        _versionLabel.AutoSize = true;
        _versionLabel.Location = new Point(190, 125);
        _versionLabel.ForeColor = SystemColors.GrayText;

        // İlerleme çubuğu
        _progressBar.Location = new Point(30, 200);
        _progressBar.Size = new Size(440, 20);
        _progressBar.Style = ProgressBarStyle.Marquee;
        _progressBar.MarqueeAnimationSpeed = 30;

        // Bilgi metni
        var infoLabel = new Label
        {
            Text = "Uygulama başlatılıyor...",
            Font = new Font("Segoe UI", 9F),
            AutoSize = true,
            Location = new Point(30, 235),
            ForeColor = SystemColors.GrayText
        };

        // Yazar bilgisi
        var authorLabel = new Label
        {
            Text = "© 2026 Serdar MSC",
            Font = new Font("Segoe UI", 8F),
            AutoSize = true,
            Location = new Point(30, 305),
            ForeColor = SystemColors.ControlDark
        };

        var githubLabel = new Label
        {
            Text = "github.com/SerdarMSC/Excel-UnProtector",
            Font = new Font("Segoe UI", 8F),
            AutoSize = true,
            Location = new Point(30, 325),
            ForeColor = Color.FromArgb(0, 120, 212)
        };

        Controls.Add(_pictureBox);
        Controls.Add(_titleLabel);
        Controls.Add(_descriptionLabel);
        Controls.Add(_versionLabel);
        Controls.Add(_progressBar);
        Controls.Add(infoLabel);
        Controls.Add(authorLabel);
        Controls.Add(githubLabel);

        // Sınır efekti
        Paint += (s, e) =>
        {
            using var pen = new Pen(Color.FromArgb(200, 200, 200), 1);
            e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        };
    }

    private void LoadIcon()
    {
        try
        {
            // Uygulamanın kurulu olduğu dizinde "Unlock Excel Security Icon.png" bul
            var exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var exeDirectory = System.IO.Path.GetDirectoryName(exePath) ?? ".";
            var iconPath = System.IO.Path.Combine(exeDirectory, "Unlock Excel Security Icon.png");

            if (System.IO.File.Exists(iconPath))
            {
                _pictureBox.Image = Image.FromFile(iconPath);
            }
            else
            {
                // Dosya bulunamazsa, fallback olarak uygulamanın ikonunu kullan
                _pictureBox.Image = Icon.ExtractAssociatedIcon(Application.ExecutablePath)?.ToBitmap();
            }
        }
        catch
        {
            // Hata oluşursa sessizce yut; UI bozulmasın
        }
    }

    /// <summary>
    /// Splash screen'i belirtilen süre boyunca gösterir, ardından kapatır.
    /// </summary>
    public void ShowWithDelay(int delayMilliseconds = 2000)
    {
        Show();
        
        var timer = new Timer { Interval = delayMilliseconds };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            timer.Dispose();
            Close();
        };
        timer.Start();

        Application.DoEvents();
    }
}
