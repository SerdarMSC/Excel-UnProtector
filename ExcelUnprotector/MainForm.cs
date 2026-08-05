using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExcelUnprotector;

public class MainForm : Form
{
    private readonly ListBox _fileList = new();
    private readonly Button _btnAdd = new();
    private readonly Button _btnRemoveSelected = new();
    private readonly Button _btnClear = new();
    private readonly Button _btnProcess = new();
    private readonly Button _btnAbout = new();
    private readonly TextBox _log = new();
    private readonly Label _hint = new();
    private readonly Label _dropHint = new();

    public MainForm()
    {
        Text = "Excel Unprotector";
        Width = 820;
        Height = 620;
        MinimumSize = new Size(620, 460);
        StartPosition = FormStartPosition.CenterScreen;
        AllowDrop = true;
        Font = new Font("Segoe UI", 9F);

        BuildLayout();
        WireEvents();
    }

    private void BuildLayout()
    {
        _hint.Text = "İşlenecek .xlsx / .xlsm / .xltx / .xltm dosyalarını ekleyin (veya bu pencereye sürükleyin):";
        _hint.AutoSize = true;
        _hint.Location = new Point(12, 12);

        _fileList.Location = new Point(12, 36);
        _fileList.Size = new Size(560, 220);
        _fileList.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
        _fileList.SelectionMode = SelectionMode.MultiExtended;
        _fileList.HorizontalScrollbar = true;
        _fileList.AllowDrop = true;

        _dropHint.Text = "Dosyaları buraya sürükleyip bırakabilirsiniz";
        _dropHint.ForeColor = SystemColors.GrayText;
        _dropHint.AutoSize = true;

        _btnAdd.Text = "Dosya Ekle…";
        _btnAdd.Location = new Point(584, 36);
        _btnAdd.Size = new Size(210, 32);
        _btnAdd.Anchor = AnchorStyles.Top | AnchorStyles.Right;

        _btnRemoveSelected.Text = "Seçilileri Kaldır";
        _btnRemoveSelected.Location = new Point(584, 76);
        _btnRemoveSelected.Size = new Size(210, 32);
        _btnRemoveSelected.Anchor = AnchorStyles.Top | AnchorStyles.Right;

        _btnClear.Text = "Listeyi Temizle";
        _btnClear.Location = new Point(584, 116);
        _btnClear.Size = new Size(210, 32);
        _btnClear.Anchor = AnchorStyles.Top | AnchorStyles.Right;

        _btnProcess.Text = "Korumaları Kaldır";
        _btnProcess.Location = new Point(584, 168);
        _btnProcess.Size = new Size(210, 40);
        _btnProcess.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        _btnProcess.Font = new Font(Font, FontStyle.Bold);
        _btnProcess.BackColor = Color.FromArgb(0, 120, 212);
        _btnProcess.ForeColor = Color.White;
        _btnProcess.FlatStyle = FlatStyle.Flat;

        _btnAbout.Text = "Hakkında";
        _btnAbout.Location = new Point(584, 216);
        _btnAbout.Size = new Size(210, 32);
        _btnAbout.Anchor = AnchorStyles.Top | AnchorStyles.Right;

        var logLabel = new Label
        {
            Text = "Sonuç günlüğü:",
            AutoSize = true,
            Location = new Point(12, 266),
            Anchor = AnchorStyles.Top | AnchorStyles.Left
        };

        _log.Multiline = true;
        _log.ReadOnly = true;
        _log.ScrollBars = ScrollBars.Vertical;
        _log.Location = new Point(12, 290);
        _log.Size = new Size(782, 270);
        _log.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        _log.Font = new Font("Consolas", 9F);
        _log.BackColor = Color.White;

        Controls.Add(_hint);
        Controls.Add(_fileList);
        Controls.Add(_dropHint);
        Controls.Add(_btnAdd);
        Controls.Add(_btnRemoveSelected);
        Controls.Add(_btnClear);
        Controls.Add(_btnProcess);
        Controls.Add(_btnAbout);
        Controls.Add(logLabel);
        Controls.Add(_log);

        // dropHint'i listenin altina, layout kesinlestikten sonra yerlestir.
        Load += (_, _) => _dropHint.Location = new Point(_fileList.Left, _fileList.Bottom + 4);
    }

    private void WireEvents()
    {
        _btnAdd.Click += (_, _) => AddFilesViaDialog();
        _btnRemoveSelected.Click += (_, _) => RemoveSelected();
        _btnClear.Click += (_, _) => _fileList.Items.Clear();
        _btnProcess.Click += (_, _) => ProcessFiles();
        _btnAbout.Click += (_, _) => ShowAboutDialog();

        DragEnter += Form_DragEnter;
        DragDrop += Form_DragDrop;
        _fileList.DragEnter += Form_DragEnter;
        _fileList.DragDrop += Form_DragDrop;
    }

    private void AddFilesViaDialog()
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Excel dosyalarını seçin",
            Filter = "Excel Makro Dosyaları (*.xlsm;*.xltm)|*.xlsm;*.xltm|" +
                     "Tüm Excel Dosyaları (*.xlsx;*.xlsm;*.xltx;*.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|" +
                     "Tüm dosyalar (*.*)|*.*",
            Multiselect = true
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
            AddFiles(dialog.FileNames);
    }

    private void AddFiles(IEnumerable<string> paths)
    {
        foreach (var path in paths)
        {
            var extension = Path.GetExtension(path);
            if (!ProtectionRemover.SupportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
                continue;

            var alreadyThere = _fileList.Items.Cast<string>()
                .Any(existing => string.Equals(existing, path, StringComparison.OrdinalIgnoreCase));

            if (!alreadyThere)
                _fileList.Items.Add(path);
        }
    }

    private void RemoveSelected()
    {
        foreach (var item in _fileList.SelectedItems.Cast<string>().ToList())
            _fileList.Items.Remove(item);
    }

    private void Form_DragEnter(object? sender, DragEventArgs e)
    {
        e.Effect = e.Data?.GetDataPresent(DataFormats.FileDrop) == true
            ? DragDropEffects.Copy
            : DragDropEffects.None;
    }

    private void Form_DragDrop(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetData(DataFormats.FileDrop) is string[] files)
            AddFiles(files);
    }

    private void ProcessFiles()
    {
        if (_fileList.Items.Count == 0)
        {
            MessageBox.Show(this, "Önce en az bir dosya ekleyin.", "Excel Unprotector",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        _btnProcess.Enabled = false;
        _log.Clear();

        var successCount = 0;
        var failCount = 0;

        foreach (var path in _fileList.Items.Cast<string>())
        {
            AppendLog($"İşleniyor: {path}");
            var result = ProtectionRemover.RemoveProtections(path);

            if (!result.Success)
            {
                failCount++;
                AppendLog($"   HATA: {result.ErrorMessage}");
                continue;
            }

            successCount++;

            if (!result.AnyProtectionFound)
            {
                AppendLog("   Bu dosyada kaldırılacak bir koruma bulunamadı.");
            }
            else
            {
                if (result.WorkbookProtectionRemoved)
                    AppendLog("   • Çalışma kitabı (yapı) koruması kaldırıldı.");

                foreach (var sheet in result.WorksheetsUnprotected)
                    AppendLog($"   • Sayfa koruması kaldırıldı: {sheet}");
            }

            if (result.VbaProjectPresent)
                AppendLog("   Not: Dosyada VBA projesi var; VBA proje (kod görüntüleme) şifresi bu araçla kaldırılmaz.");

            AppendLog($"   Çıktı: {result.OutputPath}");
            AppendLog("");
        }

        AppendLog(new string('-', 70));
        AppendLog($"Tamamlandı. Başarılı: {successCount}, Hatalı: {failCount}");

        _btnProcess.Enabled = true;

        MessageBox.Show(this,
            $"İşlem tamamlandı.\n\nBaşarılı: {successCount}\nHatalı: {failCount}\n\nÇıktı dosyaları, kaynak dosyalarla aynı klasörde \"_unprotected\" ekiyle oluşturuldu.",
            "Excel Unprotector", MessageBoxButtons.OK,
            failCount == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
    }

    private void AppendLog(string line)
    {
        _log.AppendText(line + Environment.NewLine);
    }

    private void ShowAboutDialog()
    {
        using var about = new Form
        {
            Text = "Excel Unprotector Hakkında",
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = false,
            ClientSize = new Size(360, 190),
            Font = Font
        };

        var title = new Label
        {
            Text = "Excel Unprotector",
            Font = new Font(Font, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(16, 16)
        };

        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var versionLabel = new Label
        {
            Text = $"Sürüm {version?.ToString(3) ?? "1.0.0"}",
            AutoSize = true,
            ForeColor = SystemColors.GrayText,
            Location = new Point(16, 40)
        };

        var coderLabel = new Label
        {
            Text = "Coder: SerdarMSC",
            AutoSize = true,
            Location = new Point(16, 70)
        };

        var githubLink = new LinkLabel
        {
            Text = "https://github.com/SerdarMSC/",
            AutoSize = true,
            Location = new Point(16, 94)
        };
        githubLink.Links.Add(0, githubLink.Text.Length, githubLink.Text);
        githubLink.LinkClicked += (_, _) => OpenUrl(githubLink.Text);

        var emailLink = new LinkLabel
        {
            Text = "serdarmsc@gmail.com",
            AutoSize = true,
            Location = new Point(16, 118)
        };
        emailLink.Links.Add(0, emailLink.Text.Length, emailLink.Text);
        emailLink.LinkClicked += (_, _) => OpenUrl($"mailto:{emailLink.Text}");

        var closeButton = new Button
        {
            Text = "Kapat",
            DialogResult = DialogResult.OK,
            Location = new Point(266, 148),
            Size = new Size(80, 28)
        };

        about.Controls.Add(title);
        about.Controls.Add(versionLabel);
        about.Controls.Add(coderLabel);
        about.Controls.Add(githubLink);
        about.Controls.Add(emailLink);
        about.Controls.Add(closeButton);
        about.AcceptButton = closeButton;

        about.ShowDialog(this);
    }

    private static void OpenUrl(string url)
    {
        try
        {
            Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        }
        catch
        {
            // Kullanicinin varsayilan tarayicisi/posta istemcisi acilamadiysa sessizce yut;
            // baglanti metni zaten ekranda gorunur durumda, elle kopyalanabilir.
        }
    }
}
