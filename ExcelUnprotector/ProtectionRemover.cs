using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelUnprotector;

/// <summary>
/// Excel .xlsx / .xlsm / .xltx / .xltm dosyalari aslinda bir ZIP paketidir (OOXML / Open Packaging
/// Conventions). Calisma kitabi ("Yapiyi Koru") ve sayfa ("Sayfayi Koru") korumalari bu paketin
/// icindeki XML parcalarinda saklanir:
///   - xl/workbook.xml            -> &lt;workbookProtection .../&gt;
///   - xl/worksheets/sheetN.xml   -> &lt;sheetProtection .../&gt;, &lt;protectedRanges&gt;...&lt;/protectedRanges&gt;
///
/// Bu siniftaki mantik, Excel'in kendisini hic acmadan bu etiketleri metin duzeyinde silip
/// paketin geri kalanini (formuller, bicimlendirme, VBA makrolari, ActiveX kontrolleri, resimler,
/// vs.) BAYT BAYT AYNEN korur. Boylece elle yapilan "zip ac -> xml duzelt -> zip'e geri koy"
/// islemi programatik olarak tekrarlanir.
/// </summary>
public static class ProtectionRemover
{
    /// <summary>Bu araci calistirmayi mantikli kilan uzantilar (hepsi OOXML/ZIP tabanli).</summary>
    public static readonly string[] SupportedExtensions = { ".xlsx", ".xlsm", ".xltx", ".xltm" };

    // Ad alani onekli olabilir (orn. <x:workbookProtection ...>), o yuzden (?:\w+:)? ile toleransli.
    private static readonly Regex WorkbookProtectionRegex =
        new(@"<(?:\w+:)?workbookProtection\b[^>]*/>", RegexOptions.Compiled);

    private static readonly Regex SheetProtectionRegex =
        new(@"<(?:\w+:)?sheetProtection\b[^>]*/>", RegexOptions.Compiled);

    // <protectedRanges> bir kapsayici elemandir (icinde protectedRange ogeleri olabilir) ama
    // bazen self-closing da yazilabilir; ikisini de kapsiyoruz.
    private static readonly Regex ProtectedRangesBlockRegex =
        new(@"<(?:\w+:)?protectedRanges\b[^>]*>.*?</(?:\w+:)?protectedRanges>", RegexOptions.Compiled | RegexOptions.Singleline);

    private static readonly Regex ProtectedRangesSelfClosingRegex =
        new(@"<(?:\w+:)?protectedRanges\b[^>]*/>", RegexOptions.Compiled);

    private static readonly Regex ProtectedRangeRegex =
        new(@"<(?:\w+:)?protectedRange\b[^>]*/>", RegexOptions.Compiled);

    public sealed class FileResult
    {
        public string SourcePath { get; set; } = "";
        public string? OutputPath { get; set; }
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public bool WorkbookProtectionRemoved { get; set; }
        public List<string> WorksheetsUnprotected { get; } = new();
        public bool VbaProjectPresent { get; set; }

        public bool AnyProtectionFound => WorkbookProtectionRemoved || WorksheetsUnprotected.Count > 0;
    }

    /// <summary>
    /// Verilen dosyadaki korumalari kaldirir ve sonucu (ayni klasorde, "_unprotected" ekli) yeni
    /// bir dosyaya yazar. Kaynak dosyaya hicbir sekilde dokunulmaz.
    /// </summary>
    /// <param name="sourcePath">Islenecek .xlsx/.xlsm/.xltx/.xltm dosyasi.</param>
    /// <param name="outputPath">
    /// Istege bagli hedef yol. Belirtilmezse kaynak dosyanin yaninda "<ad>_unprotected<uzanti>"
    /// adiyla, zaten varsa (1), (2)... sonekiyle olusturulur.
    /// </param>
    public static FileResult RemoveProtections(string sourcePath, string? outputPath = null)
    {
        var result = new FileResult { SourcePath = sourcePath };

        try
        {
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException("Dosya bulunamadi.", sourcePath);

            var extension = Path.GetExtension(sourcePath);
            if (!SupportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
            {
                throw new NotSupportedException(
                    $"'{extension}' uzantisi desteklenmiyor. Desteklenenler: {string.Join(", ", SupportedExtensions)}. " +
                    "(.xlsb ikili bir formattir, bu araç yalnizca XML tabanli OOXML paketlerini islemektedir.)");
            }

            outputPath ??= BuildDefaultOutputPath(sourcePath);
            EnsureDirectoryExists(outputPath);

            // Once gecici bir dosyaya yaz, basariliysa hedefe tasi. Yari yazilmis/bozuk bir
            // cikti dosyasi birakmamak icin.
            var tempPath = outputPath + ".tmp";

            try
            {
                using (var sourceStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var sourceArchive = new ZipArchive(sourceStream, ZipArchiveMode.Read))
                using (var destStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                using (var destArchive = new ZipArchive(destStream, ZipArchiveMode.Create))
                {
                    foreach (var entry in sourceArchive.Entries)
                    {
                        // Klasor girisleri (bos FullName sonu '/') - nadir ama guvenli tarafta kalalim.
                        if (entry.FullName.EndsWith("/", StringComparison.Ordinal))
                            continue;

                        if (IsVbaProjectBinary(entry.FullName))
                            result.VbaProjectPresent = true;

                        if (IsWorkbookXml(entry.FullName))
                        {
                            CopyTransformedEntry(entry, destArchive, text =>
                            {
                                var updated = WorkbookProtectionRegex.Replace(text, string.Empty);
                                if (updated != text)
                                    result.WorkbookProtectionRemoved = true;
                                return updated;
                            });
                        }
                        else if (IsWorksheetXml(entry.FullName))
                        {
                            CopyTransformedEntry(entry, destArchive, text =>
                            {
                                var original = text;
                                var updated = SheetProtectionRegex.Replace(text, string.Empty);
                                updated = ProtectedRangesBlockRegex.Replace(updated, string.Empty);
                                updated = ProtectedRangesSelfClosingRegex.Replace(updated, string.Empty);
                                updated = ProtectedRangeRegex.Replace(updated, string.Empty);

                                if (updated != original)
                                    result.WorksheetsUnprotected.Add(entry.FullName);

                                return updated;
                            });
                        }
                        else
                        {
                            CopyRawEntry(entry, destArchive);
                        }
                    }
                }

                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                File.Move(tempPath, outputPath);

                result.OutputPath = outputPath;
                result.Success = true;
            }
            catch
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
                throw;
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    private static bool IsWorkbookXml(string fullName) =>
        string.Equals(fullName, "xl/workbook.xml", StringComparison.OrdinalIgnoreCase);

    private static bool IsWorksheetXml(string fullName) =>
        fullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) &&
        fullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);

    private static bool IsVbaProjectBinary(string fullName) =>
        string.Equals(fullName, "xl/vbaProject.bin", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Bir zip girisini metin olarak okuyup donusturur ve hedef arsive yeni bir giris olarak
    /// yazar. UTF-8 BOM varsa/yoksa oldugu gibi korunur; OOXML parcalari genelde BOM'suz UTF-8'dir.
    /// </summary>
    private static void CopyTransformedEntry(ZipArchiveEntry sourceEntry, ZipArchive destArchive, Func<string, string> transform)
    {
        string text;
        bool hasBom;

        using (var entryStream = sourceEntry.Open())
        using (var memory = new MemoryStream())
        {
            entryStream.CopyTo(memory);
            var bytes = memory.ToArray();
            hasBom = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF;

            if (hasBom)
            {
                var withoutBom = new byte[bytes.Length - 3];
                Array.Copy(bytes, 3, withoutBom, 0, withoutBom.Length);
                text = new UTF8Encoding(false).GetString(withoutBom);
            }
            else
            {
                text = new UTF8Encoding(false).GetString(bytes);
            }
        }

        var transformed = transform(text);

        var destEntry = destArchive.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
        destEntry.LastWriteTime = sourceEntry.LastWriteTime;

        using var destEntryStream = destEntry.Open();
        var encoding = new UTF8Encoding(hasBom);
        var outBytes = encoding.GetBytes(transformed);
        destEntryStream.Write(outBytes, 0, outBytes.Length);
    }

    /// <summary>Bir zip girisini hicbir degisiklik yapmadan oldugu gibi hedef arsive kopyalar.</summary>
    private static void CopyRawEntry(ZipArchiveEntry sourceEntry, ZipArchive destArchive)
    {
        var destEntry = destArchive.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
        destEntry.LastWriteTime = sourceEntry.LastWriteTime;

        using var sourceEntryStream = sourceEntry.Open();
        using var destEntryStream = destEntry.Open();
        sourceEntryStream.CopyTo(destEntryStream);
    }

    private static string BuildDefaultOutputPath(string sourcePath)
    {
        var directory = Path.GetDirectoryName(sourcePath) ?? ".";
        var name = Path.GetFileNameWithoutExtension(sourcePath);
        var extension = Path.GetExtension(sourcePath);

        var candidate = Path.Combine(directory, $"{name}_unprotected{extension}");
        var counter = 1;
        while (File.Exists(candidate))
        {
            candidate = Path.Combine(directory, $"{name}_unprotected ({counter}){extension}");
            counter++;
        }

        return candidate;
    }

    private static void EnsureDirectoryExists(string filePath)
    {
        var directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            Directory.CreateDirectory(directory);
    }
}
