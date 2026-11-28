using System.Globalization;
using System.Text.Json;
using System.Xml;
using Microsoft.VisualBasic.FileIO;

namespace ReportExport;

public sealed record ReportRow(string Name, int Quantity, decimal Price)
{
    public decimal Amount => checked(Quantity * Price);
    // Проверяет название, количество и цену одной строки.
    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(Name) || Name.Length > 200)
            throw new ArgumentException("Название должно содержать от 1 до 200 символов.");
        XmlConvert.VerifyXmlChars(Name);
        if (Quantity < 0 || Quantity > 1_000_000 || Price < 0 || Price > 1_000_000_000_000m)
            throw new ArgumentException("Количество: 0–1000000; цена: 0–1000000000000.");
    }
}

public sealed record ReportTemplate
{
    public string Title { get; init; } = "Отчёт по позициям";
    public string Header { get; init; } = "UP02.02";
    public string Footer { get; init; } = "Отчёт сформирован в UP02.02";
    public string FontFamily { get; init; } = "Arial";
    public int FontSize { get; init; } = 11;
    public string AccentColor { get; init; } = "2563EB";
    public int DecimalPlaces { get; init; } = 2;
    public bool AlternateRows { get; init; } = true;
    public bool IncludeChart { get; init; } = true;
    public string NumberFormat => "#,##0" + (DecimalPlaces == 0 ? "" : "." + new string('0', DecimalPlaces));

    // Проверяет допустимые значения настроек шаблона.
    public void Validate()
    {
        foreach (var value in new[] { Title, Header, Footer })
        {
            if (value is null || value.Length > 200) throw new ArgumentException("Текст шаблона: не более 200 символов.");
            XmlConvert.VerifyXmlChars(value);
        }
        if (string.IsNullOrWhiteSpace(Title)) throw new ArgumentException("Укажите название отчёта.");
        if (FontFamily is not ("Arial" or "Segoe UI" or "Times New Roman")) throw new ArgumentException("Выберите шрифт из списка.");
        if (FontSize is < 9 or > 18 || DecimalPlaces is < 0 or > 4) throw new ArgumentException("Некорректный размер шрифта или точность.");
        if (AccentColor is null || AccentColor.Length != 6 || !AccentColor.All(Uri.IsHexDigit)) throw new ArgumentException("Цвет должен содержать 6 шестнадцатеричных символов.");
    }

    // Сохраняет настройки оформления в JSON.
    public void Save(string path)
    {
        Validate();
        File.WriteAllText(path, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
    }
    // Загружает JSON и проверяет настройки перед использованием.
    public static ReportTemplate Load(string path)
    {
        if (new FileInfo(path).Length > 64 * 1024) throw new ArgumentException("Слишком большой файл шаблона.");
        var result = JsonSerializer.Deserialize<ReportTemplate>(File.ReadAllText(path)) ?? throw new ArgumentException("Пустой шаблон.");
        result.Validate(); return result;
    }
}

public static class CsvRows
{
    // Читает CSV по одной строке, не загружая файл целиком в память.
    public static IEnumerable<ReportRow> Read(string path)
    {
        using var parser = new TextFieldParser(path, System.Text.Encoding.UTF8, true);
        parser.SetDelimiters(";"); parser.HasFieldsEnclosedInQuotes = true;
        var header = parser.ReadFields();
        if (header is null || !header.SequenceEqual(new[] { "Наименование", "Количество", "Цена" }))
            throw new ArgumentException("CSV UTF-8: первая строка должна быть Наименование;Количество;Цена");
        while (!parser.EndOfData)
        {
            var fields = parser.ReadFields();
            if (fields is null) continue;
            if (fields.Length != 3 || !int.TryParse(fields[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var count)
                || !decimal.TryParse(fields[2].Replace(',', '.'), NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var price))
                throw new ArgumentException($"Некорректные данные CSV около строки {parser.LineNumber}.");
            var row = new ReportRow(fields[0], count, price); row.Validate(); yield return row;
        }
    }
}

public sealed record ExportResult(long Rows, decimal Total);

public static class ReportExporter
{
    // Создаёт временный отчёт и заменяет итоговый файл только после успешного экспорта.
    public static ExportResult Export(string path, IEnumerable<ReportRow> rows, ReportTemplate template,
        bool pdf, CancellationToken cancellation = default, IProgress<long>? progress = null)
    {
        template.Validate();
        var fullPath = Path.GetFullPath(path);
        var temporary = Path.Combine(Path.GetDirectoryName(fullPath)!, "." + Guid.NewGuid().ToString("N") + ".tmp");
        try
        {
            cancellation.ThrowIfCancellationRequested();
            var result = pdf ? PdfReport.Write(temporary, rows, template, cancellation, progress)
                : ExcelReport.Write(temporary, rows, template, cancellation, progress);
            cancellation.ThrowIfCancellationRequested();
            if (result.Rows == 0) throw new ArgumentException("В отчёте нет данных.");
            File.Move(temporary, fullPath, true);
            return result;
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
