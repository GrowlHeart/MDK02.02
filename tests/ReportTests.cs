using System.Collections;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace ReportExport.Tests;

public sealed class ReportTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "UP0202-tests-" + Guid.NewGuid().ToString("N"));
    private readonly ReportRow[] rows = { new("Бумага А4", 5, 350m), new("Ручка", 12, 45m), new("Папка", 3, 120m) };

    // Создаёт отдельную временную папку для каждого теста.
    public ReportTests() => Directory.CreateDirectory(directory);

    // Удаляет файлы, созданные текущим тестом.
    public void Dispose() => Directory.Delete(directory, true);

    // Проверяет расчёт суммы без ошибок дробных чисел.
    [Fact]
    public void RowAmountUsesDecimal() => Assert.Equal(.3m, new ReportRow("Товар", 3, .1m).Amount);

    // Проверяет, что шаблон задаёт нужное число знаков после запятой.
    [Fact]
    public void TemplateNumberFormatMatchesPrecision()
    {
        var template = new ReportTemplate { DecimalPlaces = 3 };
        template.Validate();
        Assert.Equal("#,##0.000", template.NumberFormat);
    }

    // Проверяет заголовки, стили, колонтитулы и итог в созданном Excel.
    [Fact]
    public void ExcelFormattingMatchesTemplate()
    {
        var path = Path.Combine(directory, "report.xlsx");
        var template = new ReportTemplate { Title = "Продажи", Header = "А&Б", Footer = "Подпись", AccentColor = "FF0000", FontFamily = "Times New Roman", DecimalPlaces = 3 };
        var result = ReportExporter.Export(path, rows, template, false);
        Assert.Equal(3, result.Rows); Assert.Equal(2650m, result.Total);
        using var file = SpreadsheetDocument.Open(path, false);
        var book = file.WorkbookPart!;
        var sheet = (WorksheetPart)book.GetPartById(book.Workbook.Sheets!.Elements<Sheet>().First().Id!);
        var cells = sheet.Worksheet.Descendants<Cell>().ToDictionary(x => x.CellReference!.Value!);
        Assert.Equal("Продажи", cells["A1"].InnerText);
        Assert.Equal("Наименование", cells["A3"].InnerText);
        Assert.Equal("Количество", cells["B3"].InnerText);
        Assert.Equal("Цена, руб.", cells["C3"].InnerText);
        Assert.Equal("Сумма, руб.", cells["D3"].InnerText);
        Assert.Equal("2650", cells["D7"].InnerText);
        Assert.Equal(CellValues.Number, cells["D4"].DataType!.Value);
        Assert.Equal(1u, cells["A3"].StyleIndex!.Value);
        Assert.Equal(6u, cells["C5"].StyleIndex!.Value);
        var styles = book.WorkbookStylesPart!.Stylesheet;
        Assert.Equal("#,##0.000", styles.NumberingFormats!.Elements<NumberingFormat>().Single().FormatCode!.Value);
        Assert.Equal("Times New Roman", styles.Fonts!.Elements<Font>().First().FontName!.Val!.Value);
        Assert.NotNull(styles.Fonts.Elements<Font>().ElementAt(1).Bold);
        Assert.Equal("FFFF0000", styles.Fills!.Elements<Fill>().ElementAt(2).PatternFill!.ForegroundColor!.Rgb!.Value);
        Assert.Equal("&CА&&Б", sheet.Worksheet.GetFirstChild<HeaderFooter>()!.OddHeader!.Text);
        Assert.Contains("&P", sheet.Worksheet.GetFirstChild<HeaderFooter>()!.OddFooter!.Text);
        var errors = new OpenXmlValidator().Validate(file).Select(e => e.Description + " Path: " + e.Path?.XPath).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
    }

    // Проверяет связи графика с данными и ограничение его двадцатью позициями.
    [Fact]
    public void ChartReferencesFirstTwentyRows()
    {
        var path = Path.Combine(directory, "chart.xlsx");
        ReportExporter.Export(path, Enumerable.Range(1, 30).Select(i => new ReportRow("Товар & <" + i, 1, i)), new ReportTemplate(), false);
        using var file = SpreadsheetDocument.Open(path, false);
        Assert.Equal(2, file.WorkbookPart!.Workbook.Sheets!.Count());
        var charts = file.WorkbookPart.WorksheetParts.SelectMany(p => p.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>()).ToArray();
        Assert.Single(charts);
        using var stream = charts[0].GetStream(); var xml = XDocument.Load(stream);
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        Assert.Contains("'Отчёт'!$A$4:$A$23", xml.Descendants(c + "f").Select(x => x.Value));
        Assert.Contains("'Отчёт'!$D$4:$D$23", xml.Descendants(c + "f").Select(x => x.Value));
        Assert.Equal(20, xml.Descendants(c + "strCache").Descendants(c + "pt").Count());
        Assert.Contains("Товар & <1", xml.Descendants(c + "v").Select(x => x.Value));
        var errors = new OpenXmlValidator().Validate(file).Select(e => e.Description + " Path: " + e.Path?.XPath).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
    }

    // Проверяет однократное чтение большого источника без повторного обхода.
    [Fact]
    public void LargeExcelReadsSourceOnce()
    {
        var source = new SinglePassRows(100000);
        var path = Path.Combine(directory, "large.xlsx");
        var result = ReportExporter.Export(path, source, new ReportTemplate(), false);
        Assert.Equal(100000, result.Rows); Assert.Equal(125000m, result.Total);
        Assert.Equal(100000, source.ReadCount);
        using var file = SpreadsheetDocument.Open(path, false);
        var sheet = (WorksheetPart)file.WorkbookPart!.GetPartById(file.WorkbookPart.Workbook.Sheets!.Elements<Sheet>().First().Id!);
        int count = 0;
        using var reader = OpenXmlReader.Create(sheet);
        while (reader.Read()) if (reader.IsStartElement && reader.ElementType == typeof(Row)) count++;
        Assert.Equal(100004, count);
    }

    // Проверяет создание многостраничного PDF из однократного источника.
    [Fact]
    public void PdfAcceptsStreamingRows()
    {
        var path = Path.Combine(directory, "report.pdf");
        var result = ReportExporter.Export(path, new SinglePassRows(1000), new ReportTemplate(), true);
        Assert.Equal(1000, result.Rows); Assert.Equal(1250m, result.Total);
        using var stream = File.OpenRead(path); var bytes = new byte[5]; stream.ReadExactly(bytes);
        Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(bytes));
        Assert.True(new FileInfo(path).Length > 10000);
    }

    private sealed class SinglePassRows : IEnumerable<ReportRow>
    {
        private readonly int count;
        private bool started;
        public int ReadCount { get; private set; }

        // Запоминает нужное количество строк тестового источника.
        public SinglePassRows(int count) => this.count = count;

        // Отдаёт строки по одной и запрещает повторное чтение источника.
        public IEnumerator<ReportRow> GetEnumerator()
        {
            if (started) throw new InvalidOperationException("Источник уже прочитан"); started = true;
            for (int i = 0; i < count; i++) { ReadCount++; yield return new ReportRow("Позиция " + i, 1, 1.25m); }
        }

        // Поддерживает стандартный интерфейс перечисления строк.
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}

