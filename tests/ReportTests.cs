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

    // Проверяет построение денежного формата для разных настроек точности.
    [Theory]
    [InlineData(0, "#,##0")]
    [InlineData(2, "#,##0.00")]
    [InlineData(4, "#,##0.0000")]
    public void MoneyFormatMatchesPrecision(int digits, string expected) => Assert.Equal(expected, new ReportTemplate { DecimalPlaces = digits }.NumberFormat);

    // Проверяет отказ от некорректного названия и отрицательных значений.
    [Theory]
    [InlineData("", 1, 1)]
    [InlineData("Товар", -1, 1)]
    [InlineData("Товар", 1, -1)]
    public void InvalidRowsAreRejected(string name, int quantity, int price) => Assert.Throws<ArgumentException>(() => new ReportRow(name, quantity, price).Validate());

    // Проверяет сохранение и загрузку всех пользовательских настроек.
    [Fact]
    public void TemplateRoundTrip()
    {
        var t = new ReportTemplate { Title = "Мой отчёт", Header = "Компания", Footer = "Итоги", FontFamily = "Times New Roman", FontSize = 14, DecimalPlaces = 3, AccentColor = "FF0000", IncludeChart = false, AlternateRows = false };
        var path = Path.Combine(directory, "template.json"); t.Save(path);
        Assert.Equal(t, ReportTemplate.Load(path));
    }

    // Проверяет ограничения цвета, шрифта и точности шаблона.
    [Fact]
    public void InvalidTemplateIsRejected()
    {
        Assert.Throws<ArgumentException>(() => (new ReportTemplate { AccentColor = "oops" }).Validate());
        Assert.Throws<ArgumentException>(() => (new ReportTemplate { FontSize = 100 }).Validate());
        Assert.Throws<ArgumentException>(() => (new ReportTemplate { DecimalPlaces = -1 }).Validate());
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
        SaveSample(path, "sample.xlsx");
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

    // Проверяет отключение графика и чередующейся заливки в шаблоне.
    [Fact]
    public void OptionalFormattingCanBeDisabled()
    {
        var path = Path.Combine(directory, "plain.xlsx");
        ReportExporter.Export(path, rows, new ReportTemplate { IncludeChart = false, AlternateRows = false }, false);
        using var file = SpreadsheetDocument.Open(path, false);
        Assert.Single(file.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>());
        var sheet = file.WorkbookPart.WorksheetParts.Single();
        Assert.Null(sheet.DrawingsPart);
        Assert.Equal(2u, sheet.Worksheet.Descendants<Cell>().Single(c => c.CellReference == "C5").StyleIndex!.Value);
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

    // Проверяет сохранность старого файла и удаление временного файла при отмене.
    [Fact]
    public void CancellationPreservesExistingFile()
    {
        var path = Path.Combine(directory, "existing.xlsx"); File.WriteAllText(path, "старый файл");
        using var cancel = new CancellationTokenSource();
        // Отменяет экспорт после пятидесяти строк тестового источника.
        IEnumerable<ReportRow> Source()
        {
            for (int i = 0; i < 100; i++) { if (i == 50) cancel.Cancel(); yield return new ReportRow("Позиция", 1, 1); }
        }
        Assert.Throws<OperationCanceledException>(() => ReportExporter.Export(path, Source(), new ReportTemplate(), false, cancel.Token));
        Assert.Equal("старый файл", File.ReadAllText(path)); Assert.Single(Directory.GetFiles(directory));
    }

    // Проверяет отклонение пустого отчёта без замены старого файла.
    [Fact]
    public void EmptyReportPreservesExistingFile()
    {
        var path = Path.Combine(directory, "existing.xlsx"); File.WriteAllText(path, "original");
        Assert.Throws<ArgumentException>(() => ReportExporter.Export(path, Array.Empty<ReportRow>(), new ReportTemplate(), false));
        Assert.Equal("original", File.ReadAllText(path)); Assert.Single(Directory.GetFiles(directory));
    }

    // Проверяет CSV с кавычками, кириллицей и разными десятичными разделителями.
    [Fact]
    public void CsvSupportsQuotedNamesAndDecimals()
    {
        var path = Path.Combine(directory, "input.csv");
        File.WriteAllText(path, "Наименование;Количество;Цена\n\"Папка; архив\";2;12,50\nРучка;3;1.25");
        var result = CsvRows.Read(path).ToArray();
        Assert.Equal("Папка; архив", result[0].Name); Assert.Equal(12.50m, result[0].Price); Assert.Equal(1.25m, result[1].Price);
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
        SaveSample(path, "large.pdf");
    }

    // Проверяет пользовательский шаблон PDF и длинные названия на нескольких страницах.
    [Fact]
    public void PdfSupportsLongNamesAndCustomTemplate()
    {
        var path = Path.Combine(directory, "custom.pdf");
        var source = Enumerable.Range(1, 25).Select(i => new ReportRow(new string('Я', 180) + i, 1000000, 1000000000000m));
        var result = ReportExporter.Export(path, source, new ReportTemplate { Title = "Отчёт по продажам", Header = "Компания А&Б", Footer = "Проверка шаблона", FontFamily = "Times New Roman", FontSize = 18, DecimalPlaces = 4, AccentColor = "C2410C" }, true);
        Assert.Equal(25, result.Rows); Assert.Equal(25000000000000000000m, result.Total);
        SaveSample(path, "custom.pdf");
    }

    // Сохраняет пример для визуальной проверки, если задана папка UP0202_SAMPLE_DIR.
    private static void SaveSample(string source, string name)
    {
        var destination = Environment.GetEnvironmentVariable("UP0202_SAMPLE_DIR");
        if (string.IsNullOrEmpty(destination)) return;
        Directory.CreateDirectory(destination); File.Copy(source, Path.Combine(destination, name), true);
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

