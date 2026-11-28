using System.Globalization;
using System.Security;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ReportExport;

public static class ExcelReport
{
    // Записывает строки в Excel по одной, применяет шаблон и добавляет график.
    public static ExportResult Write(string path, IEnumerable<ReportRow> rows, ReportTemplate template,
        CancellationToken cancellation, IProgress<long>? progress)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var book = document.AddWorkbookPart(); book.Workbook = new Workbook();
        var styles = book.AddNewPart<WorkbookStylesPart>(); styles.Stylesheet = MakeStyles(template); styles.Stylesheet.Save();
        var sheet = book.AddNewPart<WorksheetPart>();
        long count = 0; decimal total = 0;
        var chartRows = new List<ReportRow>();
        using (var writer = OpenXmlWriter.Create(sheet))
        {
            writer.WriteStartElement(new Worksheet());
            writer.WriteElement(new SheetViews(new SheetView(new Pane { VerticalSplit = 3, TopLeftCell = "A4", State = PaneStateValues.Frozen, ActivePane = PaneValues.BottomLeft }) { WorkbookViewId = 0 }));
            writer.WriteElement(new Columns(new Column { Min = 1, Max = 1, Width = 48, CustomWidth = true }, new Column { Min = 2, Max = 4, Width = 23, CustomWidth = true }));
            writer.WriteStartElement(new SheetData());
            WriteRow(writer, 1, new[] { TextCell("A1", template.Title, 4) });
            WriteRow(writer, 2, new[] { TextCell("A2", template.Header) });
            WriteRow(writer, 3, new[] { TextCell("A3", "Наименование", 1), TextCell("B3", "Количество", 1), TextCell("C3", "Цена, руб.", 1), TextCell("D3", "Сумма, руб.", 1) });
            foreach (var row in rows)
            {
                cancellation.ThrowIfCancellationRequested(); row.Validate();
                if (count >= 1_048_572) throw new ArgumentException("Для одного листа доступно не более 1048572 строк данных. Разделите CSV на несколько файлов.");
                count++; uint index = (uint)count + 3;
                bool stripe = template.AlternateRows && count % 2 == 0;
                WriteRow(writer, index, new[] { TextCell($"A{index}", row.Name, stripe ? 5u : 0u), NumberCell($"B{index}", row.Quantity, stripe ? 5u : 0u), NumberCell($"C{index}", row.Price, stripe ? 6u : 2u), NumberCell($"D{index}", row.Amount, stripe ? 6u : 2u) });
                total = checked(total + row.Amount);
                if (chartRows.Count < 20) chartRows.Add(row);
                if (count % 1000 == 0) progress?.Report(count);
            }
            uint last = (uint)count + 4;
            WriteRow(writer, last, new[] { TextCell($"A{last}", "Итого", 3), NumberCell($"D{last}", total, 3) });
            writer.WriteEndElement();
            writer.WriteElement(new AutoFilter { Reference = $"A3:D{count + 3}" });
            writer.WriteElement(new MergeCells(new MergeCell { Reference = "A1:D1" }));
            writer.WriteElement(new PageMargins { Left = .3, Right = .3, Top = .7, Bottom = .7, Header = .3, Footer = .3 });
            writer.WriteElement(new PageSetup { Orientation = OrientationValues.Landscape, PaperSize = 9 });
            writer.WriteElement(new HeaderFooter(new OddHeader { Text = "&C" + EscapeHeader(template.Header) }, new OddFooter { Text = "&L" + EscapeHeader(template.Footer) + "&RСтраница &P из &N" }));
            writer.WriteEndElement();
        }
        var sheets = book.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet { Id = book.GetIdOfPart(sheet), SheetId = 1, Name = "Отчёт" });
        if (template.IncludeChart && chartRows.Count > 0) AddChart(book, sheets, chartRows);
        book.Workbook.Save(); progress?.Report(count);
        return new ExportResult(count, total);
    }

    // Создаёт набор шрифтов, заливок и денежных форматов для ячеек.
    private static Stylesheet MakeStyles(ReportTemplate t)
    {
        var normal = new Font(new FontSize { Val = t.FontSize }, new FontName { Val = t.FontFamily });
        var header = new Font(new Bold(), new FontSize { Val = t.FontSize }, new Color { Rgb = "FFFFFFFF" }, new FontName { Val = t.FontFamily });
        var bold = new Font(new Bold(), new FontSize { Val = t.FontSize }, new FontName { Val = t.FontFamily });
        var title = new Font(new Bold(), new FontSize { Val = t.FontSize + 4 }, new FontName { Val = t.FontFamily });
        // Собирает один стиль ячейки из номера шрифта, заливки и формата числа.
        CellFormat Format(uint font, uint fill, uint number) => new CellFormat(new Alignment { Vertical = VerticalAlignmentValues.Center, WrapText = number == 0, ShrinkToFit = number != 0 }) { FontId = font, FillId = fill, BorderId = 0, NumberFormatId = number, FormatId = 0, ApplyFont = true, ApplyFill = true, ApplyNumberFormat = true, ApplyAlignment = true };
        return new Stylesheet(
            new NumberingFormats(new NumberingFormat { NumberFormatId = 164, FormatCode = t.NumberFormat }) { Count = 1 },
            new Fonts(normal, header, bold, title) { Count = 4 },
            new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 }), new Fill(new PatternFill(new ForegroundColor { Rgb = "FF" + t.AccentColor }, new BackgroundColor { Indexed = 64 }) { PatternType = PatternValues.Solid }), new Fill(new PatternFill(new ForegroundColor { Rgb = "FFF1F5F9" }, new BackgroundColor { Indexed = 64 }) { PatternType = PatternValues.Solid })) { Count = 4 },
            new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())) { Count = 1 },
            new CellStyleFormats(new CellFormat { FontId = 0, FillId = 0, BorderId = 0, NumberFormatId = 0 }) { Count = 1 },
            new CellFormats(Format(0, 0, 0), Format(1, 2, 0), Format(0, 0, 164), Format(2, 0, 164), Format(3, 0, 0), Format(0, 3, 0), Format(0, 3, 164)) { Count = 7 },
            new CellStyles(new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }) { Count = 1 });
    }

    // Записывает одну строку и сразу освобождает её из памяти.
    private static void WriteRow(OpenXmlWriter writer, uint index, IEnumerable<Cell> cells)
    {
        var values = cells.ToArray();
        int longest = values.Max(cell => cell.InnerText.Length);
        writer.WriteStartElement(new Row { RowIndex = index, Height = Math.Max(32, ((longest + 27) / 28) * 25), CustomHeight = true });
        foreach (var cell in values) writer.WriteElement(cell);
        writer.WriteEndElement();
    }

    // Создаёт текстовую ячейку; текст не выполняется как формула Excel.
    private static Cell TextCell(string address, string value, uint style = 0) => new Cell(new InlineString(new Text(value) { Space = SpaceProcessingModeValues.Preserve })) { CellReference = address, DataType = CellValues.InlineString, StyleIndex = style };

    // Создаёт числовую ячейку с независимым от языка системы разделителем.
    private static Cell NumberCell(string address, decimal value, uint style) => new Cell(new CellValue(value.ToString(CultureInfo.InvariantCulture))) { CellReference = address, DataType = CellValues.Number, StyleIndex = style };

    // Экранирует управляющий символ в колонтитулах Excel.
    private static string EscapeHeader(string value) => value.Replace("&", "&&");

    // Добавляет на отдельный лист столбчатый график первых двадцати позиций.
    private static void AddChart(WorkbookPart book, Sheets sheets, List<ReportRow> rows)
    {
        var page = book.AddNewPart<WorksheetPart>();
        var drawing = page.AddNewPart<DrawingsPart>();
        var chart = drawing.AddNewPart<ChartPart>();
        // Заменяет специальные символы текста безопасными XML-последовательностями.
        string Escape(string value) => SecurityElement.Escape(value) ?? "";
        var names = new StringBuilder(); var values = new StringBuilder();
        for (int i = 0; i < rows.Count; i++)
        {
            names.Append($"<c:pt idx=\"{i}\"><c:v>{Escape(rows[i].Name)}</c:v></c:pt>");
            values.Append($"<c:pt idx=\"{i}\"><c:v>{rows[i].Amount.ToString(CultureInfo.InvariantCulture)}</c:v></c:pt>");
        }
        string chartXml = $"""
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>
            <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
              <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Сумма, руб.</c:v></c:tx>
                <c:cat><c:strRef><c:f>'Отчёт'!$A$4:$A${rows.Count + 3}</c:f><c:strCache><c:ptCount val="{rows.Count}"/>{names}</c:strCache></c:strRef></c:cat>
                <c:val><c:numRef><c:f>'Отчёт'!$D$4:$D${rows.Count + 3}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{rows.Count}"/>{values}</c:numCache></c:numRef></c:val>
              </c:ser><c:axId val="1"/><c:axId val="2"/>
            </c:barChart>
            <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>
            <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:numFmt formatCode="#,##0.00" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>
          </c:plotArea><c:plotVisOnly val="1"/></c:chart>
        </c:chartSpace>
        """;
        using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(chartXml))) chart.FeedData(stream);
        string drawingXml = $"""
        <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="8572500" cy="4286250"/>
            <xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Суммы первых 20 позиций"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{drawing.GetIdOfPart(chart)}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>
          </xdr:oneCellAnchor>
        </xdr:wsDr>
        """;
        using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(drawingXml))) drawing.FeedData(stream);
        page.Worksheet = new Worksheet(new SheetData(new Row(TextCell("A1", "Суммы первых 20 позиций отчёта")) { RowIndex = 1 }), new Drawing { Id = page.GetIdOfPart(drawing) });
        page.Worksheet.Save();
        sheets.Append(new Sheet { Id = book.GetIdOfPart(page), SheetId = 2, Name = "График" });
    }
}
