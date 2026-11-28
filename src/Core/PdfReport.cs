using System.Globalization;
using SkiaSharp;

namespace ReportExport;

public static class PdfReport
{
    // Рисует PDF постранично, не собирая все строки отчёта в список.
    public static ExportResult Write(string path, IEnumerable<ReportRow> rows, ReportTemplate template,
        CancellationToken cancellation, IProgress<long>? progress)
    {
        using var stream = File.Create(path);
        using var document = SKDocument.CreatePdf(stream);
        using var typeface = SKTypeface.FromFamilyName(template.FontFamily);
        using var paint = new SKPaint { Typeface = typeface, TextSize = template.FontSize, IsAntialias = true };
        long count = 0; decimal total = 0; int pageNumber = 0;
        SKCanvas? canvas = null; float y = 0;

        // Завершает текущую страницу и начинает новую с повторным заголовком.
        void NewPage()
        {
            if (canvas is not null) document.EndPage();
            canvas = document.BeginPage(842, 595); pageNumber++;
            paint.Color = SKColors.Black; paint.TextSize = template.FontSize + 5;
            DrawText(canvas, template.Title, 30, 32, 780, paint);
            paint.TextSize = template.FontSize;
            DrawText(canvas, template.Header, 30, 55, 780, paint);
            DrawText(canvas, template.Footer, 30, 575, 650, paint);
            canvas.DrawText($"Страница {pageNumber}", 710, 575, paint);
            paint.Color = SKColor.Parse("#" + template.AccentColor);
            canvas.DrawRect(30, 72, 782, 28, paint);
            paint.Color = SKColors.White;
            canvas.DrawText("Наименование", 35, 91, paint);
            canvas.DrawText("Количество", 445, 91, paint);
            canvas.DrawText("Цена, руб.", 565, 91, paint);
            canvas.DrawText("Сумма, руб.", 680, 91, paint);
            y = 100;
        }

        foreach (var row in rows)
        {
            cancellation.ThrowIfCancellationRequested(); row.Validate();
            var lines = WrapText(row.Name.Replace('\n', ' ').Replace('\r', ' '), 395, paint);
            float rowHeight = Math.Max(32, lines.Count * (template.FontSize + 3) + 10);
            if (canvas is null || y + rowHeight > 540) NewPage();
            count++; total = checked(total + row.Amount);
            paint.Color = template.AlternateRows && count % 2 == 0 ? SKColor.Parse("#F1F5F9") : SKColors.White;
            canvas!.DrawRect(30, y, 782, rowHeight, paint);
            paint.Color = SKColors.Black;
            for (int i = 0; i < lines.Count; i++) canvas.DrawText(lines[i], 35, y + template.FontSize + 5 + i * (template.FontSize + 3), paint);
            canvas.DrawText(row.Quantity.ToString(CultureInfo.InvariantCulture), 445, y + 21, paint);
            DrawText(canvas, row.Price.ToString("N" + template.DecimalPlaces, CultureInfo.GetCultureInfo("ru-RU")), 565, y + 21, 105, paint);
            DrawText(canvas, row.Amount.ToString("N" + template.DecimalPlaces, CultureInfo.GetCultureInfo("ru-RU")), 680, y + 21, 126, paint);
            y += rowHeight;
            if (count % 1000 == 0) progress?.Report(count);
        }
        if (canvas is null || y + 32 > 540) NewPage();
        paint.Color = SKColors.Black;
        DrawText(canvas!, $"Итого: {total.ToString("N" + template.DecimalPlaces, CultureInfo.GetCultureInfo("ru-RU"))} руб.", 35, y + 23, 770, paint);
        document.EndPage(); document.Close(); progress?.Report(count);
        return new ExportResult(count, total);
    }

    // Уменьшает шрифт длинного значения, сохраняя все его символы.
    private static void DrawText(SKCanvas canvas, string text, float x, float y, float width, SKPaint paint)
    {
        float originalSize = paint.TextSize;
        float measured = paint.MeasureText(text);
        if (measured > width) paint.TextSize *= width / measured;
        canvas.DrawText(text, x, y, paint);
        paint.TextSize = originalSize;
    }

    // Разбивает длинное название на строки по ширине столбца PDF.
    private static List<string> WrapText(string value, float width, SKPaint paint)
    {
        var lines = new List<string>(); string current = "";
        foreach (char letter in value)
        {
            if (current.Length > 0 && paint.MeasureText(current + letter) > width) { lines.Add(current); current = ""; }
            current += letter;
        }
        if (current.Length > 0) lines.Add(current);
        return lines;
    }
}
