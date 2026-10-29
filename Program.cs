using System.Data;
using ClosedXML.Excel;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace ReportExport;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        QuestPDF.Settings.License = LicenseType.Community;
        ApplicationConfiguration.Initialize();
        Application.Run(new ReportForm());
    }
}

public sealed class ReportForm : Form
{
    private readonly DataTable data = new();
    private readonly DataGridView grid = new();
    private readonly Label total = new();

    public ReportForm()
    {
        Text = "Экспорт отчётов";
        Width = 900; Height = 550; MinimumSize = new System.Drawing.Size(650, 400);
        StartPosition = FormStartPosition.CenterScreen;
        data.Columns.Add("Наименование", typeof(string));
        data.Columns.Add("Количество", typeof(int));
        data.Columns.Add("Цена", typeof(decimal));
        data.Rows.Add("Бумага А4", 5, 350m);
        data.Rows.Add("Ручка", 12, 45m);
        data.Rows.Add("Папка", 3, 120m);

        var toolbar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 48, Padding = new Padding(8), AutoSize = false };
        toolbar.Controls.Add(MakeButton("Добавить строку", (_, _) => data.Rows.Add("Новая позиция", 1, 0m)));
        toolbar.Controls.Add(MakeButton("Удалить строку", (_, _) => { if (grid.CurrentRow?.DataBoundItem is DataRowView row) row.Delete(); }));
        toolbar.Controls.Add(MakeButton("Экспорт в Excel", (_, _) => Export(false)));
        toolbar.Controls.Add(MakeButton("Экспорт в PDF", (_, _) => Export(true)));
        grid.Dock = DockStyle.Fill;
        grid.DataSource = data;
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        grid.AllowUserToAddRows = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.DataError += (_, e) => { e.ThrowException = false; MessageBox.Show("Введите корректное значение: количество — целое число, цена — число.", "Ошибка ввода"); };
        total.Dock = DockStyle.Bottom; total.Height = 35; total.Padding = new Padding(10, 8, 0, 0);
        data.RowChanged += (_, _) => UpdateTotal();
        data.RowDeleted += (_, _) => UpdateTotal();
        Controls.Add(grid); Controls.Add(total); Controls.Add(toolbar);
        UpdateTotal();
    }

    private static Button MakeButton(string text, EventHandler action)
    {
        var button = new Button { Text = text, AutoSize = true, Height = 30 };
        button.Click += action;
        return button;
    }

    private void UpdateTotal()
    {
        decimal sum = 0;
        foreach (DataRow row in data.Rows)
            if (row.RowState != DataRowState.Deleted && row[1] != DBNull.Value && row[2] != DBNull.Value)
                sum += (int)row[1] * (decimal)row[2];
        total.Text = $"Итого: {sum:N2} руб.";
    }

    private void Export(bool pdf)
    {
        if (!grid.EndEdit()) return;
        BindingContext?[data]?.EndCurrentEdit();
        var items = new List<(string Name, int Quantity, decimal Price)>();
        foreach (DataRow row in data.Rows)
        {
            if (row.RowState == DataRowState.Deleted) continue;
            if (row.IsNull(0) || string.IsNullOrWhiteSpace(row[0].ToString()) || row.IsNull(1) || row.IsNull(2) || (int)row[1] < 0 || (decimal)row[2] < 0)
            {
                MessageBox.Show("Заполните название, укажите неотрицательные количество и цену.", "Проверка данных"); return;
            }
            items.Add((row[0].ToString()!, (int)row[1], (decimal)row[2]));
        }
        using var dialog = new SaveFileDialog { Filter = pdf ? "PDF-файл (*.pdf)|*.pdf" : "Excel-файл (*.xlsx)|*.xlsx", FileName = pdf ? "Отчёт.pdf" : "Отчёт.xlsx" };
        if (dialog.ShowDialog() != DialogResult.OK) return;
        try
        {
            if (pdf) SavePdf(dialog.FileName, items); else SaveExcel(dialog.FileName, items);
            MessageBox.Show("Отчёт сохранён.", "Экспорт");
        }
        catch (Exception ex) { MessageBox.Show($"Не удалось сохранить отчёт: {ex.Message}", "Ошибка"); }
    }

    private static void SaveExcel(string path, List<(string Name, int Quantity, decimal Price)> items)
    {
        using var book = new XLWorkbook();
        var sheet = book.Worksheets.Add("Отчёт");
        string[] headers = ["Наименование", "Количество", "Цена, руб.", "Сумма, руб."];
        for (int i = 0; i < headers.Length; i++) sheet.Cell(1, i + 1).Value = headers[i];
        sheet.Range(1, 1, 1, 4).Style.Font.Bold = true;
        for (int i = 0; i < items.Count; i++)
        {
            var item = items[i]; int row = i + 2;
            sheet.Cell(row, 1).Value = item.Name; sheet.Cell(row, 2).Value = item.Quantity;
            sheet.Cell(row, 3).Value = item.Price; sheet.Cell(row, 4).Value = item.Quantity * item.Price;
        }
        int end = items.Count + 2;
        sheet.Cell(end, 1).Value = "Итого"; sheet.Cell(end, 4).Value = items.Sum(x => x.Quantity * x.Price);
        sheet.Range(end, 1, end, 4).Style.Font.Bold = true;
        sheet.Columns(3, 4).Style.NumberFormat.Format = "# ##0.00";
        sheet.Columns().AdjustToContents();
        book.SaveAs(path);
    }

    private static void SavePdf(string path, List<(string Name, int Quantity, decimal Price)> items)
    {
        Document.Create(document => document.Page(page =>
        {
            page.Margin(35); page.DefaultTextStyle(x => x.FontSize(10));
            page.Header().Text("Отчёт по позициям").FontSize(18).Bold();
            page.Content().PaddingTop(15).Table(table =>
            {
                table.ColumnsDefinition(columns => { columns.RelativeColumn(4); columns.RelativeColumn(2); columns.RelativeColumn(2); columns.RelativeColumn(2); });
                table.Header(header => { foreach (var title in new[] { "Наименование", "Количество", "Цена, руб.", "Сумма, руб." }) header.Cell().Background(Colors.Grey.Lighten2).Padding(5).Text(title).Bold(); });
                foreach (var item in items)
                {
                    table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Padding(5).Text(item.Name);
                    table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Padding(5).Text(item.Quantity.ToString());
                    table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Padding(5).Text(item.Price.ToString("N2"));
                    table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Padding(5).Text((item.Quantity * item.Price).ToString("N2"));
                }
                table.Cell().ColumnSpan(3).Padding(5).Text("Итого").Bold();
                table.Cell().Padding(5).Text(items.Sum(x => x.Quantity * x.Price).ToString("N2")).Bold();
            });
            page.Footer().AlignCenter().Text(text => { text.Span("Страница "); text.CurrentPageNumber(); });
        })).GeneratePdf(path);
    }
}
