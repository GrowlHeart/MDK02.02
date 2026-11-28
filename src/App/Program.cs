using System.Data;
using System.Globalization;

namespace ReportExport;

internal static class Program
{
    // Открывает главное окно приложения.
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new ReportForm());
    }
}

public sealed class ReportForm : Form
{
    private readonly DataTable data = new();
    private readonly DataGridView grid = new();
    private readonly Label status = new() { AutoSize = true, Text = "Источник: таблица" };
    private readonly FlowLayoutPanel toolbar = new() { Dock = DockStyle.Fill, AutoSize = true };
    private readonly FlowLayoutPanel settings = new() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };
    private readonly TextBox title = new() { Width = 280, Text = "Отчёт по позициям", MaxLength = 200 };
    private readonly TextBox header = new() { Width = 280, Text = "UP02.02", MaxLength = 200 };
    private readonly TextBox footer = new() { Width = 280, Text = "Отчёт сформирован в UP02.02", MaxLength = 200 };
    private readonly TextBox color = new() { Width = 280, Text = "2563EB", MaxLength = 6 };
    private readonly ComboBox font = new() { Width = 280, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly NumericUpDown size = new() { Minimum = 9, Maximum = 18, Value = 11, Width = 100 };
    private readonly NumericUpDown decimals = new() { Minimum = 0, Maximum = 4, Value = 2, Width = 100 };
    private readonly CheckBox stripes = new() { Text = "Чередовать фон строк", Checked = true, AutoSize = true };
    private readonly CheckBox chart = new() { Text = "График Excel: первые 20 позиций", Checked = true, AutoSize = true };
    private readonly Button cancel = new() { Text = "Отменить экспорт", AutoSize = true, Enabled = false };
    private readonly ProgressBar progressBar = new() { Width = 140, Style = ProgressBarStyle.Marquee, Visible = false };
    private string? csvPath;
    private CancellationTokenSource? cancellation;

    // Создаёт таблицу, кнопки экспорта и поля настройки шаблона.
    public ReportForm()
    {
        Text = "UP02.02 — Экспорт отчётов"; Width = 1150; Height = 780; MinimumSize = new Size(950, 650);
        StartPosition = FormStartPosition.CenterScreen;
        if (Environment.ProcessPath is { } path) Icon = Icon.ExtractAssociatedIcon(path);
        data.Columns.Add("Наименование", typeof(string)); data.Columns.Add("Количество", typeof(int)); data.Columns.Add("Цена", typeof(decimal));
        data.Rows.Add("Бумага А4", 5, 350m); data.Rows.Add("Ручка", 12, 45m); data.Rows.Add("Папка", 3, 120m);
        grid.DataSource = data; grid.Dock = DockStyle.Fill; grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false;
        grid.DataError += (_, e) => { e.ThrowException = false; MessageBox.Show("Количество должно быть целым числом, цена — числом.", "Ошибка ввода"); };
        toolbar.Controls.Add(Button("Добавить", (_, _) => data.Rows.Add("Новая позиция", 1, 0m)));
        toolbar.Controls.Add(Button("Удалить", (_, _) => { if (grid.CurrentRow?.DataBoundItem is DataRowView row) row.Delete(); }));
        toolbar.Controls.Add(Button("Открыть CSV", (_, _) => OpenCsv()));
        toolbar.Controls.Add(Button("Пример 100000 строк", async (_, _) => await CreateDemo()));
        toolbar.Controls.Add(Button("К таблице", (_, _) => { csvPath = null; grid.DataSource = data; grid.ReadOnly = false; toolbar.Controls[0].Enabled = true; toolbar.Controls[1].Enabled = true; status.Text = "Источник: таблица"; }));
        toolbar.Controls.Add(Button("Excel", async (_, _) => await Export(false)));
        toolbar.Controls.Add(Button("PDF", async (_, _) => await Export(true)));
        font.Items.AddRange(new object[] { "Arial", "Segoe UI", "Times New Roman" }); font.SelectedIndex = 0;
        Field("Название отчёта", title); Field("Верхний колонтитул", header); Field("Нижний колонтитул", footer);
        Field("Шрифт", font); Field("Размер шрифта", size); Field("Знаков после запятой", decimals); Field("Цвет заголовка (HEX)", color);
        settings.Controls.Add(stripes); settings.Controls.Add(chart);
        settings.Controls.Add(Button("Сохранить шаблон", (_, _) => SaveTemplate()));
        settings.Controls.Add(Button("Загрузить шаблон", (_, _) => LoadTemplate()));
        settings.Controls.Add(new Label { Text = "CSV: UTF-8, разделитель ;\nНаименование;Количество;Цена\nВ таблице показаны первые 100 строк CSV.\nЭкспорт обрабатывает весь файл.", AutoSize = true, MaximumSize = new Size(280, 0), Margin = new Padding(3, 15, 3, 3) });
        var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 310));
        body.Controls.Add(grid, 0, 0); body.Controls.Add(settings, 1, 0);
        var bottom = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        bottom.Controls.Add(cancel); bottom.Controls.Add(progressBar); bottom.Controls.Add(status);
        cancel.Click += (_, _) => cancellation?.Cancel();
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, Padding = new Padding(10) };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(toolbar, 0, 0); layout.Controls.Add(body, 0, 1); layout.Controls.Add(bottom, 0, 2); Controls.Add(layout);
        FormClosing += (_, e) => { if (cancellation is not null) { cancellation.Cancel(); e.Cancel = true; } };
        data.RowChanged += (_, _) => UpdateTotal(); data.RowDeleted += (_, _) => UpdateTotal(); UpdateTotal();
    }

    // Создаёт кнопку и связывает её с нужным действием.
    private static Button Button(string text, EventHandler action)
    {
        var button = new Button { Text = text, AutoSize = true, Height = 32 };
        button.Click += action; return button;
    }

    // Добавляет подпись и поле в панель шаблона.
    private void Field(string label, Control control)
    {
        settings.Controls.Add(new Label { Text = label, AutoSize = true, Margin = new Padding(3, 9, 3, 2) }); settings.Controls.Add(control);
    }

    // Собирает настройки шаблона из полей окна.
    private ReportTemplate Template()
    {
        var value = new ReportTemplate { Title = title.Text, Header = header.Text, Footer = footer.Text, FontFamily = font.Text, FontSize = (int)size.Value, DecimalPlaces = (int)decimals.Value, AccentColor = color.Text, AlternateRows = stripes.Checked, IncludeChart = chart.Checked };
        value.Validate(); return value;
    }

    // Пересчитывает общий итог после изменения данных в таблице.
    private void UpdateTotal()
    {
        if (csvPath is not null || cancellation is not null) return;
        try
        {
            decimal total = 0;
            foreach (DataRow row in data.Rows)
                if (row.RowState != DataRowState.Deleted && !row.IsNull(1) && !row.IsNull(2)) total = checked(total + (int)row[1] * (decimal)row[2]);
            status.Text = $"Источник: таблица. Итого: {total:N2} руб.";
        }
        catch (OverflowException) { status.Text = "Слишком большие значения в таблице"; }
    }

    // Сохраняет пользовательский шаблон в JSON-файл.
    private void SaveTemplate()
    {
        using var dialog = new SaveFileDialog { Filter = "Шаблон (*.json)|*.json", FileName = "Шаблон.json" };
        if (dialog.ShowDialog() != DialogResult.OK) return;
        try { Template().Save(dialog.FileName); } catch (Exception e) { MessageBox.Show(e.Message, "Ошибка шаблона"); }
    }

    // Читает шаблон из JSON и заполняет поля окна.
    private void LoadTemplate()
    {
        using var dialog = new OpenFileDialog { Filter = "Шаблон (*.json)|*.json" };
        if (dialog.ShowDialog() != DialogResult.OK) return;
        try
        {
            var t = ReportTemplate.Load(dialog.FileName);
            title.Text = t.Title; header.Text = t.Header; footer.Text = t.Footer; font.SelectedItem = t.FontFamily;
            size.Value = t.FontSize; decimals.Value = t.DecimalPlaces; color.Text = t.AccentColor; stripes.Checked = t.AlternateRows; chart.Checked = t.IncludeChart;
        }
        catch (Exception e) { MessageBox.Show(e.Message, "Ошибка шаблона"); }
    }

    // Открывает большой CSV и показывает только первые сто строк для просмотра.
    private void OpenCsv()
    {
        using var dialog = new OpenFileDialog { Filter = "Таблица CSV (*.csv)|*.csv" };
        if (dialog.ShowDialog() != DialogResult.OK) return;
        try { Preview(dialog.FileName); } catch (Exception e) { MessageBox.Show(e.Message, "Ошибка CSV"); }
    }

    // Загружает небольшой предварительный просмотр, оставляя весь CSV на диске.
    private void Preview(string path)
    {
        var preview = data.Clone();
        foreach (var row in CsvRows.Read(path).Take(100)) preview.Rows.Add(row.Name, row.Quantity, row.Price);
        if (preview.Rows.Count == 0) throw new ArgumentException("CSV не содержит данных.");
        csvPath = path; grid.DataSource = preview; grid.ReadOnly = true;
        toolbar.Controls[0].Enabled = false; toolbar.Controls[1].Enabled = false;
        status.Text = $"CSV: {Path.GetFileName(path)} — просмотр {preview.Rows.Count} строк";
    }

    // Создаёт большой пример CSV последовательной записью строк на диск.
    private async Task CreateDemo()
    {
        using var dialog = new SaveFileDialog { Filter = "CSV (*.csv)|*.csv", FileName = "Пример100000.csv" };
        if (dialog.ShowDialog() != DialogResult.OK) return;
        if (csvPath is not null && Path.GetFullPath(csvPath).Equals(Path.GetFullPath(dialog.FileName), StringComparison.OrdinalIgnoreCase)) { MessageBox.Show("Выберите другой путь, чтобы сохранить исходный CSV."); return; }
        SetBusy(true);
        string temporary = dialog.FileName + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            var token = cancellation!.Token;
            await Task.Run(() =>
            {
                using (var writer = new StreamWriter(temporary, false, System.Text.Encoding.UTF8))
                {
                    writer.WriteLine("Наименование;Количество;Цена");
                    for (int i = 1; i <= 100000; i++) { token.ThrowIfCancellationRequested(); writer.WriteLine($"Позиция {i};{i % 10 + 1};{(i % 1000 + .25m).ToString(CultureInfo.InvariantCulture)}"); }
                }
                token.ThrowIfCancellationRequested(); File.Move(temporary, dialog.FileName, true);
            });
            Preview(dialog.FileName);
        }
        catch (OperationCanceledException) { status.Text = "Создание примера отменено"; }
        catch (Exception e) { MessageBox.Show(e.Message, "Ошибка"); }
        finally { if (File.Exists(temporary)) File.Delete(temporary); SetBusy(false); }
    }

    // Запускает экспорт в фоне, показывает число обработанных строк и поддерживает отмену.
    private async Task Export(bool pdf)
    {
        if (!grid.EndEdit()) return;
        BindingContext?[data]?.EndCurrentEdit();
        try
        {
            var template = Template();
            IEnumerable<ReportRow> rows;
            if (csvPath is not null) rows = CsvRows.Read(csvPath);
            else
            {
                var snapshot = new List<ReportRow>();
                foreach (DataRow row in data.Rows)
                {
                    if (row.RowState == DataRowState.Deleted) continue;
                    if (row.IsNull(0) || row.IsNull(1) || row.IsNull(2)) throw new ArgumentException("Заполните все ячейки таблицы.");
                    var value = new ReportRow((string)row[0], (int)row[1], (decimal)row[2]); value.Validate(); snapshot.Add(value);
                }
                rows = snapshot;
            }
            using var dialog = new SaveFileDialog { Filter = pdf ? "PDF (*.pdf)|*.pdf" : "Excel (*.xlsx)|*.xlsx", FileName = pdf ? "Отчёт.pdf" : "Отчёт.xlsx" };
            if (dialog.ShowDialog() != DialogResult.OK) return;
            if (csvPath is not null && Path.GetFullPath(dialog.FileName).Equals(Path.GetFullPath(csvPath), StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Нельзя записывать отчёт поверх исходного CSV.");
            SetBusy(true);
            var token = cancellation!.Token;
            var progress = new Progress<long>(count => status.Text = $"Обработано строк: {count:N0}");
            var result = await Task.Run(() => ReportExporter.Export(dialog.FileName, rows, template, pdf, token, progress));
            status.Text = $"Сохранено строк: {result.Rows:N0}. Итого: {result.Total:N2} руб.";
            MessageBox.Show("Отчёт сохранён.", "Экспорт");
        }
        catch (OperationCanceledException) { status.Text = "Экспорт отменён"; }
        catch (Exception e) { MessageBox.Show(e.Message, "Ошибка экспорта"); }
        finally { if (cancellation is not null) SetBusy(false); }
    }

    // Блокирует редактирование на время экспорта и включает кнопку отмены.
    private void SetBusy(bool busy)
    {
        toolbar.Enabled = !busy; settings.Enabled = !busy; grid.Enabled = !busy;
        cancel.Enabled = busy; progressBar.Visible = busy;
        if (busy) { cancellation = new CancellationTokenSource(); status.Text = "Обработка данных…"; }
        else { cancellation?.Dispose(); cancellation = null; }
    }
}
