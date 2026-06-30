using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace XlsxToXlsBatch;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

internal sealed class ConvertItem
{
    public required string SourcePath { get; init; }
    public string Status { get; set; } = "等待转换";
}

internal sealed class MainForm : Form
{
    private readonly DataGridView grid = new();
    private readonly BindingSource source = new();
    private readonly List<ConvertItem> items = [];
    private readonly RadioButton sameFolder = new() { Text = "保存到原文件夹", Checked = true, AutoSize = true };
    private readonly RadioButton chosenFolder = new() { Text = "保存到指定文件夹", AutoSize = true };
    private readonly TextBox outputPath = new() { ReadOnly = true, Enabled = false };
    private readonly Button browseOutput = new() { Text = "选择…", Enabled = false, AutoSize = true };
    private readonly CheckBox recursive = new() { Text = "包含子文件夹", Checked = true, AutoSize = true };
    private readonly CheckBox overwrite = new() { Text = "覆盖已存在的 .xls 文件", AutoSize = true };
    private readonly Button addFiles = new() { Text = "添加文件", AutoSize = true };
    private readonly Button addFolder = new() { Text = "添加文件夹", AutoSize = true };
    private readonly Button removeSelected = new() { Text = "移除选中", AutoSize = true };
    private readonly Button clearAll = new() { Text = "清空列表", AutoSize = true };
    private readonly Button convert = new() { Text = "开始转换", AutoSize = true, Height = 38 };
    private readonly Button openOutput = new() { Text = "打开输出位置", AutoSize = true, Height = 38 };
    private readonly ProgressBar progress = new() { Dock = DockStyle.Fill };
    private readonly Label summary = new() { Text = "尚未添加文件", AutoSize = true };
    private volatile bool converting;

    public MainForm()
    {
        Text = "XLSX 批量转 XLS";
        MinimumSize = new Size(820, 580);
        Size = new Size(980, 680);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Microsoft YaHei UI", 9F);

        BuildLayout();
        WireEvents();
        BindGrid();
    }

    private void BuildLayout()
    {
        var header = new Label
        {
            Text = "XLSX 批量转 XLS",
            Font = new Font(Font.FontFamily, 18F, FontStyle.Bold),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 4)
        };
        var hint = new Label
        {
            Text = "调用本机 Excel / WPS 转换，最大程度保留颜色、字体、边框、公式、图片和页面设置。",
            AutoSize = true,
            ForeColor = Color.DimGray
        };
        var titlePanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            AutoSize = true,
            WrapContents = false,
            Padding = new Padding(0, 0, 0, 8)
        };
        titlePanel.Controls.Add(header);
        titlePanel.Controls.Add(hint);

        var toolbar = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        toolbar.Controls.AddRange([addFiles, addFolder, recursive, removeSelected, clearAll]);

        grid.Dock = DockStyle.Fill;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.AllowUserToResizeRows = false;
        grid.AutoGenerateColumns = false;
        grid.BackgroundColor = SystemColors.Window;
        grid.BorderStyle = BorderStyle.Fixed3D;
        grid.MultiSelect = true;
        grid.ReadOnly = true;
        grid.RowHeadersVisible = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(ConvertItem.SourcePath),
            HeaderText = "源文件",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            MinimumWidth = 420
        });
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(ConvertItem.Status),
            HeaderText = "状态",
            Width = 280
        });

        var outputGroup = new GroupBox { Text = "输出设置", Dock = DockStyle.Fill, AutoSize = true };
        var outputLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 4,
            Padding = new Padding(8)
        };
        outputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        outputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        outputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        outputLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        outputLayout.Controls.Add(sameFolder, 0, 0);
        outputLayout.SetColumnSpan(sameFolder, 4);
        outputLayout.Controls.Add(chosenFolder, 0, 1);
        outputLayout.Controls.Add(outputPath, 2, 1);
        outputPath.Dock = DockStyle.Fill;
        outputLayout.Controls.Add(browseOutput, 3, 1);
        outputLayout.Controls.Add(overwrite, 0, 2);
        outputLayout.SetColumnSpan(overwrite, 4);
        outputGroup.Controls.Add(outputLayout);

        var bottom = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true };
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        var progressPanel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, AutoSize = true };
        progressPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        progressPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
        progressPanel.Controls.Add(summary, 0, 0);
        progressPanel.Controls.Add(progress, 0, 1);
        bottom.Controls.Add(progressPanel, 0, 0);
        bottom.Controls.Add(openOutput, 1, 0);
        bottom.Controls.Add(convert, 2, 0);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(18),
            RowCount = 5,
            ColumnCount = 1
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.Controls.Add(titlePanel, 0, 0);
        root.Controls.Add(toolbar, 0, 1);
        root.Controls.Add(grid, 0, 2);
        root.Controls.Add(outputGroup, 0, 3);
        root.Controls.Add(bottom, 0, 4);
        Controls.Add(root);
    }

    private void WireEvents()
    {
        addFiles.Click += (_, _) => AddFiles();
        addFolder.Click += (_, _) => AddFolder();
        removeSelected.Click += (_, _) => RemoveSelected();
        clearAll.Click += (_, _) => { items.Clear(); RefreshGrid(); };
        sameFolder.CheckedChanged += (_, _) => UpdateOutputControls();
        chosenFolder.CheckedChanged += (_, _) => UpdateOutputControls();
        browseOutput.Click += (_, _) => ChooseOutputFolder();
        openOutput.Click += (_, _) => OpenOutputLocation();
        convert.Click += (_, _) => StartConversion();
        FormClosing += (_, e) =>
        {
            if (!converting) return;
            e.Cancel = true;
            MessageBox.Show("正在转换，请等待当前批次完成。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        };
    }

    private void BindGrid()
    {
        source.DataSource = items;
        grid.DataSource = source;
    }

    private void AddFiles()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Excel 工作簿 (*.xlsx)|*.xlsx",
            Multiselect = true,
            Title = "选择要转换的 XLSX 文件"
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
            AddPaths(dialog.FileNames);
    }

    private void AddFolder()
    {
        using var dialog = new FolderBrowserDialog { Description = "选择包含 XLSX 文件的文件夹", UseDescriptionForTitle = true };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        try
        {
            var option = recursive.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            AddPaths(Directory.EnumerateFiles(dialog.SelectedPath, "*.xlsx", option));
        }
        catch (Exception ex)
        {
            MessageBox.Show($"读取文件夹失败：\n{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void AddPaths(IEnumerable<string> paths)
    {
        var existing = new HashSet<string>(items.Select(x => x.SourcePath), StringComparer.OrdinalIgnoreCase);
        foreach (var path in paths)
        {
            var full = Path.GetFullPath(path);
            if (existing.Add(full))
                items.Add(new ConvertItem { SourcePath = full });
        }
        RefreshGrid();
    }

    private void RemoveSelected()
    {
        var selected = grid.SelectedRows.Cast<DataGridViewRow>()
            .Select(r => r.DataBoundItem as ConvertItem)
            .Where(x => x is not null)
            .ToHashSet();
        items.RemoveAll(x => selected.Contains(x));
        RefreshGrid();
    }

    private void RefreshGrid()
    {
        source.ResetBindings(false);
        summary.Text = items.Count == 0 ? "尚未添加文件" : $"共 {items.Count} 个文件";
        progress.Value = 0;
        progress.Maximum = Math.Max(1, items.Count);
    }

    private void UpdateOutputControls()
    {
        outputPath.Enabled = chosenFolder.Checked;
        browseOutput.Enabled = chosenFolder.Checked;
    }

    private void ChooseOutputFolder()
    {
        using var dialog = new FolderBrowserDialog { Description = "选择 XLS 文件的保存文件夹", UseDescriptionForTitle = true };
        if (dialog.ShowDialog(this) == DialogResult.OK)
            outputPath.Text = dialog.SelectedPath;
    }

    private void OpenOutputLocation()
    {
        string? path = chosenFolder.Checked ? outputPath.Text :
            items.FirstOrDefault() is { } item ? Path.GetDirectoryName(item.SourcePath) : null;
        if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
        {
            MessageBox.Show("请先添加文件或选择有效的输出文件夹。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        Process.Start(new ProcessStartInfo("explorer.exe", $"\"{path}\"") { UseShellExecute = true });
    }

    private void StartConversion()
    {
        if (items.Count == 0)
        {
            MessageBox.Show("请先添加要转换的 XLSX 文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        if (chosenFolder.Checked && string.IsNullOrWhiteSpace(outputPath.Text))
        {
            MessageBox.Show("请选择输出文件夹。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (chosenFolder.Checked)
            Directory.CreateDirectory(outputPath.Text);

        SetBusy(true);
        foreach (var item in items) item.Status = "等待转换";
        source.ResetBindings(false);
        progress.Maximum = items.Count;
        progress.Value = 0;

        var batch = items.ToArray();
        var destination = chosenFolder.Checked ? outputPath.Text : null;
        var allowOverwrite = overwrite.Checked;
        var thread = new Thread(() => ConvertBatch(batch, destination, allowOverwrite))
        {
            IsBackground = true,
            Name = "ExcelConverter"
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    private void ConvertBatch(ConvertItem[] batch, string? destinationFolder, bool allowOverwrite)
    {
        int success = 0, failed = 0, skipped = 0;

        try
        {
            // ── 构造批次任务 ──
            var tasks = new List<object>();
            foreach (var item in batch)
            {
                var folder = destinationFolder ?? Path.GetDirectoryName(item.SourcePath)!;
                var target = Path.Combine(folder, Path.GetFileNameWithoutExtension(item.SourcePath) + ".xls");
                tasks.Add(new
                {
                    Source = item.SourcePath,
                    Target = target,
                    TempTarget = Path.Combine(folder,
                        $".{Path.GetFileNameWithoutExtension(item.SourcePath)}.{Guid.NewGuid():N}.tmp.xls")
                });
            }

            // ── 写入输入 JSON ──
            var inputJson = System.Text.Json.JsonSerializer.Serialize(new
            {
                Tasks = tasks,
                AllowOverwrite = allowOverwrite
            });
            var jsonDir = Path.Combine(Path.GetTempPath(), "XlsxBatchConverter", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(jsonDir);
            var inputFile = Path.Combine(jsonDir, "input.json");
            var outputFile = Path.Combine(jsonDir, "output.json");
            var psScript = Path.Combine(jsonDir, "convert.ps1");
            File.WriteAllText(inputFile, inputJson);

            // ── 写入 PowerShell 脚本 ──
            File.WriteAllText(psScript, @"
param($InputFile, $OutputFile)
$ErrorActionPreference = 'Stop'
$paths = Get-Content $InputFile -Raw | ConvertFrom-Json

# 尝试 WPS(国产版) 再尝试 WPS(国际版) 最后尝试 Excel
$progIds = @('KET.Application', 'ET.Application', 'Excel.Application')
$app = $null
foreach ($id in $progIds) {
    try {
        $app = New-Object -ComObject $id
        if ($app) { break }
    } catch { }
}
if (-not $app) {
    @{ Results = @(@{ Index = -1; Status = 'FATAL'; Error = '未找到可用的 WPS/Excel COM 组件' }) } | ConvertTo-Json -Compress | Out-File $OutputFile -Encoding UTF8
    exit 1
}

$app.Visible = $false
$app.DisplayAlerts = $false

$results = @()
for ($i = 0; $i -lt $paths.Tasks.Count; $i++) {
    $task = $paths.Tasks[$i]
    $src = $task.Source
    $tgt = $task.TempTarget
    try {
        if (-not (Test-Path $src)) {
            $results += @{ Index = $i; Status = 'FAIL'; Error = '源文件不存在' }
            continue
        }
        if ((Test-Path $task.Target) -and -not $paths.AllowOverwrite) {
            $results += @{ Index = $i; Status = 'SKIP' }
            continue
        }
        $wb = $app.Workbooks.Open($src, 0, $true, 1)
        $wb.SaveAs($tgt, 56)
        $wb.Close($false)
        [Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null

        # 成功后才替换目标文件
        if (Test-Path $task.Target) { Remove-Item $task.Target -Force }
        Move-Item $tgt $task.Target -Force
        $results += @{ Index = $i; Status = 'OK' }
    } catch {
        if (Test-Path $tgt) { Remove-Item $tgt -Force -ErrorAction SilentlyContinue }
        $results += @{ Index = $i; Status = 'FAIL'; Error = $_.Exception.Message }
    }
}

try { $app.Quit() } catch { }
[Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
[GC]::Collect()

@{ Results = $results } | ConvertTo-Json -Compress | Out-File $OutputFile -Encoding UTF8
");

            // ── 执行 PowerShell ──
            var psi = new ProcessStartInfo("powershell.exe")
            {
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{psScript}\" \"{inputFile}\" \"{outputFile}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            using var process = Process.Start(psi)!;
            process.WaitForExit(120000);
            var stderr = process.StandardError.ReadToEnd();

            // ── 读取输出 JSON ──
            if (!File.Exists(outputFile))
                throw new InvalidOperationException(
                    $"PowerShell 未产生输出文件。\n错误输出：{stderr}");

            var outputJson = File.ReadAllText(outputFile);
            var resultDoc = System.Text.Json.JsonSerializer.Deserialize<BatchResult>(outputJson);

            // ── 更新每个文件状态 ──
            foreach (var r in resultDoc?.Results ?? [])
            {
                if (r.Index < 0 || r.Index >= batch.Length) continue;
                var item = batch[r.Index];
                switch (r.Status)
                {
                    case "OK":
                        item.Status = "转换成功";
                        success++;
                        break;
                    case "SKIP":
                        item.Status = "已跳过（目标文件存在）";
                        skipped++;
                        break;
                    default:
                        item.Status = "失败：" + (r.Error ?? "未知错误");
                        failed++;
                        break;
                }
                // 略作延时让 UI 能刷新
                BeginInvoke(() =>
                {
                    source.ResetBindings(false);
                    progress.Value = Math.Min(progress.Maximum, progress.Value + 1);
                    summary.Text = $"正在处理 {progress.Value} / {progress.Maximum}";
                });
                Thread.Sleep(30);
            }

            // ── 清理临时目录 ──
            try { Directory.Delete(jsonDir, true); } catch { }
        }
        catch (Exception ex)
        {
            foreach (var item in batch.Where(x => x.Status is "等待转换" or "正在转换…"))
                item.Status = "失败：" + CleanError(ex);
            BeginInvoke(() => source.ResetBindings(false));
        }
        finally
        {
            // 全量释放 COM（调用 GC 确保 PowerShell 进程退出）
            GC.Collect();
            GC.WaitForPendingFinalizers();

            BeginInvoke(() =>
            {
                SetBusy(false);
                summary.Text = $"完成：成功 {success}，失败 {failed}，跳过 {skipped}";
                var icon = failed == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning;
                MessageBox.Show(this, summary.Text, "转换完成", MessageBoxButtons.OK, icon);
            });
        }
    }

    // PowerShell 输出 JSON 的反序列化类型
    private sealed record BatchResult(ResultItem[]? Results);
    private sealed record ResultItem(int Index, string Status, string? Error);

    private void UpdateItem(ConvertItem item, string status)
    {
        item.Status = status;
        BeginInvoke(() => source.ResetBindings(false));
    }

    private void UpdateProgress(ConvertItem item)
    {
        BeginInvoke(() =>
        {
            source.ResetBindings(false);
            progress.Value = Math.Min(progress.Maximum, progress.Value + 1);
            summary.Text = $"正在处理 {progress.Value} / {progress.Maximum}";
        });
    }

    private void SetBusy(bool busy)
    {
        converting = busy;
        addFiles.Enabled = !busy;
        addFolder.Enabled = !busy;
        removeSelected.Enabled = !busy;
        clearAll.Enabled = !busy;
        sameFolder.Enabled = !busy;
        chosenFolder.Enabled = !busy;
        recursive.Enabled = !busy;
        overwrite.Enabled = !busy;
        browseOutput.Enabled = !busy && chosenFolder.Checked;
        convert.Enabled = !busy;
        convert.Text = busy ? "正在转换…" : "开始转换";
    }

    private static string CleanError(Exception ex)
    {
        while (ex.InnerException is not null) ex = ex.InnerException;
        var message = ex.Message.Replace("\r", " ").Replace("\n", " ").Trim();
        if (ex is COMException comException)
            message = $"{message} (0x{comException.ErrorCode:X8})";
        return message.Length > 140 ? message[..140] + "…" : message;
    }
}
