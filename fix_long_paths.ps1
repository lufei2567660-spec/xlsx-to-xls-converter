$ErrorActionPreference = "Stop"

$path = Join-Path $PSScriptRoot "XlsxToXlsBatch\Program.cs"
$utf8 = [Text.UTF8Encoding]::new($false)
$text = [IO.File]::ReadAllText($path).Replace("`r`n", "`n")

function Replace-Required([string]$old, [string]$new) {
    $old = $old.Replace("`r`n", "`n")
    $new = $new.Replace("`r`n", "`n")
    if (-not $script:text.Contains($old)) {
        throw ("Expected source block was not found: " + $old.Substring(0, [Math]::Min(60, $old.Length)).Replace("`n", " | "))
    }
    $script:text = $script:text.Replace($old, $new)
}

Replace-Required @'
                object? workbook = null;
                try
'@ @'
                object? workbook = null;
                string? tempFolder = null;
                try
'@

Replace-Required @'
                    dynamic books = app.Workbooks;
                    workbook = books.Open(item.SourcePath, UpdateLinks: 0, ReadOnly: true, IgnoreReadOnlyRecommended: true);
'@ @'
                    var opened = OpenWorkbookWithLongPathRetry(app, item.SourcePath);
                    workbook = opened.Workbook;
                    tempFolder = opened.TempFolder;
'@

Replace-Required @'
                    book.SaveAs(target, FileFormat: 56, ConflictResolution: 2, Local: true);
'@ @'
                    SaveWorkbookWithLongPathRetry(book, target);
'@

Replace-Required @'
                    item.Status = "失败：" + CleanError(ex);
                    failed++;
                }
                UpdateProgress(item);
'@ @'
                    item.Status = "失败：" + CleanError(ex);
                    failed++;
                }
                finally
                {
                    if (tempFolder is not null)
                    {
                        try { Directory.Delete(tempFolder, true); } catch { }
                    }
                }
                UpdateProgress(item);
'@

Replace-Required @'
                app.Visible = false;
                app.DisplayAlerts = false;
                return instance;
'@ @'
                app.Visible = false;
                app.DisplayAlerts = false;
                object books = app.Workbooks;
                try { _ = ((dynamic)books).Count; }
                finally { ReleaseCom(books); }
                return instance;
'@

$marker = @'
    private static string CleanError(Exception ex)
'@
$helpers = @'
    private static (object Workbook, string? TempFolder) OpenWorkbookWithLongPathRetry(dynamic app, string sourcePath)
    {
        object books = app.Workbooks;
        try
        {
            try
            {
                object workbook = ((dynamic)books).Open(
                    sourcePath, UpdateLinks: 0, ReadOnly: true, IgnoreReadOnlyRecommended: true);
                return (workbook, null);
            }
            catch (Exception directError)
            {
                var tempFolder = Path.Combine(
                    Path.GetTempPath(), "XlsxBatchConverter", Guid.NewGuid().ToString("N"));
                try
                {
                    Directory.CreateDirectory(tempFolder);
                    var shortSource = Path.Combine(tempFolder, "input.xlsx");
                    File.Copy(sourcePath, shortSource, true);
                    object workbook = ((dynamic)books).Open(
                        shortSource, UpdateLinks: 0, ReadOnly: true, IgnoreReadOnlyRecommended: true);
                    return (workbook, tempFolder);
                }
                catch (Exception retryError)
                {
                    try { Directory.Delete(tempFolder, true); } catch { }
                    throw new InvalidOperationException(
                        $"无法打开源文件。直接打开：{CleanError(directError)}；短路径重试：{CleanError(retryError)}");
                }
            }
        }
        finally
        {
            ReleaseCom(books);
        }
    }

    private static void SaveWorkbookWithLongPathRetry(dynamic book, string targetPath)
    {
        try
        {
            book.SaveAs(targetPath, FileFormat: 56, ConflictResolution: 2, Local: true);
        }
        catch (Exception directError)
        {
            var tempFolder = Path.Combine(
                Path.GetTempPath(), "XlsxBatchConverter", Guid.NewGuid().ToString("N"));
            try
            {
                Directory.CreateDirectory(tempFolder);
                var shortTarget = Path.Combine(tempFolder, "output.xls");
                book.SaveAs(shortTarget, FileFormat: 56, ConflictResolution: 2, Local: true);
                File.Copy(shortTarget, targetPath, true);
            }
            catch (Exception retryError)
            {
                throw new InvalidOperationException(
                    $"无法保存目标文件。直接保存：{CleanError(directError)}；短路径重试：{CleanError(retryError)}");
            }
            finally
            {
                try { Directory.Delete(tempFolder, true); } catch { }
            }
        }
    }

    private static string CleanError(Exception ex)
'@
Replace-Required $marker $helpers

Replace-Required `
    'return message.Length > 90 ? message[..90] + "…" : message;' `
    'return message.Length > 180 ? message[..180] + "…" : message;'

[IO.File]::WriteAllText($path, $text, $utf8)
