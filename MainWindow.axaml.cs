using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Input;
using Avalonia.Platform.Storage;
using Avalonia.VisualTree;
using Source2Docx.Core.Services;
using Source2Docx.Models;

namespace Source2Docx;

public partial class MainWindow : Window
{
    private const int DefaultTrimPageCount = 50;

    private const int MinTrimPageCount = 1;

    private const int MaxTrimPageCount = 100;

    private const int FixedLinesPerPage = 50;

    private readonly ObservableCollection<SourceFileItem> files = [];

    private readonly List<CodeTypeOption> codeTypes =
    [
        new("C", [".h", ".c", ".cu"]),
        new("C++", [".h", ".cpp", ".c", ".cu", ".hpp", ".cuh"]),
        new("C#", [".cs"]),
        new("C#/WPF", [".cs", ".xaml"]),
        new("Java/Eclipse", [".java", ".xml"]),
        new("Python", [".py", ".pyw"]),
        new("JavaScript/TypeScript", [".js", ".jsx", ".ts", ".tsx"]),
        new("Go", [".go"]),
        new("Rust", [".rs"]),
        new("Kotlin", [".kt", ".kts"]),
        new("Swift", [".swift"]),
        new("PHP", [".php"])
    ];

    private readonly SourceDocumentGenerator generator = new();

    private static readonly DataFormat<List<SourceFileItem>> FileListDragDataFormat =
        DataFormat.CreateInProcessFormat<List<SourceFileItem>>("Source2Docx.FileListRows");

    private CancellationTokenSource generationCts;

    private Point dragStartPoint;

    private SourceFileItem dragAnchorItem;

    private PointerPressedEventArgs dragStartEventArgs;

    private bool isGenerating;

    public MainWindow()
    {
        InitializeComponent();
        FilesListBox.ItemsSource = files;
        DragDrop.SetAllowDrop(FilesListBox, true);
        DragDrop.AddDragOverHandler(FilesListBox, FilesListBox_DragOver);
        DragDrop.AddDropHandler(FilesListBox, FilesListBox_Drop);
        CodeTypeComboBox.ItemsSource = codeTypes;
        CodeTypeComboBox.SelectedIndex = 0;
        SoftwareNameTextBox.Text = "Source2Docx";
        VersionTextBox.Text = "V1.0.0";
        OutputPathTextBox.Text = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            "AllCode.docx");
        TrimPageCountTextBox.Text = DefaultTrimPageCount.ToString();
        ProgressTextBlock.Text = "0 / 0";
        SetStatus("准备就绪。");
        UpdateTrimPageControls();
        UpdateFileSummary();
        UpdateSelectionHint();
    }

    private CodeTypeOption CurrentCodeType =>
        CodeTypeComboBox.SelectedItem as CodeTypeOption ?? codeTypes[0];

    private void TitleBar_PointerPressed(object sender, PointerPressedEventArgs e)
    {
        PointerPointProperties properties = e.GetCurrentPoint(this).Properties;
        if (!properties.IsLeftButtonPressed)
        {
            return;
        }

        if (e.ClickCount == 2)
        {
            WindowState = WindowState == WindowState.Maximized
                ? WindowState.Normal
                : WindowState.Maximized;
            return;
        }

        BeginMoveDrag(e);
    }

    private async void BrowseOutput_Click(object sender, RoutedEventArgs e)
    {
        TopLevel topLevel = TopLevel.GetTopLevel(this);
        if (topLevel?.StorageProvider == null)
        {
            SetStatus("当前环境不支持文件保存对话框。");
            return;
        }

        IStorageFile file = await topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "选择输出文档",
            SuggestedFileName = Path.GetFileName(NormalizeOutputPath(OutputPathTextBox.Text)),
            DefaultExtension = "docx",
            FileTypeChoices =
            [
                new FilePickerFileType("Word 文档")
                {
                    Patterns = ["*.docx"]
                }
            ]
        });

        string path = file?.TryGetLocalPath();
        if (!string.IsNullOrWhiteSpace(path))
        {
            OutputPathTextBox.Text = path;
        }
    }

    private async void AddFiles_Click(object sender, RoutedEventArgs e)
    {
        IReadOnlyList<string> selectedPaths = await PickFilesAsync(true);
        if (selectedPaths.Count == 0)
        {
            SetStatus($"未选中任何 {CurrentCodeType.DisplayName} 文件。");
            return;
        }

        foreach (string path in selectedPaths)
        {
            AddFile(path);
        }

        UpdateFileSummary();
        UpdateSelectionHint();
    }

    private async void AddFolder_Click(object sender, RoutedEventArgs e)
    {
        IReadOnlyList<string> selectedPaths = await PickFilesAsync(false);
        string mainFile = selectedPaths.FirstOrDefault();
        if (string.IsNullOrWhiteSpace(mainFile))
        {
            SetStatus($"未选中任何 {CurrentCodeType.DisplayName} 主文件。");
            return;
        }

        string rootDirectory = Path.GetDirectoryName(mainFile);
        if (string.IsNullOrWhiteSpace(rootDirectory) || !Directory.Exists(rootDirectory))
        {
            SetStatus("所选主文件目录无效。");
            return;
        }

        AddFile(mainFile);
        foreach (string file in EnumerateMatchingFiles(rootDirectory, CurrentCodeType.Extensions)
                     .Where(path => !path.Equals(mainFile, StringComparison.OrdinalIgnoreCase))
                     .OrderBy(static path => path, StringComparer.OrdinalIgnoreCase))
        {
            AddFile(file);
        }

        MoveItemToTop(mainFile);
        UpdateFileSummary();
        UpdateSelectionHint();
    }

    private async void ImportFolder_Click(object sender, RoutedEventArgs e)
    {
        string directoryPath = await PickFolderAsync();
        if (string.IsNullOrWhiteSpace(directoryPath))
        {
            SetStatus("未选中任何文件夹。");
            return;
        }

        List<string> matchedFiles = EnumerateMatchingFiles(directoryPath, CurrentCodeType.Extensions)
            .OrderBy(static path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (matchedFiles.Count == 0)
        {
            SetStatus($"所选文件夹中没有匹配 {CurrentCodeType.DisplayName} 的文件。");
            return;
        }

        foreach (string file in matchedFiles)
        {
            AddFile(file);
        }

        UpdateFileSummary();
        UpdateSelectionHint();
        SetStatus($"已从文件夹导入 {matchedFiles.Count} 个 {CurrentCodeType.DisplayName} 文件。");
    }

    private void ToggleCheckAll_Click(object sender, RoutedEventArgs e)
    {
        if (files.Count == 0)
        {
            return;
        }

        bool shouldCheckAll = files.Any(static file => !file.IsChecked);
        foreach (SourceFileItem file in files)
        {
            file.IsChecked = shouldCheckAll;
        }

        UpdateFileSummary();
    }

    private void RemoveChecked_Click(object sender, RoutedEventArgs e)
    {
        List<SourceFileItem> checkedItems = files.Where(static file => file.IsChecked).ToList();
        if (checkedItems.Count == 0)
        {
            SetStatus("先勾选要删除的文件。");
            return;
        }

        foreach (SourceFileItem item in checkedItems)
        {
            files.Remove(item);
        }

        UpdateFileSummary();
        UpdateSelectionHint();
    }

    private void MoveTop_Click(object sender, RoutedEventArgs e)
    {
        List<SourceFileItem> selectedItems = GetOrderedSelectedItems();
        if (selectedItems.Count == 0)
        {
            SetStatus("请先选中至少一个文件。");
            return;
        }

        for (int index = 0; index < selectedItems.Count; index++)
        {
            files.Remove(selectedItems[index]);
        }

        for (int index = 0; index < selectedItems.Count; index++)
        {
            files.Insert(index, selectedItems[index]);
        }

        ReselectItems(selectedItems);
        UpdateSelectionHint();
    }

    private void MoveUp_Click(object sender, RoutedEventArgs e)
    {
        MoveSelectedItems(-1);
    }

    private void MoveDown_Click(object sender, RoutedEventArgs e)
    {
        MoveSelectedItems(1);
    }

    private async void Generate_Click(object sender, RoutedEventArgs e)
    {
        if (isGenerating)
        {
            generationCts?.Cancel();
            SetStatus("正在停止生成，请稍候...");
            return;
        }

        if (files.Count == 0)
        {
            SetStatus("请先添加至少一个源文件。");
            return;
        }

        string softwareName = SoftwareNameTextBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(softwareName))
        {
            SetStatus("软件名称不能为空。");
            return;
        }

        string version = VersionTextBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(version))
        {
            version = "V1.0.0";
            VersionTextBox.Text = version;
        }

        string outputPath;
        try
        {
            outputPath = NormalizeOutputPath(OutputPathTextBox.Text);
        }
        catch (Exception ex)
        {
            SetStatus("输出路径无效: " + ex.Message);
            return;
        }

        if (!TryGetTrimPageCount(out int trimPageCount))
        {
            return;
        }

        List<string> orderedFiles = files.Select(static file => file.FullPath).ToList();
        FixedLineTrimOptions trimOptions = TrimPagesCheckBox.IsChecked == true
            ? new FixedLineTrimOptions(trimPageCount, FixedLinesPerPage)
            : null;

        generationCts = new CancellationTokenSource();
        SetGeneratingState(true, orderedFiles.Count);
        IProgress<int> progress = new Progress<int>(value =>
        {
            GenerationProgressBar.Value = value;
            ProgressTextBlock.Text = $"{value} / {orderedFiles.Count}";
            SetStatus($"正在处理源文件，已完成 {value} / {orderedFiles.Count}。");
        });

        try
        {
            DocumentGenerationResult result = await generator.GenerateAsync(
                softwareName,
                version,
                CurrentCodeType.DisplayName,
                outputPath,
                orderedFiles,
                trimOptions,
                progress,
                generationCts.Token);

            OutputPathTextBox.Text = result.OutputPath;

            if (result.WasCanceled)
            {
                SetStatus($"生成已停止，已保留前 {result.ProcessedCount} 个文件的结果。");
                return;
            }

            GenerationProgressBar.Value = orderedFiles.Count;
            ProgressTextBlock.Text = $"{orderedFiles.Count} / {orderedFiles.Count}";
            SetStatus(BuildSuccessMessage(result, trimOptions));

            if (File.Exists(result.OutputPath))
            {
                Process.Start(new ProcessStartInfo(result.OutputPath) { UseShellExecute = true });
            }
        }
        catch (OperationCanceledException)
        {
            SetStatus("生成已取消。");
        }
        catch (Exception ex)
        {
            SetStatus("生成失败: " + ex.Message);
        }
        finally
        {
            generationCts?.Dispose();
            generationCts = null;
            SetGeneratingState(false, orderedFiles.Count);
        }
    }

    private void TrimPagesCheckBox_Click(object sender, RoutedEventArgs e)
    {
        UpdateTrimPageControls();
    }

    private void TrimPageCountTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (int.TryParse(TrimPageCountTextBox.Text?.Trim(), out int pageCount))
        {
            TrimPageCountTextBox.Text = Math.Clamp(pageCount, MinTrimPageCount, MaxTrimPageCount).ToString();
            return;
        }

        TrimPageCountTextBox.Text = DefaultTrimPageCount.ToString();
    }

    private void FilesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateSelectionHint();
    }

    private void FilesListBox_PointerPressed(object sender, PointerPressedEventArgs e)
    {
        if (!e.GetCurrentPoint(FilesListBox).Properties.IsLeftButtonPressed)
        {
            dragAnchorItem = null;
            dragStartEventArgs = null;
            return;
        }

        dragStartPoint = e.GetPosition(FilesListBox);
        dragAnchorItem = GetItemFromSource(e.Source as Visual);
        dragStartEventArgs = e;
    }

    private async void FilesListBox_PointerMoved(object sender, PointerEventArgs e)
    {
        if (!e.GetCurrentPoint(FilesListBox).Properties.IsLeftButtonPressed || dragAnchorItem == null)
        {
            return;
        }

        Point currentPosition = e.GetPosition(FilesListBox);
        if (Math.Abs(currentPosition.X - dragStartPoint.X) < 4 &&
            Math.Abs(currentPosition.Y - dragStartPoint.Y) < 4)
        {
            return;
        }

        List<SourceFileItem> draggedItems = GetOrderedSelectedItems();
        if (!draggedItems.Contains(dragAnchorItem))
        {
            draggedItems = [dragAnchorItem];
            ReselectItems(draggedItems);
        }

        dragAnchorItem = null;

        if (dragStartEventArgs == null)
        {
            return;
        }

        DataTransfer data = new();
        data.Add(DataTransferItem.Create(FileListDragDataFormat, draggedItems));
        await DragDrop.DoDragDropAsync(dragStartEventArgs, data, DragDropEffects.Move);
        dragStartEventArgs = null;
    }

    private void FilesListBox_DragOver(object sender, DragEventArgs e)
    {
        e.DragEffects = e.DataTransfer.Contains(FileListDragDataFormat)
            ? DragDropEffects.Move
            : DragDropEffects.None;
    }

    private void FilesListBox_Drop(object sender, DragEventArgs e)
    {
        if (!e.DataTransfer.Contains(FileListDragDataFormat))
        {
            return;
        }

        List<SourceFileItem> draggedItems = e.DataTransfer.TryGetValue(FileListDragDataFormat);
        if (draggedItems == null || draggedItems.Count == 0)
        {
            return;
        }

        Control targetControl = (e.Source as Visual)?.FindAncestorOfType<ListBoxItem>();
        SourceFileItem targetItem = targetControl?.DataContext as SourceFileItem;
        if (targetItem != null && draggedItems.Contains(targetItem))
        {
            return;
        }

        bool insertAfter = false;
        if (targetControl != null)
        {
            Point position = e.GetPosition(targetControl);
            insertAfter = position.Y > targetControl.Bounds.Height / 2;
        }

        MoveItemsToDropPosition(draggedItems, targetItem, insertAfter);
    }

    private async System.Threading.Tasks.Task<IReadOnlyList<string>> PickFilesAsync(bool allowMultiple)
    {
        TopLevel topLevel = TopLevel.GetTopLevel(this);
        if (topLevel?.StorageProvider == null)
        {
            SetStatus("当前环境不支持文件选择对话框。");
            return Array.Empty<string>();
        }

        IReadOnlyList<IStorageFile> pickedFiles = await topLevel.StorageProvider.OpenFilePickerAsync(
            new FilePickerOpenOptions
            {
                Title = allowMultiple ? "添加源代码文件" : "选择主文件并递归导入同目录源码",
                AllowMultiple = allowMultiple
            });

        List<string> allPaths = pickedFiles
            .Select(static file => file.TryGetLocalPath())
            .Where(static path => !string.IsNullOrWhiteSpace(path))
            .ToList();

        List<string> filteredPaths = allPaths
            .Where(path => CurrentCodeType.Extensions.Contains(Path.GetExtension(path)))
            .ToList();

        if (allPaths.Count > 0 && filteredPaths.Count == 0)
        {
            SetStatus($"所选文件都不是 {CurrentCodeType.DisplayName} 支持的类型: {CurrentCodeType.PickerPattern}");
        }

        return filteredPaths;
    }

    private async System.Threading.Tasks.Task<string> PickFolderAsync()
    {
        TopLevel topLevel = TopLevel.GetTopLevel(this);
        if (topLevel?.StorageProvider == null)
        {
            SetStatus("当前环境不支持文件夹选择对话框。");
            return string.Empty;
        }

        IReadOnlyList<IStorageFolder> pickedFolders = await topLevel.StorageProvider.OpenFolderPickerAsync(
            new FolderPickerOpenOptions
            {
                Title = "选择要导入的文件夹",
                AllowMultiple = false
            });

        return pickedFolders.FirstOrDefault()?.TryGetLocalPath() ?? string.Empty;
    }

    private void AddFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return;
        }

        if (files.Any(file => file.FullPath.Equals(filePath, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        files.Add(new SourceFileItem
        {
            Name = Path.GetFileName(filePath),
            FullPath = filePath
        });
    }

    private void MoveItemToTop(string filePath)
    {
        SourceFileItem item = files.FirstOrDefault(file =>
            file.FullPath.Equals(filePath, StringComparison.OrdinalIgnoreCase));
        if (item == null)
        {
            return;
        }

        int index = files.IndexOf(item);
        if (index <= 0)
        {
            return;
        }

        files.Move(index, 0);
        ReselectItems([item]);
    }

    private void MoveSelectedItems(int offset)
    {
        List<SourceFileItem> selectedItems = GetOrderedSelectedItems();
        if (selectedItems.Count == 0)
        {
            SetStatus("请先选中至少一个文件。");
            return;
        }

        HashSet<SourceFileItem> selectedSet = [.. selectedItems];
        bool moved = false;
        if (offset < 0)
        {
            foreach (SourceFileItem item in selectedItems)
            {
                int currentIndex = files.IndexOf(item);
                if (currentIndex <= 0 || selectedSet.Contains(files[currentIndex - 1]))
                {
                    continue;
                }

                files.Move(currentIndex, currentIndex - 1);
                moved = true;
            }
        }
        else if (offset > 0)
        {
            for (int index = selectedItems.Count - 1; index >= 0; index--)
            {
                SourceFileItem item = selectedItems[index];
                int currentIndex = files.IndexOf(item);
                if (currentIndex < 0 || currentIndex >= files.Count - 1 || selectedSet.Contains(files[currentIndex + 1]))
                {
                    continue;
                }

                files.Move(currentIndex, currentIndex + 1);
                moved = true;
            }
        }

        if (moved)
        {
            ReselectItems(selectedItems);
        }

        UpdateSelectionHint();
    }

    private void MoveItemsToDropPosition(List<SourceFileItem> draggedItems, SourceFileItem targetItem, bool insertAfter)
    {
        if (draggedItems.Count == 0)
        {
            return;
        }

        int insertIndex = targetItem == null ? files.Count : files.IndexOf(targetItem);
        if (insertIndex < 0)
        {
            insertIndex = files.Count;
        }

        if (targetItem != null && insertAfter)
        {
            insertIndex++;
        }

        foreach (SourceFileItem item in draggedItems)
        {
            files.Remove(item);
        }

        if (targetItem != null)
        {
            insertIndex = files.IndexOf(targetItem);
            if (insertIndex < 0)
            {
                insertIndex = files.Count;
            }

            if (insertAfter)
            {
                insertIndex++;
            }
        }
        else
        {
            insertIndex = files.Count;
        }

        for (int index = 0; index < draggedItems.Count; index++)
        {
            files.Insert(Math.Min(insertIndex + index, files.Count), draggedItems[index]);
        }

        ReselectItems(draggedItems);
        UpdateSelectionHint();
    }

    private List<SourceFileItem> GetOrderedSelectedItems()
    {
        return FilesListBox.SelectedItems?
            .OfType<SourceFileItem>()
            .OrderBy(files.IndexOf)
            .ToList() ?? [];
    }

    private void ReselectItems(IEnumerable<SourceFileItem> items)
    {
        if (FilesListBox.SelectedItems == null)
        {
            return;
        }

        FilesListBox.SelectedItems.Clear();
        foreach (SourceFileItem item in items)
        {
            FilesListBox.SelectedItems.Add(item);
        }
    }

    private void UpdateTrimPageControls()
    {
        bool enabled = TrimPagesCheckBox.IsChecked == true && !isGenerating;
        TrimPageCountTextBox.IsEnabled = enabled;
    }

    private void UpdateFileSummary()
    {
        bool hasFiles = files.Count > 0;
        int checkedCount = files.Count(static file => file.IsChecked);
        bool allChecked = hasFiles && checkedCount == files.Count;

        FileSummaryTextBlock.Text = hasFiles
            ? $"共 {files.Count} 个文件，已勾选 {checkedCount} 个。"
            : "尚未添加文件";
        ToggleCheckAllButton.Content = allChecked ? "全部取消" : "全部勾选";
    }

    private void UpdateSelectionHint()
    {
        List<SourceFileItem> selectedItems = GetOrderedSelectedItems();
        SelectionHintTextBlock.Text = selectedItems.Count switch
        {
            0 => "当前未选中任何行",
            1 => "当前选中: " + selectedItems[0].Name,
            _ => $"当前选中 {selectedItems.Count} 行"
        };

        UpdateFileSummary();
    }

    private void SetGeneratingState(bool generating, int totalCount)
    {
        isGenerating = generating;
        GenerateButton.Content = generating ? "停止生成" : "开始生成";
        AddFilesButton.IsEnabled = !generating;
        ImportFolderButton.IsEnabled = !generating;
        AddFolderButton.IsEnabled = !generating;
        ToggleCheckAllButton.IsEnabled = !generating && files.Count > 0;
        RemoveCheckedButton.IsEnabled = !generating;
        MoveTopButton.IsEnabled = !generating;
        MoveUpButton.IsEnabled = !generating;
        MoveDownButton.IsEnabled = !generating;
        BrowseOutputButton.IsEnabled = !generating;
        CodeTypeComboBox.IsEnabled = !generating;
        TrimPagesCheckBox.IsEnabled = !generating;
        SoftwareNameTextBox.IsEnabled = !generating;
        VersionTextBox.IsEnabled = !generating;
        OutputPathTextBox.IsEnabled = !generating;
        UpdateTrimPageControls();

        GenerationProgressBar.Maximum = Math.Max(totalCount, 1);
        if (generating)
        {
            GenerationProgressBar.Value = 0;
            ProgressTextBlock.Text = $"0 / {totalCount}";
            SetStatus("准备开始生成...");
            return;
        }

        if (totalCount == 0)
        {
            ProgressTextBlock.Text = "0 / 0";
        }
    }

    private void SetStatus(string message)
    {
        StatusTextBlock.Text = $"[{CurrentCodeType.DisplayName}] {message}";
    }

    private bool TryGetTrimPageCount(out int trimPageCount)
    {
        trimPageCount = DefaultTrimPageCount;
        if (TrimPagesCheckBox.IsChecked != true)
        {
            return true;
        }

        if (!int.TryParse(TrimPageCountTextBox.Text?.Trim(), out trimPageCount))
        {
            SetStatus("裁剪页数必须是整数。");
            return false;
        }

        trimPageCount = Math.Clamp(trimPageCount, MinTrimPageCount, MaxTrimPageCount);
        TrimPageCountTextBox.Text = trimPageCount.ToString();
        return true;
    }

    private static IEnumerable<string> EnumerateMatchingFiles(string rootDirectory, HashSet<string> extensions)
    {
        IEnumerable<string> filesUnderRoot = Directory.EnumerateFiles(rootDirectory, "*", SearchOption.AllDirectories);
        foreach (string file in filesUnderRoot)
        {
            string extension = Path.GetExtension(file);
            if (extensions.Contains(extension))
            {
                yield return file;
            }
        }
    }

    private static string NormalizeOutputPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new InvalidOperationException("输出路径不能为空。");
        }

        string normalizedPath = Path.GetFullPath(path.Trim());
        if (!normalizedPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
        {
            normalizedPath += ".docx";
        }

        return normalizedPath;
    }

    private static string BuildSuccessMessage(DocumentGenerationResult result, FixedLineTrimOptions trimOptions)
    {
        if (trimOptions?.IsEnabled != true || !result.UsedFixedLineTrim)
        {
            return $"生成完成，已写入 {result.ProcessedCount} 个文件。";
        }

        return $"生成完成，已按固定 {trimOptions.LinesPerPage} 行/页保留前后各 {trimOptions.PageCount} 页，共扫描 {result.TotalLineCount} 行。";
    }

    private SourceFileItem GetItemFromSource(Avalonia.Visual source)
    {
        return source?
            .FindAncestorOfType<ListBoxItem>()?
            .DataContext as SourceFileItem;
    }
}
