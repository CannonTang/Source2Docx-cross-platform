using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Source2Docx.Core.Documents;
using Source2Docx.Core.Text;

namespace Source2Docx.Core.Services;

internal sealed class SourceDocumentGenerator
{
    private const int ProgressReportIntervalMilliseconds = 120;

    private const int ProgressReportStride = 8;

    public Task<DocumentGenerationResult> GenerateAsync(
        string softwareName,
        string version,
        string codeType,
        string outputPath,
        IReadOnlyList<string> files,
        FixedLineTrimOptions trimOptions,
        IProgress<int> progress,
        CancellationToken cancellationToken)
    {
        return Task.Run(() =>
        {
            string title = $"{softwareName}   {version}     源代码";
            SourceTextCleaner cleaner = CreateCleaner(codeType);
            SourceDocumentBuilder documentBuilder = new(outputPath, title, cleaner);
            FixedLineTrimBuffer trimBuffer = trimOptions?.IsEnabled == true
                ? new FixedLineTrimBuffer(trimOptions.PageCount, trimOptions.LinesPerPage)
                : null;

            int processedCount = 0;
            int lastReportedCount = 0;
            bool wasCanceled = false;
            Stopwatch progressStopwatch = Stopwatch.StartNew();
            long lastProgressReportAt = 0;
            try
            {
                foreach (string file in files.Where(static path => !string.IsNullOrWhiteSpace(path)))
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        wasCanceled = true;
                        break;
                    }

                    string cleanedText = cleaner.CleanFile(file);
                    if (trimBuffer == null)
                    {
                        documentBuilder.AppendTextBlock(cleanedText);
                    }
                    else
                    {
                        trimBuffer.Append(cleanedText);
                    }

                    processedCount++;
                    if (ShouldReportProgress(
                            processedCount,
                            lastReportedCount,
                            progressStopwatch.ElapsedMilliseconds,
                            lastProgressReportAt))
                    {
                        progress?.Report(processedCount);
                        lastReportedCount = processedCount;
                        lastProgressReportAt = progressStopwatch.ElapsedMilliseconds;
                    }
                }

                if (processedCount != lastReportedCount)
                {
                    progress?.Report(processedCount);
                }

                if (trimBuffer != null && !wasCanceled)
                {
                    documentBuilder.AppendTextBlock(trimBuffer.BuildTrimmedText());
                }

                string finalPath = documentBuilder.Complete();
                return new DocumentGenerationResult(
                    finalPath,
                    processedCount,
                    wasCanceled,
                    trimBuffer?.TotalLineCount ?? 0,
                    trimBuffer != null && !wasCanceled);
            }
            catch
            {
                documentBuilder.TryCompleteSilently();
                throw;
            }
        }, cancellationToken);
    }

    private static SourceTextCleaner CreateCleaner(string codeType)
    {
        return string.Equals(codeType, "Python", StringComparison.OrdinalIgnoreCase)
            ? new PythonSourceTextCleaner()
            : new CStyleSourceTextCleaner();
    }

    private static bool ShouldReportProgress(
        int processedCount,
        int lastReportedCount,
        long elapsedMilliseconds,
        long lastProgressReportAt)
    {
        if (processedCount <= 0 || processedCount == lastReportedCount)
        {
            return false;
        }

        if (processedCount == 1)
        {
            return true;
        }

        if (processedCount - lastReportedCount >= ProgressReportStride)
        {
            return true;
        }

        return elapsedMilliseconds - lastProgressReportAt >= ProgressReportIntervalMilliseconds;
    }

    private sealed class FixedLineTrimBuffer
    {
        private readonly List<string> headLines;

        private readonly Queue<string> tailLines;

        private readonly int preservedLineCount;

        public FixedLineTrimBuffer(int pageCount, int linesPerPage)
        {
            preservedLineCount = checked(pageCount * linesPerPage);
            headLines = new List<string>(preservedLineCount);
            tailLines = new Queue<string>(preservedLineCount);
        }

        public int TotalLineCount { get; private set; }

        public void Append(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            foreach (string line in EnumerateLines(text))
            {
                TotalLineCount++;
                if (headLines.Count < preservedLineCount)
                {
                    headLines.Add(line);
                }

                if (tailLines.Count == preservedLineCount)
                {
                    tailLines.Dequeue();
                }

                tailLines.Enqueue(line);
            }
        }

        public string BuildTrimmedText()
        {
            if (TotalLineCount <= preservedLineCount * 2)
            {
                throw new InvalidOperationException(
                    $"当前导出内容共 {TotalLineCount} 行，不足以执行前后各保留 {preservedLineCount} 行的裁剪。请调小页数或关闭该选项。");
            }

            StringBuilder builder = new();
            foreach (string line in headLines)
            {
                builder.Append(line);
            }

            foreach (string line in tailLines)
            {
                builder.Append(line);
            }

            return builder.ToString();
        }

        private static IEnumerable<string> EnumerateLines(string text)
        {
            int lineStart = 0;
            for (int index = 0; index < text.Length; index++)
            {
                if (text[index] != '\n')
                {
                    continue;
                }

                yield return text.Substring(lineStart, index - lineStart + 1);
                lineStart = index + 1;
            }

            if (lineStart < text.Length)
            {
                yield return text.Substring(lineStart);
            }
        }
    }
}

internal sealed record DocumentGenerationResult(
    string OutputPath,
    int ProcessedCount,
    bool WasCanceled,
    int TotalLineCount,
    bool UsedFixedLineTrim);

internal sealed record FixedLineTrimOptions(int PageCount, int LinesPerPage)
{
    public bool IsEnabled => PageCount > 0 && LinesPerPage > 0;
}
