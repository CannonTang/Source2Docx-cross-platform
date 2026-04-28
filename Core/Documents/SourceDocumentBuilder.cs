using System;
using Source2Docx.Core.Text;

namespace Source2Docx.Core.Documents;

internal sealed class SourceDocumentBuilder
{
	private readonly SourceTextCleaner cleaner;

	private readonly DocxTemplateWriter writer;

	public SourceDocumentBuilder(string outputPath, string title, SourceTextCleaner cleaner)
	{
		this.cleaner = cleaner ?? throw new ArgumentNullException(nameof(cleaner));
		writer = new DocxTemplateWriter(outputPath, title);
	}

	public void AppendFile(string filePath)
	{
		string cleanedText = cleaner.CleanFile(filePath);
		if (string.IsNullOrEmpty(cleanedText))
		{
			return;
		}

		writer.AppendTextBlock(cleanedText);
	}

	public void AppendTextBlock(string text)
	{
		writer.AppendTextBlock(text);
	}

	public string Complete()
	{
		writer.Complete();
		return writer.OutputPath;
	}

	public void TryCompleteSilently()
	{
		try
		{
			writer.Complete();
		}
		catch
		{
		}
	}
}
