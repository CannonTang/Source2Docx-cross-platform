using System;
using System.IO;
using System.Text;

namespace Source2Docx.Core.Text;

internal abstract class SourceTextCleaner
{
	public string CleanFile(string filePath)
	{
		Encoding encoding = FileEncodingDetector.Detect(filePath);
		string content;
		using (StreamReader reader = new StreamReader(filePath, encoding, detectEncodingFromByteOrderMarks: true))
		{
			content = reader.ReadToEnd();
		}

		string cleaned = IsMarkupFile(filePath) ? CleanMarkup(content) : CleanCode(content);
		if (cleaned.Length == 0)
		{
			return string.Empty;
		}

		return cleaned.EndsWith(Environment.NewLine, StringComparison.Ordinal)
			? cleaned
			: cleaned + Environment.NewLine;
	}

	protected abstract string RemoveCodeComments(string content);

	private static bool IsMarkupFile(string filePath)
	{
		string extension = Path.GetExtension(filePath);
		return string.Equals(extension, ".xml", StringComparison.OrdinalIgnoreCase)
			|| string.Equals(extension, ".xaml", StringComparison.OrdinalIgnoreCase);
	}

	private string CleanMarkup(string content)
	{
		return RemoveBlankLines(RemoveMarkupComments(content));
	}

	private string CleanCode(string content)
	{
		return RemoveBlankLines(RemoveCodeComments(content));
	}

	private static string RemoveMarkupComments(string content)
	{
		StringBuilder builder = new StringBuilder(content.Length);
		MarkupParsingState state = MarkupParsingState.Code;
		string remainingText = content;
		while (remainingText.Length > 2)
		{
			switch (state)
			{
				case MarkupParsingState.Code:
					int commentStart = remainingText.IndexOf("<!--", StringComparison.Ordinal);
					if (commentStart >= 0)
					{
						builder.Append(remainingText, 0, commentStart);
						remainingText = remainingText.Substring(commentStart + 4);
						state = MarkupParsingState.Comment;
					}
					else
					{
						builder.Append(remainingText);
						remainingText = string.Empty;
					}
					break;
				case MarkupParsingState.Comment:
					int commentEnd = remainingText.IndexOf("-->", StringComparison.Ordinal);
					if (commentEnd >= 0)
					{
						remainingText = remainingText.Substring(commentEnd + 3);
						state = MarkupParsingState.Code;
					}
					else
					{
						remainingText = string.Empty;
					}
					break;
			}
		}

		if (remainingText.Length > 0 && state == MarkupParsingState.Code)
		{
			builder.Append(remainingText);
		}

		return builder.ToString();
	}

	private static string RemoveBlankLines(string content)
	{
		StringBuilder builder = new StringBuilder(content.Length);
		using StringReader reader = new StringReader(content);
		string line;
		while ((line = reader.ReadLine()) != null)
		{
			string trimmed = line.Replace("\t", " ", StringComparison.Ordinal).Trim();
			if (trimmed.Length > 0)
			{
				builder.AppendLine(line);
			}
		}

		return builder.ToString();
	}

	private enum MarkupParsingState
	{
		Code,
		Comment
	}
}
