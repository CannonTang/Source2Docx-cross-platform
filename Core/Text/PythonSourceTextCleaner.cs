using System.Text;

namespace Source2Docx.Core.Text;

internal sealed class PythonSourceTextCleaner : SourceTextCleaner
{
	private enum CommentParsingState
	{
		Code,
		MultilineComment,
		SingleLineComment,
		SingleQuote,
		DoubleSingleQuote,
		SingleDoubleQuote,
		DoubleDoubleQuote,
		CommentSingleQuote,
		CommentDoubleSingleQuote,
		CommentSingleDoubleQuote,
		CommentDoubleDoubleQuote,
		LineStart,
		Backslash
	}

	protected override string RemoveCodeComments(string content)
	{
		StringBuilder builder = new StringBuilder(content.Length);
		CommentParsingState state = CommentParsingState.Code;
		for (int index = 0; index < content.Length; index++)
		{
			char current = content[index];
			switch (state)
			{
				case CommentParsingState.Code:
					switch (current)
					{
						case '\\':
							state = CommentParsingState.Backslash;
							break;
						case '#':
							state = CommentParsingState.SingleLineComment;
							break;
						case '\n':
							state = CommentParsingState.LineStart;
							builder.Append(current);
							break;
						default:
							builder.Append(current);
							break;
					}

					break;
				case CommentParsingState.LineStart:
					switch (current)
					{
						case ' ':
							builder.Append(current);
							break;
						case '\'':
							builder.AppendLine();
							state = CommentParsingState.SingleQuote;
							break;
						case '"':
							builder.AppendLine();
							state = CommentParsingState.SingleDoubleQuote;
							break;
						case '#':
							state = CommentParsingState.SingleLineComment;
							break;
						default:
							if (current != '\n' && current != '\r')
							{
								state = CommentParsingState.Code;
								builder.Append(current);
							}

							break;
					}

					break;
				case CommentParsingState.Backslash:
					builder.Append('\\');
					builder.Append(current);
					state = CommentParsingState.Code;
					break;
				case CommentParsingState.SingleLineComment:
					if (current == '\n')
					{
						builder.AppendLine();
						state = CommentParsingState.LineStart;
					}

					break;
				case CommentParsingState.SingleQuote:
					if (current == '\'')
					{
						state = CommentParsingState.DoubleSingleQuote;
						break;
					}

					builder.Append('\'');
					state = CommentParsingState.Code;
					break;
				case CommentParsingState.DoubleSingleQuote:
					if (current == '\'')
					{
						state = CommentParsingState.MultilineComment;
						break;
					}

					builder.Append('\'');
					builder.Append('\'');
					state = CommentParsingState.Code;
					break;
				case CommentParsingState.CommentSingleQuote:
					state = current == '\'' ? CommentParsingState.CommentDoubleSingleQuote : CommentParsingState.MultilineComment;
					break;
				case CommentParsingState.CommentDoubleSingleQuote:
					state = current == '\'' ? CommentParsingState.Code : CommentParsingState.MultilineComment;
					break;
				case CommentParsingState.SingleDoubleQuote:
					if (current == '"')
					{
						state = CommentParsingState.DoubleDoubleQuote;
						break;
					}

					builder.Append('"');
					state = CommentParsingState.Code;
					break;
				case CommentParsingState.DoubleDoubleQuote:
					if (current == '"')
					{
						state = CommentParsingState.MultilineComment;
						break;
					}

					builder.Append('"');
					builder.Append('"');
					state = CommentParsingState.Code;
					break;
				case CommentParsingState.CommentSingleDoubleQuote:
					state = current == '"' ? CommentParsingState.CommentDoubleDoubleQuote : CommentParsingState.MultilineComment;
					break;
				case CommentParsingState.CommentDoubleDoubleQuote:
					state = current == '"' ? CommentParsingState.Code : CommentParsingState.MultilineComment;
					break;
				case CommentParsingState.MultilineComment:
					if (current == '\'')
					{
						state = CommentParsingState.CommentSingleQuote;
					}
					else if (current == '"')
					{
						state = CommentParsingState.CommentSingleDoubleQuote;
					}

					break;
			}
		}

		return builder.ToString();
	}
}
