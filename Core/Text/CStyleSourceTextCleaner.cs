using System.Text;

namespace Source2Docx.Core.Text;

internal sealed class CStyleSourceTextCleaner : SourceTextCleaner
{
	private enum CommentParsingState
	{
		Code,
		Slash,
		MultiLineComment,
		MultiLineCommentStar,
		SingleLineComment,
		SingleLineCommentBackslash,
		CharacterLiteral,
		CharacterEscape,
		StringLiteral,
		StringEscape
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
					if (current == '/')
					{
						state = CommentParsingState.Slash;
						break;
					}

					builder.Append(current);
					if (current == '\'')
					{
						state = CommentParsingState.CharacterLiteral;
					}
					else if (current == '"')
					{
						state = CommentParsingState.StringLiteral;
					}
					break;
				case CommentParsingState.Slash:
					if (current == '*')
					{
						state = CommentParsingState.MultiLineComment;
						break;
					}

					if (current == '/')
					{
						state = CommentParsingState.SingleLineComment;
						break;
					}

					builder.Append('/');
					builder.Append(current);
					state = CommentParsingState.Code;
					break;
				case CommentParsingState.MultiLineComment:
					if (current == '*')
					{
						state = CommentParsingState.MultiLineCommentStar;
						break;
					}

					if (current == '\n')
					{
						builder.AppendLine();
					}
					break;
				case CommentParsingState.MultiLineCommentStar:
					state = current == '/'
						? CommentParsingState.Code
						: current == '*'
							? CommentParsingState.MultiLineCommentStar
							: CommentParsingState.MultiLineComment;
					break;
				case CommentParsingState.SingleLineComment:
					if (current == '\\')
					{
						state = CommentParsingState.SingleLineCommentBackslash;
					}
					else if (current == '\n')
					{
						builder.AppendLine();
						state = CommentParsingState.Code;
					}
					break;
				case CommentParsingState.SingleLineCommentBackslash:
					switch (current)
					{
						case '\\':
						case '\r':
						case '\n':
							if (current == '\n')
							{
								builder.AppendLine();
							}

							break;
						default:
							state = CommentParsingState.SingleLineComment;
							break;
					}
					break;
				case CommentParsingState.CharacterLiteral:
					builder.Append(current);
					state = current == '\\'
						? CommentParsingState.CharacterEscape
						: current == '\''
							? CommentParsingState.Code
							: CommentParsingState.CharacterLiteral;
					break;
				case CommentParsingState.CharacterEscape:
					builder.Append(current);
					state = CommentParsingState.CharacterLiteral;
					break;
				case CommentParsingState.StringLiteral:
					builder.Append(current);
					state = current == '\\'
						? CommentParsingState.StringEscape
						: current == '"'
							? CommentParsingState.Code
							: CommentParsingState.StringLiteral;
					break;
				case CommentParsingState.StringEscape:
					builder.Append(current);
					state = CommentParsingState.StringLiteral;
					break;
			}
		}

		return builder.ToString();
	}
}
