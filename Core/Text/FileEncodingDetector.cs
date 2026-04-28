using System.IO;
using System.Text;

namespace Source2Docx.Core.Text;

internal static class FileEncodingDetector
{
	public static Encoding Detect(string filePath)
	{
		using FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
		using BinaryReader binaryReader = new BinaryReader(fileStream, Encoding.Default);
		byte[] data = binaryReader.ReadBytes((int)fileStream.Length);
		if (LooksLikeUtf8(data) || HasUtf8Bom(data))
		{
			return Encoding.UTF8;
		}

		if (data.Length > 3 && data[0] == 254 && data[1] == byte.MaxValue && data[2] == 0)
		{
			return Encoding.BigEndianUnicode;
		}

		if (data.Length > 3 && data[0] == byte.MaxValue && data[1] == 254 && data[2] == 65)
		{
			return Encoding.Unicode;
		}

		return Encoding.Default;
	}

	private static bool HasUtf8Bom(byte[] data)
	{
		return data.Length > 3 && data[0] == 239 && data[1] == 187 && data[2] == 191;
	}

	private static bool LooksLikeUtf8(byte[] data)
	{
		int remainingUtf8Bytes = 1;
		for (int index = 0; index < data.Length; index++)
		{
			byte current = data[index];
			if (remainingUtf8Bytes == 1)
			{
				if (current < 128)
				{
					continue;
				}

				while (((current <<= 1) & 0x80) != 0)
				{
					remainingUtf8Bytes++;
				}

				if (remainingUtf8Bytes == 1 || remainingUtf8Bytes > 6)
				{
					return false;
				}
			}
			else
			{
				if ((current & 0xC0) != 128)
				{
					return false;
				}

				remainingUtf8Bytes--;
			}
		}

		return remainingUtf8Bytes <= 1;
	}
}
