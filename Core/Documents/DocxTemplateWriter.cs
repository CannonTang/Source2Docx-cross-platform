using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace Source2Docx.Core.Documents;

internal sealed class DocxTemplateWriter
{
	private static readonly XNamespace WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

	private static readonly XNamespace OfficeDocumentRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

	private static readonly XNamespace PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

	private readonly StringBuilder contentBuffer = new StringBuilder();

	private readonly string headerTitle;

	public DocxTemplateWriter(string outputPath, string title)
	{
		OutputPath = CreateDocumentFromTemplate(outputPath);
		headerTitle = (title ?? string.Empty).PadRight(60);
	}

	public string OutputPath { get; }

	public void AppendTextBlock(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return;
		}

		contentBuffer.Append(NormalizeLineEndings(text));
	}

	public void Complete()
	{
		using ZipArchive archive = ZipFile.Open(OutputPath, ZipArchiveMode.Update);
		XDocument documentXml = LoadXmlDocument(archive, "word/document.xml");
		XDocument relationshipXml = LoadXmlDocument(archive, "word/_rels/document.xml.rels");
		UpdateDocumentBody(documentXml, contentBuffer.ToString());
		string defaultHeaderPath = ResolveDefaultHeaderPath(documentXml, relationshipXml);
		if (!string.IsNullOrWhiteSpace(defaultHeaderPath))
		{
			XDocument headerXml = LoadXmlDocument(archive, defaultHeaderPath);
			UpdateHeaderTitle(headerXml, headerTitle);
			SaveXmlDocument(archive, defaultHeaderPath, headerXml);
		}

		SaveXmlDocument(archive, "word/document.xml", documentXml);
	}

	private static XDocument LoadXmlDocument(ZipArchive archive, string entryPath)
	{
		ZipArchiveEntry entry = archive.GetEntry(entryPath)
			?? throw new InvalidOperationException("未找到 DOCX 内容项: " + entryPath);
		using Stream stream = entry.Open();
		return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
	}

	private static void SaveXmlDocument(ZipArchive archive, string entryPath, XDocument document)
	{
		ZipArchiveEntry existingEntry = archive.GetEntry(entryPath)
			?? throw new InvalidOperationException("未找到 DOCX 内容项: " + entryPath);
		existingEntry.Delete();
		ZipArchiveEntry newEntry = archive.CreateEntry(entryPath, CompressionLevel.Optimal);
		using Stream stream = newEntry.Open();
		using StreamWriter writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
		document.Save(writer, SaveOptions.DisableFormatting);
	}

	private static void UpdateDocumentBody(XDocument documentXml, string content)
	{
		XElement body = documentXml.Root?.Element(WordNamespace + "body")
			?? throw new InvalidOperationException("DOCX 模板缺少正文 body。");
		XElement paragraph = body.Element(WordNamespace + "p")
			?? throw new InvalidOperationException("DOCX 模板缺少正文首段。");
		foreach (XElement run in CreateRuns(content, includeEmptyRunProperties: true))
		{
			paragraph.Add(run);
		}
	}

	private static string ResolveDefaultHeaderPath(XDocument documentXml, XDocument relationshipXml)
	{
		XElement body = documentXml.Root?.Element(WordNamespace + "body");
		XElement sectionProperties = body?.Element(WordNamespace + "sectPr");
		XElement defaultHeaderReference = sectionProperties?
			.Elements(WordNamespace + "headerReference")
			.FirstOrDefault(element => string.Equals((string)element.Attribute(WordNamespace + "type"), "default", StringComparison.Ordinal));
		string relationshipId = (string)defaultHeaderReference?.Attribute(OfficeDocumentRelationshipNamespace + "id");
		if (string.IsNullOrWhiteSpace(relationshipId))
		{
			return null;
		}

		XElement relationship = relationshipXml.Root?
			.Elements(PackageRelationshipNamespace + "Relationship")
			.FirstOrDefault(element => string.Equals((string)element.Attribute("Id"), relationshipId, StringComparison.Ordinal));
		string target = (string)relationship?.Attribute("Target");
		if (string.IsNullOrWhiteSpace(target))
		{
			return null;
		}

		return target.StartsWith("word/", StringComparison.OrdinalIgnoreCase) ? target : "word/" + target;
	}

	private static void UpdateHeaderTitle(XDocument headerXml, string title)
	{
		if (string.IsNullOrEmpty(title))
		{
			return;
		}

		XElement paragraph = headerXml.Root?.Element(WordNamespace + "p");
		if (paragraph == null)
		{
			return;
		}

		XElement paragraphProperties = paragraph.Element(WordNamespace + "pPr");
		XElement firstRunProperties = paragraph
			.Elements(WordNamespace + "r")
			.Select(element => element.Element(WordNamespace + "rPr"))
			.FirstOrDefault(element => element != null);
		XElement titleRun = CreateTextRun(title, firstRunProperties, includeEmptyRunProperties: firstRunProperties == null);
		if (paragraphProperties != null)
		{
			paragraphProperties.AddAfterSelf(titleRun);
			return;
		}

		paragraph.AddFirst(titleRun);
	}

	private static XElement CreateTextRun(string text, XElement runProperties, bool includeEmptyRunProperties)
	{
		XElement textElement = new XElement(WordNamespace + "t", text);
		PreserveSpace(textElement);
		XElement run = new XElement(WordNamespace + "r");
		if (runProperties != null)
		{
			run.Add(new XElement(runProperties));
		}
		else if (includeEmptyRunProperties)
		{
			run.Add(new XElement(WordNamespace + "rPr"));
		}

		run.Add(textElement);
		return run;
	}

	private static XElement[] CreateRuns(string text, bool includeEmptyRunProperties)
	{
		if (string.IsNullOrEmpty(text))
		{
			return Array.Empty<XElement>();
		}

		List<XElement> runs = new List<XElement>();
		StringBuilder currentText = new StringBuilder();
		for (int index = 0; index < text.Length; index++)
		{
			char current = text[index];
			switch (current)
			{
				case '\t':
					FlushTextRun(runs, currentText, includeEmptyRunProperties);
					runs.Add(CreateSimpleRun(new XElement(WordNamespace + "tab"), includeEmptyRunProperties));
					break;
				case '\n':
					FlushTextRun(runs, currentText, includeEmptyRunProperties);
					runs.Add(CreateSimpleRun(new XElement(WordNamespace + "br"), includeEmptyRunProperties));
					break;
				default:
					currentText.Append(current);
					break;
			}
		}

		FlushTextRun(runs, currentText, includeEmptyRunProperties);
		return runs.ToArray();
	}

	private static void FlushTextRun(List<XElement> runs, StringBuilder currentText, bool includeEmptyRunProperties)
	{
		if (currentText.Length == 0)
		{
			return;
		}

		runs.Add(CreateTextRun(currentText.ToString(), runProperties: null, includeEmptyRunProperties));
		currentText.Clear();
	}

	private static XElement CreateSimpleRun(XElement content, bool includeEmptyRunProperties)
	{
		XElement run = new XElement(WordNamespace + "r");
		if (includeEmptyRunProperties)
		{
			run.Add(new XElement(WordNamespace + "rPr"));
		}

		run.Add(content);
		return run;
	}

	private static void PreserveSpace(XElement textElement)
	{
		string value = textElement.Value;
		if (value.StartsWith(" ", StringComparison.Ordinal) || value.EndsWith(" ", StringComparison.Ordinal))
		{
			textElement.SetAttributeValue(XNamespace.Xml + "space", "preserve");
		}
	}

	private static string NormalizeLineEndings(string text)
	{
		return text
			.Replace("\r\n", "\n", StringComparison.Ordinal)
			.Replace('\r', '\n');
	}

	private static string CreateDocumentFromTemplate(string outputPath)
	{
		Assembly assembly = typeof(DocxTemplateWriter).Assembly;
		using Stream manifestResourceStream = assembly.GetManifestResourceStream("Source2Doc.a.bin")
			?? throw new FileNotFoundException("未找到嵌入的文档模板资源 Source2Doc.a.bin。");
		using MemoryStream templateBuffer = new MemoryStream();
		manifestResourceStream.CopyTo(templateBuffer);
		byte[] templateBytes = templateBuffer.ToArray();

		string normalizedOutputPath = Path.IsPathRooted(outputPath) ? outputPath : Path.Combine(Environment.CurrentDirectory, outputPath);
		string directoryPath = Path.GetDirectoryName(normalizedOutputPath);
		if (string.IsNullOrWhiteSpace(directoryPath))
		{
			directoryPath = Environment.CurrentDirectory;
			normalizedOutputPath = Path.Combine(directoryPath, Path.GetFileName(normalizedOutputPath));
		}

		Directory.CreateDirectory(directoryPath);
		try
		{
			WriteTemplateFile(normalizedOutputPath, templateBytes);
			return normalizedOutputPath;
		}
		catch
		{
			string extension = Path.GetExtension(normalizedOutputPath);
			if (string.IsNullOrWhiteSpace(extension))
			{
				extension = ".docx";
			}

			string fallbackName = Path.GetFileNameWithoutExtension(normalizedOutputPath) + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + extension;
			string fallbackPath = Path.Combine(directoryPath, fallbackName);
			WriteTemplateFile(fallbackPath, templateBytes);
			return fallbackPath;
		}
	}

	private static void WriteTemplateFile(string outputPath, byte[] templateBytes)
	{
		using FileStream fileStream = new FileStream(outputPath, FileMode.Create);
		using BinaryWriter binaryWriter = new BinaryWriter(fileStream, Encoding.UTF8);
		binaryWriter.Write(templateBytes);
		binaryWriter.Flush();
	}
}
