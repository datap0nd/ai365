using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using OutlookLocalAIChat.Security;

namespace OutlookLocalAIChat.Outlook
{
    public sealed class EmailAttachmentContent
    {
        public EmailAttachmentContent(
            string fileName,
            string kind,
            string text,
            string imageDataUrl = null)
        {
            FileName = TextBoundary.SingleLine(
                fileName ?? string.Empty,
                180);
            Kind = TextBoundary.SingleLine(
                kind ?? string.Empty,
                32);
            Text = TextBoundary.PlainText(
                text ?? string.Empty,
                EmailAttachmentReader.MaxCharactersPerAttachment);
            ImageDataUrl = TextBoundary.SingleLine(
                imageDataUrl ?? string.Empty,
                700000);
        }

        public string FileName { get; }

        public string Kind { get; }

        public string Text { get; }

        public string ImageDataUrl { get; }
    }

    public static class EmailAttachmentReader
    {
        public const int MaxAttachments = 10;
        public const int MaxBytesPerAttachment = 2 * 1024 * 1024;
        public const int MaxCharactersPerAttachment = 8000;
        public const int MaxTotalCharacters = 16000;
        private const int MaxImageBytesForBase64 = 512 * 1024;

        private static readonly HashSet<string> ImageExtensions =
            new HashSet<string>(
                new[]
                {
                    ".png", ".jpg", ".jpeg", ".gif",
                    ".bmp", ".webp", ".tif", ".tiff"
                },
                StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> ExcelExtensions =
            new HashSet<string>(
                new[]
                {
                    ".xlsx", ".xlsm", ".xls", ".csv"
                },
                StringComparer.OrdinalIgnoreCase);

        public static IReadOnlyList<EmailAttachmentContent> Read(
            object outlookApplication,
            MessageSnapshot message)
        {
            if (outlookApplication == null)
            {
                throw new ArgumentNullException(nameof(outlookApplication));
            }

            if (message == null ||
                message.EntryId.Length == 0)
            {
                return new EmailAttachmentContent[0];
            }

            object session = null;
            object item = null;
            object attachments = null;
            try
            {
                dynamic application = outlookApplication;
                session = application.Session;
                dynamic outlookSession = session;
                try
                {
                    item = message.StoreId.Length > 0
                        ? outlookSession.GetItemFromID(
                            message.EntryId,
                            message.StoreId)
                        : outlookSession.GetItemFromID(message.EntryId);
                }
                catch
                {
                    return new EmailAttachmentContent[0];
                }

                dynamic mail = item;
                attachments = mail.Attachments;
                if (attachments == null)
                {
                    return new EmailAttachmentContent[0];
                }

                dynamic outlookAttachments = attachments;
                var count = Math.Min(
                    Convert.ToInt32(outlookAttachments.Count),
                    MaxAttachments);
                var results = new List<EmailAttachmentContent>(count);
                var totalCharacters = 0;
                for (var index = 1;
                     index <= count &&
                     totalCharacters < MaxTotalCharacters;
                     index++)
                {
                    object attachment = null;
                    string tempPath = null;
                    try
                    {
                        attachment = outlookAttachments.Item(index);
                        dynamic outlookAttachment = attachment;
                        var fileName = SafeString(
                            () => outlookAttachment.FileName);
                        if (fileName.Length == 0)
                        {
                            continue;
                        }

                        var extension = Path.GetExtension(fileName);
                        if (!IsSupportedExtension(extension))
                        {
                            continue;
                        }

                        tempPath = Path.Combine(
                            Path.GetTempPath(),
                            "MailAI-" +
                            Guid.NewGuid().ToString("N") +
                            extension);
                        outlookAttachment.SaveAsFile(tempPath);

                        var fileInfo = new FileInfo(tempPath);
                        if (!fileInfo.Exists ||
                            fileInfo.Length > MaxBytesPerAttachment)
                        {
                            continue;
                        }

                        var extracted = ExtractContent(
                            tempPath,
                            fileName,
                            extension);
                        if (extracted == null ||
                            extracted.Text.Length == 0)
                        {
                            continue;
                        }

                        var remaining = MaxTotalCharacters -
                            totalCharacters;
                        var boundedText = TextBoundary.PlainText(
                            extracted.Text,
                            Math.Min(
                                MaxCharactersPerAttachment,
                                remaining));
                        if (boundedText.Length == 0)
                        {
                            break;
                        }

                        results.Add(
                            new EmailAttachmentContent(
                                fileName,
                                extracted.Kind,
                                boundedText));
                        totalCharacters += boundedText.Length;
                    }
                    catch
                    {
                    }
                    finally
                    {
                        if (tempPath != null)
                        {
                            TryDelete(tempPath);
                        }

                        Release(attachment);
                    }
                }

                return results;
            }
            catch
            {
                return new EmailAttachmentContent[0];
            }
            finally
            {
                Release(attachments);
                Release(item);
                Release(session);
            }
        }

        public static bool IsSupportedExtension(string extension)
        {
            return ImageExtensions.Contains(extension) ||
                   ExcelExtensions.Contains(extension);
        }

        private static EmailAttachmentContent ExtractContent(
            string path,
            string fileName,
            string extension)
        {
            if (ImageExtensions.Contains(extension))
            {
                return ExtractImage(path, fileName, extension);
            }

            if (ExcelExtensions.Contains(extension))
            {
                return ExtractSpreadsheet(path, fileName, extension);
            }

            return null;
        }

        private static EmailAttachmentContent ExtractImage(
            string path,
            string fileName,
            string extension)
        {
            var bytes = File.ReadAllBytes(path);
            var mimeType = ImageMimeType(extension);
            var builder = new StringBuilder();
            builder.Append("[Image attachment: ");
            builder.Append(fileName);
            builder.Append(", ");
            builder.Append(bytes.Length.ToString());
            builder.Append(" bytes, type ");
            builder.Append(mimeType);
            builder.Append(']');

            string dataUrl = null;
            if (bytes.Length <= MaxImageBytesForBase64)
            {
                builder.Append(
                    "\nVision-capable models receive this image " +
                    "through multimodal input after tool results.");
                dataUrl =
                    "data:" +
                    mimeType +
                    ";base64," +
                    Convert.ToBase64String(bytes);
            }
            else
            {
                builder.Append(
                    "\nImage exceeds the vision size limit. " +
                    "Only metadata is included.");
            }

            return new EmailAttachmentContent(
                fileName,
                "image",
                builder.ToString(),
                dataUrl);
        }

        private static EmailAttachmentContent ExtractSpreadsheet(
            string path,
            string fileName,
            string extension)
        {
            string text;
            if (extension.Equals(
                ".csv",
                StringComparison.OrdinalIgnoreCase))
            {
                text = ReadTextFile(path);
            }
            else if (extension.Equals(
                ".xls",
                StringComparison.OrdinalIgnoreCase))
            {
                text =
                    "[Excel attachment: " + fileName +
                    ". Legacy .xls binary workbooks are listed but not parsed. " +
                    "Save as .xlsx or .csv for text extraction.]";
            }
            else
            {
                text = ExtractXlsxText(path);
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                text =
                    "[Excel attachment: " + fileName +
                    ". No readable cell text was extracted.]";
            }

            return new EmailAttachmentContent(
                fileName,
                "excel",
                text);
        }

        private static string ExtractXlsxText(string path)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                var sharedStrings = ReadSharedStrings(zip);
                var sheetEntry = zip.Entries.FirstOrDefault(
                    entry => entry.FullName.Equals(
                        "xl/worksheets/sheet1.xml",
                        StringComparison.OrdinalIgnoreCase));
                if (sheetEntry == null)
                {
                    return string.Empty;
                }

                XDocument document;
                using (var stream = sheetEntry.Open())
                {
                    document = XDocument.Load(stream);
                }

                XNamespace spreadsheetNamespace =
                    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                var builder = new StringBuilder();
                foreach (var row in document
                    .Descendants(spreadsheetNamespace + "row"))
                {
                    var values = row
                        .Elements(spreadsheetNamespace + "c")
                        .Select(cell => ReadCellValue(
                            cell,
                            sharedStrings,
                            spreadsheetNamespace))
                        .Where(value => value.Length > 0)
                        .ToArray();
                    if (values.Length == 0)
                    {
                        continue;
                    }

                    builder.AppendLine(string.Join("\t", values));
                    if (builder.Length >= MaxCharactersPerAttachment)
                    {
                        break;
                    }
                }

                return builder.ToString();
            }
        }

        private static IList<string> ReadSharedStrings(ZipArchive zip)
        {
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
            {
                return new string[0];
            }

            XDocument document;
            using (var stream = entry.Open())
            {
                document = XDocument.Load(stream);
            }

            XNamespace spreadsheetNamespace =
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            return document
                .Descendants(spreadsheetNamespace + "si")
                .Select(item => string.Concat(
                    item.Descendants(spreadsheetNamespace + "t")
                        .Select(text => text.Value)))
                .ToList();
        }

        private static string ReadCellValue(
            XElement cell,
            IList<string> sharedStrings,
            XNamespace spreadsheetNamespace)
        {
            var type = cell.Attribute("t")?.Value ?? string.Empty;
            var valueElement = cell.Element(spreadsheetNamespace + "v");
            if (valueElement == null)
            {
                var inlineText = cell
                    .Descendants(spreadsheetNamespace + "t")
                    .Select(text => text.Value);
                return string.Concat(inlineText);
            }

            var raw = valueElement.Value ?? string.Empty;
            if (type == "s")
            {
                int index;
                if (int.TryParse(raw, out index) &&
                    index >= 0 &&
                    index < sharedStrings.Count)
                {
                    return sharedStrings[index];
                }
            }

            return raw;
        }

        private static string ReadTextFile(string path)
        {
            using (var reader = new StreamReader(
                path,
                Encoding.UTF8,
                true))
            {
                var buffer = new char[MaxCharactersPerAttachment];
                var read = reader.Read(
                    buffer,
                    0,
                    buffer.Length);
                return new string(buffer, 0, read);
            }
        }

        private static string ImageMimeType(string extension)
        {
            switch (extension.ToLowerInvariant())
            {
                case ".png":
                    return "image/png";
                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";
                case ".gif":
                    return "image/gif";
                case ".bmp":
                    return "image/bmp";
                case ".webp":
                    return "image/webp";
                case ".tif":
                case ".tiff":
                    return "image/tiff";
                default:
                    return "application/octet-stream";
            }
        }

        private static string SafeString(Func<object> reader)
        {
            try
            {
                return Convert.ToString(reader()) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}
