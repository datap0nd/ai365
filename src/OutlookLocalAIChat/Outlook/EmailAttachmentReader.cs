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
            // A truncated data URL is corrupt base64, so an oversized
            // image is dropped rather than bounded.
            var boundedDataUrl = (imageDataUrl ?? string.Empty).Trim();
            ImageDataUrl = boundedDataUrl.Length <=
                EmailAttachmentReader.MaxImageDataUrlCharacters
                    ? boundedDataUrl
                    : string.Empty;
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
        // Images get a higher intake ceiling because oversized ones are
        // downscaled locally before being sent as vision input.
        public const int MaxImageBytesPerAttachment = 10 * 1024 * 1024;
        public const int MaxCharactersPerAttachment = 8000;
        public const int MaxTotalCharacters = 16000;
        // 1.5 MB of image bytes is ~2.1M base64 characters plus the
        // data URL prefix; the two limits must move together.
        public const int MaxImageDataUrlCharacters = 2200000;
        // Inline images at or under this size are treated as signature
        // graphics (logos, banners) and skipped; pasted screenshots and
        // photos are far larger and always kept.
        public const int SignatureImageMaxBytes = 64 * 1024;
        private const int MaxImageBytesForBase64 = 1536 * 1024;

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

        private static readonly HashSet<string> DocumentExtensions =
            new HashSet<string>(
                new[]
                {
                    ".pdf", ".pptx", ".docx", ".ppt", ".doc",
                    ".rtf"
                },
                StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> TextExtensions =
            new HashSet<string>(
                new[]
                {
                    ".txt", ".md", ".log", ".json",
                    ".xml", ".html", ".htm", ".eml"
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
                var signatureImagesSkipped = 0;
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
                        if (IsLikelySignatureImage(
                            attachment,
                            extension,
                            SafeLong(() => outlookAttachment.Size)))
                        {
                            signatureImagesSkipped++;
                            continue;
                        }

                        // Every attachment is saved and attempted;
                        // unknown extensions are identified by content
                        // and unreadable ones produce a visible note.
                        var safeExtension =
                            System.Text.RegularExpressions.Regex
                                .IsMatch(
                                    extension,
                                    "^\\.[A-Za-z0-9]{1,10}$")
                                ? extension
                                : ".bin";
                        tempPath = Path.Combine(
                            Path.GetTempPath(),
                            "MetoAI-" +
                            Guid.NewGuid().ToString("N") +
                            safeExtension);
                        outlookAttachment.SaveAsFile(tempPath);

                        var fileInfo = new FileInfo(tempPath);
                        if (!fileInfo.Exists)
                        {
                            continue;
                        }

                        var sizeLimit =
                            ExcelExtensions.Contains(extension) ||
                            TextExtensions.Contains(extension)
                                ? MaxBytesPerAttachment
                                : MaxImageBytesPerAttachment;
                        if (fileInfo.Length > sizeLimit)
                        {
                            results.Add(new EmailAttachmentContent(
                                fileName,
                                "unreadable",
                                "[Attachment: " + fileName + ", " +
                                fileInfo.Length.ToString() +
                                " bytes. Too large for MetoAI to " +
                                "read.]"));
                            totalCharacters += 80;
                            continue;
                        }

                        var extracted = ExtractContent(
                            tempPath,
                            fileName,
                            extension);
                        if (extracted == null ||
                            extracted.Text.Length == 0)
                        {
                            results.Add(new EmailAttachmentContent(
                                fileName,
                                "unreadable",
                                "[Attachment: " + fileName + ", " +
                                fileInfo.Length.ToString() +
                                " bytes. This file type could not " +
                                "be converted to text or image " +
                                "input.]"));
                            totalCharacters += 80;
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
                                boundedText,
                                extracted.ImageDataUrl));
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

                if (signatureImagesSkipped > 0)
                {
                    results.Add(new EmailAttachmentContent(
                        "signature-images",
                        "note",
                        "[" + signatureImagesSkipped.ToString() +
                        " small inline image" +
                        (signatureImagesSkipped == 1 ? "" : "s") +
                        " ignored as signature graphics.]"));
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

        internal static bool IsLikelySignatureImage(
            object attachment,
            string extension,
            long sizeBytes)
        {
            if (!ImageExtensions.Contains(extension) ||
                sizeBytes <= 0 ||
                sizeBytes > SignatureImageMaxBytes)
            {
                return false;
            }

            return IsInlineAttachment(attachment);
        }

        private static bool IsInlineAttachment(object attachment)
        {
            object accessor = null;
            try
            {
                dynamic outlookAttachment = attachment;
                accessor = outlookAttachment.PropertyAccessor;
                if (accessor == null)
                {
                    return false;
                }

                dynamic propertyAccessor = accessor;
                try
                {
                    // PR_ATTACHMENT_HIDDEN
                    var hidden = propertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B");
                    if (hidden is bool && (bool)hidden)
                    {
                        return true;
                    }
                }
                catch
                {
                }

                try
                {
                    // PR_ATTACH_CONTENT_ID marks cid-referenced
                    // inline body images.
                    var contentId = Convert.ToString(
                        propertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x3712001F"));
                    return !string.IsNullOrEmpty(contentId);
                }
                catch
                {
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                Release(accessor);
            }
        }

        private static long SafeLong(Func<object> reader)
        {
            try
            {
                return Convert.ToInt64(reader());
            }
            catch
            {
                return 0;
            }
        }

        // Loads a user-chosen local file through the same bounded
        // extraction pipeline as email attachments (documents become
        // text, images become vision input). User-initiated only.
        public static EmailAttachmentContent LoadLocalFile(string path)
        {
            try
            {
                var fileName = Path.GetFileName(path ?? string.Empty);
                if (fileName.Length == 0)
                {
                    return null;
                }

                var info = new FileInfo(path);
                if (!info.Exists)
                {
                    return null;
                }

                var extension = Path.GetExtension(fileName);
                var sizeLimit =
                    ExcelExtensions.Contains(extension) ||
                    TextExtensions.Contains(extension)
                        ? MaxBytesPerAttachment
                        : MaxImageBytesPerAttachment;
                if (info.Length > sizeLimit)
                {
                    return new EmailAttachmentContent(
                        fileName,
                        "unreadable",
                        "[File: " + fileName + ", " +
                        info.Length.ToString() +
                        " bytes. Too large for MetoAI to read.]");
                }

                var extracted = ExtractContent(
                    path,
                    fileName,
                    extension);
                return extracted ?? new EmailAttachmentContent(
                    fileName,
                    "unreadable",
                    "[File: " + fileName +
                    ". This file type could not be converted to " +
                    "text or image input.]");
            }
            catch
            {
                return null;
            }
        }

        // Small JPEG preview for attachment tray thumbnails.
        public static string BuildThumbnailDataUrl(string path)
        {
            try
            {
                using (var original =
                    System.Drawing.Image.FromFile(path))
                {
                    var longSide = Math.Max(
                        original.Width,
                        original.Height);
                    var scale = longSide > 96
                        ? 96.0 / longSide
                        : 1.0;
                    var width = Math.Max(
                        1,
                        (int)Math.Round(original.Width * scale));
                    var height = Math.Max(
                        1,
                        (int)Math.Round(original.Height * scale));
                    using (var bitmap =
                        new System.Drawing.Bitmap(width, height))
                    {
                        using (var graphics =
                            System.Drawing.Graphics.FromImage(
                                bitmap))
                        {
                            graphics.Clear(
                                System.Drawing.Color.White);
                            graphics.DrawImage(
                                original,
                                0,
                                0,
                                width,
                                height);
                        }

                        var encoded = EncodeJpeg(bitmap, 70);
                        return encoded == null
                            ? null
                            : "data:image/jpeg;base64," +
                              Convert.ToBase64String(encoded);
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        public static bool IsSupportedExtension(string extension)
        {
            return ImageExtensions.Contains(extension) ||
                   ExcelExtensions.Contains(extension) ||
                   DocumentExtensions.Contains(extension) ||
                   TextExtensions.Contains(extension);
        }

        public static bool IsImageExtension(string extension)
        {
            return ImageExtensions.Contains(extension);
        }

        private static EmailAttachmentContent ExtractContent(
            string path,
            string fileName,
            string extension)
        {
            if (ImageExtensions.Contains(extension))
            {
                return ExtractImage(
                    path,
                    fileName,
                    ImageMimeType(extension));
            }

            if (ExcelExtensions.Contains(extension))
            {
                return ExtractSpreadsheet(path, fileName, extension);
            }

            if (DocumentExtensions.Contains(extension))
            {
                return ExtractDocument(path, fileName, extension);
            }

            if (TextExtensions.Contains(extension))
            {
                return new EmailAttachmentContent(
                    fileName,
                    "text",
                    ReadTextFile(path));
            }

            return SniffUnknownContent(path, fileName);
        }

        // Unknown or missing extensions are identified by content so
        // every attachment is at least attempted: image magic bytes,
        // OOXML zip parts, OLE compound streams, then plain text.
        private static EmailAttachmentContent SniffUnknownContent(
            string path,
            string fileName)
        {
            var sniffedMimeType = SniffImageMimeType(path);
            if (sniffedMimeType != null)
            {
                return ExtractImage(path, fileName, sniffedMimeType);
            }

            var header = ReadHeader(path, 4096);
            if (header.Length > 3 &&
                header[0] == (byte)'P' &&
                header[1] == (byte)'K')
            {
                if (ZipContainsEntry(path, "word/document.xml"))
                {
                    return ExtractDocument(path, fileName, ".docx");
                }

                if (ZipContainsEntry(path, "ppt/presentation.xml"))
                {
                    return ExtractDocument(path, fileName, ".pptx");
                }

                if (ZipContainsEntry(path, "xl/workbook.xml"))
                {
                    return ExtractSpreadsheet(
                        path,
                        fileName,
                        ".xlsx");
                }

                return null;
            }

            if (LegacyOfficeTextExtractor.LooksLikeCompoundFile(
                header))
            {
                var bytes = File.ReadAllBytes(path);
                if (LegacyOfficeTextExtractor.CompoundStreamExists(
                    bytes,
                    "WordDocument"))
                {
                    return ExtractDocument(path, fileName, ".doc");
                }

                if (LegacyOfficeTextExtractor.CompoundStreamExists(
                    bytes,
                    "PowerPoint Document"))
                {
                    return ExtractDocument(path, fileName, ".ppt");
                }

                if (LegacyOfficeTextExtractor.CompoundStreamExists(
                        bytes,
                        "Workbook") ||
                    LegacyOfficeTextExtractor.CompoundStreamExists(
                        bytes,
                        "Book"))
                {
                    return ExtractSpreadsheet(
                        path,
                        fileName,
                        ".xls");
                }

                return null;
            }

            if (header.Length > 4 &&
                header[0] == (byte)'%' &&
                header[1] == (byte)'P' &&
                header[2] == (byte)'D' &&
                header[3] == (byte)'F')
            {
                return ExtractDocument(path, fileName, ".pdf");
            }

            if (LooksLikeText(header))
            {
                return new EmailAttachmentContent(
                    fileName,
                    "text",
                    ReadTextFile(path));
            }

            return null;
        }

        private static byte[] ReadHeader(string path, int count)
        {
            using (var stream = File.OpenRead(path))
            {
                var buffer = new byte[count];
                var read = stream.Read(buffer, 0, buffer.Length);
                if (read == buffer.Length)
                {
                    return buffer;
                }

                var bounded = new byte[Math.Max(0, read)];
                Array.Copy(buffer, bounded, bounded.Length);
                return bounded;
            }
        }

        private static bool ZipContainsEntry(
            string path,
            string entryName)
        {
            try
            {
                using (var zip = ZipFile.OpenRead(path))
                {
                    return zip.GetEntry(entryName) != null;
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool LooksLikeText(byte[] header)
        {
            if (header.Length == 0)
            {
                return false;
            }

            var printable = 0;
            foreach (var value in header)
            {
                if (value == 0)
                {
                    return false;
                }

                if (value == 9 ||
                    value == 10 ||
                    value == 13 ||
                    (value >= 32 && value < 127) ||
                    value >= 160)
                {
                    printable++;
                }
            }

            return printable * 100 >= header.Length * 90;
        }

        private static string WithLegacyHeader(
            string fileName,
            string format,
            string extracted)
        {
            return extracted.Trim().Length > 0
                ? "[" + format + " attachment: " + fileName +
                  " - legacy format, best-effort text extraction]\n" +
                  extracted
                : string.Empty;
        }

        private static EmailAttachmentContent ExtractDocument(
            string path,
            string fileName,
            string extension)
        {
            var kind = "document";
            string text;
            try
            {
                switch (extension.ToLowerInvariant())
                {
                    case ".pdf":
                        kind = "pdf";
                        text = ExtractPdfText(path);
                        if (CountReadableCharacters(text) < 40)
                        {
                            text =
                                "[PDF attachment: " + fileName +
                                ". No machine-readable text could be " +
                                "extracted. The PDF is likely scanned " +
                                "pages or uses embedded font encodings. " +
                                "Ask the user to export it as text or " +
                                "paste the content into the email.]";
                        }
                        else
                        {
                            text =
                                "[PDF attachment: " + fileName +
                                " - best-effort text extraction; layout " +
                                "and some characters may be lost]\n" +
                                text;
                        }

                        break;
                    case ".pptx":
                        kind = "powerpoint";
                        text = ExtractPptxText(path);
                        break;
                    case ".docx":
                        kind = "word";
                        text = ExtractDocxText(path);
                        break;
                    case ".ppt":
                        kind = "powerpoint";
                        text = WithLegacyHeader(
                            fileName,
                            "PowerPoint",
                            LegacyOfficeTextExtractor
                                .ExtractPptText(
                                    File.ReadAllBytes(path)));
                        break;
                    case ".doc":
                        kind = "word";
                        text = WithLegacyHeader(
                            fileName,
                            "Word",
                            LegacyOfficeTextExtractor
                                .ExtractDocText(
                                    File.ReadAllBytes(path)));
                        break;
                    case ".rtf":
                        kind = "word";
                        text = LegacyOfficeTextExtractor
                            .ExtractRtfText(
                                File.ReadAllBytes(path));
                        break;
                    default:
                        text = string.Empty;
                        break;
                }
            }
            catch
            {
                text =
                    "[Attachment: " + fileName +
                    ". The file could not be parsed for text.]";
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                text =
                    "[Attachment: " + fileName +
                    ". No readable text was extracted.]";
            }

            return new EmailAttachmentContent(
                fileName,
                kind,
                text);
        }

        private static string ExtractPptxText(string path)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                XNamespace drawingNamespace =
                    "http://schemas.openxmlformats.org/drawingml/2006/main";
                var slides = zip.Entries
                    .Where(entry =>
                        entry.FullName.StartsWith(
                            "ppt/slides/slide",
                            StringComparison.OrdinalIgnoreCase) &&
                        entry.FullName.EndsWith(
                            ".xml",
                            StringComparison.OrdinalIgnoreCase))
                    .OrderBy(entry => SlideNumber(entry.FullName))
                    .ToList();
                var builder = new StringBuilder();
                foreach (var slide in slides)
                {
                    XDocument document;
                    using (var stream = slide.Open())
                    {
                        document = XDocument.Load(stream);
                    }

                    builder.AppendLine(
                        "[Slide " +
                        SlideNumber(slide.FullName).ToString() +
                        "]");
                    foreach (var paragraph in document.Descendants(
                        drawingNamespace + "p"))
                    {
                        var text = string.Concat(
                            paragraph
                                .Descendants(drawingNamespace + "t")
                                .Select(node => node.Value));
                        if (text.Trim().Length > 0)
                        {
                            builder.AppendLine(text);
                        }
                    }

                    if (builder.Length >= MaxCharactersPerAttachment)
                    {
                        break;
                    }
                }

                return builder.ToString();
            }
        }

        private static int SlideNumber(string entryName)
        {
            var digits = new StringBuilder();
            foreach (var character in entryName)
            {
                if (char.IsDigit(character))
                {
                    digits.Append(character);
                }
                else if (digits.Length > 0)
                {
                    break;
                }
            }

            int number;
            return int.TryParse(digits.ToString(), out number)
                ? number
                : 0;
        }

        private static string ExtractDocxText(string path)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                var entry = zip.GetEntry("word/document.xml");
                if (entry == null)
                {
                    return string.Empty;
                }

                XDocument document;
                using (var stream = entry.Open())
                {
                    document = XDocument.Load(stream);
                }

                XNamespace wordNamespace =
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var builder = new StringBuilder();
                foreach (var paragraph in document.Descendants(
                    wordNamespace + "p"))
                {
                    var text = string.Concat(
                        paragraph
                            .Descendants(wordNamespace + "t")
                            .Select(node => node.Value));
                    if (text.Trim().Length > 0)
                    {
                        builder.AppendLine(text);
                    }

                    if (builder.Length >= MaxCharactersPerAttachment)
                    {
                        break;
                    }
                }

                return builder.ToString();
            }
        }

        private static string ExtractPdfText(string path)
        {
            return PdfTextExtractor.Extract(
                File.ReadAllBytes(path),
                MaxCharactersPerAttachment);
        }

        private static int CountReadableCharacters(string text)
        {
            var count = 0;
            foreach (var character in text ?? string.Empty)
            {
                if (char.IsLetterOrDigit(character))
                {
                    count++;
                }
            }

            return count;
        }

        private static string SniffImageMimeType(string path)
        {
            byte[] header;
            using (var stream = File.OpenRead(path))
            {
                header = new byte[12];
                var read = stream.Read(header, 0, header.Length);
                if (read < 4)
                {
                    return null;
                }
            }

            if (header[0] == 0x89 &&
                header[1] == 0x50 &&
                header[2] == 0x4E &&
                header[3] == 0x47)
            {
                return "image/png";
            }

            if (header[0] == 0xFF &&
                header[1] == 0xD8 &&
                header[2] == 0xFF)
            {
                return "image/jpeg";
            }

            if (header[0] == 'G' &&
                header[1] == 'I' &&
                header[2] == 'F' &&
                header[3] == '8')
            {
                return "image/gif";
            }

            if (header[0] == 'B' && header[1] == 'M')
            {
                return "image/bmp";
            }

            if (header[0] == 'R' &&
                header[1] == 'I' &&
                header[2] == 'F' &&
                header[3] == 'F' &&
                header[8] == 'W' &&
                header[9] == 'E' &&
                header[10] == 'B' &&
                header[11] == 'P')
            {
                return "image/webp";
            }

            if ((header[0] == 'I' &&
                 header[1] == 'I' &&
                 header[2] == 0x2A &&
                 header[3] == 0x00) ||
                (header[0] == 'M' &&
                 header[1] == 'M' &&
                 header[2] == 0x00 &&
                 header[3] == 0x2A))
            {
                return "image/tiff";
            }

            return null;
        }

        private static EmailAttachmentContent ExtractImage(
            string path,
            string fileName,
            string mimeType)
        {
            var bytes = File.ReadAllBytes(path);
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
                dataUrl =
                    "data:" +
                    mimeType +
                    ";base64," +
                    Convert.ToBase64String(bytes);
            }
            else
            {
                var downscaled = TryDownscaleToJpeg(path);
                if (downscaled != null)
                {
                    builder.Append(
                        "\nThe image was downscaled locally to fit " +
                        "the vision size limit.");
                    dataUrl =
                        "data:image/jpeg;base64," +
                        Convert.ToBase64String(downscaled);
                }
            }

            if (dataUrl != null)
            {
                builder.Append(
                    "\nVision-capable models receive this image " +
                    "through multimodal input after tool results.");
            }
            else
            {
                builder.Append(
                    "\nImage exceeds the vision size limit and could " +
                    "not be downscaled. Only metadata is included.");
            }

            return new EmailAttachmentContent(
                fileName,
                "image",
                builder.ToString(),
                dataUrl);
        }

        private static byte[] TryDownscaleToJpeg(string path)
        {
            try
            {
                using (var original =
                    System.Drawing.Image.FromFile(path))
                {
                    var longSide = Math.Max(
                        original.Width,
                        original.Height);
                    var targetSide = Math.Min(longSide, 2048);
                    while (targetSide >= 256)
                    {
                        var scale = (double)targetSide / longSide;
                        var width = Math.Max(
                            1,
                            (int)Math.Round(original.Width * scale));
                        var height = Math.Max(
                            1,
                            (int)Math.Round(original.Height * scale));
                        using (var bitmap =
                            new System.Drawing.Bitmap(width, height))
                        {
                            using (var graphics =
                                System.Drawing.Graphics.FromImage(
                                    bitmap))
                            {
                                graphics.Clear(
                                    System.Drawing.Color.White);
                                graphics.InterpolationMode =
                                    System.Drawing.Drawing2D
                                        .InterpolationMode
                                        .HighQualityBicubic;
                                graphics.DrawImage(
                                    original,
                                    0,
                                    0,
                                    width,
                                    height);
                            }

                            foreach (var quality in
                                new long[] { 80, 55 })
                            {
                                var encoded = EncodeJpeg(
                                    bitmap,
                                    quality);
                                if (encoded != null &&
                                    encoded.Length <=
                                    MaxImageBytesForBase64)
                                {
                                    return encoded;
                                }
                            }
                        }

                        targetSide /= 2;
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static byte[] EncodeJpeg(
            System.Drawing.Bitmap bitmap,
            long quality)
        {
            var encoder = System.Drawing.Imaging.ImageCodecInfo
                .GetImageEncoders()
                .FirstOrDefault(codec =>
                    codec.FormatID ==
                    System.Drawing.Imaging.ImageFormat.Jpeg.Guid);
            if (encoder == null)
            {
                return null;
            }

            using (var parameters =
                new System.Drawing.Imaging.EncoderParameters(1))
            {
                parameters.Param[0] =
                    new System.Drawing.Imaging.EncoderParameter(
                        System.Drawing.Imaging.Encoder.Quality,
                        quality);
                using (var stream = new MemoryStream())
                {
                    bitmap.Save(stream, encoder, parameters);
                    return stream.ToArray();
                }
            }
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
                var extracted = LegacyOfficeTextExtractor
                    .ExtractXlsText(File.ReadAllBytes(path));
                text = extracted.Trim().Length > 0
                    ? "[Excel attachment: " + fileName +
                      " - legacy .xls cell text without positions]\n" +
                      extracted
                    : "[Excel attachment: " + fileName +
                      ". No readable cell text was extracted from " +
                      "the legacy workbook. Save as .xlsx or .csv " +
                      "for full extraction.]";
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
