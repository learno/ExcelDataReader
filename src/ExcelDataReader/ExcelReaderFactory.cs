using ExcelDataReader.Core.BinaryFormat;
using ExcelDataReader.Core.CompoundFormat;
using ExcelDataReader.Core.OfficeCrypto;
using ExcelDataReader.Exceptions;
using ExcelDataReader.Misc;

namespace ExcelDataReader;

/// <summary>
/// The ExcelReader Factory.
/// </summary>
public static class ExcelReaderFactory
{
    private const string DirectoryEntryWorkbook = "Workbook";
    private const string DirectoryEntryBook = "Book";
    private const string DirectoryEntryEncryptedPackage = "EncryptedPackage";
    private const string DirectoryEntryEncryptionInfo = "EncryptionInfo";

    /// <summary>
    /// Creates an instance of <see cref="ExcelBinaryReader"/> or <see cref="ExcelOpenXmlReader"/>.
    /// </summary>
    /// <param name="fileStream">The file stream.</param>
    /// <param name="configuration">The configuration object.</param>
    /// <returns>The excel data reader.</returns>
    public static IExcelDataReader CreateReader(Stream fileStream, ExcelReaderConfiguration configuration = null)
    {
        configuration ??= new ExcelReaderConfiguration();

        if (configuration.LeaveOpen)
        {
            fileStream = new LeaveOpenStream(fileStream);
        }

        var probe = new byte[8];
        fileStream.Seek(0, SeekOrigin.Begin);
        fileStream.ReadAtLeast(probe, 0, probe.Length);
        fileStream.Seek(0, SeekOrigin.Begin);

        if (CompoundDocument.IsCompoundDocument(probe))
        {
            // Can be BIFF5-8 or password protected OpenXml
            var document = new CompoundDocument(fileStream);
            if (TryGetWorkbook(fileStream, document, out var stream))
            {
                return new ExcelBinaryReader(stream, configuration.Password, configuration.FallbackEncoding);
            }

            if (TryGetEncryptedPackage(fileStream, document, configuration.Password, out stream))
            {
                return new ExcelOpenXmlReader(stream);
            }

            throw new ExcelReaderException(Errors.ErrorStreamWorkbookNotFound);
        }

        if (XlsWorkbook.IsRawBiffStream(probe))
        {
            return new ExcelBinaryReader(fileStream, configuration.Password, configuration.FallbackEncoding);
        }

        if (probe[0] == 0x50 && probe[1] == 0x4B)
        {
            // zip files start with 'PK'
            return new ExcelOpenXmlReader(fileStream);
        }

        // Check for Excel XML format (SpreadsheetML)
        if (IsExcelXmlFormat(fileStream))
        {
            return new ExcelXmlReader(fileStream);
        }

        throw new HeaderException(Errors.ErrorHeaderSignature);
    }

    /// <summary>
    /// Creates an instance of <see cref="ExcelBinaryReader"/>.
    /// </summary>
    /// <param name="fileStream">The file stream.</param>
    /// <param name="configuration">The configuration object.</param>
    /// <returns>The excel data reader.</returns>
    public static IExcelDataReader CreateBinaryReader(Stream fileStream, ExcelReaderConfiguration configuration = null)
    {
        configuration ??= new ExcelReaderConfiguration();

        if (configuration.LeaveOpen)
        {
            fileStream = new LeaveOpenStream(fileStream);
        }

        var probe = new byte[8];
        fileStream.Seek(0, SeekOrigin.Begin);
        fileStream.ReadAtLeast(probe, 0, probe.Length);
        fileStream.Seek(0, SeekOrigin.Begin);

        if (CompoundDocument.IsCompoundDocument(probe))
        {
            var document = new CompoundDocument(fileStream);
            if (TryGetWorkbook(fileStream, document, out var stream))
            {
                return new ExcelBinaryReader(stream, configuration.Password, configuration.FallbackEncoding);
            }
            else
            {
                throw new ExcelReaderException(Errors.ErrorStreamWorkbookNotFound);
            }
        }
        else if (XlsWorkbook.IsRawBiffStream(probe))
        {
            return new ExcelBinaryReader(fileStream, configuration.Password, configuration.FallbackEncoding);
        }
        else
        {
            throw new HeaderException(Errors.ErrorHeaderSignature);
        }
    }

    /// <summary>
    /// Creates an instance of <see cref="ExcelOpenXmlReader"/>.
    /// </summary>
    /// <param name="fileStream">The file stream.</param>
    /// <param name="configuration">The reader configuration -or- <see langword="null"/> to use the default configuration.</param>
    /// <returns>The excel data reader.</returns>
    public static IExcelDataReader CreateOpenXmlReader(Stream fileStream, ExcelReaderConfiguration configuration = null)
    {
        configuration ??= new ExcelReaderConfiguration();

        if (configuration.LeaveOpen)
        {
            fileStream = new LeaveOpenStream(fileStream);
        }

        var probe = new byte[8];
        fileStream.Seek(0, SeekOrigin.Begin);
        fileStream.ReadAtLeast(probe, 0, probe.Length);
        fileStream.Seek(0, SeekOrigin.Begin);

        // Probe for password protected compound document or zip file
        if (CompoundDocument.IsCompoundDocument(probe))
        {
            var document = new CompoundDocument(fileStream);
            if (TryGetEncryptedPackage(fileStream, document, configuration.Password, out var stream))
            {
                return new ExcelOpenXmlReader(stream);
            }

            throw new ExcelReaderException(Errors.ErrorCompoundNoOpenXml);
        }

        if (probe[0] == 0x50 && probe[1] == 0x4B)
        {
            // Zip files start with 'PK'
            return new ExcelOpenXmlReader(fileStream);
        }

        throw new HeaderException(Errors.ErrorHeaderSignature);
    }

    /// <summary>
    /// Creates an instance of ExcelCsvReader.
    /// </summary>
    /// <param name="fileStream">The file stream.</param>
    /// <param name="configuration">The reader configuration -or- <see langword="null"/> to use the default configuration.</param>
    /// <returns>The excel data reader.</returns>
    public static IExcelDataReader CreateCsvReader(Stream fileStream, ExcelReaderConfiguration configuration = null)
    {
        configuration ??= new ExcelReaderConfiguration();

        if (configuration.LeaveOpen)
        {
            fileStream = new LeaveOpenStream(fileStream);
        }

        return new ExcelCsvReader(fileStream, configuration.FallbackEncoding, configuration.AutodetectSeparators, configuration.AnalyzeInitialCsvRows, configuration.QuoteChar, configuration.TrimWhiteSpace);
    }

    /// <summary>
    /// Creates an instance of ExcelXmlReader.
    /// </summary>
    /// <param name="fileStream">The file stream.</param>
    /// <param name="configuration">The reader configuration -or- <see langword="null"/> to use the default configuration.</param>
    /// <returns>The excel data reader.</returns>
    public static IExcelDataReader CreateXmlReader(Stream fileStream, ExcelReaderConfiguration configuration = null)
    {
        configuration ??= new ExcelReaderConfiguration();

        if (configuration.LeaveOpen)
        {
            fileStream = new LeaveOpenStream(fileStream);
        }

        return new ExcelXmlReader(fileStream);
    }

    private static bool TryGetWorkbook(Stream fileStream, CompoundDocument document, out Stream stream)
    {
        var workbookEntry = document.FindEntry(DirectoryEntryWorkbook, DirectoryEntryBook);
        if (workbookEntry != null)
        {
            if (workbookEntry.EntryType != STGTY.STGTY_STREAM)
            {
                throw new ExcelReaderException(Errors.ErrorWorkbookIsNotStream);
            }

            stream = new CompoundStream(document, fileStream, workbookEntry.StreamFirstSector, (int)workbookEntry.StreamSize, workbookEntry.IsEntryMiniStream, false);
            return true;
        }

        stream = null;
        return false;
    }

    private static bool TryGetEncryptedPackage(Stream fileStream, CompoundDocument document, string password, out Stream stream)
    {
        var encryptedPackage = document.FindEntry(DirectoryEntryEncryptedPackage);
        var encryptionInfo = document.FindEntry(DirectoryEntryEncryptionInfo);

        if (encryptedPackage == null || encryptionInfo == null)
        {
            stream = null;
            return false;
        }

        var infoBytes = document.ReadStream(fileStream, encryptionInfo.StreamFirstSector, (int)encryptionInfo.StreamSize, encryptionInfo.IsEntryMiniStream);
        var encryption = EncryptionInfo.Create(infoBytes);

        if (encryption.VerifyPassword("VelvetSweatshop"))
        {
            // Magic password used for write-protected workbooks
            password = "VelvetSweatshop";
        }
        else if (password == null || !encryption.VerifyPassword(password))
        {
            throw new InvalidPasswordException(Errors.ErrorInvalidPassword);
        }

        var secretKey = encryption.GenerateSecretKey(password);
        var packageStream = new CompoundStream(document, fileStream, encryptedPackage.StreamFirstSector, (int)encryptedPackage.StreamSize, encryptedPackage.IsEntryMiniStream, false);

        stream = encryption.CreateEncryptedPackageStream(packageStream, secretKey);
        return true;
    }

    /// <summary>
    /// Checks if the stream contains Excel XML format (SpreadsheetML).
    /// </summary>
    /// <param name="fileStream">The file stream to check.</param>
    /// <returns>True if the stream appears to be Excel XML format.</returns>
    private static bool IsExcelXmlFormat(Stream fileStream)
    {
        if (fileStream == null || !fileStream.CanSeek || !fileStream.CanRead)
        {
            return false;
        }

        var originalPosition = fileStream.Position;
        try
        {
            fileStream.Seek(0, SeekOrigin.Begin);

            // Read first 1024 bytes to check for XML declaration and Workbook element
            var buffer = new byte[1024];
            var bytesRead = fileStream.Read(buffer, 0, buffer.Length);
            if (bytesRead < 5)
            {
                return false;
            }

            // Check for XML declaration (<?xml)
            var text = System.Text.Encoding.UTF8.GetString(buffer, 0, Math.Min(bytesRead, 1024));
            if (!text.TrimStart().StartsWith("<?xml", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            // Check for Workbook element (with or without namespace prefix)
            // Excel XML format uses <Workbook> or <ss:Workbook> as root element
            // Also check for the SpreadsheetML namespace
            var textLower = text.ToLowerInvariant();
            var hasWorkbook = textLower.Contains("<workbook") ||
                             textLower.Contains("<ss:workbook") ||
                             textLower.Contains("urn:schemas-microsoft-com:office:spreadsheet");

            if (hasWorkbook)
            {
                return true;
            }

            // If we haven't found Workbook in first 1024 bytes, try reading more
            // or use XML reader to check properly
            fileStream.Seek(0, SeekOrigin.Begin);
            return IsExcelXmlFormatUsingXmlReader(fileStream);
        }
        catch
        {
            return false;
        }
        finally
        {
            fileStream.Position = originalPosition;
        }
    }

    private static bool IsExcelXmlFormatUsingXmlReader(Stream fileStream)
    {
        try
        {
            using var xmlReader = System.Xml.XmlReader.Create(fileStream, new System.Xml.XmlReaderSettings
            {
                IgnoreWhitespace = true,
                IgnoreComments = true,
                IgnoreProcessingInstructions = false,
                DtdProcessing = System.Xml.DtdProcessing.Ignore
            });

            // Read until we find the root element
            while (xmlReader.Read())
            {
                if (xmlReader.NodeType == System.Xml.XmlNodeType.Element)
                {
                    var localName = xmlReader.LocalName;
                    var namespaceUri = xmlReader.NamespaceURI;

                    // Check for Workbook element (with or without namespace)
                    if (localName.Equals("Workbook", StringComparison.OrdinalIgnoreCase))
                    {
                        // Check if it's Excel XML namespace or no namespace (legacy format)
                        if (string.IsNullOrEmpty(namespaceUri) ||
                            namespaceUri.ToLowerInvariant().Contains("schemas-microsoft-com:office:spreadsheet"))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }
        catch
        {
            return false;
        }
    }
}
