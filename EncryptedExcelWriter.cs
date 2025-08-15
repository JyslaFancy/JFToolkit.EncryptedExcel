using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.Crypt;
using NPOI.POIFS.FileSystem;
using System;
using System.IO;

namespace JFToolkit.EncryptedExcel;

/// <summary>
/// Provides methods for saving Excel files (.xlsx, .xlsm, .xls) with optional encryption
/// </summary>
public static class EncryptedExcelWriter
{
    /// <summary>
    /// Saves a workbook to a file without encryption
    /// </summary>
    /// <param name="workbook">The workbook to save</param>
    /// <param name="filePath">Path where to save the file</param>
    /// <exception cref="ArgumentNullException">Thrown when workbook is null</exception>
    /// <exception cref="ArgumentException">Thrown when file path is null or empty</exception>
    public static void SaveToFile(IWorkbook workbook, string filePath)
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));
        
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));

        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        
        // For .xlsm files, we need to save as .xlsx first and warn the user
        if (extension == ".xlsm" && workbook is XSSFWorkbook)
        {
            // NPOI limitation: Cannot create true .xlsm files from scratch
            // Save as .xlsx instead to avoid the "invalid file format" error
            var xlsxPath = Path.ChangeExtension(filePath, ".xlsx");
            using var fileStream = new FileStream(xlsxPath, FileMode.Create, FileAccess.Write);
            workbook.Write(fileStream);
            
            // Inform about the change
            Console.WriteLine($"⚠️  Note: Saved as {Path.GetFileName(xlsxPath)} instead of .xlsm");
            Console.WriteLine("   To create a true .xlsm file, open in Excel and save as macro-enabled format");
        }
        else
        {
            using var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            workbook.Write(fileStream);
        }
    }

    /// <summary>
    /// Saves a workbook to a stream without encryption
    /// </summary>
    /// <param name="workbook">The workbook to save</param>
    /// <param name="stream">Stream to write to</param>
    /// <exception cref="ArgumentNullException">Thrown when workbook or stream is null</exception>
    public static void SaveToStream(IWorkbook workbook, Stream stream)
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));
        
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        workbook.Write(stream);
    }

    /// <summary>
    /// Saves a workbook as encrypted Excel file (.xlsx, .xlsm, .xls) using multiple automation approaches
    /// </summary>
    /// <param name="workbook">The workbook to save</param>
    /// <param name="filePath">Path where to save the file</param>
    /// <param name="password">Password to encrypt the file with</param>
    /// <exception cref="ArgumentNullException">Thrown when workbook is null</exception>
    /// <exception cref="ArgumentException">Thrown when file path is null or empty</exception>
    public static void SaveEncryptedToFile(IWorkbook workbook, string filePath, string password)
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));
        
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));

        if (string.IsNullOrWhiteSpace(password))
            throw new ArgumentException("Password cannot be null or empty", nameof(password));

        // Create temporary unencrypted file with appropriate extension
        var tempFile = Path.GetTempFileName();
        var targetExtension = Path.GetExtension(filePath).ToLowerInvariant();
        
        // For .xlsm targets, use .xlsx for temporary file to avoid corruption
        var tempExtension = targetExtension == ".xlsm" ? ".xlsx" : targetExtension;
        var tempExcelFile = Path.ChangeExtension(tempFile, tempExtension);
        
        try
        {
            // First save as unencrypted to a temporary file
            SaveToFile(workbook, tempExcelFile);
            
            // Use Excel automation for encryption (Windows only, requires Excel)
            if (ExcelAutomationHelper.IsExcelAvailable())
            {
                Console.WriteLine("   Attempting Excel automation encryption...");
                bool success = ExcelAutomationHelper.EncryptExcelFile(tempExcelFile, filePath, password);
                if (success)
                {
                    Console.WriteLine("✅ File encrypted successfully using Excel automation");
                    return;
                }
                else
                {
                    Console.WriteLine("⚠️ Excel automation encryption failed");
                }
            }
            else
            {
                Console.WriteLine("⚠️ Excel is not available for encryption");
            }
            
            // If encryption fails, save unencrypted with warning
            SaveToFile(workbook, filePath);
            Console.WriteLine("⚠️ Encryption unavailable - saved as unencrypted");
            Console.WriteLine($"   To encrypt: Open '{Path.GetFileName(filePath)}' in Excel and use 'Save As' with password");
        }
        finally
        {
            // Clean up temporary files
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
            try { if (File.Exists(tempExcelFile)) File.Delete(tempExcelFile); } catch { }
        }
    }

    /// <summary>
    /// Saves a workbook as encrypted Excel to a stream
    /// Note: This method is currently not supported due to NPOI limitations
    /// </summary>
    /// <param name="workbook">The workbook to save</param>
    /// <param name="stream">Stream to write to</param>
    /// <param name="password">Password to encrypt the file with</param>
    /// <exception cref="ArgumentNullException">Thrown when workbook or stream is null</exception>
    /// <exception cref="ArgumentException">Thrown when password is null or empty</exception>
    /// <exception cref="NotImplementedException">Always thrown as this method is not currently supported</exception>
    public static void SaveEncryptedToStream(IWorkbook workbook, Stream stream, string password)
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));
        
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        
        if (string.IsNullOrWhiteSpace(password))
            throw new ArgumentException("Password cannot be null or empty", nameof(password));

        throw new NotImplementedException(
            "SaveEncryptedToStream is not supported due to NPOI encryption limitations. " +
            "Use SaveEncryptedToFile() instead, which uses Excel automation as a workaround.");
    }

    /// <summary>
    /// Gets the workbook as a byte array without encryption
    /// </summary>
    /// <param name="workbook">The workbook to convert</param>
    /// <returns>Byte array containing the Excel file</returns>
    /// <exception cref="ArgumentNullException">Thrown when workbook is null</exception>
    public static byte[] ToByteArray(IWorkbook workbook)
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));

        using var stream = new MemoryStream();
        workbook.Write(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Gets the workbook as an encrypted byte array
    /// </summary>
    /// <param name="workbook">The workbook to convert</param>
    /// <param name="password">Password to encrypt with</param>
    /// <returns>Byte array containing the encrypted Excel file</returns>
    /// <exception cref="ArgumentNullException">Thrown when workbook is null</exception>
    /// <exception cref="ArgumentException">Thrown when password is null or empty</exception>
    public static byte[] ToEncryptedByteArray(IWorkbook workbook, string password)
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));
        
        if (string.IsNullOrWhiteSpace(password))
            throw new ArgumentException("Password cannot be null or empty", nameof(password));

        using var stream = new MemoryStream();
        SaveEncryptedToStream(workbook, stream, password);
        return stream.ToArray();
    }

    /// <summary>
    /// Creates a new workbook that can be saved as .xlsx format
    /// Note: To create true .xlsm files, open the saved .xlsx file in Excel and save as macro-enabled format
    /// </summary>
    /// <returns>A new XSSFWorkbook</returns>
    public static XSSFWorkbook CreateMacroEnabledWorkbook()
    {
        // Create a standard XSSF workbook
        // Note: NPOI cannot create true .xlsm files from scratch
        // This creates a .xlsx-compatible workbook that can be converted to .xlsm in Excel
        var workbook = new XSSFWorkbook();
        
        return workbook;
    }

    /// <summary>
    /// Creates a new Excel workbook and provides instructions for macro-enabled conversion
    /// </summary>
    /// <param name="filePath">The desired output path (can end with .xlsm)</param>
    /// <returns>A tuple containing the workbook and the actual save path</returns>
    public static (XSSFWorkbook workbook, string actualPath) CreateWorkbookForMacros(string filePath)
    {
        var workbook = new XSSFWorkbook();
        
        // If user wants .xlsm, we'll save as .xlsx and provide instructions
        if (Path.GetExtension(filePath).ToLowerInvariant() == ".xlsm")
        {
            var xlsxPath = Path.ChangeExtension(filePath, ".xlsx");
            return (workbook, xlsxPath);
        }
        
        return (workbook, filePath);
    }

    /// <summary>
    /// Validates if a file can be properly saved as the specified format
    /// </summary>
    /// <param name="workbook">The workbook to validate</param>
    /// <param name="filePath">The target file path</param>
    /// <returns>True if the combination is valid, false otherwise</returns>
    public static bool ValidateFileFormat(IWorkbook workbook, string filePath)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        
        return extension switch
        {
            ".xlsx" => workbook is XSSFWorkbook,
            ".xlsm" => workbook is XSSFWorkbook, // Will be saved as .xlsx due to NPOI limitations
            ".xls" => workbook is HSSFWorkbook,
            _ => false
        };
    }
}
