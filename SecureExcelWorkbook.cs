using System;
using System.IO;
using NPOI.SS.UserModel;

namespace JFToolkit.EncryptedExcel;

/// <summary>
/// Secure Excel Workbook API for working with encrypted macro-enabled Excel files (.xlsm)
/// Provides: Open encrypted .xlsm → Modify content → Save as encrypted .xlsm
/// 
/// IMPORTANT: Reading and modifying works on any platform. However, saving WITH encryption 
/// requires Microsoft Excel to be installed on Windows. Without Excel, files can still be 
/// saved but without encryption.
/// 
/// Supports opening password-protected Excel files, modifying cell values, and saving to new encrypted files.
/// </summary>
public class SecureExcelWorkbook : IDisposable
{
    private EncryptedExcelReader? _reader;
    private string _originalFilePath = string.Empty;
    private string _password = string.Empty;
    private bool _disposed = false;

    /// <summary>
    /// Gets the workbook for modification
    /// </summary>
    public IWorkbook? Workbook => _reader?.Workbook;

    /// <summary>
    /// Opens an encrypted .xlsm file for modification
    /// </summary>
    /// <param name="filePath">Path to the encrypted .xlsm file</param>
    /// <param name="password">Password to decrypt the file</param>
    /// <returns>SecureExcelWorkbook instance for chaining operations</returns>
    public static SecureExcelWorkbook Open(string filePath, string password)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"File not found: {filePath}");

        if (string.IsNullOrWhiteSpace(password))
            throw new ArgumentException("Password cannot be null or empty", nameof(password));

        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (extension != ".xlsm")
            throw new ArgumentException("File must be a .xlsm (macro-enabled) Excel file", nameof(filePath));

        var manager = new SecureExcelWorkbook();
        manager._originalFilePath = filePath;
        manager._password = password;
        manager._reader = EncryptedExcelReader.OpenFile(filePath, password);

        if (manager._reader.Workbook == null)
            throw new InvalidOperationException($"Could not open encrypted file: {filePath}");

        return manager;
    }

    private SecureExcelWorkbook() { }

    /// <summary>
    /// Saves the modified workbook back to the original file with the same password
    /// </summary>
    /// <returns>True if successful, false otherwise</returns>
    public bool Save()
    {
        return SaveAs(_originalFilePath);
    }

    /// <summary>
    /// Saves the modified workbook to a new file with the same password encryption.
    /// 
    /// IMPORTANT: This method requires Microsoft Excel to be installed on Windows for encryption.
    /// If Excel is not available, the method returns false and the file is not saved.
    /// 
    /// For cross-platform compatibility, use Workbook.SaveToFile() to save without encryption.
    /// </summary>
    /// <param name="filePath">Path where to save the encrypted file</param>
    /// <returns>True if successfully saved with encryption, false if Excel is unavailable</returns>
    public bool SaveAs(string filePath)
    {
        if (_disposed || Workbook == null)
            throw new ObjectDisposedException(nameof(SecureExcelWorkbook));

        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (extension != ".xlsm")
            throw new ArgumentException("Target file must be a .xlsm (macro-enabled) Excel file", nameof(filePath));

        try
        {
            // Try Excel automation first (most reliable for .xlsm encryption)
            if (ExcelAutomationHelper.IsExcelAvailable())
            {
                // Save to temporary unencrypted file first
                string tempFile = Path.GetTempFileName() + ".xlsm";
                
                try
                {
                    using (var fileStream = new FileStream(tempFile, FileMode.Create, FileAccess.Write))
                    {
                        Workbook.Write(fileStream);
                    }

                    // Use Excel automation to encrypt
                    bool success = ExcelAutomationHelper.EncryptExcelFile(tempFile, filePath, _password);
                    
                    if (success && File.Exists(filePath))
                    {
                        // Verify we can read it back
                        try
                        {
                            using var testReader = EncryptedExcelReader.OpenFile(filePath, _password);
                            return testReader.Workbook != null;
                        }
                        catch
                        {
                            return false;
                        }
                    }
                    
                    return false;
                }
                finally
                {
                    // Clean up temp file
                    try { File.Delete(tempFile); } catch { }
                }
            }
            else
            {
                // Excel not available - save unencrypted and provide guidance
                Console.WriteLine("⚠️ Excel not available for encryption.");
                Console.WriteLine("❌ No encryption method available. Saving unencrypted file.");
                Console.WriteLine($"📋 To encrypt manually:");
                Console.WriteLine($"   1. Open {filePath} in Excel");
                Console.WriteLine($"   2. Save As → Tools → General Options → Set password to: {_password}");
                
                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    Workbook.Write(fileStream);
                }
                
                return false; // Return false because file is not encrypted
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Save failed: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Quick helper to modify a cell and save
    /// </summary>
    /// <param name="sheetIndex">Sheet index (0-based)</param>
    /// <param name="row">Row index (0-based)</param>
    /// <param name="col">Column index (0-based)</param>
    /// <param name="value">New value for the cell</param>
    /// <returns>True if successful</returns>
    public bool SetCellValue(int sheetIndex, int row, int col, object value)
    {
        if (_disposed || Workbook == null)
            return false;

        try
        {
            var sheet = Workbook.GetSheetAt(sheetIndex);
            var rowObj = sheet.GetRow(row) ?? sheet.CreateRow(row);
            var cell = rowObj.GetCell(col) ?? rowObj.CreateCell(col);

            switch (value)
            {
                case string str:
                    cell.SetCellValue(str);
                    break;
                case double dbl:
                    cell.SetCellValue(dbl);
                    break;
                case int integer:
                    cell.SetCellValue(integer);
                    break;
                case DateTime date:
                    cell.SetCellValue(date);
                    break;
                case bool boolean:
                    cell.SetCellValue(boolean);
                    break;
                default:
                    cell.SetCellValue(value?.ToString() ?? "");
                    break;
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Quick helper to get a cell value
    /// </summary>
    /// <param name="sheetIndex">Sheet index (0-based)</param>
    /// <param name="row">Row index (0-based)</param>
    /// <param name="col">Column index (0-based)</param>
    /// <returns>Cell value or null if not found</returns>
    public object? GetCellValue(int sheetIndex, int row, int col)
    {
        if (_disposed || Workbook == null)
            return null;

        try
        {
            var sheet = Workbook.GetSheetAt(sheetIndex);
            var rowObj = sheet.GetRow(row);
            if (rowObj == null) return null;

            var cell = rowObj.GetCell(col);
            if (cell == null) return null;

            return cell.CellType switch
            {
                CellType.String => cell.StringCellValue,
                CellType.Numeric => DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue : cell.NumericCellValue,
                CellType.Boolean => cell.BooleanCellValue,
                CellType.Formula => cell.CellFormula,
                _ => null
            };
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Releases all resources used by the SecureExcelWorkbook.
    /// </summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _reader?.Dispose();
            _disposed = true;
        }
    }
}
