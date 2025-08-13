using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.POIFS.Crypt;
using NPOI.POIFS.FileSystem;
using System;
using System.IO;

namespace JFToolkit.EncryptedExcel;

/// <summary>
/// Provides methods for opening and reading password-encrypted Excel files
/// </summary>
public class EncryptedExcelReader : IDisposable
{
    private IWorkbook? _workbook;
    private bool _disposed;

    /// <summary>
    /// Gets the underlying NPOI workbook instance
    /// </summary>
    public IWorkbook? Workbook => _workbook;

    /// <summary>
    /// Opens an encrypted Excel file from a file path
    /// </summary>
    /// <param name="filePath">Path to the Excel file</param>
    /// <param name="password">Password to decrypt the file</param>
    /// <returns>EncryptedExcelReader instance</returns>
    /// <exception cref="ArgumentException">Thrown when file path is null or empty</exception>
    /// <exception cref="FileNotFoundException">Thrown when file doesn't exist</exception>
    /// <exception cref="InvalidOperationException">Thrown when password is incorrect or file is corrupted</exception>
    public static EncryptedExcelReader OpenFile(string filePath, string password)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));

        if (!File.Exists(filePath))
            throw new FileNotFoundException($"File not found: {filePath}");

        using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        return OpenStream(fileStream, password);
    }

    /// <summary>
    /// Opens an encrypted Excel file from a stream
    /// </summary>
    /// <param name="stream">Stream containing the Excel file data</param>
    /// <param name="password">Password to decrypt the file</param>
    /// <returns>EncryptedExcelReader instance</returns>
    /// <exception cref="ArgumentNullException">Thrown when stream is null</exception>
    /// <exception cref="InvalidOperationException">Thrown when password is incorrect or file is corrupted</exception>
    public static EncryptedExcelReader OpenStream(Stream stream, string password)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        var reader = new EncryptedExcelReader();
        
        try
        {
            // Copy stream to memory to allow multiple attempts
            var memoryStream = new MemoryStream();
            stream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            try
            {
                // Try to open as encrypted XLSX file
                var fs = new POIFSFileSystem(memoryStream);
                var info = new EncryptionInfo(fs);
                var decryptor = Decryptor.GetInstance(info);
                
                if (!decryptor.VerifyPassword(password))
                {
                    throw new InvalidOperationException("Invalid password provided.");
                }

                var dataStream = decryptor.GetDataStream(fs);
                reader._workbook = new XSSFWorkbook(dataStream);
                
                // Clean up
                dataStream?.Dispose();
                fs?.Close();
            }
            catch (Exception ex) when (!(ex is InvalidOperationException))
            {
                try
                {
                    // Fallback: try as XLS format or non-encrypted
                    memoryStream.Position = 0;
                    reader._workbook = WorkbookFactory.Create(memoryStream);
                }
                catch
                {
                    throw new InvalidOperationException("Unable to open the file. Please check the password and ensure the file is a valid Excel file.", ex);
                }
            }

            return reader;
        }
        catch
        {
            reader.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Opens an encrypted Excel file from a byte array
    /// </summary>
    /// <param name="data">Byte array containing the Excel file data</param>
    /// <param name="password">Password to decrypt the file</param>
    /// <returns>EncryptedExcelReader instance</returns>
    /// <exception cref="ArgumentNullException">Thrown when data is null</exception>
    /// <exception cref="InvalidOperationException">Thrown when password is incorrect or file is corrupted</exception>
    public static EncryptedExcelReader OpenBytes(byte[] data, string password)
    {
        if (data == null)
            throw new ArgumentNullException(nameof(data));

        using var stream = new MemoryStream(data);
        return OpenStream(stream, password);
    }

    /// <summary>
    /// Gets the number of sheets in the workbook
    /// </summary>
    public int NumberOfSheets => _workbook?.NumberOfSheets ?? 0;

    /// <summary>
    /// Gets a sheet by index
    /// </summary>
    /// <param name="index">Zero-based sheet index</param>
    /// <returns>ISheet instance</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when index is out of range</exception>
    /// <exception cref="ObjectDisposedException">Thrown when the reader has been disposed</exception>
    public ISheet GetSheetAt(int index)
    {
        ThrowIfDisposed();
        
        if (index < 0 || index >= NumberOfSheets)
            throw new ArgumentOutOfRangeException(nameof(index), $"Sheet index {index} is out of range. Valid range is 0 to {NumberOfSheets - 1}");

        return _workbook!.GetSheetAt(index);
    }

    /// <summary>
    /// Gets a sheet by name
    /// </summary>
    /// <param name="name">Sheet name</param>
    /// <returns>ISheet instance or null if not found</returns>
    /// <exception cref="ObjectDisposedException">Thrown when the reader has been disposed</exception>
    public ISheet? GetSheet(string name)
    {
        ThrowIfDisposed();
        return _workbook?.GetSheet(name);
    }

    /// <summary>
    /// Gets the name of a sheet at the specified index
    /// </summary>
    /// <param name="index">Zero-based sheet index</param>
    /// <returns>Sheet name</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when index is out of range</exception>
    /// <exception cref="ObjectDisposedException">Thrown when the reader has been disposed</exception>
    public string GetSheetName(int index)
    {
        ThrowIfDisposed();
        
        if (index < 0 || index >= NumberOfSheets)
            throw new ArgumentOutOfRangeException(nameof(index), $"Sheet index {index} is out of range. Valid range is 0 to {NumberOfSheets - 1}");

        return _workbook!.GetSheetName(index);
    }

    /// <summary>
    /// Gets all sheet names in the workbook
    /// </summary>
    /// <returns>Array of sheet names</returns>
    /// <exception cref="ObjectDisposedException">Thrown when the reader has been disposed</exception>
    public string[] GetSheetNames()
    {
        ThrowIfDisposed();
        
        var names = new string[NumberOfSheets];
        for (int i = 0; i < NumberOfSheets; i++)
        {
            names[i] = _workbook!.GetSheetName(i);
        }
        return names;
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(EncryptedExcelReader));
    }

    /// <summary>
    /// Disposes the reader and releases resources
    /// </summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _workbook?.Close();
            _workbook = null;
            _disposed = true;
        }
    }
}
