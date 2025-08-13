using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace JFToolkit.EncryptedExcel;

/// <summary>
/// Helper class for Excel automation tasks like encrypting files
/// </summary>
public static class ExcelAutomationHelper
{
    /// <summary>
    /// Encrypts an Excel file using Excel automation (requires Excel to be installed)
    /// </summary>
    /// <param name="unencryptedFilePath">Path to the unencrypted Excel file</param>
    /// <param name="encryptedFilePath">Path where to save the encrypted Excel file</param>
    /// <param name="password">Password to encrypt with</param>
    /// <returns>True if successful, false if Excel is not available</returns>
    public static bool EncryptExcelFile(string unencryptedFilePath, string encryptedFilePath, string password)
    {
        if (!File.Exists(unencryptedFilePath))
            throw new FileNotFoundException($"Source file not found: {unencryptedFilePath}");

        if (string.IsNullOrWhiteSpace(password))
            throw new ArgumentException("Password cannot be null or empty", nameof(password));

        object? excel = null;
        object? workbook = null;
        
        try
        {
            // Try to create Excel application with timeout
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                return false;

            excel = Activator.CreateInstance(excelType);
            if (excel == null)
                return false;

            // Configure Excel
            excel.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, excel, new object[] { false });
            excel.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, excel, new object[] { false });
            excel.GetType().InvokeMember("ScreenUpdating", BindingFlags.SetProperty, null, excel, new object[] { false });

            // Get workbooks collection
            var workbooks = excel.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, excel, null);
            
            // Open the unencrypted file
            workbook = workbooks?.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks, new object[] { Path.GetFullPath(unencryptedFilePath) });
            
            if (workbook == null)
                return false;
                
            // Save with password protection using reflection
            workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, workbook, new object?[] {
                Path.GetFullPath(encryptedFilePath),  // Filename
                Type.Missing,                         // FileFormat
                password,                            // Password
                password,                            // WriteResPassword
                Type.Missing,                        // ReadOnlyRecommended
                Type.Missing,                        // CreateBackup
                Type.Missing,                        // AccessMode
                Type.Missing,                        // ConflictResolution
                Type.Missing,                        // AddToMru
                Type.Missing,                        // TextCodepage
                Type.Missing,                        // TextVisualLayout
                Type.Missing                         // Local
            });
            
            return true;
        }
        catch (Exception)
        {
            // Excel not available or COM error
            return false;
        }
        finally
        {
            // Clean up COM objects
            try
            {
                if (workbook != null)
                {
                    workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, new object[] { false });
                    Marshal.ReleaseComObject(workbook);
                }
            }
            catch { }
            
            try
            {
                if (excel != null)
                {
                    excel.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, excel, null);
                    Marshal.ReleaseComObject(excel);
                }
            }
            catch { }
            
            // Force garbage collection to release COM objects
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    /// <summary>
    /// Checks if Excel is available for automation
    /// </summary>
    /// <returns>True if Excel is available, false otherwise</returns>
    public static bool IsExcelAvailable()
    {
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            return excelType != null;
        }
        catch
        {
            return false;
        }
    }
}
