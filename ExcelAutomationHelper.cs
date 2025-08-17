using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;

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

        // Kill any existing Excel processes first to avoid conflicts
        try
        {
            KillExcelProcesses();
            Thread.Sleep(1000); // Give time for cleanup
        }
        catch { /* Ignore errors in cleanup */ }

        object? excel = null;
        object? workbook = null;
        
        try
        {
            // Try to create Excel application with timeout and retries
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                return false;

            // Retry Excel creation up to 3 times with delays
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    excel = Activator.CreateInstance(excelType);
                    if (excel != null) break;
                }
                catch
                {
                    if (attempt == 3) return false;
                    Thread.Sleep(2000 * attempt); // Increasing delays
                }
            }

            if (excel == null)
                return false;

            // Configure Excel for automation with comprehensive settings
            ConfigureExcelForAutomation(excel);

            // Get workbooks collection
            var workbooks = excel.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, excel, null);
            
            // Open the unencrypted file with retry logic
            // Open the unencrypted file with retry logic
            for (int attempt = 1; attempt <= 2; attempt++)
            {
                try
                {
                    workbook = workbooks?.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks, new object[] { Path.GetFullPath(unencryptedFilePath) });
                    if (workbook != null) break;
                }
                catch
                {
                    if (attempt == 2) return false;
                    Thread.Sleep(1000);
                }
            }
            
            if (workbook == null)
                return false;

            // Determine the Excel file format based on target extension (.xlsx/.xls supported)
            var targetExtension = Path.GetExtension(encryptedFilePath).ToLowerInvariant();
            return SaveAsEncryptedStandard(workbook, encryptedFilePath, password, targetExtension);
        }
        catch (Exception)
        {
            // Excel not available or COM error
            return false;
        }
        finally
        {
            // Clean up COM objects with enhanced cleanup
            CleanupExcelObjects(workbook, excel);
        }
    }


    /// <summary>
    /// Standard save method for non-.xlsm files
    /// </summary>
    private static bool SaveAsEncryptedStandard(object workbook, string filePath, string password, string targetExtension)
    {
        try
        {
            object fileFormat = Type.Missing;
            
            // Set appropriate file format
            switch (targetExtension)
            {
                case ".xlsx":
                    fileFormat = 51; // xlOpenXMLWorkbook
                    break;
                case ".xls":
                    fileFormat = 56; // xlExcel8 (Excel 97-2003)
                    break;
                default:
                    fileFormat = 51; // Default to .xlsx
                    break;
            }

            workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, workbook, new object?[] {
                Path.GetFullPath(filePath),  // Filename
                fileFormat,                  // FileFormat
                password,                    // Password (read password only)
                Type.Missing,                // WriteResPassword (don't set - avoids double prompt)
                Type.Missing,                // ReadOnlyRecommended
                Type.Missing,                // CreateBackup
                Type.Missing,                // AccessMode
                Type.Missing,                // ConflictResolution
                Type.Missing,                // AddToMru
                Type.Missing,                // TextCodepage
                Type.Missing,                // TextVisualLayout
                Type.Missing                 // Local
            });
            
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Configure Excel application for reliable automation
    /// </summary>
    private static void ConfigureExcelForAutomation(object excel)
    {
        try
        {
            excel.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, excel, new object[] { false });
            excel.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, excel, new object[] { false });
            excel.GetType().InvokeMember("ScreenUpdating", BindingFlags.SetProperty, null, excel, new object[] { false });
            excel.GetType().InvokeMember("EnableEvents", BindingFlags.SetProperty, null, excel, new object[] { false });
            excel.GetType().InvokeMember("Interactive", BindingFlags.SetProperty, null, excel, new object[] { false });
            
            // Additional settings for better stability
            try
            {
                excel.GetType().InvokeMember("Calculation", BindingFlags.SetProperty, null, excel, new object[] { -4135 }); // xlCalculationManual
                excel.GetType().InvokeMember("StatusBar", BindingFlags.SetProperty, null, excel, new object[] { false });
            }
            catch { /* Some settings might not be available in all Excel versions */ }
        }
        catch { /* Continue even if some configuration fails */ }
    }

    /// <summary>
    /// Enhanced cleanup of Excel COM objects
    /// </summary>
    private static void CleanupExcelObjects(object? workbook, object? excel)
    {
        try
        {
            if (workbook != null)
            {
                try
                {
                    workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, new object[] { false });
                }
                catch { }
                
                try
                {
                    Marshal.ReleaseComObject(workbook);
                }
                catch { }
            }
        }
        catch { }
        
        try
        {
            if (excel != null)
            {
                try
                {
                    excel.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, excel, null);
                }
                catch { }
                
                try
                {
                    Marshal.ReleaseComObject(excel);
                }
                catch { }
            }
        }
        catch { }
        
        // Force multiple garbage collections for better COM cleanup
        try
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            
            // Give time for cleanup
            Thread.Sleep(500);
        }
        catch { }
    }

    /// <summary>
    /// Kill existing Excel processes to avoid conflicts
    /// </summary>
    private static void KillExcelProcesses()
    {
        try
        {
            var processes = Process.GetProcessesByName("EXCEL");
            foreach (var process in processes)
            {
                try
                {
                    process.Kill();
                    process.WaitForExit(2000); // Wait up to 2 seconds
                }
                catch { }
                finally
                {
                    process.Dispose();
                }
            }
        }
        catch { }
    }

    /// <summary>
    /// Checks if Excel is available for automation with enhanced detection
    /// </summary>
    /// <returns>True if Excel is available, false otherwise</returns>
    public static bool IsExcelAvailable()
    {
        try
        {
            // First check if Excel is registered as COM server
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                return false;

            // Try to actually create an instance to verify it works
            object? excel = null;
            try
            {
                excel = Activator.CreateInstance(excelType);
                if (excel == null)
                    return false;

                // Try to access a basic property to ensure it's functional
                var version = excel.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, excel, null);
                return version != null;
            }
            finally
            {
                if (excel != null)
                {
                    try
                    {
                        excel.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, excel, null);
                        Marshal.ReleaseComObject(excel);
                    }
                    catch { }
                }
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Gets the version of Excel if available
    /// </summary>
    /// <returns>Excel version string or null if not available</returns>
    public static string? GetExcelVersion()
    {
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                return null;

            object? excel = null;
            try
            {
                excel = Activator.CreateInstance(excelType);
                if (excel == null)
                    return null;

                var version = excel.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, excel, null);
                return version?.ToString();
            }
            finally
            {
                if (excel != null)
                {
                    try
                    {
                        excel.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, excel, null);
                        Marshal.ReleaseComObject(excel);
                    }
                    catch { }
                }
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        catch
        {
            return null;
        }
    }
}
