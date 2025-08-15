using System;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using NPOI.SS.UserModel;

namespace JFToolkit.EncryptedExcel;

/// <summary>
/// Enhanced encryption helper that provides multiple fallback options
/// when Excel is not available
/// </summary>
public static class EnhancedEncryptionHelper
{
    /// <summary>
    /// Attempts to encrypt an Excel file using the best available method
    /// </summary>
    /// <param name="workbook">The workbook to encrypt</param>
    /// <param name="filePath">Target file path</param>
    /// <param name="password">Encryption password</param>
    /// <returns>EncryptionResult indicating what method was used</returns>
    public static EncryptionResult SaveEncryptedWithFallbacks(IWorkbook workbook, string filePath, string password)
    {
        Console.WriteLine("🔐 Attempting to save encrypted file...");
        Console.WriteLine($"   Target: {Path.GetFileName(filePath)}");
        Console.WriteLine($"   Trying multiple encryption methods...");
        Console.WriteLine();
        
        // Method 1: Excel COM Automation (best compatibility)
        if (ExcelAutomationHelper.IsExcelAvailable())
        {
            Console.WriteLine("Method 1: Excel COM Automation");
            Console.WriteLine("✅ Excel is available - using COM automation");
            
            var tempFile = SaveToTempFile(workbook, filePath);
            try
            {
                if (ExcelAutomationHelper.EncryptExcelFile(tempFile, filePath, password))
                {
                    Console.WriteLine("✅ SUCCESS: File encrypted using Excel COM automation");
                    return new EncryptionResult(true, EncryptionMethod.ExcelCOM, "Encrypted using Excel COM automation");
                }
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }
        else
        {
            Console.WriteLine("Method 1: Excel COM Automation");
            Console.WriteLine("❌ Excel not available - skipping COM automation");
        }
        
        // Method 2: LibreOffice Automation (cross-platform alternative)
        if (IsLibreOfficeAvailable())
        {
            Console.WriteLine("Method 2: LibreOffice Automation");
            Console.WriteLine("✅ LibreOffice detected - attempting automation");
            
            if (TryLibreOfficeEncryption(workbook, filePath, password))
            {
                Console.WriteLine("✅ SUCCESS: File encrypted using LibreOffice");
                return new EncryptionResult(true, EncryptionMethod.LibreOffice, "Encrypted using LibreOffice automation");
            }
            else
            {
                Console.WriteLine("❌ LibreOffice encryption failed");
            }
        }
        else
        {
            Console.WriteLine("Method 2: LibreOffice Automation");
            Console.WriteLine("❌ LibreOffice not available");
        }
        
        // Method 3: Python openpyxl (if available)
        if (IsPythonWithOpenpyxlAvailable())
        {
            Console.WriteLine("Method 3: Python openpyxl");
            Console.WriteLine("✅ Python with openpyxl detected");
            
            if (TryPythonEncryption(workbook, filePath, password))
            {
                Console.WriteLine("✅ SUCCESS: File encrypted using Python openpyxl");
                return new EncryptionResult(true, EncryptionMethod.PythonOpenpyxl, "Encrypted using Python openpyxl");
            }
            else
            {
                Console.WriteLine("❌ Python encryption failed");
            }
        }
        else
        {
            Console.WriteLine("Method 3: Python openpyxl");
            Console.WriteLine("❌ Python with openpyxl not available");
        }
        
        // Method 4: Fallback - Save unencrypted with detailed instructions
        Console.WriteLine("Method 4: Fallback - Unencrypted with instructions");
        Console.WriteLine("⚠️ All encryption methods failed - saving unencrypted");
        
        EncryptedExcelWriter.SaveToFile(workbook, filePath);
        
        var instructions = GenerateEncryptionInstructions(filePath);
        Console.WriteLine();
        Console.WriteLine("📋 ENCRYPTION INSTRUCTIONS:");
        Console.WriteLine("============================");
        foreach (var instruction in instructions)
        {
            Console.WriteLine($"   {instruction}");
        }
        
        return new EncryptionResult(false, EncryptionMethod.None, 
            "File saved unencrypted. Manual encryption required. See instructions above.");
    }
    
    private static string SaveToTempFile(IWorkbook workbook, string targetPath)
    {
        var tempFile = Path.GetTempFileName();
        var extension = Path.GetExtension(targetPath);
        var tempExcelFile = Path.ChangeExtension(tempFile, extension);
        
        EncryptedExcelWriter.SaveToFile(workbook, tempExcelFile);
        return tempExcelFile;
    }
    
    private static bool IsLibreOfficeAvailable()
    {
        try
        {
            // Check for LibreOffice installation
            var possiblePaths = new[]
            {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                "/usr/bin/libreoffice",
                "/usr/local/bin/libreoffice",
                "/Applications/LibreOffice.app/Contents/MacOS/soffice"
            };
            
            foreach (var path in possiblePaths)
            {
                if (File.Exists(path)) return true;
            }
            
            // Try to find it in PATH
            return CanExecuteCommand("soffice", "--version");
        }
        catch
        {
            return false;
        }
    }
    
    private static bool TryLibreOfficeEncryption(IWorkbook workbook, string filePath, string password)
    {
        try
        {
            var tempFile = SaveToTempFile(workbook, filePath);
            
            // LibreOffice command to open and save with password
            var arguments = $"--headless --convert-to xlsx --outdir \"{Path.GetDirectoryName(filePath)}\" " +
                           $"--password \"{password}\" \"{tempFile}\"";
            
            var result = ExecuteCommand("soffice", arguments);
            
            if (File.Exists(tempFile)) File.Delete(tempFile);
            
            return result && File.Exists(filePath);
        }
        catch
        {
            return false;
        }
    }
    
    private static bool IsPythonWithOpenpyxlAvailable()
    {
        try
        {
            return CanExecuteCommand("python", "-c \"import openpyxl; print('OK')\"") ||
                   CanExecuteCommand("python3", "-c \"import openpyxl; print('OK')\"");
        }
        catch
        {
            return false;
        }
    }
    
    private static bool TryPythonEncryption(IWorkbook workbook, string filePath, string password)
    {
        try
        {
            var tempFile = SaveToTempFile(workbook, filePath);
            
            // Create Python script for encryption
            var pythonScript = $@"
import openpyxl
from openpyxl.workbook.protection import WorkbookProtection

wb = openpyxl.load_workbook('{tempFile.Replace("\\", "\\\\")}')
wb.security = WorkbookProtection(workbookPassword='{password}')
wb.save('{filePath.Replace("\\", "\\\\")}')
print('SUCCESS')
";
            
            var scriptFile = Path.GetTempFileName() + ".py";
            File.WriteAllText(scriptFile, pythonScript);
            
            try
            {
                var result = ExecuteCommand("python", $"\"{scriptFile}\"") ||
                            ExecuteCommand("python3", $"\"{scriptFile}\"");
                
                return result && File.Exists(filePath);
            }
            finally
            {
                if (File.Exists(scriptFile)) File.Delete(scriptFile);
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }
        catch
        {
            return false;
        }
    }
    
    private static bool CanExecuteCommand(string command, string arguments)
    {
        try
        {
            using var process = new Process();
            process.StartInfo.FileName = command;
            process.StartInfo.Arguments = arguments;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            
            process.Start();
            process.WaitForExit(5000); // 5 second timeout
            
            return process.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }
    
    private static bool ExecuteCommand(string command, string arguments)
    {
        try
        {
            using var process = new Process();
            process.StartInfo.FileName = command;
            process.StartInfo.Arguments = arguments;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            
            process.Start();
            process.WaitForExit(30000); // 30 second timeout
            
            return process.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }
    
    private static string[] GenerateEncryptionInstructions(string filePath)
    {
        var fileName = Path.GetFileName(filePath);
        var isWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
        var isMac = RuntimeInformation.IsOSPlatform(OSPlatform.OSX);
        
        var instructions = new[]
        {
            $"Your file '{fileName}' was saved WITHOUT encryption.",
            "To encrypt it manually, choose one of these options:",
            "",
            "🏢 OPTION 1: Microsoft Excel",
            "   1. Open the file in Excel",
            "   2. Click File → Info → Protect Workbook → Encrypt with Password",
            "   3. Enter your password and save",
            "",
            "🆓 OPTION 2: LibreOffice Calc (Free)",
            "   1. Download LibreOffice from https://www.libreoffice.org/",
            "   2. Open the file in LibreOffice Calc", 
            "   3. File → Save As → Check 'Save with password'",
            "   4. Enter your password and save",
            "",
            "🐍 OPTION 3: Python Script (Technical)",
            "   1. Install Python: pip install openpyxl",
            "   2. Run: python -c \"import openpyxl; wb=openpyxl.load_workbook('file.xlsx'); wb.security=openpyxl.workbook.protection.WorkbookProtection(workbookPassword='your_password'); wb.save('encrypted.xlsx')\"",
            "",
            "💡 TIP: Install Excel, LibreOffice, or Python for automatic encryption"
        };
        
        return instructions;
    }
}

/// <summary>
/// Result of an encryption attempt
/// </summary>
/// <summary>
/// Result of an encryption operation
/// </summary>
public class EncryptionResult
{
    /// <summary>
    /// Gets a value indicating whether the encryption operation was successful
    /// </summary>
    public bool Success { get; }
    
    /// <summary>
    /// Gets the encryption method that was used
    /// </summary>
    public EncryptionMethod Method { get; }
    
    /// <summary>
    /// Gets a message describing the result of the operation
    /// </summary>
    public string Message { get; }
    
    /// <summary>
    /// Initializes a new instance of the EncryptionResult class
    /// </summary>
    /// <param name="success">Whether the operation was successful</param>
    /// <param name="method">The encryption method used</param>
    /// <param name="message">A message describing the result</param>
    public EncryptionResult(bool success, EncryptionMethod method, string message)
    {
        Success = success;
        Method = method;
        Message = message;
    }
}

/// <summary>
/// Available encryption methods
/// </summary>
public enum EncryptionMethod
{
    /// <summary>
    /// No encryption method available
    /// </summary>
    None,
    
    /// <summary>
    /// Excel COM automation encryption
    /// </summary>
    ExcelCOM,
    
    /// <summary>
    /// LibreOffice automation encryption
    /// </summary>
    LibreOffice,
    
    /// <summary>
    /// Python openpyxl library encryption
    /// </summary>
    PythonOpenpyxl,
    
    /// <summary>
    /// OpenXML SDK encryption (for future implementation)
    /// </summary>
    OpenXMLSDK
}
