using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace JFToolkit.EncryptedExcel;

/// <summary>
/// PowerShell-based Excel automation for encryption tasks
/// </summary>
public static class PowerShellExcelHelper
{
    /// <summary>
    /// Encrypts an Excel file using PowerShell and Excel COM automation
    /// </summary>
    /// <param name="unencryptedFilePath">Path to the unencrypted Excel file</param>
    /// <param name="encryptedFilePath">Path where to save the encrypted Excel file</param>
    /// <param name="password">Password to encrypt with</param>
    /// <returns>True if successful, false otherwise</returns>
    public static bool EncryptExcelFile(string unencryptedFilePath, string encryptedFilePath, string password)
    {
        if (!File.Exists(unencryptedFilePath))
            throw new FileNotFoundException($"Source file not found: {unencryptedFilePath}");

        if (string.IsNullOrWhiteSpace(password))
            throw new ArgumentException("Password cannot be null or empty", nameof(password));

        try
        {
            var script = CreateEncryptionScript(unencryptedFilePath, encryptedFilePath, password);
            return ExecutePowerShellScript(script);
        }
        catch (Exception)
        {
            return false;
        }
    }

    private static string CreateEncryptionScript(string inputFile, string outputFile, string password)
    {
        var script = new StringBuilder();
        script.AppendLine("try {");
        script.AppendLine("    $excel = New-Object -ComObject Excel.Application");
        script.AppendLine("    $excel.Visible = $false");
        script.AppendLine("    $excel.DisplayAlerts = $false");
        script.AppendLine("    $excel.ScreenUpdating = $false");
        script.AppendLine();
        script.AppendLine($"    $inputPath = '{inputFile.Replace("'", "''")}'");
        script.AppendLine($"    $outputPath = '{outputFile.Replace("'", "''")}'");
        script.AppendLine($"    $password = '{password.Replace("'", "''")}'");
        script.AppendLine();
        script.AppendLine("    # Open the unencrypted file");
        script.AppendLine("    $workbook = $excel.Workbooks.Open($inputPath)");
        script.AppendLine();
        script.AppendLine("    # Save with password protection");
        script.AppendLine("    $workbook.SaveAs($outputPath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook, $password, $password)");
        script.AppendLine();
        script.AppendLine("    # Clean up");
        script.AppendLine("    $workbook.Close($false)");
        script.AppendLine("    $excel.Quit()");
        script.AppendLine("    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null");
        script.AppendLine("    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null");
        script.AppendLine("    [System.GC]::Collect()");
        script.AppendLine("    [System.GC]::WaitForPendingFinalizers()");
        script.AppendLine();
        script.AppendLine("    Write-Output 'SUCCESS'");
        script.AppendLine("} catch {");
        script.AppendLine("    Write-Error $_.Exception.Message");
        script.AppendLine("    if ($workbook) { $workbook.Close($false) }");
        script.AppendLine("    if ($excel) { $excel.Quit() }");
        script.AppendLine("    exit 1");
        script.AppendLine("}");

        return script.ToString();
    }

    private static bool ExecutePowerShellScript(string script)
    {
        try
        {
            using var process = new Process();
            process.StartInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -Command -",
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            process.Start();
            
            // Send the script to PowerShell
            process.StandardInput.WriteLine(script);
            process.StandardInput.Close();

            // Wait for completion with timeout
            bool finished = process.WaitForExit(30000); // 30 second timeout
            
            if (!finished)
            {
                process.Kill();
                return false;
            }

            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            return process.ExitCode == 0 && output.Contains("SUCCESS");
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Checks if Excel is available for PowerShell automation
    /// </summary>
    /// <returns>True if Excel is available, false otherwise</returns>
    public static bool IsExcelAvailable()
    {
        try
        {
            var testScript = @"
                try {
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                    Write-Output 'AVAILABLE'
                } catch {
                    Write-Output 'NOT_AVAILABLE'
                }";

            using var process = new Process();
            process.StartInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -Command -",
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            process.Start();
            process.StandardInput.WriteLine(testScript);
            process.StandardInput.Close();
            process.WaitForExit(10000); // 10 second timeout

            string output = process.StandardOutput.ReadToEnd();
            return output.Contains("AVAILABLE");
        }
        catch
        {
            return false;
        }
    }
}
