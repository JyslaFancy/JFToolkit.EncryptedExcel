using JFToolkit.EncryptedExcel;
using NPOI.SS.UserModel;

namespace RealWorldTest;

class Program
{
    static void Main()
    {
        Console.WriteLine("🔄 Real-World Test: Open → Modify → Save with Same Password");
        Console.WriteLine("=".PadRight(60, '='));

        string originalFile = @"C:\test\Encyption Test sheet 1.xlsx";
        string modifiedFile = @"C:\test\ProcessedWithSamePassword.xlsx";
        string password = "TestPassword123";

        try
        {
            // Step 1: Open encrypted file
            Console.WriteLine($"📂 Opening encrypted file: {Path.GetFileName(originalFile)}");
            using var reader = EncryptedExcelReader.OpenFile(originalFile, password);
            var workbook = reader.Workbook!;
            var sheet = workbook.GetSheetAt(0);
            
            Console.WriteLine($"✅ File opened successfully");
            Console.WriteLine($"   Rows: {sheet.LastRowNum + 1}");
            
            // Step 2: Show current data
            Console.WriteLine("\n📊 Current Data:");
            for (int i = 0; i <= Math.Min(sheet.LastRowNum, 4); i++)
            {
                var row = sheet.GetRow(i);
                if (row != null)
                {
                    string name = row.GetCell(0)?.ToString() ?? "";
                    string age = row.GetCell(1)?.ToString() ?? "";
                    string salary = row.GetCell(2)?.ToString() ?? "";
                    Console.WriteLine($"   Row {i}: {name} | {age} | {salary}");
                }
            }
            
            // Step 3: Modify data (business logic)
            Console.WriteLine("\n✏️ Applying Business Logic:");
            
            // Example: Give everyone a 5% salary increase
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row?.GetCell(2) != null)
                {
                    var salaryCell = row.GetCell(2);
                    if (double.TryParse(salaryCell.ToString(), out double currentSalary))
                    {
                        double newSalary = currentSalary * 1.05; // 5% increase
                        salaryCell.SetCellValue(newSalary);
                        Console.WriteLine($"   Updated Row {i}: Salary {currentSalary:F0} → {newSalary:F0}");
                    }
                }
            }
            
            // Add processing timestamp
            var lastRow = sheet.LastRowNum + 1;
            var timestampRow = sheet.CreateRow(lastRow);
            timestampRow.CreateCell(0).SetCellValue($"Processed on: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            Console.WriteLine($"   Added timestamp at row {lastRow}");
            
            // Step 4: Save with same password (this is where we need the solution)
            Console.WriteLine("\n💾 Saving with same password...");
            
            // Try the enhanced encrypted save
            bool success = SaveWithEncryption(workbook, modifiedFile, password);
            
            if (success)
            {
                Console.WriteLine($"✅ SUCCESS: Saved encrypted file: {Path.GetFileName(modifiedFile)}");
                
                // Step 5: Verify by reopening
                Console.WriteLine("\n🔍 Verification: Reopening saved file...");
                using var verifyReader = EncryptedExcelReader.OpenFile(modifiedFile, password);
                var verifySheet = verifyReader.Workbook!.GetSheetAt(0);
                Console.WriteLine($"✅ Verification successful - {verifySheet.LastRowNum + 1} rows found");
                
                Console.WriteLine("\n🎉 REAL-WORLD TEST COMPLETE!");
                Console.WriteLine("   Your application can now:");
                Console.WriteLine("   • Open encrypted Excel files");
                Console.WriteLine("   • Apply business logic/modifications"); 
                Console.WriteLine("   • Save back with same password protection");
                Console.WriteLine("   • Maintain full encryption in your workflow");
            }
            else
            {
                Console.WriteLine("❌ Encrypted save failed - using fallback approach");
                ShowFallbackOptions(workbook, modifiedFile, password);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error: {ex.Message}");
        }
    }
    
    static bool SaveWithEncryption(IWorkbook workbook, string filePath, string password)
    {
        try
        {
            // Try multiple approaches for encrypted saving
            return TryExcelAutomation(workbook, filePath, password) ||
                   TryPowerShellApproach(workbook, filePath, password) ||
                   TryNativeNPOI(workbook, filePath, password);
        }
        catch
        {
            return false;
        }
    }
    
    static bool TryExcelAutomation(IWorkbook workbook, string filePath, string password)
    {
        try
        {
            Console.WriteLine("   Trying Excel Automation...");
            return ExcelAutomationHelper.EncryptExcelFile(
                SaveTemporaryFile(workbook), filePath, password);
        }
        catch
        {
            return false;
        }
    }
    
    static bool TryPowerShellApproach(IWorkbook workbook, string filePath, string password)
    {
        try
        {
            Console.WriteLine("   Trying PowerShell approach...");
            return PowerShellExcelHelper.EncryptExcelFile(
                SaveTemporaryFile(workbook), filePath, password);
        }
        catch
        {
            return false;
        }
    }
    
    static bool TryNativeNPOI(IWorkbook workbook, string filePath, string password)
    {
        try
        {
            Console.WriteLine("   Trying native NPOI...");
            EncryptedExcelWriter.SaveEncryptedToFile(workbook, filePath, password);
            return true;
        }
        catch
        {
            return false;
        }
    }
    
    static string SaveTemporaryFile(IWorkbook workbook)
    {
        var tempFile = Path.GetTempFileName();
        var excelFile = Path.ChangeExtension(tempFile, ".xlsx");
        workbook.SaveToFile(excelFile);
        return excelFile;
    }
    
    static void ShowFallbackOptions(IWorkbook workbook, string filePath, string password)
    {
        Console.WriteLine("\n📋 FALLBACK OPTIONS for your application:");
        Console.WriteLine("\n1. 🤖 Automated Solution (Recommended):");
        Console.WriteLine("   • Save as unencrypted");
        Console.WriteLine("   • Use PowerShell to encrypt with same password");
        Console.WriteLine("   • Delete unencrypted temporary file");
        
        Console.WriteLine("\n2. 🔄 Two-Step Process:");
        Console.WriteLine("   • Save unencrypted version"); 
        Console.WriteLine("   • Notify user to manually encrypt");
        Console.WriteLine("   • Or queue for batch encryption");
        
        Console.WriteLine("\n3. 🛠️ Custom Integration:");
        Console.WriteLine("   • Integrate with Excel COM objects");
        Console.WriteLine("   • Use third-party encryption libraries");
        Console.WriteLine("   • Implement file-level encryption");
        
        // Save unencrypted as fallback
        workbook.SaveToFile(filePath.Replace(".xlsx", "_unencrypted.xlsx"));
        Console.WriteLine($"\n💾 Saved unencrypted version: {Path.GetFileName(filePath.Replace(".xlsx", "_unencrypted.xlsx"))}");
    }
}
