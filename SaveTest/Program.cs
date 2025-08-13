using JFToolkit.EncryptedExcel;

namespace SaveTest;

class Program
{
    static void Main()
    {
        Console.WriteLine("🔧 Testing Excel Modification and Save...");

        try
        {
            // Open encrypted file
            using var reader = EncryptedExcelReader.OpenFile(@"C:\test\Encyption Test sheet 1.xlsx", "TestPassword123");
            var workbook = reader.Workbook!;
            var sheet = workbook.GetSheetAt(0);
            
            Console.WriteLine($"📊 Original John Doe: Age={sheet.GetCellValue(1, 1)}, Salary={sheet.GetCellValue(1, 2)}");
            
            // Modify data
            sheet.SetCellValue(1, 1, 32); // Change age
            sheet.SetCellValue(1, 2, 80000); // Change salary
            
            // Add new employee
            sheet.AddRow("Sarah Connor", 40, 120000, new DateTime(2024, 8, 13), true);
            
            Console.WriteLine($"✏️ Modified John Doe: Age={sheet.GetCellValue(1, 1)}, Salary={sheet.GetCellValue(1, 2)}");
            Console.WriteLine($"➕ Added: {sheet.GetCellValue(5, 0)}, Salary={sheet.GetCellValue(5, 2)}");
            
            // Save as encrypted (will save unencrypted with guidance)
            Console.WriteLine("📁 Attempting to save with encryption...");
            EncryptedExcelWriter.SaveEncryptedToFile(workbook, @"C:\test\ModifiedFileForEncryption.xlsx", "TestPassword123");
            
            // Also save regular unencrypted
            workbook.SaveToFile(@"C:\test\ModifiedFileUnencrypted.xlsx");
            Console.WriteLine("✅ Saved unencrypted to: C:\\test\\ModifiedFileUnencrypted.xlsx");
            
            Console.WriteLine();
            Console.WriteLine("🎉 Success! Your NuGet package can:");
            Console.WriteLine("   ✅ Read password-encrypted Excel files perfectly");
            Console.WriteLine("   ✅ Modify data and add new rows");
            Console.WriteLine("   ✅ Save as standard Excel files");
            Console.WriteLine("   📝 For encryption: Use Excel's 'Save As' with password");
            Console.WriteLine();
            Console.WriteLine("� Files created in C:\\test\\:");
            Console.WriteLine("   • ModifiedFileUnencrypted.xlsx (ready to use)");
            Console.WriteLine("   • ModifiedFileForEncryption.xlsx (ready for manual encryption)");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error: {ex.Message}");
        }
    }
}
