using JFToolkit.EncryptedExcel;

class SimpleTest
{
    static void Main()
    {
        try
        {
            Console.WriteLine("🔧 Testing Excel file modification and saving...");
            
            // Test: Open, modify, and save as unencrypted
            Console.WriteLine("\n📝 Opening encrypted file, modifying data, and saving as unencrypted");
            
            using var reader = EncryptedExcelReader.OpenFile(@"C:\test\Encyption Test sheet 1.xlsx", "TestPassword123");
            var workbook = reader.Workbook!;
            var sheet = workbook.GetSheetAt(0);
            
            Console.WriteLine("📊 Original data:");
            Console.WriteLine($"   John Doe: Age={sheet.GetCellValue(1, 1)}, Salary={sheet.GetCellValue(1, 2)}");
            
            // Modify John's data
            sheet.SetCellValue(1, 1, 31); // Age 30 -> 31
            sheet.SetCellValue(1, 2, 77000); // Salary 75000 -> 77000
            
            // Add a new employee
            sheet.AddRow("Mike Johnson", 35, 85000, new DateTime(2024, 8, 13), true);
            
            // Add modification timestamp
            sheet.SetCellValue(6, 0, "Modified on:");
            sheet.SetCellValue(6, 1, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            
            Console.WriteLine("✏️ After modifications:");
            Console.WriteLine($"   John Doe: Age={sheet.GetCellValue(1, 1)}, Salary={sheet.GetCellValue(1, 2)}");
            Console.WriteLine($"   New Employee: {sheet.GetCellValue(5, 0)}, Age={sheet.GetCellValue(5, 1)}, Salary={sheet.GetCellValue(5, 2)}");
            Console.WriteLine($"   Modification time: {sheet.GetCellValue(6, 1)}");
            Console.WriteLine($"   Total rows now: {sheet.LastRowNum + 1}");
            
            // Save as unencrypted
            workbook.SaveToFile(@"C:\test\Modified_File.xlsx");
            Console.WriteLine("✅ Saved modified file as unencrypted: C:\\test\\Modified_File.xlsx");
            
            // Verify by opening the saved file
            Console.WriteLine("\n🔍 Verifying saved file...");
            var verifyWorkbook = new NPOI.XSSF.UserModel.XSSFWorkbook(@"C:\test\Modified_File.xlsx");
            var verifySheet = verifyWorkbook.GetSheetAt(0);
            
            Console.WriteLine($"✅ Successfully opened saved file!");
            Console.WriteLine($"   Verified John's new age: {verifySheet.GetCellValue(1, 1)}");
            Console.WriteLine($"   Verified John's new salary: {verifySheet.GetCellValue(1, 2)}");
            Console.WriteLine($"   Verified new employee: {verifySheet.GetCellValue(5, 0)}");
            Console.WriteLine($"   Verified total rows: {verifySheet.LastRowNum + 1}");
            
            verifyWorkbook.Close();
            
            Console.WriteLine("\n🎉 Test completed successfully!");
            Console.WriteLine("\n💡 Key achievements:");
            Console.WriteLine("   ✅ Opened password-encrypted Excel file");
            Console.WriteLine("   ✅ Modified existing cell values");
            Console.WriteLine("   ✅ Added new rows with data");
            Console.WriteLine("   ✅ Saved as unencrypted Excel file");
            Console.WriteLine("   ✅ Verified the saved file opens correctly");
            
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
    }
}
