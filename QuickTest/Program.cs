using JFToolkit.EncryptedExcel;

class Program
{
    static void Main()
    {
        try
        {
            Console.WriteLine("� Testing Excel file modification and saving...");
            
            // Test 1: Open, modify, and save unencrypted
            Console.WriteLine("\n📝 Test 1: Modify and save as unencrypted");
            using (var reader = EncryptedExcelReader.OpenFile(@"C:\test\Encyption Test sheet 1.xlsx", "TestPassword123"))
            {
                var workbook = reader.Workbook!;
                var sheet = workbook.GetSheetAt(0);
                
                Console.WriteLine("� Original data:");
                Console.WriteLine($"   John Doe: Age={sheet.GetCellValue(1, 1)}, Salary={sheet.GetCellValue(1, 2)}");
                
                // Modify John's data
                sheet.SetCellValue(1, 1, 31); // Age 30 -> 31
                sheet.SetCellValue(1, 2, 77000); // Salary 75000 -> 77000
                
                // Add a new employee
                sheet.AddRow("New Employee", 29, 68000, DateTime.Now, true);
                
                Console.WriteLine("✏️ Modified data:");
                Console.WriteLine($"   John Doe: Age={sheet.GetCellValue(1, 1)}, Salary={sheet.GetCellValue(1, 2)}");
                Console.WriteLine($"   New Employee: {sheet.GetCellValue(5, 0)}, Age={sheet.GetCellValue(5, 1)}");
                
                // Save as unencrypted
                workbook.SaveToFile(@"C:\test\Modified_Unencrypted.xlsx");
                Console.WriteLine("✅ Saved unencrypted file: C:\\test\\Modified_Unencrypted.xlsx");
            }
            
            // Test 2: Save as encrypted with new password
            Console.WriteLine("\n� Test 2: Save as encrypted with new password");
            using (var reader = EncryptedExcelReader.OpenFile(@"C:\test\Encyption Test sheet 1.xlsx", "TestPassword123"))
            {
                var workbook = reader.Workbook!;
                var sheet = workbook.GetSheetAt(0);
                
                // Add modification timestamp
                sheet.SetCellValue(6, 0, "Modified on:");
                sheet.SetCellValue(6, 1, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                
                // Save as encrypted with new password
                workbook.SaveEncryptedToFile(@"C:\test\Modified_Encrypted.xlsx", "NewPassword456");
                Console.WriteLine("✅ Saved encrypted file: C:\\test\\Modified_Encrypted.xlsx");
                Console.WriteLine("   New password: NewPassword456");
            }
            
            // Test 3: Verify the encrypted file can be opened
            Console.WriteLine("\n🔍 Test 3: Verify encrypted file can be opened");
            using (var reader = EncryptedExcelReader.OpenFile(@"C:\test\Modified_Encrypted.xlsx", "NewPassword456"))
            {
                var sheet = reader.GetSheetAt(0);
                Console.WriteLine($"✅ Successfully opened encrypted file!");
                Console.WriteLine($"   Modification timestamp: {sheet.GetCellValue(6, 1)}");
                Console.WriteLine($"   Number of rows: {sheet.LastRowNum + 1}");
            }
            
            // Test 4: Work with byte arrays
            Console.WriteLine("\n💾 Test 4: Byte array operations");
            var fileBytes = File.ReadAllBytes(@"C:\test\Encyption Test sheet 1.xlsx");
            using (var reader = EncryptedExcelReader.OpenBytes(fileBytes, "TestPassword123"))
            {
                var workbook = reader.Workbook!;
                var sheet = workbook.GetSheetAt(0);
                
                // Add a note
                sheet.SetCellValue(7, 0, "Processed via byte array");
                
                // Convert to encrypted byte array
                var encryptedBytes = workbook.ToEncryptedByteArray("ByteArrayPassword789");
                
                // Save the byte array to file
                File.WriteAllBytes(@"C:\test\FromByteArray.xlsx", encryptedBytes);
                Console.WriteLine($"✅ Created file from byte array: {encryptedBytes.Length} bytes");
            }
            
            // Test 5: Verify byte array file
            Console.WriteLine("\n✅ Test 5: Verify byte array file");
            using (var reader = EncryptedExcelReader.OpenFile(@"C:\test\FromByteArray.xlsx", "ByteArrayPassword789"))
            {
                var sheet = reader.GetSheetAt(0);
                Console.WriteLine($"✅ Byte array file opened successfully!");
                Console.WriteLine($"   Note: {sheet.GetCellValue(7, 0)}");
            }
            
            Console.WriteLine("\n🎉 All tests completed successfully!");
            Console.WriteLine("\nFiles created:");
            Console.WriteLine("   📄 C:\\test\\Modified_Unencrypted.xlsx (no password)");
            Console.WriteLine("   🔐 C:\\test\\Modified_Encrypted.xlsx (password: NewPassword456)");
            Console.WriteLine("   💾 C:\\test\\FromByteArray.xlsx (password: ByteArrayPassword789)");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
    }
}
