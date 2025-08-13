using JFToolkit.EncryptedExcel;
using NPOI.SS.UserModel;

namespace JFToolkit.EncryptedExcel.Examples;

/// <summary>
/// Examples showing how to modify and save Excel files with encryption
/// </summary>
public static class ModifyAndSaveExamples
{
    /// <summary>
    /// Example showing how to open, modify, and save an encrypted Excel file
    /// </summary>
    /// <param name="inputFilePath">Path to the input encrypted Excel file</param>
    /// <param name="inputPassword">Password for the input file</param>
    /// <param name="outputFilePath">Path where to save the modified file</param>
    /// <param name="outputPassword">Password for the output file (null for unencrypted)</param>
    public static void ModifyAndSaveExample(string inputFilePath, string inputPassword, string outputFilePath, string? outputPassword = null)
    {
        try
        {
            Console.WriteLine("📖 Opening encrypted Excel file...");
            
            // Open the encrypted file
            using var reader = EncryptedExcelReader.OpenFile(inputFilePath, inputPassword);
            var workbook = reader.Workbook!;
            
            Console.WriteLine($"✅ Opened file with {reader.NumberOfSheets} sheets");
            
            // Get the first sheet
            var sheet = workbook.GetSheetAt(0);
            
            // Modify existing data
            Console.WriteLine("✏️ Modifying existing data...");
            
            // Update John Doe's age and salary
            sheet.SetCellValue(1, 1, 31); // Age: 30 -> 31
            sheet.SetCellValue(1, 2, 77000); // Salary: 75000 -> 77000
            
            // Add a new employee
            Console.WriteLine("➕ Adding new employee...");
            sheet.AddRow("Mike Wilson", 28, 70000, DateTime.Now, true);
            
            // Modify Jane Smith's status
            sheet.SetCellValue(2, 4, false); // Set Jane as inactive
            
            // Add a comment/note
            sheet.SetCellValue(6, 0, "Modified on");
            sheet.SetCellValue(6, 1, DateTime.Now.ToString("yyyy-MM-dd"));
            
            Console.WriteLine("💾 Saving modified file...");
            
            // Save the file
            if (!string.IsNullOrEmpty(outputPassword))
            {
                // Save as encrypted
                workbook.SaveEncryptedToFile(outputFilePath, outputPassword);
                Console.WriteLine($"✅ Saved encrypted file to: {outputFilePath}");
            }
            else
            {
                // Save without encryption
                workbook.SaveToFile(outputFilePath);
                Console.WriteLine($"✅ Saved unencrypted file to: {outputFilePath}");
            }
            
            Console.WriteLine("✅ Modification completed successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Example showing how to create a new Excel file from scratch
    /// </summary>
    /// <param name="filePath">Path where to save the new file</param>
    /// <param name="password">Password to encrypt with (null for unencrypted)</param>
    public static void CreateNewFileExample(string filePath, string? password = null)
    {
        try
        {
            Console.WriteLine("📝 Creating new Excel file...");
            
            // Create a new workbook
            var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            
            // Create first sheet
            var employeeSheet = workbook.CreateSheet("Employees");
            
            // Add headers
            employeeSheet.SetRowValues(0, "Name", "Department", "Salary", "Start Date", "Active");
            
            // Add data
            employeeSheet.AddRow("Alice Johnson", "Engineering", 95000, new DateTime(2023, 1, 15), true);
            employeeSheet.AddRow("Bob Smith", "Marketing", 70000, new DateTime(2023, 2, 20), true);
            employeeSheet.AddRow("Carol Davis", "Finance", 85000, new DateTime(2022, 11, 10), false);
            employeeSheet.AddRow("David Wilson", "Engineering", 92000, new DateTime(2023, 3, 5), true);
            
            // Create second sheet with summary
            var summarySheet = workbook.CreateSheet("Summary");
            summarySheet.SetRowValues(0, "Metric", "Value");
            summarySheet.AddRow("Total Employees", 4);
            summarySheet.AddRow("Active Employees", 3);
            summarySheet.AddRow("Average Salary", 85500);
            summarySheet.AddRow("Report Date", DateTime.Now);
            
            Console.WriteLine("💾 Saving new file...");
            
            // Save the file
            if (!string.IsNullOrEmpty(password))
            {
                workbook.SaveEncryptedToFile(filePath, password);
                Console.WriteLine($"✅ Created encrypted file: {filePath}");
            }
            else
            {
                workbook.SaveToFile(filePath);
                Console.WriteLine($"✅ Created unencrypted file: {filePath}");
            }
            
            workbook.Close();
            Console.WriteLine("✅ File creation completed!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Example showing bulk data operations
    /// </summary>
    /// <param name="inputFilePath">Path to input file</param>
    /// <param name="inputPassword">Input file password</param>
    /// <param name="outputFilePath">Path to output file</param>
    /// <param name="outputPassword">Output file password</param>
    public static void BulkDataOperationsExample(string inputFilePath, string inputPassword, string outputFilePath, string? outputPassword = null)
    {
        try
        {
            Console.WriteLine("📊 Performing bulk data operations...");
            
            using var reader = EncryptedExcelReader.OpenFile(inputFilePath, inputPassword);
            var workbook = reader.Workbook!;
            var sheet = workbook.GetSheetAt(0);
            
            // Read all current data
            var allData = sheet.ToArray();
            Console.WriteLine($"📋 Read {allData.GetLength(0)} rows × {allData.GetLength(1)} columns");
            
            // Give everyone a 10% raise
            Console.WriteLine("💰 Giving everyone a 10% salary increase...");
            for (int row = 1; row <= sheet.LastRowNum; row++) // Skip header
            {
                var currentSalary = sheet.GetCell(row, 2)?.GetNumericValue() ?? 0;
                if (currentSalary > 0)
                {
                    var newSalary = currentSalary * 1.1;
                    sheet.SetCellValue(row, 2, Math.Round(newSalary, 2));
                }
            }
            
            // Add audit trail
            var auditSheet = workbook.CreateSheet("Audit Trail");
            auditSheet.SetRowValues(0, "Action", "Date", "Details");
            auditSheet.AddRow("Salary Increase", DateTime.Now, "Applied 10% increase to all employees");
            auditSheet.AddRow("File Modified", DateTime.Now, $"Original file: {Path.GetFileName(inputFilePath)}");
            
            // Save with audit trail
            if (!string.IsNullOrEmpty(outputPassword))
            {
                workbook.SaveEncryptedToFile(outputFilePath, outputPassword);
            }
            else
            {
                workbook.SaveToFile(outputFilePath);
            }
            
            Console.WriteLine($"✅ Bulk operations completed! Saved to: {outputFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Example showing how to work with byte arrays (useful for web applications)
    /// </summary>
    /// <param name="inputFilePath">Path to input file</param>
    /// <param name="inputPassword">Input password</param>
    /// <param name="outputPassword">Output password</param>
    /// <returns>Encrypted byte array</returns>
    public static byte[]? ByteArrayExample(string inputFilePath, string inputPassword, string outputPassword)
    {
        try
        {
            Console.WriteLine("🔄 Working with byte arrays...");
            
            // Read file as bytes
            var fileBytes = File.ReadAllBytes(inputFilePath);
            Console.WriteLine($"📁 Read {fileBytes.Length} bytes from file");
            
            // Open from byte array
            using var reader = EncryptedExcelReader.OpenBytes(fileBytes, inputPassword);
            var workbook = reader.Workbook!;
            
            // Make some modifications
            var sheet = workbook.GetSheetAt(0);
            sheet.SetCellValue(0, 5, "Last Modified");
            sheet.SetCellValue(1, 5, DateTime.Now);
            
            // Convert back to encrypted byte array
            var encryptedBytes = workbook.ToEncryptedByteArray(outputPassword);
            Console.WriteLine($"💾 Generated {encryptedBytes.Length} encrypted bytes");
            
            return encryptedBytes;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Error: {ex.Message}");
            return null;
        }
    }
}
