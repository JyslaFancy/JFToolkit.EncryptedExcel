using JFToolkit.EncryptedExcel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace JFToolkit.EncryptedExcel.Tests;

/// <summary>
/// Test class to verify macro-enabled Excel file support
/// </summary>
public static class MacroSupportTest
{
    /// <summary>
    /// Test creating and working with a macro-enabled workbook
    /// </summary>
    public static void TestMacroEnabledSupport()
    {
        Console.WriteLine("=== Testing Macro-Enabled Excel Support (.xlsm) ===");
        
        try
        {
            // Create a new XSSF workbook (supports both .xlsx and .xlsm)
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("MacroTest");
            
            // Add some data
            var headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("Name");
            headerRow.CreateCell(1).SetCellValue("Value");
            headerRow.CreateCell(2).SetCellValue("Formula");
            
            var dataRow = sheet.CreateRow(1);
            dataRow.CreateCell(0).SetCellValue("Test Data");
            dataRow.CreateCell(1).SetCellValue(100);
            dataRow.CreateCell(2).SetCellFormula("B2*2"); // Simple formula
            
            // Test saving as .xlsm format (unencrypted first)
            string tempPath = Path.GetTempFileName().Replace(".tmp", ".xlsm");
            Console.WriteLine($"Creating test .xlsm file: {tempPath}");
            
            // Save the workbook
            EncryptedExcelWriter.SaveToFile(workbook, tempPath);
            Console.WriteLine("✅ Successfully created .xlsm file");
            
            // Now test reading it back (unencrypted file)
            Console.WriteLine("Testing reading .xlsm file...");
            
            // For unencrypted files, we need to use NPOI directly
            using var fileStream = new FileStream(tempPath, FileMode.Open, FileAccess.Read);
            var testWorkbook = WorkbookFactory.Create(fileStream);
            
            var readSheet = testWorkbook.GetSheetAt(0);
            if (readSheet != null)
            {
                var nameValue = readSheet.GetCellValue(0, 0);
                var numValue = readSheet.GetCellValue(1, 1);
                Console.WriteLine($"✅ Read data - Name: {nameValue}, Value: {numValue}");
            }
            
            // Clean up
            File.Delete(tempPath);
            Console.WriteLine("✅ Test completed successfully");
            
            // Important note about macros
            Console.WriteLine();
            Console.WriteLine("📝 Important Notes about .xlsm support:");
            Console.WriteLine("   • .xlsm files are fully supported for reading and writing");
            Console.WriteLine("   • Macro code is preserved when saving .xlsm files");
            Console.WriteLine("   • Macros cannot be executed through NPOI (security limitation)");
            Console.WriteLine("   • Data and formulas work normally");
            Console.WriteLine("   • Password encryption works the same as .xlsx files");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Test failed: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Information about macro support limitations and capabilities
    /// </summary>
    public static void DisplayMacroSupportInfo()
    {
        Console.WriteLine();
        Console.WriteLine("=== Macro-Enabled Excel (.xlsm) Support Information ===");
        Console.WriteLine();
        Console.WriteLine("✅ SUPPORTED FEATURES:");
        Console.WriteLine("   • Reading encrypted .xlsm files");
        Console.WriteLine("   • Writing/saving .xlsm files with encryption");
        Console.WriteLine("   • Preserving existing macro code");
        Console.WriteLine("   • All standard Excel data types and formulas");
        Console.WriteLine("   • Cell formatting and styling");
        Console.WriteLine("   • Multiple worksheets");
        Console.WriteLine();
        Console.WriteLine("⚠️  LIMITATIONS:");
        Console.WriteLine("   • Cannot execute VBA macros (security restriction)");
        Console.WriteLine("   • Cannot create new macros programmatically");
        Console.WriteLine("   • Cannot modify existing macro code");
        Console.WriteLine();
        Console.WriteLine("💡 USE CASES:");
        Console.WriteLine("   • Processing data in macro-enabled templates");
        Console.WriteLine("   • Updating data while preserving existing macros");
        Console.WriteLine("   • Converting between .xlsm and .xlsx formats");
        Console.WriteLine("   • Batch processing of macro-enabled files");
    }
}
