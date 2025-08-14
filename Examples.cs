using JFToolkit.EncryptedExcel;
using NPOI.SS.UserModel;
using System;
using System.IO;

namespace JFToolkit.EncryptedExcel.Examples;

/// <summary>
/// Example usage of the EncryptedExcelReader library
/// </summary>
public static class Examples
{
    /// <summary>
    /// Basic example of opening and reading an encrypted Excel file
    /// </summary>
    /// <param name="filePath">Path to the encrypted Excel file</param>
    /// <param name="password">Password for the file</param>
    public static void BasicExample(string filePath, string password)
    {
        try
        {
            // Open the encrypted Excel file
            using var reader = EncryptedExcelReader.OpenFile(filePath, password);
            
            Console.WriteLine($"Successfully opened file with {reader.NumberOfSheets} sheet(s)");
            
            // List all sheet names
            var sheetNames = reader.GetSheetNames();
            Console.WriteLine("Sheets:");
            foreach (var sheetName in sheetNames)
            {
                Console.WriteLine($"  - {sheetName}");
            }
            
            // Work with the first sheet
            var firstSheet = reader.GetSheetAt(0);
            Console.WriteLine($"\nReading from sheet: {reader.GetSheetName(0)}");
            
            // Read some cells
            Console.WriteLine($"Cell A1: {firstSheet.GetCellValue(0, 0)}");
            Console.WriteLine($"Cell B1: {firstSheet.GetCellValue(0, 1)}");
            Console.WriteLine($"Cell A2: {firstSheet.GetCellValue(1, 0)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Example of reading different cell types with type safety
    /// </summary>
    /// <param name="filePath">Path to the encrypted Excel file</param>
    /// <param name="password">Password for the file</param>
    public static void TypedReadingExample(string filePath, string password)
    {
        using var reader = EncryptedExcelReader.OpenFile(filePath, password);
        var sheet = reader.GetSheetAt(0);
        
        // Read different types of data
        for (int row = 0; row <= Math.Min(sheet.LastRowNum, 10); row++)
        {
            for (int col = 0; col < 5; col++)
            {
                var cell = sheet.GetCell(row, col);
                if (cell != null)
                {
                    Console.WriteLine($"Cell [{row},{col}]:");
                    Console.WriteLine($"  String: {cell.GetStringValue()}");
                    Console.WriteLine($"  Numeric: {cell.GetNumericValue()}");
                    Console.WriteLine($"  Date: {cell.GetDateTimeValue()}");
                    Console.WriteLine($"  Boolean: {cell.GetBooleanValue()}");
                    Console.WriteLine($"  Type: {cell.CellType}");
                    Console.WriteLine();
                }
            }
        }
    }
    
    /// <summary>
    /// Example of reading entire rows and columns
    /// </summary>
    /// <param name="filePath">Path to the encrypted Excel file</param>
    /// <param name="password">Password for the file</param>
    public static void BulkReadingExample(string filePath, string password)
    {
        using var reader = EncryptedExcelReader.OpenFile(filePath, password);
        var sheet = reader.GetSheetAt(0);
        
        // Read the header row
        var headerRow = sheet.GetRow(0);
        if (headerRow != null)
        {
            var headers = headerRow.GetRowValues();
            Console.WriteLine("Headers: " + string.Join(", ", headers));
        }
        
        // Read the first column (usually IDs or names)
        var firstColumnValues = sheet.GetColumnValues(0, startRow: 1); // Skip header
        Console.WriteLine("\nFirst column values:");
        foreach (var value in firstColumnValues.Take(10)) // Show first 10
        {
            Console.WriteLine($"  {value}");
        }
        
        // Convert entire sheet to array (useful for small sheets)
        if (sheet.LastRowNum < 100) // Only for small sheets
        {
            var allData = sheet.ToArray();
            Console.WriteLine($"\nSheet converted to array: {allData.GetLength(0)} rows x {allData.GetLength(1)} columns");
        }
    }
    
    /// <summary>
    /// Example of working with multiple sheets
    /// </summary>
    /// <param name="filePath">Path to the encrypted Excel file</param>
    /// <param name="password">Password for the file</param>
    public static void MultipleSheetsExample(string filePath, string password)
    {
        using var reader = EncryptedExcelReader.OpenFile(filePath, password);
        
        // Process each sheet
        for (int i = 0; i < reader.NumberOfSheets; i++)
        {
            var sheet = reader.GetSheetAt(i);
            var sheetName = reader.GetSheetName(i);
            
            Console.WriteLine($"\nProcessing sheet: {sheetName}");
            Console.WriteLine($"  Rows: {sheet.LastRowNum + 1}");
            
            // Find the maximum number of columns
            int maxColumns = 0;
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                var currentRow = sheet.GetRow(row);
                if (currentRow != null && currentRow.LastCellNum > maxColumns)
                {
                    maxColumns = currentRow.LastCellNum;
                }
            }
            Console.WriteLine($"  Columns: {maxColumns}");
            
            // Show some sample data
            if (sheet.LastRowNum >= 0)
            {
                var firstRow = sheet.GetRow(0);
                if (firstRow != null)
                {
                    var rowData = firstRow.GetRowValues();
                    Console.WriteLine($"  First row: {string.Join(" | ", rowData.Take(5))}");
                }
            }
        }
    }
    
    /// <summary>
    /// Example of working with macro-enabled Excel files (.xlsm)
    /// </summary>
    /// <param name="xlsmFilePath">Path to the encrypted .xlsm file</param>
    /// <param name="password">Password for the file</param>
    public static void MacroEnabledExample(string xlsmFilePath, string password)
    {
        try
        {
            // Open the encrypted macro-enabled Excel file
            using var reader = EncryptedExcelReader.OpenFile(xlsmFilePath, password);
            
            Console.WriteLine($"Successfully opened macro-enabled file: {Path.GetFileName(xlsmFilePath)}");
            Console.WriteLine($"Number of sheets: {reader.NumberOfSheets}");
            
            // Read data just like any other Excel file
            var firstSheet = reader.GetFirstSheet();
            if (firstSheet != null)
            {
                Console.WriteLine($"First sheet name: {firstSheet.SheetName}");
                
                // Read some data
                var cellValue = firstSheet.GetCellValue(0, 0);
                Console.WriteLine($"Cell A1 value: {cellValue}");
                
                // Note: Macros are preserved but cannot be executed through NPOI
                // The file structure and data remain intact
            }
            
            // Save as macro-enabled file (preserves macros)
            var outputPath = xlsmFilePath.Replace(".xlsm", "_modified.xlsm");
            EncryptedExcelWriter.SaveEncryptedToFile(reader.Workbook!, outputPath, password);
            
            Console.WriteLine($"Saved macro-enabled file to: {outputPath}");
            Console.WriteLine("Note: Macros are preserved in the saved file");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing macro-enabled file: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Example of reading from a stream (useful for web applications)
    /// </summary>
    /// <param name="fileBytes">Excel file as byte array</param>
    /// <param name="password">Password for the file</param>
    public static void StreamExample(byte[] fileBytes, string password)
    {
        // Method 1: From byte array
        using var reader1 = EncryptedExcelReader.OpenBytes(fileBytes, password);
        Console.WriteLine($"Opened from bytes: {reader1.NumberOfSheets} sheets");
        
        // Method 2: From stream
        using var stream = new MemoryStream(fileBytes);
        using var reader2 = EncryptedExcelReader.OpenStream(stream, password);
        Console.WriteLine($"Opened from stream: {reader2.NumberOfSheets} sheets");
    }
}
