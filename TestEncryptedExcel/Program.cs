using JFToolkit.EncryptedExcel;

namespace TestEncryptedExcel;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("JFToolkit.EncryptedExcel Demo");
        Console.WriteLine("============================");
        
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: TestEncryptedExcel <excel-file-path> <password>");
            Console.WriteLine();
            Console.WriteLine("This demo shows how to use the JFToolkit.EncryptedExcel library");
            Console.WriteLine("to open and read password-protected Excel files.");
            Console.WriteLine();
            Console.WriteLine("Example:");
            Console.WriteLine("  TestEncryptedExcel \"C:\\path\\to\\encrypted.xlsx\" \"mypassword\"");
            return;
        }
        
        string filePath = args[0];
        string password = args[1];
        
        try
        {
            // Demo the library functionality
            DemoLibrary(filePath, password);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        
        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }
    
    static void DemoLibrary(string filePath, string password)
    {
        Console.WriteLine($"Opening file: {filePath}");
        Console.WriteLine($"Using password: {new string('*', password.Length)}");
        Console.WriteLine();
        
        // Open the encrypted Excel file
        using var reader = EncryptedExcelReader.OpenFile(filePath, password);
        
        Console.WriteLine($"✅ Successfully opened file!");
        Console.WriteLine($"📊 Number of sheets: {reader.NumberOfSheets}");
        Console.WriteLine();
        
        // List all sheet names
        var sheetNames = reader.GetSheetNames();
        Console.WriteLine("📋 Sheet names:");
        for (int i = 0; i < sheetNames.Length; i++)
        {
            Console.WriteLine($"  {i + 1}. {sheetNames[i]}");
        }
        Console.WriteLine();
        
        // Work with the first sheet
        if (reader.NumberOfSheets > 0)
        {
            var firstSheet = reader.GetSheetAt(0);
            var sheetName = reader.GetSheetName(0);
            
            Console.WriteLine($"📄 Reading from sheet: '{sheetName}'");
            Console.WriteLine($"📏 Rows: {firstSheet.LastRowNum + 1}");
            
            // Find max columns
            int maxColumns = 0;
            for (int row = 0; row <= Math.Min(firstSheet.LastRowNum, 10); row++)
            {
                var currentRow = firstSheet.GetRow(row);
                if (currentRow != null && currentRow.LastCellNum > maxColumns)
                {
                    maxColumns = currentRow.LastCellNum;
                }
            }
            Console.WriteLine($"📏 Columns: {maxColumns}");
            Console.WriteLine();
            
            // Show sample data (first 5x5 cells)
            Console.WriteLine("📊 Sample data (first 5x5 cells):");
            Console.WriteLine(new string('-', 80));
            
            for (int row = 0; row <= Math.Min(firstSheet.LastRowNum, 4); row++)
            {
                var rowData = new List<string>();
                for (int col = 0; col < Math.Min(maxColumns, 5); col++)
                {
                    var cellValue = firstSheet.GetCellValue(row, col);
                    // Truncate long values
                    if (cellValue.Length > 15)
                        cellValue = cellValue.Substring(0, 12) + "...";
                    rowData.Add(cellValue.PadRight(15));
                }
                Console.WriteLine($"Row {row + 1}: {string.Join(" | ", rowData)}");
            }
            
            Console.WriteLine(new string('-', 80));
            Console.WriteLine();
            
            // Demo different cell types
            Console.WriteLine("🔍 Cell type analysis (first 10 cells with data):");
            int cellCount = 0;
            for (int row = 0; row <= firstSheet.LastRowNum && cellCount < 10; row++)
            {
                for (int col = 0; col < maxColumns && cellCount < 10; col++)
                {
                    var cell = firstSheet.GetCell(row, col);
                    if (cell != null && !string.IsNullOrWhiteSpace(cell.GetStringValue()))
                    {
                        var stringValue = cell.GetStringValue();
                        var numericValue = cell.GetNumericValue();
                        var dateValue = cell.GetDateTimeValue();
                        var boolValue = cell.GetBooleanValue();
                        
                        Console.WriteLine($"  Cell [{row + 1},{col + 1}] ({cell.CellType}):");
                        Console.WriteLine($"    String: \"{stringValue}\"");
                        if (numericValue != 0)
                            Console.WriteLine($"    Numeric: {numericValue}");
                        if (dateValue != DateTime.MinValue)
                            Console.WriteLine($"    Date: {dateValue:yyyy-MM-dd}");
                        Console.WriteLine($"    Boolean: {boolValue}");
                        Console.WriteLine();
                        
                        cellCount++;
                    }
                }
            }
            
            // Demo row operations
            if (firstSheet.LastRowNum >= 0)
            {
                Console.WriteLine("📊 First row values:");
                var firstRow = firstSheet.GetRow(0);
                if (firstRow != null)
                {
                    var rowValues = firstRow.GetRowValues();
                    for (int i = 0; i < Math.Min(rowValues.Length, 10); i++)
                    {
                        if (!string.IsNullOrWhiteSpace(rowValues[i]))
                        {
                            Console.WriteLine($"  Column {i + 1}: \"{rowValues[i]}\"");
                        }
                    }
                }
                Console.WriteLine();
            }
            
            // Demo column operations
            if (maxColumns > 0)
            {
                Console.WriteLine("📊 First column values (first 10):");
                var columnValues = firstSheet.GetColumnValues(0);
                for (int i = 0; i < Math.Min(columnValues.Length, 10); i++)
                {
                    if (!string.IsNullOrWhiteSpace(columnValues[i]))
                    {
                        Console.WriteLine($"  Row {i + 1}: \"{columnValues[i]}\"");
                    }
                }
            }
        }
        
        Console.WriteLine("\n✅ Demo completed successfully!");
    }
}
