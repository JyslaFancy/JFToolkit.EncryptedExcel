# JFToolkit.EncryptedExcel

A simple and easy-to-use .NET library for opening and reading password-encrypted Excel files using the free NPOI library.

## Features

- ✅ Open password-encrypted Excel files (.xlsx and .xls)
- ✅ Read from file paths, streams, or byte arrays
- ✅ Easy-to-use extension methods for working with cells
- ✅ Handle different cell types (string, numeric, date, boolean)
- ✅ Convert sheets to arrays for easy data processing
- ✅ Free and open-source (uses NPOI, no license required)

## Installation

Install via NuGet Package Manager:

```
Install-Package JFToolkit.EncryptedExcel
```

Or via .NET CLI:

```
dotnet add package JFToolkit.EncryptedExcel
```

## Quick Start

### Basic Usage

```csharp
using JFToolkit.EncryptedExcel;

// Open an encrypted Excel file
using var reader = EncryptedExcelReader.OpenFile("encrypted_file.xlsx", "password123");

// Get the first sheet
var sheet = reader.GetSheetAt(0);

// Read a specific cell
var cellValue = sheet.GetCellValue(0, 0); // Row 0, Column 0

// Get all sheet names
var sheetNames = reader.GetSheetNames();
```

### Working with Cells

```csharp
using var reader = EncryptedExcelReader.OpenFile("data.xlsx", "mypassword");
var sheet = reader.GetSheetAt(0);

// Get cell values with type safety
var cell = sheet.GetCell(1, 2);
var stringValue = cell.GetStringValue();
var numericValue = cell.GetNumericValue();
var dateValue = cell.GetDateTimeValue();
var boolValue = cell.GetBooleanValue();
```

### Reading Entire Rows or Columns

```csharp
using var reader = EncryptedExcelReader.OpenFile("data.xlsx", "mypassword");
var sheet = reader.GetSheetAt(0);

// Get all values from a row
var row = sheet.GetRow(0);
var rowValues = row.GetRowValues();

// Get all values from a column
var columnValues = sheet.GetColumnValues(0); // Column A

// Convert entire sheet to 2D array
var allData = sheet.ToArray();
```

### Opening from Different Sources

```csharp
// From file path
using var reader1 = EncryptedExcelReader.OpenFile("path/to/file.xlsx", "password");

// From stream
using var fileStream = new FileStream("file.xlsx", FileMode.Open);
using var reader2 = EncryptedExcelReader.OpenStream(fileStream, "password");

// From byte array
byte[] excelData = GetExcelDataFromSomewhere();
using var reader3 = EncryptedExcelReader.OpenBytes(excelData, "password");
```

## Error Handling

The library throws meaningful exceptions for common scenarios:

```csharp
try
{
    using var reader = EncryptedExcelReader.OpenFile("file.xlsx", "wrongpassword");
}
catch (FileNotFoundException)
{
    Console.WriteLine("File not found");
}
catch (InvalidOperationException ex)
{
    Console.WriteLine($"Could not open file: {ex.Message}");
    // Usually means wrong password or corrupted file
}
```

## API Reference

### EncryptedExcelReader

- `OpenFile(string filePath, string password)` - Open from file path
- `OpenStream(Stream stream, string password)` - Open from stream
- `OpenBytes(byte[] data, string password)` - Open from byte array
- `GetSheetAt(int index)` - Get sheet by index
- `GetSheet(string name)` - Get sheet by name
- `GetSheetNames()` - Get all sheet names
- `NumberOfSheets` - Total number of sheets

### Extension Methods

- `GetStringValue()` - Get cell value as string
- `GetNumericValue()` - Get cell value as double
- `GetDateTimeValue()` - Get cell value as DateTime
- `GetBooleanValue()` - Get cell value as boolean
- `GetCellValue(row, col)` - Get cell value at coordinates
- `GetRowValues()` - Get all values from a row
- `GetColumnValues()` - Get all values from a column
- `ToArray()` - Convert sheet to 2D string array

## Requirements

- .NET 9.0 or later
- Encrypted Excel files (.xlsx or .xls format)

## Dependencies

- NPOI 2.7.4 (free, no license required)

## License

This project is licensed under the MIT License.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
