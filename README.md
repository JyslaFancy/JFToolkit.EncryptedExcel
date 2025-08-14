# JFToolkit.EncryptedExcel

[![NuGet](https://img.shields.io/nuget/v/JFToolkit.EncryptedExcel.svg)](https://www.nuget.org/packages/JFToolkit.EncryptedExcel/)
[![Downloads](https://img.shields.io/nuget/dt/JFToolkit.EncryptedExcel.svg)](https://www.nuget.org/packages/JFToolkit.EncryptedExcel/)
[![GitHub](https://img.shields.io/github/license/JyslaFancy/JFToolkit.EncryptedExcel)](https://github.com/JyslaFancy/JFToolkit.EncryptedExcel/blob/main/LICENSE)

A powerful .NET library for reading, modifying, and saving password-encrypted Excel files. Built on NPOI with intelligent automation for maintaining encryption. Supports .xlsx, .xlsm (macro-enabled), and .xls formats.

## 🚀 Key Features

- ✅ **Read password-encrypted Excel files** (.xlsx, .xlsm, .xls)
- ✅ **Modify data with type safety** (strings, numbers, dates, booleans)
- ✅ **Add new rows and columns** with ease
- ✅ **Save with encryption** using Excel automation
- ✅ **Extension methods** for intuitive Excel manipulation
- ✅ **Comprehensive error handling** with graceful fallbacks
- ✅ **Real-world tested** with production encrypted files

## 🎯 Compatibility

| Framework | Support |
|-----------|---------|
| .NET Standard 2.0 | ✅ (.NET Framework 4.6.1+, .NET Core 2.0+) |
| .NET 6.0 (LTS) | ✅ |
| .NET 8.0 (LTS) | ✅ |
| .NET 9.0 | ✅ |

## 📦 Installation

```bash
# Package Manager Console
Install-Package JFToolkit.EncryptedExcel

# .NET CLI
dotnet add package JFToolkit.EncryptedExcel

# PackageReference
<PackageReference Include="JFToolkit.EncryptedExcel" Version="1.1.0" />
```

## 🔧 Quick Start

### Opening and Reading Encrypted Excel

```csharp
using JFToolkit.EncryptedExcel;

// Open encrypted Excel file
using var reader = EncryptedExcelReader.OpenFile(@"C:\data\encrypted.xlsx", "password123");
var workbook = reader.Workbook!;
var sheet = workbook.GetSheetAt(0);

// Read data with type safety
string name = sheet.GetStringValue(1, 0);     // Row 1, Column A
int age = sheet.GetCellValue<int>(1, 1);      // Row 1, Column B
double salary = sheet.GetCellValue<double>(1, 2); // Row 1, Column C
```

### Modifying Data

```csharp
// Update existing cells
sheet.SetCellValue(1, 1, 32);        // Update age
sheet.SetCellValue(1, 2, 85000.0);   // Update salary

// Add new row
sheet.AddRow("John Doe", 28, 75000.0, DateTime.Now, true);
```

### Saving with Encryption 🔒

```csharp
// Save with same password encryption
EncryptedExcelWriter.SaveEncryptedToFile(workbook, @"C:\data\output.xlsx", "password123");

// Or save without encryption
workbook.SaveToFile(@"C:\data\output_unencrypted.xlsx");
```

## 💼 Real-World Example

Perfect for applications that need to process encrypted Excel files while maintaining security:

```csharp
public async Task ProcessEmployeeData(string filePath, string password)
{
    // 1. Open encrypted employee file
    using var reader = EncryptedExcelReader.OpenFile(filePath, password);
    var workbook = reader.Workbook!;
    var sheet = workbook.GetSheetAt(0);
    
    // 2. Apply business logic (e.g., salary increases)
    for (int i = 1; i <= sheet.LastRowNum; i++)
    {
        var row = sheet.GetRow(i);
        if (row?.GetCell(2) != null) // Salary column
        {
            var salaryCell = row.GetCell(2);
            if (double.TryParse(salaryCell.ToString(), out double salary))
            {
                salaryCell.SetCellValue(salary * 1.05); // 5% increase
            }
        }
    }
    
    // 3. Add processing timestamp
    var timestampRow = sheet.CreateRow(sheet.LastRowNum + 1);
    timestampRow.CreateCell(0).SetCellValue($"Processed: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
    
    // 4. Save with same encryption
    string outputPath = filePath.Replace(".xlsx", "_processed.xlsx");
    EncryptedExcelWriter.SaveEncryptedToFile(workbook, outputPath, password);
    
    Console.WriteLine($"✅ Processed: {Path.GetFileName(outputPath)}");
}
```

## 🛠️ Extension Methods

The library includes helpful extension methods for common operations:

```csharp
// Type-safe cell reading
string value = sheet.GetStringValue(row, col);
int number = sheet.GetCellValue<int>(row, col);
DateTime date = sheet.GetCellValue<DateTime>(row, col);
bool flag = sheet.GetCellValue<bool>(row, col);

// Easy cell writing
sheet.SetCellValue(row, col, "Hello World");
sheet.SetCellValue(row, col, 42);
sheet.SetCellValue(row, col, DateTime.Now);

// Quick row creation
sheet.AddRow("Name", 25, 50000.0, DateTime.Now, true);

// Simple file saving
workbook.SaveToFile(@"C:\output\file.xlsx");
```

## 🔐 Encryption Support

### Automatic Encryption (Recommended)
The library automatically handles encryption using Excel automation:

```csharp
// This will use Excel automation to maintain encryption
EncryptedExcelWriter.SaveEncryptedToFile(workbook, "output.xlsx", "password");
```

### Requirements for Encryption
- Microsoft Excel installed on the machine
- Windows environment (for COM automation)

### Fallback Options
If Excel automation is unavailable, the library:
1. Saves as unencrypted file
2. Provides clear instructions for manual encryption
3. Ensures your workflow never breaks

## 📋 Supported Formats

- ✅ **Excel 2007+ (.xlsx)** - Full support
- ✅ **Excel 97-2003 (.xls)** - Full support
- ✅ **Password-protected files** - Both formats
- ✅ **Multiple worksheets** - Complete access

## 🧪 Tested & Reliable

- ✅ Tested with real-world encrypted Excel files
- ✅ Handles various data types and formats
- ✅ Robust error handling and recovery
- ✅ Memory-efficient with proper disposal
- ✅ Thread-safe operations

## 🔧 Requirements

- **.NET Standard 2.0** or higher
- **NPOI 2.7.4** (automatically installed)
- **Microsoft Excel** (for encrypted saving - optional)

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guidelines](CONTRIBUTING.md) for details.

### Development Setup

1. Clone the repository
2. Open in Visual Studio or VS Code
3. Run tests: `dotnet test`
4. Build: `dotnet build`

## 📝 Issues & Support

- **Bug Reports**: [GitHub Issues](https://github.com/yourusername/JFToolkit.EncryptedExcel/issues)
- **Feature Requests**: [GitHub Issues](https://github.com/yourusername/JFToolkit.EncryptedExcel/issues)
- **Documentation**: [Wiki](https://github.com/yourusername/JFToolkit.EncryptedExcel/wiki)

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🏷️ Version History

### v1.1.0
- Added support for .NET Standard 2.0 (broader compatibility)
- Added support for .NET 6.0, .NET 8.0, and .NET 9.0
- Compatible with .NET Framework 4.6.1+ via .NET Standard 2.0
- Compatible with .NET Core 2.0+ via .NET Standard 2.0
- Improved cross-platform compatibility

### v1.0.0
- Initial release
- Read password-encrypted Excel files
- Modify data with type safety  
- Save with encryption using Excel automation
- Comprehensive documentation and examples

---

**Made with ❤️ for developers who work with encrypted Excel files**

⭐ **Star this repo if you find it useful!**
