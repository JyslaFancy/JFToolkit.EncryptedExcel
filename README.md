# JFToolkit.EncryptedExcel (2.0.0-dev Experimental Reset)

This branch (develop / 2.0.0-dev) is an experimental **ground‑up redesign** removing all dependencies on:
* NPOI
* Excel COM automation (Microsoft Office)

Current state: a **minimal placeholder** (`WorkbookPlaceholder`) that demonstrates the intended lightweight API surface while a new cross‑platform encrypted Excel engine is researched.

## Why the Reset?
The 1.x line relied on NPOI + Excel Automation for encryption. Goals for 2.x:
1. Eliminate Excel installation requirement.
2. Reduce dependency footprint.
3. Provide predictable, testable core with pluggable format/encryption providers.
4. Enable future streaming & zero‑copy scenarios.

## What Exists Now
* `WorkbookPlaceholder` – in‑memory single sheet, tab‑delimited export.
* No real Excel parsing or encryption yet.

## Near-Term Roadmap (Subject to Change)
1. Define core abstractions: IWorkbookModel, IWorksheetModel, ICellModel.
2. Implement thin OpenXML (.xlsx) reader (shared strings + basic cells).
3. Add write support (simple sheet, limited types).
4. Pluggable encryption layer (Agile encryption spec) without Excel.
5. Streaming load/save for large sheets.

## Example (Placeholder)
```csharp
using JFToolkit.EncryptedExcel;

var wb = WorkbookPlaceholder.Create();
wb.SetCell(0,0,"Hello");
wb.SetCell(1,1,123);
wb.SavePlainText("out.tsv");
```

## Not Production Ready
Do not publish packages from this branch. It is intentionally incomplete.

## Legacy (1.x)
For stable functionality (encrypted .xlsx / .xls using NPOI + Excel), use the `main` branch and released NuGet versions (<=1.5.1).

## License
MIT (see LICENSE)

---
© 2025 JFToolkit – Experimental development in progress.

## 🚀 Key Features

- ✅ **SecureExcelWorkbook API** - Simple workflow for encrypted Excel files
- ✅ **Read encrypted .xlsx / .xls files** - Works on any platform (no Excel required)
- ✅ **Modify data** - Full editing capabilities with type safety
- ⚠️ **Save with encryption (.xlsx / .xls)** - Requires Microsoft Excel (Windows)
- ⚠️ **.xlsm note** - Macro-enabled files can be opened if already decrypted, but encrypted .xlsm save is not supported
- ✅ **Save to separate files** - Modify and save without overwriting originals

## 🎯 Platform Compatibility

| Feature | Windows + Excel | Windows Only | Any Platform |
|---------|-----------------|--------------|--------------|
| Read encrypted files | ✅ | ✅ | ✅ |
| Modify data | ✅ | ✅ | ✅ |
| Save with encryption | ✅ | ❌ | ❌ |
| Save without encryption | ✅ | ✅ | ✅ |

### Framework Support
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
<PackageReference Include="JFToolkit.EncryptedExcel" Version="1.5.0" />
```

## 🔧 Quick Start with SecureExcelWorkbook

### Simple Workflow: Open → Modify → Save

```csharp
using JFToolkit.EncryptedExcel;

// Open encrypted Excel file
using var workbook = SecureExcelWorkbook.Open(@"C:\path\to\encrypted.xlsx", "password123");

// Read current values
var currentValue = workbook.GetCellValue(0, 0, 2); // Sheet 0, Row 0, Column C
Console.WriteLine($"Current value: {currentValue}");

// Modify content
workbook.SetCellValue(0, 0, 2, "New Value");
workbook.SetCellValue(0, 1, 2, "Another Value");
workbook.SetCellValue(0, 2, 2, DateTime.Now);

// Save to a separate file (keeps original unchanged)
bool saved = workbook.SaveAs(@"C:\path\to\modified.xlsx");
if (saved)
{
    Console.WriteLine("✅ File saved with encryption (.xlsx)!");
    // When you open this file in Excel: SINGLE password prompt only!
}
else
{
    Console.WriteLine("❌ Encryption failed - Excel not available");
    // File operations will still work, but without encryption
}
```

> **💡 Note**: The `SaveAs()` method requires Microsoft Excel to be installed for encryption. If Excel is not available, the save operation will return `false` and you'll need to save without encryption or install Excel.

### Advanced Usage with Direct NPOI Access

```csharp
using JFToolkit.EncryptedExcel;

// For advanced scenarios, you can access the underlying NPOI workbook
using var workbook = SecureExcelWorkbook.Open(@"C:\data\encrypted.xlsx", "password123");

// Direct NPOI access for complex operations
var npoiWorkbook = workbook.Workbook;
var sheet = npoiWorkbook.GetSheetAt(0);

// Use full NPOI functionality
var row = sheet.CreateRow(10);
var cell = row.CreateCell(0);
cell.SetCellValue("Advanced modification");
```

### Legacy API (Still Supported)

```csharp
// The original EncryptedExcelReader is still available for backward compatibility
using var reader = EncryptedExcelReader.OpenFile(@"C:\data\encrypted.xlsx", "password123");
var workbook = reader.Workbook!;
var sheet = workbook.GetSheetAt(0);

// Use extension methods for easier data access
string name = sheet.GetStringValue(1, 0);     // Row 1, Column A
int age = sheet.GetCellValue<int>(1, 1);      // Row 1, Column B
double salary = sheet.GetCellValue<double>(1, 2); // Row 1, Column C
```

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

- ✅ **Excel 2007+ (.xlsx)** - Full support (encrypted read/write)
- ✅ **Excel 97-2003 (.xls)** - Full support (encrypted read/write)
- ⚠️ **Macro-enabled (.xlsm)** - Not supported for encrypted workflows in 1.5.0
- ✅ **Password-protected files** - .xlsx / .xls
- ✅ **Multiple worksheets** - Complete access

## 🧪 Tested & Reliable

- ✅ Tested with real-world encrypted Excel files
- ✅ Handles various data types and formats
- ✅ Robust error handling and recovery
- ✅ Memory-efficient with proper disposal
- ✅ Thread-safe operations

## 🔧 System Requirements

### ✅ For Reading & Modifying Encrypted Files
- Any .NET-compatible platform
- No additional software required

### ⚠️ For Saving WITH Encryption
- **Windows** operating system
- **Microsoft Excel** installed (any recent version)
- .NET Framework or .NET Core/.NET 5+

### Alternative: Save Without Encryption
```csharp
// Works on any platform - no Excel required
using var workbook = SecureExcelWorkbook.Open("encrypted.xlsx", "password");
workbook.SetCellValue(0, 0, 2, "Modified");

// Save without encryption (works everywhere)
workbook.Workbook.SaveToFile("output.xlsx");
```

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
