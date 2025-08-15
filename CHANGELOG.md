# Changelog

All notable changes to JFToolkit.EncryptedExcel will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Planned
- Cross-platform encrypted saving alternatives
- Performance optimizations for large files
- Additional Excel format support

## [1.4.0] - 2025-08-15

### 🚨 BREAKING CHANGES
- **Class Rename**: `EncryptedXlsmManager` renamed to `SecureExcelWorkbook` for better clarity
- **API Update**: All references to `EncryptedXlsmManager` must be updated to `SecureExcelWorkbook`

### ✨ Major Improvements
- **Repository Cleanup**: Removed 9 unnecessary files from main library for cleaner, focused codebase
- **Enhanced Documentation**: Added prominent Excel installation requirements and platform compatibility tables
- **Zero Build Warnings**: Resolved all 61 build warnings (XML documentation + platform-specific COM warnings)
- **Professional Package**: Streamlined library focused on core functionality

### 🗂️ Files Removed (Repository Cleanup)
- `Examples.cs` - Moved examples to separate documentation
- `ModifyAndSaveExamples.cs` - Consolidated into main examples
- `MacroSupportTest.cs` - Moved to test projects
- `AsposeVsNpoiAnalysis.cs` - Development analysis removed
- `AlternativeEncryptionMethods.cs` - Future feature preparation removed
- `PowerShellExcelHelper.cs` - Functionality integrated into main classes
- `OpenXmlEncryptionHelper.cs` - Future cross-platform feature removed
- `TestMacroFix.cs` - Development test removed
- `XlsmDiagnosticTest.cs` - Development diagnostic removed

### 📚 Documentation Enhancements
- **Excel Requirements**: Added clear warnings that Microsoft Excel is required for encryption features
- **Platform Compatibility**: Added comprehensive compatibility table (Windows vs. other platforms)
- **Installation Guide**: Updated with clear requirements for different deployment scenarios
- **Package Description**: Updated NuGet description to clearly state Excel dependency

### 🔧 Technical Improvements
- **Complete XML Documentation**: Added comprehensive XML docs for all public members
- **Warning-Free Builds**: Resolved CS1591 (missing XML documentation) and CA1416 (platform-specific) warnings
- **Better IntelliSense**: All public APIs now have detailed documentation for improved developer experience
- **Clean Project Structure**: Focused on essential files for production use

### 💡 API Clarity
- **SecureExcelWorkbook**: New name better describes the complete encrypted Excel workflow
- **Method Documentation**: All methods now include Excel requirement warnings where applicable
- **Error Messages**: Improved error messages for environments without Excel

## [1.2.0] - 2025-08-14

### Added
- **Macro-Enabled Excel Support**: Explicit support for .xlsm (macro-enabled) files
- **Enhanced Documentation**: Clear indication of .xlsm format support throughout
- **Macro Example**: Added dedicated example for working with .xlsm files
- **Comprehensive Testing**: Added MacroSupportTest.cs for validation

### Enhanced
- **Package Tags**: Added 'xlsm' and 'macro' tags for better discoverability
- **File Format Documentation**: Updated all documentation to mention .xlsx, .xlsm, and .xls support
- **API Documentation**: Enhanced XML comments to specify supported formats

### Technical Details
- NPOI's XSSFWorkbook natively supports both .xlsx and .xlsm formats
- Macros are preserved when reading and saving .xlsm files
- VBA code execution is not supported (security limitation)
- All standard Excel features work normally with .xlsm files

## [1.1.1] - 2025-08-14

### Fixed
- **GitHub Repository Links**: Fixed NuGet package metadata to correctly link to GitHub repository
- **Package Cleanup**: Removed references to deleted development files from package includes
- **Professional Package**: NuGet package now properly links to https://github.com/JyslaFancy/JFToolkit.EncryptedExcel

### Technical Details
- Updated `PackageProjectUrl` and `RepositoryUrl` in project file
- Cleaned up package file references after repository cleanup
- Maintained all existing functionality and compatibility

## [1.1.0] - 2024-12-28

### Added
- **Multi-Framework Support**: Expanded compatibility to support .NET Standard 2.0, .NET 6.0, .NET 8.0, and .NET 9.0
- **Broader Compatibility**: Now supports legacy .NET Framework applications through .NET Standard 2.0
- **Enhanced Package Metadata**: Improved NuGet package information and documentation

### Changed
- **Target Frameworks**: Changed from single .NET 9.0 target to multi-target framework support
- **Dependencies**: Updated NPOI package references for framework-specific compatibility
- **Build Configuration**: Improved multi-framework build process

### Technical Details
- Supports 95%+ of .NET developers across different framework versions
- Maintains backward compatibility with existing code
- Symbol packages included for debugging support

## [1.0.0] - 2024-12-28

### Added
- **Core Functionality**: Initial release with encrypted Excel reading capabilities
- **EncryptedExcelReader**: Main class for opening password-protected Excel files
- **EncryptedExcelWriter**: Advanced saving with encryption support
- **Multiple Automation Methods**: PowerShell and COM automation for encrypted saving
- **Extension Methods**: Convenient Excel manipulation helpers
- **Comprehensive Examples**: Real-world usage examples and test cases

### Features
- 📖 **Read encrypted Excel files** with password protection
- ✏️ **Modify Excel data** using familiar NPOI syntax
- 💾 **Save with encryption** using multiple fallback methods
- 🔧 **Extension methods** for easier Excel manipulation
- 🎯 **Real-world tested** with actual encrypted Excel files

### Core Classes
- `EncryptedExcelReader`: Opens encrypted Excel files
- `EncryptedExcelWriter`: Saves Excel files with encryption
- `ExcelExtensions`: Helper extension methods
- `PowerShellExcelHelper`: PowerShell-based automation
- `ExcelAutomationHelper`: COM-based automation fallback

### Dependencies
- NPOI 2.7.4: Free Excel library for .NET
- System.Management.Automation: PowerShell integration
- Microsoft.Office.Interop.Excel: COM automation fallback

### Documentation
- Comprehensive README with examples
- Real-world usage scenarios
- PowerShell verification scripts
- NuGet package guide

### Tested Scenarios
- ✅ Opening encrypted Excel files
- ✅ Reading and modifying data
- ✅ Saving with original encryption
- ✅ Multiple sheet handling
- ✅ Cross-framework compatibility

---

## Version History Summary

| Version | Release Date | Key Features |
|---------|-------------|--------------|
| 1.4.0   | 2025-08-15  | 🚨 BREAKING: Renamed to SecureExcelWorkbook, major cleanup, zero warnings |
| 1.2.0   | 2025-08-14  | Added explicit macro-enabled Excel (.xlsm) support |
| 1.1.1   | 2025-08-14  | Fixed GitHub repository links in NuGet package |
| 1.1.0   | 2024-12-28  | Multi-framework support (.NET Standard 2.0+) |
| 1.0.0   | 2024-12-28  | Initial release with encrypted Excel support |

## Migration Guide

### From 1.2.0 to 1.4.0 (BREAKING CHANGES)
This version contains breaking changes due to class rename:

**Required Changes:**
```csharp
// OLD (v1.2.0 and earlier)
using (var manager = new EncryptedXlsmManager())
{
    // ... your code
}

// NEW (v1.4.0+)
using (var workbook = new SecureExcelWorkbook())
{
    // ... your code (same methods, just different class name)
}
```

**Update your package reference:**
```xml
<PackageReference Include="JFToolkit.EncryptedExcel" Version="1.4.0" />
```

**What's New in 1.4.0:**
- Much cleaner repository structure (9 unnecessary files removed)
- Better documentation with clear Excel requirements
- Zero build warnings for professional development
- Same reliable functionality with improved clarity

### From 1.1.1 to 1.2.0
No breaking changes. Simply update your package reference:

```xml
<PackageReference Include="JFToolkit.EncryptedExcel" Version="1.2.0" />
```

New features in 1.2.0:
- Explicit .xlsm (macro-enabled) file support
- Enhanced documentation for all supported formats
- New macro-enabled example in Examples.cs

### From 1.1.0 to 1.1.1
No breaking changes. Simply update your package reference:

```xml
<PackageReference Include="JFToolkit.EncryptedExcel" Version="1.1.1" />
```

This update only improves the NuGet package metadata with correct GitHub links.

### From 1.0.0 to 1.1.0
No breaking changes. Simply update your package reference:

```xml
<PackageReference Include="JFToolkit.EncryptedExcel" Version="1.1.0" />
```

The API remains identical, but you now get broader framework compatibility.

## Support Matrix

| Framework | 1.0.0 | 1.1.0 | 1.1.1 | 1.2.0 | 1.4.0 |
|-----------|-------|-------|-------|-------|-------|
| .NET 9.0 | ✅ | ✅ | ✅ | ✅ | ✅ |
| .NET 8.0 | ❌ | ✅ | ✅ | ✅ | ✅ |
| .NET 6.0 | ❌ | ✅ | ✅ | ✅ | ✅ |
| .NET Standard 2.0 | ❌ | ✅ | ✅ | ✅ | ✅ |
| .NET Framework 4.6.1+ | ❌ | ✅ (via .NET Standard 2.0) | ✅ (via .NET Standard 2.0) | ✅ (via .NET Standard 2.0) | ✅ (via .NET Standard 2.0) |

## File Format Support

| Format | Description | 1.0.0 | 1.1.0+ | 1.2.0+ | 1.4.0+ |
|--------|-------------|-------|--------|--------|--------|
| .xlsx | Excel Workbook | ✅ | ✅ | ✅ | ✅ |
| .xls | Excel 97-2003 | ✅ | ✅ | ✅ | ✅ |
| .xlsm | Excel Macro-Enabled | ✅* | ✅* | ✅ (Explicit) | ✅ (Explicit) |

*Supported but not explicitly documented

## Acknowledgments

- **NPOI Team**: For the excellent free Excel library
- **Community**: For testing and feedback
- **Microsoft**: For PowerShell and Excel automation capabilities
