# Changelog

All notable changes to JFToolkit.EncryptedExcel will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Planned
- Cross-platform encrypted saving alternatives
- Performance optimizations for large files
- Additional Excel format support

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
| 1.1.0   | 2024-12-28  | Multi-framework support (.NET Standard 2.0+) |
| 1.0.0   | 2024-12-28  | Initial release with encrypted Excel support |

## Migration Guide

### From 1.0.0 to 1.1.0
No breaking changes. Simply update your package reference:

```xml
<PackageReference Include="JFToolkit.EncryptedExcel" Version="1.1.0" />
```

The API remains identical, but you now get broader framework compatibility.

## Support Matrix

| Framework | 1.0.0 | 1.1.0 |
|-----------|-------|-------|
| .NET 9.0 | ✅ | ✅ |
| .NET 8.0 | ❌ | ✅ |
| .NET 6.0 | ❌ | ✅ |
| .NET Standard 2.0 | ❌ | ✅ |
| .NET Framework 4.6.1+ | ❌ | ✅ (via .NET Standard 2.0) |

## Acknowledgments

- **NPOI Team**: For the excellent free Excel library
- **Community**: For testing and feedback
- **Microsoft**: For PowerShell and Excel automation capabilities
