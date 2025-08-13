# Contributing to JFToolkit.EncryptedExcel

Thank you for your interest in contributing! This document provides guidelines for contributing to the project.

## 🚀 Getting Started

### Prerequisites
- .NET 6.0 SDK or later
- Visual Studio 2022 or VS Code
- Git

### Development Setup
1. Fork the repository
2. Clone your fork: `git clone https://github.com/yourusername/JFToolkit.EncryptedExcel.git`
3. Create a branch: `git checkout -b feature/your-feature-name`
4. Make your changes
5. Test your changes: `dotnet test`
6. Commit: `git commit -m "Add your feature"`
7. Push: `git push origin feature/your-feature-name`
8. Create a Pull Request

## 🧪 Testing

### Running Tests
```bash
# Run all tests
dotnet test

# Run tests for specific framework
dotnet test --framework net6.0
dotnet test --framework netstandard2.0
```

### Test Requirements
- All new features must include unit tests
- Tests should cover both success and failure scenarios
- Test with encrypted Excel files when possible

## 📝 Code Style

### C# Guidelines
- Follow standard C# naming conventions
- Use meaningful variable and method names
- Add XML documentation for public APIs
- Include using statements explicitly (for .NET Standard 2.0 compatibility)

### Example:
```csharp
/// <summary>
/// Opens an encrypted Excel file and returns a reader instance
/// </summary>
/// <param name="filePath">Path to the Excel file</param>
/// <param name="password">Password for decryption</param>
/// <returns>EncryptedExcelReader instance</returns>
public static EncryptedExcelReader OpenFile(string filePath, string password)
{
    // Implementation
}
```

## 🐛 Bug Reports

When reporting bugs, please include:
- Steps to reproduce
- Expected behavior
- Actual behavior
- .NET version and framework
- Excel file details (if applicable)
- Error messages or stack traces

### Bug Report Template:
```markdown
**Describe the bug**
A clear description of what the bug is.

**To Reproduce**
Steps to reproduce the behavior:
1. Go to '...'
2. Click on '....'
3. See error

**Expected behavior**
What you expected to happen.

**Environment**
- .NET Version: [e.g. .NET 6.0]
- OS: [e.g. Windows 11]
- Package Version: [e.g. 1.1.0]

**Additional context**
Any other context about the problem.
```

## ✨ Feature Requests

We welcome feature requests! Please:
- Check existing issues first
- Describe the use case
- Explain why it would benefit other users
- Consider if it fits the project scope

## 🔧 Pull Request Guidelines

### Before Submitting
- [ ] Code builds successfully
- [ ] All tests pass
- [ ] New tests added for new features
- [ ] Documentation updated
- [ ] XML comments added for public APIs

### Pull Request Process
1. Update documentation if needed
2. Add tests for new functionality
3. Ensure backwards compatibility
4. Update CHANGELOG.md if applicable
5. Follow the pull request template

## 📚 Documentation

### XML Documentation
All public APIs must have XML documentation:
```csharp
/// <summary>
/// Brief description of what the method does
/// </summary>
/// <param name="parameterName">Description of the parameter</param>
/// <returns>Description of what is returned</returns>
/// <exception cref="ExceptionType">When this exception is thrown</exception>
```

### README Updates
- Keep examples current
- Update version information
- Add new features to the feature list

## 🏗️ Project Structure

```
JFToolkit.EncryptedExcel/
├── src/
│   ├── EncryptedExcelReader.cs     # Main reader class
│   ├── EncryptedExcelWriter.cs     # Main writer class
│   ├── ExcelExtensions.cs          # Extension methods
│   ├── ExcelAutomationHelper.cs    # COM automation
│   └── PowerShellExcelHelper.cs    # PowerShell automation
├── examples/                       # Example files
├── tests/                         # Unit tests
├── docs/                          # Documentation
└── README.md
```

## 🎯 Areas for Contribution

### High Priority
- [ ] Improve encrypted saving reliability
- [ ] Add more unit tests
- [ ] Performance optimizations
- [ ] Cross-platform Excel automation alternatives

### Medium Priority
- [ ] Support for more Excel features
- [ ] Better error messages
- [ ] Additional extension methods
- [ ] Documentation improvements

### Low Priority
- [ ] Code refactoring
- [ ] Additional examples
- [ ] Benchmark tests

## 🤝 Code of Conduct

### Our Pledge
We are committed to making participation in this project a harassment-free experience for everyone.

### Our Standards
- Be respectful and inclusive
- Focus on constructive feedback
- Help others learn and grow
- Keep discussions on-topic

## 📧 Contact

- **Issues**: Use GitHub Issues for bugs and feature requests
- **Discussions**: Use GitHub Discussions for general questions
- **Security**: Email security issues privately

## 🎉 Recognition

Contributors will be recognized in:
- CONTRIBUTORS.md file
- GitHub contributor list
- Release notes (for significant contributions)

Thank you for contributing to JFToolkit.EncryptedExcel! 🙏
