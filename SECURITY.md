# Security Policy

## Supported Versions

We actively support the following versions with security updates:

| Version | Supported          |
| ------- | ------------------ |
| 1.1.x   | :white_check_mark: |
| 1.0.x   | :white_check_mark: |
| < 1.0   | :x:                |

## Reporting a Vulnerability

We take security seriously. If you discover a security vulnerability, please follow these steps:

### 🔒 Private Disclosure

**DO NOT** create a public GitHub issue for security vulnerabilities.

Instead, please:
1. Email the maintainer privately
2. Include detailed information about the vulnerability
3. Provide steps to reproduce if possible
4. Allow reasonable time for a fix before public disclosure

### 📧 Contact Information

For security issues, contact:
- **Primary**: Create a private security advisory on GitHub
- **Alternative**: Open a draft security advisory in the Security tab

### 🕒 Response Timeline

- **Initial Response**: Within 48 hours
- **Status Update**: Within 7 days
- **Fix Timeline**: Depends on severity
  - Critical: 1-3 days
  - High: 1-2 weeks
  - Medium: 2-4 weeks
  - Low: Next planned release

### 🎯 Scope

This security policy covers:
- The main JFToolkit.EncryptedExcel library
- Any bundled dependencies
- Example code that could lead to security issues

### 🛡️ Security Considerations

#### Password Handling
- Passwords are passed as strings (not SecureString) for NPOI compatibility
- Passwords are not logged or persisted
- Memory is not explicitly cleared (relies on GC)

#### File System Access
- Library requires file system access to read/write Excel files
- No validation of file paths (caller responsibility)
- Temporary files may be created during automation

#### External Dependencies
- NPOI library for Excel manipulation
- PowerShell for encrypted saving (Windows only)
- Excel COM automation as fallback

#### Known Limitations
- PowerShell execution for encrypted saving
- COM object creation for Excel automation
- No input validation on file paths

### 🔍 Vulnerability Assessment

#### Common Attack Vectors
- **Path Traversal**: Not validated - caller must sanitize paths
- **Code Injection**: PowerShell scripts are parameterized
- **Memory Disclosure**: Passwords stored in managed strings
- **Privilege Escalation**: Requires PowerShell/Excel permissions

#### Security Recommendations
1. **Validate file paths** before calling library methods
2. **Use secure password sources** (not hardcoded)
3. **Run with minimal permissions** when possible
4. **Monitor PowerShell execution** in restricted environments

### 📋 Security Checklist for Contributors

When contributing code, ensure:
- [ ] No hardcoded passwords or secrets
- [ ] Input validation where appropriate
- [ ] No unsafe file operations
- [ ] PowerShell scripts are parameterized
- [ ] Error messages don't leak sensitive information
- [ ] Dependencies are from trusted sources

### 🔄 Security Updates

Security updates will be:
- Released as patch versions (e.g., 1.1.1)
- Documented in CHANGELOG.md
- Announced in release notes
- Tagged with "security" label

### 📚 Additional Resources

- [OWASP Secure Coding Practices](https://owasp.org/www-project-secure-coding-practices-quick-reference-guide/)
- [Microsoft Security Development Lifecycle](https://www.microsoft.com/en-us/securityengineering/sdl)
- [NuGet Package Security](https://docs.microsoft.com/en-us/nuget/policies/ecosystem)

---

Thank you for helping keep JFToolkit.EncryptedExcel secure! 🔐
