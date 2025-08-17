# TEST PLAN – Encrypted .xlsm Redevelopment (2.0.0-dev)

## 1. Purpose
Define the testing strategy ensuring reliable open → modify → save (re‑encrypt) workflow for password‑protected .xlsm workbooks without relying on NPOI or Excel COM.

## 2. Scope
In-scope:
- Agile encrypted .xlsm parsing, decryption, modification, encryption
- Preservation of non-edited content (macros, relationships, styles)
- Minimal cell edit API correctness
- Error taxonomy behavior
- Performance & memory boundaries (baseline metrics)

Out-of-scope (initial phases):
- Standard (legacy RC4) encryption
- Digital signature / VBA signing validation
- Complex formula recalculation
- Rich editing (styles, merges, charts modifications)

## 3. Test Categories
| Category | Goal |
|----------|------|
| Unit – Binary Parsing | Correctly extract fields from EncryptionInfo structures |
| Unit – Key Derivation (KDF) | Produce expected keys/hashes for known vectors |
| Unit – XML Editing | Insert/update cell values; sharedStrings consistency |
| Unit – Error Handling | Trigger and validate specific exceptions |
| Integration – Decrypt/Inspect | Full decrypt; verify sheet & macro presence |
| Integration – Round Trip (No Edit) | Decrypt→Encrypt yields structurally equivalent content (macro binary equality) |
| Integration – Round Trip (With Edit) | Targeted cell changes only; others untouched |
| Security – Wrong Password | Reject with InvalidPasswordException reliably |
| Security – Tampered EncryptionInfo | Detect and fail fast |
| Performance – Large File | Validate throughput & memory ceiling |
| Stress – High Spin Count | Ensure reasonable slowdown but no hang/crash |
| Regression | Protect against previously fixed defects |

## 4. Environments
| Dimension | Requirement |
|----------|-------------|
| OS | Windows, Linux (for decrypt/edit). Mac optional later |
| .NET Runtimes | net8.0 primary; validate net6.0; netstandard2.0 (where feasible) |
| Excel Installation | NOT required (independent solution) |

## 5. Test Data Management
Fixtures:
1. `sample_agile_basic.xlsm` – Small workbook, 1 sheet, few cells, password: `Test123!`
2. `sample_agile_macros.xlsm` – Contains VBA project with simple macro.
3. `sample_agile_multisheet.xlsm` – 10 sheets, varied data types.
4. `sample_agile_large.xlsm` – ~50k rows * 10 columns (performance).
5. `sample_agile_highspin.xlsm` – Same as basic but high spinCount (e.g., 500k).

Storage:
- Not in repo (potentially sensitive). Use encrypted zip or generate via internal script.
- Provide script `tools/generate-fixtures.ps1` (future) to build them using Excel (one-time) then compress.

Verification Artifacts:
- Baseline JSON descriptor per fixture: sheet count, sharedStrings count, macro hash (SHA256), workbook rel list.
- Expected output descriptor after edits.

## 6. Unit Test Specifications
### 6.1 EncryptionInfoParserTests
- Parse_ValidAgileInfo_ReturnsExpectedFields
- Parse_UnsupportedVersion_ThrowsUnsupportedEncryptionException
- Parse_TruncatedStream_ThrowsEncryptionInfoCorruptException

### 6.2 AgileKeyDerivationTests
- DeriveKey_KnownVector_SHA1_ProducesExpectedBytes
- DeriveKey_InvalidSpinCount_Throws
- DeriveKey_LargeSpinCount_CompletesWithinThreshold (e.g., < 3s @ 100k spins)

### 6.3 WorkbookEditorTests
- SetCell_NewString_AddsSharedStringAndUpdatesCounts
- SetCell_ExistingNumeric_OverwritesValueWithoutExtraSharedString
- SetCell_NewRowAndCell_CreatesStructure

### 6.4 ErrorHandlingTests
- Open_WrongPassword_ThrowsInvalidPasswordException
- Open_TamperedEncryptionInfo_ThrowsEncryptionInfoCorruptException

## 7. Integration Test Specifications
### 7.1 DecryptOnlyTests
- Decrypt_BasicWorkbook_VerifiesSheetAndMacroPresence
- Decrypt_MacroWorkbook_PreservesVbaProjectBinHash

### 7.2 RoundTripNoEditTests
- RoundTrip_NoEdits_OutputDecryptsAndAllHashesMatch (macro & unchanged parts)

### 7.3 RoundTripWithEditTests
- RoundTrip_EditSingleCell_OnlyTargetSheetXmlDiffers
- RoundTrip_AddNewString_SharedStringsCountIncrementedOne

### 7.4 PerformanceTests
- RoundTrip_LargeWorkbook_CompletesWithinTimeBudget (define baseline, e.g., < 10s)
- Memory_LargeWorkbook_PeakBelow( X ) (monitor via diagnostic hooks)

### 7.5 StressTests
- Open_HighSpinCount_CompletesWithinReasonableTime (warn if > threshold)

### 7.6 NegativeSecurityTests
- Decrypt_ModifiedEncryptedPackageTail_ThrowsEncryptionIntegrityException
- Decrypt_ReplacedMacroPart_DetectsMismatchIfIntegrityEnabled

## 8. Tooling & Frameworks
- Test framework: xUnit (lightweight & parallelizable)
- FluentAssertions for readable assertions
- BenchmarkDotNet (optional) for KDF & encryption micro-benchmarks
- Custom diagnostic hooks (events or logger interface) for performance capture

## 9. Automation & CI
Pipeline (GitHub Actions suggested):
1. Matrix: OS (windows-latest, ubuntu-latest) × TFMs (net8.0, net6.0)
2. Steps:
   - Restore
   - Build
   - Run unit tests
   - (Windows only) Integration tests with fixtures (Linux if fixtures added & cross‑platform logic stable)
3. Artifacts: Test results (trx/junit), code coverage (coverlet), performance summary JSON.

Quality Gates:
- Unit test pass rate: 100%
- Code coverage (core logic): target 80%+ for parser, KDF, encryptor
- Lint / analyzers: zero warnings (or documented suppressions)
- No allocations > defined threshold in tight loops (benchmark baseline locked after M3)

## 10. Metrics & Monitoring
Collected per integration run:
- Decrypt time (ms)
- Edit time (ms) for N cell operations
- Re-encrypt time (ms)
- Total round trip time (ms)
- Peak working set (MB) (Windows PerformanceCounter / cross-platform proc stats)
- Macro part hash equality (bool)

## 11. Test Data Integrity Strategy
- Each fixture accompanied by baseline JSON (committed if license-safe or generated at runtime)
- Hash algorithm: SHA256
- Validation rejects unexpected new/unremoved parts (unless explicitly allowed list)

## 12. Risk-Based Prioritization
| Feature | Risk | Priority |
|---------|------|---------|
| Parser correctness | High (foundation) | P0 |
| KDF accuracy | High | P0 |
| AES encryption integrity | High | P0 |
| Macro part preservation | Medium | P1 |
| SharedStrings edit logic | Medium | P1 |
| Performance large file | Medium | P2 |
| High spinCount stress | Low | P3 |

## 13. Phased Exit Criteria
| Phase | Criteria |
|-------|----------|
| M1 (Parse+KDF) | Parser + KDF unit tests green; wrong password test passes |
| M2 (Decrypt) | Can decrypt & list sheets + macro hash |
| M3 (Round Trip No Edit) | Re-encryption yields openable file; macro hash identical |
| M4 (Edit) | Cell edit diff limited to target sheet & sharedStrings |
| M5 (Integrity) | Integrity checks optional mode implemented |
| M6 (Perf Hardening) | Large workbook within performance budget |

## 14. Open Questions
- Do we snapshot coverage thresholds early (freeze after M3)?
- Should macro hash verification be opt-in (performance)?
- Where to store encrypted fixtures securely in CI (private repo or encrypted artifact)?

## 15. Immediate Actions
1. Create test project `JFToolkit.EncryptedExcel.Tests` (xUnit)
2. Add skeleton tests (ParserTests, KeyDerivationTests)
3. Add helper for loading binary test fixtures
4. Add GitHub Actions workflow yaml (build + unit tests)

---
*Living document – update as implementation detail evolves.*
