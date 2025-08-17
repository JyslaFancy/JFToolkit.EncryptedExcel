# Gap Analysis: Encrypted .xlsm Workflow (Open → Modify → Save Re‑Encrypted)

## 1. Business / Functional Requirement
Open an existing **password‑protected .xlsm (macro‑enabled)** Excel workbook, copy it to a new file, modify selected cell values, and save the new file as **password‑protected .xlsm** preserving:
- All original worksheets (names, order, visibility)
- Macros / VBA project intact
- Workbook / sheet protection states
- Pivot caches, charts, drawings, defined names, styles, relationships
- Encryption strength (at least Agile / AES based) equivalent to source

High‑level steps expected:
1. Input: (path, password)
2. Decrypt & load workbook (structure + VBA + relationships)
3. Clone to in‑memory editable representation
4. Apply cell edits (API surface to set values reliably; type inference or explicit types)
5. Re‑encrypt and write to new .xlsm output using the SAME (or specified new) password
6. Validate output (openable by Excel with single password prompt; macros preserved; no repair dialog)

## 2. Current State (develop branch 2.0.0‑dev)
| Aspect | Current Status | Notes |
|--------|----------------|-------|
| Dependency stack | NPOI & Excel COM removed | Clean slate using Open XML SDK only (DocumentFormat.OpenXml) |
| Decryption capability | NOT IMPLEMENTED | `XlsmEncryptionHelper.OpenModifyAndSaveEncryptedXlsm` throws NotImplementedException |
| Encryption capability | NOT IMPLEMENTED | No writer for Agile / Standard Office encryption container |
| OOXML package parsing | PARTIAL INFRA (via OpenXml SDK planned) | Can open unencrypted packages only (not wired yet) |
| Extraction of EncryptionInfo | NOT IMPLEMENTED | Need binary parser for EncryptionInfo + key derivation params |
| Key derivation (SHA1 / SHA512 / etc.) | NOT IMPLEMENTED | Spin count + salt logic missing |
| Stream cipher / block encryption pipeline | NOT IMPLEMENTED | Need AES-CBC or AES-ECB per spec with HMAC as required |
| VBA project preservation | AT RISK | Requires copying binary parts (vbaProject.bin, relationships) untouched |
| Macro signatures / digital signatures | NOT IN SCOPE (initial) | If signed, resigning or signature preservation policy needed |
| Cell modification layer | NOT IMPLEMENTED | Need minimal worksheet editing abstraction once decrypted |
| API surface | PLACEHOLDER ONLY | No public stable methods; only stub helper |
| Error handling / diagnostics | NONE | Need structured exception taxonomy |
| Tests (unit/integration) | NONE | Need fixtures: encrypted sample .xlsm + golden verification |
| Performance considerations | N/A | Must avoid full in‑memory duplication where possible |
| Security considerations | N/A | Must use secure disposal of derived keys / zeroization policy |
| Documentation | MINIMAL | Gap file (this) + project README (experimental) |

## 3. Required Capability Breakdown
### 3.1 Decrypt Phase
- Read OLE/ZIP compound (Office encryption wrapper) header
- Parse `EncryptionInfo` stream (Agile preferred; fallback Standard) including:
  - Version / flags
  - KeyData (salt, hash algo, cipher algo, block size, hash size, spin count)
  - KeyEncryptors (password/key based)
- Derive key from password: iterative hash (spinCount) + XOR folding (Standard) or algorithm chain (Agile)
- Validate by decrypting and checking package signature / integrity (e.g., compare decrypted stream header `[PK]` ZIP magic)

### 3.2 Load & Represent
- Mount decrypted ZIP (OPC / Open Packaging Conventions)
- Enumerate parts: workbook, worksheets, sharedStrings, styles, theme, relationships, vbaProject.bin, custom UI
- Provide lightweight model for targeted cell edits (avoid implementing full NPOI)

### 3.3 Modify
- Selective cell set/update with:
  - Shared string table update (if text) OR inline string option
  - Numeric / date types (store as double with appropriate style) 
  - Automatic creation of rows / cells if missing
- Avoid altering formula cells unless explicitly requested

### 3.4 Re‑Encrypt & Save
- Serialize modified OPC package to a memory stream
- Produce new encryption structure:
  - For Agile: generate new salt, IV, encryption key (unless preserving original parameters) OR re‑use original salt per requirement (decide policy)
  - Encrypt package stream in 4096‑byte blocks (per spec) with AES
  - Write updated `EncryptionInfo` and `EncryptedPackage` streams
- Ensure Excel opens without repair; test with multiple Excel versions (prefer 2016+)

### 3.5 Validation
- Post‑write: attempt to decrypt with password and open (self‑check)
- Structural checks: number of sheets, presence of vbaProject.bin, sharedStrings count difference only if new strings
- Optional hash of each original non‑edited binary part to ensure untouched integrity (especially macro binary)

### 3.6 Diagnostics & Errors
Define exception types:
- `EncryptionFormatNotSupportedException`
- `InvalidPasswordException`
- `EncryptionIntegrityException`
- `UnsupportedCipherException`
- `MacroPreservationException`
Provide verbose logging hooks (interface) but default silent unless enabled.

## 4. Technical Gaps & Tasks
| Category | Gap | Task Ideas | Priority |
|----------|-----|-----------|----------|
| Spec Research | Precise Agile + Standard encryption spec details | Collect MS-OFFCRYPTO sections (Agile, ECMA-376) | High |
| Binary Parsing | `EncryptionInfo` parser | Implement structured reader with unit tests vs fixtures | High |
| KDF | Agile password key derivation | Implement SHA1/SHA512 chain & spin loop | High |
| Crypto | AES block encryption (CBC) + HMAC (if needed) | Use BCL `Aes.Create()`; ensure padding per spec | High |
| Container Build | Write `EncryptedPackage` stream assembly | Stream pipeline writer (block processing) | High |
| Decrypt Flow | Full password validation | Round‑trip decrypt test harness | High |
| OPC Handling | Open decrypted ZIP from stream | Use `Package.Open(memoryStream)` | High |
| Cell Edit API | Minimal edit abstraction | Provide `SetCellValue(sheetName, row, col, object)` | Medium |
| Shared Strings | Add / reuse shared strings | Implement table loader & append logic | Medium |
| Macro Preservation | Copy binary parts untouched | Explicit part list & integrity hash check | Medium |
| Testing | Encrypted fixture generator | Use Excel to generate sample encrypted .xlsm for baseline | High |
| Test Automation | CI test matrix | Password verify + modification assert | Medium |
| Security Hygiene | Key zeroization | Overwrite arrays after use | Low |
| Performance | Streaming encryption (avoid large memory) | Blockwise transform + temp file fallback | Medium |
| API Design | Public surface definition | Draft `EncryptedMacroWorkbook` class contract | Medium |
| Documentation | Developer spec | Author DESIGN_ENCRYPTION.md | Medium |

## 5. Proposed Minimal Viable API (Draft)
```csharp
public sealed class EncryptedMacroWorkbook : IDisposable
{
    public static EncryptedMacroWorkbook Open(string path, string password);
    public IReadOnlyList<string> SheetNames { get; }
    public object? GetCell(string sheetName, int rowIndex, int columnIndex);
    public void SetCell(string sheetName, int rowIndex, int columnIndex, object? value);
    public void SaveAs(string outputPath, string? newPassword = null, EncryptionOptions? options = null);
}
```
Supporting types:
- `EncryptionOptions` (cipher, hash, spinCount override, preserveSalt)
- `EncryptionDiagnostics` (timings, block counts, integrity flags)

## 6. Risks & Mitigations
| Risk | Impact | Mitigation |
|------|--------|-----------|
| Spec misinterpretation | Corrupt or unrecoverable files | Build incremental validator & compare with Excel output |
| Incomplete macro preservation | Broken VBA / lost signatures | Treat binary macro parts as opaque; hash verify pre/post |
| Performance on large files | High memory usage | Stream blocks, temp file fallback |
| Crypto mistakes | Security weakness / failure to open | Reproduce known test vectors; unit test KDF & encryption |
| Time to implement | Delivery delays | Parallelize: parsing, KDF, block writer, OPC edit |

## 7. Milestone Roadmap
1. Parsing & Decrypt (M1)
2. Modify Unencrypted Workbook (M2)
3. Re‑Encrypt Round Trip (M3)
4. Cell Edit API & Tests (M4)
5. Macro Integrity & Extended Validation (M5)
6. Optimization & Public Preview (M6)

## 8. Acceptance Criteria (Done Definition)
- Given encrypted .xlsm + password → Open API returns workbook object (sheets accessible)
- After SetCell operations → SaveAs(newFile) produces encrypted .xlsm that:
  - Opens in Excel with password; no repair dialog
  - Preserves macros (VBA project hash identical)
  - Updates only targeted cells
- Supports at least Agile AES 128; optionally extends to AES 256 if source used it
- 100% deterministic round‑trip when no modifications (binary equality of decrypted payload)
- Unit & integration tests pass on CI (Windows, maybe Linux for parse/edit without encrypt until cross-platform strategy defined)

## 9. Design Decisions (Resolved Former Open Questions)
- Password change on save: Supported (optional new password parameter). Default = reuse original.
- Salt / IV policy: Always generate fresh cryptographic salt & IV for re-encryption; fallback to original only if Excel interoperability issues detected (diagnostic mode).
- Encryption formats: Agile only initial implementation; legacy Standard RC4 explicitly out-of-scope (may throw `EncryptionFormatNotSupportedException`).
- Streaming/bulk edit API: Deferred; will design after core round-trip (tracked as enhancement backlog item).

## 10. Immediate Next Steps
1. Add DESIGN_ENCRYPTION.md skeleton with spec references
2. Implement `EncryptionInfo` binary parser + tests (fixture from real encrypted .xlsm)
3. Implement Agile KDF (derive key test vs known vector)
4. Decrypt to raw OPC stream and validate workbook open
5. Draft public API class (internal until stable) & simple cell read support

---
*This document will evolve as implementation proceeds.*
