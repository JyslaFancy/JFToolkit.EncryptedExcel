# DESIGN_ENCRYPTION (Agile OOXML .xlsm Password Workflow)

> Status: Draft (2.0.0-dev)

## 1. Objective
Implement cross‑platform (no Excel COM, no NPOI) ability to:
1. Decrypt password‑protected .xlsm (Agile encryption)
2. Modify limited cell values (targeted write scope)
3. Re‑encrypt + save preserving macros & structure

## 2. Scope (Phase 1)
Included:
- Agile encryption only (ECMA-376 / MS-OFFCRYPTO spec reference)
- AES 128/256 CBC
- SHA1 (initial) + optional SHA512 if source uses it
- Preserve all non-edited parts byte-for-byte (vbaProject.bin, relationships, styles, theme, drawings, pivot caches)
- Basic cell value updates (string, numeric, boolean, date) via sharedStrings table or inline strings

Excluded (Phase 1):
- Standard (legacy) encryption
- Changing macro project
- Digital signature/VBA signing preservation (strip or copy as-is)
- Cell formulas recalculation / formula editing
- Row styling / merged cells manipulation

## 3. High-Level Architecture
```
EncryptedFile (.xlsm)
│
├─ EncryptionInfo (binary)  <-- Parse => derive key
├─ EncryptedPackage (binary) -- decrypt --> Raw OPC (ZIP in memory)
│                                 │
│                                 ├─ [Content_Types].xml
│                                 ├─ xl/workbook.xml
│                                 ├─ xl/worksheets/sheetN.xml (target edits)
│                                 ├─ xl/sharedStrings.xml (maybe)
│                                 ├─ xl/styles.xml
│                                 ├─ xl/vbaProject.bin (opaque)
│                                 └─ _rels & xl/_rels/.rels
│
└─ Rebuild: modify XML parts -> serialize -> encrypt -> write new EncryptionInfo + EncryptedPackage
```

## 4. Key Components
| Component | Responsibility |
|-----------|---------------|
| `EncryptionInfoParser` | Read binary header, extract algorithms, salt, spinCount, key size |
| `AgileKeyDerivation` | Derive encryption key from password + salt + spinCount |
| `PackageDecryptor` | Produce decrypted ZIP stream with integrity validation |
| `OoxmlWorkbookEditor` | Minimal layer to read & modify sheets and shared strings |
| `PackageRebuilder` | Serialize updated OPC into single stream |
| `AgileEncryptor` | Encrypt stream block-wise into `EncryptedPackage` |
| `EncryptedMacroWorkbook` | Public façade orchestrating operations |
| `IntegrityValidator` | Compare hashes of preserved parts pre/post |

## 5. Detailed Flow
### 5.1 Decrypt
1. Open input file as compound stream (encrypted OOXML container is actually a flat file – treat as binary)
2. Locate & parse `EncryptionInfo` (Agile version 4/5 structures)
3. Derive key: `hash = H(salt + passwordUnicodeLE); repeat spinCount: hash = H(iterationIndex + hash); finalKey = Truncate(hash)`
4. Initialize AES (CBC) with IV from `KeyData`/block header
5. Decrypt `EncryptedPackage` payload
6. Validate decrypted first bytes are `50 4B 03 04` (ZIP local file header) else throw

### 5.2 Load + Modify
1. Open decrypted memory stream with `Package.Open`
2. Map target sheet(s) by name or index
3. For each edit:
   - Locate row element (create if missing)
   - Locate cell (e.g., `r="A1"`); create with proper `t` attribute
   - For string: either add to sharedStrings.xml and set cell to `s` index OR add inline string (decision: prefer sharedStrings for consistency)
4. Update sharedStrings count & uniqueCount if new strings added

### 5.3 Re-Encrypt
1. Serialize updated package to memory stream (or temp file for large size)
2. Generate new salt + IV (unless preserving original; configurable)
3. Derive key with new salt (if changed) & provided password
4. Encrypt package in blocks (Agile uses segmenting—validate spec: typical is whole-stream AES; confirm)
5. Write new `EncryptionInfo` structure & encrypted payload to output

### 5.4 Validation Path (Optional)
After write, optionally decrypt produced file using same logic to ensure integrity (dev/diagnostic mode only).

## 6. Data Structures
```csharp
internal sealed record EncryptionInfo(
    ushort VersionMajor,
    ushort VersionMinor,
    uint Flags,
    string CipherAlgorithm,
    int CipherKeySize,
    string HashAlgorithm,
    int HashSize,
    byte[] Salt,
    int SpinCount,
    byte[] EncryptedVerifier,
    byte[] EncryptedVerifierHash
);
```

## 7. Error Taxonomy
| Error | Condition |
|-------|-----------|
| `InvalidPasswordException` | Key derivation succeeds but verifier mismatch |
| `UnsupportedEncryptionException` | Non-Agile or unsupported cipher/hash |
| `EncryptionInfoCorruptException` | Structural parse failure |
| `MacroIntegrityException` | vbaProject.bin altered unexpectedly |
| `WorkbookEditException` | XML edit failure |

## 8. Security Considerations
- Zero sensitive buffers (salt, key, intermediate hashes) after use (`Array.Clear` + `CryptographicOperations.ZeroMemory`).
- Avoid writing decrypted temp to disk unless size > threshold; if written, ensure secure delete (best-effort overwrite).
- Parameter validation: reject extremely large spinCount (> 10 million) to prevent DoS.

## 9. Performance
- Stream transformations to avoid large LOH allocations.
- Reuse `SHA1`/`SHA512` instances via `IncrementalHash` where possible.
- Optional span-based parsing for `EncryptionInfo`.

## 10. Testing Strategy
| Test Type | Purpose |
|-----------|---------|
| Unit: Parser | Known binary fixtures -> expected fields |
| Unit: KDF | Reproduce documented test vector outputs |
| Unit: Editor | Edit cell; verify XML delta |
| Integration: Round Trip | Decrypt->Edit->Encrypt->Open in Excel manually/CI smoke |
| Integrity: Macro Hash | Hash vbaProject.bin pre/post no-op save |
| Negative: Wrong Password | Ensure `InvalidPasswordException` |

## 11. Incremental Delivery Plan
- Phase A: Parser + KDF + verify wrong/right password
- Phase B: Full decrypt & open as OPC
- Phase C: Minimal cell edit + shared string update
- Phase D: Re-encrypt (no edits) binary-equivalent test
- Phase E: Re-encrypt with edits
- Phase F: Hardening + diagnostics + public preview

## 12. Open Issues
- Confirm Agile block size & segmentation details vs entire stream encryption
- Determine if HMAC integrity check required for chosen sub-version
- Decide salt/IV preservation vs regeneration policy (default: regenerate)
- Multi-threading benefits for spinCount loops? (Probably not – ensure deterministic and safe)

## 13. References
- ECMA-376 Part 2: Open Packaging Conventions & Encryption
- MS-OFFCRYPTO: [https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto]
- Historical reverse-engineering notes (to be gathered)

---
*End of draft*
