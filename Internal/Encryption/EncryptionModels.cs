using System;

namespace JFToolkit.EncryptedExcel.Internal.Encryption;

/// <summary>
/// Parsed subset of Agile EncryptionInfo necessary for key derivation & decrypt.
/// Spec sections: MS-OFFCRYPTO 2.3.4 (Agile Encryption)
/// </summary>
internal sealed record AgileEncryptionInfo(
    ushort VersionMajor,
    ushort VersionMinor,
    uint Flags,
    string CipherAlgorithm,
    int CipherKeyBits,
    string HashAlgorithm,
    int HashSize,
    byte[] Salt,
    int SpinCount,
    byte[] EncryptedVerifier,
    byte[] EncryptedVerifierHash
);

/// <summary>
/// Indicates which high-level encryption flavor was detected.
/// </summary>
internal enum EncryptionFlavor
{
    Unknown = 0,
    Standard,
    Agile
}
