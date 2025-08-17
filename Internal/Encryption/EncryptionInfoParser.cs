using System;
using System.IO;

namespace JFToolkit.EncryptedExcel.Internal.Encryption;

/// <summary>
/// Parses Office EncryptionInfo stream (Agile focus). Standard encryption will raise not supported.
/// </summary>
internal static class EncryptionInfoParser
{
    public static (EncryptionFlavor Flavor, AgileEncryptionInfo? Agile) Parse(Stream encryptionInfoStream)
    {
        if (encryptionInfoStream == null) throw new ArgumentNullException(nameof(encryptionInfoStream));
        using var ms = new MemoryStream();
        encryptionInfoStream.CopyTo(ms);
        var data = ms.ToArray();
        if (data.Length < 8)
            throw new EncryptionInfoCorruptException("EncryptionInfo too short");

        var reader = new BinarySpanReader(data);
        var versionMinor = reader.ReadUInt16();
        var versionMajor = reader.ReadUInt16();
        var flags = reader.ReadUInt32();

        // Heuristic: Agile often stores an XML section after a fixed header (version >= 4?). We'll refine later.
        // For now treat all as Agile if versionMajor >= 4.
        if (versionMajor < 4)
        {
            // Likely Standard encryption
            throw new EncryptionFormatNotSupportedException($"Standard/legacy encryption (version {versionMajor}.{versionMinor}) not supported");
        }

        // Placeholder parsing: need to parse remaining XML structure. For now throw NotImplemented to signal incomplete implementation.
        throw new NotImplementedException("Agile EncryptionInfo parsing not implemented yet");
    }
}
