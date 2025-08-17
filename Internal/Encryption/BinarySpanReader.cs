using System;
using System.Buffers.Binary;

namespace JFToolkit.EncryptedExcel.Internal.Encryption;

/// <summary>
/// Lightweight span-based reader for little-endian binary parsing.
/// </summary>
internal ref struct BinarySpanReader
{
    private ReadOnlySpan<byte> _span;
    private int _offset;

    public BinarySpanReader(ReadOnlySpan<byte> span)
    {
        _span = span;
        _offset = 0;
    }

    public int Remaining => _span.Length - _offset;
    public int Position => _offset;

    private ReadOnlySpan<byte> Slice(int count)
    {
        if (count < 0 || _offset + count > _span.Length)
            throw new EncryptionInfoCorruptException($"Unexpected end of data at {Position} (need {count} bytes)");
        var slice = _span.Slice(_offset, count);
        _offset += count;
        return slice;
    }

    public ushort ReadUInt16() => BinaryPrimitives.ReadUInt16LittleEndian(Slice(2));
    public uint ReadUInt32() => BinaryPrimitives.ReadUInt32LittleEndian(Slice(4));
    public int ReadInt32() => BinaryPrimitives.ReadInt32LittleEndian(Slice(4));
    public byte[] ReadBytes(int count) => Slice(count).ToArray();

    public string ReadLengthPrefixedUnicodeString()
    {
        // Some Agile segments contain length-prefixed UTF-16LE strings (length in bytes or chars differs by context)
        // Placeholder: implement once exact field encountered; for now throw if called.
        throw new NotImplementedException("String parsing not yet implemented");
    }
}
