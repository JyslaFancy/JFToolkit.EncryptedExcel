using System;

namespace JFToolkit.EncryptedExcel.Internal.Encryption;

internal class InvalidPasswordException : Exception
{
    public InvalidPasswordException(string message) : base(message) { }
}

internal class EncryptionFormatNotSupportedException : Exception
{
    public EncryptionFormatNotSupportedException(string message) : base(message) { }
}

internal class EncryptionInfoCorruptException : Exception
{
    public EncryptionInfoCorruptException(string message) : base(message) { }
}

internal class UnsupportedCipherException : Exception
{
    public UnsupportedCipherException(string message) : base(message) { }
}

internal class EncryptionIntegrityException : Exception
{
    public EncryptionIntegrityException(string message) : base(message) { }
}
