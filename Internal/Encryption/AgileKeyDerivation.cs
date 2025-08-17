using System;
using System.Security.Cryptography;
using System.Text;

namespace JFToolkit.EncryptedExcel.Internal.Encryption;

/// <summary>
/// Agile password key derivation (placeholder). Will implement spin loop + hash chaining.
/// </summary>
internal static class AgileKeyDerivation
{
    public static byte[] DeriveKey(string password, byte[] salt, int spinCount, int keyBytes, string hashAlgorithm)
    {
        if (string.IsNullOrEmpty(password)) throw new ArgumentException("Password required", nameof(password));
        if (salt == null || salt.Length == 0) throw new ArgumentException("Salt required", nameof(salt));
        if (spinCount <= 0) throw new ArgumentOutOfRangeException(nameof(spinCount));
        if (keyBytes <= 0) throw new ArgumentOutOfRangeException(nameof(keyBytes));

        // Temporary naive placeholder: NOT SPEC-COMPLIANT.
        // Will be replaced with: H^spinCount( salt || UTF16LE(password) ) and final block derivation.
        using var sha1 = SHA1.Create();
        var pwd = Encoding.Unicode.GetBytes(password); // UTF-16LE
        var working = new byte[salt.Length + pwd.Length];
        Buffer.BlockCopy(salt, 0, working, 0, salt.Length);
        Buffer.BlockCopy(pwd, 0, working, salt.Length, pwd.Length);
        byte[] hash = sha1.ComputeHash(working);
        for (int i = 0; i < Math.Min(spinCount, 1000); i++) // cap for now to avoid runaway
        {
            var iteration = BitConverter.GetBytes(i);
            var concat = new byte[iteration.Length + hash.Length];
            Buffer.BlockCopy(iteration, 0, concat, 0, iteration.Length);
            Buffer.BlockCopy(hash, 0, concat, iteration.Length, hash.Length);
            hash = sha1.ComputeHash(concat);
        }
        if (hash.Length == keyBytes) return hash;
        if (hash.Length > keyBytes)
        {
            var trimmed = new byte[keyBytes];
            Buffer.BlockCopy(hash, 0, trimmed, 0, keyBytes);
            return trimmed;
        }
        // Expand (simple repeat) – placeholder only
        var expanded = new byte[keyBytes];
        for (int i = 0; i < keyBytes; i++) expanded[i] = hash[i % hash.Length];
        return expanded;
    }
}
