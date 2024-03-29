﻿Imports System.Security.Cryptography
Imports System.Management

Public Class clsCrypto

    Private TripleDes As New TripleDESCryptoServiceProvider

    Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()

        Dim sha1 As New SHA1CryptoServiceProvider

        ' Hash the key.
        Dim keyBytes() As Byte =
            System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)

        ' Truncate or pad the hash.
        ReDim Preserve hash(length - 1)
        Return hash
    End Function

    ''' <summary>
    ''' Initializes the symmetric crypto provider.
    ''' </summary>
    ''' <param name="key">The key to encode and decode data with</param>
    ''' <remarks></remarks>
    Sub New(ByVal key As String)
        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub

    Private Function EncryptData(ByVal plaintext As String) As String

        ' Convert the plaintext string to a byte array.
        Dim plaintextBytes() As Byte =
            System.Text.Encoding.Unicode.GetBytes(plaintext)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the encoder to write to the stream.
        Dim encStream As New CryptoStream(ms,
            TripleDes.CreateEncryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()

        ' Convert the encrypted stream to a printable string.
        Return Convert.ToBase64String(ms.ToArray)
    End Function

    Private Function DecryptData(ByVal encryptedtext As String) As String

        ' Convert the encrypted text string to a byte array.
        Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the decoder to write to the stream.
        Dim decStream As New CryptoStream(ms,
            TripleDes.CreateDecryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
        decStream.FlushFinalBlock()

        ' Convert the plaintext stream to a string.
        Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
    End Function

    ''' <summary>
    ''' Decrypts a cipher into a plain text representation
    ''' </summary>
    ''' <param name="cipher">The cipher to decrypt</param>
    ''' <returns>The decrypted plain text representation of the cipher entered</returns>
    Public Function Decrypt(cipher As String) As String
        Return DecryptData(cipher)
    End Function

    ''' <summary>
    ''' Encrypts plain text to cipher
    ''' </summary>
    ''' <param name="txt">The plain text to encrypt</param>
    ''' <returns>The encrpyted cipher representation of the plain text entered</returns>
    Public Function Encrypt(txt As String) As String
        Return EncryptData(txt)
    End Function

    ''' <summary>
    ''' generate login code from computer data
    ''' </summary>
    ''' <param name="_coding">encryted string</param>
    ''' <returns>login code dez / string</returns>
    Private Function generateCode(ByVal _coding As String) As Int64
        Dim _code As Int64
        For I As Integer = 1 To Len(_coding)
            _code += (Asc(Mid(_coding, I, 1)) * 915734)
        Next
        Return _code
    End Function

End Class
