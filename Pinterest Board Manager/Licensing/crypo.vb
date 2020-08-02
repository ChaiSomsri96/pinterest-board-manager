Imports System.Text
Imports System.Security.Cryptography
Imports System.IO

Friend Class crypo
    Private _data As String
    Private _key As String
    Private _LicenseKey As String
    Private _MachID As String
    Private _LocalIP As String

    Friend Property Data() As String
        Get
            Return _data
        End Get
        Set(ByVal value As String)
            _data = value
        End Set
    End Property

    Friend Property Key As String
        Get
            Return _key
        End Get
        Set(ByVal value As String)
            _key = value
        End Set
    End Property

    Friend Property LicenseKey As String
        Get
            Return _LicenseKey
        End Get
        Set(ByVal value As String)
            _LicenseKey = value
        End Set
    End Property

    Friend Property MachID As String
        Get
            Return _MachID
        End Get
        Set(ByVal value As String)
            _MachID = value
        End Set
    End Property

    Friend Property LocalIP As String
        Get
            Return _LocalIP
        End Get
        Set(ByVal value As String)
            _LocalIP = value
        End Set
    End Property

    Friend Function Decrypt(ByVal str As String, ByVal key As String) As String
        Dim cipher As Byte() = Convert.FromBase64String(str)
        Dim btKey As Byte() = Encoding.ASCII.GetBytes(key)
        Dim decryptor As ICryptoTransform = New RijndaelManaged() With { _
            .Mode = CipherMode.ECB, _
            .Padding = PaddingMode.None _
        }.CreateDecryptor(btKey, Nothing)
        Dim ms As New MemoryStream(cipher)
        Dim cs As New CryptoStream(ms, decryptor, CryptoStreamMode.Read)
        Dim plain As Byte() = New Byte(cipher.Length - 1) {}
        Dim count As Integer = cs.Read(plain, 0, plain.Length)
        ms.Close()
        cs.Close()
        Return Encoding.UTF8.GetString(plain, 0, count)
    End Function

    Friend Function Encrypt(ByVal Message As String, ByVal pssw As String) As String
        Dim buffer As Byte()
        Dim encoding As New UTF8Encoding
        Dim bytes As Byte() = encoding.GetBytes(pssw)
        Dim managed As New RijndaelManaged With { _
            .Key = bytes, _
            .Mode = CipherMode.ECB, _
            .Padding = PaddingMode.PKCS7 _
        }
        Dim inputBuffer As Byte() = encoding.GetBytes(Message)
        Dim transform As ICryptoTransform = managed.CreateEncryptor
        Try
            buffer = managed.CreateEncryptor.TransformFinalBlock(inputBuffer, 0, inputBuffer.Length)
        Finally
            managed.Clear()
        End Try
        Return Convert.ToBase64String(buffer)
    End Function

End Class
