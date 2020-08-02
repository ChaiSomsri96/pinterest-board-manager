Imports System.Security.Cryptography

Friend Class Randomizer(Of T)
    Implements IComparer(Of T)

    ''// Ensures different instances are sorted in different orders
    Private Shared Salter As New Random() ''// only as random as your seed
    Private Salt As Integer
    Friend Sub New()
        Salt = Salter.Next(Integer.MinValue, Integer.MaxValue)
    End Sub

    Private Shared sha As New SHA1CryptoServiceProvider()
    Private Function HashNSalt(ByVal x As Integer) As Integer
        Dim b() As Byte = sha.ComputeHash(BitConverter.GetBytes(x))
        Dim r As Integer = 0
        For i As Integer = 0 To b.Length - 1 Step 4
            r = r Xor BitConverter.ToInt32(b, i)
        Next

        Return r Xor Salt
    End Function

    Friend Function Compare(ByVal x As T, ByVal y As T) As Integer _
        Implements IComparer(Of T).Compare

        Return HashNSalt(x.GetHashCode()).CompareTo(HashNSalt(y.GetHashCode()))
    End Function
End Class
