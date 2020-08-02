Imports System.IO

Namespace ExifRemover
    Public Class JpegPatcher
        Public Function PatchAwayExif(inStream As Stream, outStream As Stream) As Stream
            Dim jpegHeader As Byte() = New Byte(1) {}
            jpegHeader(0) = CByte(inStream.ReadByte())
            jpegHeader(1) = CByte(inStream.ReadByte())
            If jpegHeader(0) = &HFF AndAlso jpegHeader(1) = &HD8 Then
                'check if it's a jpeg file
                SkipAppHeaderSection(inStream)
            End If
            outStream.WriteByte(&HFF)
            outStream.WriteByte(&HD8)

            Dim readCount As Integer
            Dim readBuffer As Byte() = New Byte(4095) {}
            While (InlineAssignHelper(readCount, inStream.Read(readBuffer, 0, readBuffer.Length))) > 0
                outStream.Write(readBuffer, 0, readCount)
            End While

            Return outStream
        End Function

        Private Sub SkipAppHeaderSection(inStream As Stream)
            Dim header As Byte() = New Byte(1) {}
            header(0) = CByte(inStream.ReadByte())
            header(1) = CByte(inStream.ReadByte())

            While header(0) = &HFF AndAlso (header(1) >= &HE0 AndAlso header(1) <= &HEF)
                Dim exifLength As Integer = inStream.ReadByte()
                exifLength = exifLength << 8
                exifLength = exifLength Or inStream.ReadByte()

                For i As Integer = 0 To exifLength - 3
                    inStream.ReadByte()
                Next
                header(0) = CByte(inStream.ReadByte())
                header(1) = CByte(inStream.ReadByte())
            End While
            inStream.Position -= 2
            'skip back two bytes
        End Sub
        Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
            target = value
            Return value
        End Function
    End Class
End Namespace

