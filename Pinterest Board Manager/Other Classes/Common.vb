Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Security.Cryptography

Friend Module Common

    ' Create Strong Password
    Friend Function CreateStrongPassword() As String
        Dim NewPassword As String = ""
        While NewPassword = ""
            NewPassword = GetMixtureAllChars(GetRandomNumber(8, 11))
            Dim upper As Integer
            Dim lower As Integer
            Dim numbers As Integer
            For i = 0 To NewPassword.Length - 1
                If Char.IsLetter(NewPassword(i)) Then
                    If Char.IsUpper(NewPassword(i)) Then
                        upper += 1
                    Else
                        lower += 1
                    End If
                ElseIf Char.IsNumber(NewPassword(i)) Then
                    numbers += 1
                Else
                    ' Other Chars
                End If
            Next
            ' Check Ok
            If upper < 1 OrElse lower < 1 OrElse numbers < 1 Then
                NewPassword = ""
            End If
        End While
        Return NewPassword
    End Function

    ' Get Random Mixture Function
    Friend Function GetMixtureAllChars(ByVal Length As Integer) As String
        Randomize()
        Dim Letters() = {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
        Dim Phrase As String = ""
        For i = 0 To (Length - 1)
            Phrase = Phrase & Letters(GetRandomNumber(0, (Letters.Count() - 1)))
        Next
        Return Phrase
    End Function

    ' TextFileToList
    Friend Sub TextFileToList(ByRef Data As List(Of String), ByVal FilePath As String)
        ' Load Accounts
        If File.Exists(FilePath) Then
            Dim objReader As New System.IO.StreamReader(FilePath)
            Dim Line As String
            Do While objReader.Peek() <> -1
                Line = objReader.ReadLine()
                If Line = Nothing OrElse Line = "" Then Continue Do
                Line.Replace("	", "")
                Data.Add(Line)
            Loop
            objReader.Close()
        End If
    End Sub

    ' Check For Alphabetic Characters
    Friend Function CheckForAlphaCharacters(ByVal StringToCheck As String) As Boolean
        For i = 0 To StringToCheck.Length - 1
            If Char.IsLetter(StringToCheck.Chars(i)) Then
                Return True
            End If
        Next
        Return False
    End Function

    ' Base 64 Decode
    Friend Function ConvertFromBase64(ByVal Input As String) As String
        Dim B As Byte() = System.Convert.FromBase64String(Input)
        Return System.Text.Encoding.UTF8.GetString(B)
    End Function

    ' MD5 Function
    Friend Function MD5(ByVal input As String) As String
        Dim x As New System.Security.Cryptography.MD5CryptoServiceProvider()
        Dim bs As Byte() = System.Text.Encoding.UTF8.GetBytes(input)
        bs = x.ComputeHash(bs)
        Dim s As New System.Text.StringBuilder()
        For Each b As Byte In bs
            s.Append(b.ToString("x2").ToLower())
        Next
        Dim hash As String = s.ToString()
        Return hash
    End Function

    ' Get Session Token From Cookies
    Friend Function GetSessionToken(ByVal a_Process As CurlFunctions, Optional ByVal DomainName As String = "pinterest.com") As String
        Dim token As String = ""
        For Each cook In a_Process.cookies.GetCookies(New Uri("https://www." & DomainName & "/"))
            Try
                If cook.ToString().Contains("csrftoken=") Then
                    token = cook.ToString().Replace("csrftoken=", "")
                    Exit For
                End If
            Catch ex As Exception
                ' Ignore Issues
            End Try
        Next
        If token = Nothing Then
            For Each cook In a_Process.cookies.GetCookies(New Uri("http://www." & DomainName & "/"))
                Try
                    If cook.ToString().Contains("csrftoken=") Then
                        token = cook.ToString().Replace("csrftoken=", "")
                        Exit For
                    End If
                Catch ex As Exception
                    ' Ignore Issues
                End Try
            Next
        End If
        Return token
    End Function

    ' Get Random Mixture Function
    Friend Function GetMixtureLetter(ByVal Length As Integer) As String
        Randomize()
        Dim Letters() = {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
        Dim Phrase As String = ""
        For i = 0 To (Length - 1)
            Phrase = Phrase & Letters(GetRandomNumber(0, (Letters.Count() - 1)))
        Next
        Return Phrase
    End Function

    ' Resize & Format Image For Instagram
    Friend Function ResizeWatermark(ByVal FilePath As String, ByVal WatermarkText As String, ByVal newColor As Color, ByVal wmSize As Integer, ByVal wmFontStyle As FontStyle, ByVal wmFontFamily As String, ByVal XPositioning As String, ByVal YPositioning As String) As Image
        Try
            ' Load Image
            Dim ImageMemoryStream As New MemoryStream(My.Computer.FileSystem.ReadAllBytes(FilePath))
            Dim OriginalImage As New Bitmap(ImageMemoryStream)
            ' Calculate Size
            Dim newImg As New Bitmap(OriginalImage, OriginalImage.Width, OriginalImage.Height) '' blank canvas
            ' Create New Bitmap
            Dim bmp As New Drawing.Bitmap(OriginalImage.Width, OriginalImage.Height)
            Dim grap As Drawing.Graphics = Drawing.Graphics.FromImage(bmp)
            grap.Clear(Drawing.Color.Black)
            Dim g As Graphics = Graphics.FromImage(bmp)
            ' Calculate Points To Insert Resized Image
            Dim InsertX As Integer = 0
            Dim InsertY As Integer = 0
            ' Add Resized Image To Canvas
            g.DrawImage(newImg, New Point(InsertX, InsertY))
            ImageMemoryStream.Close()
            ' Create Watermark
            Dim sf As New StringFormat
            Dim padding As Integer = 5
            ' Set X Positioning
            Select Case XPositioning
                Case "Left"
                    sf.Alignment = StringAlignment.Near
                Case "Right"
                    sf.Alignment = StringAlignment.Far
                Case "Center"
                    sf.Alignment = StringAlignment.Center
            End Select
            ' Set Y Positioning
            Select Case YPositioning
                Case "Top"
                    sf.LineAlignment = StringAlignment.Near
                Case "Middle"
                    sf.LineAlignment = StringAlignment.Center
                Case "Bottom"
                    sf.LineAlignment = StringAlignment.Far
            End Select
            ' Set Color
            Dim brush As SolidBrush = New SolidBrush(newColor)
            ' Measure the size of the image for alignment of text
            Dim wmrect As New Rectangle(New Point(padding, padding), bmp.Size)
            wmrect.Height = wmrect.Height - padding
            wmrect.Width = wmrect.Width - padding
            ' Set Font
            Dim font As Font = New Font(wmFontFamily, wmSize, wmFontStyle, GraphicsUnit.Pixel)
            Dim color As Color = newColor
            ' Add Watermark
            g.DrawString(WatermarkText, font, brush, wmrect, sf)
            ' Return New Image
            Return bmp
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ' Randomize Array
    Friend Sub RandomizeArray(ByRef items() As String)
        Dim max_index As Integer = items.Length - 1
        Dim rnd As New Random
        For i As Integer = 0 To max_index - 1
            ' Pick an item for position i.
            Dim j As Integer = rnd.Next(i, max_index + 1)

            ' Swap them.
            Dim temp As String = items(i)
            items(i) = items(j)
            items(j) = temp
        Next i
    End Sub

    ' Get Milliseconds
    Friend Function GetMilliSeconds() As Long
        Dim span As TimeSpan = (DateTime.UtcNow - New DateTime(&H7B2, 1, 1, 0, 0, 0))
        Return Convert.ToInt64(span.TotalMilliseconds)
    End Function

    ' Create Unix MicroTime Function
    Friend Function GetUnixMicroTime(ByVal dates As DateTime) As String
        If dates.IsDaylightSavingTime = True Then
            dates = DateAdd(DateInterval.Hour, -1, dates)
        End If
        Dim origin As DateTime = New DateTime(1970, 1, 1, 0, 0, 0, 0)
        Dim diff As TimeSpan = dates - origin
        Return CStr(diff.TotalSeconds).Replace(".", "")
    End Function

    ' Valid URL Function
    Friend Function ValidateURL(ByVal URL As String) As Boolean
        If URL.Contains(" ") Then Return False
        If Not URL.StartsWith("http://") AndAlso Not URL.StartsWith("https://") Then Return False
        If Not URL.Contains(".") Then Return False
        ' If Not Regex.IsMatch(URL, "^((ht|f)tp(s?)\:\/\/|~/|/)?([\w]+:\w+@)?([a-zA-Z]{1}([\w\-]+\.)+([\w]{2,5}))(:[\d]{1,5})?((/?\w+/)+|/?)(\w+\.[\w]{3,4})?((\?\w+=\w+)?(&\w+=\w+)*)?") Then Return False
        Return True
    End Function

    ' Detect Pinterest Error Message
    Friend Function DetectPinterestError(ByVal Content As String) As String
        Dim ErrorMsg As String = Nothing
        If Content.Contains("""message"":""") Then
            ErrorMsg = GetBetween(Content, """message"":""", """")
            If ErrorMsg.Contains("ul class") Then Return Nothing
        End If
        Return ErrorMsg
    End Function

    ' Get Random Number Function
    Dim objRandom As New System.Random(CType(System.DateTime.Now.Ticks Mod System.Int32.MaxValue, Integer))
    Friend Function GetRandomNumber(Optional ByVal Low As Integer = 1, Optional ByVal High As Integer = 100) As Integer
        Return objRandom.Next(Low, High + 1)
    End Function

    ' Get Random Letter Function
    Friend Function GetRandomLetter() As String
        Randomize()
        Dim Letters() = {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"}
        Return Letters(CInt(Rnd() * (Letters.Count() - 1)))
    End Function

    ' URL Encode Function
    Friend Function URLEncode(ByVal URL As String) As String
        If URL = Nothing OrElse URL = "" Then Return ""
        Return Uri.EscapeDataString(URL)
    End Function

    ' GetBetween Function
    Friend Function GetBetween(ByVal haystack As String, ByVal needle As String, ByVal needle_two As String) As String
        Dim istart As Integer = InStr(haystack, needle)
        If istart > 0 Then
            ' Dim istop As Integer = InStr(istart, haystack, needle_two)
            Dim istop As Integer = InStr(istart + Len(needle), haystack, needle_two)
            If istop > 0 Then
                Try
                    Dim value As String = haystack.Substring(istart + Len(needle) - 1, istop - istart - Len(needle))
                    Return value
                Catch ex As Exception
                    Return ""
                End Try
            End If
        ElseIf haystack.Contains("RETRY_GETBETWEEN") = False Then
            ' Re-Attempt With No Spaces
            haystack = "RETRY_GETBETWEEN" & vbCrLf & haystack.Replace(" ", "")
            Return GetBetween(haystack, needle.Replace(" ", ""), needle_two.Replace(" ", ""))
        End If
        Return ""
    End Function

    ' GetAllStringsBetween Function
    Friend Function GetAllStringsBetween(ByVal Haystack As String, ByVal StartSearch As String, ByVal EndSearch As String, Optional ByVal ScrapeInt As Integer = 1) As String()
        Try
            If StartSearch.Contains("(.*?)") Then ScrapeInt += 1
            Dim rx As New Regex(StartSearch & "(.+?)" & EndSearch)
            Dim mc As MatchCollection = rx.Matches(Haystack.Trim())
            Dim FoundStrings(mc.Count) As String
            Dim i As Integer = 0
            For Each m As Match In mc
                If m.Groups(ScrapeInt).Value.ToString() <> Nothing Then
                    FoundStrings(i) = m.Groups(ScrapeInt).Value.ToString()
                    i += 1
                End If
            Next
            ' Check New Content
            If FoundStrings.Count() < 2 AndAlso Haystack.Contains("RETRY_GETBETWEEN") = False Then
                ' Re-Attempt With No Spaces
                Haystack = "RETRY_GETBETWEEN" & vbCrLf & Haystack.Replace(" ", "")
                Return GetAllStringsBetween(Haystack, StartSearch.Replace(" ", ""), EndSearch.Replace(" ", ""))
            End If
            ' Return Array
            Return FoundStrings
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ' Strip Tags Function
    Friend Function StripTags(ByVal html As String) As String
        If html = Nothing OrElse html.Trim() = Nothing OrElse html.Trim() = "" Then Return html
        Return Regex.Replace(html, "<.*?>", "")
    End Function

    ' Log_Show_Error Function
    Dim ErrorLogLock As New Object
    Friend Sub Log_Show_Error(ByVal Content As String, ByVal SaveFileAs As String, ByVal ActiveDatabase As String)
        SyncLock ErrorLogLock
            Dim strFile As String = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\error_logs\" & "ErrorLog_ECODE_" & SaveFileAs & ".txt"
            Dim fileExists As Boolean = File.Exists(strFile)
            Using sw As New StreamWriter(File.Open(strFile, FileMode.OpenOrCreate))
                sw.WriteLine(If(fileExists, Content, Content))
            End Using
            ' System.Diagnostics.Process.Start(strFile)
        End SyncLock
    End Sub

    ' Tokenize Function
    Function Tokenize(ByVal SourceString As String) As String
        Dim ReturnString As New StringBuilder
        Dim I As Integer = 0
        Randomize()
        Do While I < SourceString.Length
            Dim c As Char = SourceString.Chars(I)
            Select Case c
                Case ("{"c)
                    ReturnString.Append(Tokenize(SourceString.Substring(I + 1)))
                    Exit Do
                Case ("}"c)
                    Dim Options As String() = ReturnString.ToString().Split("|"c)
                    ReturnString = New StringBuilder
                    ReturnString.Append(Options(CType(Math.Floor(Rnd() * CType(Options.Length, Single)), Integer)))
                    If I < SourceString.Length - 1 Then
                        SourceString = SourceString.Substring(I + 1)
                        I = -1
                    Else
                        SourceString = ""
                    End If
                Case Else
                    ReturnString.Append(c)
            End Select
            I += 1
        Loop
        Return ReturnString.ToString()
    End Function

    ' New Spintax Parse Function
    Friend Function spintaxParse(rand As Random, s As String) As String
        If s.Contains("{"c) Then
            Dim closingBracePosition As Integer = s.IndexOf("}"c)
            Dim openingBracePosition As Integer = closingBracePosition

            While Not s(openingBracePosition).Equals("{"c)
                openingBracePosition -= 1
            End While

            Dim spintaxBlock As String = s.Substring(openingBracePosition, closingBracePosition - openingBracePosition + 1)

            Dim items As String() = spintaxBlock.Substring(1, spintaxBlock.Length - 2).Split("|"c)

            s = s.Replace(spintaxBlock, items(rand.[Next](items.Length)))

            Return spintaxParse(rand, s)
        Else
            Return s
        End If
    End Function

End Module