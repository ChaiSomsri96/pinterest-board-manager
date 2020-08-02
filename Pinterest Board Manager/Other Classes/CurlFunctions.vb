Imports System.Net
Imports System.Text
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

Friend Class CurlFunctions
    Friend AccountID As String
    Friend Proxy As String
    Friend Port As Integer
    Friend ProxyUser As String
    Friend ProxyPass As String
    Friend Useragent As String
    Friend cookies As New CookieContainer
    Private MaxTimeout As Integer = My.Settings.MaxTimeout * 1000
    Friend ActiveDatabase As String = ""

    Friend Sub SetUseragent()
        Randomize()
        If My.Settings.Useragents = Nothing Then Throw New Exception("No Useragents In Settings")
        Dim TmpAgents() As String = Split(My.Settings.Useragents, vbCrLf)
        Useragent = TmpAgents(CInt(Int(Rnd() * (TmpAgents.Count() - 1))))
    End Sub

    ' Set Account Values New Sub - Also Loads Account Cookies
    Friend Sub New(ByVal AccID As String, ByVal TmpActiveDatabase As String, Optional ByVal Prox As String = "", Optional ByVal Por As Integer = 0, Optional ByVal ProxyUse As String = "", Optional ByVal ProxyPas As String = "", Optional ByVal UAgent As String = "")
        AccountID = AccID
        Proxy = Prox
        Port = Por
        ProxyUser = ProxyUse
        ProxyPass = ProxyPas
        ActiveDatabase = TmpActiveDatabase
        Try
            cookies = LoadCookies()
        Catch ex As Exception
            Throw New Exception("Failed To Load Cookies For Account: " & AccID & " - " & ex.Message)
        End Try
        If UAgent <> "" Then
            Useragent = UAgent
        Else
            SetUseragent()
        End If
        ' Request Settings
        ServicePointManager.ServerCertificateValidationCallback = (Function(sender, certificate, chain, sslPolicyErrors) True)
        System.Net.ServicePointManager.Expect100Continue = False
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
    End Sub

    ' Dispose Sub - Saves Cookies Upon Class Instance Closure / Program Exit
    Friend Overridable Sub Dispose()
        If cookies.Count > 0 Then
            Try
                SaveCookie()
            Catch ex As Exception
                ' Throw New Exception("Failed To Save Cookies For Account: " & AccountID & " - " & ex.Message)
            End Try
        End If
    End Sub

    ' Load Cookies From DAT File Function
    Function LoadCookies() As CookieContainer
        Dim retrievedCookies As CookieContainer = Nothing
        If AccountID <> "TempCookie" AndAlso AccountID <> "0" Then
            Dim bform As New BinaryFormatter
            Dim Folder As String = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\" & "\cookies\"
            If Not (Directory.Exists(Folder)) Then
                Directory.CreateDirectory(Folder)
            End If
            AccountID = AccountID.Replace("?", "").Replace("@", "").Replace("*", "").Replace("/", "")
            Dim DatFile As String = Folder & AccountID & ".dat"
            Dim fileExists As Boolean = File.Exists(DatFile)
            If fileExists = False Then
                File.Create(DatFile).Dispose()
                retrievedCookies = New CookieContainer
            Else
                Using s As Stream = File.OpenRead((DatFile))
                    If Not s.Length = 0 Then
                        retrievedCookies = DirectCast(bform.Deserialize(s), CookieContainer)
                    Else
                        retrievedCookies = New CookieContainer
                    End If
                End Using
            End If
        Else
            retrievedCookies = New CookieContainer
        End If
        Return retrievedCookies
    End Function

    ' Save Cookies To DAT File Sub
    Private Sub SaveCookie()
        If AccountID <> "TempCookie" Then
            Dim formatter = New BinaryFormatter()
            Dim Folder As String = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\" & "\cookies\"
            If Not (Directory.Exists(Folder)) Then
                Directory.CreateDirectory(Folder)
            End If
            AccountID = AccountID.Replace("?", "").Replace("@", "").Replace("*", "").Replace("/", "")
            Dim DatFile As String = Folder & AccountID & ".dat"
            Dim fileExists As Boolean = File.Exists(DatFile)
            If fileExists = False Then File.Create(DatFile).Dispose()
            Dim file__1 As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), DatFile)
            Using s As Stream = File.Create(file__1)
                formatter.Serialize(s, cookies)
            End Using
        End If
    End Sub

    ' Return Content Function
    Friend Function Return_Content(ByVal URL As String, Optional ByVal Referrer As String = "", Optional ByVal CheckRateLimit As Boolean = False) As String
        Threading.Thread.Sleep(My.Settings.PageLoadDelay * 1000)
        Dim Content As String = Nothing
        For i = 0 To My.Settings.MaxRetry
            Try
                Dim request As HttpWebRequest = CType(Net.WebRequest.Create(URL), HttpWebRequest)
                ' Set Proxy Details
                If Proxy <> Nothing AndAlso Proxy <> "0" AndAlso Proxy <> "" Then
                    Dim myProxy As New WebProxy(Proxy, Port)
                    If ProxyUser <> "0" AndAlso ProxyUser <> "" Then myProxy.Credentials = New NetworkCredential(ProxyUser, ProxyPass)
                    request.Proxy = myProxy
                End If
                ' Request Settings
                request.Method = "GET"
                request.ServicePoint.Expect100Continue = False
                request.KeepAlive = True
                request.AllowAutoRedirect = True
                request.Timeout = MaxTimeout
                request.CookieContainer = cookies
                If Referrer <> "" Then request.Referer = Referrer
                request.UserAgent = Useragent
                Dim response As System.Net.HttpWebResponse
                Try
                    response = CType(request.GetResponse(), HttpWebResponse)
                Catch ex As WebException
                    If ex.Message.Contains("400") OrElse ex.Message.Contains("403") OrElse ex.Message.Contains("500") OrElse ex.Message.Contains("503") Then
                        response = CType(ex.Response, HttpWebResponse)
                        Dim postreqreader As New StreamReader(response.GetResponseStream())
                        Content = postreqreader.ReadToEnd
                        Return Content
                    Else
                        Throw New Exception(ex.Message)
                    End If
                End Try
                If response.StatusCode = HttpStatusCode.OK Then
                    Dim responseStream As New StreamReader(response.GetResponseStream())
                    Content = responseStream.ReadToEnd()
                End If
                ' Add Cookies
                If response.Cookies.Count > 0 Then cookies.Add(response.Cookies)
                ' Close Connection
                response.Close()
                ' Exit Loop (Successful Connection)
                Exit For
            Catch ex As Exception
                If i = My.Settings.MaxRetry Then Throw New Exception(ex.Message) Else Continue For
            End Try
        Next
        If Content = Nothing Then Return "" Else Return Content
    End Function

    ' POST Content Function
    Friend Function Post_Content(ByVal URL As String, ByVal Data As String, Optional ByVal Referrer As String = "", Optional ByVal MultiPart As Boolean = False, Optional ByVal CheckRateLimit As Boolean = False, Optional ByVal SkipRetries As Boolean = False, Optional ByVal AcceptHeader As String = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8") As String
        Threading.Thread.Sleep(My.Settings.PageLoadDelay * 1000)
        Dim Content As String = Nothing
        Dim Retries As Integer = My.Settings.MaxRetry
        If SkipRetries = True Then Retries = 1
        For i = 0 To Retries
            Try
                Dim encoding As New UTF8Encoding
                Dim byteData As Byte() = encoding.GetBytes(Data)
                ' New Request
                Dim request As HttpWebRequest = DirectCast(WebRequest.Create(URL), HttpWebRequest)
                ' Set Proxy Details
                If Proxy <> Nothing AndAlso Proxy <> "0" AndAlso Proxy <> "" Then
                    Dim myProxy As New WebProxy(Proxy, Port)
                    myProxy.Credentials = New NetworkCredential(ProxyUser, ProxyPass)
                    request.Proxy = myProxy
                End If
                ' Request Settings
                request.Method = "POST"
                request.ServicePoint.Expect100Continue = False
                request.KeepAlive = True
                request.AllowAutoRedirect = False
                request.Timeout = MaxTimeout
                request.CookieContainer = cookies
                If Referrer.Contains(".com") OrElse URL.Contains(".com") Then
                    request.Headers.Add("Origin", "https://www.pinterest.com")
                Else
                    request.Headers.Add("Origin", "https://www.pinterest.com")
                End If
                ' Headers
                request.UserAgent = Useragent
                request.Accept = AcceptHeader
                ' request.Headers.Add("Accept-Encoding: gzip, deflate")
                request.Headers.Add("Accept-Language: en-GB,en-US;q=0.9,en;q=0.8")
                request.ContentLength = byteData.Length
                If MultiPart = False Then
                    request.ContentType = "application/x-www-form-urlencoded; charset=UTF-8"
                Else
                    request.ContentType = "multipart/form-data"
                End If
                If Referrer <> "" Then request.Referer = Referrer
                ' Read Stream
                Dim postreqstream As Stream = request.GetRequestStream()
                postreqstream.Write(byteData, 0, byteData.Length)
                postreqstream.Close()
                Dim postresponse As HttpWebResponse
                Try
                    postresponse = CType(request.GetResponse(), HttpWebResponse)
                Catch ex As WebException
                    If ex.Message.Contains("400") OrElse ex.Message.Contains("403") OrElse ex.Message.Contains("500") OrElse ex.Message.Contains("503") OrElse ex.Message.Contains("401") OrElse ex.Message.Contains("409") Then
                        postresponse = CType(ex.Response, HttpWebResponse)
                        Dim postreqreader As New StreamReader(postresponse.GetResponseStream())
                        Content = postreqreader.ReadToEnd
                        Return Content
                    Else
                        Throw New Exception(ex.Message)
                    End If
                End Try
                ' Check HTTP Status
                If postresponse.StatusCode = Net.HttpStatusCode.OK Then
                    Dim postreqreader As New StreamReader(postresponse.GetResponseStream())
                    Content = postreqreader.ReadToEnd
                End If
                If postresponse.StatusCode = Net.HttpStatusCode.Redirect OrElse postresponse.StatusCode = Net.HttpStatusCode.Moved OrElse postresponse.StatusCode = Net.HttpStatusCode.MovedPermanently Then
                    ' Add Cookies
                    If postresponse.Cookies.Count > 0 Then cookies.Add(postresponse.Cookies)
                    URL = postresponse.Headers("Location")
                    If CheckRateLimit = True AndAlso URL.Contains("ratelimit") Then
                        Return "Rate Limit"
                    End If
                    Content = Return_Content(URL, Referrer)
                End If
                ' Close Connection
                postresponse.Close()
                ' Exit Loop (Successful Connection)
                Exit For
            Catch ex As Exception
                If i = My.Settings.MaxRetry Then Throw New Exception(ex.Message) Else Continue For
            End Try
        Next
        ' Return HTML Content
        If Content = Nothing Then Return "" Else Return Content
    End Function

    ' PINTEREST - NEW XMLHttpRequest_Post_Request (With Data)
    Friend Function XMLHttpRequest_Post_Request(ByVal URL As String, ByVal Data As String, ByVal CSRFToken As String, Optional ByVal Referrer As String = "", Optional ByVal NewAppHeader As Boolean = True, Optional ByVal NewAppVersion As String = "", Optional ByVal AppState As String = "") As String
        Threading.Thread.Sleep(My.Settings.PageLoadDelay * 1000)
        Dim Content As String = Nothing
        For i = 0 To My.Settings.MaxRetry
            Try
                Dim encoding As New UTF8Encoding
                Dim byteData As Byte() = encoding.GetBytes(Data)
                ' New Request
                Dim request As HttpWebRequest = DirectCast(WebRequest.Create(URL), HttpWebRequest)
                ' Set Proxy Details
                If Proxy <> Nothing AndAlso Proxy <> "" AndAlso Proxy <> "0" Then
                    Dim myProxy As New WebProxy(Proxy, Port)
                    myProxy.Credentials = New NetworkCredential(ProxyUser, ProxyPass)
                    request.Proxy = myProxy
                End If
                ' Request Settings
                request.AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate
                request.ServicePoint.Expect100Continue = False
                request.Method = "POST"
                request.KeepAlive = True
                request.Timeout = MaxTimeout
                request.CookieContainer = cookies
                request.UserAgent = Useragent
                request.ContentLength = byteData.Length
                request.Accept = "application/json, text/javascript, */*; q=0.01"
                ' request.Headers.Add("Accept-Encoding", "gzip, deflate")
                request.Headers.Add("Accept-Language", "en-gb,en;q=0.5")
                request.ContentType = "application/x-www-form-urlencoded; charset=UTF-8"
                If AppState <> "" Then request.Headers.Add("X-Pinterest-AppState", AppState)
                If CSRFToken <> Nothing Then request.Headers.Add("X-CSRFToken", CSRFToken)
                If NewAppHeader = True Then request.Headers.Add("X-NEW-APP", "1")
                If NewAppVersion <> "" Then request.Headers.Add("X-APP-VERSION", NewAppVersion)
                request.Headers.Add("X-Requested-With", "XMLHttpRequest")
                request.AllowAutoRedirect = False
                If Referrer <> "" Then request.Referer = Referrer
                ' Read Stream
                Dim postreqstream As Stream = request.GetRequestStream()
                postreqstream.Write(byteData, 0, byteData.Length)
                postreqstream.Close()
                Dim postresponse As HttpWebResponse
                Try
                    postresponse = DirectCast(request.GetResponse(), HttpWebResponse)
                Catch ex As WebException
                    If ex.Message.Contains("400") OrElse ex.Message.Contains("401") OrElse ex.Message.Contains("403") OrElse ex.Message.Contains("500") OrElse ex.Message.Contains("503") OrElse ex.Message.Contains("409") OrElse ex.Message.Contains("429") Then
                        postresponse = CType(ex.Response, HttpWebResponse)
                        Dim postreqreader As New StreamReader(postresponse.GetResponseStream())
                        Content = postreqreader.ReadToEnd
                        Return Content
                    Else
                        Throw New Exception(ex.Message)
                    End If
                End Try
                ' Check HTTP Status
                If postresponse.StatusCode = Net.HttpStatusCode.OK Then
                    Dim postreqreader As New StreamReader(postresponse.GetResponseStream())
                    Content = postreqreader.ReadToEnd
                End If
                ' Add Cookies
                If postresponse.Cookies.Count > 0 Then cookies.Add(postresponse.Cookies)
                ' Close Connection
                postresponse.Close()
                ' Exit Loop (Successful Connection)
                Exit For
            Catch ex As Exception
                If i = My.Settings.MaxRetry Then Throw New Exception(ex.Message) Else Continue For
            End Try
        Next
        ' Return HTML Content
        If Content = Nothing Then Return "" Else Return Content
    End Function

    ' PINTEREST - XMLHttpRequest_Post_Request_NoData (Without Data)
    Friend Function XMLHttpRequest_Post_Request_NoData(ByVal URL As String, ByVal CSRFToken As String, Optional ByVal Referrer As String = "") As String
        Threading.Thread.Sleep(My.Settings.PageLoadDelay * 1000)
        Dim Content As String = Nothing
        For i = 0 To My.Settings.MaxRetry
            Try
                Dim encoding As New UTF8Encoding
                ' New Request
                Dim request As HttpWebRequest = DirectCast(WebRequest.Create(URL), HttpWebRequest)
                ' Set Proxy Details
                If Proxy <> Nothing AndAlso Proxy <> "" AndAlso Proxy <> "0" Then
                    Dim myProxy As New WebProxy(Proxy, Port)
                    myProxy.Credentials = New NetworkCredential(ProxyUser, ProxyPass)
                    request.Proxy = myProxy
                End If
                ' Request Settings
                request.ServicePoint.Expect100Continue = False
                request.Method = "POST"
                request.KeepAlive = False
                request.Timeout = MaxTimeout
                request.CookieContainer = cookies
                request.UserAgent = Useragent
                request.ContentLength = 0
                request.Accept = "application/json, text/javascript, */*; q=0.01"
                ' request.Headers.Add("Accept-Encoding", "gzip, deflate")
                request.Headers.Add("Accept-Language", "en-gb,en;q=0.5")
                request.ContentType = "application/x-www-form-urlencoded"
                request.Headers.Add("X-CSRFToken", CSRFToken)
                request.Headers.Add("X-Requested-With", "XMLHttpRequest")
                request.AllowAutoRedirect = False
                If Referrer <> "" Then request.Referer = Referrer
                ' Read Stream
                Dim postreqstream As Stream = request.GetRequestStream()
                postreqstream.Close()
                Dim postresponse As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
                ' Check HTTP Status
                If postresponse.StatusCode = Net.HttpStatusCode.OK Then
                    Dim postreqreader As New StreamReader(postresponse.GetResponseStream())
                    Content = postreqreader.ReadToEnd
                End If
                ' Add Cookies
                If postresponse.Cookies.Count > 0 Then cookies.Add(postresponse.Cookies)
                ' Close Connection
                postresponse.Close()
                ' Exit Loop (Successful Connection)
                Exit For
            Catch ex As Exception
                If i = My.Settings.MaxRetry Then Throw New Exception(ex.Message) Else Continue For
            End Try
        Next
        ' Return HTML Content
        If Content = Nothing Then Return "" Else Return Content
    End Function

    ' PINTEREST - NEW XMLHttpRequest_Get_Request
    Friend Function XMLHttpRequest_Get_Request(ByVal URL As String, ByVal CSRFToken As String, ByVal Referrer As String, Optional ByVal Accept As String = "", Optional ByVal NewAppHeader As Boolean = True, Optional ByVal NewAppVersion As String = "", Optional ByVal AppState As String = "") As String
        Threading.Thread.Sleep(My.Settings.PageLoadDelay * 1000)
        Dim Content As String = Nothing
        For i = 0 To My.Settings.MaxRetry
            Try
                Dim request As Net.HttpWebRequest = CType(Net.WebRequest.Create(URL), HttpWebRequest)
                ' Set Proxy Details
                If Proxy <> Nothing AndAlso Proxy <> "0" AndAlso Proxy <> "" Then
                    Dim myProxy As New WebProxy(Proxy, Port)
                    myProxy.Credentials = New NetworkCredential(ProxyUser, ProxyPass)
                    request.Proxy = myProxy
                End If
                ' Request Settings
                request.Method = "GET"
                request.ServicePoint.Expect100Continue = False
                request.KeepAlive = False
                request.AllowAutoRedirect = True
                request.Timeout = MaxTimeout
                request.CookieContainer = cookies
                If Accept <> "" Then request.Accept = Accept
                If CSRFToken <> "" Then request.Headers.Add("X-CSRFToken", CSRFToken)
                If NewAppHeader = True Then request.Headers.Add("X-NEW-APP", "1")
                If NewAppVersion <> "" Then request.Headers.Add("X-APP-VERSION", NewAppVersion)
                If AppState <> "" Then request.Headers.Add("X-Pinterest-AppState", AppState)
                request.Headers.Add("X-Requested-With", "XMLHttpRequest")
                If Referrer <> "" Then request.Referer = Referrer
                request.UserAgent = Useragent
                request.AllowAutoRedirect = True
                Dim response As Net.HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                If response.StatusCode = Net.HttpStatusCode.OK Then
                    Dim responseStream As IO.StreamReader = New IO.StreamReader(response.GetResponseStream())
                    Content = responseStream.ReadToEnd()
                End If
                ' Add Cookies
                If response.Cookies.Count > 0 Then cookies.Add(response.Cookies)
                ' Close Connection
                response.Close()
                ' Exit Loop (Successful Connection)
                Exit For
            Catch ex As Exception
                If i = My.Settings.MaxRetry Then Throw New Exception(ex.Message) Else Continue For
            End Try
        Next
        ' Return HTML Content
        If Content = Nothing Then Return "" Else Return Content
    End Function

    ' Download Image From Web
    Friend Sub GetImage(ByVal URL As String, ByVal FilePath As String)
        Threading.Thread.Sleep(My.Settings.PageLoadDelay * 1000)
        Dim TmpImage As System.Drawing.Image = Nothing
        Dim request As Net.HttpWebRequest = CType(Net.WebRequest.Create(URL), HttpWebRequest)
        ' Set Proxy Details
        If Proxy <> Nothing AndAlso Proxy <> "0" AndAlso Proxy <> "" Then
            Dim myProxy As New WebProxy(Proxy, Port)
            myProxy.Credentials = New NetworkCredential(ProxyUser, ProxyPass)
            request.Proxy = myProxy
        End If
        request.Method = "GET"
        request.ServicePoint.Expect100Continue = False
        request.KeepAlive = False
        request.AllowAutoRedirect = True
        request.Timeout = MaxTimeout
        request.CookieContainer = cookies
        request.UserAgent = Useragent
        request.AllowAutoRedirect = True
        Dim response As Net.HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
        If response.StatusCode = Net.HttpStatusCode.OK Then
            TmpImage = System.Drawing.Image.FromStream(response.GetResponseStream)
        End If
        ' Add Cookies
        If response.Cookies.Count > 0 Then cookies.Add(response.Cookies)
        ' Close Connection
        response.Close()
        ' Save Image To Hard Drive
        TmpImage.Save(FilePath)
        TmpImage = Nothing
    End Sub

    ' Download Image From Web & Remove Exif
    Friend Sub GetImage_RemoveExif(ByVal URL As String, ByVal FilePath As String, Optional ByVal Referrer As String = "", Optional ByVal AcceptHeader As String = "image/png,image/*;q=0.8,*/*;q=0.5")
        Threading.Thread.Sleep(My.Settings.PageLoadDelay * 1000)
        ' Get Folder
        Dim Folder As String = Path.GetDirectoryName(FilePath)
        ' Download Image
        Dim TmpImage As System.Drawing.Image = Nothing
        Dim request As Net.HttpWebRequest = CType(Net.WebRequest.Create(URL), HttpWebRequest)
        ' Set Proxy Details
        If Proxy <> Nothing AndAlso Proxy <> "0" AndAlso Proxy <> "" Then
            Dim myProxy As New WebProxy(Proxy, Port)
            myProxy.Credentials = New NetworkCredential(ProxyUser, ProxyPass)
            request.Proxy = myProxy
        End If
        request.UserAgent = Useragent
        request.AllowWriteStreamBuffering = False
        request.Method = "GET"
        request.AllowAutoRedirect = True
        request.KeepAlive = False
        request.Timeout = MaxTimeout
        request.CookieContainer = cookies
        request.Accept = AcceptHeader
        If Referrer <> "" Then request.Referer = Referrer
        Dim response As Net.HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
        If response.StatusCode = Net.HttpStatusCode.OK Then
            Dim HTTPStream As Stream = response.GetResponseStream()
            Dim memoryStream As New MemoryStream()
            HTTPStream.CopyTo(memoryStream)
            memoryStream.Position = 0
            HTTPStream = memoryStream
            ' Remove EXIF Data
            Dim Patcher As New ExifRemover.JpegPatcher
            Dim new_memoryStream As New MemoryStream()
            new_memoryStream = Patcher.PatchAwayExif(memoryStream, new_memoryStream)
            ' Save Image
            TmpImage = System.Drawing.Image.FromStream(new_memoryStream)
        End If
        ' Add Cookies
        If response.Cookies.Count > 0 Then cookies.Add(response.Cookies)
        ' Save Image To Hard Drive
        Try
            If FilePath.ToLower().EndsWith(".gif") Then
                TmpImage.Save(FilePath, System.Drawing.Imaging.ImageFormat.Gif)
            Else
                TmpImage.Save(FilePath, System.Drawing.Imaging.ImageFormat.Jpeg)
            End If
            ' Close Connection
            response.Close()
        Catch ex As Exception
            Throw New Exception("Failed To Save Image To Hard Drive: " & ex.Message)
        End Try
        TmpImage = Nothing
    End Sub

    ' Get_URL Function (No Proxy)
    Friend Function Get_URL(ByVal URL As String, Optional ByVal Referrer As String = "") As String
        Threading.Thread.Sleep(My.Settings.PageLoadDelay * 1000)
        Dim Content As String = Nothing
        For i = 0 To My.Settings.MaxRetry
            Try
                Dim request As Net.HttpWebRequest = CType(Net.WebRequest.Create(URL), HttpWebRequest)
                ' Set Proxy Details
                If Proxy <> Nothing AndAlso Proxy <> "0" AndAlso Proxy <> "" Then
                    Dim myProxy As New WebProxy(Proxy, Port)
                    myProxy.Credentials = New NetworkCredential(ProxyUser, ProxyPass)
                    request.Proxy = myProxy
                End If
                ' Request Settings
                request.Method = "GET"
                request.KeepAlive = False
                request.AllowAutoRedirect = True
                request.Timeout = MaxTimeout
                request.CookieContainer = cookies
                If Referrer <> "" Then request.Referer = Referrer
                request.UserAgent = Useragent
                request.AllowAutoRedirect = True
                Dim response As Net.HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                If response.StatusCode = Net.HttpStatusCode.OK Then
                    Dim responseStream As IO.StreamReader = New IO.StreamReader(response.GetResponseStream())
                    Content = responseStream.ReadToEnd()
                End If
                ' Add Cookies
                If response.Cookies.Count > 0 Then cookies.Add(response.Cookies)
                ' Close Connection
                response.Close()
                ' Exit Loop (Successful Connection)
                Exit For
            Catch ex As Exception
                If i = My.Settings.MaxRetry Then Throw New Exception(ex.Message) Else Continue For
            End Try
        Next
        ' Return HTML Content
        If Content = Nothing Then Return "" Else Return Content
    End Function

    ' PINTEREST - Upload Profile Image
    Friend Function UploadProfileImage(ByVal FilePath As String, ByVal CSRFToken As String, ByVal Username As String) As String
        Threading.Thread.Sleep(My.Settings.PageLoadDelay * 1000)
        Dim ImageData As System.Drawing.Image
        Dim fs As New System.IO.FileStream(FilePath, System.IO.FileMode.Open)
        ImageData = Image.FromStream(fs)

        Dim Content As String = ""
        Dim boundary As String = "---------------------------" & (DateTime.Now.Ticks.ToString("x").Substring(0, DateTime.Now.Ticks.ToString("x").Length() - 3))

        Dim builder As New StringBuilder()

        builder.Append("--" & boundary & vbCrLf & "Content-Disposition: form-data; name=""img""; filename=""" & FilePath.Split(CChar("\")).Last() & """")
        builder.Append(vbCrLf)
        builder.Append("Content-Type: image/jpeg")
        builder.Append(vbCrLf & vbCrLf)

        Dim headerBytes As Byte() = Encoding.ASCII.GetBytes(builder.ToString())

        Dim ms As New MemoryStream()
        ImageData.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)

        Dim imageBytes As Byte() = ms.ToArray()

        Dim footerBytes As Byte() = Encoding.ASCII.GetBytes((vbCrLf & "--" & boundary & "--" & vbCrLf))

        Dim postBytes As Byte() = New Byte(headerBytes.Length + imageBytes.Length + (footerBytes.Length - 1)) {}
        fs.Close()
        Buffer.BlockCopy(headerBytes, 0, postBytes, 0, headerBytes.Length)
        Buffer.BlockCopy(imageBytes, 0, postBytes, headerBytes.Length, imageBytes.Length)
        Buffer.BlockCopy(footerBytes, 0, postBytes, headerBytes.Length + imageBytes.Length, footerBytes.Length)
        Dim req As HttpWebRequest = DirectCast(WebRequest.Create("https://www.pinterest.com/upload-image/?img=" & FilePath.Split(CChar("\")).Last()), HttpWebRequest)
        ' Set Proxy Details
        If Proxy <> Nothing AndAlso Proxy <> "0" AndAlso Proxy <> "" Then
            Dim myProxy As New WebProxy(Proxy, Port)
            myProxy.Credentials = New NetworkCredential(ProxyUser, ProxyPass)
            req.Proxy = myProxy
        End If
        req.Method = "POST"
        req.UserAgent = Useragent
        req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        req.Headers.Add("X-Requested-With", "XMLHttpRequest")
        req.Headers.Add("X-File-Name", FilePath.Split(CChar("\")).Last())
        req.Headers.Add("X-CSRFToken", CSRFToken)
        req.ServicePoint.Expect100Continue = False
        req.KeepAlive = True
        req.Referer = "https://www.pinterest.com/" & URLEncode(Username.ToLower()) & "/"
        req.CookieContainer = cookies
        req.ContentType = "multipart/form-data; boundary=" & boundary
        req.ContentLength = postBytes.Length
        req.AllowAutoRedirect = False

        ' Read Stream
        Dim postreqstream As Stream
        Dim postresponse As HttpWebResponse
        postreqstream = req.GetRequestStream()
        postreqstream.Write(postBytes, 0, postBytes.Length)
        postreqstream.Close()
        Try
            postresponse = DirectCast(req.GetResponse(), HttpWebResponse)
        Catch ex As WebException
            If ex.Message.Contains("400") OrElse ex.Message.Contains("403") OrElse ex.Message.Contains("500") Then
                postresponse = CType(ex.Response, HttpWebResponse)
                Dim postreqreader As New StreamReader(postresponse.GetResponseStream())
                Content = postreqreader.ReadToEnd
                Return Content
            Else
                Throw New Exception(ex.Message)
            End If
        End Try
        ' Check HTTP Status
        If postresponse.StatusCode = Net.HttpStatusCode.OK Then
            Dim postreqreader As New StreamReader(postresponse.GetResponseStream())
            Content = postreqreader.ReadToEnd
        End If
        ' Close Connection
        postresponse.Close()
        'fs.Close()
        Return Content
    End Function

    ' PINTEREST - Upload Image
    Friend Function UploadImage(ByVal FilePath As String, ByVal token As String, ByVal Username As String) As String
        Threading.Thread.Sleep(My.Settings.PageLoadDelay * 1000)
        ' Variables
        Dim ImageData As System.Drawing.Image
        ' Save Image To Memory
        Try
            Dim ImageMemoryStream As New MemoryStream(My.Computer.FileSystem.ReadAllBytes(FilePath))
            ImageData = New Bitmap(ImageMemoryStream)
        Catch ex As Exception
            Throw New Exception("Failed To Load / Reformat Image: " & ex.Message & " (" & FilePath & ")")
        End Try
        Dim ms As New MemoryStream()
        ImageData.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
        ' Variables
        Dim Content As String = ""
        Dim boundary As String = "---------------------------" & DateTime.Now.Ticks.ToString("x")
        Dim footer_boundary As String = boundary & "--"
        Dim req As HttpWebRequest = DirectCast(WebRequest.Create("https://www.pinterest.com/upload-image/?img=" & FilePath.Split(CChar("\")).Last()), HttpWebRequest)
        ' Set Proxy Details
        If Proxy <> Nothing AndAlso Proxy <> "0" AndAlso Proxy <> "" Then
            Dim myProxy As New WebProxy(Proxy, Port)
            myProxy.Credentials = New NetworkCredential(ProxyUser, ProxyPass)
            req.Proxy = myProxy
        End If
        ' Set Request Settings
        System.Net.ServicePointManager.Expect100Continue = False
        req.Method = "POST"
        req.UserAgent = Useragent
        req.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        req.Headers.Add("X-Requested-With", "XMLHttpRequest")
        req.Headers.Add("X-File-Name", FilePath.Split(CChar("\")).Last())
        req.Headers.Add("X-CSRFToken", token)
        req.ServicePoint.Expect100Continue = False
        req.KeepAlive = True
        req.Referer = "https://www.pinterest.com/"
        req.CookieContainer = cookies
        req.ContentType = "multipart/form-data; boundary=" & boundary
        req.AllowAutoRedirect = False

        Dim builder As New StringBuilder()

        builder.Append("--" & boundary & vbCrLf & "Content-Disposition: form-data; name=""img""; filename=""" & FilePath.Split(CChar("\")).Last() & """")
        builder.Append(vbCrLf)
        builder.Append("Content-Type: image/jpeg")
        builder.Append(vbCrLf & vbCrLf)

        Dim headerBytes As Byte() = Encoding.ASCII.GetBytes(builder.ToString())

        Dim imageBytes As Byte() = ms.ToArray()

        Dim footerBytes As Byte() = Encoding.ASCII.GetBytes((vbCrLf & "--" & boundary & "--" & vbCrLf))

        Dim postBytes As Byte() = New Byte(headerBytes.Length + imageBytes.Length + (footerBytes.Length - 1)) {}

        req.ContentLength = postBytes.Length
        req.CookieContainer = cookies
        Buffer.BlockCopy(headerBytes, 0, postBytes, 0, headerBytes.Length)
        Buffer.BlockCopy(imageBytes, 0, postBytes, headerBytes.Length, imageBytes.Length)
        Buffer.BlockCopy(footerBytes, 0, postBytes, headerBytes.Length + imageBytes.Length, footerBytes.Length)
        ms.Close()
        ' Read Stream
        Dim postreqstream As Stream
        Dim postresponse As HttpWebResponse
        postreqstream = req.GetRequestStream()
        postreqstream.Write(postBytes, 0, postBytes.Length)
        postreqstream.Close()
        Try
            postresponse = DirectCast(req.GetResponse(), HttpWebResponse)
        Catch ex As WebException
            If ex.Message.Contains("400") OrElse ex.Message.Contains("401") OrElse ex.Message.Contains("403") OrElse ex.Message.Contains("500") Then
                postresponse = CType(ex.Response, HttpWebResponse)
                Dim postreqreader As New StreamReader(postresponse.GetResponseStream())
                Content = postreqreader.ReadToEnd
                Return Content
            Else
                Throw New Exception(ex.Message)
            End If
        End Try
        ' Check HTTP Status
        If postresponse.StatusCode = Net.HttpStatusCode.OK Then
            Dim postreqreader As New StreamReader(postresponse.GetResponseStream())
            Content = postreqreader.ReadToEnd
        End If
        ' Close Connection
        postresponse.Close()
        'fs.Close()
        Return Content
    End Function

End Class