Imports System.IO
Imports System.Text
Imports System.Threading

Friend Class ThreadManager

#Region "Class Variables"
    ' Class Variables
    Private m_MainWindow As Form1
    Friend Campaign_Account As String
    Friend Campaign_BoardName As String
    Friend Campaign_Status As String
    Friend Campaign_Keywords As String
    Friend Campaign_Website As String
    Friend Campaign_Delay As Integer
    Friend Campaign_Inuse As String
    Friend Campaign_PinsCreated As Integer
    Friend NextPostTime As DateTime
    Friend ActiveDatabase As String

#End Region

#Region "Class Functions"
    ' New Sub
    Friend Sub New(ByRef MainWindow As Form1)
        m_MainWindow = MainWindow
    End Sub
    ' Update Successful Stats
    Private Sub IncreaseSuccessful()
        m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Pins Created", "+1")
        ' Campaign_PinsCreated = CInt(Campaign_PinsCreated) + 1
        ' m_MainWindow.UpdateDGV(Campaign_Account, Campaign_BoardName, "Pins Created", CType(Campaign_PinsCreated, String))
    End Sub
    ' Close Thread Sub
    Private Sub CloseThread(Optional ByVal CampaignDisabled As Boolean = False)
        Try
            If m_MainWindow.BoardManager_StopProcess = True Then
                m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Status", "Not Running")
            Else
                If CampaignDisabled = True Then
                    m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Status", "Task Disabled")
                Else
                    m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Status", "Awaiting Timer", True)
                End If
            End If
            m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Next Post Time", CStr(DateTime.Now.AddSeconds(CDbl((Campaign_Delay * 60)))))
            m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Inuse", "0")
            If CampaignDisabled = True Then
                m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Status", "Task Disabled")
            Else
                If m_MainWindow.BoardManager_StopProcess = True Then
                    m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Status", "Not Running")
                Else
                    m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Status", "Awaiting Timer", True)
                End If
            End If
            ' Reduce Thread Count
            m_MainWindow.ThreadsRunning -= 1
        Catch ex As Exception
            Log_Show_Error(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, System.Reflection.MethodInfo.GetCurrentMethod().Name, ActiveDatabase)
            MessageBox.Show("A Fatal Error Occured Closing A Thread - " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Main Processes"
    ' Start New Thread Sub
    Friend Sub RunProcess()
        Try
            ' Check For Free Thread
            Dim FoundFreeThread As Boolean = False
            SyncLock m_MainWindow.ThreadStarter_SyncLock
                If m_MainWindow.ThreadsRunning < My.Settings.MaxThreads Then
                    FoundFreeThread = True
                    m_MainWindow.ThreadsRunning += 1
                End If
            End SyncLock
            ' Wait For Free Thread
            If FoundFreeThread = False Then
                Dim WaitingMsgSent As Boolean = False
                While m_MainWindow.ThreadsRunning >= My.Settings.MaxThreads
                    If WaitingMsgSent = False Then
                        m_MainWindow.UpdateDGV(Campaign_Account, Campaign_BoardName, "Status", "Awaiting Free Thread")
                        m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Status", "Awaiting Free Thread")
                        WaitingMsgSent = True
                    End If
                    Thread.Sleep(2000)
                    SyncLock m_MainWindow.ThreadStarter_SyncLock
                        If m_MainWindow.ThreadsRunning >= My.Settings.MaxThreads Then
                            Continue While
                        Else
                            m_MainWindow.ThreadsRunning += 1
                            Exit While
                        End If
                    End SyncLock
                End While
            End If
            ' Already Running In A Thread, Now Call Starter Function
            m_MainWindow.UpdateDGV(Campaign_Account, Campaign_BoardName, "Status", "Starting Thread")
            m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Status", "Starting Thread")
            MainRunProcess()
        Catch ex As Exception
            Log_Show_Error(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, System.Reflection.MethodInfo.GetCurrentMethod().Name, ActiveDatabase)
            MessageBox.Show("A Fatal Error Occured With RunProcess - " & ex.Message)
            CloseThread(False)
        End Try
    End Sub
    ' Main Process
    Private Sub MainRunProcess()
        Try
            Randomize()
            ' Set Account Variables
            Dim Web_Username As String = Campaign_Account
            Dim Web_FriendlyUsername As String = ""
            Dim Web_UserID As String = ""
            Dim Web_Password As String = ""
            Dim Web_Proxy As String = ""
            Dim Web_Port As String = "0"
            Dim Web_ProxyUsername As String = ""
            Dim Web_ProxyPassword As String = ""
            Dim AccountActiveStatus As String = ""
            Dim FoundAccount As Boolean = False
            ' Wrap In Locker For Inuse Variable
            SyncLock m_MainWindow.AccountUpdaterLock
                ' Find Account In Main Form
                For Each TmpAcc() As String In m_MainWindow.Accounts
                    If TmpAcc(0) = Web_Username Then
                        Web_Password = TmpAcc(1).Trim()
                        Web_Proxy = TmpAcc(2).Trim()
                        Web_Port = TmpAcc(3).Trim()
                        Web_ProxyUsername = TmpAcc(4).Trim()
                        Web_ProxyPassword = TmpAcc(5).Trim()
                        AccountActiveStatus = TmpAcc(6).Trim()
                        FoundAccount = True
                        Exit For
                    End If
                Next
                ' Confirm Found Account Ok
                If FoundAccount = False Then
                    m_MainWindow.AddMessage("Failed To Find Account In System: " & Web_Username & " - Task Disabled")
                    ' Close Thread & Disable Campaign
                    CloseThread(True)
                    Exit Sub
                End If
                ' Check Account Active
                If AccountActiveStatus = "0" Then
                    m_MainWindow.AddMessage("Account Is Disabled: " & Web_Username & " - Task Disabled")
                    ' Close Thread & Disable Campaign
                    CloseThread(True)
                    Exit Sub
                ElseIf AccountActiveStatus = "2" Then
                    m_MainWindow.AddMessage("Account Is Locked: " & Web_Username & " - Task Disabled")
                    ' Close Thread & Disable Campaign
                    CloseThread(True)
                    Exit Sub
                End If
            End SyncLock
            ' Lookup Board ID
            Dim BoardID As String = ""
            Try
                BoardID = m_MainWindow.GetBoardId(Campaign_Account, Campaign_BoardName)
            Catch ex As Exception
                m_MainWindow.AddMessage("Failed To Find Board ID For Account: " & Web_Username & " - " & ex.Message & " - Task Disabled")
                ' Close Thread & Disable Campaign
                CloseThread(True)
                Exit Sub
            End Try
            ' Start New HTTP Session
            Dim a_Process As New CurlFunctions(Web_Username, ActiveDatabase, Web_Proxy, Web_Port, Web_ProxyUsername, Web_ProxyPassword)
            ' Login To Account
            Try
                m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Status", "Logging In To Account")
                m_MainWindow.Login(a_Process, Web_FriendlyUsername, Web_Password, Web_Username, Web_UserID)
            Catch ex As Exception
                m_MainWindow.AddMessage("Failed To Login To Account: " & Web_Username & " - " & ex.Message & " - Task Disabled")
                ' Close Thread & Disable Campaign
                CloseThread(True)
                Exit Sub
            End Try
            ' Lookup Stats
            Try
                Dim TotalImpressions As String = m_MainWindow.LookupStats(a_Process, Web_UserID)
                m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Total Impressions", TotalImpressions)
            Catch ex As Exception
                m_MainWindow.AddMessage("Failed To Lookup Total Impressions Count For Account: " & Web_Username & " - " & ex.Message)
            End Try
            ' Clear Notifications
            Try
                m_MainWindow.ClearNotifications(a_Process)
            Catch ex As Exception
                ' Ignore Issues
            End Try
            ' Select Random Keyword
            Dim TmpKeyword As String = Split(Campaign_Keywords, ",")(Rnd() * (Split(Campaign_Keywords, ",").Count() - 1))
            ' Scrape Pins
            Dim ScrapedData As New List(Of String)
            Try
                m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Status", "Scraping Pins For Keyword: " & TmpKeyword)
                ScrapePins(a_Process, TmpKeyword, ScrapedData)
            Catch ex As Exception
                m_MainWindow.AddMessage("Failed To Scrape Pins For Keyword:  " & TmpKeyword & " Using Account: " & Web_Username & " - " & ex.Message)
                GoTo CloseThisThread
            End Try
            ' Select Random Pin
            Dim FoundPin As Boolean = False
            Dim RandPin As String = ""
            ' Add To Used List (Thread Safe)
            SyncLock m_MainWindow.UsedListCheckLocker
                Dim AlreadyBeingUsedFinder As Boolean = True
                While AlreadyBeingUsedFinder = True
                    If ScrapedData.Count() < 1 Then
                        m_MainWindow.AddMessage("Scraped No New Pins To Use For Keyword: " & TmpKeyword & " Using Account: " & Web_Username)
                        GoTo CloseThisThread
                    Else
                        Dim RandInt As Integer = Rnd() * (ScrapedData.Count() - 1)
                        RandPin = ScrapedData(RandInt)
                        ScrapedData.RemoveAt(RandInt)
                        If m_MainWindow.AlreadyUsedPinIDs.Contains(RandPin) = False Then Exit While
                    End If
                End While
                m_MainWindow.AlreadyUsedPinIDs.Add(RandPin)
                WriteToFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\cache\" & "AlreadyUsedPinIDs.txt", RandPin)
            End SyncLock
            ' Set Pin Variables
            Dim PinID As String = RandPin
            Dim PinTitle As String = ""
            Dim PinLink As String = ""
            Dim PinImageLink As String = ""
            Dim PinDescription As String = ""
            ' Get Pin Information
            If m_MainWindow.BoardManager_StopProcess = True Then GoTo CloseThisThread
            Try
                m_MainWindow.GetPinInformation(a_Process, PinID, PinTitle, PinLink, PinImageLink, PinDescription)
            Catch ex As Exception
                m_MainWindow.AddMessage("Failed To Lookup Pin Information For Pin ID: " & PinID & " Using Account: " & Web_Username & " - " & ex.Message)
                ' Close Thread
                GoTo CloseThisThread
            End Try
            ' Download Image
            Dim OriginalImage_LocalFilePath As String = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\images\" & "." & Guid.NewGuid.ToString() & Split(PinImageLink, ".").Last()
            Try
                a_Process.GetImage_RemoveExif(PinImageLink, OriginalImage_LocalFilePath)
            Catch ex As Exception
                m_MainWindow.AddMessage("Failed To Download Pin Image For Pin ID: " & PinID & " Using Account: " & Web_Username & " - " & ex.Message)
                ' Close Thread
                GoTo CloseThisThread
            End Try
            ' Replace Variables
            Dim TmpPinLink As String = spintaxParse(New Random(), Campaign_Website.Replace("[PINID]", PinID))
            If TmpPinLink.Contains("[RAND") Then
                Dim RandInt As String = GetBetween(TmpPinLink, "[RAND", "]")
                If RandInt = Nothing OrElse Integer.TryParse(RandInt, 0) = False Then
                    m_MainWindow.AddMessage("Found Bad Integer Token In Link: " & TmpPinLink)
                    GoTo CloseThisThread
                End If
                ' Replace Token
                TmpPinLink = TmpPinLink.Replace("[RAND" & RandInt & "]", "")
                ' Replace Integers
                For i As Integer = 1 To CInt(RandInt)
                    TmpPinLink = TmpPinLink & GetRandomNumber(0, 9)
                Next
            End If
            ' Update Status
            m_MainWindow.UpdateDGV(Campaign_Account, Campaign_BoardName, "Status", "Posting New Pin")
            m_MainWindow.UpdateCampaignData(Campaign_Account, Campaign_BoardName, "Status", "Posting New Pin")
            m_MainWindow.AddMessage("Posting New Pin To: " & Web_Username)
            ' Post Pin
            If m_MainWindow.BoardManager_StopProcess = True Then GoTo CloseThisThread
            Dim UploadedPinURL As String = ""
            Try
                UploadedPinURL = m_MainWindow.UploadImage(a_Process, Web_Username, OriginalImage_LocalFilePath, BoardID, PinTitle, PinDescription, TmpPinLink)
            Catch ex As Exception
                m_MainWindow.AddMessage("Failed To Post Pin To Account: " & Web_Username & " - " & ex.Message)
                ' Log Temporary Error
                Log_Show_Error(Web_Username & vbCrLf & OriginalImage_LocalFilePath & vbCrLf & BoardID & vbCrLf & PinTitle & vbCrLf & PinDescription & vbCrLf & TmpPinLink & vbCrLf & ex.Message, "PinFailed_1", ActiveDatabase)
                ' Delete Image Files
                DeleteImageFiles(New String() {OriginalImage_LocalFilePath, OriginalImage_LocalFilePath})
                ' Close Thread
                GoTo CloseThisThread
            End Try
            ' Add Success
            IncreaseSuccessful()
            m_MainWindow.AddMessage("Successfully Posted New Pin To Account: " & Web_Username & " - " & UploadedPinURL)
            ' Save Cookies
            a_Process.Dispose()
            ' Delete Image Files
            DeleteImageFiles(New String() {OriginalImage_LocalFilePath})
CloseThisThread:
            ' Close Thread
            CloseThread()
        Catch ex As Exception
            Log_Show_Error(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, System.Reflection.MethodInfo.GetCurrentMethod().Name, ActiveDatabase)
            MessageBox.Show("A Fatal Error Occured In MainRunProcess - " & ex.Message)
            CloseThread()
        End Try
    End Sub
    ' Scrape Pins Sub
    Private Sub ScrapePins(ByRef a_Process As CurlFunctions, ByVal Keyword As String, ByRef ScrapedData As List(Of String))
        ' Set Info
        Dim MainDomain As String = "pinterest.com"
        ' Variables
        Dim URL As String = ""
        Dim Content As String
        Dim NextMaxID As String = "1"
        Dim MorePagesAvailable As Boolean = True
        Keyword = Keyword.Replace(" ", "+")
        ' Load Homepage
        Try
            Content = a_Process.Return_Content("https://www." & MainDomain & "/")
        Catch ex As Exception
            Throw New Exception("Failed To Load Homepage: " & ex.Message)
        End Try
        ' Scrape App Version
        Dim AppVersion As String = GetBetween(Content, """app_version"": """, """")
        If AppVersion = Nothing Then AppVersion = GetBetween(Content, """app_version"":""", """")
        ' Get Token
        Dim Token As String = GetSessionToken(a_Process)
        ' Confirm Token Exists
        If Token = Nothing Then
            Throw New Exception("Failed To Scrape Session Token")
        End If
        ' Check Keyword
        If Keyword.ToLower().Contains("user:") AndAlso Split(Keyword.ToLower(), "user:").Count() < 2 Then
            Throw New Exception("Invalid Format For User Search: " & Keyword & " - User Search Should Be - user:username")
        End If
        ' Loop Pages
        For i As Integer = 1 To My.Settings.MaxPages
            If m_MainWindow.BoardManager_StopProcess = True OrElse MorePagesAvailable = False OrElse NextMaxID = Nothing Then Exit For
            ' Load Search Page
            Try
                If Keyword.ToLower().Contains("user:") Then
                    If i = 1 Then
                        URL = "https://www." & MainDomain & "/resource/UserPinsResource/get/?source_url=%2F" & URLEncode(Split(Keyword.ToLower(), "user:")(1)) & "%2Fpins%2F&data=%7B%22options%22%3A%7B%22username%22%3A%22" & URLEncode(Split(Keyword.ToLower(), "user:")(1)) & "%22%7D%2C%22context%22%3A%7B%7D%7D&_=" & GetMilliSeconds()
                        Content = a_Process.XMLHttpRequest_Get_Request(URL, Token, "https://www." & MainDomain & "/search/?q=" & URLEncode(Keyword.Trim()), "", True, AppVersion, "active")
                    Else
                        URL = "https://www." & MainDomain & "/resource/UserPinsResource/get/?source_url=%2F" & URLEncode(Split(Keyword.ToLower(), "user:")(1)) & "%2Fpins%2F&data=%7B%22options%22%3A%7B%22username%22%3A%22" & URLEncode(Split(Keyword.ToLower(), "user:")(1)) & "%22%2C%22bookmarks%22%3A%5B%22" & URLEncode(NextMaxID) & "%22%5D%7D%2C%22context%22%3A%7B%7D%7D&_=" & GetMilliSeconds()
                        Content = a_Process.XMLHttpRequest_Get_Request(URL, Token, "https://www." & MainDomain & "/search/?q=" & URLEncode(Keyword.Trim()), "", True, AppVersion, "active")
                    End If
                Else
                    ' Search Pins Mode
                    If i = 1 Then
                        URL = "https://www." & MainDomain & "/resource/BaseSearchResource/get/?source_url=%2Fsearch%2Fpins%2F%3Fq%3D" & URLEncode(Keyword.Trim()) & "&data=%7B%22options%22%3A%7B%22filters%22%3Anull%2C%22query%22%3A%22" & URLEncode(Keyword.Trim()) & "%22%2C%22rs%22%3Anull%2C%22scope%22%3A%22pins%22%7D%2C%22context%22%3A%7B%7D%7D&_=" & GetMilliSeconds()
                        Content = a_Process.XMLHttpRequest_Get_Request(URL, Token, "https://www." & MainDomain & "/search/pins/?q=" & URLEncode(Keyword.Trim()), "", True, AppVersion, "active")
                    Else
                        URL = "https://www." & MainDomain & "/resource/SearchResource/get/?source_url=%2Fsearch%2Fpins%2F%3Fq%3D" & URLEncode(Keyword.Trim()) & "&data=%7B%22options%22%3A%7B%22bookmarks%22%3A%5B%22" & URLEncode(NextMaxID) & "%22%5D%2C%22query%22%3A%22" & URLEncode(Keyword.Trim()) & "%22%2C%22scope%22%3A%22pins%22%7D%2C%22context%22%3A%7B%7D%7D&_=" & GetMilliSeconds()
                        Content = a_Process.XMLHttpRequest_Get_Request(URL, Token, "https://www." & MainDomain & "/search/?q=" & URLEncode(Keyword.Trim()))
                    End If
                End If
            Catch ex As Exception
                If ex.Message.Contains("404") Then Exit For
                m_MainWindow.AddMessage("Failed To Load URL: " & ex.Message)
                Continue For
            End Try
            ' Scrape Next ID
            If Content <> Nothing Then
                NextMaxID = GetBetween(Content, """nextBookmark"": """, """")
                If NextMaxID = Nothing Then NextMaxID = GetBetween(Content, """bookmarks"":[""", """")
            End If
            ' Confirm Available
            If NextMaxID = Nothing OrElse NextMaxID = "-end-" Then MorePagesAvailable = False
            ' Scrape Usernames
            Dim PinIDs() As String = GetAllStringsBetween(Content, "class=""socialItem"" href=""\/pin\/", "\/")
            If PinIDs.Count() < 2 Then PinIDs = GetAllStringsBetween(Content, "class=\\""socialItem\\"" href=\\""\/pin\/", "\/")
            If PinIDs.Count() < 2 Then PinIDs = GetAllStringsBetween(Content, """uri"": ""/v3/pins/", "/")
            ' User Search
            If PinIDs.Count() < 2 Then PinIDs = GetAllStringsBetween(Content, "}}, ""id"": """, """")
            ' MessageBox.Show(CStr("Page " & i & " - ") & CStr(PinIDs.Count()))
            ' Increase Scraped Counter
            For Each PinID As String In PinIDs
                ' Confirm Available
                If PinID = Nothing Then Continue For
                ' Other Checks
                If m_MainWindow.AlreadyUsedPinIDs.Contains(PinID) OrElse PinID.Length < 3 OrElse PinID = "" OrElse PinID.Trim() = "" OrElse CheckForAlphaCharacters(PinID) = True Then
                    Continue For
                End If
                If ScrapedData.Contains(PinID) = True Then Continue For
                ScrapedData.Add(PinID)
            Next
            ' Exit Early If Enough Data
            If ScrapedData.Count() > 20 Then Exit For
            If MorePagesAvailable = False Then Exit For
        Next
    End Sub
    ' Write To File
    Friend Sub WriteToFile(ByVal FilePath As String, ByVal ContentToAdd As String, Optional ByVal AppendData As Boolean = True)
        Try
            SyncLock m_MainWindow.FileEditorLock
                If File.Exists(FilePath) = False Then
                    Dim writer As New System.IO.StreamWriter(FilePath, OpenMode.Append, System.Text.Encoding.UTF8)
                    writer.Close()
                Else
                    ' Get Current Encoding
                    Dim CurrentEncoding As Encoding
                    Using sr As New StreamReader(FilePath, True)
                        sr.Read()
                        CurrentEncoding = sr.CurrentEncoding
                    End Using
                    ' Check Encoding
                    If CurrentEncoding Is Encoding.UTF8 = False Then

                        ' Read All File Contents
                        Dim AllFileContents As String = ""
                        Using sr As New StreamReader(FilePath, True)
                            AllFileContents = sr.ReadToEnd
                        End Using
                        ' Delete File
                        File.Delete(FilePath)
                        ' Write New File
                        Dim writer As New System.IO.StreamWriter(FilePath, OpenMode.Append, System.Text.Encoding.UTF8)
                        writer.WriteLine(AllFileContents)
                        writer.Close()
                    End If
                End If
                Dim objWriter As New System.IO.StreamWriter(FilePath, AppendData, Encoding.UTF8)
                Dim unicodeEncoding As New System.Text.UTF8Encoding()
                Dim encodedString() As Byte
                encodedString = unicodeEncoding.GetBytes(ContentToAdd)
                objWriter.WriteLine(unicodeEncoding.GetString(encodedString))
                objWriter.Close()
            End SyncLock
        Catch ex As Exception
            Log_Show_Error(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, System.Reflection.MethodInfo.GetCurrentMethod().Name, ActiveDatabase)
        End Try
    End Sub
    ' Delete Image Files
    Private Sub DeleteImageFiles(ByVal ImagesToDelete() As String)
        For Each TmpImage As String In ImagesToDelete
            If TmpImage = Nothing Then Continue For
            Try
                File.Delete(TmpImage)
            Catch ex As Exception
                ' Ignore Issues
            End Try
        Next
    End Sub
#End Region

End Class
