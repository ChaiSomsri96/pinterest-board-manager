Imports System.IO
Imports System.Text
Imports System.Threading

Friend Class FollowerThreadManager

#Region "Class Variables"
    ' Class Variables
    Private m_MainWindow As Form1
    Friend Campaign_Account As String
    Friend Campaign_Keywords As String
    Friend Campaign_MinFollow As String
    Friend Campaign_MaxFollow As String
    Friend Campaign_MinDelay As String
    Friend Campaign_MaxDelay As String
    Friend Campaign_Inuse As String
    Friend Campaign_FollowsSent As Integer
    Friend NextTaskTime As DateTime
    Friend Campaign_Status As String
    Friend Campaign_DelayTime As String
    Friend ActiveDatabase As String

#End Region

#Region "Class Functions"
    ' New Sub
    Friend Sub New(ByRef MainWindow As Form1)
        m_MainWindow = MainWindow
    End Sub
    ' Update Successful Stats
    Private Sub IncreaseSuccessful()
        m_MainWindow.UpdateFollowerData(Campaign_Account, "Follows Sent", "+1")
        ' SyncLock m_MainWindow.FollowerTaskUpdaterLock
        Campaign_FollowsSent = CInt(Campaign_FollowsSent) + 1
        ' m_MainWindow.UpdateFollowerDGV(Campaign_Account, "Follows Sent", CType(Campaign_FollowsSent, String))
        ' End SyncLock
    End Sub
    ' Close Thread Sub
    Private Sub CloseThread(Optional ByVal CampaignDisabled As Boolean = False)
        Try
            If m_MainWindow.Follower_StopProcess = True Then
                m_MainWindow.UpdateFollowerData(Campaign_Account, "Status", "Not Running")
            Else
                If CampaignDisabled = True Then
                    m_MainWindow.UpdateFollowerData(Campaign_Account, "Status", "Task Disabled")
                Else
                    m_MainWindow.UpdateFollowerData(Campaign_Account, "Status", "Awaiting Timer", True)
                End If
            End If
            m_MainWindow.UpdateFollowerData(Campaign_Account, "Next Task Time", CStr(DateTime.Now.AddSeconds(CDbl((Campaign_DelayTime * 60)))))
            m_MainWindow.UpdateFollowerData(Campaign_Account, "Inuse", "0")
            If CampaignDisabled = True Then
                m_MainWindow.UpdateFollowerDGV(Campaign_Account, "Status", "Task Disabled")
            Else
                If m_MainWindow.Follower_StopProcess = True Then
                    m_MainWindow.UpdateFollowerData(Campaign_Account, "Status", "Not Running")
                Else
                    m_MainWindow.UpdateFollowerData(Campaign_Account, "Status", "Awaiting Timer")
                End If
            End If
            ' Reduce Thread Count
            m_MainWindow.Follower_ThreadsRunning -= 1
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
                If m_MainWindow.Follower_ThreadsRunning < My.Settings.MaxFollowerThreads Then
                    FoundFreeThread = True
                    m_MainWindow.Follower_ThreadsRunning += 1
                End If
            End SyncLock
            ' Wait For Free Thread
            If FoundFreeThread = False Then
                Dim WaitingMsgSent As Boolean = False
                While m_MainWindow.Follower_ThreadsRunning >= My.Settings.MaxFollowerThreads
                    If WaitingMsgSent = False Then
                        m_MainWindow.UpdateFollowerDGV(Campaign_Account, "Status", "Awaiting Free Thread")
                        m_MainWindow.UpdateFollowerData(Campaign_Account, "Status", "Awaiting Free Thread")
                        WaitingMsgSent = True
                    End If
                    Thread.Sleep(1000)
                    SyncLock m_MainWindow.ThreadStarter_SyncLock
                        If m_MainWindow.Follower_ThreadsRunning >= My.Settings.MaxFollowerThreads Then
                            Continue While
                        Else
                            m_MainWindow.Follower_ThreadsRunning += 1
                            Exit While
                        End If
                    End SyncLock
                End While
            End If
            ' Already Running In A Thread, Now Call Starter Function
            m_MainWindow.UpdateFollowerDGV(Campaign_Account, "Status", "Starting Thread")
            m_MainWindow.UpdateFollowerData(Campaign_Account, "Status", "Starting Thread")
            MainRunProcess()
        Catch ex As Exception
            Log_Show_Error(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, System.Reflection.MethodInfo.GetCurrentMethod().Name, ActiveDatabase)
            MessageBox.Show("A Fatal Error Occured With RunProcess - " & ex.Message)
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
            Dim ConsecutiveFollowFailures As String = ""
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
                        ConsecutiveFollowFailures = TmpAcc(8)
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
            ' Start New HTTP Session
            Dim a_Process As New CurlFunctions(Web_Username, ActiveDatabase, Web_Proxy, Web_Port, Web_ProxyUsername, Web_ProxyPassword)
            ' Login To Account
            Try
                m_MainWindow.UpdateFollowerData(Campaign_Account, "Status", "Logging In To Account")
                m_MainWindow.Login(a_Process, Web_FriendlyUsername, Web_Password, Web_Username, Web_UserID)
            Catch ex As Exception
                m_MainWindow.AddMessage("Failed To Login To Account: " & Web_Username & " - " & ex.Message & " - Task Disabled")
                ' Close Thread & Disable Campaign
                CloseThread(True)
                Exit Sub
            End Try
            ' Clear Notifications
            Try
                m_MainWindow.UpdateFollowerData(Campaign_Account, "Status", "Clearing Notifications")
                m_MainWindow.ClearNotifications(a_Process)
            Catch ex As Exception
                ' Ignore Issues
            End Try
            ' Select Random Keyword
            Dim TmpKeyword As String = Split(Campaign_Keywords, ",")(Rnd() * (Split(Campaign_Keywords, ",").Count() - 1))
            ' Scrape Pins
            Dim ScrapedData As New List(Of String)
            Try
                m_MainWindow.UpdateFollowerData(Campaign_Account, "Status", "Scraping Users For Keyword: " & TmpKeyword)






                ScrapeUsers(a_Process, TmpKeyword, ScrapedData)
            Catch ex As Exception
                m_MainWindow.AddMessage("Failed To Scrape Users For Keyword:  " & TmpKeyword & " Using Account: " & Web_Username & " - " & ex.Message)
                GoTo CloseThisThread
            End Try
            ' Set Follows To Send
            Dim TmpUsersToFollow As Integer = GetRandomNumber(Campaign_MinFollow, Campaign_MaxFollow)
            ' Loop Processes
            For i As Integer = 1 To TmpUsersToFollow
                If m_MainWindow.Follower_StopProcess = True Then Exit For

                ' Select Random User
                Dim RandUser As String = ""
                ' Add To Used List (Thread Safe)
                SyncLock m_MainWindow.UsedListCheckLocker
                    Dim AlreadyBeingUsedFinder As Boolean = True
                    While AlreadyBeingUsedFinder = True
                        If ScrapedData.Count() < 1 Then
                            m_MainWindow.AddMessage("Scraped No New Users To Follow Keyword: " & TmpKeyword & " Using Account: " & Web_Username)
                            GoTo CloseThisThread
                        Else
                            Dim RandInt As Integer = Rnd() * (ScrapedData.Count() - 1)
                            RandUser = ScrapedData(RandInt)
                            ScrapedData.RemoveAt(RandInt)
                            If m_MainWindow.AlreadyFollowing.Contains(RandUser) = False Then Exit While
                        End If
                    End While
                    m_MainWindow.AlreadyFollowing.Add(RandUser)
                    WriteToFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\cache\" & "AlreadyFollowing.txt", RandUser)
                End SyncLock
                ' Set Pin Variables
                Dim PinID As String = RandUser




                ' Update Status
                m_MainWindow.UpdateFollowerData(Campaign_Account, "Status", "Following User: " & PinID)
                ' Follow User
                If m_MainWindow.Follower_StopProcess = True Then GoTo CloseThisThread
                Try
                    m_MainWindow.FollowUser(a_Process, PinID)
                Catch ex As Exception
                    If ex.Message.Contains("could not complete that request") Then
                        m_MainWindow.AddMessage("Failed To Follow User: " & PinID & " From Account: " & Web_Username & " - Pinterest Reported That The Request Could Not Be Completed")
                        ' Increase Failure Count
                        m_MainWindow.UpdateAccountData(Web_Username, "ConsecutiveFollowFailures", "+1")
                        ConsecutiveFollowFailures = CStr(CInt(ConsecutiveFollowFailures) + 1)
                        ' Check Failures
                        If ConsecutiveFollowFailures > 2 Then
                            WriteToFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\cache\" & "BadFollowingAccounts.txt", Web_Username & ":" & Web_Password & ":" & Web_Proxy & ":" & Web_Port & ":" & Web_ProxyUsername & ":" & Web_ProxyPassword, True, True)
                            CloseThread(True)
                            Exit Sub
                        End If
                        ' Close Thread
                        GoTo WaitTimer
                    Else
                        m_MainWindow.AddMessage("Failed To Follow User: " & PinID & " From Account: " & Web_Username & " - " & ex.Message)
                        ' Close Thread
                        GoTo WaitTimer
                    End If
                End Try
                ' Reset Failures
                ConsecutiveFollowFailures = "0"
                m_MainWindow.UpdateAccountData(Web_Username, "ConsecutiveFollowFailures", "0")
                ' Add Success
                IncreaseSuccessful()
                m_MainWindow.AddMessage("Successfully Followed User: " & PinID & " From Account: " & Web_Username)
                ' Sleep
WaitTimer:
                m_MainWindow.UpdateFollowerData(Campaign_Account, "Status", "Waiting Delay Time")
                m_MainWindow.Sleep(GetRandomNumber(Campaign_MinDelay, Campaign_MaxDelay), "Follower")
            Next
            ' Save Cookies
            a_Process.Dispose()
CloseThisThread:
            ' Close Thread
            CloseThread()
        Catch ex As Exception
            Log_Show_Error(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, System.Reflection.MethodInfo.GetCurrentMethod().Name, ActiveDatabase)
            MessageBox.Show("A Fatal Error Occured In Follower MainRunProcess - " & ex.Message)
            CloseThread()
        End Try
    End Sub
    ' Scrape Pins Sub
    Private Sub ScrapeUsers(ByRef a_Process As CurlFunctions, ByVal Keyword As String, ByRef ScrapedData As List(Of String))
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
        ' Loop Pages
        For i As Integer = 1 To My.Settings.MaxPages
            If m_MainWindow.Follower_StopProcess = True OrElse MorePagesAvailable = False OrElse NextMaxID = Nothing Then Exit For
            ' Load Search Page
            Try
                ' Search Pins Mode
                If i = 1 Then
                    URL = "https://www." & MainDomain & "/resource/BaseSearchResource/get/?source_url=%2Fsearch%2Fpins%2F%3Fq%3D" & URLEncode(Keyword.Trim()) & "&data=%7B%22options%22%3A%7B%22filters%22%3Anull%2C%22query%22%3A%22" & URLEncode(Keyword.Trim()) & "%22%2C%22rs%22%3Anull%2C%22scope%22%3A%22pins%22%7D%2C%22context%22%3A%7B%7D%7D&_=" & GetMilliSeconds()
                    Content = a_Process.XMLHttpRequest_Get_Request(URL, Token, "https://www." & MainDomain & "/search/pins/?q=" & URLEncode(Keyword.Trim()), "", True, AppVersion, "active")
                Else
                    URL = "https://www." & MainDomain & "/resource/SearchResource/get/?source_url=%2Fsearch%2Fpins%2F%3Fq%3D" & URLEncode(Keyword.Trim()) & "&data=%7B%22options%22%3A%7B%22bookmarks%22%3A%5B%22" & URLEncode(NextMaxID) & "%22%5D%2C%22query%22%3A%22" & URLEncode(Keyword.Trim()) & "%22%2C%22scope%22%3A%22pins%22%7D%2C%22context%22%3A%7B%7D%7D&_=" & GetMilliSeconds()
                    Content = a_Process.XMLHttpRequest_Get_Request(URL, Token, "https://www." & MainDomain & "/search/?q=" & URLEncode(Keyword.Trim()))
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
            Dim PinIDs() As String = GetAllStringsBetween(Content, """username"":""", """")
            ' Increase Scraped Counter
            For Each PinID As String In PinIDs
                ' Confirm Available
                If PinID = Nothing Then Continue For
                ' Other Checks
                If m_MainWindow.AlreadyFollowing.Contains(PinID) OrElse PinID.Length < 3 OrElse PinID = "" OrElse PinID.Trim() = "" Then
                    Continue For
                End If
                If ScrapedData.Contains(PinID) = True Then Continue For
                ScrapedData.Add(PinID)
            Next
            ' Exit Early If Enough Data
            If ScrapedData.Count() > (Campaign_MaxFollow * 3) Then Exit For
            If MorePagesAvailable = False Then Exit For
        Next
    End Sub
    ' Write To File
    Friend Sub WriteToFile(ByVal FilePath As String, ByVal ContentToAdd As String, Optional ByVal AppendData As Boolean = True, Optional ByVal CheckIfAlreadyExists As Boolean = False)
        Try
            SyncLock m_MainWindow.FileEditorLock
                If File.Exists(FilePath) = False Then
                    Dim writer As New System.IO.StreamWriter(FilePath, OpenMode.Append, System.Text.Encoding.UTF8)
                    writer.Close()
                Else
                    ' Check If Already Exists
                    If CheckIfAlreadyExists = True Then
                        Dim FileContent = My.Computer.FileSystem.ReadAllText(FilePath)
                        If FileContent.Contains(ContentToAdd) Then
                            Exit Sub ' Already Exists
                        End If
                    End If
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
#End Region

End Class
