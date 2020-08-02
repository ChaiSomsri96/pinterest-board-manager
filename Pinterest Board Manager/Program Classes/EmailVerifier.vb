Imports System.Net
Imports System.Threading
Imports System.IO
Imports System.Text

Friend Class EmailVerifier
    ' Variables
    Private m_MainWindow As Form1
    Private ActiveDatabase As String
    Friend EmailAddress As String = ""
    Friend Password As String = ""
    Friend Proxy As String = ""
    Friend Port As Integer = 0
    Friend ProxyUser As String = ""
    Friend ProxyPass As String = ""
    Friend Username As String = ""
    Friend EmailPassword As String = ""
    Friend AccountCreatedTime As String = ""

    ' Sub New - Set Settings
    Friend Sub New(ByRef MainWindow As Form1, ByVal TmpActiveDatabase As String)
        m_MainWindow = MainWindow
        ActiveDatabase = TmpActiveDatabase
    End Sub
    ' Run Process
    Friend Sub RunProcess()
        Randomize()
        Try
            ' Start New HTTP Session
            Dim a_Process As New CurlFunctions(Username, ActiveDatabase, Proxy, Port, ProxyUser, ProxyPass)
            ' Variables
            Dim Content As String = ""
            Dim EmailBody As String = ""
            Dim VerificationURL As String = ""
            Dim EmailVerified As Boolean = False
            ' Start POP Process
            Dim PopClass As New POP3(EmailAddress, EmailPassword, ActiveDatabase)
            ' Lookup Verification Email
            Try
                EmailBody = PopClass.FindEmail(EmailAddress, "confirm your email")
            Catch ex As Exception
                m_MainWindow.AddMessage("Failed To Find Verification Email For Account: " & Username & " (" & EmailAddress & ") - " & ex.Message.Trim())
                GoTo CloseThread
            End Try
            ' Scrape Link Area  
            Dim LinkArea As String = GetBetween(EmailBody, "ededed", "</a>")
            If LinkArea = Nothing Then LinkArea = GetBetween(EmailBody, "comment_text", "</a>")
            If LinkArea = Nothing Then LinkArea = GetBetween(EmailBody, "padding:14px", "</a>")
            If LinkArea = Nothing Then LinkArea = GetBetween(EmailBody, "background-color:#bd081c", "</a>")
            If LinkArea = Nothing Then LinkArea = GetBetween(EmailBody, "border-radius:4px", "</a>")
            If LinkArea = Nothing Then LinkArea = GetBetween(EmailBody, "class=""m_plr0""", "</a>")
            ' Confirm Available
            If LinkArea = Nothing Then
                Log_Show_Error(EmailBody, "EmailVerifier_1", ActiveDatabase)
                Log_Show_Error(LinkArea, "EmailVerifier_2", ActiveDatabase)
                m_MainWindow.AddMessage("Failed To Parse Verification URL For Account: " & Username & " - Unknown Issue")
                GoTo CloseThread
            End If
            ' Parse Link
            VerificationURL = GetBetween(LinkArea, "href=3D""", """")
            If VerificationURL = Nothing Then VerificationURL = GetBetween(LinkArea, "href=""", """")
            ' Disconnect From Server
            Try
                PopClass.Disconnect()
            Catch ex As Exception

            End Try
            ' Check Confirmation Link Exists
            If VerificationURL = "" Then
                Log_Show_Error(EmailBody, "EmailVerifier_3", ActiveDatabase)
                m_MainWindow.AddMessage("Failed To Parse Confirmation Link For Account: " & Username & " - Unknown Issue")
                GoTo CloseThread
            End If
            ' Go To Verification URL
            Try
                Content = a_Process.Return_Content(VerificationURL)
            Catch ex As Exception
                m_MainWindow.AddMessage("Failed To Email Verify Account: " & Username & ", Unknown Issue")
                GoTo CloseThread
            End Try
            ' Check Content Ok
            If Not Content.Contains("Logout") AndAlso Not Content.Contains("LoggedInUser(true)") AndAlso Not Content.Contains("userIsAuthenticated = true") AndAlso Not Content.Contains("isLoggedIn"": true") Then
                ' Invalid Verification Link
                If Content.Contains("clicked on an invalid link") Then
                    m_MainWindow.AddMessage("Email Verification Failed For Account: " & Username & " - Invalid Confirmation Link. Link Used: " & VerificationURL)
                    GoTo CloseThread
                End If
                ' No Email Verification Code Supplied
                If Content.Contains("email_verification_code not specified") Then
                    m_MainWindow.AddMessage("Email Verification Failed For Account: " & Username & " - email_verification_code Not Specified. Link Used: " & VerificationURL)
                    GoTo CloseThread
                End If
                ' Incorrect Email Address In Link
                If Content.Contains("Error finding your account information") Then
                    m_MainWindow.AddMessage("Email Verification Failed For Account: " & Username & " - email_verification_code Not Specified. Link Used: " & VerificationURL)
                    GoTo CloseThread
                End If
                ' Unknown Issue
                Log_Show_Error(VerificationURL & vbCrLf & vbCrLf & Content, "EmailVerifier_4", ActiveDatabase)
            End If
            ' Report Success
            EmailVerified = True
            m_MainWindow.RefreshBoards("AccountsVerified")
            m_MainWindow.AddMessage("Successfully Email Verified Account: " & Username)
            ' Set Full Account Data
            Dim FullAccountData() As String = New String() {EmailAddress, Password, Proxy, Port, ProxyUser, ProxyPass, Username, "1", AccountCreatedTime}
            ' Backup Acount
            WriteToFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\backups\" & "VerifiedAccountsBackup.txt", Join(FullAccountData, ":"))
CloseThread:
            ' Set Old Account Data
            Dim OldAccountData() As String = New String() {EmailAddress, Password, Proxy, Port, ProxyUser, ProxyPass, Username, "0", AccountCreatedTime}
            ' Remove Old Account
            m_MainWindow.AccountCreator_CreatedAccounts.Remove(OldAccountData)
            ' Update Existing Account Data
            If EmailVerified = True Then
                Dim TmpAccountData() As String = New String() {EmailAddress, Password, Proxy, Port, ProxyUser, ProxyPass, Username, "1", AccountCreatedTime}
                m_MainWindow.AccountCreator_CreatedAccounts.Add(TmpAccountData)
            Else
                Dim TmpAccountData() As String = New String() {EmailAddress, Password, Proxy, Port, ProxyUser, ProxyPass, Username, "2", AccountCreatedTime}
                m_MainWindow.AccountCreator_CreatedAccounts.Add(TmpAccountData)
            End If
            ' Close Thread
            m_MainWindow.RefreshBoards("AccountCreator_ProcessComplete")
        Catch ex As Exception
            MessageBox.Show("A Fatal Error Occured: " & ex.Message)
            m_MainWindow.RefreshBoards("AccountCreator_ProcessComplete")
        End Try
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
End Class
