Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Threading

Friend Class Form1

#Region "Form Variables"
    ' Form Variables
    Friend MyVersion As String = "1.0.3.4"
    Friend MyWebsite As String = "www.BlackHatToolz.com"
    Friend SoftwareName As String = "Pinterest Board Manager"
    Friend ActiveDatabase As String = "datastore"
    Friend Accounts As New List(Of String())
    ' Board Manager Variables
    Friend Full_Campaign_Data As New DataTable
    Friend BoardInformation As New List(Of String())
    Friend AlreadyUsedPinIDs As New HashSet(Of String)
    Friend BoardManager_StopProcess As Boolean = False
    ' Follower Variables
    Friend Full_Following_Task_Data As New DataTable
    Friend AlreadyFollowing As New HashSet(Of String)
    Friend Follower_StopProcess As Boolean = False
    ' Account Creator Variables
    Friend AccountCreator_Usernames As New List(Of String)
    Friend AccountCreator_BusinessNames As New List(Of String)
    Friend AccountCreator_FirstNames As New List(Of String)
    Friend AccountCreator_LastNames As New List(Of String)
    Friend AccountCreator_EmailAccounts As New List(Of String)
    Friend AccountCreator_BoardsToCreate As New List(Of String)
    Friend AccountCreator_CreatedAccounts As New List(Of String())
    Friend AccountCreator_Proxies As New List(Of String())
    Friend AccountCreator_TempProxies As New List(Of String())
    Friend AccountCreator_UsernameBlacklist As New HashSet(Of String)
    Friend AccountCreator_EmailBlacklist As New HashSet(Of String)
    Friend AccountCreator_BusinessNameBlacklist As New HashSet(Of String)
    Friend AccountCreator_LastChosenProxy As Integer = 0
    Friend AccountCreator_ProfileImagesFolder As String
    Friend AccountCreator_ProfileGender As String
    Friend AccountCreator_BusinessType As String
    Friend AccountCreator_MaxAccountsToCreate As String
    Friend AccountCreator_DelayBetweenCreation As String
    Friend AccountCreator_EmailVerifyDelay As String
    Friend AccountCreator_AutoAddAccounts As Boolean
    Friend AccountCreator_StopProcess As Boolean = False
    Friend AccountCreator_AccountsCreated As Integer = 0
    Friend AccountCreator_AccountsVerified As Integer = 0
    Friend AccountCreator_ThreadsRunning As Integer = 0
    ' Global Multi-Threading Variables
    Friend ThreadsRunning As Integer = 0
    Friend Follower_ThreadsRunning As Integer = 0
    Friend AccountUpdaterLock As New Object
    Friend TaskUpdaterLock As New Object
    Friend FollowerTaskUpdaterLock As New Object
    Friend BoardInformationLock As New Object
    Friend FileEditorLock As New Object
    Friend ThreadStarter_SyncLock As New Object
    Friend UsedListCheckLocker As New Object
    ' Glogal Delegates
    Private m_MainWindow As Form1
    Private Delegate Sub d_AddMessage(ByVal Message As String)
    Private Delegate Sub d_RefreshBoards(ByVal Message As String)
    Private Delegate Sub d_UpdateAccountData(ByVal AccountUsername As String, ByVal FieldToUpdate As String, ByVal NewValue As String, ByVal ForceUpdate As Boolean)
    ' Licensing Variables
    Friend MainResponse As String = ""
    Friend MachineResponse As String = ""
    Friend Unlocked As Boolean = False
    Friend MachineName As String
    Friend LocalIP As String
    Friend MainDomains As String
    Friend checker As Class1
    Friend crytpo As crypo
    ' Forms
    Friend Form_AccountManager As New AccountManager

#End Region

#Region "Delegate Procedures"
    ' Add Message
    Friend Sub AddMessage(ByVal Msg As String)
        If Me.InvokeRequired Then
            Me.Invoke(New d_AddMessage(AddressOf AddMessage), Msg)
        Else
            Dim strTime As String = "[" & DateTime.Now.ToString("t") & "] "
            Msg = strTime & Msg
            TextBox4.Text = Msg & vbCrLf & TextBox4.Text
        End If
    End Sub
    ' Refresh Boards
    Friend Sub RefreshBoards(ByVal AccountUsername As String, Optional ByVal TextField As String = "BoardManager")
        If Me.InvokeRequired Then
            Me.Invoke(New d_RefreshBoards(AddressOf RefreshBoards), AccountUsername)
        Else
            ' Select Module That Has Completed
            Select Case AccountUsername
                Case "ProcessComplete"
                    ' Reset LinkLabel
                    LinkLabel1.Text = "Refresh Board Information"
                    LinkLabel1.Enabled = True
                    LinkLabel1.Location = New Point(162, 60)
                    AddMessage("Completed Looking Up Board Information Processes")
                Case "AccountCreator_ProcessComplete"
                    AccountCreator_ThreadsRunning -= 1
                    If AccountCreator_ThreadsRunning <= 0 Then
                        AccountCreator_ThreadsRunning = 0
                        Button16.Enabled = True
                        Button14.Enabled = False
                        AddMessage("Account Creator Completed")
                    End If
                Case "AccountsVerified"
                    AccountCreator_AccountsVerified += 1
                    Label26.Text = AccountCreator_AccountsVerified.ToString()
                Case "AccountsCreated"
                    AccountCreator_AccountsCreated += 1
                    Label24.Text = AccountCreator_AccountsCreated.ToString()
                Case "RefreshAccounts"
                    ' RefreshAccounts()
                    Form_AccountManager.RefreshAccounts()
                Case Else
                    ' Remove Old Results
                    If TextField = "BoardManager" Then
                        ComboBox2.Items.Clear()
                    Else
                        TextBox10.Text = ""
                    End If
                    ' Add Boards To List & ComboBox
                    For Each TmpBoardInfo() As String In BoardInformation
                        If TmpBoardInfo(0) <> AccountUsername Then Continue For
                        Dim BoardName As String = TmpBoardInfo(1)
                        Dim BoardID As String = TmpBoardInfo(2)
                        ' Add To ComboBox
                        If TextField = "BoardManager" Then
                            ComboBox2.Items.Add(BoardName)
                            ' Check If Already Exists
                            Dim userToRemove = BoardInformation.Cast(Of Object()).Where(Function(j) j(2).Equals(BoardID)).Any()
                            If userToRemove = True Then Continue For
                            ' Add To List
                            BoardInformation.Add(New String() {AccountUsername, BoardName, BoardID})
                        Else
                            TextBox10.Text = BoardName & vbCrLf & TextBox10.Text
                        End If
                    Next
                    ' Select Default
                    If ComboBox2.Items.Count() > 0 Then ComboBox2.SelectedIndex = 0
            End Select
        End If
    End Sub
    ' Update Board Manager Task Data
    Friend Sub UpdateAccountData(ByVal AccountUsername As String, ByVal FieldToUpdate As String, ByVal NewValue As String, Optional ByVal ForceUpdate As Boolean = False)
        If Me.InvokeRequired Then
            Me.Invoke(New d_UpdateAccountData(AddressOf UpdateAccountData), {AccountUsername, FieldToUpdate, NewValue, ForceUpdate})
        Else
            Try
                SyncLock AccountUpdaterLock
                    For Each TmpAccountData() As String In Accounts
                        If AccountUsername = "ALL_ACCOUNTS" OrElse (CStr(TmpAccountData(0)) = AccountUsername) Then
                            Select Case FieldToUpdate
                                Case "ConsecutiveFollowFailures"
                                    Dim TmpUpdatedValue As String = NewValue
                                    If NewValue = "+1" Then
                                        TmpUpdatedValue = CStr(CInt(TmpAccountData(8)) + 1)
                                        TmpAccountData(8) = TmpUpdatedValue
                                    Else
                                        TmpAccountData(8) = TmpUpdatedValue
                                    End If
                            End Select
                            ' Exit Loop
                            If AccountUsername <> "ALL_ACCOUNTS" Then Exit For
                        End If
                    Next
                    ' Save Campaign Data
                    Form_AccountManager.SaveAccounts()
                End SyncLock
            Catch ex As Exception
                MessageBox.Show("An Issue Occured Updating Account Data: " & ex.Message)
            End Try
        End If
    End Sub
#End Region

#Region "Data Management Functions"
    ' Load File Sub
    Friend Sub LoadFile(ByVal TxtFile As String, ByVal DataName As String)
        ' Load Accounts
        If File.Exists(TxtFile) Then
            Dim objReader As New System.IO.StreamReader(TxtFile)
            Dim Line As String
            Dim AccountsLoaded As Integer = 0
            Do While objReader.Peek() <> -1
                Line = objReader.ReadLine()
                If Line = Nothing OrElse Line = "" Then Continue Do
                Line.Replace("	", "")
                ' Select Data Type
                Select Case DataName
                    Case "Accounts"
                        LoadAccount(Line)
                    Case "Tasks"
                        LoadTask(Line)
                    Case "FollowerTask"
                        LoadFollowerTask(Line)
                    Case "BoardInformation"
                        LoadBoardInformation(Line)
                    Case "AlreadyUsed"
                        If Line <> Nothing AndAlso Line.Trim() <> Nothing AndAlso AlreadyUsedPinIDs.Contains(Line) = False Then AlreadyUsedPinIDs.Add(Line)
                    Case "AlreadyFollowing"
                        If Line <> Nothing AndAlso Line.Trim() <> Nothing AndAlso AlreadyFollowing.Contains(Line) = False Then AlreadyFollowing.Add(Line)
                    Case "UsernameBlacklist"
                        If Line <> Nothing AndAlso Line.Trim() <> Nothing AndAlso AccountCreator_UsernameBlacklist.Contains(Line) = False Then AccountCreator_UsernameBlacklist.Add(Line)
                    Case "EmailBlacklist"
                        If Line <> Nothing AndAlso Line.Trim() <> Nothing AndAlso AccountCreator_EmailBlacklist.Contains(Line) = False Then AccountCreator_EmailBlacklist.Add(Line)
                    Case "AccountCreator_Proxies"
                        If Line = Nothing OrElse Line = "" OrElse Line = vbCrLf OrElse Line.Trim() = "" Then Continue Do
                        Line.Replace("	", "")
                        Dim Account As String() = Split(Line.Trim(), ":")
                        If Account.Count < 2 Then Continue Do
                        Dim Proxy As String = Account(0)
                        Dim Port As String = Account(1)
                        Dim i_ProxyUsername As String = "0"
                        Dim i_ProxyPassword As String = "0"
                        If Account.Count > 2 Then
                            i_ProxyUsername = Account(2)
                            i_ProxyPassword = Account(3)
                        End If
                        ' Check For Valid Port
                        If Integer.TryParse(Port, 0) = False Then Continue Do
                        ' Check To See If Already Exists
                        AccountCreator_Proxies.Add(New String() {Proxy, Port, i_ProxyUsername, i_ProxyPassword})
                    Case "AccountCreator_Usernames"
                        If Line <> Nothing AndAlso Line.Trim() <> Nothing AndAlso AccountCreator_UsernameBlacklist.Contains(Line) = False Then AccountCreator_Usernames.Add(Line)
                    Case "AccountCreator_EmailAccounts"
                        If Line <> Nothing AndAlso Line.Trim() <> Nothing AndAlso AccountCreator_EmailBlacklist.Contains(Line) = False Then AccountCreator_EmailAccounts.Add(Line)
                    Case "AccountCreator_FirstNames"
                        If Line <> Nothing AndAlso Line.Trim() <> Nothing Then AccountCreator_FirstNames.Add(Line)
                    Case "AccountCreator_LastNames"
                        If Line <> Nothing AndAlso Line.Trim() <> Nothing Then AccountCreator_LastNames.Add(Line)
                    Case "AccountCreator_BoardsToCreate"
                        If Line <> Nothing AndAlso Line.Trim() <> Nothing Then AccountCreator_BoardsToCreate.Add(Line)
                    Case "AccountCreator_BusinessNames"
                        If Line <> Nothing AndAlso Line.Trim() <> Nothing AndAlso AccountCreator_BusinessNameBlacklist.Contains(Line) = False Then AccountCreator_BusinessNames.Add(Line)
                    Case "BusinessNameBlacklist"
                        If Line <> Nothing AndAlso Line.Trim() <> Nothing AndAlso AccountCreator_BusinessNameBlacklist.Contains(Line) = False Then AccountCreator_BusinessNameBlacklist.Add(Line)
                    Case Else
                        Throw New Exception("Unknown Data Source: " & DataName)
                End Select
            Loop
            objReader.Close()
        End If
    End Sub
    ' Create Task DataTable
    Private Sub CreateTaskDataTable()
        Try
            ' Create Table
            Full_Campaign_Data.Columns.Add("Account")
            Full_Campaign_Data.Columns.Add("Board Name")
            Full_Campaign_Data.Columns.Add("Keywords")
            Full_Campaign_Data.Columns.Add("Website")
            Full_Campaign_Data.Columns.Add("Delay", GetType(Integer))
            Full_Campaign_Data.Columns.Add("Inuse", GetType(Integer))
            Full_Campaign_Data.Columns.Add("Pins Created", GetType(Integer))
            Full_Campaign_Data.Columns.Add("Next Post Time", GetType(DateTime))
            Full_Campaign_Data.Columns.Add("Status")
            Full_Campaign_Data.Columns.Add("Total Impressions")
            ' Add To DGV
            DataGridView1.DataSource = Full_Campaign_Data
            ' Hide Unwanted Columns
            DataGridView1.Columns("Keywords").Visible = False
            DataGridView1.Columns("Website").Visible = False
            DataGridView1.Columns("Delay").Visible = False
            DataGridView1.Columns("Inuse").Visible = False
            ' Resize Columns
            DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Catch ex As Exception
            MessageBox.Show("Failed To Create Task DataTable - " & ex.Message & " - " & ex.StackTrace)
        End Try
    End Sub
    ' Create Follower Task DataTable
    Private Sub CreateFollowerTaskDataTable()
        Try
            ' Create Table
            Full_Following_Task_Data.Columns.Add("Account")
            Full_Following_Task_Data.Columns.Add("Keywords")
            Full_Following_Task_Data.Columns.Add("MinFollow", GetType(Integer))
            Full_Following_Task_Data.Columns.Add("MaxFollow", GetType(Integer))
            Full_Following_Task_Data.Columns.Add("MinDelay", GetType(Integer))
            Full_Following_Task_Data.Columns.Add("MaxDelay", GetType(Integer))
            Full_Following_Task_Data.Columns.Add("Inuse", GetType(Integer))
            Full_Following_Task_Data.Columns.Add("Follows Sent", GetType(Integer))
            Full_Following_Task_Data.Columns.Add("Next Task Time", GetType(DateTime))
            Full_Following_Task_Data.Columns.Add("Status")
            Full_Following_Task_Data.Columns.Add("DelayTime")
            ' Add To DGV
            DataGridView2.DataSource = Full_Following_Task_Data
            ' Hide Unwanted Columns
            DataGridView2.Columns("MinFollow").Visible = False
            DataGridView2.Columns("MaxFollow").Visible = False
            DataGridView2.Columns("MinDelay").Visible = False
            DataGridView2.Columns("MaxDelay").Visible = False
            DataGridView2.Columns("Inuse").Visible = False
            DataGridView2.Columns("DelayTime").Visible = False
            ' Resize Columns
            DataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Catch ex As Exception
            MessageBox.Show("Failed To Create Following Task DataTable - " & ex.Message & " - " & ex.StackTrace)
        End Try
    End Sub
    ' Load Database
    Private Sub LoadDatabase(ByVal DB_Name As String)
        ' Create Required Folders
        Dim Folders() As String = {ActiveDatabase, ActiveDatabase & "\backups", ActiveDatabase & "\cache", ActiveDatabase & "\cookies", ActiveDatabase & "\error_logs", ActiveDatabase & "\exports", ActiveDatabase & "\session", ActiveDatabase & "\images"}
        For Each Folder In Folders
            If Not (Directory.Exists(My.Application.Info.DirectoryPath & "\" & Folder & "\")) Then
                Try
                    Directory.CreateDirectory(My.Application.Info.DirectoryPath & "\" & Folder & "\")
                Catch ex As Exception
                    MessageBox.Show("Failed To Create Required Folders:" & vbCrLf & ex.Message)
                End Try
            End If
        Next
        ' Remove Old Accounts
        ComboBox1.Items.Clear()
        Accounts.Clear()
        ' Load Accounts
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "Accounts.txt", "Accounts")
        ' Refresh Account Data
        RefreshAccounts()
        ' Remove Old Tasks
        Full_Campaign_Data.Rows.Clear()
        ' Load Tasks
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "Tasks.txt", "Tasks")
        ' Remove Old Tasks
        Full_Following_Task_Data.Rows.Clear()
        ' Load Follower Tasks
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "FollowerTasks.txt", "FollowerTask")
        ' Load Already Used Data
        AlreadyUsedPinIDs.Clear()
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\cache\" & "AlreadyUsedPinIDs.txt", "AlreadyUsed")
        ' Load Already Used Data
        AlreadyFollowing.Clear()
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\cache\" & "AlreadyFollowing.txt", "AlreadyFollowing")
        ' Load Already Used Data
        AccountCreator_UsernameBlacklist.Clear()
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\cache\" & "UsernameBlacklist.txt", "UsernameBlacklist")
        ' Load Already Used Data
        AccountCreator_EmailBlacklist.Clear()
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\cache\" & "EmailBlacklist.txt", "EmailBlacklist")
        ' Load Already Used Data
        AccountCreator_BusinessNameBlacklist.Clear()
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\cache\" & "BusinessNameBlacklist.txt", "BusinessNameBlacklist")
        ' Load Account Creator Proxies
        AccountCreator_Proxies.Clear()
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_Proxies.txt", "AccountCreator_Proxies")
        ' Load Account Creator Usernames
        AccountCreator_Usernames.Clear()
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_Usernames.txt", "AccountCreator_Usernames")
        ' Load Account Creator Email Accounts
        AccountCreator_EmailAccounts.Clear()
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_EmailAccounts.txt", "AccountCreator_EmailAccounts")
        ' Load Account Creator First Names
        AccountCreator_FirstNames.Clear()
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_FirstNames.txt", "AccountCreator_FirstNames")
        ' Load Account Creator Last Names
        AccountCreator_LastNames.Clear()
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_LastNames.txt", "AccountCreator_LastNames")
        ' Load Account Creator Boards To Create
        AccountCreator_BoardsToCreate.Clear()
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_BoardsToCreate.txt", "AccountCreator_BoardsToCreate")
        ' Load Business Names To Set
        AccountCreator_BusinessNames.Clear()
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_BusinessNames.txt", "AccountCreator_BusinessNames")
        ' Clear Board Information
        ComboBox2.Items.Clear()
        BoardInformation.Clear()
        ' Display Account Creator Counts
        Label40.Text = AccountCreator_Proxies.Count().ToString()
        Label8.Text = AccountCreator_Usernames.Count().ToString()
        Label9.Text = AccountCreator_EmailAccounts.Count().ToString()
        Label19.Text = AccountCreator_FirstNames.Count().ToString()
        Label21.Text = AccountCreator_LastNames.Count().ToString()
        Label38.Text = AccountCreator_BoardsToCreate.Count().ToString()
        Label42.Text = AccountCreator_BusinessNames.Count().ToString()
        ' Load New Information
        LoadFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "BoardInformation.txt", "BoardInformation")
    End Sub
    ' Load Load Task Sub
    Private Sub LoadTask(ByVal Line As String)
        ' Format Data Tokens
        Line = Line.Replace("\\|", "{PIPE_KEY}")
        ' Split Data
        Dim Account As String() = Split(Line, "|")
        ' Check Length
        If Account.Count < 10 Then Exit Sub
        ' Set Account Data
        Dim Campaign_Account As String = Account(0).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_BoardName As String = Account(1).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_Keywords As String = Account(2).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_Website As String = Account(3).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_Delay As String = Account(4).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_Inuse As String = Account(5).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_PinsCreated As String = Account(6).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim TmpNextPostTime As String = Account(7).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_Status As String = Account(8).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim TotalImpressions As String = Account(9).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        ' New Variables
        Campaign_Status = "Not Running"
        Campaign_Inuse = "0"
        ' Check For Next Post Date
        Dim Campaign_NextPostDate As DateTime = CDate(DateTime.Now.ToString())
        Try
            Campaign_NextPostDate = CDate(Account(7))
        Catch ex As Exception
            Campaign_NextPostDate = CDate(DateTime.Now.ToString())
        End Try
        ' Add Account
        Full_Campaign_Data.Rows.Add(Campaign_Account, Campaign_BoardName, Campaign_Keywords, Campaign_Website, Campaign_Delay, Campaign_Inuse, Campaign_PinsCreated, Campaign_NextPostDate, Campaign_Status, TotalImpressions)
    End Sub
    ' Load Load Follower Task Sub
    Private Sub LoadFollowerTask(ByVal Line As String)
        ' Format Data Tokens
        Line = Line.Replace("\\|", "{PIPE_KEY}")
        ' Split Data
        Dim Account As String() = Split(Line, "|")
        ' Check Length
        If Account.Count < 11 Then Exit Sub
        ' Set Account Data
        Dim Campaign_Account As String = Account(0).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_Keywords As String = Account(1).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_MinFollow As String = Account(2).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_MaxFollow As String = Account(3).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_MinDelay As String = Account(4).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_MaxDelay As String = Account(5).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_Inuse As String = Account(6).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_FollowsSent As String = Account(7).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim TmpNextPostTime As String = Account(8).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_Status As String = Account(9).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim Campaign_DelayTime As String = Account(10).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        ' New Variables
        If Campaign_Status <> "Task Disabled" Then Campaign_Status = "Not Running"
        Campaign_Inuse = "0"
        ' Check For Next Post Date
        Dim Campaign_NextPostDate As DateTime = CDate(DateTime.Now.ToString())
        Try
            Campaign_NextPostDate = CDate(Account(8))
        Catch ex As Exception
            Campaign_NextPostDate = CDate(DateTime.Now.ToString())
        End Try
        ' Add Account
        Full_Following_Task_Data.Rows.Add(Campaign_Account, Campaign_Keywords, Campaign_MinFollow, Campaign_MaxFollow, Campaign_MinDelay, Campaign_MaxDelay, Campaign_Inuse, Campaign_FollowsSent, Campaign_NextPostDate, Campaign_Status, Campaign_DelayTime)
    End Sub
    ' Load Account Sub
    Private Sub LoadAccount(ByVal Line As String)
        Line.Replace("	", "")
        Dim Account As String() = Split(Line, ":")
        If Account.Count < 2 Then Exit Sub
        Dim Username As String = Account(0)
        Dim EmailAddress As String = ""
        Dim Password As String = Account(1)
        Dim Proxy As String = ""
        Dim Port As String = "0"
        If Account.Count > 2 Then
            Proxy = Account(2)
            If Account.Count() < 4 Then Exit Sub
            If Integer.TryParse(Account(3), 0) = False Then Exit Sub
            Port = Account(3)
        End If
        Dim i_ProxyUsername As String = "0"
        Dim i_ProxyPassword As String = "0"
        If Account.Count > 4 Then
            i_ProxyUsername = Account(4)
            If Account.Count() < 6 Then Exit Sub
            i_ProxyPassword = Account(5)
        End If
        ' Check For Follower Failures
        Dim ConsecutiveFollowFailures As String = "0"
        If Account.Count > 6 Then ConsecutiveFollowFailures = Account(6)
        ' Check For Valid Port
        If Integer.TryParse(Port, 0) = False Then Exit Sub
        ' Check To See If Already Exists
        Dim userToRemove = Accounts.Cast(Of Object()).Where(Function(j) j(0).Equals(Username)).Any()
        If userToRemove = True Then Exit Sub
        Accounts.Add(New String() {Username, Password, Proxy, Port, i_ProxyUsername, i_ProxyPassword, "1", "0", ConsecutiveFollowFailures})
    End Sub
    ' Load Board Information
    Private Sub LoadBoardInformation(ByVal Line As String)
        ' Format Data Tokens
        Line = Line.Replace("\\|", "{PIPE_KEY}")
        ' Split Data
        Dim Account As String() = Split(Line, "|")
        ' Check Length
        If Account.Count < 3 Then Exit Sub
        ' Set Data
        Dim AccountUsername As String = Account(0).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim BoardName As String = Account(1).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        Dim BoardID As String = Account(2).Replace("{PIPE_KEY}", "|").Replace("\r\n", vbCrLf)
        ' Check Not Already Exists
        Dim userToRemove = BoardInformation.Cast(Of Object()).Where(Function(j) j(2).Equals(BoardID)).Any()
        If userToRemove = True Then Exit Sub
        ' Add Data
        BoardInformation.Add(New String() {AccountUsername, BoardName, BoardID})
    End Sub
    ' Update Board Manager Task Data
    Friend Sub UpdateCampaignData(ByVal AccountUsername As String, ByVal BoardName As String, ByVal FieldToUpdate As String, ByVal NewValue As String, Optional ByVal ForceUpdate As Boolean = False)
        Try
            SyncLock TaskUpdaterLock
                For Each row As DataRow In Full_Campaign_Data.Rows
                    If AccountUsername = "ALL_CAMPAIGNS" OrElse (CStr(row("Account")) = AccountUsername AndAlso CStr(row("Board Name")) = BoardName) Then
                        Select Case FieldToUpdate
                            Case "Status"
                                row("Status") = NewValue
                            Case "Inuse"
                                row("Inuse") = CInt(NewValue)
                            Case "Next Post Time"
                                row("Next Post Time") = CDate(NewValue)
                            Case "Pins Created"
                                row("Pins Created") = CInt(row("Pins Created")) + 1
                            Case "Total Impressions"
                                row("Total Impressions") = NewValue
                        End Select
                        ' Exit Loop
                        If AccountUsername <> "ALL_CAMPAIGNS" Then Exit For
                    End If
                Next
                ' Save Campaign Data
                m_MainWindow.SaveTasks()
            End SyncLock
        Catch ex As Exception
            MessageBox.Show("An Issue Occured Updating Account Data: " & ex.Message)
        End Try
    End Sub
    ' Update Board Manager Task Data
    Friend Sub UpdateFollowerData(ByVal AccountUsername As String, ByVal FieldToUpdate As String, ByVal NewValue As String, Optional ByVal ForceUpdate As Boolean = False)
        Try
            SyncLock FollowerTaskUpdaterLock
                For Each row As DataRow In Full_Following_Task_Data.Rows
                    If AccountUsername = "ALL_CAMPAIGNS" OrElse CStr(row("Account")) = AccountUsername Then
                        Select Case FieldToUpdate
                            Case "Status"
                                If row("Status") <> "Task Disabled" Then row("Status") = NewValue
                            Case "Inuse"
                                row("Inuse") = CInt(NewValue)
                            Case "Next Task Time"
                                row("Next Task Time") = CDate(NewValue)
                            Case "Follows Sent"
                                row("Follows Sent") = CInt(row("Follows Sent")) + 1
                        End Select
                        ' Exit Loop
                        If AccountUsername <> "ALL_CAMPAIGNS" Then Exit For
                    End If
                Next
                ' Save Campaign Data
                m_MainWindow.SaveFollowerTasks()
            End SyncLock
        Catch ex As Exception
            MessageBox.Show("An Issue Occured Updating Account Data: " & ex.Message)
        End Try
    End Sub
    ' Update Board Manager DGV Cell Sub
    Friend Sub UpdateDGV(ByVal AccountUsername As String, ByVal BoardName As String, ByVal ColumnName As String, ByVal TextContent As String, Optional ByVal WhereCurrentValueIs As String = "", Optional ByVal ForceUpdate As Boolean = False)

        Exit Sub
        Try
            ' Update DGV
            ' SyncLock Update_DGV_SyncLock
            For Each row As DataRow In Full_Campaign_Data.Rows
                    If AccountUsername = "ALL_CAMPAIGNS" OrElse (CStr(row("Account")) = AccountUsername AndAlso CStr(row("Board Name")) = BoardName) Then
                        If ColumnName = "Delay" OrElse ColumnName = "Inuse" OrElse ColumnName = "Pins Created" OrElse ColumnName = "Enabled" Then
                            If WhereCurrentValueIs <> Nothing AndAlso WhereCurrentValueIs <> CStr(row(ColumnName)) Then Continue For
                            row(ColumnName) = CInt(TextContent)
                        ElseIf ColumnName = "Status" Then
                            row(ColumnName) = TextContent
                        Else
                            If WhereCurrentValueIs <> Nothing AndAlso WhereCurrentValueIs <> CStr(row(ColumnName)) Then Continue For
                            row(ColumnName) = TextContent
                        End If
                        Full_Campaign_Data.AcceptChanges()
                        If AccountUsername <> "ALL_CAMPAIGNS" Then Exit For
                    End If
                Next
            ' End SyncLock
        Catch ex As Exception
            Log_Show_Error(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, Me.Name & "_" & System.Reflection.MethodInfo.GetCurrentMethod().Name, ActiveDatabase)
            MessageBox.Show("A Fatal Error Occured Updating The DGV - " & ex.Message)
        End Try
    End Sub
    ' Update Follower DGV Cell Sub
    Friend Sub UpdateFollowerDGV(ByVal AccountUsername As String, ByVal ColumnName As String, ByVal TextContent As String, Optional ByVal WhereCurrentValueIs As String = "", Optional ByVal ForceUpdate As Boolean = False)
        Exit Sub
        Try
            ' Update DGV
            ' SyncLock Update_DGV_SyncLock
            For Each row As DataRow In Full_Following_Task_Data.Rows
                    If AccountUsername = "ALL_CAMPAIGNS" OrElse CStr(row("Account")) = AccountUsername Then
                        If ColumnName = "Inuse" OrElse ColumnName = "Follows Sent" Then
                            If WhereCurrentValueIs <> Nothing AndAlso WhereCurrentValueIs <> CStr(row(ColumnName)) Then Continue For
                            row(ColumnName) = CInt(TextContent)
                        ElseIf ColumnName = "Status" Then
                            row(ColumnName) = TextContent
                        Else
                            If WhereCurrentValueIs <> Nothing AndAlso WhereCurrentValueIs <> CStr(row(ColumnName)) Then Continue For
                            row(ColumnName) = TextContent
                        End If
                        Full_Following_Task_Data.AcceptChanges()
                        If AccountUsername <> "ALL_CAMPAIGNS" Then Exit For
                    End If
                Next
            ' End SyncLock
        Catch ex As Exception
            Log_Show_Error(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, Me.Name & "_" & System.Reflection.MethodInfo.GetCurrentMethod().Name, ActiveDatabase)
            MessageBox.Show("A Fatal Error Occured Updating The Follower DGV - " & ex.Message)
        End Try
    End Sub
    ' Save Board Manager Task Data
    Friend Sub SaveTasks()
        Try
            Dim FilePath As String = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "Tasks.txt"
            SyncLock TaskUpdaterLock
                Full_Campaign_Data.AcceptChanges()
                Using sw As New StreamWriter(FilePath, False)
                    For Each drow As DataRow In Full_Campaign_Data.Rows
                        If drow.RowState = DataRowState.Deleted Then Continue For
                        Dim lineoftext As String = String.Join("|", drow.ItemArray.Select(Function(s) s.ToString.Replace("|", "{PIPE_KEY}")).ToArray)
                        sw.WriteLine(lineoftext.Replace(vbCrLf, "\r\n").Replace(vbCr, "\r\n").Replace(vbLf, "\r\n"))
                    Next
                End Using
            End SyncLock
        Catch ex As Exception
            ' Ignore Issues
        End Try
    End Sub
    ' Save Follower Task Data
    Friend Sub SaveFollowerTasks()
        Try
            Dim FilePath As String = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "FollowerTasks.txt"
            SyncLock FollowerTaskUpdaterLock
                Full_Following_Task_Data.AcceptChanges()
                Using sw As New StreamWriter(FilePath, False)
                    For Each drow As DataRow In Full_Following_Task_Data.Rows
                        If drow.RowState = DataRowState.Deleted Then Continue For
                        Dim lineoftext As String = String.Join("|", drow.ItemArray.Select(Function(s) s.ToString.Replace("|", "{PIPE_KEY}")).ToArray)
                        sw.WriteLine(lineoftext.Replace(vbCrLf, "\r\n").Replace(vbCr, "\r\n").Replace(vbLf, "\r\n"))
                    Next
                End Using
            End SyncLock
        Catch ex As Exception
            ' Ignore Issues
        End Try
    End Sub
    ' Save Board Information
    Friend Sub SaveBoardInformation()
        Try
            Dim FilePath As String = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "BoardInformation.txt"
            SyncLock BoardInformationLock
                Full_Campaign_Data.AcceptChanges()
                Using sw As New StreamWriter(FilePath, False)
                    For Each BoardInfo() As String In BoardInformation
                        Dim lineoftext As String = String.Join("|", BoardInfo)
                        sw.WriteLine(lineoftext.Replace(vbCrLf, "\r\n").Replace(vbCr, "\r\n").Replace(vbLf, "\r\n"))
                    Next
                End Using
            End SyncLock
        Catch ex As Exception
            ' Ignore Issues
        End Try
    End Sub
    ' Set Account As Disabled
    Friend Sub DisabledAccount(ByVal Username As String)
        SyncLock AccountUpdaterLock
            For Each Acc In Accounts
                If Acc(0) = Username Then
                    Acc(6) = "0"
                    Exit For
                End If
            Next
        End SyncLock
    End Sub
    ' Set Account As Read-Only Mode
    Friend Sub SetReadOnly(ByVal Username As String)
        SyncLock AccountUpdaterLock
            For Each Acc In Accounts
                If Acc(0) = Username Then
                    Acc(6) = "2"
                    Exit For
                End If
            Next
        End SyncLock
    End Sub
    ' Refresh Accounts
    Friend Sub RefreshAccounts()
        ComboBox1.Items.Clear()
        ComboBox9.Items.Clear()
        For Each TmpAccount() As String In Accounts
            ComboBox1.Items.Add(TmpAccount(0))
            ComboBox9.Items.Add(TmpAccount(0))
        Next
    End Sub

#End Region

#Region "DGV Functions"
    ' Record Cell Values
    Dim SelectedRows As DataGridViewSelectedRowCollection
    Dim Follower_SelectedRows As DataGridViewSelectedRowCollection
    ' Handle Data Errors For Board Manager DGV
    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal anError As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        Try
            Log_Show_Error(anError.Context.ToString, "Form1_DataGridView1_DataError", ActiveDatabase)
        Catch ex As Exception
            ' Ignore Issues
        End Try
    End Sub
    ' Handle Data Errors For Follower DGV
    Private Sub DataGridView2_DataError(ByVal sender As Object, ByVal anError As DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError
        Try
            Log_Show_Error(anError.Context.ToString, "Form1_DataGridView2_DataError", ActiveDatabase)
        Catch ex As Exception
            ' Ignore Issues
        End Try
    End Sub
    ' Manage Board Manager Selected Rows
    Private Sub DataGridView1_SelectedIndexChanged(sender As System.Object, e As MouseEventArgs) Handles DataGridView1.MouseUp
        Try
            SelectedRows = DataGridView1.SelectedRows
            If e.Button = MouseButtons.Right Then
                ContextMenuStrip1.Show(DataGridView1, e.Location)
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Manage Follower Selected Rows
    Private Sub DataGridView2_SelectedIndexChanged(sender As System.Object, e As MouseEventArgs) Handles DataGridView2.MouseUp
        Try
            Follower_SelectedRows = DataGridView2.SelectedRows
            If e.Button = MouseButtons.Right Then
                ' ContextMenuStrip1.Show(DataGridView2, e.Location)
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Edit Board Manager Selected Task
    Private Sub EditSelectedTaskToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditSelectedTaskToolStripMenuItem.Click
        Try
            If DataGridView1.SelectedRows.Count() > 0 Then
                ' Set Data To Process
                Dim DataToProcess As New List(Of String())
                ' Enable Selected Campaigns
                For Each TmpRow As DataGridViewRow In SelectedRows
                    Dim AccountUsername As String = TmpRow.Cells(0).Value.ToString()
                    Dim BoardName As String = TmpRow.Cells(1).Value.ToString()
                    ' Lookup Campaign Info
                    For Each TmpCampData As DataRow In Full_Campaign_Data.Rows
                        If TmpCampData("Account").ToString() = AccountUsername AndAlso TmpCampData("Board Name").ToString() = BoardName Then
                            ComboBox1.SelectedItem = AccountUsername
                            ComboBox2.SelectedItem = BoardName
                            TextBox1.Text = TmpCampData("Keywords").ToString().Replace(",", vbCrLf)
                            TextBox3.Text = TmpCampData("Website").ToString()
                            TextBox2.Text = TmpCampData("Delay").ToString()
                            Exit For
                        End If
                    Next
                    Exit For
                Next
            Else
                MessageBox.Show("No Rows Selected")
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Form Load/Close Events"
    ' Form Load
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' Randomize
            Randomize()
            ' Set Text
            Me.Text = SoftwareName & " V." & MyVersion & " - " & MyWebsite
            ' Hide Process Log
            Me.Size = New Size(1144, 497)
            ' Set Max Threads
            ' My.Settings.MaxThreads = 10
            ' Create Task DataTable
            CreateTaskDataTable()
            ' Create Following Task DataTable
            CreateFollowerTaskDataTable()
            ' Load Database
            LoadDatabase(ActiveDatabase)
            ' Load Current Database Names
            For Each Dir As String In Directory.GetDirectories(My.Application.Info.DirectoryPath)
                ComboBox4.Items.Add(Dir.Replace(My.Application.Info.DirectoryPath & "\", ""))
                ComboBox7.Items.Add(Dir.Replace(My.Application.Info.DirectoryPath & "\", ""))
            Next
            ' Add Create New Database Option
            ComboBox4.Items.Add("Create New Database")
            ComboBox7.Items.Add("Create New Database")
            ' Add Threads To ComboBox's
            For i As Integer = 1 To 100
                ComboBox6.Items.Add(i.ToString())
            Next
            ' Select Default Selections
            If ComboBox1.Items.Count() > 0 Then ComboBox1.SelectedIndex = 0
            If ComboBox2.Items.Count() > 0 Then ComboBox2.SelectedIndex = 0
            If ComboBox4.Items.Count() > 0 Then ComboBox4.SelectedIndex = 0
            If ComboBox7.Items.Count() > 0 Then ComboBox7.SelectedIndex = 0
            If ComboBox3.Items.Count() > 0 Then ComboBox3.SelectedIndex = 0
            If ComboBox5.Items.Count() > 0 Then ComboBox5.SelectedIndex = 0
            If ComboBox6.Items.Count() > 0 Then ComboBox6.SelectedIndex = 0
            ' Load Defaults
            LoadDefaults()
            ' Load Settings
            LoadSettings()
            ' Check Latest Version On Machine
            If GetSetting(SoftwareName.Replace(" ", "_"), "settings", "CurrentVersion", "") <> MyVersion Then
                ' If This Is The First Time New Version Has Run Delete Old Error Log Files
                Try
                    Dim newfiles() As String = Directory.GetFiles(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\error_logs\")
                    For Each TmpFile As String In newfiles
                        File.Delete(TmpFile)
                    Next
                Catch ex As Exception
                    ' Ignore Issues
                End Try
                ' Set Latest Version
                SaveSetting(SoftwareName.Replace(" ", "_"), "settings", "CurrentVersion", MyVersion)
            End If
            ' Check Ok
            If Unlocked = False Then End
            ' Check License
            Dim sb As StringBuilder = New StringBuilder
            ' Second Check
            If checker.m_ok = False Then Exit Sub
            sb.AppendLine(checker.m_key1)
            sb.AppendLine(checker.m_key)
            sb.AppendLine(checker.m_check)
            sb.AppendLine(checker.m_ok)
            sb.AppendLine(checker.m_result)
            If MainDomains <> MD5(sb.ToString()) Then Exit Sub
            ' Set Active Control
            Me.ActiveControl = GroupBox1
            ' Set Multi-Threading Delegates
            m_MainWindow = Me
            Form_AccountManager.m_MainForm = Me
        Catch ex As Exception
            MessageBox.Show("An Issue Occured Loading The Main Form - " & ex.Message)
        End Try
    End Sub
    ' Form Close Events
    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            ' Reset Process Log
            TextBox4.Text = ""
            ' Save Defaults
            SaveDefaults()
            ' End Program
            End
        Catch ex As Exception
            MessageBox.Show("An Issue Occured: " & ex.Message)
        End Try
    End Sub
    ' Load Defaults
    Private Sub LoadDefaults()
        ' Load Defaults
        Try
            Dim btn As Control = Me.GetNextControl(Me, True)
            While Not btn Is Nothing
                If TypeOf btn Is TextBox Then
                    If btn.Name <> "" Then
                        If GetSetting(SoftwareName.Replace(" ", "_"), "textboxes", btn.Name, "") <> Nothing Then btn.Text = GetSetting(SoftwareName.Replace(" ", "_"), "textboxes", btn.Name, "")
                    End If
                ElseIf TypeOf btn Is ComboBox Then
                    ' Ignore ComboBoxes As They Are Dynamic
                ElseIf TypeOf btn Is RichTextBox Then
                    If btn.Name <> "" Then
                        If GetSetting(SoftwareName.Replace(" ", "_"), "textboxes", btn.Name, "") <> Nothing Then btn.Text = GetSetting(SoftwareName.Replace(" ", "_"), "textboxes", btn.Name, "")
                    End If
                ElseIf TypeOf btn Is CheckBox Then
                    If btn.Name <> "" Then
                        If GetSetting(SoftwareName.Replace(" ", "_"), "checkboxes", btn.Name, "") <> Nothing Then TryCast(btn, CheckBox).Checked = CBool(GetSetting(SoftwareName.Replace(" ", "_"), "checkboxes", btn.Name, ""))
                    End If
                ElseIf TypeOf btn Is GroupBox Then
                    For Each ChildCtl As Control In btn.Controls
                        If TypeOf ChildCtl Is TextBox Then
                            If ChildCtl.Name <> "" Then
                                If GetSetting(SoftwareName.Replace(" ", "_"), "textboxes", ChildCtl.Name, "") <> Nothing Then ChildCtl.Text = GetSetting(SoftwareName.Replace(" ", "_"), "textboxes", ChildCtl.Name, "")
                            End If
                        ElseIf TypeOf ChildCtl Is ComboBox Then
                            ' Ignore ComboBoxes As They Are Dynamic
                        ElseIf TypeOf ChildCtl Is RichTextBox Then
                            If ChildCtl.Name <> "" Then
                                If GetSetting(SoftwareName.Replace(" ", "_"), "textboxes", ChildCtl.Name, "") <> Nothing Then ChildCtl.Text = GetSetting(SoftwareName.Replace(" ", "_"), "textboxes", ChildCtl.Name, "")
                            End If
                        ElseIf TypeOf btn Is CheckBox Then
                            If btn.Name <> "" Then
                                If GetSetting(SoftwareName.Replace(" ", "_"), "checkboxes", btn.Name, "") <> Nothing Then TryCast(btn, CheckBox).Checked = CBool(GetSetting(SoftwareName.Replace(" ", "_"), "checkboxes", btn.Name, ""))
                            End If
                        End If
                    Next
                End If
                btn = Me.GetNextControl(btn, True)
            End While
        Catch ex As Exception
            MessageBox.Show("An Issue Occured Loading Default Settings: " & ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub
    ' Save Defaults
    Private Sub SaveDefaults()
        ' Save Control Values
        Dim btn As Control = Me.GetNextControl(Me, True)
        While Not btn Is Nothing
            If TypeOf btn Is TextBox Then
                If btn.Name <> "" Then SaveSetting(SoftwareName.Replace(" ", "_"), "textboxes", btn.Name, btn.Text)
            ElseIf TypeOf btn Is ComboBox Then
                ' Ignore ComboBoxes As They Are Dynamic
            ElseIf TypeOf btn Is CheckBox Then
                Dim TmpBtn As CheckBox = CType(btn, CheckBox)
                Dim TrueFalse As String = "False"
                If TmpBtn.Checked = True Then TrueFalse = "True"
                If btn.Name <> "" Then SaveSetting(SoftwareName.Replace(" ", "_"), "checkboxes", TmpBtn.Name, TrueFalse)
            ElseIf TypeOf btn Is RichTextBox Then
                If btn.Name <> "" Then SaveSetting(SoftwareName.Replace(" ", "_"), "textboxes", btn.Name, btn.Text)
            ElseIf TypeOf btn Is GroupBox Then
                For Each ChildCtl As Control In btn.Controls
                    If TypeOf ChildCtl Is TextBox Then
                        If ChildCtl.Name <> "" Then SaveSetting(SoftwareName.Replace(" ", "_"), "textboxes", ChildCtl.Name, ChildCtl.Text)
                    ElseIf TypeOf ChildCtl Is ComboBox Then
                        ' Ignore ComboBoxes As They Are Dynamic
                    ElseIf TypeOf ChildCtl Is CheckBox Then
                        Dim TmpBtn As CheckBox = CType(ChildCtl, CheckBox)
                        Dim TrueFalse As String = "False"
                        If TmpBtn.Checked = True Then TrueFalse = "True"
                        If ChildCtl.Name <> "" Then SaveSetting(SoftwareName.Replace(" ", "_"), "checkboxes", TmpBtn.Name, TrueFalse)
                    ElseIf TypeOf ChildCtl Is RichTextBox Then
                        If ChildCtl.Name <> "" Then SaveSetting(SoftwareName.Replace(" ", "_"), "textboxes", ChildCtl.Name, ChildCtl.Text)
                    End If
                Next
            Else
                ' MessageBox.Show(btn.GetType.ToString())
            End If
            btn = Me.GetNextControl(btn, True)
        End While
    End Sub
#End Region

#Region "Board Manager Buttons"
    ' Load Database Button
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, Button18.Click
        Try
            Dim TmpDbNameToLoad As String = ""
            Dim btn As Button = CType(sender, Button)
            If btn.Name = "Button7" Then
                TmpDbNameToLoad = ComboBox4.SelectedItem.ToString()
            ElseIf btn.Name = "Button18" Then
                TmpDbNameToLoad = ComboBox7.SelectedItem.ToString()
            End If
            ActiveDatabase = TmpDbNameToLoad.Replace("/", "").Replace("\", "")
            LoadDatabase(ActiveDatabase)
            Form_AccountManager.RefreshAccounts()
            ' Set Current Database
            ComboBox4.SelectedItem = ActiveDatabase
            ComboBox7.SelectedItem = ActiveDatabase
            ' Select First Account
            If ComboBox1.Items.Count() > 0 Then ComboBox1.SelectedIndex = 0
            If ComboBox9.Items.Count() > 0 Then ComboBox9.SelectedIndex = 0
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Create New Database Option
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Try
            If ComboBox4.SelectedItem.ToString() = "Create New Database" Then
                Dim NewCampaignName As String = InputBox("Please Enter A Name For A New Database")
                If NewCampaignName <> Nothing Then
                    ' Check Does Not Already Exist
                    If ComboBox4.Items.Contains(NewCampaignName) Then
                        MessageBox.Show("You Already Have A Database With This Name")
                        Exit Sub
                    End If
                    ComboBox4.Items.Remove("Create New Database")
                    ComboBox4.Items.Add(NewCampaignName)
                    ComboBox4.Items.Add("Create New Database")
                    ComboBox4.SelectedItem = NewCampaignName
                    ComboBox7.Items.Remove("Create New Database")
                    ComboBox7.Items.Add(NewCampaignName)
                    ComboBox7.Items.Add("Create New Database")
                    ComboBox7.SelectedItem = NewCampaignName
                    ActiveDatabase = NewCampaignName
                    LoadDatabase(ActiveDatabase)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured: " & ex.Message)
        End Try
    End Sub
    ' Create New Database Option (Follower Tab)
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        Try
            If ComboBox7.SelectedItem.ToString() = "Create New Database" Then
                Dim NewCampaignName As String = InputBox("Please Enter A Name For A New Database")
                If NewCampaignName <> Nothing Then
                    ' Check Does Not Already Exist
                    If ComboBox7.Items.Contains(NewCampaignName) Then
                        MessageBox.Show("You Already Have A Database With This Name")
                        Exit Sub
                    End If
                    ComboBox4.Items.Remove("Create New Database")
                    ComboBox4.Items.Add(NewCampaignName)
                    ComboBox4.Items.Add("Create New Database")
                    ComboBox4.SelectedItem = NewCampaignName
                    ComboBox7.Items.Remove("Create New Database")
                    ComboBox7.Items.Add(NewCampaignName)
                    ComboBox7.Items.Add("Create New Database")
                    ComboBox7.SelectedItem = NewCampaignName
                    ActiveDatabase = NewCampaignName
                    LoadDatabase(ActiveDatabase)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured: " & ex.Message)
        End Try
    End Sub
    ' Add Task
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If ComboBox1.SelectedIndex = -1 Then
                MessageBox.Show("Please Select An Account To Use")
                Exit Sub
            End If
            If ComboBox2.SelectedIndex = -1 Then
                MessageBox.Show("Please Select A Board Name To Post To")
                Exit Sub
            End If
            If TextBox1.Text = "" Then
                MessageBox.Show("Please Enter Keywords To Search For Pins With")
                Exit Sub
            End If
            If TextBox3.Text = "" Then
                MessageBox.Show("Please Enter A Website URL")
                Exit Sub
            End If
            If TextBox2.Text = "" Then
                MessageBox.Show("Please Enter A Time To Create New Pins With")
                Exit Sub
            End If
            ' Set Variables
            Dim AccountName As String = ComboBox1.SelectedItem.ToString()
            Dim BoardName As String = ComboBox2.SelectedItem.ToString()
            Dim Keywords As String = Join(Split(TextBox1.Text, vbCrLf), ",")
            Dim WebsiteURL As String = TextBox3.Text
            Dim DelayTime As Integer = CInt(TextBox2.Text)
            ' Set Existing Variables
            Dim PinsCreated As String = "0"
            Dim NextPostTime As String = DateTime.Now.ToString()
            Dim TotalImpressions As String = "0"
            Dim UpdatingExistingTask As Boolean = False
            Dim tmprow As DataRow = Nothing
            ' Check Task Does Not Already Exist
            For Each row As DataRow In Full_Campaign_Data.Rows
                If CStr(row("Account")) = AccountName AndAlso CStr(row("Board Name")) = BoardName Then
                    Dim WarningMessage As String = "You Already Have A Task For This Account & Board" & vbCrLf & vbCrLf & "Do You Want To Update The Existing Task Data?"
                    Dim result As Integer = MessageBox.Show(WarningMessage, "", MessageBoxButtons.YesNo)
                    If result = DialogResult.Yes Then
                        ' Set As Updating
                        tmprow = row
                        UpdatingExistingTask = True
                        PinsCreated = row("Pins Created").ToString()
                        NextPostTime = row("Next Post Time").ToString()
                        TotalImpressions = row("Total Impressions").ToString()
                    Else
                        Exit Sub
                    End If
                End If
            Next
            ' Check If Updating
            If UpdatingExistingTask = True Then
                ' Delete Existing Task
                Full_Campaign_Data.Rows.Remove(tmprow)
            End If
            ' Add Task
            If NextPostUpdater.Enabled = True Then
                Full_Campaign_Data.Rows.Add(AccountName, BoardName, Keywords, WebsiteURL, DelayTime, "0", PinsCreated, NextPostTime, "Awaiting Timer", TotalImpressions)
            Else
                Full_Campaign_Data.Rows.Add(AccountName, BoardName, Keywords, WebsiteURL, DelayTime, "0", PinsCreated, NextPostTime, "Not Running", TotalImpressions)
            End If
            ' Show Message
            If UpdatingExistingTask = True Then
                MessageBox.Show("Task Successfully Updated")
            End If
            ' Backup Tasks
            SaveTasks()
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message & " - " & ex.StackTrace)
        End Try
    End Sub
    ' Start Task Scheduler
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            ' Confirm Tasks Available
            If Full_Campaign_Data.Rows.Count() < 1 Then
                MessageBox.Show("Please Create Tasks Before Starting Processes")
                Exit Sub
            End If
            ' Enable All Tasks
            UpdateCampaignData("ALL_CAMPAIGNS", "", "Status", "Awaiting Timer")
            m_MainWindow.UpdateDGV("ALL_CAMPAIGNS", "", "Status", "Awaiting Timer")
            ' Reset Variables
            BoardManager_StopProcess = False
            ' Reset Button
            Button2.Enabled = False
            Button3.Enabled = True
            ' Start Timer
            NextPostUpdater.Interval = CInt(3 * 1000) ' 3 Second Interval
            NextPostUpdater.Enabled = True
            NextPostUpdater_Tick(Nothing, Nothing)
        Catch ex As Exception
            MessageBox.Show("An Issue Occured:  " & ex.Message)
        End Try
    End Sub
    ' Stop Task Scheduler
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            ' Reset Variables
            BoardManager_StopProcess = True
            NextPostUpdater.Enabled = False
            ' Reset Button
            Button2.Enabled = True
            Button3.Enabled = False
            AddMessage("Waiting For All Running Tasks To End")
            UpdateDGV("ALL_CAMPAIGNS", "", "Status", "Not Running", "Awaiting Timer", True)
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Delete Selected Tasks
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            ' Set Data To Process
            Dim DataToProcess As New List(Of String())
            ' Disable Selected Tasks
            For Each TmpRow As DataGridViewRow In SelectedRows
                Dim AccountName As String = TmpRow.Cells(0).Value.ToString()
                Dim BoardName As String = TmpRow.Cells(1).Value.ToString()
                DataToProcess.Add(New String() {AccountName, BoardName})
            Next
            For Each TmpData() As String In DataToProcess
                Dim AccountName As String = TmpData(0)
                Dim BoardName As String = TmpData(1)
                ' Delete Task
                For i As Integer = 0 To (Full_Campaign_Data.Rows.Count() - 1)
                    If CStr(Full_Campaign_Data.Rows(i)("Account")) = AccountName AndAlso CStr(Full_Campaign_Data.Rows(i)("Board Name")) = BoardName Then
                        Full_Campaign_Data.Rows.RemoveAt(i)
                        Exit For
                    End If
                Next
            Next
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Account Manager
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, Button23.Click
        Try
            Form_AccountManager.WindowState = WindowState.Normal
            Form_AccountManager.Show()
            Form_AccountManager.Focus()
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' View/Hide Process Log
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, Button17.Click, Button21.Click
        Try
            If Button6.Text = "View Process Log" Then
                GroupBox3.Visible = True
                Me.Size = New Size(1144, 719)
                Button6.Text = "Hide Process Log"
                Button17.Text = "Hide Process Log"
                Button21.Text = "Hide Process Log"
            Else
                GroupBox3.Visible = False
                Me.Size = New Size(1144, 497)
                Button6.Text = "View Process Log"
                Button17.Text = "View Process Log"
                Button21.Text = "View Process Log"
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Selected Account Changed
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            ' Set Account
            Dim AccountUsername As String = ComboBox1.SelectedItem.ToString()
            RefreshBoards(AccountUsername)
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Selected Board Name Changed
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            For Each TmpTaskData As DataRow In Full_Campaign_Data.Rows
                If CStr(TmpTaskData("Board Name")) = ComboBox2.SelectedItem.ToString() Then
                    Dim Campaign_Keywords As String = CStr(TmpTaskData("Keywords"))
                    TextBox1.Text = TmpTaskData("Keywords").ToString().Replace(",", vbCrLf)
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Refresh Board Information
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            If ComboBox1.SelectedIndex = -1 Then
                MessageBox.Show("Please Select An Account First")
                Exit Sub
            End If
            LinkLabel1.Enabled = False
            LinkLabel1.Text = "Loading Boards..."
            LinkLabel1.Location = New Point(202, 60)
            AccountsToRefreshForBoardInformation.Clear()
            Dim AccountUsername As String = ComboBox1.SelectedItem.ToString()
            AccountsToRefreshForBoardInformation.Add(AccountUsername)
            ' Start Threaded Sub
            Dim objNewThread As New Thread(AddressOf RefreshBoardInformation)
            objNewThread.IsBackground = True
            objNewThread.Start()
        Catch ex As Exception
            LinkLabel1.Text = "Refresh Board Information"
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Follower Task Buttons"
    ' Selected Account Changed
    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        Try
            ' Set Account
            Dim AccountUsername As String = ComboBox9.SelectedItem.ToString()
            RefreshBoards(AccountUsername, "Follower")
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Add Follower Task
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Try
            If ComboBox9.SelectedIndex = -1 Then
                MessageBox.Show("Please Select An Account To Use")
                Exit Sub
            End If
            If TextBox10.Text = "" Then
                MessageBox.Show("Please Enter Keywords To Search For Pins With")
                Exit Sub
            End If
            If TextBox8.Text = "" OrElse TextBox11.Text = "" Then
                MessageBox.Show("Please Enter Follows To Send Per Run")
                Exit Sub
            End If
            If TextBox13.Text = "" OrElse TextBox12.Text = "" Then
                MessageBox.Show("Please Enter Delays Between Sending Follows")
                Exit Sub
            End If
            If TextBox9.Text = "" Then
                MessageBox.Show("Please Enter A Repeat Task Time Delay")
                Exit Sub
            End If
            ' Set Variables
            Dim AccountName As String = ComboBox9.SelectedItem.ToString()
            Dim Keywords As String = Join(Split(TextBox10.Text, vbCrLf), ",")
            Dim MinFollows As Integer = CInt(TextBox8.Text)
            Dim MaxFollows As Integer = CInt(TextBox11.Text)
            Dim MinDelay As Integer = CInt(TextBox13.Text)
            Dim MaxDelay As Integer = CInt(TextBox12.Text)
            Dim DelayTime As Integer = CInt(TextBox9.Text)
            ' Set Existing Variables
            Dim FollowsSent As String = "0"
            Dim NextTaskTime As String = DateTime.Now.ToString()
            Dim UpdatingExistingTask As Boolean = False
            Dim tmprow As DataRow = Nothing
            ' Check Task Does Not Already Exist
            For Each row As DataRow In Full_Following_Task_Data.Rows
                If CStr(row("Account")) = AccountName Then
                    Dim WarningMessage As String = "You Already Have A Task For This Account" & vbCrLf & vbCrLf & "Do You Want To Update The Existing Task Data?"
                    Dim result As Integer = MessageBox.Show(WarningMessage, "", MessageBoxButtons.YesNo)
                    If result = DialogResult.Yes Then
                        ' Set As Updating
                        tmprow = row
                        UpdatingExistingTask = True
                        FollowsSent = row("Follows Sent").ToString()
                        NextTaskTime = row("Next Task Time").ToString()
                    Else
                        Exit Sub
                    End If
                End If
            Next
            ' Check If Updating
            If UpdatingExistingTask = True Then
                ' Delete Existing Task
                Full_Following_Task_Data.Rows.Remove(tmprow)
            End If
            ' Add Task
            If NextPostUpdater.Enabled = True Then
                Full_Following_Task_Data.Rows.Add(AccountName, Keywords, MinFollows, MaxFollows, MinDelay, MaxDelay, "0", FollowsSent, NextTaskTime, "Awaiting Timer", DelayTime)
            Else
                Full_Following_Task_Data.Rows.Add(AccountName, Keywords, MinFollows, MaxFollows, MinDelay, MaxDelay, "0", FollowsSent, NextTaskTime, "Not Running", DelayTime)
            End If
            ' Show Message
            If UpdatingExistingTask = True Then
                MessageBox.Show("Task Successfully Updated")
            End If
            ' Backup Tasks
            SaveFollowerTasks()
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Start Follower Task Scheduler
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Try
            ' Confirm Tasks Available
            If Full_Following_Task_Data.Rows.Count() < 1 Then
                MessageBox.Show("Please Create Tasks Before Starting Processes")
                Exit Sub
            End If
            ' Enable All Tasks
            UpdateFollowerData("ALL_CAMPAIGNS", "Status", "Awaiting Timer")
            m_MainWindow.UpdateFollowerDGV("ALL_CAMPAIGNS", "Status", "Awaiting Timer")
            ' Reset Variables
            Follower_StopProcess = False
            ' Reset Button
            Button20.Enabled = False
            Button22.Enabled = True
            ' Start Timer
            FollowerTimer.Interval = CInt(3 * 1000) ' 3 Second Interval
            FollowerTimer.Enabled = True
            FollowerTimer_Tick(Nothing, Nothing)
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Stop Follower Task Scheduler
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Try
            ' Reset Variables
            Follower_StopProcess = True
            FollowerTimer.Enabled = False
            ' Reset Button
            Button20.Enabled = True
            Button22.Enabled = False
            AddMessage("Waiting For All Running Follower Tasks To End")
            UpdateFollowerDGV("ALL_CAMPAIGNS", "Status", "Not Running", "Awaiting Timer", True)
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Delete Selected Follower Tasks
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Try
            ' Set Data To Process
            Dim DataToProcess As New List(Of String)
            ' Disable Selected Tasks
            For Each TmpRow As DataGridViewRow In Follower_SelectedRows
                Dim AccountName As String = TmpRow.Cells(0).Value.ToString()
                DataToProcess.Add(AccountName)
            Next
            For Each TmpData As String In DataToProcess
                Dim AccountName As String = TmpData
                ' Delete Task
                For i As Integer = 0 To (Full_Following_Task_Data.Rows.Count() - 1)
                    If CStr(Full_Following_Task_Data.Rows(i)("Account")) = AccountName Then
                        Full_Following_Task_Data.Rows.RemoveAt(i)
                        Exit For
                    End If
                Next
            Next
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Account Creator Buttons"
    ' Load Proxies
    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Try
            MessageBox.Show("Please Load Proxies From A Text File, One Per Line In The Following Formats: " & vbCrLf & "Proxy:Port " & vbCrLf & "Proxy:Port:ProxyUsername:ProxyPassword" & vbCrLf & "E.g. 127.0.0.1:8888" & vbCrLf & "127.0.0.2:8888:user:pass")
            OpenFileDialog1.Filter = "*txt Text Files|*.txt"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                AccountCreator_Proxies.Clear()
                ' Set Temporary File Name To Allow Automatic Load On DB Load
                Dim TmpFileName As String = OpenFileDialog1.FileName
                ' Copy File
                Try
                    File.Copy(TmpFileName, My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_Proxies.txt")
                    TmpFileName = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_Proxies.txt"
                Catch ex As Exception
                    ' Ignore Issues
                End Try
                ' Read File
                Dim objReader As New System.IO.StreamReader(TmpFileName)
                Dim Line As String
                Dim AccountsLoaded As Integer = 0
                Do While objReader.Peek() <> -1
                    Line = objReader.ReadLine()
                    If Line = Nothing OrElse Line = "" OrElse Line = vbCrLf OrElse Line.Trim() = "" Then Continue Do
                    Line.Replace("	", "")
                    Dim Account As String() = Split(Line.Trim(), ":")
                    If Account.Count < 2 Then Continue Do
                    Dim Proxy As String = Account(0)
                    Dim Port As String = Account(1)
                    Dim i_ProxyUsername As String = "0"
                    Dim i_ProxyPassword As String = "0"
                    If Account.Count > 2 Then
                        i_ProxyUsername = Account(2)
                        i_ProxyPassword = Account(3)
                    End If
                    ' Check For Valid Port
                    If Integer.TryParse(Port, 0) = False Then Continue Do
                    ' Check To See If Already Exists
                    AccountCreator_Proxies.Add(New String() {Proxy, Port, i_ProxyUsername, i_ProxyPassword})
                    AccountsLoaded += 1
                Loop
                ' Randomize Proxies
                AccountCreator_Proxies.Sort(New Randomizer(Of String())())
                Label40.Text = AccountCreator_Proxies.Count().ToString()
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Load Usernames
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            MessageBox.Show("Please Load Usernames To Create From A Text File, One Username Per Line")
            OpenFileDialog1.Filter = "*txt Text Files|*.txt"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                AccountCreator_Usernames.Clear()
                If File.Exists(OpenFileDialog1.FileName) Then
                    ' Set Temporary File Name To Allow Automatic Load On DB Load
                    Dim TmpFileName As String = OpenFileDialog1.FileName
                    ' Copy File
                    Try
                        File.Copy(TmpFileName, My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_Usernames.txt")
                        TmpFileName = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_Usernames.txt"
                    Catch ex As Exception
                        ' Ignore Issues
                    End Try
                    ' Read File
                    Dim objReader As New System.IO.StreamReader(TmpFileName)
                    Dim Line As String
                    Do While objReader.Peek() <> -1
                        Line = objReader.ReadLine()
                        If Line = Nothing OrElse Line = "" Then Continue Do
                        Line.Replace("	", "")
                        If AccountCreator_UsernameBlacklist.Contains(Line) = False Then AccountCreator_Usernames.Add(Line)
                    Loop
                    objReader.Close()
                End If
                Label8.Text = AccountCreator_Usernames.Count().ToString()
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Load Email Accounts
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Try
            MessageBox.Show("Please Load Email Accounts From A Text File, One Per Line In The Following Formats: " & vbCrLf & "email@email.com " & vbCrLf & "email@email.com:email_password")
            OpenFileDialog1.Filter = "*txt Text Files|*.txt"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                AccountCreator_EmailAccounts.Clear()
                If File.Exists(OpenFileDialog1.FileName) Then
                    ' Set Temporary File Name To Allow Automatic Load On DB Load
                    Dim TmpFileName As String = OpenFileDialog1.FileName
                    ' Copy File
                    Try
                        File.Copy(TmpFileName, My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_EmailAccounts.txt")
                        TmpFileName = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_EmailAccounts.txt"
                    Catch ex As Exception
                        ' Ignore Issues
                    End Try
                    ' Read File
                    Dim objReader As New System.IO.StreamReader(TmpFileName)
                    Dim Line As String
                    Do While objReader.Peek() <> -1
                        Line = objReader.ReadLine()
                        If Line = Nothing OrElse Line = "" Then Continue Do
                        Line.Replace("	", "")
                        If AccountCreator_EmailBlacklist.Contains(Line) = False Then AccountCreator_EmailAccounts.Add(Line)
                    Loop
                    objReader.Close()
                End If
                Label9.Text = AccountCreator_EmailAccounts.Count().ToString()
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Load First Names
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Try
            MessageBox.Show("Please Load First Names From A Text File, One Per Line")
            OpenFileDialog1.Filter = "*txt Text Files|*.txt"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                AccountCreator_FirstNames.Clear()
                If File.Exists(OpenFileDialog1.FileName) Then
                    ' Set Temporary File Name To Allow Automatic Load On DB Load
                    Dim TmpFileName As String = OpenFileDialog1.FileName
                    ' Copy File
                    Try
                        File.Copy(TmpFileName, My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_FirstNames.txt")
                        TmpFileName = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_FirstNames.txt"
                    Catch ex As Exception
                        ' Ignore Issues
                    End Try
                    ' Read File
                    Dim objReader As New System.IO.StreamReader(TmpFileName)
                    Dim Line As String
                    Do While objReader.Peek() <> -1
                        Line = objReader.ReadLine()
                        If Line = Nothing OrElse Line = "" Then Continue Do
                        Line.Replace("	", "")
                        If AccountCreator_UsernameBlacklist.Contains(Line) = False Then AccountCreator_FirstNames.Add(Line)
                    Loop
                    objReader.Close()
                End If
                Label19.Text = AccountCreator_FirstNames.Count().ToString()
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Load Last Names
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Try
            MessageBox.Show("Please Load Last Names From A Text File, One Per Line")
            OpenFileDialog1.Filter = "*txt Text Files|*.txt"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                AccountCreator_LastNames.Clear()
                If File.Exists(OpenFileDialog1.FileName) Then
                    ' Set Temporary File Name To Allow Automatic Load On DB Load
                    Dim TmpFileName As String = OpenFileDialog1.FileName
                    ' Copy File
                    Try
                        File.Copy(TmpFileName, My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_LastNames.txt")
                        TmpFileName = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_LastNames.txt"
                    Catch ex As Exception
                        ' Ignore Issues
                    End Try
                    ' Read File
                    Dim objReader As New System.IO.StreamReader(TmpFileName)
                    Dim Line As String
                    Do While objReader.Peek() <> -1
                        Line = objReader.ReadLine()
                        If Line = Nothing OrElse Line = "" Then Continue Do
                        Line.Replace("	", "")
                        If AccountCreator_UsernameBlacklist.Contains(Line) = False Then AccountCreator_LastNames.Add(Line)
                    Loop
                    objReader.Close()
                End If
                Label21.Text = AccountCreator_LastNames.Count().ToString()
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Load Boards To Create
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Try
            MessageBox.Show("Please Load First Names From A Text File, One Per Line In The Following Format: " & vbCrLf & "BoardName:BoardCategory" & vbCrLf & "E.g. My Board Name:Women's Fashion")
            OpenFileDialog1.Filter = "*txt Text Files|*.txt"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                AccountCreator_BoardsToCreate.Clear()
                If File.Exists(OpenFileDialog1.FileName) Then
                    ' Set Temporary File Name To Allow Automatic Load On DB Load
                    Dim TmpFileName As String = OpenFileDialog1.FileName
                    ' Copy File
                    Try
                        File.Copy(TmpFileName, My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_BoardsToCreate.txt")
                        TmpFileName = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_BoardsToCreate.txt"
                    Catch ex As Exception
                        ' Ignore Issues
                    End Try
                    ' Read File
                    Dim objReader As New System.IO.StreamReader(TmpFileName)
                    Dim Line As String
                    Do While objReader.Peek() <> -1
                        Line = objReader.ReadLine()
                        If Line = Nothing OrElse Line = "" Then Continue Do
                        Line.Replace("	", "")
                        If AccountCreator_UsernameBlacklist.Contains(Line) = False Then AccountCreator_BoardsToCreate.Add(Line)
                    Loop
                    objReader.Close()
                End If
                Label38.Text = AccountCreator_BoardsToCreate.Count().ToString()
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Load Business Names
    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        Try
            MessageBox.Show("Please Load Business Names To Set From A Text File, One Business Name Per Line")
            OpenFileDialog1.Filter = "*txt Text Files|*.txt"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                AccountCreator_BusinessNames.Clear()
                If File.Exists(OpenFileDialog1.FileName) Then
                    ' Set Temporary File Name To Allow Automatic Load On DB Load
                    Dim TmpFileName As String = OpenFileDialog1.FileName
                    ' Copy File
                    Try
                        File.Copy(TmpFileName, My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_BusinessNames.txt")
                        TmpFileName = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\session\" & "AccountCreator_BusinessNames.txt"
                    Catch ex As Exception
                        ' Ignore Issues
                    End Try
                    ' Read File
                    Dim objReader As New System.IO.StreamReader(TmpFileName)
                    Dim Line As String
                    Do While objReader.Peek() <> -1
                        Line = objReader.ReadLine()
                        If Line = Nothing OrElse Line = "" Then Continue Do
                        Line.Replace("	", "")
                        If AccountCreator_BusinessNameBlacklist.Contains(Line) = False Then AccountCreator_BusinessNames.Add(Line)
                    Loop
                    objReader.Close()
                End If
                Label42.Text = AccountCreator_BusinessNames.Count().ToString()
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Select Profile Images Folder
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click, TextBox86.DoubleClick
        Try
            If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
                TextBox86.Text = FolderBrowserDialog1.SelectedPath
            End If
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Start Process
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Try
            ' Confirm All Data Available
            If AccountCreator_Usernames.Count() < 1 Then
                MessageBox.Show("Please Load Usernames")
                Exit Sub
            End If
            ' Confirm All Data Available
            If AccountCreator_BusinessNames.Count() < 1 Then
                MessageBox.Show("Please Load Business Names")
                Exit Sub
            End If
            ' Confirm All Data Available
            If AccountCreator_FirstNames.Count() < 1 Then
                MessageBox.Show("Please Load First Names")
                Exit Sub
            End If
            ' Confirm All Data Available
            If AccountCreator_LastNames.Count() < 1 Then
                MessageBox.Show("Please Load Last Names")
                Exit Sub
            End If
            ' Confirm All Data Available
            If AccountCreator_EmailAccounts.Count() < 1 Then
                MessageBox.Show("Please Load Email Accounts")
                Exit Sub
            End If
            ' Confirm All Data Available
            If TextBox86.Text = Nothing OrElse TextBox5.Text = Nothing OrElse TextBox6.Text = Nothing OrElse TextBox7.Text = Nothing Then
                MessageBox.Show("Please Enter Values Into All Available Text Boxes")
                Exit Sub
            End If
            ' Set Variables
            AccountCreator_ProfileImagesFolder = TextBox86.Text
            AccountCreator_ProfileGender = ComboBox3.SelectedItem.ToString()
            AccountCreator_MaxAccountsToCreate = CInt(TextBox5.Text)
            AccountCreator_DelayBetweenCreation = CInt(TextBox6.Text)
            AccountCreator_EmailVerifyDelay = CInt(TextBox7.Text)
            If CheckBox2.Checked = True Then
                AccountCreator_TempProxies = AccountCreator_Proxies
            End If
            ' Set Friendly Name For Business Type
            If ComboBox5.SelectedItem.ToString() = "Blogger" Then
                AccountCreator_BusinessType = "blogger"
            ElseIf ComboBox5.SelectedItem.ToString() = "Consumer Good, Product, or Service" Then
                AccountCreator_BusinessType = "consumer_good_product_or_service"
            ElseIf ComboBox5.SelectedItem.ToString() = "Contractor or Service Provider" Then
                AccountCreator_BusinessType = "contractor_or_service_provider"
            ElseIf ComboBox5.SelectedItem.ToString() = "Influencer, Public Figure, or Celebrity" Then
                AccountCreator_BusinessType = "influencer_public_figure_or_celebrity"
            ElseIf ComboBox5.SelectedItem.ToString() = "Institution/Non-profit" Then
                AccountCreator_BusinessType = "institution_or_non_prof"
            ElseIf ComboBox5.SelectedItem.ToString() = "Local Retail Store" Then
                AccountCreator_BusinessType = "local_retail_store"
            ElseIf ComboBox5.SelectedItem.ToString() = "Local Service" Then
                AccountCreator_BusinessType = "local_service"
            ElseIf ComboBox5.SelectedItem.ToString() = "Online Retail or Marketplace" Then
                AccountCreator_BusinessType = "online_retail_or_marketplace"
            ElseIf ComboBox5.SelectedItem.ToString() = "Other" Then
                AccountCreator_BusinessType = "other"
            ElseIf ComboBox5.SelectedItem.ToString() = "Publisher or Media" Then
                AccountCreator_BusinessType = "publisher_or_media"
            ElseIf ComboBox5.SelectedItem.ToString() = "Brand" Then
                AccountCreator_BusinessType = "brand"
            ElseIf ComboBox5.SelectedItem.ToString() = "Institution/Non-profit" Then
                AccountCreator_BusinessType = "institution_or_non_prof"
            ElseIf ComboBox5.SelectedItem.ToString() = "Local Business" Then
                AccountCreator_BusinessType = "local_business"
            ElseIf ComboBox5.SelectedItem.ToString() = "Media" Then
                AccountCreator_BusinessType = "media"
            ElseIf ComboBox5.SelectedItem.ToString() = "Online Marketplace" Then
                AccountCreator_BusinessType = "online_marketplace"
            ElseIf ComboBox5.SelectedItem.ToString() = "Professional" Then
                AccountCreator_BusinessType = "professional"
            ElseIf ComboBox5.SelectedItem.ToString() = "Public Figure" Then
                AccountCreator_BusinessType = "public_figure"
            ElseIf ComboBox5.SelectedItem.ToString() = "Retailer" Then
                AccountCreator_BusinessType = "retailer"
            End If
            ' Reset Variables
            AccountCreator_StopProcess = False
            Button16.Enabled = False
            Button14.Enabled = True
            AccountCreator_AccountsCreated = 0
            AccountCreator_AccountsVerified = 0
            AccountCreator_AutoAddAccounts = CheckBox1.Checked
            ' Enable Verifier Timer
            EmailVerifierTimer.Interval = 5000
            EmailVerifierTimer.Enabled = True
            EmailVerifierTimer.Start()
            ' Start New Class
            Dim AccountCreatorClass As New AccountCreator(Me, ActiveDatabase, CheckBox2.Checked)
            ' Start Threads
            For i As Integer = 1 To ComboBox6.SelectedItem.ToString()
                Dim objNewThread As New Thread(AddressOf AccountCreatorClass.RunProcess)
                objNewThread.IsBackground = True
                objNewThread.Start()
                AccountCreator_ThreadsRunning += 1
            Next
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message & " - " & ex.StackTrace)
        End Try
    End Sub
    ' Stop Process
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Try
            ' Reset Variables
            AccountCreator_StopProcess = True
            AddMessage("Waiting For All Account Creator Threads To End")
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Export Unused Email Addresses
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Try
            ' Open CSV File
            Dim strFile As String = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\exports\" & "Unused_EmailAccounts.txt"
            Dim objStreamWriter As StreamWriter
            objStreamWriter = New StreamWriter(strFile, False, Encoding.Unicode)
            ' Loop Accounts
            For Each TmpLine As String In AccountCreator_EmailAccounts
                If TmpLine = Nothing Then Continue For
                objStreamWriter.WriteLine(TmpLine)
            Next
            ' Close File
            objStreamWriter.Close()
            ' Read File
            System.Diagnostics.Process.Start(strFile)
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Export Unused Usernames
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Try
            ' Open CSV File
            Dim strFile As String = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\exports\" & "Unused_Usernames.txt"
            Dim objStreamWriter As StreamWriter
            objStreamWriter = New StreamWriter(strFile, False, Encoding.Unicode)
            ' Loop Accounts
            For Each TmpLine As String In AccountCreator_Usernames
                If TmpLine = Nothing Then Continue For
                objStreamWriter.WriteLine(TmpLine)
            Next
            ' Close File
            objStreamWriter.Close()
            ' Read File
            System.Diagnostics.Process.Start(strFile)
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Export Created Accounts
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Try
            ' Open CSV File
            Dim strFile As String = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\exports\" & "Created_Accounts.txt"
            Dim objStreamWriter As StreamWriter
            objStreamWriter = New StreamWriter(strFile, False, Encoding.Unicode)
            ' Write Column Titles
            objStreamWriter.WriteLine("{0}:{1}:{2}:{3}:{4}:{5}:{6}:{7}", "Username", "Password", "Proxy", "Port", "Proxy Username", "Proxy Password", "Email Address", "Email Password")
            ' Loop Accounts
            For Each Acc As String() In AccountCreator_CreatedAccounts
                objStreamWriter.WriteLine("{0}:{1}:{2}:{3}:{4}:{5}:{6}:{7}", Acc(0), Acc(1), Acc(2), Acc(3), Acc(4), Acc(5), Acc(6), Acc(7))
            Next
            ' Close File
            objStreamWriter.Close()
            ' Read File
            System.Diagnostics.Process.Start(strFile)
        Catch ex As Exception
            MessageBox.Show("An Issue Occured: " & ex.Message)
        End Try
    End Sub
    ' Export Verified Accounts
    Private Sub Button56_Click(sender As Object, e As EventArgs) Handles Button56.Click
        Try
            ' Open CSV File
            Dim strFile As String = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\exports\" & "Verified_Accounts.txt"
            Dim objStreamWriter As StreamWriter
            objStreamWriter = New StreamWriter(strFile, False, Encoding.Unicode)
            ' Write Column Titles
            objStreamWriter.WriteLine("{0}:{1}:{2}:{3}:{4}:{5}:{6}:{7}", "Username", "Password", "Proxy", "Port", "Proxy Username", "Proxy Password", "Email Address", "Email Password")
            ' Loop Accounts
            For Each Acc As String() In AccountCreator_CreatedAccounts
                If Acc(8) <> "1" Then Continue For
                objStreamWriter.WriteLine("{0}:{1}:{2}:{3}:{4}:{5}:{6}:{7}", Acc(0), Acc(1), Acc(2), Acc(3), Acc(4), Acc(5), Acc(6), Acc(7))
            Next
            ' Close File
            objStreamWriter.Close()
            ' Read File
            System.Diagnostics.Process.Start(strFile)
        Catch ex As Exception
            MessageBox.Show("An Issue Occured: " & ex.Message)
        End Try
    End Sub
    ' Export Un-Verified Accounts
    Private Sub Button74_Click(sender As Object, e As EventArgs) Handles Button74.Click
        Try
            ' Open CSV File
            Dim strFile As String = My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\exports\" & "Unverified_Accounts.txt"
            Dim objStreamWriter As StreamWriter
            objStreamWriter = New StreamWriter(strFile, False, Encoding.Unicode)
            ' Write Column Titles
            objStreamWriter.WriteLine("{0}:{1}:{2}:{3}:{4}:{5}:{6}:{7}", "Username", "Password", "Proxy", "Port", "Proxy Username", "Proxy Password", "Email Address", "Email Password")
            ' Loop Accounts
            For Each Acc As String() In AccountCreator_CreatedAccounts
                If Acc(8) <> "0" Then Continue For
                objStreamWriter.WriteLine("{0}:{1}:{2}:{3}:{4}:{5}:{6}:{7}", Acc(0), Acc(1), Acc(2), Acc(3), Acc(4), Acc(5), Acc(6), Acc(7))
            Next
            ' Close File
            objStreamWriter.Close()
            ' Read File
            System.Diagnostics.Process.Start(strFile)
        Catch ex As Exception
            MessageBox.Show("An Issue Occured: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Settings Manager"
    ' Save All Settings Button
    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        Try
            My.Settings.PageLoadDelay = CInt(TextBox15.Text)
            My.Settings.Useragents = TextBox14.Text
            My.Settings.MaxThreads = CInt(TextBox16.Text)
            My.Settings.MaxFollowerThreads = CInt(TextBox17.Text)
            My.Settings.MaxPages = CInt(TextBox18.Text)
            My.Settings.MaxTimeout = CInt(TextBox19.Text)
            SaveSettings()
            MessageBox.Show("Settings Saved")
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
    ' Load Settings
    Private Sub LoadSettings()
        If GetSetting(SoftwareName.Replace(" ", "_"), "settings", "MaxThreads", "") <> Nothing Then My.Settings.MaxThreads = CInt(GetSetting(SoftwareName.Replace(" ", "_"), "settings", "MaxThreads", ""))
        If GetSetting(SoftwareName.Replace(" ", "_"), "settings", "Useragents", "") <> Nothing Then My.Settings.Useragents = GetSetting(SoftwareName.Replace(" ", "_"), "settings", "Useragents", "")
        If GetSetting(SoftwareName.Replace(" ", "_"), "settings", "PageLoadDelay", "") <> Nothing Then My.Settings.PageLoadDelay = CInt(GetSetting(SoftwareName.Replace(" ", "_"), "settings", "PageLoadDelay", ""))
        If GetSetting(SoftwareName.Replace(" ", "_"), "settings", "DisplayMessages", "") <> Nothing Then My.Settings.DisplayMessages = CBool(GetSetting(SoftwareName.Replace(" ", "_"), "settings", "DisplayMessages", ""))
        If GetSetting(SoftwareName.Replace(" ", "_"), "settings", "MaxFollowerThreads", "") <> Nothing Then My.Settings.MaxFollowerThreads = CInt(GetSetting(SoftwareName.Replace(" ", "_"), "settings", "MaxFollowerThreads", ""))
        If GetSetting(SoftwareName.Replace(" ", "_"), "settings", "MaxPages", "") <> Nothing Then My.Settings.MaxPages = CInt(GetSetting(SoftwareName.Replace(" ", "_"), "settings", "MaxPages", ""))
        If GetSetting(SoftwareName.Replace(" ", "_"), "settings", "MaxTimeout", "") <> Nothing Then My.Settings.MaxTimeout = CInt(GetSetting(SoftwareName.Replace(" ", "_"), "settings", "MaxTimeout", ""))
        ' Set Data In Settings Tab
        TextBox15.Text = CStr(My.Settings.PageLoadDelay)
        TextBox14.Text = My.Settings.Useragents
        TextBox16.Text = CStr(My.Settings.MaxThreads)
        TextBox17.Text = CStr(My.Settings.MaxFollowerThreads)
        TextBox18.Text = CStr(My.Settings.MaxPages)
        TextBox19.Text = CStr(My.Settings.MaxTimeout)
    End Sub
    ' Save Settings
    Private Sub SaveSettings()
        Try
            SaveSetting(SoftwareName.Replace(" ", "_"), "settings", "MaxThreads", CStr(My.Settings.MaxThreads))
            SaveSetting(SoftwareName.Replace(" ", "_"), "settings", "Useragents", CStr(My.Settings.Useragents))
            SaveSetting(SoftwareName.Replace(" ", "_"), "settings", "PageLoadDelay", CStr(My.Settings.PageLoadDelay))
            SaveSetting(SoftwareName.Replace(" ", "_"), "settings", "DisplayMessages", CStr(My.Settings.DisplayMessages))
            SaveSetting(SoftwareName.Replace(" ", "_"), "settings", "MaxFollowerThreads", CStr(My.Settings.MaxFollowerThreads))
            SaveSetting(SoftwareName.Replace(" ", "_"), "settings", "MaxPages", CStr(My.Settings.MaxPages))
            SaveSetting(SoftwareName.Replace(" ", "_"), "settings", "MaxTimeout", CStr(My.Settings.MaxTimeout))
        Catch ex As Exception
            MessageBox.Show("Failed To Save Settings To Registry: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Site Functions"
    ' Pinterest - Get Board ID
    Friend Function GetBoardId(ByVal AccountUsername As String, ByVal BoardName As String) As String
        For Each TmpBoardInfo() As String In BoardInformation
            If TmpBoardInfo(0) = AccountUsername AndAlso TmpBoardInfo(1) = BoardName Then
                Return TmpBoardInfo(2)
            End If
        Next
        ' Nothing Found
        Throw New Exception("Board ID Does Not Exist In System - Refresh Board Information Or Check Board Name Still Exists (Board Name: " & BoardName & ")")
    End Function
    ' Pinterest - Login
    Friend Sub Login(ByRef a_Process As CurlFunctions, ByRef Username As String, ByVal Password As String, ByVal EmailAddress As String, ByRef UserID As String)
        ' Check Info
        Dim MainDomain As String = "pinterest.com"
        ' Check Login System
        Dim TmpUsername As String = EmailAddress
        ' Variables
        Dim URL As String
        Dim Referrer As String
        Dim Data As String
        Dim Content As String
        Dim Attempt As Integer = 0
        Dim UKIP As Boolean = False
        ' Load Login Page
        URL = "https://www." & MainDomain & "/"
        Referrer = "https://www." & MainDomain & "/"
        Try
            Content = a_Process.Return_Content(URL, "")
            Content = a_Process.Return_Content(URL, "")
        Catch ex As Exception
            Throw New Exception("Failed To Load Login Page - " & ex.Message)
        End Try
        ' Check For UK IP
        If Content.Contains("pinterest.com") Then
            UKIP = True
            ' Load Login Page
            URL = "https://www." & MainDomain.Replace(".com", ".com") & "/"
            Referrer = "https://www." & MainDomain & "/"
            Try
                Content = a_Process.Return_Content(URL, "")
            Catch ex As Exception
                Throw New Exception("Failed To Load Login Page - " & ex.Message)
            End Try
        End If
        ' Scrape App Version
        Dim AppVersion As String = GetBetween(Content, """app_version"": """, """")
        If AppVersion = Nothing Then AppVersion = GetBetween(Content, """app_version"":""", """")
        ' Check For Bot Page
        If Content.Contains("We've detected a bot") Then Throw New Exception("IP: " & a_Process.Proxy & " Has Been Detected As Spam / A Bot IP")
        ' Check To See If Already Logged In
        If Content.Contains("LoggedInUser(true)") OrElse Content.Contains("userIsAuthenticated = true") OrElse Content.Contains("isLoggedIn"": true") OrElse Content.Contains("isLoggedIn"":true") Then
            ' Check For Read-Only Mode Account
            If Content.Contains("read-only mode") Then
                Throw New Exception("Account Is In Read-Only Mode")
            End If
            ' Scrape User ID
            UserID = GetBetween(Content, """isAuth"":true,""id"":""", """")
            If UserID = Nothing Then UserID = GetBetween(Content, """type"": ""partner"", ""id"": """, """")
            Username = GetBetween(Content, UserID & """,""username"":""", """")
            If Username = Nothing Then Username = GetBetween(Content, """profile"", ""username"": """, """")
            If Username = Nothing Then Username = GetBetween(Content, "is_employee"": false, ""username"": """, """")
            ' Confirm Available
            If UserID = Nothing OrElse Username = Nothing Then
                Log_Show_Error(Content, "6", ActiveDatabase)
                Throw New Exception("Failed To Scrape Username, Unknown Issue")
            End If
            ' Report Success
            AddMessage("Account: " & Username & " Already Logged In")
            Exit Sub
        End If
        ' Scrape Token
        Dim token As String = GetBetween(Content, "name='csrfmiddlewaretoken' value='", "'")
        If UKIP = True Then
            If token = Nothing Then token = GetSessionToken(a_Process, "pinterest.com")
        Else
            If token = Nothing Then token = GetSessionToken(a_Process)
        End If
        ' Confirm Token Exists
        If token = Nothing Then
            Log_Show_Error(Content, "7", ActiveDatabase)
            Throw New Exception("Failed To Scrape Login Token")
        End If
        ' Scrape Next Page
        Dim NextPage As String = GetBetween(Content, "name=""next"" value=""", """")
        ' Set Referrer
        If UKIP = True Then
            Referrer = "https://www.pinterest.com/"
        Else
            Referrer = "https://www.pinterest.com/"
        End If
        ' Post To Handshake URL
        URL = "https://accounts.pinterest.com/v3/login/handshake/"
        Data = "username_or_email=" & URLEncode(EmailAddress) & "&password=" & URLEncode(Password) & "&"
        Try
            Content = a_Process.Post_Content(URL, Data, Referrer, False, False, False, "*/*")
        Catch ex As Exception
            Throw New Exception("Failed To Post Login Data - " & ex.Message)
        End Try
        ' Scrape Token
        Dim LoginToken As String = GetBetween(Content, """data"": """, """")
        ' Confirm Available
        If LoginToken = Nothing Then
            GoTo CheckError
        End If
        ' Post Login Information
        If UKIP = True Then
            URL = "https://www.pinterest.com/resource/HandshakeSessionResource/create/"
            Referrer = "https://www.pinterest.com/"
        Else
            URL = "https://www.pinterest.com/resource/HandshakeSessionResource/create/"
            Referrer = "https://www.pinterest.com/"
        End If
        Data = "source_url=%2F&data=%7B%22options%22%3A%7B%22token%22%3A%22" & LoginToken & "%22%7D%2C%22context%22%3A%7B%7D%7D"
        Try
            Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, token, Referrer, True, AppVersion, "active")
        Catch ex As Exception
            If ex.Message.Contains("404") OrElse ex.Message.Contains("401") Then
                DisabledAccount(EmailAddress)
                Throw New Exception("Incorrect Login Details")
            End If
            Throw New Exception("Failed To Post Login Data - " & ex.Message)
        End Try
        ' Confirm Ok
        If Not Content.Contains("error"": null") AndAlso Not Content.Contains("error"":null") Then
CheckError:
            ' Check For Bot Page
            If Content.Contains("We've detected a bot") Then
                Throw New Exception("IP: " & a_Process.Proxy & " Has Been Detected As Spam / A Bot IP")
            ElseIf Content.Contains("been suspended") Then
                DisabledAccount(EmailAddress)
                Throw New Exception("Account: " & Username & " Has Been Suspended From Pinterest")
            ElseIf Content.Contains("User not found") OrElse Content.Contains("password you entered") OrElse Content.Contains("Enter a valid email address") OrElse Content.Contains("we couldn't verify your email and password") OrElse Content.Contains("recognize that email") OrElse Content.Contains("Invalid email or password") OrElse Content.Contains("wrong email or password") OrElse Content.Contains("password you entered is incorrect") OrElse Content.Contains("email you entered does not belong to any account") Then
                DisabledAccount(EmailAddress)
                Throw New Exception("Incorrect Login Details")
            ElseIf Content.Contains("Login with username is currently disabled") Then
                Throw New Exception("Login with username is currently disabled")
            ElseIf Content.Contains("exceeded your rate limit") Then
                Throw New Exception("Login Rate Limit Exceeded From This IP (" & a_Process.Proxy & ")")
            ElseIf Content.Contains("reset your password") Then
                SetReadOnly(Username)
                Throw New Exception("Account Needs Password Reset Via Email")
            Else
                Log_Show_Error(Data & vbCrLf & vbCrLf & Content, "8", ActiveDatabase)
                Throw New Exception("Unknown Issue")
            End If
        End If
        ' Load Homepage
        URL = "https://www." & MainDomain & "/"
        Try
            Content = a_Process.Return_Content(URL)
        Catch ex As Exception
            Throw New Exception("Failed To Load Homepage - " & ex.Message)
        End Try
        ' Confirm Logged In
        If Not Content.Contains("LoggedInUser(true)") AndAlso Not Content.Contains("userIsAuthenticated = true") AndAlso Not Content.Contains("isLoggedIn"": true") AndAlso Not Content.Contains("isLoggedIn"":true") Then
            ' Check For Disabled Acount
            If Content.Contains("been deactivated") = True Then
                DisabledAccount(EmailAddress)
                Throw New Exception("Account Disabled")
            End If
            ' Check For Read-Only Mode Account
            If Content.Contains("read-only mode") Then
                SetReadOnly(Username)
                Throw New Exception("Account In Read-Only Mode")
            End If
            ' Find Other Issues
            If Content.Contains("we couldn't verify your email and password") OrElse Content.Contains("recognize that email") Then
                DisabledAccount(EmailAddress)
                Throw New Exception("Incorrect Login Details")
            End If
            If Content.Contains("Login with username is currently disabled") Then Throw New Exception("Login with username is currently disabled")
            If Content.Contains("exceeded your rate limit") Then Throw New Exception("Login Rate Limit Exceeded From This IP (" & a_Process.Proxy & ")")
            If Content.Contains("Enter a valid email address") Then
                DisabledAccount(EmailAddress)
                Throw New Exception("Account: " & Username & " Not A Valid Username / Email Address")
            End If
            Log_Show_Error(Content, "9", ActiveDatabase)
            Throw New Exception("Unknown Issue")
        End If
        ' Scrape User ID
        UserID = GetBetween(Content, """isAuth"":true,""id"":""", """")
        If UserID = Nothing Then UserID = GetBetween(Content, """type"": ""partner"", ""id"": """, """")
        Username = GetBetween(Content, UserID & """,""username"":""", """")
        If Username = Nothing Then Username = GetBetween(Content, """profile"", ""username"": """, """")
        If Username = Nothing Then Username = GetBetween(Content, """is_employee"": false, ""username"": """, """")
        ' Confirm Available
        If UserID = Nothing OrElse Username = Nothing Then
            Log_Show_Error(Content, "11", ActiveDatabase)
            Throw New Exception("Failed To Scrape Username, Unknown Issue")
        End If
        ' Report Success
        AddMessage("Successfully Logged Into Account: " & Username)
        ' Save Cookies
        a_Process.Dispose()
    End Sub
    ' Pinterest - Upload Image
    Friend Function UploadImage(ByRef a_Process As CurlFunctions, ByVal Username As String, ByVal ProfileImagePath As String, ByVal BoardID As String, ByVal PinTitle As String, ByVal PinDescription As String, ByVal New_PinLink As String) As String
        ' Set Info
        Dim MainDomain As String = "pinterest.com"
        ' Variables
        Dim URL As String
        Dim Data As String
        Dim Referrer As String
        Dim Content As String
        Dim PinIDNum As String = ""
        Dim AppVersion As String = ""
        Dim UploadedImage As String = ""


        Try
            ' Format Variables
            If PinTitle <> Nothing Then PinTitle = PinTitle.Replace("\", "")
            If PinDescription <> Nothing Then PinDescription = PinDescription.Replace("\", "")
            ' Load Homepage
            URL = "https://www." & MainDomain & "/"
            Content = a_Process.Return_Content(URL)
            ' Scrape App Version
            AppVersion = GetBetween(Content, """app_version"": """, """")
            If AppVersion = Nothing Then AppVersion = GetBetween(Content, """app_version"":""", """")
            ' Upload Pin
            Try
                Content = a_Process.UploadImage(ProfileImagePath, GetSessionToken(a_Process), Username)
            Catch ex As Exception
                Throw New Exception("Failed To Upload Image - " & ex.Message)
            End Try
        Catch ex As Exception
            Throw New Exception("A - " & ex.Message)
        End Try


        Try
            ' Scrape Uploaded Image URL
            UploadedImage = GetBetween(Content, """image_url"": """, """")
            ' Confirm Available
            If UploadedImage = Nothing Then
                If Content.Contains("image is too small") Then
                    Throw New Exception("Failed To Upload Image For Account: " & Username & " - Image File Was Rejected By Pinterest")
                Else
                    Log_Show_Error(Content, "UploadImage_2", ActiveDatabase)
                    Throw New Exception("Failed To Upload Image For Account: " & Username & " - Unknown Issue")
                End If
            End If
            ' Format Data
            If New_PinLink = Nothing Then New_PinLink = ""
            Try
                If PinTitle.Length() > 100 Then PinTitle = PinTitle.Substring(0, 99)
            Catch ex As Exception
                ' Ignore Issues
            End Try
        Catch ex As Exception
            Throw New Exception("B: " & ex.Message)
        End Try

        Try

            ' Upload Pin
            URL = "https://www." & MainDomain & "/resource/PinResource/create/"
            Data = "source_url=%2Fpin-builder%2F&data=%7B%22options%22%3A%7B%22board_id%22%3A%22" & BoardID & "%22%2C%22field_set_key%22%3A%22create_success%22%2C%22skip_pin_create_log%22%3Atrue%2C%22description%22%3A%22" & URLEncode(PinDescription) & "%22%2C%22link%22%3A%22" & URLEncode(New_PinLink) & "%22%2C%22title%22%3A%22" & URLEncode(PinTitle) & "%22%2C%22image_url%22%3A%22" & URLEncode(UploadedImage) & "%22%2C%22method%22%3A%22uploaded%22%2C%22upload_metric%22%3A%7B%22source%22%3A%22partner_upload_standalone%22%7D%7D%2C%22context%22%3A%7B%7D%7D"
            Referrer = "https://www." & MainDomain & "/"
            Try
                Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
            Catch ex As Exception
                Log_Show_Error(URL & vbCrLf & vbCrLf & Data & vbCrLf & vbCrLf & ex.Message, "UploadImage_3", ActiveDatabase)
                Throw New Exception(ex.Message)
            End Try
        Catch ex As Exception
            Throw New Exception("C: " & ex.Message)
        End Try


        Try
            ' Confirm Pin OK
            If Not Content.Contains("""error"": null") AndAlso Not Content.Contains("""error"":null") Then
                Dim ErrMsg As String = DetectPinterestError(Content)
                If ErrMsg <> Nothing Then
                    If ErrMsg.Contains("The image url is not valid") OrElse ErrMsg.Contains("not a valid image") Then Throw New Exception("Invalid Image URL")
                    If ErrMsg.Contains("You are pinning really fast") OrElse ErrMsg.Contains("combat spam") Then Throw New Exception("You Are Pinning Too Fast")
                    If ErrMsg.Contains("We love your enthusiasm") Then Throw New Exception("You Are Pinning Too Fast")
                    If ErrMsg.Contains("Enter a valid URL") Then Throw New Exception("Please Enter A Valid URL")
                    If ErrMsg.Contains("hiccups") Then Throw New Exception("Pin Failed, Pinterest Issue")
                    If ErrMsg.Contains("could not fetch the image") Then Throw New Exception("Pin Failed, Pinterest Could Not Fetch The Image")
                    If ErrMsg.Contains("account has been deactivated") Then Throw New Exception("Account Banned")
                    If ErrMsg.Contains("Users have reported") Then Throw New Exception("Link Has Been Reported As Spam")
                    If ErrMsg.Contains("safe mode to protect") Then Throw New Exception("Account Is In Safe Mode")
                    If ErrMsg.Contains("request could not be completed") Then Throw New Exception("Pinterest Reported That The Request Could Not Be Completed")
                    Throw New Exception("Failed To Post Pin - New Issue: " & ErrMsg)
                End If
                If Content.Contains("We've detected a bot") Then Throw New Exception("BOT DETECTION")
                Log_Show_Error(Content, "UploadImage_4", ActiveDatabase)
                Throw New Exception("Failed To Post Pin - Unknown Issue")
            End If
        Catch ex As Exception
            Throw New Exception("D: " & ex.Message)
        End Try

        Try
            ' Scrape Pin ID
            PinIDNum = GetBetween(Content, """pin"",""id"":""", """")
            If PinIDNum = Nothing Then PinIDNum = GetBetween(Content, """pin"",""id"":""", """")
            If PinIDNum = Nothing Then PinIDNum = GetBetween(Content, """object_id_str"":""", """")
            If PinIDNum = Nothing Then PinIDNum = GetBetween(Content, """section"":null,""id"":""", """")
        Catch ex As Exception
            Throw New Exception("E: " & ex.Message)
        End Try
        ' Successfully Added Pin
        Return PinIDNum
    End Function
    ' Pinterest - Get Pin Information
    Friend Sub GetPinInformation(ByRef a_Process As CurlFunctions, ByVal PinID As String, ByRef PinTitle As String, ByRef PinLink As String, ByRef PinImageLink As String, ByRef PinDescription As String)
        ' Variables
        Dim URL As String
        Dim Content As String
        ' Load Pin URL
        URL = "https://www.pinterest.com/pin/" & PinID & "/"
        Content = a_Process.Return_Content(URL)
        ' Scrape Data
        PinTitle = GetBetween(Content, """grid_title"": """, """")
        PinLink = GetBetween(Content, """link"": """, """")
        PinImageLink = GetBetween(Content, """736x"": {""url"": """, """")
        PinDescription = GetBetween(Content, """description"": """, """")
        If PinDescription = Nothing Then PinDescription = GetBetween(Content, """closeup_user_note"": """, """")
        If PinDescription = Nothing Then PinDescription = GetBetween(Content, "null, ""description"": """, """")
        ' Strip Tags
        If PinDescription <> Nothing Then PinDescription = StripTags(PinDescription)
        ' Check For Pin Title
        If PinTitle = Nothing AndAlso PinDescription <> Nothing Then
            Dim DescriptionWords() As String = Split(PinDescription, " ")
            For i As Integer = 0 To 3
                If DescriptionWords.Count() > i Then
                    If i = 0 Then
                        PinTitle = DescriptionWords(i)
                    Else
                        PinTitle = PinTitle & " " & DescriptionWords(i)
                    End If
                End If
            Next
        End If
        ' Confirm Information
        If PinImageLink = Nothing OrElse PinDescription = Nothing Then
            Log_Show_Error(PinID & vbCrLf & vbCrLf & Content, "GetPinInformation_1", ActiveDatabase)
            Throw New Exception("Failed To Scrape Required Variables - Unknown Issue")
        End If
    End Sub
    ' Pinterest - Scrape User's Boards
    Friend Sub ScrapeUserBoards(ByRef a_Process As CurlFunctions, ByVal UserID As String, ByRef AvailableBoards As List(Of String()))
        ' Set Info
        Dim MainDomain As String = "pinterest.com"
        ' Variables
        Dim URL As String
        Dim Content As String
        ' Load Homepage
        URL = "https://www." & MainDomain & "/"
        Try
            Content = a_Process.Return_Content(URL)
        Catch ex As Exception
            Throw New Exception("Failed To Load Pinterest Homepage - " & ex.Message)
        End Try
        ' Get Token
        Dim Token As String = GetSessionToken(a_Process)
        ' Scrape App Version
        Dim AppVersion As String = GetBetween(Content, """app_version"": """, """")
        If AppVersion = Nothing Then AppVersion = GetBetween(Content, """app_version"":""", """")
        ' Confirm Token Exists
        If Token = Nothing Then
            Throw New Exception("Failed To Scrape Session Token")
        End If
        ' Load Pin Resources URL
        ' URL = "https://www.pinterest.com/_ngjs/resource/BoardsResource/get/?source_url=%2F" & UserID & "%2F&data=%7B%22options%22%3A%7B%22isPrefetch%22%3Afalse%2C%22privacy_filter%22%3A%22all%22%2C%22sort%22%3A%22last_pinned_to%22%2C%22field_set_key%22%3A%22profile_grid_item%22%2C%22username%22%3A%22" & UserID & "%22%2C%22page_size%22%3A25%2C%22group_by%22%3A%22mix_public_private%22%2C%22include_archived%22%3Atrue%2C%22redux_normalize_feed%22%3Atrue%7D%2C%22context%22%3A%7B%7D%7D&_=" & GetMilliSeconds()
        ' URL = "https://www.pinterest.com/_ngjs/resource/BoardsResource/get/?source_url=%2F" & UserID & "%2Fboards%2F&data=%7B%22options%22%3A%7B%22isPrefetch%22%3Afalse%2C%22privacy_filter%22%3A%22all%22%2C%22sort%22%3A%22last_pinned_to%22%2C%22field_set_key%22%3A%22profile_grid_item%22%2C%22username%22%3A%22" & UserID & "%22%2C%22page_size%22%3A25%2C%22group_by%22%3A%22visibility%22%2C%22include_archived%22%3Atrue%2C%22redux_normalize_feed%22%3Atrue%7D%2C%22context%22%3A%7B%7D%7D&_=" & GetMilliSeconds()
        URL = "https://www.pinterest.com/_ngjs/resource/BoardPickerBoardsResource/get/?source_url=%2Fpin-builder%2F&data=%7B%22options%22%3A%7B%22isPrefetch%22%3Afalse%2C%22field_set_key%22%3A%22board_picker%22%7D%2C%22context%22%3A%7B%7D%7D&_=" & GetMilliSeconds()
        Try
            Content = a_Process.XMLHttpRequest_Get_Request(URL, Token, "https://www." & MainDomain & "/", "application/json, text/javascript, */*; q=0.01", True, AppVersion, "active")
        Catch ex As Exception
            Throw New Exception("Failed To Load Pin Resources URL: " & ex.Message)
        End Try
        ' New Method
        Dim ScrapeArea As String = Content
        ' Loop Content
        Dim AlreadyFound As New List(Of String)
        While ScrapeArea.Contains("section_count")
            '  Dim TmpSizeArea As String = GetBetween(ScrapeArea, "is_collaborative", """},""id""")
            Dim TmpSizeArea As String = GetBetween(ScrapeArea, "section_count", "section_count")
            If TmpSizeArea = Nothing Then TmpSizeArea = GetBetween(ScrapeArea, "section_count", "boards_shortlist")
            Dim BoardName As String = GetBetween(TmpSizeArea, """name"":""", """")
            Dim BoardID As String = GetBetween(TmpSizeArea, "board"",""id"":""", """")
            ' Set New Content
            If Split(ScrapeArea, TmpSizeArea).Count() < 2 Then
                ScrapeArea = ""
            Else
                ScrapeArea = Split(ScrapeArea, TmpSizeArea)(1)
            End If
            ' Check If Already Found
            If BoardID = Nothing OrElse AlreadyFound.Contains(BoardID) Then Continue While
            AlreadyFound.Add(BoardID)
            ' Format Data
            Dim Bad_Chars() As String = {"\u2764", "\u2606", "\u2601", "\u2605", "\u2665", "\u0f66", "\u0f7a", "\u0f58", "\u0f0b", "\u0f45", "\u0f53", "\u263c"}
            ' Replace Bad Chars
            For Each TmpChar In Bad_Chars
                BoardName = BoardName.Replace(TmpChar, "")
            Next
            ' Check For Spaces
            If BoardName.StartsWith(" ") Then BoardName = BoardName.Substring(1, BoardName.Length() - 1)
            If BoardName.EndsWith(" ") Then BoardName = BoardName.Substring(0, BoardName.Length() - 2)
            ' Add To Hashtable
            Try
                If BoardName <> Nothing AndAlso BoardID <> Nothing Then AvailableBoards.Add(New String() {BoardName.ToLower().Trim(), BoardID})
            Catch ex As Exception
                ' Ignore Duplicates
            End Try
        End While
    End Sub
    ' Pinterest - Switch To AnalyticsRefresh Fashboard
    Friend Sub SwitchToNewAnalyticsRefresh(ByRef a_Process As CurlFunctions)
        ' Variables
        Dim URL As String = ""
        Dim Data As String = ""
        Dim Content As String = ""
        ' Replace Cookies
        Dim FoundCookie As Boolean = False
        For Each cook As Cookie In a_Process.cookies.GetCookies(New Uri("https://pinterest.com/"))
            If cook.Name.Contains("analyticsOverviewIsDefault") Then
                cook.Value = "AnalyticsRefresh"
                FoundCookie = True
            End If
        Next
        ' Add Cookie If Required
        If FoundCookie = False Then
            Dim TmpCook As New Cookie()
            TmpCook.Domain = "pinterest.com"
            TmpCook.Name = "AnalyticsRefresh"
            TmpCook.Value = "AnalyticsRefresh"
            a_Process.cookies.Add(TmpCook)
        End If
    End Sub
    ' Pinterest - Lookup Stats
    Friend Function LookupStats(ByRef a_Process As CurlFunctions, ByVal Web_UserID As String) As String
        ' Variables
        Dim URL As String = ""
        Dim Content As String = ""
        Dim DateNow As DateTime = DateTime.Now()
        Dim EndDate As String = DateNow.ToString("yyyy-MM-dd")
        Dim StartDate As String = DateNow.AddMonths(-1).ToString("yyyy-MM-dd")
        ' Switch To New Dashboard
        SwitchToNewAnalyticsRefresh(a_Process)
        ' Load Stats URL
        URL = "https://analytics.pinterest.com/_ngjs/resource/AnalyticsTimeSeriesResource/get/?source_url=%2Foverview%2F&data=%7B%22options%22%3A%7B%22isPrefetch%22%3Afalse%2C%22app_types%22%3A%22all%22%2C%22paid%22%3A2%2C%22in_profile%22%3A2%2C%22from_owned_content%22%3A2%2C%22end_date%22%3A%22" & EndDate & "%22%2C%22metric_types%22%3A%22IMPRESSION%22%2C%22owned_content_list%22%3A%22%22%2C%22split_field%22%3A%22NO_SPLIT%22%2C%22start_date%22%3A%22" & StartDate & "%22%2C%22user_id%22%3A%22" & Web_UserID & "%22%2C%22pin_format%22%3Anull%7D%2C%22context%22%3A%7B%7D%7D&_=" & GetMilliSeconds()
        Content = a_Process.XMLHttpRequest_Get_Request(URL, GetSessionToken(a_Process), "", "application/json, text/javascript, */*, q=0.01", False, "", "active")
        ' Scrape Count
        Dim TotalImpressions As String = GetBetween(Content, """summary_metrics"":{""IMPRESSION"":", "}")
        ' Confirm Available
        If TotalImpressions = Nothing Then
            Log_Show_Error(Content, "LookupStats_1", ActiveDatabase)
            Throw New Exception("Failed To Scrape Total Impressions Count - Unknown Issue")
        End If
        ' Return Count
        Return TotalImpressions
    End Function
    ' Pinterest - Clear Notifications
    Friend Sub ClearNotifications(ByRef a_Process As CurlFunctions)
        ' Variables
        Dim URL As String = ""
        Dim Data As String = ""
        Dim Content As String = ""
        Dim Referrer As String = "https://www.pinterest.com/"
        ' Load Homepage
        Content = a_Process.Return_Content("https://www.pinterest.com/")
        ' Scrape App Version
        Dim AppVersion As String = GetBetween(Content, """app_version"": """, """")
        If AppVersion = Nothing Then AppVersion = GetBetween(Content, """app_version"":""", """")
        ' Load Stats URL
        URL = "https://www.pinterest.com/_ngjs/resource/NewsHubBadgeResource/delete/"
        Data = "source_url=%2F&data=%7B%22options%22%3A%7B%7D%2C%22context%22%3A%7B%7D%7D"
        Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
    End Sub
    ' Pinterest - Follow User
    Friend Sub FollowUser(ByRef a_Process As CurlFunctions, ByVal Username As String)
        ' Set Info
        Dim MainDomain As String = "pinterest.com"
        ' Variables
        Dim URL As String
        Dim Data As String
        Dim Referrer As String
        Dim Content As String
        ' Format Username
        Username = Username.Replace("_", "").Replace("-", "")
        ' Load Profile Page
        URL = "https://www." & MainDomain & "/" & URLEncode(Username) & "/"
        Referrer = "https://www." & MainDomain & "/"
        Try
            Content = a_Process.Return_Content(URL, Referrer)
        Catch ex As Exception
            Throw New Exception("Failed To Load User Profile: " & ex.Message)
        End Try
        ' Get Token
        Dim Token As String = GetSessionToken(a_Process)
        ' Scrape App Version
        Dim AppVersion As String = GetBetween(Content, """app_version"": """, """")
        If AppVersion = Nothing Then AppVersion = GetBetween(Content, """app_version"":""", """")
        ' Confirm Token Exists
        If Token = Nothing Then
            Throw New Exception("Failed To Scrape Session Token")
        End If
        ' Check Not Already Following
        If Content.Contains("class=""buttonText"">Unfollow") Then Throw New Exception("Already Following User: " & Username)
        ' Load Get User Data URL
        ' URL = "https://www." & MainDomain & "/resource/UserResource/get/?data={%22options%22%3A{%22username%22%3A%22" & Username & "%22}%2C%22module%22%3A{%22name%22%3A%22UserProfilePage%22%2C%22options%22%3A{%22username%22%3A%22" & Username & "%22%2C%22tab%22%3A%22boards%22}%2C%22append%22%3Afalse%2C%22errorStrategy%22%3A0}%2C%22context%22%3A{%22app_version%22%3A%22" & AppVersion & "%22}}&_=" & GetMilliSeconds()
        URL = "https://www." & MainDomain & "/_ngjs/resource/UserResource/get/?source_url=%2F" & Username & "%2F&data=%7B%22options%22%3A%7B%22isPrefetch%22%3Afalse%2C%22username%22%3A%22" & Username & "%22%2C%22field_set_key%22%3A%22profile%22%7D%2C%22context%22%3A%7B%7D%7D&_=" & GetMilliSeconds()
        Try
            Content = a_Process.XMLHttpRequest_Get_Request(URL, Token, "https://www." & MainDomain & "/" & URLEncode(Username))
        Catch ex As Exception
            If ex.Message.Contains("404") Then
                Throw New Exception("User Not Found")
            Else
                Throw New Exception("Failed To Load Noop URL: " & ex.Message)
            End If
        End Try
        ' Scrape User ID
        Dim UserID As String = GetBetween(Content, """user_id"": """, """")
        If UserID = Nothing Then UserID = GetBetween(Content, """id"":""", """")
        If UserID = Nothing Then UserID = GetBetween(Content, ".jpg"",""id"":""", """")
        ' Confirm Available
        If UserID = Nothing Then
            Log_Show_Error(Content, "FollowUser_1", ActiveDatabase)
            Throw New Exception("Failed To Scrape UserID For Username: " & Username)
        End If
        ' Follow User
        URL = "https://www." & MainDomain & "/resource/UserFollowResource/create/"
        ' Data = "source_url=%2F" & Username & "%2F&data=%7B%22options%22%3A%7B%22user_id%22%3A%22" & UserID & "%22%7D%2C%22context%22%3A%7B%7D%7D&module_path=App()%3EUserProfilePage(resource%3DUserResource(username%3D" & Username & "%2C+invite_code%3Dnull))%3EUserProfileHeader(resource%3DUserResource(username%3D" & Username & "%2C+invite_code%3Dnull))%3EUserFollowButton(followed%3Dfalse%2C+is_me%3Dfalse%2C+unfollow_text%3DUnfollow%2C+memo%3D%5Bobject+Object%5D%2C+follow_ga_category%3Duser_follow%2C+unfollow_ga_category%3Duser_unfollow%2C+disabled%3Dfalse%2C+color%3Dprimary%2C+text%3DFollow%2C+user_id%3D" & UserID & "%2C+follow_text%3DFollow%2C+follow_class%3Dprimary)"
        Data = "source_url=%2F" & Username & "%2F&data=%7B%22options%22%3A%7B%22user_id%22%3A%22" & UserID & "%22%7D%2C%22context%22%3A%7B%7D%7D"
        Referrer = "https://www." & MainDomain & "/" & Username
        Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, Token, Referrer, True, AppVersion, "active")
        ' Confirm Successful Follow
        If Not Content.Contains("""error"": null") AndAlso Not Content.Contains("""error"":null") AndAlso Not Content.Contains("status"":""success") Then
            Dim ErrMsg As String = DetectPinterestError(Content)
            If Content.Contains("We've detected a bot") Then
                Throw New Exception("BOT DETECTION")
            ElseIf Content.Contains("could not complete that request") Then
                Throw New Exception("could not complete that request")
            ElseIf ErrMsg <> Nothing Then
                If Content.Contains("You have exceeded the maximum rate of users followed") Then Throw New Exception("Exceeded Maximum Follow Rate")
                Throw New Exception("Follow Failed - New Issue: " & ErrMsg)
            Else
                Log_Show_Error(Content, "FollowUser_2", ActiveDatabase)
                Throw New Exception("Follow Failed, Unknown Issue")
            End If
        End If
    End Sub
    ' Pinterest - Create Board
    Friend Sub CreateBoard(ByRef a_Process As CurlFunctions, ByVal NewBoardName As String, ByVal Username As String, Optional ByVal BoardCategory As String = "")
        ' Set Info
        Dim MainDomain As String = "pinterest.com"
        ' Variables
        Dim URL As String
        Dim Data As String
        Dim Referrer As String
        Dim Content As String
        ' Change Board Name
        NewBoardName = NewBoardName.Replace("-", " ").Replace(vbCrLf, "").Trim()
        ' Load Pin Page
        URL = "https://www." & MainDomain & "/" & URLEncode(Username.Replace("_", "").Replace("-", "")) & "/boards/"
        Referrer = "https://www." & MainDomain & "/" & URLEncode(Username.Replace("_", "").Replace("-", "")) & "/"
        Content = a_Process.Return_Content(URL, Referrer)
        ' Get Token
        Dim token As String = GetSessionToken(a_Process)
        Dim AppVersion As String = GetBetween(Content, """app_version"": """, """")
        If AppVersion = Nothing Then AppVersion = GetBetween(Content, """app_version"":""", """")
        ' Confirm Token Exists
        If token = Nothing Then
            Log_Show_Error(Content, "CreateBoard_1", ActiveDatabase)
            Throw New Exception("Failed To Scrape Session Token")
        End If
        ' Select Category
        Dim Categories As New List(Of String())
        Categories.Add({"animals", "Animals"})
        Categories.Add({"architecture", "Architecture"})
        Categories.Add({"art", "Art"})
        Categories.Add({"cars_motorcycles", "Cars Motorcycles"})
        Categories.Add({"celebrities", "Celebrities"})
        Categories.Add({"design", "Design"})
        Categories.Add({"diy_crafts", "DIY Crafts"})
        Categories.Add({"education", "Education"})
        Categories.Add({"film_music_books", "Film Music Books"})
        Categories.Add({"food_drink", "Food Drink"})
        Categories.Add({"gardening", "Gardening"})
        Categories.Add({"geek", "Geek"})
        Categories.Add({"hair_beauty", "Hair Beauty"})
        Categories.Add({"health_fitness", "Health Fitness"})
        Categories.Add({"history", "History"})
        Categories.Add({"holidays_events", "Holidays Events"})
        Categories.Add({"home_decor", "Home Decor"})
        Categories.Add({"humor", "Humor"})
        Categories.Add({"illustrations_posters", "Illustrations Posters"})
        Categories.Add({"kids", "Kids"})
        Categories.Add({"mens_fashion", "Men's"})
        Categories.Add({"outdoors", "Outdoors"})
        Categories.Add({"photography", "Photography"})
        Categories.Add({"products", "Products"})
        Categories.Add({"quotes", "Quotes"})
        Categories.Add({"science_nature", "Science Nature"})
        Categories.Add({"sports", "Sports"})
        Categories.Add({"tattoos", "Tattoos"})
        Categories.Add({"technology", "Technology"})
        Categories.Add({"travel", "Travel Places"})
        Categories.Add({"weddings", "Wedding"})
        Categories.Add({"womens_fashion", "Women's"})
        ' Loop Categories
        Dim MainCategory As String = ""
        If BoardCategory <> "" Then
            For Each Category As String() In Categories
                Dim TmpCats() As String = Category(1).Split(CChar(" "))
                For Each Tmp In TmpCats
                    If BoardCategory.Contains(Tmp) Then
                        MainCategory = Category(0)
                        Exit For
                    End If
                Next
            Next
            If MainCategory = Nothing OrElse MainCategory.Trim() = "" Then MainCategory = Categories(CInt(Rnd() * (Categories.Count() - 1)))(0)
        End If
        ' POST Data
        URL = "https://www." & MainDomain & "/resource/BoardResource/create/"
        Data = "data=%7B%22options%22%3A%7B%22name%22%3A%22" & URLEncode(NewBoardName) & "%22%2C%22category%22%3A%22" & MainCategory & "%22%2C%22description%22%3A%22" & URLEncode(NewBoardName) & "%22%2C%22privacy%22%3A%22public%22%2C%22collab_board_email%22%3Afalse%7D%2C%22context%22%3A%7B%22app_version%22%3A%22" & AppVersion & "%22%7D%7D&source_url=%2F" & Username.Replace("_", "").Replace("-", "") & "%2Fboards%2F&module_path=App()%3EUserProfilePage(resource%3DUserResource(username%3D" & Username.Replace("_", "").Replace("-", "") & "))%3EUserProfileContent(resource%3DUserResource(username%3D" & Username.Replace("_", "").Replace("-", "") & "))%3EUserBoards()%3EGrid(resource%3DProfileBoardsResource(username%3D" & Username.Replace("_", "").Replace("-", "") & "))%3EGridItems(resource%3DProfileBoardsResource(username%3D" & Username.Replace("_", "").Replace("-", "") & "))%3EBoardCreateRep(submodule%3D%5Bobject+Object%5D%2C+ga_category%3Dboard_create%2C+tagName%3Da)%23Modal(module%3DBoardCreate())"
        Referrer = "https://www." & MainDomain & "/" & URLEncode(Username.Replace("_", "").Replace("-", "")) & "/boards/"
        Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, token, Referrer)
        ' Confirm Pin OK
        If Content.Contains("read-only mode") Then Throw New Exception("Account In Read-Only Mode")
        If Content.Contains("This field is required") Then Throw New Exception("Missing Required Field")
        If Content.Contains("maximum number of boards") Then Throw New Exception("Maximum")
        If Content.Contains("hiccups") Then Throw New Exception("Hiccups")
        If Content.Contains("already have a board") Then Throw New Exception("Already")
        If Content.Contains("Your category is invalid") Then Throw New Exception("Board Category Is Invalid (" & MainCategory & ")")
        If Content.Contains("Something went wrong on our end") Then Throw New Exception("Pinterest Issue (No Error Message Provided)")
        If Content.Contains("We've detected a bot") Then Throw New Exception("BOT DETECTION")
        If Content.Contains("email address needs to be confirmed") Then Throw New Exception("Email Address Needs To Be Verified")
        ' Scrape Board ID
        Dim TmpBoardID As String = GetBetween(Content, "pin_count"": ", "error")
        ' Confirm Board ID Exists
        If TmpBoardID = Nothing Then
            If Content.Contains("already have a board with that name") OrElse Content.Contains("already got one") Then
                Throw New Exception("Issue Parsing Board Details")
            ElseIf Content.Contains("Something went wrong on our end") Then
                Throw New Exception("Pinterest Issue (No Error Message Provided)")
            End If
            Log_Show_Error(Data & vbCrLf & vbCrLf & Content, "CreateBoard_2", ActiveDatabase)
            Throw New Exception("Failed To Scrape New Board ID")
        End If
        ' Scrape True ID
        Dim BoardID As String = GetBetween(TmpBoardID, """object_id_str"":""", """")
        If BoardID = Nothing Then BoardID = GetBetween(Content, "},""id"":""", """")
        ' Confirm Board ID Exists
        If BoardID = Nothing Then
            If Content.Contains("already have a board with that name") OrElse Content.Contains("already got one") Then
                Throw New Exception("Issue Parsing Board Details")
            ElseIf Content.Contains("Something went wrong on our end") Then
                Throw New Exception("Pinterest Issue (No Error Message Provided)")
            End If
            Log_Show_Error(Content, "CreateBoard_3", ActiveDatabase)
            Throw New Exception("Failed To Scrape New Board ID")
        End If
        ' Check Pin Count
        Dim Pin_Count As String = GetBetween(Content, """pin_count"": ", ",")
        ' Successfully Created Board
        If Pin_Count <> Nothing AndAlso Pin_Count <> "0" Then AddMessage("Successfully Created New Board: " & NewBoardName)
    End Sub
#End Region

#Region "Main Processes"
    ' Variables
    Friend AccountsToRefreshForBoardInformation As New List(Of String)
    ' Board Manager Timer Thread Manager
    Private Sub NextPostUpdater_Tick(sender As Object, e As EventArgs) Handles NextPostUpdater.Tick
        Try
            ' Disable Timer To Prevent Multiple Runs
            NextPostUpdater.Enabled = False
            ' Loop Tasks
            SyncLock TaskUpdaterLock
                For Each row As DataRow In Full_Campaign_Data.Rows
                    Dim Campaign_Account As String = CStr(row("Account"))
                    Dim Campaign_BoardName As String = CStr(row("Board Name"))
                    Dim Campaign_Status As String = CStr(row("Status"))
                    Dim Campaign_Keywords As String = CStr(row("Keywords"))
                    Dim Campaign_Website As String = CStr(row("Website"))
                    Dim Campaign_Delay As String = CStr(row("Delay"))
                    Dim Campaign_Inuse As String = CStr(row("Inuse"))
                    Dim Campaign_PinsCreated As String = CStr(row("Pins Created"))
                    Dim NextPostTime As String = CStr(row("Next Post Time"))
                    ' Check If Inuse
                    If Campaign_Inuse = "1" OrElse Campaign_Status = "Task Disabled" Then Continue For
                    ' Check If 0 Seconds & Start Thread
                    ' Dim DateDiff As Integer = (CDate(NextUpload) - DateTime.Now).TotalSeconds
                    If CDate(NextPostTime) <= DateTime.Now Then
                        ' Set Campaign Status
                        row("Status") = "Starting Thread"
                        ' Set Campaign Inuse
                        row("Inuse") = "1"
                        ' Start Thread Object
                        Dim Threaded_Class As New ThreadManager(Me)
                        ' Set Thread Variables
                        Threaded_Class.Campaign_Account = Campaign_Account
                        Threaded_Class.Campaign_BoardName = Campaign_BoardName
                        Threaded_Class.Campaign_Status = Campaign_Status
                        Threaded_Class.Campaign_Keywords = Campaign_Keywords
                        Threaded_Class.Campaign_Website = Campaign_Website
                        Threaded_Class.Campaign_Delay = Campaign_Delay
                        Threaded_Class.Campaign_Inuse = Campaign_Inuse
                        Threaded_Class.Campaign_PinsCreated = Campaign_PinsCreated
                        Threaded_Class.Campaign_Delay = Campaign_Delay
                        Threaded_Class.NextPostTime = NextPostTime
                        Threaded_Class.ActiveDatabase = ActiveDatabase
                        ' Start Thread
                        Dim objNewThread As New Thread(AddressOf Threaded_Class.RunProcess)
                        objNewThread.IsBackground = True
                        objNewThread.Start()
                        ' Exit Loop To Prevent Thread Creations Lagging Program
                        Exit For
                    End If
                Next
            End SyncLock
            ' Re-Enable Timer
            NextPostUpdater.Enabled = True
        Catch ex As Exception
            Log_Show_Error(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, Me.Name & "_" & System.Reflection.MethodInfo.GetCurrentMethod().Name, ActiveDatabase)
            MessageBox.Show("A Fatal Error Occured With The Timer Process - " & ex.Message)
        End Try
    End Sub
    ' Lookup Board Information For Main Form 
    Friend Sub RefreshBoardInformation()
        Try
            For Each TmpAccount As String In AccountsToRefreshForBoardInformation
                Dim AccountUsername As String = TmpAccount
                AddMessage("Refreshing Board Information For Account " & AccountUsername)
                ' Set Account Variables
                Dim Web_Username As String = AccountUsername
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
                        m_MainWindow.AddMessage("Failed To Find Account In System " & Web_Username)
                        Continue For
                    End If
                    ' Check Account Active
                    If AccountActiveStatus = "0" Then
                        m_MainWindow.AddMessage("Account Is Disabled " & Web_Username)
                        Continue For
                    ElseIf AccountActiveStatus = "2" Then
                        m_MainWindow.AddMessage("Account Is Locked " & Web_Username)
                        Continue For
                    End If
                End SyncLock
                ' Start New HTTP Session
                Dim a_Process As New CurlFunctions(Web_Username, ActiveDatabase, Web_Proxy, Web_Port, Web_ProxyUsername, Web_ProxyPassword)
                ' Login To Account
                Try
                    m_MainWindow.Login(a_Process, Web_FriendlyUsername, Web_Password, Web_Username, Web_UserID)
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Login To Account " & Web_Username & " - " & ex.Message)
                    Continue For
                End Try
                ' Scrape Users Boards
                Dim AvailableBoards As New List(Of String())
                Try
                    m_MainWindow.ScrapeUserBoards(a_Process, Web_FriendlyUsername, AvailableBoards)
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Scrape User's Boards: " & ex.Message)
                    Continue For
                End Try
                ' Loop Boards
                For Each TmpBoard() As String In AvailableBoards
                    ' Set Variables
                    Dim TmpBoardName As String = TmpBoard(0)
                    Dim TmpBoardId As String = TmpBoard(1)
                    ' Check Not Already Exists
                    Dim userToRemove = BoardInformation.Cast(Of Object()).Where(Function(j) j(2).Equals(TmpBoardId)).Any()
                    If userToRemove = True Then Continue For
                    ' Add To Main List
                    BoardInformation.Add(New String() {AccountUsername, TmpBoardName, TmpBoardId})
                Next
                ' Refresh Board Info For Account
                RefreshBoards(AccountUsername)
                ' Save Data
                SaveBoardInformation()
            Next
            ' Mark As Process Completed
            RefreshBoards("ProcessComplete")
        Catch ex As Exception
            Log_Show_Error(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, Me.Name & "_" & System.Reflection.MethodInfo.GetCurrentMethod().Name, ActiveDatabase)
            RefreshBoards("Process Completed")
            MessageBox.Show("An Issue Occured: " & ex.Message)
        End Try
    End Sub
    ' Follower Timer Thread Manager
    Private Sub FollowerTimer_Tick(sender As Object, e As EventArgs) Handles FollowerTimer.Tick
        Try
            ' Disable Timer To Prevent Multiple Runs
            FollowerTimer.Enabled = False
            ' Loop Tasks
            SyncLock FollowerTaskUpdaterLock
                For Each row As DataRow In Full_Following_Task_Data.Rows
                    Dim Campaign_Account As String = CStr(row("Account"))
                    Dim Campaign_Keywords As String = CStr(row("Keywords"))
                    Dim Campaign_MinFollow As String = CStr(row("MinFollow"))
                    Dim Campaign_MaxFollow As String = CStr(row("MaxFollow"))
                    Dim Campaign_MinDelay As String = CStr(row("MinDelay"))
                    Dim Campaign_MaxDelay As String = CStr(row("MaxDelay"))
                    Dim Campaign_Inuse As String = CStr(row("Inuse"))
                    Dim Campaign_FollowsSent As String = CStr(row("Follows Sent"))
                    Dim NextTaskTime As String = CStr(row("Next Task Time"))
                    Dim Campaign_Status As String = CStr(row("Status"))
                    Dim Campaign_DelayTime As String = CStr(row("DelayTime"))
                    ' Check If Inuse
                    If Campaign_Inuse = "1" OrElse Campaign_Status = "Task Disabled" Then Continue For
                    ' Check If 0 Seconds & Start Thread
                    ' Dim DateDiff As Integer = (CDate(NextUpload) - DateTime.Now).TotalSeconds
                    If CDate(NextTaskTime) <= DateTime.Now Then
                        ' Set Campaign Status
                        row("Status") = "Starting Thread"
                        ' Set Campaign Inuse
                        row("Inuse") = "1"
                        ' Start Thread Object
                        Dim Threaded_Class As New FollowerThreadManager(Me)
                        ' Set Thread Variables
                        Threaded_Class.Campaign_Account = Campaign_Account
                        Threaded_Class.Campaign_Keywords = Campaign_Keywords
                        Threaded_Class.Campaign_MinFollow = Campaign_MinFollow
                        Threaded_Class.Campaign_MaxFollow = Campaign_MaxFollow
                        Threaded_Class.Campaign_MinDelay = Campaign_MinDelay
                        Threaded_Class.Campaign_MaxDelay = Campaign_MaxDelay
                        Threaded_Class.Campaign_Inuse = Campaign_Inuse
                        Threaded_Class.Campaign_FollowsSent = Campaign_FollowsSent
                        Threaded_Class.NextTaskTime = NextTaskTime
                        Threaded_Class.Campaign_Status = Campaign_Status
                        Threaded_Class.Campaign_DelayTime = Campaign_DelayTime
                        Threaded_Class.ActiveDatabase = ActiveDatabase
                        ' Start Thread
                        Dim objNewThread As New Thread(AddressOf Threaded_Class.RunProcess)
                        objNewThread.IsBackground = True
                        objNewThread.Start()
                        ' Exit Loop To Prevent Thread Creations Lagging Program
                        Exit For
                    End If
                Next
            End SyncLock
            ' Re-Enable Timer
            FollowerTimer.Enabled = True
        Catch ex As Exception
            Log_Show_Error(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, Me.Name & "_" & System.Reflection.MethodInfo.GetCurrentMethod().Name, ActiveDatabase)
            MessageBox.Show("A Fatal Error Occured With The Timer Process - " & ex.Message)
        End Try
    End Sub
    ' Email Verifier Timer
    Private Sub EmailVerifierTimer_Tick(sender As Object, e As EventArgs) Handles EmailVerifierTimer.Tick
        Try
            ' Disable Timer To Prevent Multiple Runs
            EmailVerifierTimer.Enabled = False
            ' Loop Accounts
            For Each TmpAccount() As String In AccountCreator_CreatedAccounts
                ' Check Ok To Process Account
                If TmpAccount(8) <> "0" Then Continue For
                ' Check Time Has Passed To Allow For Verification
                Dim Starttime As DateTime = CDate(TmpAccount(9))
                Dim EndTime As DateTime = DateTime.Now
                Dim TimeDifference As TimeSpan = EndTime - Starttime
                If TimeDifference.TotalSeconds >= AccountCreator_EmailVerifyDelay Then
                    ' Set As Inuse
                    TmpAccount(8) = "3"
                    ' Set Account Variables
                    Dim EmailAddress As String = TmpAccount(0)
                    Dim Password As String = TmpAccount(1)
                    Dim Proxy As String = TmpAccount(2)
                    Dim Port As Integer = TmpAccount(3)
                    Dim ProxyUser As String = TmpAccount(4)
                    Dim ProxyPass As String = TmpAccount(5)
                    Dim Username As String = TmpAccount(6)
                    Dim EmailPassword As String = TmpAccount(7)
                    Dim CurrentStatus As String = TmpAccount(8) ' 0 = Not Verified Yet | 1 = Verified | 2 = Email Verification Failed | 3 = InUse
                    Dim AccountCreatedTime As String = TmpAccount(9)
                    ' Start New Class Instance
                    Dim TmpEmailVerifier As New EmailVerifier(m_MainWindow, ActiveDatabase)
                    ' Set Variables
                    With TmpEmailVerifier
                        .EmailAddress = EmailAddress
                        .Password = Password
                        .Proxy = Proxy
                        .Port = Port
                        .ProxyUser = ProxyUser
                        .ProxyPass = ProxyPass
                        .Username = Username
                        .EmailPassword = EmailPassword
                        .AccountCreatedTime = AccountCreatedTime
                    End With
                    ' Start Process In Background Thread
                    Dim objNewThread As New Thread(AddressOf TmpEmailVerifier.RunProcess)
                    objNewThread.IsBackground = True
                    objNewThread.Start()
                    ' Exit Loop To Allow Next Timer To Pick Up Next Account
                    Exit For
                End If
            Next
            ' Re-Enable Timer
            EmailVerifierTimer.Enabled = True
        Catch ex As Exception
            Log_Show_Error(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, Me.Name & "_" & System.Reflection.MethodInfo.GetCurrentMethod().Name, ActiveDatabase)
            MessageBox.Show("A Fatal Error Occured With The Timer Process - " & ex.Message)
        End Try
    End Sub
    ' Sleep
    Friend Sub Sleep(ByVal SleepDelay As Integer, ByVal CurPage As String)
        For i = 0 To SleepDelay - 1
            If CurPage = "Follower" AndAlso Follower_StopProcess = True Then Exit For
            If CurPage = "Board Manager" AndAlso BoardManager_StopProcess = True Then Exit For
            If CurPage = "AccountCreator" AndAlso AccountCreator_StopProcess = True Then Exit For
            Thread.Sleep(1000)
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
#End Region

End Class
