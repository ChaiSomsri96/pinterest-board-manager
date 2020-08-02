Imports System.IO
Imports System.Text
Imports System.Threading

Friend Class AccountManager

#Region "Account Structure"
    ' 0 => Username
    ' 1 => Password
    ' 2 => Proxy
    ' 3 => Port
    ' 4 => Proxy Username
    ' 5 => Proxy Password
    ' 6 => Active

#End Region

#Region "Form Variables"
    ' Variables
    Friend m_MainForm As Form1
#End Region

#Region "Data Management Subs"
    ' Refresh Accounts Sub
    Friend Sub RefreshAccounts()
        Try
            CheckedListBox1.Items.Clear()
            For Each Account In m_MainForm.Accounts
                CheckedListBox1.Items.Add(Account(0))
            Next
            ' Refresh Accounts In Main Form
            m_MainForm.RefreshAccounts()
        Catch ex As Exception
            MessageBox.Show("A Fatal Error Occured Loading Accounts Into Form - " & ex.Message)
        End Try
    End Sub
    ' Save Accounts Sub - Saves Accounts To Session Folder
    Friend Sub SaveAccounts()
        Try
            Dim strFile As String = My.Application.Info.DirectoryPath & "\" & m_MainForm.ActiveDatabase & "\session\" & "Accounts.txt"
            Dim objStreamWriter As StreamWriter
            objStreamWriter = New StreamWriter(strFile, False, Encoding.Unicode)
            ' Loop Accounts
            For Each Acc In m_MainForm.Accounts
                objStreamWriter.WriteLine("{0}:{1}:{2}:{3}:{4}:{5}:{6}", Acc(0), Acc(1), Acc(2), Acc(3), Acc(4), Acc(5), Acc(8))
            Next
            ' Close File
            objStreamWriter.Close()
        Catch ex As Exception
            MessageBox.Show("Note: Accounts Failed To Save To File (\" & m_MainForm.ActiveDatabase & "\session\Accounts.txt" & " - " & ex.Message)
        End Try
    End Sub

#End Region

#Region "Form Load / Close Events"
    ' Form Load
    Private Sub Accounts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RefreshAccounts()
        Me.ActiveControl = CheckedListBox1
    End Sub
    ' Prevent Form Close
    Private Sub MyForm_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            Hide()
        End If
    End Sub
#End Region

#Region "Buttons"
    ' Add Account Button
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            ' Check For Dupe
            For Each Account In m_MainForm.Accounts
                If TextBox1.Text = Account(0) Then
                    MessageBox.Show("You Have Already Added This Account")
                    Exit Sub
                End If
            Next
            ' Set Account Variables
            Dim i_Username As String = TextBox1.Text
            Dim i_Password As String = TextBox2.Text
            Dim i_Proxy As String = "0"
            If TextBox3.Text <> "Proxy (Optional)" Then i_Proxy = TextBox3.Text
            Dim i_Port As String = "0"
            If TextBox4.Text <> "Port (Optional)" Then i_Port = TextBox4.Text
            Dim i_ProxyUsername As String = "0"
            Dim i_ProxyPassword As String = "0"
            If TextBox5.Text <> "Proxy Username (Optional)" Then i_ProxyUsername = TextBox5.Text
            If TextBox6.Text <> "Proxy Password (Optional)" Then i_ProxyPassword = TextBox6.Text
            If Integer.TryParse(i_Port, 0) = False Then
                MessageBox.Show("Please Enter A Valid Port")
                Exit Sub
            End If
            m_MainForm.Accounts.Add({i_Username, i_Password, i_Proxy, i_Port, i_ProxyUsername, i_ProxyPassword, "1", "0", "0"})
            SaveAccounts()
            RefreshAccounts()
        Catch ex As Exception
            MessageBox.Show("A Fatal Error Occured Adding An Account - " & ex.Message)
        End Try
    End Sub
    ' Remove Selected Accounts Button
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            For Each Item As String In CheckedListBox1.CheckedItems
                For i As Integer = 0 To m_MainForm.Accounts.Count() - 1
                    If m_MainForm.Accounts(i)(0) = Item Then
                        m_MainForm.Accounts.RemoveAt(i)
                        Exit For
                    End If
                Next
            Next
            SaveAccounts()
            RefreshAccounts()
        Catch ex As Exception
            MessageBox.Show("A Fatal Error Occured Removing The Account - " & ex.Message)
        End Try
    End Sub
    ' Remove All Accounts Button
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            CheckedListBox1.Items.Clear()
            m_MainForm.Accounts.Clear()
            SaveAccounts()
        Catch ex As Exception
            MessageBox.Show("A Fatal Error Occured Removing Accounts - " & ex.Message)
        End Try
    End Sub
    ' Load Accounts Button
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            ' Display Account Type Form Before Import
            MessageBox.Show("Load Accounts From A Text File In The Following Format (1 Per Line)" & vbCrLf & "EmailAddress:Password:Proxy:Port:ProxyUsername:ProxyPassword" & vbCrLf & "Example:" & vbCrLf &
                            "joebloggs@email.com:qwerty:127.0.0.1:8080:0:0")
            OpenFileDialog1.Filter = "*txt Text Files|*.txt"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                Dim objReader As New System.IO.StreamReader(OpenFileDialog1.FileName)
                Dim Line As String
                Dim AccountsLoaded As Integer = 0
                Dim SkippedInt As Integer = 0
                Do While objReader.Peek() <> -1
                    Line = objReader.ReadLine()
                    If Line = Nothing OrElse Line = "" Then Continue Do
                    Line.Replace("	", "")
                    Dim Account As String() = Split(Line, ":")
                    If Account.Count < 6 Then Continue Do
                    Dim Username As String = Account(0)
                    Dim Password As String = Account(1)
                    Dim Proxy As String = Account(2)
                    Dim Port As String = Account(3)
                    Dim i_ProxyUsername As String = Account(4)
                    Dim i_ProxyPassword As String = Account(5)
                    ' Check For Valid Port
                    If Integer.TryParse(Port, 0) = False Then Continue Do
                    ' Check To See If Already Exists
                    m_MainForm.Accounts.Add({Username, Password, Proxy, Port, i_ProxyUsername, i_ProxyPassword, "1", "0", "0"})
                    AccountsLoaded += 1
                Loop
                objReader.Close()
                ' Reload CheckedBoxList
                SaveAccounts()
                RefreshAccounts()
            End If
        Catch ex As Exception
            MessageBox.Show("A Fatal Error Occured Importing Accounts - " & ex.Message)
        End Try
    End Sub
    ' Export Accounts Button
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            ' Open CSV File
            Dim strFile As String = My.Application.Info.DirectoryPath & "\" & m_MainForm.ActiveDatabase & "\exports\" & "Accounts.txt"
            Dim objStreamWriter As StreamWriter
            objStreamWriter = New StreamWriter(strFile, False, Encoding.Unicode)
            ' Write Column Titles
            objStreamWriter.WriteLine("{0}:{1}:{2}:{3}:{4}:{5}", "Username", "Password", "Proxy", "Port", "Proxy Username", "Proxy Password")
            ' Loop Accounts
            For Each Account In m_MainForm.Accounts
                objStreamWriter.WriteLine("{0}:{1}:{2}:{3}:{4}:{5}", Account(0), Account(1), Account(2), Account(3), Account(4), Account(5))
            Next
            ' Close File
            objStreamWriter.Close()
            ' Read File
            System.Diagnostics.Process.Start(strFile)
        Catch ex As Exception
            MessageBox.Show("A Fatal Error Occured Exporting Accounts - " & ex.Message)
        End Try
    End Sub
    ' Export Disabled Accounts Buttons
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            ' Open CSV File
            Dim strFile As String = My.Application.Info.DirectoryPath & "\" & m_MainForm.ActiveDatabase & "\exports\" & "DisabledAccounts.txt"
            Dim objStreamWriter As StreamWriter
            objStreamWriter = New StreamWriter(strFile, False, Encoding.Unicode)
            ' Write Column Titles
            objStreamWriter.WriteLine("{0}:{1}:{2}:{3}:{4}:{5}", "Username", "Password", "Proxy", "Port", "Proxy Username", "Proxy Password")
            ' Loop Accounts
            For Each Account In m_MainForm.Accounts
                If Account(6) <> "0" Then Continue For
                objStreamWriter.WriteLine("{0}:{1}:{2}:{3}:{4}:{5}", Account(0), Account(1), Account(2), Account(3), Account(4), Account(5))
            Next
            ' Close File
            objStreamWriter.Close()
            ' Read File
            System.Diagnostics.Process.Start(strFile)
        Catch ex As Exception
            MessageBox.Show("A Fatal Error Occured Exporting Accounts - " & ex.Message)
        End Try
    End Sub
    ' Refresh Board Information
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            ' Remove Old Accounts
            m_MainForm.AccountsToRefreshForBoardInformation.Clear()
            ' Add All Accounts
            For Each TmpAccount() As String In m_MainForm.Accounts
                Dim AccountUsername As String = TmpAccount(0)
                m_MainForm.AccountsToRefreshForBoardInformation.Add(AccountUsername)
            Next
            ' Start Threaded Sub
            Dim objNewThread As New Thread(AddressOf m_MainForm.RefreshBoardInformation)
            objNewThread.IsBackground = True
            objNewThread.Start()
        Catch ex As Exception
            MessageBox.Show("An Issue Occured - " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Got / Lost Focus Events"
    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.GotFocus
        If TextBox1.Text = "Login Email Address" Then TextBox1.Text = ""
    End Sub
    Private Sub TextBox7_TextChanged_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.LostFocus
        If TextBox1.Text = "" Then TextBox1.Text = "Login Email Address"
    End Sub
    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.GotFocus
        If TextBox2.Text = "Login Password" Then TextBox2.Text = ""
    End Sub
    Private Sub TextBox2_TextChanged_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.LostFocus
        If TextBox2.Text = "" Then TextBox2.Text = "Login Password"
    End Sub
    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.GotFocus
        If TextBox3.Text = "Proxy (Optional)" Then TextBox3.Text = ""
    End Sub
    Private Sub TextBox3_TextChanged_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.LostFocus
        If TextBox3.Text = "" Then TextBox3.Text = "Proxy (Optional)"
    End Sub
    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.GotFocus
        If TextBox4.Text = "Port (Optional)" Then TextBox4.Text = ""
    End Sub
    Private Sub TextBox4_TextChanged_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.LostFocus
        If TextBox4.Text = "" Then TextBox4.Text = "Port (Optional)"
    End Sub
    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.GotFocus
        If TextBox5.Text = "Proxy Username (Optional)" Then TextBox5.Text = ""
    End Sub
    Private Sub TextBox5_TextChanged_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.LostFocus
        If TextBox5.Text = "" Then TextBox5.Text = "Proxy Username (Optional)"
    End Sub
    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.GotFocus
        If TextBox6.Text = "Proxy Password (Optional)" Then TextBox6.Text = ""
    End Sub
    Private Sub TextBox6_TextChanged_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.LostFocus
        If TextBox6.Text = "" Then TextBox6.Text = "Proxy Password (Optional)"
    End Sub

#End Region

End Class