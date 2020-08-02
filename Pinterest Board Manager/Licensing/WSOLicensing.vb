Option Strict Off

Imports System.Web
Imports System.Net
Imports System.IO
Imports System.Text

Friend NotInheritable Class WSOLicensing
    Private _LicenseKey As String
    Private _UnlimitedCamps As String
    Private _Error As String
    Private _Validated As Boolean
    Private _MachID As String = System.Environment.MachineName
    Private UseBackupServer As Boolean = False
    Private Crypto As crypo = New crypo

    Friend Sub New(ByVal LicenseKey As String, ByVal t_CallServer As Boolean, ByRef check As Class1)
        Try
            ' Set Username & Password
            _LicenseKey = LicenseKey
            ' Check For Bad Applications
            Dim Banned_Processes() As String = {"fiddler", "xcharles", "simpleassemblyexplorer", "reflector", "hxd", "ollydb", "wireshark", "burp", "webscarab", "paros", "networkminer", "ratproxy", "andiparos", "owasp zap", "mitmproxy", "imglory"}
            For Each p As Process In Process.GetProcesses()
                If Banned_Processes.Contains(p.ProcessName.ToLower()) Then
                    MessageBox.Show("Please Close Any Programs Or Loaders That May Be Monitoring Web Activity Before Starting This Application")
                    End
                End If
            Next
            Try
                ' Check For Modified Hosts File
                Dim HostsFile As String = My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.System) & "\drivers\etc\hosts").ToLower()
                ' If HostsFile.Contains("blackhattoolz.com") OrElse HostsFile.Contains("statuslol.com") OrElse HostsFile.Contains("184.172.169.134") Then End
            Catch ex As Exception

            End Try
            ' Set Crypto Variables
            Dim random As New Random
            Dim str As String = "abcdefghijklmnopqrstuvwyxzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
            Dim chArray As Char() = New Char(&H10 - 1) {}
            Dim i As Integer
            For i = 0 To &H10 - 1
                chArray(i) = str.Chars(random.Next(str.Length))
            Next i
            Dim strRnd As String = chArray
            Crypto.LicenseKey = LicenseKey
            Crypto.Key = strRnd
            check.m_key = strRnd
            ' Send Login Request To Server
            If t_CallServer = True Then
                ' Check For Updates
                ' CheckVersion(Form1, True)
                CallServer(check)
            End If
        Catch ex As Exception
            MessageBox.Show("An Error Occured Launching Main Core (3) - " & ex.Message)
            End
        End Try
    End Sub

    Friend Property LicenseKey() As String
        Get
            Return _LicenseKey
        End Get
        Set(ByVal value As String)
            _LicenseKey = value
        End Set
    End Property

    Friend Property ErrorMsg() As String
        Get
            Return _Error
        End Get
        Set(ByVal value As String)
            _Error = value
        End Set
    End Property

    Friend Property Validated() As Boolean
        Get
            Return _Validated
        End Get
        Set(ByVal value As Boolean)
            _Validated = value
        End Set
    End Property

    Private Sub CallServer(ByRef check As Class1)
        Try
            ' Variables
            Dim webresponse As String = ""
            ' Get Local IP
            Dim LocalIP As String = ""
            Dim hostName = System.Net.Dns.GetHostName()
            For Each hostAdr In System.Net.Dns.GetHostEntry(hostName).AddressList()
                If hostAdr.ToString().StartsWith("192.168.") Then
                    LocalIP = hostAdr.ToString()
                    Exit For
                End If
            Next
            ' Set Crypto Variables
            Crypto.LocalIP = LocalIP
            Crypto.MachID = _MachID
            ' Get Response
            Dim request As System.Net.HttpWebRequest
            Dim response As System.Net.HttpWebResponse
            Dim sr As System.IO.StreamReader
            Dim URL As String = ""
            Try
                ' Get Response
                Dim str As String = Crypto.Encrypt(Crypto.LicenseKey & "," & Crypto.MachID & "," & Crypto.LocalIP & "," & Form1.SoftwareName.Replace(" ", "_"), Crypto.Key)
                URL = "http://www.statuslol.com/licensing/custom/crypto_new.php?th=" & Me.Crypto.Key & URLEncode(str)
                request = CType(System.Net.HttpWebRequest.Create(URL), HttpWebRequest)
                request.Proxy = Nothing
                response = CType(request.GetResponse(), HttpWebResponse)
                sr = New System.IO.StreamReader(response.GetResponseStream())
                webresponse = sr.ReadToEnd()
            Catch ex As Exception
                MessageBox.Show("An Error Occured Launching Main Core (2A) - " & ex.Message)
                Process.Start(URL)
                End
            End Try
            ValidateLicense(webresponse, check)
        Catch ex As Exception
            MessageBox.Show("An Error Occured Launching Main Core (2) - " & ex.Message)
            End
        End Try
    End Sub

    Private Sub ValidateLicense(ByVal EncryptedResponse As String, ByRef check As Class1)
        Try
            ' Variables
            Dim First_Encrypted_Response As String = ""
            Dim Second_Encrypted_Response As String = ""
            ' Decrypt Response
            Dim key As String = ""
            Dim tk As String = ""
            Dim Content As String = ""
            Dim NewContent() As String
            ' Set Errors
            Dim Errors(15) As String
            Errors(0) = "License Key Not Recognized"
            Errors(1) = "License Key Already Been Activated On 2 PC's. Contact Admin@BlackHatToolz.com Deactivate A PC"
            Errors(2) = "License Key Already Been Activated On 2 PC's. Contact Admin@BlackHatToolz.com Deactivate A PC"
            Errors(3) = "The Licensing Server Is Currently Unavailable, Please Try Again In A Few Minutes"
            Errors(4) = "Could Not Connect To Licensing Server. Please Make Sure An Internet Connection Is Available"
            Errors(5) = "An Issue Occured With The Licensing Server" & vbCrLf & "The Updater Program Is Now Starting To Check For Available Updates"
            Errors(6) = "The Licensing Server Is Currently Unavailable, Please Try Again In A Few Minutes"
            Errors(7) = "A Licensing Error Occured. Please Run Updater.exe To Check For A Program Update"
            Errors(8) = "Could Not Connect To Licensing Server. Please Make Sure An Internet Connection Is Available"
            Errors(9) = "Could Not Connect To Licensing Server. Please Make Sure An Internet Connection Is Available 9"
            Errors(10) = "Could Not Connect To Licensing Server. Please Make Sure An Internet Connection Is Available 10"
            Errors(11) = "Could Not Connect To Licensing Server. Please Make Sure An Internet Connection Is Available 11"
            Errors(12) = "Could Not Connect To Licensing Server. Please Make Sure An Internet Connection Is Available 12"
            Errors(13) = "Licensing IP Checks Failed - Please Contact Admin - admin@blackhattoolz.com"
            Errors(14) = "The Licensing Server Returned A Bad Response - Please Contact Administrator - admin@blackhattoolz.com " & vbCrLf & "Server Response: " & EncryptedResponse
            ' Confirm Valid Response
            If EncryptedResponse = Nothing OrElse EncryptedResponse = "" Then
                _Validated = False
                _Error = Errors(15)
                Exit Sub
            End If
            ' Split Encrypted Responses
            Try
                First_Encrypted_Response = EncryptedResponse.Split(CChar(":"))(0)
                Second_Encrypted_Response = EncryptedResponse.Split(CChar(":"))(1)
            Catch ex As Exception
                MessageBox.Show("An Error Occured Launching Main Core (4) - " & ex.Message)
                End
            End Try
            Try
                ' Decrypt Response
                key = First_Encrypted_Response.Substring(0, &H10)
            Catch ex As Exception
                MessageBox.Show("An Error Occured Launching Main Core (5A) - " & ex.Message & vbCrLf & First_Encrypted_Response & vbCrLf & vbCrLf & EncryptedResponse)
                End
            End Try
            Try
                ' Dim tk As String = EncryptedResponse.Substring(&H10, &H18)
                tk = First_Encrypted_Response.Remove(0, 16)
            Catch ex As Exception
                MessageBox.Show("An Error Occured Launching Main Core (5B) - " & ex.Message)
                End
            End Try
            Try
                Content = Me.Crypto.Decrypt(tk, key)
                ' Split Response
                NewContent = Split(Content, ":")
            Catch ex As Exception
                MessageBox.Show("An Error Occured Launching Main Core (5C) - " & ex.Message)
                End
            End Try
            ' Confirm Valid Content
            If NewContent.Length <> 6 Then
                _Validated = False
                _Error = Errors(14)
                Exit Sub
            End If
            ' Set Vars
            Dim Err As String = NewContent(1)
            Dim Lower As String = NewContent(0)
            Dim Upper As String = NewContent(2)
            Dim AllowedAccounts As Integer = CInt(NewContent(3))
            Dim MachID As String = NewContent(4)
            Dim LocalIP As String = NewContent(5)
            ' Check For Error
            If Err.Length = 1 Then
                _Error = Errors(CInt(Err))
                _Validated = False
                If CInt(Err) < 2 Then Process.Start("http://www.BlackHatToolz.com/")
                Exit Sub
            End If
            ' Get System IP Address
            Dim req As HttpWebRequest
            Dim res As HttpWebResponse
            Dim Stream As IO.Stream
            Dim sr As StreamReader
            Dim IP As String = ""
            Try
                Dim URL As String
                If UseBackupServer = True Then
                    URL = "http://www.blackhattoolz.com/showip.php"
                Else
                    URL = "http://www.statuslol.com/licensing/showip.php"
                End If
                req = CType(WebRequest.Create(URL), HttpWebRequest)
                req.Proxy = Nothing
                res = CType(req.GetResponse(), HttpWebResponse)
                Stream = res.GetResponseStream()
                sr = New StreamReader(Stream)
                IP = sr.ReadToEnd()
                res.Close()
            Catch ex As Exception
                MessageBox.Show("An Issue Occured Validating Your License Key:" & vbCrLf & ex.Message)
                Exit Sub
            End Try
            ' Split IP
            Dim IP_Parts() As String = Split(IP, ".")
            ' Confirm License Validation OK
            If Err <> IP Then
                _Validated = False
                _Error = Errors(CInt(Err))
                Exit Sub
            End If
            ' Confirm Security
            If Lower <> MD5(IP_Parts(0)) Then
                _Validated = False
                _Error = Errors(9)
                Exit Sub
            End If
            If Upper <> MD5(IP_Parts(2)) Then
                _Validated = False
                _Error = Errors(10)
                Exit Sub
            End If
            If MachID <> MD5(_MachID) Then
                _Validated = False
                _Error = Errors(11)
                Exit Sub
            End If
            ' Check Local IP
            Dim ValidatedIP As String = ""
            Dim LocalIPValidated As Boolean = False
            Dim hostName = System.Net.Dns.GetHostName()
            For Each hostAdr In System.Net.Dns.GetHostEntry(hostName).AddressList()
                If MD5(hostAdr.ToString()) = LocalIP Then
                    ValidatedIP = hostAdr.ToString()
                    LocalIPValidated = True
                    Exit For
                End If
            Next
            ' Decrypt Response
            Dim MainDomain As String = Me.Crypto.Decrypt(Second_Encrypted_Response, Crypto.Key).Replace(Chr(0), "")
            ' License Validated
            SaveSetting(Form1.SoftwareName, "textboxes", "LicenseKey", LicenseKey)
            ' Create New Form
            Dim MainForm As New Form1
            MainForm.Unlocked = True
            MainForm.MachineResponse = MachID
            MainForm.MachineName = MachID
            MainForm.MainResponse = Content
            MainForm.LocalIP = ValidatedIP
            ' MainForm.MainDomain = MainDomain.Replace(Chr(0), "")
            ' MainForm.Button1.Enabled = True
            MainForm.crytpo = Me.Crypto
            check.m_key1 = Crypto.LicenseKey
            check.m_key = Crypto.Key
            check.m_check = Second_Encrypted_Response
            check.m_ok = True
            check.m_result = EncryptedResponse
            Dim sb As StringBuilder = New StringBuilder
            sb.AppendLine(check.m_key1)
            sb.AppendLine(check.m_key)
            sb.AppendLine(check.m_check)
            sb.AppendLine(check.m_ok)
            sb.AppendLine(check.m_result)
            MainForm.MainDomains = MD5(sb.ToString())
            MainForm.checker = check
            My.Settings.LicenseKey = LicenseKey
            _Validated = True
            MainForm.Show()
        Catch ex As Exception
            MessageBox.Show("An Error Occured Launching Main Core (1) - " & ex.Message)
            End
        End Try
    End Sub

End Class
