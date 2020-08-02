Imports System.IO
Imports System.Net.Sockets
Imports System.Net.Security
Imports System.Net

Friend Class POP3

    Friend ActiveDatabase As String
    Private IsConnected As Boolean = False
    Private PopHost As String
    Private UserName As String
    Friend Original_UserName As String
    Private Password As String
    Private PortNm As Integer
    Private POP3 As New TcpClient
    Private Read_Stream As StreamReader
    Private NetworkS_tream As NetworkStream
    Private m_sslStream As SslStream
    Private server_Command As String
    Private ret_Val As Integer
    Private m_buffer() As Byte
    Private StatResp As String
    Private Server_Reponse As String
    Private Num_Emails As Integer = 0
    Friend MaxEmailsToScan As Integer = 10
    Private MailDotComDomains() As String = New String() {"mail.com", "email.com", "usa.com", "myself.com", "consultant.com", "post.com", "europe.com", "asia.com", "iname.com", "writeme.com", "dr.com", "engineer.com", "cheerful.com", "accountant.com", "techie.com", "uymail.com", "activist.com", "adexec.com", "allergist.com", "alumni.com", "alumnidirector.com", "angelic.com", "archaeologist.com", "arcticmail.com", "artlover.com", "bikerider.com", "birdlover.com", "brew-meister.com", "cash4u.com", "chemist.com", "clerk.com", "columnist.com", "comic.com", "computer4u.com", "counsellor.com", "cyberservices.com", "deliveryman.com", "diplomats.com", "disposable.com", "execs.com", "fastservice.com", "financier.com", "gardener.com", "geologist.com", "graphic-designer.com", "groupmail.com", "homemail.com", "hot-shot.com", "instruction.com", "insurer.com", "job4u.com", "journalist.com", "legislator.com", "lobbyist.com", "minister.com", "net-shopping.com", "optician.com", "pediatrician.com", "planetmail.com", "politician.com", "presidency.com", "priest.com", "publicist.com", "qualityservice.com", "realtyagent.com", "registerednurses.com", "repairman.com", "representative.com", "rescueteam.com", "sociologist.com", "solution4u.com", "tech-center.com", "technologist.com", "theplate.com", "toothfairy.com", "tvstar.com", "umpire.com", "webname.com", "worker.com", "workmail.com", "2trom.com", "aircraftmail.com", "atheist.com", "blader.com", "boardermail.com", "brew-master.com", "bsdmail.com", "catlover.com", "cutey.com", "dbzmail.com", "doglover.com", "doramail.com", "galaxyhit.com", "hackermail.com", "hilarious.com", "keromail.com", "kittymail.com", "lovecat.com", "marchmail.com", "nonpartisan.com", "petlover.com", "snakebite.com", "toke.com", "cyberdude.com", "cybergal.com", "cyber-wizard.com", "housemail.com", "inorbit.com", "mail-me.com", "rocketship.com", "acdcfan.com", "discofan.com", "elvisfan.com", "hiphopfan.com", "kissfans.com", "madonnafan.com", "metalfan.com", "ninfan.com", "ravemail.com", "reborn.com", "reggaefan.com", "californiamail.com", "dallasmail.com", "nycmail.com", "pacific-ocean.com", "pacificwest.com", "sanfranmail.com", "africamail.com", "asia-mail.com", "australiamail.com", "berlin.com", "brazilmail.com", "chinamail.com", "dublin.com", "dutchmail.com", "englandmail.com", "europemail.com", "germanymail.com", "irelandmail.com", "israelmail.com", "italymail.com", "koreamail.com", "mexicomail.com", "moscowmail.com", "munich.com", "polandmail.com", "safrica.com", "samerica.com", "scotlandmail.com", "spainmail.com", "swedenmail.com", "swissmail.com", "torontomail.com", "disciples.com", "innocent.com", "muslim.com", "protestant.com", "reincarnate.com", "religious.com", "saintly.com", "contractor.net", "appraiser.net", "auctioneer.net", "bartender.net", "chef.net", "coolsite.net", "fireman.net", "hairdresser.net", "instructor.net", "orthodontist.net", "photographer.net", "physicist.net", "planetmail.net", "programmer.net", "radiologist.net", "salesperson.net", "secretary.net", "socialworker.net", "songwriter.net", "surgical.net", "therapist.net", "greenmail.net", "humanoid.net", "null.net", "bellair.net", "linuxmail.org", "clubmember.org", "collector.org", "graduate.org", "musician.org", "teachers.org"}

    Friend Sub New(ByVal EmailAddress As String, ByVal EmailPassword As String, ByVal TmpActiveDatabase As String, Optional ByVal CustomPopHost As String = "", Optional ByVal CustomPopPort As String = "", Optional ByVal CustomPopUser As String = "")
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls
        ' Set Variables
        Original_UserName = EmailAddress
        UserName = EmailAddress
        Password = EmailPassword
        ActiveDatabase = TmpActiveDatabase
        If CustomPopUser <> Nothing Then UserName = CustomPopUser
        ' Remove Catchall
        If UserName.Contains("+") Then
            Dim PartToRemove As String = GetBetween(UserName, "+", "@")
            UserName = UserName.Replace("+" & PartToRemove, "")
        End If
        ' POP Settings
        If EmailAddress.ToLower().Contains("hotmail.") OrElse EmailAddress.ToLower().Contains("live.") OrElse EmailAddress.ToLower().Contains("outlook.") Then
            PopHost = "pop3.live.com"
            PortNm = 995
        ElseIf EmailAddress.ToLower().Contains("gmail.") OrElse EmailAddress.ToLower().Contains("googlemail.") Then
            PopHost = "pop.gmail.com"
            PortNm = 995
        ElseIf EmailAddress.ToLower().Contains("aol.") Then
            PopHost = "pop.aol.com"
            PortNm = 995
        ElseIf EmailAddress.ToLower().Contains("yahoo.") OrElse EmailAddress.ToLower().Contains("ymail.") Then
            PopHost = "pop.mail.yahoo.com"
            PortNm = 995
        ElseIf EmailAddress.ToLower().Contains("gmx.") Then
            PopHost = "pop.gmx.com"
            PortNm = 995
        ElseIf EmailAddress.ToLower().Contains("chewiemail.") OrElse EmailAddress.ToLower().Contains("clovermail.") OrElse EmailAddress.ToLower().Contains("expressmail.") OrElse EmailAddress.ToLower().Contains("mail-on.") OrElse EmailAddress.ToLower().Contains("manlymail.") OrElse EmailAddress.ToLower().Contains("offcolormail.") OrElse EmailAddress.ToLower().Contains("openmail.") OrElse EmailAddress.ToLower().Contains("powdermail.") OrElse EmailAddress.ToLower().Contains("tightmail.") OrElse EmailAddress.ToLower().Contains("toothandmail.") OrElse EmailAddress.ToLower().Contains("tushmail.") OrElse EmailAddress.ToLower().Contains("vfemail.") OrElse EmailAddress.ToLower().Contains("isonews2.") Then
            PopHost = "mail.vfemail.net"
            PortNm = 995
        ElseIf EmailAddress.ToLower().Contains("mail.ru") OrElse EmailAddress.ToLower().Contains("list.ru") OrElse EmailAddress.ToLower().Contains("mail.ua") OrElse EmailAddress.ToLower().Contains("bk.ru") OrElse EmailAddress.ToLower().Contains("inbox.ru") Then
            PopHost = "pop.mail.ru"
            PortNm = 995
        ElseIf EmailAddress.ToLower().Contains("mail.") Then
            PopHost = "pop.mail.com"
            PortNm = 995
        ElseIf EmailAddress.ToLower().Contains("daum.") Then
            PopHost = "pop.daum.net"
            PortNm = 995
        Else
            ' Check For Mail.com Other Domains
            Dim EmailProvider As String = Split(EmailAddress, "@")(1).ToLower()
            If MailDotComDomains.Contains(EmailProvider) Then
                PopHost = "pop.mail.com"
                PortNm = 995
            End If
        End If
        ' Check For Custom POP
        If PopHost = Nothing AndAlso CustomPopHost <> Nothing AndAlso CustomPopPort <> Nothing Then
            PopHost = CustomPopHost
            Try
                PortNm = CInt(CustomPopPort)
            Catch ex As Exception
                Throw New Exception("Invalid Port Provided For Email Account: " & EmailAddress & " (Port Provided: " & CustomPopPort & ")")
            End Try
        End If
        ' Check Host Avaiable
        If PopHost = Nothing Then
            Throw New Exception("Email Provider Not Currently Supported: " & EmailAddress)
        End If
    End Sub

    Friend Function FindEmail(ByVal TextToFind_1 As String, ByVal TextToFind_2 As String) As String
        ' Variables
        Dim EmailsChecked As Integer = 0
        ' Check Mail Server
        If PopHost = Nothing Then Throw New Exception("Email Account: " & UserName & " Not Supported For This Application")
        ' Check Connected
        If IsConnected = False Then Connect()
        ' Loop List
        For i As Integer = (CInt(Num_Emails) - 1) To 0 Step -1
            If EmailsChecked >= MaxEmailsToScan Then Exit For
            ' Download Email
            Dim TempEmailBody As String
            Try
                TempEmailBody = Retrieve((i + 1))
            Catch ex As Exception
                ' MessageBox.Show("Failed To Download Email Message: " & ex.Message.Trim())
                Continue For
            End Try
            If TempEmailBody.ToLower().Contains("" & TextToFind_1.ToLower() & "") AndAlso TempEmailBody.ToLower().Contains(TextToFind_2.ToLower()) Then
                Dim NewFormat As String = TempEmailBody.Replace("=" & vbCrLf, "")
                Return NewFormat
            End If
            EmailsChecked += 1
        Next
        ' Not Found Email
        Throw New Exception("Verification Email Not In Inbox")
    End Function

    Friend Overloads Sub Connect(Optional ByVal RetryAttempts As Integer = 0)
        Try
            ' Connect
            POP3 = New TcpClient(PopHost, PortNm)
            NetworkS_tream = POP3.GetStream()
            m_sslStream = New SslStream(NetworkS_tream)
            m_sslStream.AuthenticateAsClient(PopHost)
            Read_Stream = New StreamReader(m_sslStream)
            Read_Stream.BaseStream.ReadTimeout = 10000
            Server_Reponse = Read_Stream.ReadLine()
            ' If TestingMode = True Then MessageBox.Show("Connect_1 - " & Server_Reponse)
            ' Send Username
            server_Command = "USER " + UserName + vbCrLf
            m_buffer = System.Text.Encoding.ASCII.GetBytes(server_Command.ToCharArray())
            m_sslStream.Write(m_buffer, 0, m_buffer.Length)
            Server_Reponse = Read_Stream.ReadLine()
            ' If TestingMode = True Then MessageBox.Show("Connect_2 - " & Server_Reponse)
            ' Send Password
            server_Command = "PASS " + Password + vbCrLf
            m_buffer = System.Text.Encoding.ASCII.GetBytes(server_Command.ToCharArray())
            m_sslStream.Write(m_buffer, 0, m_buffer.Length)
            Server_Reponse = Read_Stream.ReadLine()
            ' If TestingMode = True Then MessageBox.Show("Connect_3 - " & Server_Reponse)
            ' Check Ok
            If Not Server_Reponse.Contains("OK") Then
                If Server_Reponse.Contains("password not accepted") OrElse Server_Reponse.Contains("Incorrect username") OrElse Server_Reponse.ToLower().Contains("authentication failed") OrElse Server_Reponse.Contains("unknown user name") Then ' Bad Username / Password / Disabled Account
                    Throw New Exception("Incorrect Login Information")
                ElseIf Server_Reponse.Contains("Web login required") Then ' Gmail Additional Authentication Required
                    Throw New Exception("Failed To Login To Account - Server Response: " & Server_Reponse)
                ElseIf Server_Reponse.Contains("SYS/TEMP") Then
                    Throw New Exception("Failed To Login To Account - Server Response: " & Server_Reponse & " (Allow Less Secure Apps Potentially Disabled)")
                ElseIf Server_Reponse.Contains("password has been compromised") Then
                    Throw New Exception("Account Requires A Password Reset Before Access")
                ElseIf Server_Reponse.Contains("too many active sessions") Then
                    If RetryAttempts > 3 Then
                        Throw New Exception("3 Failed Connection Attempts Occured In 30 Seconds Due To Too Many Active POP Connections To: " & PopHost)
                    Else
                        Threading.Thread.Sleep(10000)
                        RetryAttempts += 1
                        Connect(RetryAttempts)
                    End If
                Else
                    Log_Show_Error(Server_Reponse, "POP3_Connect_1", ActiveDatabase)
                    Throw New Exception("Failed To Login To Account - Server Response: " & Server_Reponse)
                End If
            End If
            ' Request Email Count
            server_Command = "STAT " + vbCrLf
            m_buffer = System.Text.Encoding.ASCII.GetBytes(server_Command.ToCharArray())
            m_sslStream.Write(m_buffer, 0, m_buffer.Length)
            Server_Reponse = Read_Stream.ReadLine()
            ' If TestingMode = True Then MessageBox.Show("Connect_4 - " & Server_Reponse)
            ' Set Number Of Emails
            Dim msgInfo() As String = Split(Server_Reponse, " "c)
            Num_Emails = CInt(msgInfo(1))
        Catch ex As Exception
            If ex.Message.Contains("period of time") Then
                Throw New Exception("The POP Server Timed Out / POP May Not Be Enabled On This Account")
            Else
                Throw New Exception(ex.Message)
            End If
        End Try
    End Sub

    Function SendCommand(ByVal SslStrem As SslStream, ByVal Server_Command As String) As String
        Server_Command = Server_Command & vbCrLf
        m_buffer = System.Text.Encoding.ASCII.GetBytes(Server_Command.ToCharArray())
        m_sslStream.Write(m_buffer, 0, m_buffer.Length)
        Read_Stream = New StreamReader(m_sslStream)
        Server_Reponse = Read_Stream.ReadLine()
        Return Server_Reponse
    End Function

    Public Function ListEmails() As List(Of String())
        Dim retval As New List(Of String())
        Dim List_Resp As String
        For i As Integer = 1 To Num_Emails
            List_Resp = SendCommand(m_sslStream, "LIST " & i.ToString)
            retval.Add(Split(List_Resp, " "))
        Next
        Return retval
    End Function

    Friend Sub Disconnect()
        StatResp = SendCommand(m_sslStream, "QUIT ") & vbCrLf
        POP3.Close()
        ret_Val = 0
    End Sub

    Public Function Retrieve(ByVal Index_Num As Integer) As String
        Dim XX As String
        Dim sZTMP As String
        Dim msg As String = ""
        XX = ("RETR " + Index_Num.ToString & vbCrLf)
        m_buffer = System.Text.Encoding.ASCII.GetBytes(XX.ToCharArray())
        m_sslStream.Write(m_buffer, 0, m_buffer.Length)
        Read_Stream = New StreamReader(m_sslStream)
        Read_Stream.ReadLine()
        sZTMP = Read_Stream.ReadLine + vbCrLf
        Do While Read_Stream.Peek <> -1
            sZTMP = Read_Stream.ReadLine + vbCrLf
            msg += (sZTMP)
        Loop
        Return msg
    End Function

End Class
