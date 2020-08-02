Imports System.Net
Imports System.Threading
Imports System.IO
Imports System.Text

Friend Class AccountCreator
    ' Variables
    Private m_MainWindow As Form1
    Private ActiveDatabase As String
    Private ItemSelector As New Object
    Private ProfileImageSelectionLock As New Object
    Private ProfileUpdater_UniqueImages As Boolean = True
    Private UseEachProxyOnceOnly As Boolean

    ' Sub New - Set Settings
    Friend Sub New(ByRef MainWindow As Form1, ByVal TmpActiveDatabase As String, ByVal TmpUseEachProxyOnceOnly As Boolean)
        m_MainWindow = MainWindow
        ActiveDatabase = TmpActiveDatabase
        UseEachProxyOnceOnly = TmpUseEachProxyOnceOnly
    End Sub
    ' Run Process
    Friend Sub RunProcess()
        Randomize()
        Try
            ' Loop Attempts
            While m_MainWindow.AccountCreator_AccountsCreated < m_MainWindow.AccountCreator_MaxAccountsToCreate AndAlso m_MainWindow.AccountCreator_StopProcess = False
                If m_MainWindow.AccountCreator_StopProcess = True Then Exit While
                ' Check Thread Count
                Dim AccountsLeftToCreate As Integer = (m_MainWindow.AccountCreator_MaxAccountsToCreate - m_MainWindow.AccountCreator_AccountsCreated)
                If AccountsLeftToCreate < m_MainWindow.AccountCreator_ThreadsRunning Then Exit While
                ' Account Variables
                Dim Gender As String = m_MainWindow.AccountCreator_ProfileGender
                Dim Username As String = ""
                Dim FirstName As String = ""
                Dim LastName As String = ""
                Dim Password As String = ""
                Dim Proxy As String = "0"
                Dim Port As String = "0"
                Dim ProxyUser As String = "0"
                Dim ProxyPass As String = "0"
                Dim EmailAddress As String = ""
                Dim EmailPassword As String = ""
                Dim BusinessName As String = ""
                ' Lock Account Variable Selection Process
                SyncLock ItemSelector
                    ' Check Counts
                    Try
                        If m_MainWindow.AccountCreator_Usernames.Count() < 1 Then Throw New Exception("No More Usernames Available")
                        If m_MainWindow.AccountCreator_EmailAccounts.Count() < 1 Then Throw New Exception("No More Email Available")
                        If m_MainWindow.AccountCreator_BusinessNames.Count() < 1 Then Throw New Exception("No More Business Names Available")
                    Catch ex As Exception
                        m_MainWindow.AddMessage(ex.Message)
                        Exit While
                    End Try
                    ' Select Proxy
                    If UseEachProxyOnceOnly = True Then
                        If m_MainWindow.AccountCreator_TempProxies.Count() < 1 Then
                            m_MainWindow.AddMessage("No More Proxies Available - Ending Account Creator Thread")
                            Exit While
                        Else
                            Proxy = m_MainWindow.AccountCreator_TempProxies(0)(0).Trim()
                            Port = CInt(m_MainWindow.AccountCreator_TempProxies(0)(1))
                            ProxyUser = m_MainWindow.AccountCreator_TempProxies(0)(2).Trim()
                            ProxyPass = m_MainWindow.AccountCreator_TempProxies(0)(3).Trim()
                            m_MainWindow.AccountCreator_TempProxies.RemoveAt(0)
                        End If
                    Else
                        If m_MainWindow.AccountCreator_Proxies.Count() > 0 Then
                            If m_MainWindow.AccountCreator_LastChosenProxy >= (m_MainWindow.AccountCreator_Proxies.Count() - 1) Then m_MainWindow.AccountCreator_LastChosenProxy = 0
                            Proxy = m_MainWindow.AccountCreator_Proxies(m_MainWindow.AccountCreator_LastChosenProxy)(0).Trim()
                            Port = CInt(m_MainWindow.AccountCreator_Proxies(m_MainWindow.AccountCreator_LastChosenProxy)(1))
                            ProxyUser = m_MainWindow.AccountCreator_Proxies(m_MainWindow.AccountCreator_LastChosenProxy)(2).Trim()
                            ProxyPass = m_MainWindow.AccountCreator_Proxies(m_MainWindow.AccountCreator_LastChosenProxy)(3).Trim()
                            m_MainWindow.AccountCreator_LastChosenProxy += 1
                        End If
                    End If
                    ' Select Variables
                    Username = m_MainWindow.AccountCreator_Usernames(0)
                    EmailAddress = m_MainWindow.AccountCreator_EmailAccounts(0)
                    BusinessName = m_MainWindow.AccountCreator_BusinessNames(0)
                    FirstName = m_MainWindow.AccountCreator_FirstNames(CInt(Rnd() * (m_MainWindow.AccountCreator_FirstNames.Count() - 1)))
                    LastName = m_MainWindow.AccountCreator_LastNames(CInt(Rnd() * (m_MainWindow.AccountCreator_LastNames.Count() - 1)))
                    ' Remove Variables
                    m_MainWindow.AccountCreator_Usernames.RemoveAt(0)
                    m_MainWindow.AccountCreator_EmailAccounts.RemoveAt(0)
                    m_MainWindow.AccountCreator_BusinessNames.RemoveAt(0)
                End SyncLock
                ' Check For Email Password
                If EmailAddress.Contains(":") Then
                    EmailPassword = Split(EmailAddress, ":")(1)
                    EmailAddress = Split(EmailAddress, ":")(0)
                End If
                ' Generate Strong Password
                Password = CreateStrongPassword()
                ' Start New HTTP Session
                Dim a_Process As CurlFunctions
                Try
                    a_Process = New CurlFunctions(Username, ActiveDatabase, Proxy, Port, ProxyUser, ProxyPass)
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Create New HTTP Process: " & ex.Message & " - " & ex.StackTrace)
                    m_MainWindow.Sleep(2, "AccountCreator")
                    Continue While
                End Try
                ' Set New CookieContainer In Case Of Existing Account
                a_Process.cookies = New CookieContainer
                ' Set Variables
                Dim URL As String = ""
                Dim Data As String = ""
                Dim Referrer As String = ""
                Dim Content As String = ""
                ' Load Homepage
                URL = "https://www.pinterest.com/"
                Try
                    Content = a_Process.Return_Content(URL)
                    Content = a_Process.Return_Content(URL)
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Pinterest Home Page URL: " & ex.Message)
                    m_MainWindow.Sleep(2, "AccountCreator")
                    Continue While
                End Try
                ' Scrape App Version
                Dim AppVersion As String = GetBetween(Content, """app_version"": """, """")
                If AppVersion = Nothing Then AppVersion = GetBetween(Content, """app_version"":""", """")
                ' Scrape CSRF Token
                Dim csrfmiddlewaretoken As String = GetSessionToken(a_Process)
                ' Mark Username & Email As Used
                WriteToFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\cache\" & "UsernameBlacklist.txt", Username)
                WriteToFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\cache\" & "EmailBlacklist.txt", EmailAddress & ":" & EmailPassword)
                WriteToFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\cache\" & "BusinessNameBlacklist.txt", BusinessName)
                ' Post Registration
                URL = "https://accounts.pinterest.com/v3/register/partner/handshake/"
                Data = "email=" & URLEncode(EmailAddress) & "&username=&password=" & URLEncode(Password) & "&first_name=" & URLEncode(FirstName) & "&last_name=" & URLEncode(LastName) & "&country=&locale=en-US&&business_name=" & URLEncode(BusinessName)
                Referrer = "https://www.pinterest.com/"
                Try
                    Content = a_Process.Post_Content(URL, Data, Referrer)
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Post To Register URL For Account:  " & EmailAddress & " - " & ex.Message)
                    m_MainWindow.Sleep(2, "AccountCreator")
                    Continue While
                End Try
                ' Confirm Ok
                ' Scrape Handshake Token
                Dim HandshakeToken As String = GetBetween(Content, """data"": """, """")
                ' Confirm Ok
                If Not Content.Contains("status"": ""success") OrElse HandshakeToken = Nothing Then
                    ' Check For Bot Detection
                    If Content.Contains("Sorry, we've blocked this registration") Then
                        m_MainWindow.AddMessage("Registration Blocked By Pinterest - IP: " & Proxy & " Has Been Detected As Spam / A Bot IP")
                        m_MainWindow.Sleep(2, "AccountCreator")
                        Continue While
                    ElseIf Content.Contains("Email already taken") Then
                        m_MainWindow.AddMessage("Registration Failed - Email Address Is Already In Use (" & EmailAddress & ")")
                        m_MainWindow.Sleep(2, "AccountCreator")
                        Continue While
                    Else
                        Log_Show_Error(Data & vbCrLf & vbCrLf & Content, "AccountCreator_1", ActiveDatabase)
                        m_MainWindow.AddMessage("Account Creation Failed For Account: " & EmailAddress & " - Unknown Issue")
                        m_MainWindow.Sleep(2, "AccountCreator")
                        Continue While
                    End If
                End If
                ' Report Success
                m_MainWindow.RefreshBoards("AccountsCreated")
                m_MainWindow.AddMessage("Successfully Created Account: " & EmailAddress)
                ' Set Full Account Data
                Dim FullAccountData() As String = New String() {EmailAddress, Password, Proxy, Port, ProxyUser, ProxyPass, Username, EmailPassword, "0", DateTime.Now.ToString()}
                ' Backup Acount
                WriteToFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\backups\" & "CreatedAccountsBackup.txt", Join(FullAccountData, ":"))
                m_MainWindow.AccountCreator_CreatedAccounts.Add(FullAccountData)
                ' Save Cookies For Later Use
                a_Process.Dispose()
                ' Post To Login With Handshake
                URL = "https://www.pinterest.com/resource/HandshakeSessionResource/create/"
                Data = "source_url=%2F&data=%7B%22options%22%3A%7B%22token%22%3A%22" & HandshakeToken & "%22%2C%22isRegistration%22%3Atrue%7D%2C%22context%22%3A%7B%7D%7D"
                Referrer = "https://www.pinterest.com/"
                Try
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, csrfmiddlewaretoken, Referrer, True, AppVersion, "active")
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Post Login Handshake For Account: " & EmailAddress & " - " & ex.Message)
                    m_MainWindow.Sleep(2, "AccountCreator")
                    Continue While
                End Try
                ' Load Homepage
                URL = "https://www.pinterest.com/"
                Try
                    Content = a_Process.Return_Content(URL, Referrer)
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Load Homepage After Creation For Account: " & EmailAddress & " - " & ex.Message)
                    m_MainWindow.Sleep(2, "AccountCreator")
                    Continue While
                End Try
                ' Scrape  Experience ID
                Dim ExperienceID As String = GetBetween(Content, """experience_id"":", ",")
                ' Confirm Available
                If ExperienceID = Nothing Then
                    Log_Show_Error(Content, "AccountCreator_2", ActiveDatabase)
                    m_MainWindow.AddMessage("Failed To Scrape Experience ID For Account: " & EmailAddress & " - Unknown Issue")
                    m_MainWindow.Sleep(2, "AccountCreator")
                    Continue While
                End If
                ' Set Locale To US
                URL = "https://www.pinterest.com/_ngjs/resource/UserSettingsResource/update/"
                Data = "source_url=%2F&data=%7B%22options%22%3A%7B%22country%22%3A%22US%22%2C%22locale%22%3A%22en-US%22%7D%2C%22context%22%3A%7B%7D%7D"
                Try
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Post Locale Data For Account: " & EmailAddress & " - " & ex.Message)
                    m_MainWindow.Sleep(2, "AccountCreator")
                    Continue While
                End Try
                ' Post Page 2
                URL = "https://www.pinterest.com/_ngjs/resource/UserStateResource/create/"
                Data = "source_url=%2F&data=%7B%22options%22%3A%7B%22state%22%3A%22NUX_COUNTRY_LOCALE_STEP_STATUS%22%2C%22value%22%3A2%7D%2C%22context%22%3A%7B%7D%7D"
                Try
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Post Locale Data For Account: " & EmailAddress & " - " & ex.Message)
                    m_MainWindow.Sleep(2, "AccountCreator")
                    Continue While
                End Try
                ' Set Business Name & Type
                URL = "https://www.pinterest.com/_ngjs/resource/UserSettingsResource/update/"
                Data = "source_url=%2F&data=%7B%22options%22%3A%7B%22account_type%22%3A%22" & m_MainWindow.AccountCreator_BusinessType & "%22%2C%22business_name%22%3A%22" & URLEncode(BusinessName) & "%22%7D%2C%22context%22%3A%7B%7D%7D"
                Try
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Set Business Name For Account: " & EmailAddress & " - " & ex.Message)
                    m_MainWindow.Sleep(2, "AccountCreator")
                    Continue While
                End Try
                ' Set Advertising Intent
                URL = "https://www.pinterest.com/_ngjs/resource/UserSettingsResource/update/"
                Data = "source_url=%2F&data=%7B%22options%22%3A%7B%22advertising_intent%22%3A3%7D%2C%22context%22%3A%7B%7D%7D"
                Try
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Set Advertising Intent For Account: " & EmailAddress & " - " & ex.Message)
                    m_MainWindow.Sleep(2, "AccountCreator")
                    Continue While
                End Try
                ' Complete Experience
                Try
                    ' Post Next Page
                    URL = "https://www.pinterest.com/_ngjs/resource/UserStateResource/create/"
                    Data = "source_url=%2F&data=%7B%22options%22%3A%7B%22state%22%3A%22NUX_TOPIC_PICKER_STEP_STATUS%22%2C%22value%22%3A1%7D%2C%22context%22%3A%7B%7D%7D"
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                    ' Select Random Interest
                    Dim AllInterests As List(Of String) = New String() {"935249274030", "927552077382", "961238559656", "948264532184", "902065567321", "898037969690", "924255633983", "908182459161", "955506047789", "895028262457", "922203297757", "905860166503", "956143059915", "948192800438", "903260720461", "949916069751", "906537459533", "944938374183", "922134410098", "916477076637", "932441562566", "898977956499", "920923736524", "903884479167", "918644201389", "925056443165", "906615418398", "944798509834", "920779196416", "905661505034", "903847091747", "902794668380", "936760250909", "901809349518", "921358921047", "960796800398", "911373384406", "896604368580", "906224180441", "959339805947", "918105274631", "919296720236", "916919860042", "895527461239", "951457072235", "950635673676", "897357639793", "932986458342", "933384436092", "955847418782", "900357664296", "925089719884", "907551614055", "941062423970", "924742740590", "954798921443", "916216778476", "922993631942", "953543301552", "929229643961", "913976725415", "954982871150", "901682362753", "945391946569", "929900110692", "946028385883", "958001935551", "920236059316", "901080083157", "951844471953", "943263428449", "951597319263", "901374657300", "908410600882", "909721864236", "912966607302", "923753523743", "952738239437", "901468450490", "929113501072", "955209809744", "899049071324", "928107717621", "935793879023", "955651259944", "946765704754", "901179409185", "943097620043", "905708509733", "920800886925", "956464956665", "951797652334", "928284437325", "899802232626", "954929326772", "909588660405", "897943516645", "901766345242", "945584629411", "959317080636"}.ToList()
                    Dim Interests As String = ""
                    Dim RandInterestsToSet As Integer = GetRandomNumber(1, 5)
                    For i As Integer = 1 To RandInterestsToSet
                        Dim TmpInterest As String = AllInterests(Rnd() * (AllInterests.Count() - 1))
                        AllInterests.Remove(TmpInterest)
                        Interests = TmpInterest & "," & Interests
                    Next
                    ' Remove Trailing Comma
                    Interests = Interests.Substring(0, Interests.Length - 1)
                    ' Set Interests
                    URL = "https://www.pinterest.com/_ngjs/resource/OrientationContextResource/create/"
                    Data = "source_url=%2F&data=%7B%22options%22%3A%7B%22interests%22%3A%22" & Interests & "%22%2C%22log_data%22%3A%22%5B%5C%22%7B%5C%5C%5C%22source%5C%5C%5C%22%3A%5C%5C%5C%22smart_recs%5C%5C%5C%22%2C%5C%5C%5C%22pos%5C%5C%5C%22%3A0%2C%5C%5C%5C%22gender%5C%5C%5C%22%3A%5C%5C%5C%22unspecified%5C%5C%5C%22%7D%5C%22%2C%5C%22%7B%5C%5C%5C%22source%5C%5C%5C%22%3A%5C%5C%5C%22smart_recs%5C%5C%5C%22%2C%5C%5C%5C%22pos%5C%5C%5C%22%3A1%2C%5C%5C%5C%22gender%5C%5C%5C%22%3A%5C%5C%5C%22unspecified%5C%5C%5C%22%7D%5C%22%2C%5C%22%7B%5C%5C%5C%22source%5C%5C%5C%22%3A%5C%5C%5C%22smart_recs%5C%5C%5C%22%2C%5C%5C%5C%22pos%5C%5C%5C%22%3A2%2C%5C%5C%5C%22gender%5C%5C%5C%22%3A%5C%5C%5C%22unspecified%5C%5C%5C%22%7D%5C%22%2C%5C%22%7B%5C%5C%5C%22source%5C%5C%5C%22%3A%5C%5C%5C%22smart_recs%5C%5C%5C%22%2C%5C%5C%5C%22pos%5C%5C%5C%22%3A4%2C%5C%5C%5C%22gender%5C%5C%5C%22%3A%5C%5C%5C%22unspecified%5C%5C%5C%22%7D%5C%22%2C%5C%22%7B%5C%5C%5C%22source%5C%5C%5C%22%3A%5C%5C%5C%22smart_recs%5C%5C%5C%22%2C%5C%5C%5C%22pos%5C%5C%5C%22%3A3%2C%5C%5C%5C%22gender%5C%5C%5C%22%3A%5C%5C%5C%22unspecified%5C%5C%5C%22%7D%5C%22%2C%5C%22%7B%5C%5C%5C%22source%5C%5C%5C%22%3A%5C%5C%5C%22smart_recs%5C%5C%5C%22%2C%5C%5C%5C%22pos%5C%5C%5C%22%3A9%2C%5C%5C%5C%22gender%5C%5C%5C%22%3A%5C%5C%5C%22unspecified%5C%5C%5C%22%7D%5C%22%2C%5C%22%7B%5C%5C%5C%22source%5C%5C%5C%22%3A%5C%5C%5C%22smart_recs%5C%5C%5C%22%2C%5C%5C%5C%22pos%5C%5C%5C%22%3A8%2C%5C%5C%5C%22gender%5C%5C%5C%22%3A%5C%5C%5C%22unspecified%5C%5C%5C%22%7D%5C%22%2C%5C%22%7B%5C%5C%5C%22source%5C%5C%5C%22%3A%5C%5C%5C%22smart_recs%5C%5C%5C%22%2C%5C%5C%5C%22pos%5C%5C%5C%22%3A7%2C%5C%5C%5C%22gender%5C%5C%5C%22%3A%5C%5C%5C%22unspecified%5C%5C%5C%22%7D%5C%22%2C%5C%22%7B%5C%5C%5C%22source%5C%5C%5C%22%3A%5C%5C%5C%22smart_recs%5C%5C%5C%22%2C%5C%5C%5C%22pos%5C%5C%5C%22%3A5%2C%5C%5C%5C%22gender%5C%5C%5C%22%3A%5C%5C%5C%22unspecified%5C%5C%5C%22%7D%5C%22%2C%5C%22%7B%5C%5C%5C%22source%5C%5C%5C%22%3A%5C%5C%5C%22smart_recs%5C%5C%5C%22%2C%5C%5C%5C%22pos%5C%5C%5C%22%3A6%2C%5C%5C%5C%22gender%5C%5C%5C%22%3A%5C%5C%5C%22unspecified%5C%5C%5C%22%7D%5C%22%5D%22%2C%22referrer%22%3A%22nux%22%2C%22user_behavior_data%22%3A%22%7B%5C%22signupInterestsPickerTimeSpent%5C%22%3A" & GetRandomNumber(100, 2000).ToString() & "%7D%22%7D%2C%22context%22%3A%7B%7D%7D"
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                    ' Post Next Page
                    URL = "https://www.pinterest.com/_ngjs/resource/UserStateResource/create/"
                    Data = "source_url=%2F&data=%7B%22options%22%3A%7B%22state%22%3A%22NUX_TOPIC_PICKER_STEP_STATUS%22%2C%22value%22%3A2%7D%2C%22context%22%3A%7B%7D%7D"
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                    ' Post Next Page
                    URL = "https://www.pinterest.com/_ngjs/resource/UserStateResource/create/"
                    Data = "source_url=%2F&data=%7B%22options%22%3A%7B%22state%22%3A%22NUX_BROWSER_EXT_STEP_STATUS%22%2C%22value%22%3A1%7D%2C%22context%22%3A%7B%7D%7D"
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                    ' Post Next Page
                    URL = "https://www.pinterest.com/_ngjs/resource/UserStateResource/create/"
                    Data = "source_url=%2F&data=%7B%22options%22%3A%7B%22state%22%3A%22NUX_BROWSER_EXT_STEP_STATUS%22%2C%22value%22%3A2%7D%2C%22context%22%3A%7B%7D%7D"
                    ' Post Next Page
                    URL = "https://www.pinterest.com/_ngjs/resource/UserExperienceCompletedResource/update/"
                    Data = "source_url=%2F&data=%7B%22options%22%3A%7B%22placed_experience_id%22%3A%2211%253A" & ExperienceID & "%22%2C%22extra_context%22%3A%7B%7D%7D%2C%22context%22%3A%7B%7D%7D"
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                    ' Post Next Page
                    URL = "https://www.pinterest.com/_ngjs/resource/UserExperienceViewedResource/update/"
                    Data = "source_url=%2F" & Username.ToLower() & "&data=%7B%22options%22%3A%7B%22placed_experience_id%22%3A%2219%253A" & ExperienceID & "%22%7D%2C%22context%22%3A%7B%7D%7D"
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                    ' Post Next Page
                    URL = "https://www.pinterest.com/_ngjs/resource/UserExperienceResource/delete/"
                    Data = "source_url=%2F" & Username.ToLower() & "%2F&data=%7B%22options%22%3A%7B%22placed_experience_id%22%3A%2219%253A500925%22%7D%2C%22context%22%3A%7B%7D%7D"
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                    ' Set Account Gender
                    URL = "https://www.pinterest.com/_ngjs/resource/UserSettingsResource/update/"
                    If Gender = "Female" Then
                        Data = "source_url=%2Fsettings%2Faccount-settings&data=%7B%22options%22%3A%7B%22gender%22%3A%22female%22%7D%2C%22context%22%3A%7B%7D%7D"
                    Else
                        Data = "source_url=%2Fsettings%2Faccount-settings&data=%7B%22options%22%3A%7B%22gender%22%3A%22male%22%7D%2C%22context%22%3A%7B%7D%7D"
                    End If
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                    ' Set Business Description
                    URL = "https://www.pinterest.com/_ngjs/resource/UserSettingsResource/update/"
                    Data = "source_url=%2Fsettings%2F&data=%7B%22options%22%3A%7B%22about%22%3A%22" & URLEncode(BusinessName) & "%22%7D%2C%22context%22%3A%7B%7D%7D"
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                    ' Set Business Type
                    URL = "https://www.pinterest.com/_ngjs/resource/UserSettingsResource/update/"
                    Data = "source_url=%2Fsettings%2Faccount-settings&data=%7B%22options%22%3A%7B%22account_type%22%3A%22" & m_MainWindow.AccountCreator_BusinessType & "%22%7D%2C%22context%22%3A%7B%7D%7D"
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                    ' Set Location To US
                    URL = "https://www.pinterest.com/_ngjs/resource/UserSettingsResource/update/"
                    Data = "source_url=%2Fsettings%2Fedit-profile&data=%7B%22options%22%3A%7B%22location%22%3A%22US%22%7D%2C%22context%22%3A%7B%7D%7D"
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Set Experience Complete For Account: " & EmailAddress & " - " & ex.Message)
                    m_MainWindow.Sleep(2, "AccountCreator")
                    Continue While
                End Try
                ' Report Success
                m_MainWindow.AddMessage("Successfully Completed Experience Setup For Account: " & EmailAddress)
                ' Load Edit Profile URL
                URL = "https://www.pinterest.com/settings/edit-profile"
                Try
                    Content = a_Process.Return_Content(URL)
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Load Profile Page After Creation For Account: " & EmailAddress & " - " & ex.Message)
                    GoTo BoardCreation
                End Try
                ' Scrape Current Username
                Dim CurUsername As String = GetBetween(Content, "name=""username"" value=""", """")
                ' Check If Username Ok
                If CurUsername <> Nothing AndAlso CurUsername.ToLower() = Username.ToLower() Then GoTo BoardCreation
                ' Set Account Username
                Dim FoundOkUsername As Boolean = False
                Dim UsernameAttempts As Integer = 0
                While FoundOkUsername = False
                    ' Add To Blacklist
                    WriteToFile(My.Application.Info.DirectoryPath & "\" & ActiveDatabase & "\cache\" & "UsernameBlacklist.txt", Username)
                    ' Update Username
                    URL = "https://www.pinterest.com/_ngjs/resource/UserSettingsResource/update/"
                    Data = "source_url=%2Fsettings%2Fedit-profile&data=%7B%22options%22%3A%7B%22username%22%3A%22" & Username.ToLower() & "%22%7D%2C%22context%22%3A%7B%7D%7D"
                    Try
                        Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                    Catch ex As Exception
                        m_MainWindow.AddMessage("Failed To Set Username For Account: " & EmailAddress & " - " & ex.Message)
                        GoTo BoardCreation
                    End Try
                    ' Confirm Ok
                    If Not Content.Contains("status"":""success") Then
                        If Content.Contains("already taken") Then
                            m_MainWindow.AddMessage("Failed To Set Username For Account: " & EmailAddress & " - Username Already Taken: " & Username)
                        ElseIf Content.Contains("Invalid username") Then
                            m_MainWindow.AddMessage("Failed To Set Username For Account: " & EmailAddress & " - Username Is Not Valid: " & Username)
                        Else
                            Log_Show_Error(Content, "AccountCreator_5", ActiveDatabase)
                            m_MainWindow.AddMessage("Failed To Set Username For Account: " & EmailAddress & " - Unknown Issue")
                        End If
                    Else
                        m_MainWindow.AddMessage("Successfully Set Username: " & Username & " For Account: " & EmailAddress)
                        Exit While
                    End If
                    ' Wait
                    UsernameAttempts += 1
                    If UsernameAttempts > 5 Then
                        m_MainWindow.AddMessage("Failed To Set Username For Account: " & EmailAddress & " After 5 Consecutive Attempts")
                        GoTo BoardCreation
                    End If
                    ' Add Additional Character
                    Dim LastChar As String = Username(Username.Length - 1)
                    Username = Username & LastChar
                    ' Wait
                    m_MainWindow.Sleep(2, "AccountCreator")
                End While
BoardCreation:
                ' Loop Boards To Create
                For Each TmpBoard As String In m_MainWindow.AccountCreator_BoardsToCreate
                    ' Set Data
                    Dim BoardName As String = TmpBoard
                    Dim BoardCategory As String = ""
                    ' Check For Category
                    If BoardName.Contains(":") Then
                        BoardCategory = Split(BoardName, ":")(1)
                        BoardName = Split(BoardName, ":")(0)
                    End If
                    ' Create Board
                    Try
                        m_MainWindow.CreateBoard(a_Process, BoardName, Username, BoardCategory)
                        m_MainWindow.AddMessage("Successfully Created Board: " & BoardName & " For Account: " & EmailAddress)
                    Catch ex As Exception
                        m_MainWindow.AddMessage("Failed To Create Board: " & BoardName & " For Account: " & EmailAddress & " - " & ex.Message)
                        m_MainWindow.Sleep(2, "AccountCreator")
                        Continue For
                    End Try
                    ' Wait
                    m_MainWindow.Sleep(2, "AccountCreator")
                Next
                ' Wrap Image Selection Process
                Dim ProfileImagePath As String
                Try
                    ' Get Image Files
                    Dim files() As String = Directory.GetFiles(m_MainWindow.AccountCreator_ProfileImagesFolder)
                    Dim ImageFiles = files.Where(Function(f) f.ToLower().EndsWith(".png"))
                    If ImageFiles.Count() < 1 Then
                        ImageFiles = files.Where(Function(f) f.ToLower().EndsWith(".jpg"))
                        If ImageFiles.Count() < 1 Then
                            ImageFiles = files.Where(Function(f) f.ToLower().EndsWith(".jpeg"))
                        End If
                    End If
                    ' Check Images Available
                    If ImageFiles.Count() < 1 Then
                        m_MainWindow.AddMessage("Out Of Profile Images To Upload For Account: " & EmailAddress)
                        GoTo WaitDelay
                    End If
                    ' Select Profile Image
                    SyncLock ProfileImageSelectionLock
                        ProfileImagePath = ImageFiles(CInt(Rnd() * (ImageFiles.Count() - 1)))
                        Dim Filename As String = ProfileImagePath.Split(CChar("\")).Last
                        Dim NewFilePath As String = m_MainWindow.AccountCreator_ProfileImagesFolder & "\used\" & Filename
                        ' Create Directory
                        Try
                            If Directory.Exists(m_MainWindow.AccountCreator_ProfileImagesFolder & "\used\") = False Then Directory.CreateDirectory(m_MainWindow.AccountCreator_ProfileImagesFolder & "\used\")
                        Catch ex As Exception
                            Throw New Exception("Failed To Create Directory: " & m_MainWindow.AccountCreator_ProfileImagesFolder & "\used\" & " - " & ex.Message)
                        End Try
                        ' If Unique Move Image To Used Folder
                        If ProfileUpdater_UniqueImages = True Then
                            ' Move Image
                            File.Copy(ProfileImagePath, NewFilePath, True)
                            ' Delete Old Image
                            File.Delete(ProfileImagePath)
                            ' Change File Path For Upload
                            ProfileImagePath = NewFilePath
                        End If
                    End SyncLock
                Catch ex As Exception
                    m_MainWindow.AddMessage("An Error Occured While Selecting A Profile Image: " & ex.Message & ". Skipping Account (" & EmailAddress & ")")
                    GoTo WaitDelay
                End Try
                ' Upload Image
                Try
                    Content = a_Process.UploadProfileImage(ProfileImagePath, GetSessionToken(a_Process), Username)
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Upload Profile Image For Account: " & EmailAddress & " - " & ex.Message)
                    GoTo WaitDelay
                End Try
                ' Scrape Uploaded Image URL
                Dim UploadedImage As String = GetBetween(Content, """image_url"": """, """")
                ' Confirm Available
                If UploadedImage = Nothing Then
                    Log_Show_Error(Content, "AccountCreator_3", ActiveDatabase)
                    m_MainWindow.AddMessage("Failed To Upload Profile Image For Account: " & EmailAddress & " - Unknown Issue")
                    GoTo WaitDelay
                    Continue While
                End If
                ' Update Profile Image
                URL = "https://www.pinterest.com/resource/UserSettingsResource/update/"
                Data = "source_url=%2Fsettings%2F&data=%7B%22options%22%3A%7B%22profile_image_url%22%3A%22" & URLEncode(UploadedImage) & "%22%7D%2C%22context%22%3A%7B%7D%7D"
                Referrer = "https://www.pinterest.com/" & URLEncode(Username.ToLower()) & "/"
                Try
                    Content = a_Process.XMLHttpRequest_Post_Request(URL, Data, GetSessionToken(a_Process), Referrer, True, AppVersion, "active")
                Catch ex As Exception
                    m_MainWindow.AddMessage("Failed To Upload Profile Image (2) For Account: " & EmailAddress & " - " & ex.Message)
                    GoTo WaitDelay
                End Try
                ' Confirm Successful
                If Not Content.Contains("""error"": null") AndAlso Not Content.Contains("""error"":null") AndAlso Not Content.Contains("status"":""success") Then
                    m_MainWindow.AddMessage("Failed To Upload Profile Image For Account: " & EmailAddress & " - Unknown Issue")
                    Log_Show_Error(Content, "AccountCreator_4", ActiveDatabase)
                    GoTo WaitDelay
                End If
                ' Report Success
                m_MainWindow.AddMessage("Successfully Uploaded Profile Image For Account: " & EmailAddress)
WaitDelay:
                ' Add Account To Main List
                If m_MainWindow.AccountCreator_AutoAddAccounts = True Then
                    m_MainWindow.Accounts.Add({EmailAddress, Password, Proxy, Port, ProxyUser, ProxyPass, "1", "0", "0"})
                    ' Refresh Account List
                    m_MainWindow.RefreshBoards("RefreshAccounts")
                End If
                ' Wait Delay
                m_MainWindow.Sleep(m_MainWindow.AccountCreator_DelayBetweenCreation, "AccountCreator")
            End While
            ' Close Thread
            m_MainWindow.RefreshBoards("AccountCreator_ProcessComplete")
        Catch ex As Exception
            MessageBox.Show("A Fatal Error Occured: " & ex.Message & " - " & ex.StackTrace)
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
