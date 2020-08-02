Imports System.Threading

Friend Class WSOLoginForm

    Friend checker As Class1

    Private Sub LoginForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim frm As Form1 = New Form1
        Me.Icon = frm.Icon
        frm.Dispose()
        frm = Nothing

        Dim LicenseKey As String = GetSetting(Form1.SoftwareName, "textboxes", "LicenseKey", "")
        If LicenseKey <> "" Then
            Dim m_check As Class1 = New Class1
            Dim Lic As WSOLicensing = New WSOLicensing(LicenseKey, True, m_check)
            If Lic.Validated <> True Or m_check.m_ok <> True _
                Or String.IsNullOrEmpty(m_check.m_key1) _
                Or String.IsNullOrEmpty(m_check.m_key) _
                Or String.IsNullOrEmpty(m_check.m_result) _
                Or String.IsNullOrEmpty(m_check.m_check) Then
                MessageBox.Show(Lic.ErrorMsg)
                TextBox1.Text = LicenseKey
            Else
                ' Form1.Show()
                checker = New Class1
                checker.m_ok = True
                Me.Close()
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim m_check As Class1 = New Class1
        Dim Lic As WSOLicensing = New WSOLicensing(TextBox1.Text, True, m_check)
        If Lic.Validated <> True Or m_check.m_ok <> True _
                Or String.IsNullOrEmpty(m_check.m_key1) _
                Or String.IsNullOrEmpty(m_check.m_key) _
                Or String.IsNullOrEmpty(m_check.m_result) _
                Or String.IsNullOrEmpty(m_check.m_check) Then
            MessageBox.Show(Lic.ErrorMsg)
        Else
            ' Form1.Show()
            checker = New Class1
            checker.m_ok = True
            Me.Close()
        End If
    End Sub
End Class