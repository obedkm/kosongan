Imports System.Security.Cryptography
Imports System.Text

Public Class FormSetting
    Dim fnc As New helper


    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        My.Settings.DatabaseName = fnc.AES_Encrypt(txtdbname.Text)
        My.Settings.DataSource = fnc.AES_Encrypt(txtserver.Text)
        My.Settings.Password = fnc.AES_Encrypt(txtpassword.Text)
        My.Settings.User = fnc.AES_Encrypt(txtuser.Text)
        My.Settings.reportsource = fnc.AES_Encrypt(txtreport.Text)
        splash.Timer1.Enabled = True
        Me.Close()
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If fnc.cekKonfigurasi = False Then
            txtdbname.Text = ""
            txtpassword.Text = ""
            txtreport.Text = ""
            txtserver.Text = ""
            txtuser.Text = ""
        Else
            txtdbname.Text = fnc.AES_Decrypt(My.Settings.DatabaseName)
            txtpassword.Text = fnc.AES_Decrypt(My.Settings.Password)
            txtreport.Text = fnc.AES_Decrypt(My.Settings.reportsource)
            txtserver.Text = fnc.AES_Decrypt(My.Settings.DataSource)
            txtuser.Text = fnc.AES_Decrypt(My.Settings.User)
        End If
    End Sub
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        splash.Timer1.Enabled = True
        Me.Close()
    End Sub


   

    End Class