Imports System
Imports System.Data.SqlClient
Imports System.Text


Public Class LoginForm

    Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer) As Integer
    Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
    Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Integer, ByVal hRgn As Integer, ByVal bRedraw As Boolean) As Integer

    Dim help As New helper

    Private Sub LoginForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim i As Integer
        i = CreateRoundRectRgn(0, 0, Me.Width, Me.Height, 20, 20)
        SetWindowRgn(Me.Handle, i, 0)
    End Sub

    Private Sub btexit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btexit.Click
        Me.Close()
    End Sub

    Private Sub btlogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btlogin.Click
        'cek password, cek loginmodul, cek accesslist
        MenuUtama.loginAsCode = help.getlogincode(cmbUser.Text)
        MenuUtama.loginModul = "01"

        If cmbUser.Text = "" Then
            MsgBox("user name masih kosong")
        Else
            MenuUtama.tsuser.Text = cmbUser.Text

            MenuUtama.Show()
            Me.Close()
        End If
    End Sub

End Class
