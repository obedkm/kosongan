Public Class splash
    Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer) As Integer
    Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
    Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Integer, ByVal hRgn As Integer, ByVal bRedraw As Boolean) As Integer

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ProgressBar1.Increment(2)
        If ProgressBar1.Value = 2 Then
            Label1.Text = "read conf"
        End If
        If ProgressBar1.Value = 22 Then
            Label1.Text = "test connection "
        End If
        If ProgressBar1.Value = 42 Then
            Label1.Text = "acces database "
        End If
        If ProgressBar1.Value = 62 Then
            Label1.Text = "akses user "
        End If
        If ProgressBar1.Value = 98 Then
            Label1.Text = "finished "
        End If
        If ProgressBar1.Value = 100 Then
            LoginForm.Show()
            Timer1.Stop()
            Me.Close()

        End If
    End Sub

    Private Sub splash_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        If e.KeyCode = Keys.F10 Then
            Timer1.Enabled = False
            FormSetting.Show()

        End If
    End Sub

    Private Sub splash_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim i As Integer
        i = CreateRoundRectRgn(0, 0, Me.Width, Me.Height, 20, 20)
        SetWindowRgn(Me.Handle, i, 0)
        Timer1.Start()
    End Sub


    Private Sub btExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btExit.Click
        Me.Close()
    End Sub
End Class