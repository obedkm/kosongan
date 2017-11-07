Imports System.Windows.Forms
Imports System.Data.SqlClient
Imports System.Reflection

Public Class MenuUtama

    Public loginModul As String = ""
    Public loginAs As String = ""
    Public loginAsCode As String = ""
    Dim help As New helper
    Dim accesmenu As String
    Dim arrAccesmenu() As String

#Region "Window menu strip"
    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileVerticalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileHorizontalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ArrangeIconsToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub
#End Region

    Private Sub Menu_UtamaClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If (MessageBox.Show("Close?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No) Then
            e.Cancel = True
        End If
    End Sub

    Private Sub MenuUtama_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Interval = 1
        Timer1.Start()
        '// get main form title as module login
        Me.Text = help.getmodulname(loginModul)
        '//cek acces as login and get all menu
        accesmenu = help.getaccesslist(loginModul, loginAsCode)
        Me.tsLogin.Text = help.getloginas(loginAsCode)
        Me.TsTanggal.Text = Now.Date.ToShortDateString
        arrAccesmenu = accesmenu.Split(";")
        isiHeaderMenu(MenuStrip)
    End Sub

    'load menu header berdasarkan modul dan login user
    Private Sub isiHeaderMenu(ByRef menustrp As MenuStrip)
        Try
            Using Hmenu As New Connectio
                Hmenu.SQL = String.Format("select MenuItem,MenuNama from TMenuItem where MenuJenis='H' and MenuKode=@kode")
                Hmenu.OpenConnection()
                Hmenu.InitializeCommand()
                Dim dmn As SqlDataReader
                Hmenu.AddParameterWithValue("@kode", loginModul)
                dmn = Hmenu.Command.ExecuteReader
                Dim x As Integer = 1
                If dmn.HasRows Then
                    While dmn.Read
                        If cek_access_list(dmn(0).ToString) = True Then
                            menustrp.Items.Add(dmn(1).ToString)
                            Me.Controls.Add(menustrp)
                            sub_menu(menustrp.Items(x))
                            x += 1
                        End If
                    End While
                    Me.MenuStrip.Items.Add(WindowsMenu)
                End If
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "error load menu")
        End Try
    End Sub

    'load sub menu berdasarkn header menu, login user, dan accesslist dari [TAccessItem]
    Public Sub sub_menu(ByVal sender As Object)
        Dim cms As New ContextMenuStrip
        Dim smnu() As String
        Dim pmid As String = ""
        Try
            Using smenu As New Connectio
                smenu.SQL = ""
                smenu.SQL = "select MenuItem from TMenuItem where MenuNama='" & sender.ToString & "' and MenuJenis='H' and MenuKode=@kode"
                smenu.OpenConnection()
                smenu.InitializeCommand()
                Dim dsm As SqlDataReader
                smenu.AddParameterWithValue("@kode", loginModul)
                dsm = smenu.Command.ExecuteReader
                If dsm.HasRows Then
                    dsm.Read()
                    pmid = dsm(0)
                End If
                smenu.CloseConnection()
            End Using
        Catch
        End Try

        Try
            Using smenu As New Connectio
                smenu.SQL = "'"
                smenu.SQL = "select MenuItem,MenuNama from TMenuITem where MenuItem like '" & pmid.Substring(0, 1) & "%' and  MenuJenis='M' and MenuKode=@kode"
                smenu.OpenConnection()
                smenu.InitializeCommand()
                Dim smn As SqlDataReader
                smenu.AddParameterWithValue("@kode", loginModul)
                smn = smenu.Command.ExecuteReader
                ReDim Preserve smnu(0)
                Dim i As Integer
                If smn.HasRows Then
                    i = 0
                    While smn.Read()
                        If cek_access_list(smn(0).ToString) = True Then
                            ReDim Preserve smnu(i)
                            smnu(i) = smn("MenuNama")
                            i = i + 1
                        End If
                    End While
                End If
                smenu.CloseConnection()
                For Each anakmenu As String In smnu
                    cms.Items.Add(anakmenu, Nothing, New System.EventHandler(AddressOf SelectedChildMenu_OnClick))
                Next
                Dim tsi As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
                tsi.DropDown = cms
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, , "error load submenu")
        End Try
    End Sub

    'open  form from menu name
    Private Sub SelectedChildMenu_OnClick(ByVal sender As Object, ByVal e As System.EventArgs)
        '// dari MenuNama cari MenuItem dari tabel - menuitem juga di simpan di tag masing-masing form
        '// loop dari assembly unntuk dapat nama form nya
        Dim idItem As String = ""
        Try
            Using menuitem As New Connectio
                menuitem.SQL = String.Format("select * from TMenuItem where MenuKode='{0}' and MenuJenis='M' and MenuNama='{1}'", loginModul, sender.ToString)
                menuitem.OpenConnection()
                menuitem.InitializeCommand()
                Dim mnitem As SqlDataReader
                mnitem = menuitem.Command.ExecuteReader
                If mnitem.HasRows Then
                    mnitem.Read()
                    idItem = mnitem("MenuKode").ToString & mnitem("MenuItem").ToString & mnitem("MenuJenis").ToString
                End If
                menuitem.CloseConnection()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        Try
            Dim objForm As Form
            Dim sValue As String
            Dim FullTypeName As String
            Dim FormInstanceType As Type
            sValue = sender.ToString.Replace(" ", "_")
            FullTypeName = Application.ProductName & "." & sValue
            FormInstanceType = Type.GetType(FullTypeName, True, True)
            objForm = CType(Activator.CreateInstance(FormInstanceType), Form)

            Dim asmName As String = My.Application.Info.AssemblyName.ToString
            Dim asm As Assembly = Assembly.GetExecutingAssembly
            For Each aType As Type In asm.GetTypes
                If aType.BaseType.Name.ToUpper = "FORM" Then
                    If aType.Name.Contains("Menu") Or aType.Name.Contains("Setting") Or aType.Name.Contains("splash") Or aType.Name.Contains("Login") Then Continue For
                    Dim type As Type = Assembly.GetExecutingAssembly().GetType(asmName & "." & aType.Name.ToString)
                    Dim form As Form = DirectCast(Activator.CreateInstance(type), Form)
                    If form.Tag.ToString = idItem Then
                        form.MdiParent = Me
                        form.Text = sender.ToString
                        form.Show()
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("Menu " & sender.ToString & " belum tersedia", MsgBoxStyle.Information, "Maaf")
        End Try

        '// jika langsung akses ke nama form nya pakai dibawah ini
        'Try
        '    Dim objForm As Form
        '    Dim sValue As String
        '    Dim FullTypeName As String
        '    Dim FormInstanceType As Type
        '    sValue = sender.ToString.Replace(" ", "_")
        '    FullTypeName = Application.ProductName & "." & sValue
        '    FormInstanceType = Type.GetType(FullTypeName, True, True)
        '    objForm = CType(Activator.CreateInstance(FormInstanceType), Form)
        '    'MsgBox(objForm.AccessibleDescription.ToString)
        '    objForm.MdiParent = Me
        '    objForm.Text = sender.ToString
        '    objForm.Show()
        'Catch ex As Exception
        '    MsgBox("Menu " & sender.ToString & " belum tersedia", MsgBoxStyle.Information, "Maaf")
        'End Try
    End Sub

    'cek menuitem in accesslit
    Private Function cek_access_list(ByVal menuitem As String)
        Dim boolmnitem As Boolean = False
        For Each item In arrAccesmenu
            If item = menuitem Then
                boolmnitem = True
            End If
        Next
        Return boolmnitem
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.tsJam.Text = DateTime.Now.ToString("HH:mm")
    End Sub
End Class
