Imports System.Data.SqlClient
Imports System.Text
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports System.IO
Imports System.Drawing


Public Class helper
    Public conn As SqlConnection
    Public da As SqlDataAdapter
    Public cmd As SqlCommand
    Public ds As DataSet
    Public dr As SqlDataReader
    Public Shared pass As String = "?!@#$%^&*()_+|;:,’.-`~1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

    Public Sub New()
        Koneksi()
    End Sub

    Public Sub Koneksi()
        Try
            Dim strcn = "Data Source=.\SEHATI; Initial Catalog=Sehati; Integrated Security=True;"
            conn = New SqlConnection(strcn)
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

#Region "Login"
    'Method untuk mengisi nama user
    Public Sub LoginIsiUser(ByVal objek As ComboBox)
        conn.Open()
        Dim query = "Select UserID From [TUser]"
        cmd = New SqlCommand(query, conn)
        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TUser]")
        objek.DataSource = ds.Tables(0)
        objek.DisplayMember = "UserID"
        conn.Close()
    End Sub

    'MD5 Method
    Private Function md5Hash(ByVal SourceText As String) As String
        'Create an encoding object to ensure the encoding standard for the source text
        Dim Ue As New UnicodeEncoding()
        'Retrieve a byte array based on the source text
        Dim ByteSourceText() As Byte = Ue.GetBytes(SourceText)
        'Instantiate an MD5 Provider object
        Dim Md5 As New MD5CryptoServiceProvider()
        'Compute the hash value from the source
        Dim ByteHash() As Byte = Md5.ComputeHash(ByteSourceText)
        'And convert it to String format for return
        Return Convert.ToBase64String(ByteHash)
    End Function
    'SHA Method
    Private Function passwordEncryptSHA(ByVal password As String) As String
        Dim sha As New SHA1CryptoServiceProvider ' declare sha as a new SHA1CryptoServiceProvider
        Dim bytesToHash() As Byte ' and here is a byte variable

        bytesToHash = System.Text.Encoding.ASCII.GetBytes(password) ' covert the password into ASCII code

        bytesToHash = sha.ComputeHash(bytesToHash) ' this is where the magic starts and the encryption begins

        Dim encPassword As String = ""

        For Each b As Byte In bytesToHash
            encPassword += b.ToString("x2")
        Next

        Return encPassword ' boom there goes the encrypted password!

    End Function

    'Method untuk cek password
    Public Function LoginCekPassword(ByVal pass As String) As Integer
        conn.Open()
        Dim query = "Select UserID, UserPass From [TUser] Where UserPass=@UserPass"
        cmd = New SqlCommand(query, conn)
        cmd.Parameters.AddWithValue("UserPass", md5Hash(pass))
        Return cmd.ExecuteNonQuery()
        conn.Close()
    End Function

    Public Function getloginas(ByVal usr As String) As String
        Dim ret As String
        conn.Open()
        Dim q = String.Format("select [AccessName] from [TAccess] where [AccessCode]='{0}'", usr)
        cmd = New SqlCommand(q, conn)
        ret = cmd.ExecuteScalar()
        conn.Close()
        Return ret
    End Function

    Public Function getlogincode(ByVal usr As String) As String
        Dim ret As String
        conn.Open()
        Dim q = String.Format("select [AccessCode] from TUser where UserID='{0}'", usr)
        cmd = New SqlCommand(q, conn)
        ret = cmd.ExecuteScalar()
        conn.Close()
        Return ret
    End Function

    Public Function getaccesslist(ByVal code As String, ByVal usr As String)
        Dim str As String = ""
        Try
            conn.Open()
            Dim q = "select [AccessList] from [TAccessItem] where [AccessCode]=@user and [AccessMenu]=@kode"
            cmd = New SqlCommand(q, conn)
            cmd.Parameters.AddWithValue("@user", usr)
            cmd.Parameters.AddWithValue("@kode", code)
            str = cmd.ExecuteScalar
        Catch ex As SqlException
            MessageBox.Show(ex.Message)
        Finally
            conn.Close()
        End Try
       

        Return str
    End Function

    Public Function getmodulname(ByVal str As String)
        conn.Open()
        Dim strs As String
        Dim query = "select MenuNama from TMenu where MenuKode=@kode"
        cmd = New SqlCommand(query, conn)
        cmd.Parameters.AddWithValue("@kode", str)
        strs = cmd.ExecuteScalar
        conn.Close()
        Return strs
    End Function

#End Region


#Region "enkrip dekrip"
    Public Function cekKonfigurasi()
        Dim strHasil As Boolean = False
        If My.Settings.DatabaseName <> "" Or My.Settings.DataSource <> "" Or My.Settings.User <> "" Or My.Settings.Password <> "" Then
            strHasil = True
        End If
        Return strHasil
    End Function

    Public Shared Function Encrypt(ByVal input As String) As String
        Dim Teks_Asli As String, Teks_Sandi As String
        Dim PanjangTeks As Long
        Dim Pos As Long, EnkripsiKarakter, EnkripsiText

        Teks_Asli = " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890?!@#$%^&*()_+|;:,’.-`~"
        Teks_Sandi = "?!@#$%^&*()_+|;:,’.-`~1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ "

        For PanjangTeks = 1 To Len(input)
            Pos = InStr(Teks_Asli, Mid(input, PanjangTeks, 1))
            If Pos > 0 Then
                EnkripsiKarakter = Mid(Teks_Sandi, Pos, 1)
                EnkripsiText = EnkripsiText + EnkripsiKarakter
            Else
                EnkripsiText = EnkripsiText + Mid(input, PanjangTeks, 1)
            End If
        Next
        Return EnkripsiText
    End Function

    Public Shared Function Decrypt(ByVal input As String) As String
        Dim Teks_Asli As String, Teks_Sandi As String
        Dim PanjangTeks As Long
        Dim Pos As Long, DekripsiKarakter, DekripsiText

        Teks_Sandi = "?!@#$%^&*()_+|;:,’.-`~1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ "
        Teks_Asli = " ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890?!@#$%^&*()_+|;:,’.-`~"

        For PanjangTeks = 1 To Len(input)
            Pos = InStr(Teks_Sandi, Mid(input, PanjangTeks, 1))
            If Pos > 0 Then
                DekripsiKarakter = Mid(Teks_Asli, Pos, 1)
                DekripsiText = DekripsiText + DekripsiKarakter
            Else
                DekripsiText = DekripsiText + Mid(input, PanjangTeks, 1)
            End If
        Next
        Return DekripsiText
    End Function

    Public Function AES_Encrypt(ByVal input As String) As String
        Dim AES As New System.Security.Cryptography.RijndaelManaged
        Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim encrypted As String = ""
        Try
            Dim hash(31) As Byte
            Dim temp As Byte() = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(pass))
            Array.Copy(temp, 0, hash, 0, 16)
            Array.Copy(temp, 0, hash, 15, 16)
            AES.Key = hash
            AES.Mode = CipherMode.ECB
            Dim DESEncrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateEncryptor
            Dim Buffer As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(input)
            encrypted = Convert.ToBase64String(DESEncrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Return encrypted
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function AES_Decrypt(ByVal input As String) As String
        Dim AES As New System.Security.Cryptography.RijndaelManaged
        Dim Hash_AES As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim decrypted As String = ""
        Try
            Dim hash(31) As Byte
            Dim temp As Byte() = Hash_AES.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(pass))
            Array.Copy(temp, 0, hash, 0, 16)
            Array.Copy(temp, 0, hash, 15, 16)
            AES.Key = hash
            AES.Mode = CipherMode.ECB
            Dim DESDecrypter As System.Security.Cryptography.ICryptoTransform = AES.CreateDecryptor
            Dim Buffer As Byte() = Convert.FromBase64String(input)
            decrypted = System.Text.ASCIIEncoding.ASCII.GetString(DESDecrypter.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Return decrypted
        Catch ex As Exception
            Return Nothing
        End Try

    End Function
#End Region

    'Method untuk isi data grid berdasakan filter dari dua parameter dan kata kunci
    Public Sub IsiDataGridFilter(ByVal cari As String, ByVal cari2 As String, ByVal parameter As String, ByVal parameter2 As String, ByVal objek As DataGridView, ByVal query As String, ByVal dataset As String)
        If cari = "" And cari2 = "" Then
        ElseIf cari = "" Then
            query = String.Format(query & " Where {0} Like '%'+@katakunci2+'%'", parameter2)
        ElseIf cari2 = "" Then
            query = String.Format(query & " Where {0} Like '%'+@katakunci+'%'", parameter)
        ElseIf Not (cari = "" And cari2 = "") Then
            query = String.Format(query & " Where {0} Like '%'+@katakunci+'%' AND {1} Like '%'+@katakunci2+'%'", parameter, parameter2)
        End If
        conn.Open()
        cmd = New SqlCommand(query, conn)
        cmd.Parameters.AddWithValue("@katakunci", cari)
        cmd.Parameters.AddWithValue("@katakunci2", cari2)
        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, dataset)
        objek.DataSource = ds.Tables(dataset)

        conn.Close()
    End Sub
    Public Function hitungUmur(ByVal objek As DataGridView) As String()
        Dim umur(2) As String

        Dim tahun As String = ""
        Dim bulan As String = ""
        Dim hari As String = ""
        If Not (objek.CurrentRow.Cells(8).Value Is Nothing) Then
            Dim tgllahir As String = objek.CurrentRow.Cells(8).Value
            If tgllahir.Substring(0, 4) = "0000" Then
                tahun = Now.Year - tgllahir.Substring(4, 4)
                bulan = "0"
                hari = "0"
            Else
                Dim dt = DateTime.ParseExact(tgllahir, "ddMMyyyy", Nothing)
                Dim dt2 As Date = Now.ToShortDateString
                Dim dt3 As TimeSpan = (dt2 - dt)
                Dim diff As Double = dt3.Days
                tahun = Str(Int(diff / 365))
                diff = diff Mod 365
                bulan = Str(Int(diff / 30))
                diff = diff Mod 30
                hari = Str(diff)
            End If
        End If

        umur(0) = tahun
        umur(1) = bulan
        umur(2) = hari
        Return umur
    End Function
    Public Function BuatNoRegis()
        'Get Date
        Dim tanggal As String = Now().Year().ToString().Substring(2, 2) & "0" & Now().Month().ToString() & "0" & Now().Day.ToString()

        'Get Urutan Sebelumnya
        Dim noUrut As Integer = 0
        Dim query As String = "Select Top 1 JalanNoReg From [TRawatJalan] Where JalanTanggal = GETDATE() Order By JalanNoReg desc"
        conn.Open()
        cmd = New SqlCommand(query, conn)
        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        Dim hasil = cmd.ExecuteScalar()
        If hasil = "" Then
            noUrut = 1
        Else

            noUrut = CInt(hasil.ToString.Substring(13, 1)) + 1
        End If

        conn.Close()

        'Susun Nomer Reg
        Dim noreg As String = ""
        noreg &= "RT-" & tanggal & "-" & noUrut.ToString.PadLeft(4, "0")

        Return noreg
    End Function
    Public Sub IsiUnit(ByVal objek As ComboBox)
        Dim query As String = "Select UnitKode, UnitNama From [TUnit] where UnitGrup ='KUM' OR UnitGrup = 'KSP' UNION Select '' as UnitKode, 'ALL' as 'UnitNama' Order By UnitNama asc "
        conn.Open()
        cmd = New SqlCommand(query, conn)


        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TUnit]")

        objek.DataSource = ds.Tables("[TUnit]")
        objek.ValueMember = "UnitKode"
        objek.DisplayMember = "UnitNama"

        conn.Close()
    End Sub
    Public Sub IsiDokter(ByVal objek As ComboBox, ByVal pelaku As String)
        Dim query As String = "Select PelakuKode, PelakuNama From [TPelaku] where UnitKode=@unit Order By PelakuNama Asc "

        conn.Open()
        cmd = New SqlCommand(query, conn)
        cmd.Parameters.AddWithValue("@unit", pelaku)

        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TPelaku]")

        objek.DataSource = ds.Tables("[TPelaku]")
        objek.ValueMember = "PelakuKode"
        objek.DisplayMember = "PelakuNama"

        conn.Close()
    End Sub
    Public Sub IsiPenjamin(ByVal objek As ComboBox)
        Dim query As String = "Select TOP 3 VarKode, VarNama From [TAdmVar] where VarSeri='JENISPAS' Order By VarKode asc"
        conn.Open()
        cmd = New SqlCommand(query, conn)


        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TAdmVar]")

        objek.DataSource = ds.Tables("[TAdmVar]")
        objek.ValueMember = "VarKode"
        objek.DisplayMember = "VarNama"

        conn.Close()

    End Sub
    Public Sub IsiPenjamin2(ByVal objek As ComboBox, ByVal penjamin As String)
        Dim query As String = "Select PrshKode, PrshNama From [TPerusahaan] where Substring(PrshKode,1,1)=@penjamin1 Order By PrshKode Asc "
        conn.Open()
        cmd = New SqlCommand(query, conn)
        cmd.Parameters.AddWithValue("@penjamin1", penjamin)
        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TPerusahaan]")

        objek.DataSource = ds.Tables("[TPerusahaan]")
        objek.ValueMember = "PrshKode"
        objek.DisplayMember = "PrshNama"

        conn.Close()
    End Sub


#Region "PasienBaru"
    Public Sub buatNomorRM(ByVal txtObjek As TextBox)
        Dim query As String = "Select Top 1 * From [TPasien] order by PasienNomorRM Desc"
        conn.Open()
        cmd = New SqlCommand(query, conn)
        da = New SqlDataAdapter
        ds = New DataSet
        da.SelectCommand = cmd
        da.Fill(ds, "[TPasien]")
        txtObjek.Text = (CInt(ds.Tables("[TPasien]").Rows(0).Item("PasienNomorRM")) + 1).ToString()



        conn.Close()
    End Sub
    Public Sub isiJenisPasien(ByVal objek As ComboBox)
        Dim query As String = "Select VarKode +' - '+ VarNama as 'KodeNama', VarKode, VarNama From [TAdmVar] where VarSeri ='JENISPAS'"
        conn.Open()
        cmd = New SqlCommand(query, conn)


        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TAdmVar]")

        objek.DataSource = ds.Tables("[TAdmVar]")
        objek.ValueMember = "VarKode"
        objek.DisplayMember = "KodeNama"

        conn.Close()
    End Sub
    Public Sub isiStatusKawin(ByVal objek As ComboBox)
        Dim query As String = "Select VarKode +' - '+ VarNama as 'KodeNama', VarKode, VarNama From [TAdmVar] where VarSeri ='KAWIN'"
        conn.Open()
        cmd = New SqlCommand(query, conn)


        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TAdmVar1]")

        objek.DataSource = ds.Tables("[TAdmVar1]")
        objek.ValueMember = "VarKode"
        objek.DisplayMember = "KodeNama"

        conn.Close()
    End Sub
    Public Sub isiAgama(ByVal objek As ComboBox)
        Dim query As String = "Select VarKode +' - '+ VarNama as 'KodeNama', VarKode, VarNama From [TAdmVar] where VarSeri ='AGAMA'"
        conn.Open()
        cmd = New SqlCommand(query, conn)


        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TAdmVar2]")

        objek.DataSource = ds.Tables("[TAdmVar2]")
        objek.ValueMember = "VarKode"
        objek.DisplayMember = "KodeNama"

        conn.Close()
    End Sub
    Public Sub isiDarah(ByVal objek As ComboBox)
        Dim query As String = "Select VarKode, VarNama From [TAdmVar] where VarSeri ='DARAH'"
        conn.Open()
        cmd = New SqlCommand(query, conn)


        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TAdmVar3]")

        objek.DataSource = ds.Tables("[TAdmVar3]")
        objek.ValueMember = "VarKode"
        objek.DisplayMember = "VarNama"

        conn.Close()
    End Sub
    Public Sub isiGender(ByVal objek As ComboBox)
        Dim query As String = "Select VarKode +' - '+ VarNama as 'KodeNama', VarKode, VarNama From [TAdmVar] where VarSeri ='GENDER'"
        conn.Open()
        cmd = New SqlCommand(query, conn)


        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TAdmVar4]")

        objek.DataSource = ds.Tables("[TAdmVar4]")
        objek.ValueMember = "VarKode"
        objek.DisplayMember = "KodeNama"

        conn.Close()
    End Sub
    Public Sub isiPendidikan(ByVal objek As ComboBox)
        Dim query As String = "Select VarKode +' - '+ VarNama as 'KodeNama', VarKode, VarNama From [TAdmVar] where VarSeri ='PENDIDIKAN'"
        conn.Open()
        cmd = New SqlCommand(query, conn)


        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TAdmVar5]")

        objek.DataSource = ds.Tables("[TAdmVar5]")
        objek.ValueMember = "VarKode"
        objek.DisplayMember = "KodeNama"

        conn.Close()
    End Sub
    Public Sub isiPekerjaan(ByVal objek As ComboBox)
        Dim query As String = "Select VarKode +' - '+ VarNama as 'KodeNama', VarKode, VarNama From [TAdmVar] where VarSeri ='PEKERJAAN'"
        conn.Open()
        cmd = New SqlCommand(query, conn)


        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TAdmVar6]")

        objek.DataSource = ds.Tables("[TAdmVar6]")
        objek.ValueMember = "VarKode"
        objek.DisplayMember = "KodeNama"

        conn.Close()
    End Sub
    Public Sub isiHubKeluarga(ByVal objek As ComboBox)
        Dim query As String = "Select VarKode +' - '+ VarNama as 'KodeNama', VarKode, VarNama From [TAdmVar] where VarSeri ='Keluarga'"
        conn.Open()
        cmd = New SqlCommand(query, conn)


        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TAdmVar7]")

        objek.DataSource = ds.Tables("[TAdmVar7]")
        objek.ValueMember = "VarKode"
        objek.DisplayMember = "KodeNama"

        conn.Close()
    End Sub
    Public Sub isiProvinsi(ByVal objek As ComboBox)
        Dim query As String = "Select WilKode, WilJenis , WilNama From [TWilayah] where WilJenis =1"
        conn.Open()
        Try
            cmd = New SqlCommand(query, conn)


            ds = New DataSet
            da = New SqlDataAdapter
            da.SelectCommand = cmd
            da.Fill(ds, "[TWilayah]")

            objek.DataSource = ds.Tables("[TWilayah]")
            objek.ValueMember = "WilKode"
            objek.DisplayMember = "WilNama"
            objek.Text = "Yogyakarta"
        Catch ex As Exception
        Finally
            conn.Close()
        End Try


    End Sub
    Public Sub isiKota(ByVal objek1 As ComboBox, ByVal objek2 As ComboBox)
        Dim query As String = "Select WilKode, WilJenis , WilNama From [TWilayah] where WilJenis =2 And Substring(WilKode,1,2)=@Prov"
        conn.Open()
        Try
            cmd = New SqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@Prov", objek2.SelectedValue.ToString.Substring(0, 2))

            ds = New DataSet
            da = New SqlDataAdapter
            da.SelectCommand = cmd
            da.Fill(ds, "[TWilayah]")

            objek1.DataSource = ds.Tables("[TWilayah]")
            objek1.ValueMember = "WilKode"
            objek1.DisplayMember = "WilNama"

        Catch ex As Exception
        Finally
            conn.Close()
        End Try


    End Sub
    Public Sub isiKecamatan(ByVal objek1 As ComboBox, ByVal objek2 As ComboBox)
        Dim query As String = "Select WilKode, WilJenis , WilNama From [TWilayah] where WilJenis =3 And Substring(WilKode,3,2)=@Kota"
        'If cmd.Connection.State = ConnectionState.Open Then
        '    cmd.Connection.Close()
        'End If

        conn.Open()
        Try
            cmd = New SqlCommand(query, conn)
            cmd.Parameters.AddWithValue("@Kota", objek2.SelectedValue.ToString.Substring(2, 2))


            ds = New DataSet
            da = New SqlDataAdapter
            da.SelectCommand = cmd
            da.Fill(ds, "[TWilayah]")

            objek1.DataSource = ds.Tables("[TWilayah]")
            objek1.ValueMember = "WilKode"
            objek1.DisplayMember = "WilNama"

        Catch ex As Exception
        Finally
            conn.Close()
        End Try


    End Sub
#End Region


    Public Sub InsertPasienBaru(ByVal txtPasienNomorRM As TextBox, ByVal txtNamaLengkap As TextBox, ByVal cmbJenisKelamin As ComboBox, ByVal txtAlamat As TextBox, ByVal cmbKelurahan As ComboBox,
                              ByVal cmbKecamatan As ComboBox, ByVal cmbKota As ComboBox, ByVal cmbProvinsi As ComboBox, ByVal txtRT As TextBox, ByVal txtRW As TextBox, ByVal txtTelepon As TextBox, ByVal txtHP As TextBox,
                              ByVal txtTempatLahir As TextBox, ByVal tgl As String, ByVal cmbJenisPekerjaan As ComboBox, ByVal txtNamaPekerjaan As TextBox, ByVal txtAlamatKerja As TextBox, ByVal cmbGoldar As ComboBox, ByVal cmbAgama As ComboBox, ByVal cmbPendidikan As ComboBox,
                              ByVal cmbStatusKawin As ComboBox, ByVal txtNamaKeluarga As TextBox, ByVal cmbHubKeluarga As ComboBox, ByVal txtKtp As TextBox, ByVal cmbJenisPas As ComboBox, ByVal txtNamaPanggilan As TextBox,
                              ByVal txtTeleponKeluarga As TextBox, ByVal txtAlamatKeluarga As TextBox)

        Dim query = "Insert into [TPasien] (PasienNomorRM,PasienNama,PasienGender,PasienAlamat,PasienKelurahan,PasienKecamatan,PasienKota,PasienProv,PasienRT,PasienRW,PasienTelp,PasienHP,PasienTmpLahir,PasienTglLahir,PasienKerjaKode,PasienKerja,PasienKerjaAlamat,PasienGolDarah,PasienAgama,PasienPdk,PasienStatusKw,PasienKlgNama,PasienKlgHub,PasienNOID,PasienJenis,PasienTglInput,UserID,PasienPanggilan,PasienKlgTelp,PasienKlgAlamat) Values (@PasienNomorRM,@PasienNama,@PasienGender,@PasienAlamat,@PasienKelurahan,@PasienKecamatan,@PasienKota,@PasienProv,@PasienRT,@PasienRW,@PasienTelp,@PasienHP,@PasienTmpLahir,@PasienTglLahir,@PasienKerjaKode,@PasienKerja,@PasienKerjaAlamat,@PasienGolDarah,@PasienAgama,@PasienPdk,@PasienStatusKw,@PasienKlgNama,@PasienKlgHub,@PasienNOID,@PasienJenis,@PasienTglInput,@UserID,@PasienPanggilan,@PasienKlgTelp,@PasienKlgAlamat)"

        conn.Open()
        cmd = New SqlCommand(query, conn)
        cmd.Parameters.AddWithValue("@PasienNomorRM", txtPasienNomorRM.Text)
        cmd.Parameters.AddWithValue("@PasienNama", txtNamaLengkap.Text)
        cmd.Parameters.AddWithValue("@PasienGender", cmbJenisKelamin.SelectedValue)
        cmd.Parameters.AddWithValue("@PasienAlamat", txtAlamat.Text)
        cmd.Parameters.AddWithValue("@PasienKelurahan", cmbKelurahan.Text)
        cmd.Parameters.AddWithValue("@PasienKecamatan", cmbKecamatan.Text)
        cmd.Parameters.AddWithValue("@PasienKota", cmbKota.Text)
        cmd.Parameters.AddWithValue("@PasienProv", cmbProvinsi.Text)
        cmd.Parameters.AddWithValue("@PasienRT", txtRT.Text)
        cmd.Parameters.AddWithValue("@PasienRW", txtRW.Text)
        cmd.Parameters.AddWithValue("@PasienTelp", txtTelepon.Text)
        cmd.Parameters.AddWithValue("@PasienHP", txtHP.Text)
        cmd.Parameters.AddWithValue("@PasienTmpLahir", txtTempatLahir.Text)
        cmd.Parameters.AddWithValue("@PasienTglLahir", tgl)
        cmd.Parameters.AddWithValue("@PasienKerjaKode", cmbJenisPekerjaan.SelectedValue)
        cmd.Parameters.AddWithValue("@PasienKerja", txtNamaPekerjaan.Text)
        cmd.Parameters.AddWithValue("@PasienKerjaAlamat", txtAlamatKerja.Text)
        cmd.Parameters.AddWithValue("@PasienGolDarah", cmbGoldar.SelectedValue)
        cmd.Parameters.AddWithValue("@PasienAgama", cmbAgama.SelectedValue)
        cmd.Parameters.AddWithValue("@PasienPdk", cmbPendidikan.SelectedValue)
        cmd.Parameters.AddWithValue("@PasienStatusKw", cmbStatusKawin.SelectedValue)
        cmd.Parameters.AddWithValue("@PasienKlgNama", txtNamaKeluarga.Text)
        cmd.Parameters.AddWithValue("@PasienKlgHub", cmbHubKeluarga.SelectedValue)
        cmd.Parameters.AddWithValue("@PasienNOID", txtKtp.Text)
        cmd.Parameters.AddWithValue("@PasienJenis", cmbJenisPas.SelectedValue)
        cmd.Parameters.AddWithValue("@PasienTglInput", Now().ToString)
        cmd.Parameters.AddWithValue("@UserID", "Obed")
        cmd.Parameters.AddWithValue("@PasienPanggilan", txtNamaPanggilan.Text)
        cmd.Parameters.AddWithValue("@PasienKlgTelp", txtTeleponKeluarga.Text)
        cmd.Parameters.AddWithValue("@PasienKlgAlamat", txtAlamatKeluarga.Text)

        If cmd.ExecuteNonQuery = 1 Then
            MessageBox.Show("Insert Berhasil")
        Else
            MessageBox.Show("Insert Gagal")
        End If
        conn.Close()

    End Sub


#Region "Wewenang"
    Public Sub isiCPanel(ByVal txtNama As TextBox, ByVal txtNamaLengkap As TextBox, ByVal txtAlamatLengkap As TextBox, ByVal txtAlamatPendek As TextBox, ByVal pbLogoBesar1 As PictureBox, ByVal pbLogoBesar2 As PictureBox, ByVal pbLogoKecil1 As PictureBox, ByVal pbLogoKecil2 As PictureBox)
        Dim query = "Select * from [TCpanel] "
        conn.Open()
        cmd = New SqlCommand(query, conn)
        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[TCpanel]")
        txtNama.Text = ds.Tables("[TCpanel]").Rows(0).Item("KodeRS")
        txtNamaLengkap.Text = ds.Tables("[TCpanel]").Rows(0).Item("NamaRS")
        txtAlamatLengkap.Text = ds.Tables("[TCpanel]").Rows(0).Item("AlamatLengkap")
        txtAlamatPendek.Text = ds.Tables("[TCpanel]").Rows(0).Item("AlamatPendek")
        Dim image As Byte() = DirectCast(ds.Tables("[TCpanel]").Rows(0).Item("LogoBesarWarna"), Byte())
        Dim ms1 As New MemoryStream(image)
        pbLogoBesar1.Image = Bitmap.FromStream(ms1)
        Dim image2 As Byte() = DirectCast(ds.Tables("[TCpanel]").Rows(0).Item("LogoBesarBW"), Byte())
        Dim ms2 As New MemoryStream(image2)
        pbLogoBesar2.Image = Bitmap.FromStream(ms2)
        Dim image3 As Byte() = DirectCast(ds.Tables("[TCpanel]").Rows(0).Item("LogoKecilWarna"), Byte())
        Dim ms3 As New MemoryStream(image3)
        pbLogoKecil1.Image = Bitmap.FromStream(ms3)
        Dim image4 As Byte() = DirectCast(ds.Tables("[TCpanel]").Rows(0).Item("LogoKecilBW"), Byte())
        Dim ms4 As New MemoryStream(image4)
        pbLogoKecil2.Image = Bitmap.FromStream(ms4)
        conn.Close()
    End Sub

    Public Function UpdateCPanel(ByVal txtNama As TextBox, ByVal txtNamaLengkap As TextBox, ByVal txtAlamatLengkap As TextBox, ByVal txtAlamatPendek As TextBox, ByVal arr As Byte(), ByVal arr2 As Byte(), ByVal arr3 As Byte(), ByVal arr4 As Byte())
        Dim hasil As Integer
        Dim query = "Update [TCpanel] set KodeRS=@kodeRS, NamaRS=@namaRS, AlamatLengkap=@AlamatLengkap, AlamatPendek=@AlamatPendek, LogoBesarWarna=@logobesarwarna, LogoBesarBW=@logobesarbw, LogoKecilWarna=@logokecilwarna, LogoKecilBW=@logokecilbw  where ID=01"
        conn.Open()
        cmd = New SqlCommand(query, conn)
        cmd.Parameters.AddWithValue("@kodeRS", txtNama.Text)
        cmd.Parameters.AddWithValue("@namaRS", txtNamaLengkap.Text)
        cmd.Parameters.AddWithValue("@AlamatLengkap", txtAlamatLengkap.Text)
        cmd.Parameters.AddWithValue("@AlamatPendek", txtAlamatPendek.Text)
        cmd.Parameters.AddWithValue("@LogoBesarWarna", arr)
        cmd.Parameters.AddWithValue("@LogoBesarBW", arr2)
        cmd.Parameters.AddWithValue("@LogoKecilWarna", arr3)
        cmd.Parameters.AddWithValue("@LogoKecilBW", arr4)


        hasil = cmd.ExecuteNonQuery()
        conn.Close()
        Return hasil
    End Function

    Public Function isiPengguna()
        Dim query = "Select AccessName from [TAccess] "
        Dim output As New List(Of String)
        conn.Open()
        cmd = New SqlCommand(query, conn)
        dr = cmd.ExecuteReader()
        While dr.Read()
            output.Add(dr.Item("AccessName"))
        End While
        conn.Close()
        Return output
    End Function

    Public Function isiUserName(ByVal access As String)

        Dim query2 = "Select UserName from [SEHATI].dbo.TUser Where AccessCode=(Select top 1 AccessCode from Sehati.dbo.TAccess Where AccessName=@Access)"
        Dim user As New List(Of String)
        conn.Open()
        cmd = New SqlCommand(query2, conn)
        cmd.Parameters.AddWithValue("@Access", access)
        dr = cmd.ExecuteReader()
        While dr.Read()
            user.Add(dr.Item("UserName"))
        End While
        conn.Close()
        Return user
    End Function

    Public Sub loadMenuHItem(ByVal usercode As String, ByRef menuH As MenuStrip)
        Dim Hmenu As New MenuStrip
        Dim qlistmenu = "select MenuNama from TMenuItem where MenuKode=@usercode and MenuJenis='H'"
        conn.Open()
        cmd = New SqlCommand(qlistmenu, conn)
        cmd.Parameters.AddWithValue("@usercode", usercode)
        dr = cmd.ExecuteReader
        If dr.HasRows Then
            While dr.Read
                Hmenu.Items.Add(dr(0).ToString, Nothing, New System.EventHandler(AddressOf MainMenu_OnClick))
                menuH.Controls.Add(Hmenu)
            End While
        End If
    End Sub

    Public Sub MainMenu_OnClick(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim cms As New ContextMenuStrip()
        'Dim sMenu() As String
        'Dim sMenuRD As SqlClient.SqlDataReader
        'Dim sQry As String = ""
        'sQry = "Select MenuID from MenuMaster Where MenuText = '" & sender.ToString & "'"
        'sMenuRD = Conn.ReaderData(sQry)
        'Dim parentMenuID As Integer
        'If sMenuRD.HasRows Then
        '    sMenuRD.Read()
        '    parentMenuID = sMenuRD("MenuID")
        'End If
        'sMenuRD.Close()
        'sQry = ""
        'sQry = "Select MenuText from MenuMaster Where MainMenuID ="
        '" & parentMenuID & "'" & _
        '       " And isActive = 1" & _
        '       " And MenuID in (
        '           Select MenuID from Access Where AccessId =
        '           " & CInt(iUserAccessMode) & ")" & _
        '       " Order BY MenuOrder"
        'sMenuRD = Conn.ReaderData(sQry)
        'ReDim Preserve sMenu(0)
        'Dim i As Integer
        'If sMenuRD.HasRows Then
        '    ReDim Preserve sMenu(0)
        '    i = 0
        '    While sMenuRD.Read()
        '        ReDim Preserve sMenu(i)
        '        sMenu(i) = sMenuRD("MenuText")
        '        i = i + 1
        '        End While
        '    End If
        'sMenuRD.Close()
        'For Each sMn As String In sMenu
        '    cms.Items.Add(sMn, Nothing,
        '        New System.EventHandler(AddressOf SelectedChildMenu_OnClick))
        'Next
        'Dim tsi As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        'tsi.DropDown = cms
    End Sub
#End Region

    Public Sub IsiVRawatJalan(ByVal objek As DataGridView)
        Dim query As String = "Select * From VRawatJalan"
        conn.Open()
        cmd = New SqlCommand(query, conn)


        ds = New DataSet
        da = New SqlDataAdapter
        da.SelectCommand = cmd
        da.Fill(ds, "[VRawatJalan]")

        objek.DataSource = ds.Tables("[VRawatJalan]")


        conn.Close()
    End Sub

End Class

