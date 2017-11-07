Imports System.Data.SqlClient
Imports System.Text
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports System.IO
Imports System.Drawing

Public Class Connectio
    Implements IDisposable

    Private _SQL As String
    Private _Connection As SqlConnection
    Private _Command As SqlCommand
    Private _DataAdapter As SqlDataAdapter
    Private _DataReader As SqlDataReader
    Private strServer As String = ""
    Private strDatabase As String = ""


    Public Property Connection() As SqlConnection
        Get
            Return _Connection
        End Get
        Set(ByVal value As SqlConnection)
            _Connection = value
        End Set
    End Property

    Public Property Command() As SqlCommand
        Get
            Return _Command
        End Get
        Set(ByVal value As SqlCommand)
            _Command = value
        End Set
    End Property

    Public Property DataAdapter() As SqlDataAdapter
        Get
            Return _DataAdapter
        End Get
        Set(ByVal value As SqlDataAdapter)
            _DataAdapter = value
        End Set
    End Property

    Public Property DataReader() As SqlDataReader
        Get
            Return _DataReader
        End Get
        Set(ByVal value As SqlDataReader)
            _DataReader = value
        End Set
    End Property

    Public Property SQL() As String
        Get
            Return _SQL
        End Get
        Set(ByVal value As String)
            _SQL = value
        End Set
    End Property

    Public Sub New()
        Dim cnstring As String = "Data Source=.\SEHATI;Initial Catalog=Sehati;Integrated Security=True"
        Connection = New SqlConnection(cnstring)
    End Sub

    Public Sub OpenConnection()
        Try
            Connection.Open()
        Catch sqlException As sqlException
            Throw New System.Exception(sqlException.Message, sqlException.InnerException)
        Catch InvalidOperationExceptionErr As InvalidOperationException
            Throw New System.Exception(InvalidOperationExceptionErr.Message, InvalidOperationExceptionErr.InnerException)
        End Try
    End Sub

    Public Sub CloseConnection()
        Try
            Connection.Close()
        Catch ex As SqlException
        End Try
    End Sub

    Public Sub InitializeCommand()
        If Command Is Nothing Then
            Try
                Command = New sqlCommand(SQL, Connection)
                'cek apakah tipe commandnya store procedure atau bukan
                If Not SQL.ToUpper.StartsWith("SELECT") _
                    And Not SQL.ToUpper.StartsWith("UPDATE") _
                    And Not SQL.ToUpper.StartsWith("DELETE") _
                    And Not SQL.ToUpper.StartsWith("INSERT") Then
                    Command.CommandType = CommandType.StoredProcedure
                End If
            Catch OldeDbExceptionErr As sqlException
                Throw New System.Exception(OldeDbExceptionErr.Message, OldeDbExceptionErr.InnerException)
            End Try
        End If
    End Sub

    Public Sub AddParameter(ByVal strName As String, ByVal sqlType As SqlDbType, _
    ByVal intSize As Integer, ByVal objValue As Object)
        Try
            Command.Parameters.Add(strName, sqlType, intSize).Value = objValue
        Catch OldeDbExceptionErr As sqlException
            Throw New System.Exception(OldeDbExceptionErr.Message, OldeDbExceptionErr.InnerException)
        End Try
    End Sub

    Public Sub AddParameterWithValue(ByVal strName As String, ByVal objValue As Object)
        Try
            Command.Parameters.AddWithValue(strName, objValue)
        Catch OldeDbExceptionErr As sqlException
            Throw New System.Exception(OldeDbExceptionErr.Message, OldeDbExceptionErr.InnerException)
        End Try
    End Sub

    Public Sub InitializeDataAdapter()
        Try
            DataAdapter = New sqlDataAdapter
            DataAdapter.SelectCommand = Command
        Catch OldeDbExceptionErr As sqlException
            Throw New System.Exception(OldeDbExceptionErr.Message, OldeDbExceptionErr.InnerException)
        End Try
    End Sub

    Public Sub FillDataSet(ByRef dsDataSet As DataSet, ByVal strTableName As String)
        Try
            InitializeCommand()
            InitializeDataAdapter()
            DataAdapter.Fill(dsDataSet, strTableName)
        Catch OldeDbExceptionErr As sqlException
            Throw New System.Exception(OldeDbExceptionErr.Message, OldeDbExceptionErr.InnerException)
        Finally
            Command.Dispose()
            Command = Nothing
            DataAdapter.Dispose()
            DataAdapter = Nothing
        End Try
    End Sub

    Public Sub FillDataTable(ByRef dtDataTable As DataTable)
        Try
            InitializeCommand()
            InitializeDataAdapter()
            DataAdapter.Fill(dtDataTable)
        Catch OldeDbExceptionErr As sqlException
            Throw New System.Exception(OldeDbExceptionErr.Message, OldeDbExceptionErr.InnerException)
        Finally
            Command.Dispose()
            Command = Nothing
            DataAdapter.Dispose()
            DataAdapter = Nothing
        End Try
    End Sub

    Public Function ExecuteStoreProcedure() As Integer
        Try
            OpenConnection()
            ExecuteStoreProcedure = Command.ExecuteNonQuery
        Catch ex As Exception
            Throw New System.Exception(ex.Message, ex.InnerException)
        Finally
            CloseConnection()
        End Try
    End Function

    Private disposedValue As Boolean = False        ' To detect redundant calls
    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free managed resources when explicitly called
                If Not DataReader Is Nothing Then
                    DataReader.Close()
                    DataReader = Nothing
                End If
                If Not DataAdapter Is Nothing Then
                    DataAdapter.Dispose()
                    DataAdapter = Nothing
                End If
                If Not Command Is Nothing Then
                    Command.Dispose()
                    Command = Nothing
                End If
                If Not Connection Is Nothing Then
                    Connection.Close()
                    Connection = Nothing
                End If
            End If
            ' TODO: free shared unmanaged resources
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class