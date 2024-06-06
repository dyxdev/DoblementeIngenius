Imports System.Data.SqlClient


Public Class DataAccess
    Private _connectionString As String
    Private _connection As SqlConnection

    Public Property UconnectionString As String

        Get

            Return _connectionString


        End Get

        Set(value As String)

            If IsValidConnectionPipeline(value) And IsValidConnectionString(value) Then Me._connectionString = value
            Me.CheckConnection()

        End Set

    End Property


    ' Open connection method
    Public Function GetConnection()
        If _connection IsNot Nothing Then
            Dim connection As New SqlConnection(_connectionString)
            _connection = connection
        End If

        Return _connection

    End Function

    ' Close connection method
    Public Sub CloseConnection()
        If _connection IsNot Nothing Then
            _connection.Close()
            _connection.Dispose()
        End If
    End Sub



    Public Function IsValidConnectionPipeline(connectionString As String) As Boolean
        Dim hasDataSource As Boolean = InStr(connectionString, "Data Source=") > 0
        Dim hasInitialCatalog As Boolean = InStr(connectionString, "Initial Catalog=") > 0

        ' Contain source and table
        Return hasDataSource And hasInitialCatalog
    End Function

    Public Function IsValidConnectionString(connectionString As String) As Boolean
        Dim builder As New SqlConnectionStringBuilder(connectionString)
        Try
            builder.ConnectionString = connectionString ' for avoid sql injection
            Return True
        Catch ex As Exception
            Throw New ArgumentException("Invalid connection strings")
        End Try
    End Function

    Public Sub CheckConnection()

        Dim connection As New SqlConnection(Me._connectionString)

        Try
            connection.Open()
        Catch ex As SqlException
            ' Handle connection string errors
            If ex.Message.Contains("Login failed for user") Then  ' Example check for login error
                Throw New Exception("Incorrect username or password!")
            End If
            Throw New Exception("Connection string error: " & ex.Message)
        Finally
            If connection.State = ConnectionState.Open Then
                connection.Close()
            End If
        End Try
    End Sub

    Public Function fetch(query As String) As DataTable
        Dim dataTable As New DataTable()
        Try
            Using _connection As New SqlConnection(_connectionString)
                _connection.Open()
                Using command As New SqlCommand(query, _connection)
                    Dim adapter As New SqlDataAdapter(command)
                    adapter.Fill(dataTable)
                End Using
            End Using
        Catch ex As Exception
            dataTable = Nothing
        End Try

        Return dataTable
    End Function

    Public Function ExistTable(tableName As String) As Boolean
        Using _connection
            _connection.Open()

            Dim tableExists As Boolean = False

            Dim schemaRestrictions() As String = {Nothing, Nothing, tableName}
            Dim table As DataTable = _connection.GetSchema("Tables", schemaRestrictions)

            If table.Rows.Count > 0 Then
                tableExists = True
            End If

            _connection.Close()

            Return tableExists
        End Using
    End Function
End Class
