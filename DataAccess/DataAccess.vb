Public Class DataAccess
    ' Connection string property
    Private Property _connectionString As String
    Private Property _connection As SqlConnection

    ' Open connection method
    Public Sub OpenConnection()
        ' Implement logic to open connection to Access database using OleDbConnection
    End Sub

    ' Close connection method
    Public Sub CloseConnection()
        ' Implement logic to close connection to Access database
    End Sub
End Class
