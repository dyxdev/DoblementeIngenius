Public Class AlojamientoRepository
    Implements IRepository

    Private _DB As DataAccess

    Public Sub New()

    End Sub
    Public Sub New(DB As DataAccess)
        Me._DB = DB
    End Sub

    Public Function GetAll() As DataTable Implements IRepository.GetAll

        Dim query As String = "SELECT * FROM ALOJAMIENTO"
        Return _DB.fetch(query)

    End Function
End Class
