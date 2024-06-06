Imports ClassLibrary1

Public Class AlojamientoService
    Implements IAlojamientoService

    Private ReadOnly _alojamientoRepository As AlojamientoRepository

    Public Sub New(alojamientoRepository As IRepository)
        Me._alojamientoRepository = alojamientoRepository
    End Sub

    Public Function getAllAlojamientos() As DataTable Implements IAlojamientoService.getAllAlojamientos
        Return _alojamientoRepository.GetAll()
    End Function
End Class
