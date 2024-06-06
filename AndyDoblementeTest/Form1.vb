Imports Autofac
Imports Autofac.Core
Imports ClassLibrary1
Imports Services

Public Class Form1

    Private _container As Container
    Private _configuration As IConfigurationDB
    Private _alojamientoService As IAlojamientoService
    Private _dataAccess As DataAccess


    Private Sub InitializeContainer()

        Dim builder As New ContainerBuilder()

        builder.RegisterType(Of DataAccess).As(Of DataAccess)()
        builder.RegisterType(Of AlojamientoRepository).As(Of IRepository)()
        builder.RegisterType(Of AlojamientoService).As(Of IAlojamientoService)()
        builder.RegisterType(Of ConfigurationDB).As(Of IConfigurationDB)()

        _container = builder.Build()

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim server As String = TextBox1.Text
        Dim dbName As String = TextBox2.Text
        Dim user As String = TextBox3.Text
        Dim pass As String = TextBox4.Text

        Try
            Dim stringConnection As String = _configuration.connect(server, user, pass, dbName)
            Console.WriteLine(stringConnection)
            Console.WriteLine("Connected")
            _dataAccess.UconnectionString = stringConnection
            Dim table As DataTable = _alojamientoService.getAllAlojamientos()
            Console.WriteLine("Fetched all alojamientos")
            If table.Rows.Count > 0 Then
                DataGridView1.DataSource = table
                DataGridView1.Visible = True
            End If

        Catch ex As Exception
            Messages.ShowErrorMessageBox(ex.Message)
            DataGridView1.Visible = False
        End Try

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeContainer()
        _dataAccess = _container.Resolve(Of DataAccess)
        _configuration = _container.Resolve(Of IConfigurationDB)
        _alojamientoService = _container.Resolve(Of IAlojamientoService)

    End Sub

End Class
