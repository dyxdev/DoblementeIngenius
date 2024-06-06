Imports ClassLibrary1
Imports Services

Public Class _Default
    Inherits Page

    Private _dataAccess As DataAccess
    Private _alojamientoService As IAlojamientoService
    Private _configuration As IConfigurationDB

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Dim container = HttpContext.Current.Items("AutofacContainer")

        If container IsNot Nothing Then
            ' Resolve dependencies
            _dataAccess = container.Resolve(Of DataAccess)
            _configuration = container.Resolve(Of IConfigurationDB)
            _alojamientoService = container.Resolve(Of IAlojamientoService)
        Else
            _dataAccess = New DataAccess()
            _configuration = New ConfigurationDB()
            _alojamientoService = New AlojamientoService(New AlojamientoRepository(_dataAccess))

        End If

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label2.Visible = True
        lblError.Visible = False

        Dim server As String = TextBoxserver.Text
        Dim dbName As String = TextBoxDatabase.Text
        Dim user As String = TextBoxUser.Text
        Dim pass As String = TextBoxPassword.Text

        Try
            Dim stringConnection As String = _configuration.connect(server, user, pass, dbName)
            Console.WriteLine(stringConnection)
            Console.WriteLine("Connected")
            _dataAccess.UconnectionString = stringConnection
            Dim table As DataTable = _alojamientoService.getAllAlojamientos()
            Console.WriteLine("Fetched all alojamientos")
            If table.Rows.Count > 0 Then
                TableGrid.Visible = True
                TableGrid.DataSource = table
                TableGrid.DataBind()
            End If
            Label2.Visible = False

        Catch ex As Exception
            lblError.Text = ex.Message
            lblError.Visible = True
            Label2.Visible = False
        End Try


    End Sub
End Class