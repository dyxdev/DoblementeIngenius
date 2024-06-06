Imports System.Web.Optimization
Imports ClassLibrary1
Imports Services
Imports Autofac
Imports Autofac.Core

Public Class Global_asax
    Inherits HttpApplication

    Private _container As Container


    Sub Application_Start(sender As Object, e As EventArgs)
        ' Fires when the application is started
        RouteConfig.RegisterRoutes(RouteTable.Routes)
        BundleConfig.RegisterBundles(BundleTable.Bundles)
        Me.InitializeContainer()
        HttpContext.Current.Items.Add("AutofacContainer", Me._container)
    End Sub


    Private Sub InitializeContainer()

        Dim builder As New ContainerBuilder()

        builder.RegisterType(Of DataAccess).As(Of DataAccess)()
        builder.RegisterType(Of AlojamientoRepository).As(Of IRepository)()
        builder.RegisterType(Of AlojamientoService).As(Of IAlojamientoService)()
        builder.RegisterType(Of ConfigurationDB).As(Of IConfigurationDB)()

        _container = builder.Build()

    End Sub
End Class