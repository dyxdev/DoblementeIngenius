Imports System.Data.SqlClient
Public Class ModelProveedor
    Public Property Identifier As Integer
    Public Property Proveedor As String
    Public Property ActivoPC As Boolean
    Public Property CodProveedor As String
    Public Property NombreFiscal As String
    Public Property CIF As String
    Public Property Direccion As String
    Public Property Poblacion As String
    Public Property Provincia As String
    Public Property Pais As String
    Public Property CPOS As String
    Public Property Contacto As String
    Public Property Telefono As String
    Public Property Telefono2 As String
    Public Property Fax As String
    Public Property Email As String
    Public Property CContable As String
    Public Property NPeticionR As Boolean
    Public Property NConfirmacionR As Boolean
    Public Property NCancelacionR As Boolean
    Public Property IDMoneda As Integer
    Public Property EmailCorfirmacion As String
    Public Property ContactoConfirmacion As String
    Public Property Observaciones As String
    Public Property Web As String
    Public Property SWIFT As String
    Public Property IBAN As String
    Public Property Banco As String
    Public Property IDFormaPago As Integer
    Public Property Comisionable As Boolean
    Public Property IDEmpresa As Integer
    Public Property EmailCancelacion As String
    Public Property ContactoCancelacion As String
    Public Property EmailPeticion As String
    Public Property ContactoPeticion As String
    Public Property PuedeConfPeticion As Boolean
    Public Property IDIdioma As Integer
    Public Property Propio As Boolean
    Public Property MeFactura As Boolean
    Public Property CobraAlCliente As Byte
    Public Property PrioridadFCobro As Byte
    Public Property FacturaCliente As Boolean
    Public Property IDProvincia As Integer
    Public Property IDPais As Integer
    Public Property CContableComision As String
    Public Property ServicioSuplido As Boolean
    Public Property EnviaObservaciones As Boolean
    Public Property CtaPrevision As String
    Public Property IDPoliticaCancelacion As Integer
    Public Property Sectores As String
    Public Property IDClienteFacturacion As Integer
    Public Property IDGrupoServicio As Integer
    Public Property OmitirPCxDspNoReembolsable As Boolean
    Public Property EnviarReservasAgrupadas As Boolean
    Public Property PuedeCancPeticion As Boolean
    Public Property IDRetencion As Integer
    Public Property IDSerie As Integer
    Public Property PorcGuia As Decimal
    Public Property CtaContableGuia As String
    Public Property CodigoContabilidad As String
    Public Property Logo As Byte()
    Public Property CodigoExterno As String
    Public Property IDSociedad As Integer
    Public Property IDProveedorFacturacion As Integer
    Public Property MostrarBonoClienteIdiomaProveedor As Boolean
    Public Property NAgrupacionTrasladoR As Boolean
    Public Property EmailAgrupacionTraslado As String
    Public Property BloquearServiciosConfEnviada As Boolean
    Public Property IDIva As Integer
    Public Property OtrosDatosBancarios As String
    Public Property NoEnviarPkpass As Boolean
    Public Property IncluirObservacionesEnMailBono As Boolean
    Public Property CorregirHyperlinksInserguros As Boolean
    Public Property IncluirObservacionesSoloConHyperlink As Boolean
    Public Property NSolicitudCancelacionR As Boolean
    Public Property EmailSolicitudCancelacionR As String
    Public Property ContactoSolicitudCancelacionR As String
    Public Property NoIncluirUrlBonoEnObservaciones As Boolean
End Class

Public Class Proveedor
    Private Property _dataAccess As DataAccess


    Public Sub New(dataAccess As DataAccess)
        Me._dataAccess = dataAccess
    End Sub

    Public Function SelectAll() As List(Of ModelProveedor)
        Dim query As String = "SELECT * FROM PROVEEDOR"
        Dim proveedores As New List(Of ModelProveedor)()

        Dim dt As DataTable = Me._dataAccess.fetch(query)

        For Each row As DataRow In dt.Rows
            Dim proveedor As New ModelProveedor()

            proveedor.Identifier = Convert.ToInt32(row("Identifier"))
            proveedor.Proveedor = row("Proveedor").ToString()
            proveedor.ActivoPC = Convert.ToBoolean(row("ActivoPC"))
            proveedor.CodProveedor = row("CodProveedor").ToString()
            proveedor.NombreFiscal = row("NombreFiscal").ToString()
            proveedor.CIF = row("CIF").ToString()
            proveedor.Direccion = row("Direccion").ToString()
            proveedor.Poblacion = row("Poblacion").ToString()
            proveedor.Provincia = row("Provincia").ToString()
            proveedor.Pais = row("Pais").ToString()
            proveedor.CPOS = row("CPOS").ToString()
            proveedor.Contacto = row("Contacto").ToString()
            proveedor.Telefono = row("Telefono").ToString()
            proveedor.Telefono2 = row("Telefono2").ToString()
            proveedor.Fax = row("Fax").ToString()
            proveedor.Email = row("Email").ToString()
            proveedor.CContable = row("CContable").ToString()
            proveedor.NPeticionR = Convert.ToBoolean(row("NPeticionR"))
            proveedor.NConfirmacionR = Convert.ToBoolean(row("NConfirmacionR"))
            proveedor.NCancelacionR = Convert.ToBoolean(row("NCancelacionR"))
            proveedor.IDMoneda = Convert.ToInt32(row("IDMoneda"))
            proveedor.EmailCorfirmacion = row("EmailCorfirmacion").ToString()
            proveedor.ContactoConfirmacion = row("ContactoConfirmacion").ToString()
            proveedor.Observaciones = row("Observaciones").ToString()
            proveedor.Web = row("Web").ToString()
            proveedor.SWIFT = row("SWIFT").ToString()
            proveedor.IBAN = row("IBAN").ToString()
            proveedor.Banco = row("Banco").ToString()
            proveedor.IDFormaPago = Convert.ToInt32(row("IDFormaPago"))
            proveedor.Comisionable = Convert.ToBoolean(row("Comisionable"))
            proveedor.IDEmpresa = Convert.ToInt32(row("IDEmpresa"))
            proveedor.EmailCancelacion = row("EmailCancelacion").ToString()
            proveedor.ContactoCancelacion = row("ContactoCancelacion").ToString()
            proveedor.EmailPeticion = row("EmailPeticion").ToString()
            proveedor.ContactoPeticion = row("ContactoPeticion").ToString()
            proveedor.PuedeConfPeticion = Convert.ToBoolean(row("PuedeConfPeticion"))
            proveedor.IDIdioma = Convert.ToInt32(row("IDIdioma"))
            proveedor.Propio = Convert.ToBoolean(row("Propio"))
            proveedor.MeFactura = Convert.ToBoolean(row("MeFactura"))
            proveedor.CobraAlCliente = Convert.ToByte(row("CobraAlCliente"))
            proveedor.PrioridadFCobro = Convert.ToByte(row("PrioridadFCobro"))
            proveedor.FacturaCliente = Convert.ToBoolean(row("FacturaCliente"))
            proveedor.IDProvincia = Convert.ToInt32(row("IDProvincia"))
            proveedor.IDPais = Convert.ToInt32(row("IDPais"))
            proveedor.CContableComision = row("CContableComision").ToString()
            proveedor.ServicioSuplido = Convert.ToBoolean(row("ServicioSuplido"))
            proveedor.EnviaObservaciones = Convert.ToBoolean(row("EnviaObservaciones"))
            proveedor.CtaPrevision = row("CtaPrevision").ToString()
            proveedor.IDPoliticaCancelacion = Convert.ToInt32(row("IDPoliticaCancelacion"))
            proveedor.Sectores = row("Sectores").ToString()
            proveedor.IDClienteFacturacion = Convert.ToInt32(row("IDClienteFacturacion"))
            proveedor.IDGrupoServicio = Convert.ToInt32(row("IDGrupoServicio"))
            proveedor.OmitirPCxDspNoReembolsable = Convert.ToBoolean(row("OmitirPCxDspNoReembolsable"))
            proveedor.EnviarReservasAgrupadas = Convert.ToBoolean(row("EnviarReservasAgrupadas"))
            proveedor.PuedeCancPeticion = Convert.ToBoolean(row("PuedeCancPeticion"))
            proveedor.IDRetencion = Convert.ToInt32(row("IDRetencion"))
            proveedor.IDSerie = Convert.ToInt32(row("IDSerie"))
            proveedor.PorcGuia = Convert.ToDecimal(row("PorcGuia"))
            proveedor.CtaContableGuia = row("CtaContableGuia").ToString()
            proveedor.CodigoContabilidad = row("CodigoContabilidad").ToString()
            proveedor.Logo = DirectCast(row("Logo"), Byte())
            proveedor.CodigoExterno = row("CodigoExterno").ToString()
            proveedor.IDSociedad = Convert.ToInt32(row("IDSociedad"))
            proveedor.IDProveedorFacturacion = Convert.ToInt32(row("IDProveedorFacturacion"))
            proveedor.MostrarBonoClienteIdiomaProveedor = Convert.ToBoolean(row("MostrarBonoClienteIdiomaProveedor"))
            proveedor.NAgrupacionTrasladoR = Convert.ToBoolean(row("NAgrupacionTrasladoR"))
            proveedor.EmailAgrupacionTraslado = row("EmailAgrupacionTraslado").ToString()
            proveedor.BloquearServiciosConfEnviada = Convert.ToBoolean(row("BloquearServiciosConfEnviada"))
            proveedor.IDIva = Convert.ToInt32(row("IDIva"))
            proveedor.OtrosDatosBancarios = row("OtrosDatosBancarios").ToString()
            proveedor.NoEnviarPkpass = Convert.ToBoolean(row("NoEnviarPkpass"))
            proveedor.IncluirObservacionesEnMailBono = Convert.ToBoolean(row("IncluirObservacionesEnMailBono"))
            proveedor.CorregirHyperlinksInserguros = Convert.ToBoolean(row("CorregirHyperlinksInserguros"))
            proveedor.IncluirObservacionesSoloConHyperlink = Convert.ToBoolean(row("IncluirObservacionesSoloConHyperlink"))
            proveedor.NSolicitudCancelacionR = Convert.ToBoolean(row("NSolicitudCancelacionR"))
            proveedor.EmailSolicitudCancelacionR = row("EmailSolicitudCancelacionR").ToString()
            proveedor.ContactoSolicitudCancelacionR = row("ContactoSolicitudCancelacionR").ToString()
            proveedor.NoIncluirUrlBonoEnObservaciones = Convert.ToBoolean(row("NoIncluirUrlBonoEnObservaciones"))

            proveedores.Add(proveedor)
        Next

        Return proveedores
    End Function

    Public Function Delete(ByVal Identifier As Int32) As Boolean
        Dim query As String = "DELETE FROM PROVEEDOR WHERE Identifier = @Identifier"

        Using connection As New SqlConnection(Me._dataAccess.UconnectionString)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@Identifier", Identifier)

                connection.Open()
                Dim rowsAffected As Integer = command.ExecuteNonQuery()

                Return rowsAffected > 0
            End Using
        End Using
    End Function

    Public Function Update(ByVal identifier As Int32, ByVal proveedor As String, ByVal activoPC As Boolean, ByVal codProveedor As String, ByVal nombreFiscal As String, ByVal cif As String, ByVal direccion As String, ByVal poblacion As String, ByVal provincia As String, ByVal pais As String, ByVal cpos As String, ByVal contacto As String, ByVal telefono As String, ByVal telefono2 As String, ByVal fax As String, ByVal email As String, ByVal cContable As String, ByVal nPeticionR As Boolean, ByVal nConfirmacionR As Boolean, ByVal nCancelacionR As Boolean, ByVal idMoneda As Integer, ByVal emailCorfirmacion As String, ByVal contactoConfirmacion As String, ByVal observaciones As String, ByVal web As String, ByVal swift As String, ByVal iban As String, ByVal banco As String, ByVal idFormaPago As Integer, ByVal comisionable As Boolean, ByVal idEmpresa As Integer, ByVal emailCancelacion As String, ByVal contactoCancelacion As String, ByVal emailPeticion As String, ByVal contactoPeticion As String, ByVal puedeConfPeticion As Boolean, ByVal idIdioma As Integer, ByVal propio As Boolean, ByVal meFactura As Boolean, ByVal cobraAlCliente As Byte, ByVal prioridadFCobro As Byte, ByVal facturaCliente As Boolean, ByVal idProvincia As Integer, ByVal idPais As Integer, ByVal cContableComision As String, ByVal servicioSuplido As Boolean, ByVal enviaObservaciones As Boolean, ByVal ctaPrevision As String, ByVal idPoliticaCancelacion As Integer, ByVal sectores As String, ByVal idClienteFacturacion As Integer, ByVal idGrupoServicio As Integer, ByVal omitirPCxDspNoReembolsable As Boolean, ByVal enviarReservasAgrupadas As Boolean, ByVal puedeCancPeticion As Boolean, ByVal idRetencion As Integer, ByVal idSerie As Integer, ByVal porcGuia As Decimal, ByVal ctaContableGuia As String, ByVal codigoContabilidad As String, ByVal logo As Byte(), ByVal codigoExterno As String, ByVal idSociedad As Integer, ByVal idProveedorFacturacion As Integer, ByVal mostrarBonoClienteIdiomaProveedor As Boolean, ByVal nAgrupacionTrasladoR As Boolean, ByVal emailAgrupacionTraslado As String, ByVal bloquearServiciosConfEnviada As Boolean, ByVal idIva As Integer, ByVal otrosDatosBancarios As String, ByVal noEnviarPkpass As Boolean, ByVal incluirObservacionesEnMailBono As Boolean, ByVal corregirHyperlinksInserguros As Boolean, ByVal incluirObservacionesSoloConHyperlink As Boolean, ByVal nSolicitudCancelacionR As Boolean, ByVal emailSolicitudCancelacionR As String, ByVal contactoSolicitudCancelacionR As String, ByVal noIncluirUrlBonoEnObservaciones As Boolean) As Boolean
        Dim query As String = "UPDATE FROM PROVEEDOR SET Proveedor  = @Proveedor,  ActivoPC = @ActivoPC, CodProveedor = @CodProveedor,
                                NombreFiscal = @NombreFiscal, CIF = @CIF, Direccion = @Direccion, Poblacion = @Poblacion, Provincia = @Provincia,
                                Pais = @Pais, CPOS = @CPOS, Contacto = @Contacto , Telefono = @Telefono, Telefono2 = @Telefono2, Fax = @Fax, email = @email, CContable = @CContable, NPeticionR = @NPeticionR,
                                NConfirmacionR = @NConfirmacionR, NCancelacionR = @NCancelacionR, IDMoneda = @IDMoneda, emailCorfirmacion = @emailCorfirmacion, ContactoConfirmacion = @ContactoConfirmacion,
                                Observaciones = @Observaciones, Web = @Web, SWIFT = @SWIFT, IBAN = @IBAN, Banco = @Banco, IDFormaPago = @IDFormaPago, Comisionable = @Comisionable, IDEmpresa = @IDEmpresa,
                                emailCancelacion = @emailCancelacion, ContactoCancelacion = @ContactoCancelacion, emailPeticion = @emailPeticion, ContactoPeticion = @ContactoPeticion, PuedeConfPeticion = @PuedeConfPeticion,
                                IDIdioma = @IDIdioma, Propio = @Propio, MeFactura = @MeFactura, CobraAlCliente = @CobraAlCliente, PrioridadFCobro = @PrioridadFCobro, FacturaCliente = @FacturaCliente,
                                IDProvincia = @IDProvincia, IDPais = @IDPais, CContableComision = @CContableComision, ServicioSuplido = @ServicioSuplido, EnviaObservaciones = @EnviaObservaciones,
                                CtaPrevision = @CtaPrevision, IDPoliticaCancelacion = @IDPoliticaCancelacion, Sectores = @Sectores, IDClienteFacturacion = @IDClienteFacturacion, IDGrupoServicio = @IDGrupoServicio,
                                OmitirPCxDspNoReembolsable = @OmitirPCxDspNoReembolsable, EnviarReservasAgrupadas = @EnviarReservasAgrupadas, PuedeCancPeticion = @PuedeCancPeticion, IDRetencion = @IDRetencion,
                                IDSerie = @IDSerie, PorcGuia = @PorcGuia, CtaContableGuia = @CtaContableGuia, CodigoContabilidad = @CodigoContabilidad, Logo = @Logo, CodigoExterno = @CodigoExterno,
                                IDSociedad = @IDSociedad, IDProveedorFacturacion = @IDProveedorFacturacion, MostrarBonoClienteIdiomaProveedor = @MostrarBonoClienteIdiomaProveedor, NAgrupacionTrasladoR = @NAgrupacionTrasladoR,
                                emailAgrupacionTraslado = @emailAgrupacionTraslado, BloquearServiciosConfEnviada = @BloquearServiciosConfEnviada, IDIva = @IDIva, OtrosDatosBancarios = @OtrosDatosBancarios,
                                NoEnviarPkpass = @NoEnviarPkpass, IncluirObservacionesEnMailBono = @IncluirObservacionesEnMailBono, CorregirHyperlinksInserguros = @CorregirHyperlinksInserguros,
                                IncluirObservacionesSoloConHyperlink = @IncluirObservacionesSoloConHyperlink, NSolicitudCancelacionR = @NSolicitudCancelacionR, EmailSolicitudCancelacionR = @EmailSolicitudCancelacionR,
                                ContactoSolicitudCancelacionR = @ContactoSolicitudCancelacionR, NoIncluirUrlBonoEnObservaciones = @NoIncluirUrlBonoEnObservaciones  WHERE Identifier = @Identifier"

        Using connection As New SqlConnection(Me._dataAccess.UconnectionString)
            Using command As New SqlCommand(query)
                command.Parameters.AddWithValue("@Proveedor", proveedor)
                command.Parameters.AddWithValue("@ActivoPC", activoPC)
                command.Parameters.AddWithValue("@CodProveedor", codProveedor)
                command.Parameters.AddWithValue("@NombreFiscal", nombreFiscal)
                command.Parameters.AddWithValue("@CIF", cif)
                command.Parameters.AddWithValue("@Direccion", direccion)
                command.Parameters.AddWithValue("@Poblacion", poblacion)
                command.Parameters.AddWithValue("@Provincia", provincia)
                command.Parameters.AddWithValue("@Pais", pais)
                command.Parameters.AddWithValue("@CPOS", cpos)
                command.Parameters.AddWithValue("@Contacto", contacto)
                command.Parameters.AddWithValue("@Telefono", telefono)
                command.Parameters.AddWithValue("@Telefono2", telefono2)
                command.Parameters.AddWithValue("@Fax", fax)
                command.Parameters.AddWithValue("@Email", email)
                command.Parameters.AddWithValue("@CContable", cContable)
                command.Parameters.AddWithValue("@NPeticionR", nPeticionR)
                command.Parameters.AddWithValue("@NConfirmacionR", nConfirmacionR)
                command.Parameters.AddWithValue("@NCancelacionR", nCancelacionR)
                command.Parameters.AddWithValue("@IDMoneda", idMoneda)
                command.Parameters.AddWithValue("@EmailCorfirmacion", emailCorfirmacion)
                command.Parameters.AddWithValue("@ContactoConfirmacion", contactoConfirmacion)
                command.Parameters.AddWithValue("@Observaciones", observaciones)
                command.Parameters.AddWithValue("@Web", web)
                command.Parameters.AddWithValue("@SWIFT", swift)
                command.Parameters.AddWithValue("@IBAN", iban)
                command.Parameters.AddWithValue("@Banco", banco)
                command.Parameters.AddWithValue("@IDFormaPago", idFormaPago)
                command.Parameters.AddWithValue("@Comisionable", comisionable)
                command.Parameters.AddWithValue("@IDEmpresa", idEmpresa)
                command.Parameters.AddWithValue("@EmailCancelacion", emailCancelacion)
                command.Parameters.AddWithValue("@ContactoCancelacion", contactoCancelacion)
                command.Parameters.AddWithValue("@EmailPeticion", emailPeticion)
                command.Parameters.AddWithValue("@ContactoPeticion", contactoPeticion)
                command.Parameters.AddWithValue("@PuedeConfPeticion", puedeConfPeticion)
                command.Parameters.AddWithValue("@IDIdioma", idIdioma)
                command.Parameters.AddWithValue("@Propio", propio)
                command.Parameters.AddWithValue("@MeFactura", meFactura)
                command.Parameters.AddWithValue("@CobraAlCliente", cobraAlCliente)
                command.Parameters.AddWithValue("@PrioridadFCobro", prioridadFCobro)
                command.Parameters.AddWithValue("@FacturaCliente", facturaCliente)
                command.Parameters.AddWithValue("@IDProvincia", idProvincia)
                command.Parameters.AddWithValue("@IDPais", idPais)
                command.Parameters.AddWithValue("@CContableComision", cContableComision)
                command.Parameters.AddWithValue("@ServicioSuplido", servicioSuplido)
                command.Parameters.AddWithValue("@EnviaObservaciones", enviaObservaciones)
                command.Parameters.AddWithValue("@ctaPrevision", ctaPrevision)
                command.Parameters.AddWithValue("@IDPoliticaCancelacion", idPoliticaCancelacion)
                command.Parameters.AddWithValue("@Sectores", sectores)
                command.Parameters.AddWithValue("@IDClienteFacturacion", idClienteFacturacion)
                command.Parameters.AddWithValue("@IDGrupoServicio", idGrupoServicio)
                command.Parameters.AddWithValue("@omitirPCxDspNoReembolsable", omitirPCxDspNoReembolsable)
                command.Parameters.AddWithValue("@EnviarReservasAgrupadas", enviarReservasAgrupadas)
                command.Parameters.AddWithValue("@PuedeCancPeticion", puedeCancPeticion)
                command.Parameters.AddWithValue("@IDRetencion", idRetencion)
                command.Parameters.AddWithValue("@idSerie", idSerie)
                command.Parameters.AddWithValue("@PorcGuia", porcGuia)
                command.Parameters.AddWithValue("@CtaContableGuia", ctaContableGuia)
                command.Parameters.AddWithValue("@CodigoContabilidad", codigoContabilidad)
                command.Parameters.AddWithValue("@Logo", logo)
                command.Parameters.AddWithValue("@CodigoExterno", codigoExterno)
                command.Parameters.AddWithValue("@idSociedad", idSociedad)
                command.Parameters.AddWithValue("@IDProveedorFacturacion", idProveedorFacturacion)
                command.Parameters.AddWithValue("@MostrarBonoClienteIdiomaProveedor", mostrarBonoClienteIdiomaProveedor)
                command.Parameters.AddWithValue("@NAgrupacionTrasladoR", nAgrupacionTrasladoR)
                command.Parameters.AddWithValue("@emailAgrupacionTraslado", emailAgrupacionTraslado)
                command.Parameters.AddWithValue("@BloquearServiciosConfEnviada", bloquearServiciosConfEnviada)
                command.Parameters.AddWithValue("@IDIva", idIva)
                command.Parameters.AddWithValue("@OtrosDatosBancarios", otrosDatosBancarios)
                command.Parameters.AddWithValue("@noEnviarPkpass", noEnviarPkpass)
                command.Parameters.AddWithValue("@IncluirObservacionesEnMailBono", incluirObservacionesEnMailBono)
                command.Parameters.AddWithValue("@CorregirHyperlinksInserguros", corregirHyperlinksInserguros)
                command.Parameters.AddWithValue("@incluirObservacionesSoloConHyperlink", incluirObservacionesSoloConHyperlink)
                command.Parameters.AddWithValue("@NSolicitudCancelacionR", nSolicitudCancelacionR)
                command.Parameters.AddWithValue("@EmailSolicitudCancelacionR", emailSolicitudCancelacionR)
                command.Parameters.AddWithValue("@contactoSolicitudCancelacionR", contactoSolicitudCancelacionR)
                command.Parameters.AddWithValue("@NoIncluirUrlBonoEnObservaciones", noIncluirUrlBonoEnObservaciones)

                ' Ejecutar el comando y obtener el ID del nuevo registro insertado
                connection.Open()
                Dim rowsAffected As Integer = command.ExecuteNonQuery()

                Return rowsAffected > 0
            End Using
        End Using
    End Function

    Public Function Create(ByVal proveedor As String, ByVal activoPC As Boolean, ByVal codProveedor As String, ByVal nombreFiscal As String, ByVal cif As String, ByVal direccion As String, ByVal poblacion As String, ByVal provincia As String, ByVal pais As String, ByVal cpos As String, ByVal contacto As String, ByVal telefono As String, ByVal telefono2 As String, ByVal fax As String, ByVal email As String, ByVal cContable As String, ByVal nPeticionR As Boolean, ByVal nConfirmacionR As Boolean, ByVal nCancelacionR As Boolean, ByVal idMoneda As Integer, ByVal emailCorfirmacion As String, ByVal contactoConfirmacion As String, ByVal observaciones As String, ByVal web As String, ByVal swift As String, ByVal iban As String, ByVal banco As String, ByVal idFormaPago As Integer, ByVal comisionable As Boolean, ByVal idEmpresa As Integer, ByVal emailCancelacion As String, ByVal contactoCancelacion As String, ByVal emailPeticion As String, ByVal contactoPeticion As String, ByVal puedeConfPeticion As Boolean, ByVal idIdioma As Integer, ByVal propio As Boolean, ByVal meFactura As Boolean, ByVal cobraAlCliente As Byte, ByVal prioridadFCobro As Byte, ByVal facturaCliente As Boolean, ByVal idProvincia As Integer, ByVal idPais As Integer, ByVal cContableComision As String, ByVal servicioSuplido As Boolean, ByVal enviaObservaciones As Boolean, ByVal ctaPrevision As String, ByVal idPoliticaCancelacion As Integer, ByVal sectores As String, ByVal idClienteFacturacion As Integer, ByVal idGrupoServicio As Integer, ByVal omitirPCxDspNoReembolsable As Boolean, ByVal enviarReservasAgrupadas As Boolean, ByVal puedeCancPeticion As Boolean, ByVal idRetencion As Integer, ByVal idSerie As Integer, ByVal porcGuia As Decimal, ByVal ctaContableGuia As String, ByVal codigoContabilidad As String, ByVal logo As Byte(), ByVal codigoExterno As String, ByVal idSociedad As Integer, ByVal idProveedorFacturacion As Integer, ByVal mostrarBonoClienteIdiomaProveedor As Boolean, ByVal nAgrupacionTrasladoR As Boolean, ByVal emailAgrupacionTraslado As String, ByVal bloquearServiciosConfEnviada As Boolean, ByVal idIva As Integer, ByVal otrosDatosBancarios As String, ByVal noEnviarPkpass As Boolean, ByVal incluirObservacionesEnMailBono As Boolean, ByVal corregirHyperlinksInserguros As Boolean, ByVal incluirObservacionesSoloConHyperlink As Boolean, ByVal nSolicitudCancelacionR As Boolean, ByVal emailSolicitudCancelacionR As String, ByVal contactoSolicitudCancelacionR As String, ByVal noIncluirUrlBonoEnObservaciones As Boolean) As Int32
        Dim query As String = "INSERT INTO Proveedor (
                                Proveedor, ActivoPC, CodProveedor, NombreFiscal, CIF, Direccion, Poblacion, Provincia,
                                Pais, CPOS, Contacto, Telefono, Telefono2, Fax, email, CContable, NPeticionR,
                                NConfirmacionR, NCancelacionR, IDMoneda, emailCorfirmacion, ContactoConfirmacion,
                                Observaciones, Web, SWIFT, IBAN, Banco, IDFormaPago, Comisionable, IDEmpresa,
                                emailCancelacion, ContactoCancelacion, emailPeticion, ContactoPeticion, PuedeConfPeticion,
                                IDIdioma, Propio, MeFactura, CobraAlCliente, PrioridadFCobro, FacturaCliente,
                                IDProvincia, IDPais, CContableComision, ServicioSuplido, EnviaObservaciones,
                                CtaPrevision, IDPoliticaCancelacion, Sectores, IDClienteFacturacion, IDGrupoServicio,
                                OmitirPCxDspNoReembolsable, EnviarReservasAgrupadas, PuedeCancPeticion, IDRetencion,
                                IDSerie, PorcGuia, CtaContableGuia, CodigoContabilidad, Logo, CodigoExterno,
                                IDSociedad, IDProveedorFacturacion, MostrarBonoClienteIdiomaProveedor, NAgrupacionTrasladoR,
                                emailAgrupacionTraslado, BloquearServiciosConfEnviada, IDIva, OtrosDatosBancarios,
                                NoEnviarPkpass, IncluirObservacionesEnMailBono, CorregirHyperlinksInserguros,
                                IncluirObservacionesSoloConHyperlink, NSolicitudCancelacionR, EmailSolicitudCancelacionR,
                                ContactoSolicitudCancelacionR, NoIncluirUrlBonoEnObservaciones
                                )
                                VALUES (@Proveedor, @ActivoPC, @CodProveedor, @NombreFiscal, @CIF, @Direccion, @Poblacion, @Provincia,
                                @pais, @CPOS, @Contacto, @Telefono, @Telefono2, @Fax, @email, @CContable, @NPeticionR,
                                @nConfirmacionR, @NCancelacionR, @IDMoneda, @emailCorfirmacion, @ContactoConfirmacion,
                                @observaciones, @Web, @SWIFT, @IBAN, @Banco, @IDFormaPago, @Comisionable, @IDEmpresa,
                                @emailCancelacion, @ContactoCancelacion, @emailPeticion, @ContactoPeticion, @PuedeConfPeticion,
                                @idIdioma, @Propio, @MeFactura, @CobraAlCliente, @PrioridadFCobro, @FacturaCliente,
                                @idProvincia, @IDPais, @CContableComision, @ServicioSuplido, @EnviaObservaciones,
                                @ctaPrevision, @IDPoliticaCancelacion, @Sectores, @IDClienteFacturacion, @IDGrupoServicio,
                                @omitirPCxDspNoReembolsable, @EnviarReservasAgrupadas, @PuedeCancPeticion, @IDRetencion,
                                @idSerie, @PorcGuia, @CtaContableGuia, @CodigoContabilidad, @Logo, @CodigoExterno,
                                @idSociedad, @IDProveedorFacturacion, @MostrarBonoClienteIdiomaProveedor, @NAgrupacionTrasladoR,
                                @emailAgrupacionTraslado, @BloquearServiciosConfEnviada, @IDIva, @OtrosDatosBancarios,
                                @noEnviarPkpass, @IncluirObservacionesEnMailBono, @CorregirHyperlinksInserguros,
                                @incluirObservacionesSoloConHyperlink, @NSolicitudCancelacionR, @EmailSolicitudCancelacionR,
                                @contactoSolicitudCancelacionR, @NoIncluirUrlBonoEnObservaciones)"
        Using connection As New SqlConnection(Me._dataAccess.UconnectionString)
            Using command As New SqlCommand(query)
                command.Parameters.AddWithValue("@Proveedor", proveedor)
                command.Parameters.AddWithValue("@ActivoPC", activoPC)
                command.Parameters.AddWithValue("@CodProveedor", codProveedor)
                command.Parameters.AddWithValue("@NombreFiscal", nombreFiscal)
                command.Parameters.AddWithValue("@CIF", cif)
                command.Parameters.AddWithValue("@Direccion", direccion)
                command.Parameters.AddWithValue("@Poblacion", poblacion)
                command.Parameters.AddWithValue("@Provincia", provincia)
                command.Parameters.AddWithValue("@Pais", pais)
                command.Parameters.AddWithValue("@CPOS", cpos)
                command.Parameters.AddWithValue("@Contacto", contacto)
                command.Parameters.AddWithValue("@Telefono", telefono)
                command.Parameters.AddWithValue("@Telefono2", telefono2)
                command.Parameters.AddWithValue("@Fax", fax)
                command.Parameters.AddWithValue("@Email", email)
                command.Parameters.AddWithValue("@CContable", cContable)
                command.Parameters.AddWithValue("@NPeticionR", nPeticionR)
                command.Parameters.AddWithValue("@NConfirmacionR", nConfirmacionR)
                command.Parameters.AddWithValue("@NCancelacionR", nCancelacionR)
                command.Parameters.AddWithValue("@IDMoneda", idMoneda)
                command.Parameters.AddWithValue("@EmailCorfirmacion", emailCorfirmacion)
                command.Parameters.AddWithValue("@ContactoConfirmacion", contactoConfirmacion)
                command.Parameters.AddWithValue("@Observaciones", observaciones)
                command.Parameters.AddWithValue("@Web", web)
                command.Parameters.AddWithValue("@SWIFT", swift)
                command.Parameters.AddWithValue("@IBAN", iban)
                command.Parameters.AddWithValue("@Banco", banco)
                command.Parameters.AddWithValue("@IDFormaPago", idFormaPago)
                command.Parameters.AddWithValue("@Comisionable", comisionable)
                command.Parameters.AddWithValue("@IDEmpresa", idEmpresa)
                command.Parameters.AddWithValue("@EmailCancelacion", emailCancelacion)
                command.Parameters.AddWithValue("@ContactoCancelacion", contactoCancelacion)
                command.Parameters.AddWithValue("@EmailPeticion", emailPeticion)
                command.Parameters.AddWithValue("@ContactoPeticion", contactoPeticion)
                command.Parameters.AddWithValue("@PuedeConfPeticion", puedeConfPeticion)
                command.Parameters.AddWithValue("@IDIdioma", idIdioma)
                command.Parameters.AddWithValue("@Propio", propio)
                command.Parameters.AddWithValue("@MeFactura", meFactura)
                command.Parameters.AddWithValue("@CobraAlCliente", cobraAlCliente)
                command.Parameters.AddWithValue("@PrioridadFCobro", prioridadFCobro)
                command.Parameters.AddWithValue("@FacturaCliente", facturaCliente)
                command.Parameters.AddWithValue("@IDProvincia", idProvincia)
                command.Parameters.AddWithValue("@IDPais", idPais)
                command.Parameters.AddWithValue("@CContableComision", cContableComision)
                command.Parameters.AddWithValue("@ServicioSuplido", servicioSuplido)
                command.Parameters.AddWithValue("@EnviaObservaciones", enviaObservaciones)
                command.Parameters.AddWithValue("@ctaPrevision", ctaPrevision)
                command.Parameters.AddWithValue("@IDPoliticaCancelacion", idPoliticaCancelacion)
                command.Parameters.AddWithValue("@Sectores", sectores)
                command.Parameters.AddWithValue("@IDClienteFacturacion", idClienteFacturacion)
                command.Parameters.AddWithValue("@IDGrupoServicio", idGrupoServicio)
                command.Parameters.AddWithValue("@omitirPCxDspNoReembolsable", omitirPCxDspNoReembolsable)
                command.Parameters.AddWithValue("@EnviarReservasAgrupadas", enviarReservasAgrupadas)
                command.Parameters.AddWithValue("@PuedeCancPeticion", puedeCancPeticion)
                command.Parameters.AddWithValue("@IDRetencion", idRetencion)
                command.Parameters.AddWithValue("@idSerie", idSerie)
                command.Parameters.AddWithValue("@PorcGuia", porcGuia)
                command.Parameters.AddWithValue("@CtaContableGuia", ctaContableGuia)
                command.Parameters.AddWithValue("@CodigoContabilidad", codigoContabilidad)
                command.Parameters.AddWithValue("@Logo", logo)
                command.Parameters.AddWithValue("@CodigoExterno", codigoExterno)
                command.Parameters.AddWithValue("@idSociedad", idSociedad)
                command.Parameters.AddWithValue("@IDProveedorFacturacion", idProveedorFacturacion)
                command.Parameters.AddWithValue("@MostrarBonoClienteIdiomaProveedor", mostrarBonoClienteIdiomaProveedor)
                command.Parameters.AddWithValue("@NAgrupacionTrasladoR", nAgrupacionTrasladoR)
                command.Parameters.AddWithValue("@emailAgrupacionTraslado", emailAgrupacionTraslado)
                command.Parameters.AddWithValue("@BloquearServiciosConfEnviada", bloquearServiciosConfEnviada)
                command.Parameters.AddWithValue("@IDIva", idIva)
                command.Parameters.AddWithValue("@OtrosDatosBancarios", otrosDatosBancarios)
                command.Parameters.AddWithValue("@noEnviarPkpass", noEnviarPkpass)
                command.Parameters.AddWithValue("@IncluirObservacionesEnMailBono", incluirObservacionesEnMailBono)
                command.Parameters.AddWithValue("@CorregirHyperlinksInserguros", corregirHyperlinksInserguros)
                command.Parameters.AddWithValue("@incluirObservacionesSoloConHyperlink", incluirObservacionesSoloConHyperlink)
                command.Parameters.AddWithValue("@NSolicitudCancelacionR", nSolicitudCancelacionR)
                command.Parameters.AddWithValue("@EmailSolicitudCancelacionR", emailSolicitudCancelacionR)
                command.Parameters.AddWithValue("@contactoSolicitudCancelacionR", contactoSolicitudCancelacionR)
                command.Parameters.AddWithValue("@NoIncluirUrlBonoEnObservaciones", noIncluirUrlBonoEnObservaciones)

                ' Ejecutar el comando y obtener el ID del nuevo registro insertado
                command.ExecuteNonQuery()
                command.CommandText = "SELECT SCOPE_IDENTITY()"
                Dim newId As Integer = Convert.ToInt32(command.ExecuteScalar())

                ' Devolver el ID del nuevo registro insertado
                Return newId
            End Using
        End Using
    End Function

    Public Function SelectOne(Identifier As Int32) As ModelProveedor

        Dim query As String = "SELECT * FROM PROVEEDOR"
        Dim dt As DataTable = Me._dataAccess.fetch(query)
        Dim proveedor As New ModelProveedor()

        For Each row As DataRow In dt.Rows
            If Convert.ToInt32(row("Identifier") = Identifier) Then

                proveedor.Identifier = Convert.ToInt32(row("Identifier"))
                proveedor.Proveedor = row("Proveedor").ToString()
                proveedor.ActivoPC = Convert.ToBoolean(row("ActivoPC"))
                proveedor.CodProveedor = row("CodProveedor").ToString()
                proveedor.NombreFiscal = row("NombreFiscal").ToString()
                proveedor.CIF = row("CIF").ToString()
                proveedor.Direccion = row("Direccion").ToString()
                proveedor.Poblacion = row("Poblacion").ToString()
                proveedor.Provincia = row("Provincia").ToString()
                proveedor.Pais = row("Pais").ToString()
                proveedor.CPOS = row("CPOS").ToString()
                proveedor.Contacto = row("Contacto").ToString()
                proveedor.Telefono = row("Telefono").ToString()
                proveedor.Telefono2 = row("Telefono2").ToString()
                proveedor.Fax = row("Fax").ToString()
                proveedor.Email = row("Email").ToString()
                proveedor.CContable = row("CContable").ToString()
                proveedor.NPeticionR = Convert.ToBoolean(row("NPeticionR"))
                proveedor.NConfirmacionR = Convert.ToBoolean(row("NConfirmacionR"))
                proveedor.NCancelacionR = Convert.ToBoolean(row("NCancelacionR"))
                proveedor.IDMoneda = Convert.ToInt32(row("IDMoneda"))
                proveedor.EmailCorfirmacion = row("EmailCorfirmacion").ToString()
                proveedor.ContactoConfirmacion = row("ContactoConfirmacion").ToString()
                proveedor.Observaciones = row("Observaciones").ToString()
                proveedor.Web = row("Web").ToString()
                proveedor.SWIFT = row("SWIFT").ToString()
                proveedor.IBAN = row("IBAN").ToString()
                proveedor.Banco = row("Banco").ToString()
                proveedor.IDFormaPago = Convert.ToInt32(row("IDFormaPago"))
                proveedor.Comisionable = Convert.ToBoolean(row("Comisionable"))
                proveedor.IDEmpresa = Convert.ToInt32(row("IDEmpresa"))
                proveedor.EmailCancelacion = row("EmailCancelacion").ToString()
                proveedor.ContactoCancelacion = row("ContactoCancelacion").ToString()
                proveedor.EmailPeticion = row("EmailPeticion").ToString()
                proveedor.ContactoPeticion = row("ContactoPeticion").ToString()
                proveedor.PuedeConfPeticion = Convert.ToBoolean(row("PuedeConfPeticion"))
                proveedor.IDIdioma = Convert.ToInt32(row("IDIdioma"))
                proveedor.Propio = Convert.ToBoolean(row("Propio"))
                proveedor.MeFactura = Convert.ToBoolean(row("MeFactura"))
                proveedor.CobraAlCliente = Convert.ToByte(row("CobraAlCliente"))
                proveedor.PrioridadFCobro = Convert.ToByte(row("PrioridadFCobro"))
                proveedor.FacturaCliente = Convert.ToBoolean(row("FacturaCliente"))
                proveedor.IDProvincia = Convert.ToInt32(row("IDProvincia"))
                proveedor.IDPais = Convert.ToInt32(row("IDPais"))
                proveedor.CContableComision = row("CContableComision").ToString()
                proveedor.ServicioSuplido = Convert.ToBoolean(row("ServicioSuplido"))
                proveedor.EnviaObservaciones = Convert.ToBoolean(row("EnviaObservaciones"))
                proveedor.CtaPrevision = row("CtaPrevision").ToString()
                proveedor.IDPoliticaCancelacion = Convert.ToInt32(row("IDPoliticaCancelacion"))
                proveedor.Sectores = row("Sectores").ToString()
                proveedor.IDClienteFacturacion = Convert.ToInt32(row("IDClienteFacturacion"))
                proveedor.IDGrupoServicio = Convert.ToInt32(row("IDGrupoServicio"))
                proveedor.OmitirPCxDspNoReembolsable = Convert.ToBoolean(row("OmitirPCxDspNoReembolsable"))
                proveedor.EnviarReservasAgrupadas = Convert.ToBoolean(row("EnviarReservasAgrupadas"))
                proveedor.PuedeCancPeticion = Convert.ToBoolean(row("PuedeCancPeticion"))
                proveedor.IDRetencion = Convert.ToInt32(row("IDRetencion"))
                proveedor.IDSerie = Convert.ToInt32(row("IDSerie"))
                proveedor.PorcGuia = Convert.ToDecimal(row("PorcGuia"))
                proveedor.CtaContableGuia = row("CtaContableGuia").ToString()
                proveedor.CodigoContabilidad = row("CodigoContabilidad").ToString()
                proveedor.Logo = DirectCast(row("Logo"), Byte())
                proveedor.CodigoExterno = row("CodigoExterno").ToString()
                proveedor.IDSociedad = Convert.ToInt32(row("IDSociedad"))
                proveedor.IDProveedorFacturacion = Convert.ToInt32(row("IDProveedorFacturacion"))
                proveedor.MostrarBonoClienteIdiomaProveedor = Convert.ToBoolean(row("MostrarBonoClienteIdiomaProveedor"))
                proveedor.NAgrupacionTrasladoR = Convert.ToBoolean(row("NAgrupacionTrasladoR"))
                proveedor.EmailAgrupacionTraslado = row("EmailAgrupacionTraslado").ToString()
                proveedor.BloquearServiciosConfEnviada = Convert.ToBoolean(row("BloquearServiciosConfEnviada"))
                proveedor.IDIva = Convert.ToInt32(row("IDIva"))
                proveedor.OtrosDatosBancarios = row("OtrosDatosBancarios").ToString()
                proveedor.NoEnviarPkpass = Convert.ToBoolean(row("NoEnviarPkpass"))
                proveedor.IncluirObservacionesEnMailBono = Convert.ToBoolean(row("IncluirObservacionesEnMailBono"))
                proveedor.CorregirHyperlinksInserguros = Convert.ToBoolean(row("CorregirHyperlinksInserguros"))
                proveedor.IncluirObservacionesSoloConHyperlink = Convert.ToBoolean(row("IncluirObservacionesSoloConHyperlink"))
                proveedor.NSolicitudCancelacionR = Convert.ToBoolean(row("NSolicitudCancelacionR"))
                proveedor.EmailSolicitudCancelacionR = row("EmailSolicitudCancelacionR").ToString()
                proveedor.ContactoSolicitudCancelacionR = row("ContactoSolicitudCancelacionR").ToString()
                proveedor.NoIncluirUrlBonoEnObservaciones = Convert.ToBoolean(row("NoIncluirUrlBonoEnObservaciones"))
                Exit For
            End If
        Next

        Return proveedor
    End Function
End Class