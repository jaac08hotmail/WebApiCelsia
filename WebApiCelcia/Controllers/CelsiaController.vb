Imports System.Net
Imports System.Web.Http
Imports WebApiCelcia.Conexion

Namespace Controllers


    <RoutePrefix("api/Celcia")>
    Public Class CelciaController
        Inherits ApiController

        ' GET: api/Celcia
        <Route("GetValues")>
        Public Function GetValues() As IEnumerable(Of String)
            Dim ds As New DataSet
            Dim dt As DataTable
            Dim Result As String

            ds = Conexion.OptenerPuertos()
            dt = ds.Tables(0)
            If Not dt Is Nothing Then
                For Each dRow As DataRow In dt.Rows
                    'Agrega la consulta a la grilla.
                    Result = Result & " " & CInt(dRow("TONUMERO")) & " " & dRow("TOPUERTO") & " " & dRow("TODESCRIPCION") & " " _
                     & Format(Date.Now, "MMM/dd/yyyy HH:mm")
                Next
            End If

            Return New String() {Result, "value2"}
        End Function

        ' GET: api/Celcia/5
        Public Function GetValue(ByVal id As Integer) As String
            Return "value"
        End Function

        ' POST: api/Celcia
        Public Sub PostValue(<FromBody()> ByVal value As String)

        End Sub

        ' POST: api/Celcia
        <Route("PostPrueba")>
        Public Function PostPrueba(ByVal user As Usuario) As IHttpActionResult
            Dim ado As New Conexion
            Dim ds As DataSet
            Dim dt As DataTable
            Dim respuesta As String
            Try
                ds = Conexion.OptenerPuertos()
                dt = ds.Tables(0)
                If Not dt Is Nothing Then
                    For Each dRow As DataRow In dt.Rows
                        'Agrega la consulta a la grilla.
                        respuesta = respuesta & " " & CInt(dRow("TONUMERO")) & " " & dRow("TOPUERTO") & " " & dRow("TODESCRIPCION") & " " _
                     & Format(Date.Now, "MMM/dd/yyyy HH:mm")
                    Next
                End If
            Catch ex As Exception
                Me.BadRequest("Mensaje personalizado" & ex.Message)
            End Try
            Return Ok(respuesta)
        End Function


        ' PUT: api/Celcia/5
        Public Sub PutValue(ByVal id As Integer, <FromBody()> ByVal value As String)

        End Sub

        ' DELETE: api/Celcia/5
        Public Sub DeleteValue(ByVal id As Integer)

        End Sub


        <Route("PostConsultaUsuario")>
        Public Function PostConsultaUsuario(ByVal user As Usuario) As IHttpActionResult
            Dim usuario As New Usuario

            Try

                user = CRUD.Consultarusuario(user)

            Catch ex As Exception
                Me.BadRequest("Mensaje personalizado" & ex.Message)
            End Try
            Return Ok(user)
        End Function


    End Class
End Namespace