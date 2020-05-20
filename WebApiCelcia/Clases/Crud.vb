Public Class CRUD


    Public Shared Function Consultarusuario(ByVal user As Usuario) As Usuario

        Try
            Dim ado As New Conexion
            Dim ds As DataSet
            Dim dt As DataTable
            Dim respuesta As String
            Dim usuario As New Usuario
            Dim modulos As New List(Of Modulo)

            Dim QRY As String

            QRY = "select IDUser,Unombres,IDRol,Rnombre,PfkMod from Usuario,Rol,perfil " &
                                " where Ulogin='" & user.login & "' and Upasw = '" & user.passw & "' and UFkRol = IDRol and IDRol = PfkRol "

            ds = Conexion.QryDatos(QRY)
            dt = ds.Tables(0)
            If Not dt Is Nothing Then
                For Each dRow As DataRow In dt.Rows
                    usuario.IDUser = dRow("IDUser")
                    usuario.Rnombre = dRow("Rnombre")
                    usuario.UFkRol = dRow("IDRol")
                    usuario.Rnombre = dRow("Rnombre")

                    Dim modulo As New Modulo

                    modulo.IDMod = dRow("PfkMod")
                    modulos.Add(modulo)
                Next
            End If

            usuario.Modulos = modulos

            Return usuario

        Catch ex As Exception

            Return Nothing
        End Try

    End Function

End Class
