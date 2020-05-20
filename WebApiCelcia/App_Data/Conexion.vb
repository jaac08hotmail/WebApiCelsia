Imports System.Data.SqlClient

Public Class Conexion

    Public Shared Function QryDatos(ByVal qry As String) As DataSet


        Dim cnn = My.Settings.StringConnection.ToString()
        Dim Cn As New SqlConnection(cnn.ToString)
        Dim dat As New SqlDataAdapter
        Dim ds As New DataSet
        Dim cm As New SqlCommand

        Cn.Open()
        cm.Connection = Cn
        cm.CommandText = qry
        dat.SelectCommand = cm

        dat.Fill(ds, "QryDatos")
        Cn.Close()
        Return ds
    End Function
    Public Shared Function OptenerPuertos() As DataSet

        Dim cnn = My.Settings.StringConnection.ToString()
        Dim Cn As New SqlConnection(cnn.ToString)
        Dim ds As New DataSet
        Try
            Dim QRY As String = "SELECT tonumero, toip, topuerto,todescripcion,toestado FROM spepuertos "
            Dim da As SqlDataAdapter = New SqlDataAdapter(QRY, Cn)

            da.Fill(ds, "spepuertos")
        Catch ex As Exception

        End Try


        Return ds
    End Function
    Public Sub QryInsertAudit(ByVal cadqry As String, ByVal Mensaje As String, ByVal puerto As String)
        Dim cnn = My.Settings.StringConnection.ToString()
        Dim Cn As New SqlConnection(cnn)
        Dim dat As New SqlDataAdapter
        Dim ds As New DataSet
        Dim cm As New SqlCommand
        Try

            Using Cn
                cadqry = Replace(cadqry, "'", "\")
                Mensaje = Replace(Mensaje, "'", "\")
                cadqry = "INSERT INTO Auditoria (Mensaje,Fecha,puerto) VALUES ('[" & cadqry & "]''[" & Mensaje & "]',convert(datetime,GETDATE(),102),'" & puerto & "')"
                cm.Connection = Cn
                Cn.Open()
                cm.CommandText = cadqry
                cm.ExecuteNonQuery()
                Cn.Close()
            End Using

        Catch ex As Exception

        End Try

        Cn.Close()
    End Sub
    Public Shared Sub QryInsertError(ByVal cadqry As String, ByVal Mensaje As String, ByVal puerto As String)
        Dim cnn = My.Settings.StringConnection.ToString()
        Dim Cn As New SqlConnection(cnn)
        Dim dat As New SqlDataAdapter
        Dim ds As New DataSet
        Dim cm As New SqlCommand
        Try

            Using Cn
                cadqry = Replace(cadqry, "'", "\")
                Mensaje = Replace(Mensaje, "'", "\")
                cadqry = "INSERT INTO errmov (Mensaje,Fecha,puerto) VALUES ('[" & Mensaje & "] " & cadqry & "',convert(datetime,GETDATE(),102),'" & puerto & "')"
                Cn.Open()
                cm.Connection = Cn
                cm.CommandText = cadqry
                cm.ExecuteNonQuery()
                Cn.Close()
            End Using

        Catch ex As Exception

        End Try

        Cn.Close()
    End Sub

    Public Shared Sub QryInsert(ByVal cadqry As String, ByVal PUERTO As String)
        Dim cnn = My.Settings.StringConnection.ToString()
        Dim Cn As New SqlConnection(cnn)
        Dim dat As New SqlDataAdapter
        Dim cm As New SqlCommand
        Try

            Using Cn

                Cn.Open()
                cm.Connection = Cn
                cm.CommandText = cadqry
                cm.ExecuteNonQuery()
                Cn.Close()
            End Using

        Catch ex As Exception
            QryInsertError(cadqry, ex.Message, PUERTO)
        End Try

        Cn.Close()
    End Sub

    Public Shared Sub QryInsertImp(ByVal cadqry As String)
        Dim cnn = My.Settings.StringConnection.ToString()
        Dim Cn As New SqlConnection(cnn)
        Dim dat As New SqlDataAdapter
        Dim cm As New SqlCommand
        Try

            Using Cn

                Cn.Open()
                cm.Connection = Cn
                cm.CommandText = cadqry
                cm.ExecuteNonQuery()
                Cn.Close()
            End Using

        Catch ex As Exception
            QryInsertError(cadqry, ex.Message, "")
        End Try

        Cn.Close()
    End Sub

    Public Shared Function Sincronizar() As Integer
        Dim sDestino As String, nDestino As String, auxtam As String, Qry As String
        Dim i As Integer
        Dim Entry, txt_cadena(4) As String
        Dim numtiqmax As Double
        Dim VarCodRut As Integer 'Andres Mesa 11/19/2010 -> Para almacenar la ruta que viene del qry
        Dim ConDes As Integer, ConNomDes As Integer  'Andres Mesa 12/06/2010 -> Control que determina numero de linea de una trama
        Dim ds As New DataSet
        Dim ds2 As New DataSet
        numtiqmax = 0


        Try

            'Limpiamos las cajas de texto donde se desplagara la informacion generda
            txt_cadena(0) = ""
            txt_cadena(1) = ""
            txt_cadena(2) = ""
            txt_cadena(3) = ""

            i = 0

            '*******************************************************************************************************
            '*********************************** INICIO - 2000 - EMPRESAS ******************************************
            '*******************************************************************************************************
            Qry = "SELECT dscoddet, dsdes, dsval " _
                         & "FROM gesuptip,gedetsuptip " _
                         & "WHERE stcodtip=dscodtip and stdes = 'EMPRESAS' and dsest='A' " _
                         & "ORDER BY dsdes "
            ds = Conexion.QryDatos(Qry)

            If Not ds.Tables("QryDatos").Rows.Count = 0 Then
                Conexion.QryInsertImp("DELETE FROM spedatos ")
                Entry = ""


                'CICLO DE EMPRESAS
                For i = 0 To ds.Tables("QryDatos").Rows.Count - 1

                    Entry = CStr(ds.Tables("QryDatos").Rows(i)("dscoddet")) & ";" & CStr(ds.Tables("QryDatos").Rows(i)("dsdes")) & ";" & CStr(ds.Tables("QryDatos").Rows(i)("dsval"))

                    If txt_cadena(0) = "" Then
                        txt_cadena(0) = Entry
                    Else
                        txt_cadena(0) = txt_cadena(0) & ";>" & Entry
                    End If

                Next

                txt_cadena(0) = txt_cadena(0) & ";>"


                'LEVILLA. NOV 18-2010
                'EL CODIGO 2000 ES LA EMPRESA
                Qry = "INSERT INTO spedatos (" &
                      "macodigo, macodemp, madatos,manombre,maestado) " &
                      "VALUES ('" & 2000 & "', '0', '" & txt_cadena(0) & "', 'EMPRESA' ,'A')"
                Conexion.QryInsertImp(Qry)
            End If
            '*******************************************************************************************************
            '************************************ FINAL - 2000 - EMPRESAS ******************************************
            '*******************************************************************************************************

            '*******************************************************************************************************
            '********************************* INICIO - 9000 - PTO. VENTA ******************************************
            '*******************************************************************************************************
            'Andres Mesa 12/01/2010
            'Ahora consultamos la parametrizacion de las spectras y por cada registro, creamos un codigo 9000
            'con el punto de venta de dicha spectra.
            Qry = "SELECT dscoddet, dsdes,dsval FROM gesuptip, gedetsuptip " _
                & "WHERE stdes = 'PARAMETRIZACION SPECTRAS' " _
                & "AND stcodtip = dscodtip and dsest='A' " _
                & "ORDER BY dscoddet"
            ds = Conexion.QryDatos(Qry)
            Entry = ""
            If Not ds.Tables("QryDatos").Rows.Count = 0 Then

                For i = 0 To ds.Tables("QryDatos").Rows.Count - 1
                    VarCodRut = 9000 + (CInt(ds.Tables("QryDatos").Rows(i)("dscoddet")) - 4000)
                    Entry = CStr(ds.Tables("QryDatos").Rows(i)("dsval")) & ";" & CStr(ds.Tables("QryDatos").Rows(i)("dsdes"))

                    Qry = "INSERT INTO spedatos (" &
                            "macodigo, macodemp, madatos,manombre,maestado) " &
                         "VALUES ('" & VarCodRut & "', '0', '" & Entry & "', 'PTO VENTA','A' )"
                    Conexion.QryInsertImp(Qry)

                Next

            End If
            '*******************************************************************************************************
            '*********************************** FINAL - 9000 - PTO VENTA ******************************************
            '*******************************************************************************************************

            '*******************************************************************************************************
            '*********************************** INICIO - 5000 - DESTINOS *****************************************
            '*******************************************************************************************************
            'Andres Mesa 11/29/2010
            'Ahora, hacemos una consulta del catalogo PARAMETRIZACION SPECTRAS y por cada registro
            'Busca la ruta que le corresponde en la consulta que se abre en el tbl6
            Qry = "SELECT DSCODDET FROM gesuptip, gedetsuptip " _
                 & "WHERE stdes = 'PARAMETRIZACION SPECTRAS' " _
                 & "AND stcodtip = dscodtip " _
                 & "ORDER BY dscoddet"
            ds = Conexion.QryDatos(Qry)
            Entry = ""
            If Not ds.Tables("QryDatos").Rows.Count = 0 Then


                For i = 0 To ds.Tables("QryDatos").Rows.Count - 1
                    Dim cod As Integer, conNom As Integer
                    sDestino = ""

                    txt_cadena(0) = ""
                    txt_cadena(1) = ""
                    txt_cadena(2) = ""

                    'LEE LAS TARIFAS DE LA EMPRESA
                    Qry = "SELECT dscoddet,dsdes,dsval FROM gesuptip, gedetsuptip " _
                    & "WHERE stdes = 'DESTINOS' " _
                    & "AND stcodtip = dscodtip and dsest='A' " _
                    & "and dscoddet = " & CInt(ds.Tables("QryDatos").Rows(i)("dscoddet")) - 4000 & " " _
                    & "ORDER BY dscoddet"
                    ds2 = Conexion.QryDatos(Qry)
                    If Not ds2.Tables("QryDatos").Rows.Count = 0 Then
                        sDestino = ""
                        ConDes = 0
                        ConNomDes = 0
                        cod = (5000 + CInt(ds2.Tables("QryDatos").Rows(0)("dscoddet")))
                        conNom = (6000 + CInt(ds2.Tables("QryDatos").Rows(0)("dscoddet")))

                        For f = 0 To ds2.Tables("QryDatos").Rows.Count - 1

                            'CONCATENA TODOS  LOS DESTINOS
                            sDestino = CStr(ds2.Tables("QryDatos").Rows(f)("dsval"))
                            nDestino = CStr(ds2.Tables("QryDatos").Rows(f)("dsdes"))

                            If txt_cadena(2) = "" Then
                                txt_cadena(2) = sDestino
                            Else
                                txt_cadena(2) = txt_cadena(2) & ";" & sDestino
                            End If

                            If txt_cadena(1) = "" Then
                                txt_cadena(1) = nDestino
                            Else
                                txt_cadena(1) = txt_cadena(1) & ";" & nDestino
                            End If

                            If txt_cadena(1).Length > 1000 Then
                                ConNomDes = ConNomDes + 1
                                txt_cadena(1) = txt_cadena(1) & ";>"

                                Qry = "INSERT INTO ltmaestrom (" &
                                                "macodigo, macodemp, madatos,manombre,maestado, macontrol) " &
                                      "VALUES ('" & conNom & "@" & ConNomDes & "', '0', '" & txt_cadena(1) & "', 'DESTINOS','A', " & ConNomDes & ")"
                                Conexion.QryInsertImp(Qry)
                                txt_cadena(2) = ""
                                nDestino = ""
                            End If

                            'INSERTA BAJO LA SERIE 6000 LOS CODIGOS DE LOS DESTINOS
                            If txt_cadena(2).Length > 1000 Then
                                ConDes = ConDes + 1
                                txt_cadena(2) = txt_cadena(2) & ";>"

                                Qry = "INSERT INTO ltmaestrom (" &
                                                "macodigo, macodemp, madatos,manombre,maestado, macontrol) " &
                                      "VALUES ('" & cod & "@" & ConDes & "', '0', '" & txt_cadena(2) & "', 'CODIGOS','A', " & ConDes & ")"
                                Conexion.QryInsertImp(Qry)
                                txt_cadena(2) = ""
                                sDestino = ""
                            End If
                        Next

                        'INSERTA BAJO LA SERIE 6000 LOS CODIGOS DE LOS DESTINOS
                        txt_cadena(2) = txt_cadena(2) & ";>"
                        txt_cadena(1) = txt_cadena(1) & ";>"

                        If txt_cadena(1).Length > 2 Then
                            ConNomDes = ConNomDes + 1
                            Qry = "INSERT INTO spedatos (" &
                                        "macodigo, macodemp, madatos,manombre,maestado,macontrol) " &
                                  "VALUES ('" & conNom & "@" & ConNomDes & "', '0', '" & txt_cadena(1) & "', 'DESTINOS','A', " & ConNomDes & ")"
                            Conexion.QryInsertImp(Qry)
                        End If

                        If txt_cadena(2).Length > 2 Then
                            ConDes = ConDes + 1
                            Qry = "INSERT INTO spedatos (" &
                                        "macodigo, macodemp, madatos,manombre,maestado,macontrol) " &
                                  "VALUES ('" & cod & "@" & ConDes & "', '0', '" & txt_cadena(2) & "', 'CODIGOS','A', " & ConDes & ")"
                            Conexion.QryInsertImp(Qry)
                        End If

                        Qry = "INSERT INTO spedatos (" &
                                   "macodigo, macodemp, madatos,manombre,maestado,macontrol) " &
                             "VALUES ('" & cod & "', '0', '" & ConDes & "', 'CODIGOS','A', " & ConDes & ")"
                        Conexion.QryInsertImp(Qry)

                        Qry = "INSERT INTO spedatos (" &
                                   "macodigo, macodemp, madatos,manombre,maestado,macontrol) " &
                             "VALUES ('" & conNom & "', '0', '" & ConNomDes & "', 'DESTINOS','A', " & ConNomDes & ")"
                        Conexion.QryInsertImp(Qry)

                    End If

                Next
            End If


            Return 1


        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Shared Function BuscarAgencia(ByVal Criterio As String) As DataSet

        Dim cnn = My.Settings.StringConnection.ToString()
        Dim Cn As New SqlConnection(cnn.ToString)
        Dim ds As New DataSet
        Try
            Dim QRY As String = "SELECT dscoddet,dsdes FROM gedetsuptip WHERE  dscodtip = 3 and dsval =" & Criterio
            Dim da As SqlDataAdapter = New SqlDataAdapter(QRY, Cn)

            da.Fill(ds, "gedetsuptip")


        Catch ex As Exception

        End Try


        Return ds
    End Function

End Class




