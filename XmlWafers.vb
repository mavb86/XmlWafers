Imports System.Xml
Imports System.IO
Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop

Public Class XmlWafers

    Dim arrayXML(0) As String
    Dim nomArchivos(0) As String
    Dim valores(0) As String
    Dim jobName As String
    Dim classe As String
    Dim lifeTime As String
    Dim resistivity As String
    Dim thickness As String
    Dim ttv As String
    Dim microCrackResult As String
    Dim sideLength0 As String
    Dim sideLength1 As String
    Dim sideLength2 As String
    Dim sideLength3 As String
    Dim dimensionX As String
    Dim dimensionY As String
    Dim ChamferLength0 As String
    Dim ChamferLength1 As String
    Dim ChamferLength2 As String
    Dim ChamferLength3 As String
    Dim CornerAngle0 As String
    Dim CornerAngle1 As String
    Dim CornerAngle2 As String
    Dim CornerAngle3 As String
    Dim ClassifyResult As String
    Dim RearBWResult As String
    Dim PpNo As String
    Dim dset As New DataSet
    Dim conexion As New MySqlConnection
    Dim rutaError As String
    Dim nombreError As String
    Dim swError As String
    Dim contadorErrores As Byte

    Private Sub ButtonCargarXML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CargarXML.Click

        Dim totalArchivos As Integer
        totalArchivos = UBound(arrayXML)

        Me.barraProgreso.Minimum = 0
        Me.barraProgreso.Maximum = totalArchivos

        Dim i As Integer
        Dim XMLReader As Xml.XmlReader

        conexion.ConnectionString = "address=" & servidor & "; user id=" & usuario & "; password=" & contraseña & "; database=xml_manz"
        conexion.Open()

        For j = 0 To UBound(arrayXML)

            XMLReader = New Xml.XmlTextReader(arrayXML(j))

            While XMLReader.Read

                Select Case XMLReader.NodeType
                    Case Xml.XmlNodeType.Element
                        If XMLReader.AttributeCount > 0 Then
                            While XMLReader.MoveToNextAttribute
                                valores(i) = XMLReader.Value
                                ReDim Preserve valores(i + 1)
                                i = i + 1
                            End While
                        End If
                End Select

            End While

            XMLReader.Close()
            comprob_valores()
            rutaError = arrayXML(j)
            nombreError = nomArchivos(j)

            If contadorErrores >= 10 Then
                MsgBox("Se ha detenido el proceso porque han ocurrido más de 10 errores en la importación de archivos, compruebe que no esta importando archivos que ya han sido importados anteriormente", MsgBoxStyle.Critical, "DcWafers DATABASE XML_MANZ")
                ReDim valores(0)
                i = 0
                conexion.Close()
                ReDim arrayXML(0)
                ReDim nomArchivos(0)
                Me.CargarXML.Enabled = False
                Me.barraProgreso.Value = 0
                Exit Sub
            End If

            conexion_sql()
            ReDim valores(0)
            i = 0
            Me.barraProgreso.Value = j

        Next

        conexion.Close()
        ReDim arrayXML(0)
        ReDim nomArchivos(0)
        MsgBox("Proceso completado", MsgBoxStyle.Information, "DATABASE XML_MANZ")

        If swError = True Then
            MsgBox("Han ocurrido errores durante la importacion de archivos a la base de datos, por favor compruebe su email para poder analizarlos", MsgBoxStyle.Exclamation, "DATABASE XML_MANZ")
        End If

        Me.CargarXML.Enabled = False
        Me.barraProgreso.Value = 0
        swError = False

    End Sub

    Private Sub CargarXMLsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CargarXMLsToolStripMenuItem.Click
        Me.OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

        Try

            Dim archivos_seleccionados As Integer
            archivos_seleccionados = Me.OpenFileDialog1.FileNames.Length
            ReDim arrayXML(archivos_seleccionados - 1)
            ReDim nomArchivos(archivos_seleccionados - 1)
            Dim valor As String()
            Dim nombres As String()

            valor = Me.OpenFileDialog1.FileNames
            nombres = Me.OpenFileDialog1.SafeFileNames

            For i = 0 To UBound(arrayXML)
                arrayXML(i) = valor(i)
                nomArchivos(i) = nombres(i)
            Next

            MsgBox(UBound(arrayXML) + 1 & " archivos cargados en memoria correctamente", MsgBoxStyle.Information, "DATABASE XML_MANZ")
            Me.CargarXML.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message)
            contadorErrores += 1

        End Try

    End Sub

    Public Sub conexion_sql()

        Dim strComando As String

        Try
            Dim comandoSentencia As MySqlCommand
            Dim fecha As String
            Dim segundos As String
            fecha = valores(1).Substring(0, 8)
            segundos = valores(1).Substring(9, 6)
            strComando = "insert into manz values('" & valores(3) & "','" & valores(1) & "','" & fecha & "','" & segundos & "','" & jobName & "'," & classe & "," & lifeTime & "," & resistivity & "," & thickness & "," & ttv & ",'" & microCrackResult & "'," & sideLength0 & "," & sideLength1 & "," & sideLength2 & "," & sideLength3 & "," & dimensionX & "," & dimensionY & "," & ChamferLength0 & "," & ChamferLength1 & "," & ChamferLength2 & "," & ChamferLength3 & "," & CornerAngle0 & "," & CornerAngle1 & "," & CornerAngle2 & "," & CornerAngle3 & ",'" & ClassifyResult & "','" & RearBWResult & "');"
            comandoSentencia = New MySqlCommand(strComando, conexion)
            comandoSentencia.ExecuteNonQuery()
        Catch ex As Exception
            errorMail(ex.ToString, rutaError, nombreError)
            swError = True
            contadorErrores += 1
        End Try

    End Sub

    Public Sub comprob_valores()

        jobName = ""
        classe = ""
        lifeTime = 0
        resistivity = 0
        thickness = 0
        ttv = 0
        microCrackResult = ""
        sideLength0 = 0
        sideLength1 = 0
        sideLength2 = 0
        sideLength3 = 0
        dimensionX = 0
        dimensionY = 0
        ChamferLength0 = 0
        ChamferLength1 = 0
        ChamferLength2 = 0
        ChamferLength3 = 0
        CornerAngle0 = 0
        CornerAngle1 = 0
        CornerAngle2 = 0
        CornerAngle3 = 0
        ClassifyResult = ""
        RearBWResult = ""

        For i = 0 To UBound(valores)
            If valores(i) = "JOBNAME" Then
                jobName = valores(i + 1)
            ElseIf valores(i) = "CLASS" Then
                classe = valores(i + 1)
            ElseIf valores(i) = "SEMILAB-LIFETIME.LIFETIME" Then
                lifeTime = valores(i + 1)
            ElseIf valores(i) = "SEMILAB-THICKNESS.RESISTIVITY" Then
                resistivity = valores(i + 1)
            ElseIf valores(i) = "SEMILAB-THICKNESS.AVERAGE" Then
                thickness = valores(i + 1)
            ElseIf valores(i) = "SEMILAB-THICKNESS.TTV" Then
                ttv = valores(i + 1)
            ElseIf valores(i) = "MICROCRACK.CLASSIFYRESULT" Then
                microCrackResult = valores(i + 1)
            ElseIf valores(i) = "GEOMETRY.SIDE.LENGTHS[0]" Then
                sideLength0 = valores(i + 1)
            ElseIf valores(i) = "GEOMETRY.SIDE.LENGTHS[1]" Then
                sideLength1 = valores(i + 1)
            ElseIf valores(i) = "GEOMETRY.SIDE.LENGTHS[2]" Then
                sideLength2 = valores(i + 1)
            ElseIf valores(i) = "GEOMETRY.SIDE.LENGTHS[3]" Then
                sideLength3 = valores(i + 1)
            ElseIf valores(i) = "GEOMETRY.DIMENSION.X_0" Then
                dimensionX = valores(i + 1)
            ElseIf valores(i) = "GEOMETRY.DIMENSION.Y_0" Then
                dimensionY = valores(i + 1)
            ElseIf valores(i) = "REARBW.CHAMFER.LENGTHS[0]" Then
                ChamferLength0 = valores(i + 1)
            ElseIf valores(i) = "REARBW.CHAMFER.LENGTHS[1]" Then
                ChamferLength1 = valores(i + 1)
            ElseIf valores(i) = "REARBW.CHAMFER.LENGTHS[2]" Then
                ChamferLength2 = valores(i + 1)
            ElseIf valores(i) = "REARBW.CHAMFER.LENGTHS[3]" Then
                ChamferLength3 = valores(i + 1)
            ElseIf valores(i) = "GEOMETRY.CORNER.ANGLES[0]" Then
                CornerAngle0 = valores(i + 1)
            ElseIf valores(i) = "GEOMETRY.CORNER.ANGLES[1]" Then
                CornerAngle1 = valores(i + 1)
            ElseIf valores(i) = "GEOMETRY.CORNER.ANGLES[2]" Then
                CornerAngle2 = valores(i + 1)
            ElseIf valores(i) = "GEOMETRY.CORNER.ANGLES[3]" Then
                CornerAngle3 = valores(i + 1)
            ElseIf valores(i) = "GEOMETRY.CLASSIFYRESULT" Then
                ClassifyResult = valores(i + 1)
            ElseIf valores(i) = "REARBW.CLASSIFYRESULT" Then
                RearBWResult = valores(i + 1)
            End If
        Next

        If classe = "" Then
            classe = 10
        End If

    End Sub

    Private Sub SubirXML_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.CargarXML.Enabled = False
        swError = False
    End Sub

    Private Sub errorMail(ByVal cuerpo As String, ByVal ruta As String, ByVal nombre As String)

        Dim oApp As Outlook._Application
        oApp = New Outlook.Application()

        Dim oMsg As Outlook._MailItem
        oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)
        oMsg.Subject = "Error MANZ en archivo: " & nombre
        oMsg.Body = "Servidor: " & servidor & vbCrLf & "Usuario del programa: " & usuario & vbCrLf & "Fecha y hora del error: " & Now & vbCrLf & "Archivo: " & nombre & vbCrLf & "Ruta del archivo: " & ruta & vbCrLf & vbCrLf & "##################################################################################################################################################" & vbCrLf & cuerpo & vbCrLf & "##################################################################################################################################################"
        oMsg.To = "user@company.domain"

        Dim sSource As String = ruta
        Dim sDisplayName As String = nombre
        Dim sBodyLen As String = "1000000"
        Dim oAttachs As Outlook.Attachments = oMsg.Attachments
        Dim oAttach As Outlook.Attachment

        oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)
        oMsg.Send()
        oApp = Nothing
        oMsg = Nothing
        oAttach = Nothing
        oAttachs = Nothing

    End Sub
End Class
