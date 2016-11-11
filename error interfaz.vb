Public fec As String
Sub errorinterfaz()
'
' errorinterfaz Macro
'

'
    

    
    ' abrimos el archivo descargado de DUSOFT  lo organizamos y lo guardamos en el escritorio
    Application.ScreenUpdating = False
    strArchivo = Application.GetOpenFilename
    If strArchivo = False Then Exit Sub
    Workbooks.OpenText Filename:=strArchivo
    Application.DisplayAlerts = False
    ruta = ActiveWorkbook.Path
    arch = ActiveWorkbook.Name
    n_arch = Split(arch, ".")
    fec = n_arch(0)
    ult = Range("A1048576").End(xlUp).Row
    Workbooks.OpenText Filename:=strArchivo, Origin _
        :=xlMSDOS, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False _
        , Comma:=False, Space:=False, Other:=True, OtherChar:="@", FieldInfo _
        :=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), _
        Array(7, 2), Array(8, 2), Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2 _
        ), Array(14, 2), Array(15, 2), Array(16, 2), Array(17, 2), Array(18, 2), Array(19, 2), Array _
        (20, 2), Array(21, 2), Array(22, 2), Array(23, 2), Array(24, 2), Array(25, 2), Array(26, 2), _
        Array(27, 2), Array(28, 2), Array(29, 2), Array(30, 2), Array(31, 2), Array(32, 2)), _
        TrailingMinusNumbers:=True
    Application.WindowState = xlMaximized
    ActiveWorkbook.SaveAs Filename:="C:\Users\SOP1\Documents\errores interfaz\descargados\error_interfaz_" & fec & ".xlsx", _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Set errorint = ActiveWorkbook
    errorint.Activate
    
    ' creamos el archivo para enviar por correo
    Workbooks.Add
    Set enverrorint = ActiveWorkbook
    enverrorint.Activate
    Application.WindowState = xlMaximized
    ActiveCell.FormulaR1C1 = "FARMACIA"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "NUMERO DE FORMULA"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "NUMERO DE RADICADO"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "CODIGO ASIGNADO EN APLICATIVO EMILIO"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "MX"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "CANTIDAD DESPACHADA"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "OBSERVACIONES"
    errorint.Activate
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "=(MID(R[1]C[7],1,10))"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.NumberFormat = "@"
    Selection.Copy
    enverrorint.Activate
    Range("H1").Select
    ActiveSheet.Paste
    errorint.Activate
    Range("C2:C" & ult - 1).Select
    Selection.Copy
    enverrorint.Activate
    Range("A2").Select
    ActiveSheet.Paste
    errorint.Activate
    Range("E2:F" & ult - 1).Select
    Selection.Copy
    enverrorint.Activate
    Range("B2").Select
    ActiveSheet.Paste
    errorint.Activate
    Range("O2:O" & ult - 1).Select
    Selection.Copy
    enverrorint.Activate
    Range("D2").Select
    ActiveSheet.Paste
    errorint.Activate
    Range("Q2:Q" & ult - 1).Select
    Selection.Copy
    enverrorint.Activate
    Range("E2").Select
    ActiveSheet.Paste
    errorint.Activate
    Range("U2:U" & ult - 1).Select
    Selection.Copy
    enverrorint.Activate
    Range("F2").Select
    ActiveSheet.Paste
    errorint.Activate
    Range("B2:B" & ult - 1).Select
    Selection.Copy
    enverrorint.Activate
    Range("H2").Select
    ActiveSheet.Paste
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=TRIM(RC[-1])"
    Range("I2").Select
    Selection.AutoFill Destination:=Range("I2:I" & ult)
    Range("I2:I" & ult - 1).Select
    Selection.Copy
    Range("H2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=LOWER(RC[-1])"
    Range("I2").Select
    Selection.AutoFill Destination:=Range("I2:I" & ult - 1)
    Range("I2:I" & ult - 1).Select
    Selection.Copy
    Range("H2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("I2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(IFERROR(FIND(""no puede ser vaciono"",RC[-1]),0)<>0,""SIN CODIGO ASIGNADO EN APLICATIVO DE EMILIO"",IF(IFERROR(FIND(""no se encontro el producto"",RC[-1]),0)<>0,""CODIGO DUSOFT NO EXISTE"",IF(IFERROR(FIND(""no esta activo en la bodega"",RC[-1]),0)<>0,""CODIGO NO ESTA ACTIVO EN LA FARMACIA"",IF(IFERROR(FIND(""siendo usado en una formula de especiales"",RC[-1]),0)" & _
        "<>0,""CODIGO PACTADO EN FORMULA ESPECIAL"",IF(IFERROR(FIND(""siendo usado en una formula de acuerdo"",RC[-1]),0)<>0,""CODIGO ESPECIAL EN FORMULA NORMAL"",IF(IFERROR(FIND(""no es ni MR ni NPT"",RC[-1]),0)<>0,""CODIGO NO AUTORIZADO EN UNIDOSIS"",IF(IFERROR(FIND(""no se encuentra dispensacion previa"",RC[-1]),0)<>0,""MO SIN SOPORTE DE DISPENSACION"","""")))))))" & _
        ""
    Range("I2").Select
    Selection.AutoFill Destination:=Range("I2:I" & ult - 1)
    Range("I2:I" & ult - 1).Select
    Selection.Copy
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=+RC[-4]-0"
    Selection.AutoFill Destination:=Range("J2:J" & ult - 1)
    Range("J2:J" & ult - 1).Select
    Selection.Copy
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2").Select
    Selection.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = 1
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]>0,RC[-1],""----DEVOLUCION----  "" & RC[-1])"
    Selection.AutoFill Destination:=Range("J2:J" & ult - 1)
    Range("J2:J" & ult - 1).Select
    Selection.Copy
    Range("I2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("J2:J" & ult - 1).Select
    Selection.ClearContents
    Range("I2:I" & ult - 1).Select
    Selection.Copy
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H2:I" & ult - 1).Select
    Selection.ClearContents
    Cells.Select
    Cells.EntireColumn.AutoFit
        Range("A1:G1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("D:D").ColumnWidth = 17
    Columns("C:C").ColumnWidth = 11
    Columns("E:E").Select
    Selection.ColumnWidth = 50
    Columns("F:F").Select
    Selection.ColumnWidth = 13
    Rows("1:1").EntireRow.AutoFit
    Range("A2").Select
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs Filename:="C:\Users\SOP1\Documents\errores interfaz\enviados\error_interfaz_envio_" & fec & ".xlsx", _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    errorint.Activate
    Range("A2").Select
    ActiveWorkbook.Save
    ActiveWindow.Close
    enverrorint.Activate
    resultado = MsgBox("Desea enviar el archivo del dia  " & fec, vbYesNo + vbQuestion)
    If resultado = vbYes Then
        Call ENVIARCORREOMEDIPOL
    Else
        Exit Sub
    End If
    
End Sub
    
    Sub EnviarMailMEDIPOL()
        Application.ScreenUpdating = False
        Dim MailExitoso As Boolean
        'llamo a la funcion:
        MailExitoso = EnviarMailMEDIPOL_CDO()
        'si me devuelve un resultado Verdadero, todo salió bien:
        If MailExitoso = True Then
            MsgBox "El correo fué enviado satisfactoriamente", vbInformation, "Informe"
        Else
            MsgBox "No se pudo enviar el correo", vbCritical, "Informe"
        End If
    End Sub
    
Sub ENVIARCORREOMEDIPOL()
Dim iMsg As Object
Dim iConf As Object
Dim strbody As String
Dim diradjun As String
Dim dest As String
Dim Flds As Variant
Dim objShell As Object
Dim respuesta1 As Integer
Set objShell = CreateObject("WScript.Shell")

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")

iConf.Load -1 ' CDO Source Defaults
Set Flds = iConf.Fields
With Flds
.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "soporte.sistemas@medmfen.com"
.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "medmfen123"
.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"

.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
.Update
End With

strbody = "Saludos," & vbNewLine & vbNewLine & _
"En el archivo de excel adjunto están los radicados con fecha " & fec & " que no se han podido cargar al aplicativo dusoft y por ende no aparecen en los cortes de formulación." & vbNewLine & _
"Para cada linea de error se indica el motivo por el cual no cargo el producto" & vbNewLine & vbNewLine & _
"Gracias" & vbNewLine & vbNewLine & vbNewLine & vbNewLine & _
"SOPORTE SISTEMAS DUSOFT MEDIPOL." & vbNewLine & _
"TEL 3152571347" & vbNewLine & vbNewLine & _
"Antes de imprimir este u otro correo electrónico piense bien si es necesario hacerlo: El medioambiente es asunto de todos." & vbNewLine & vbNewLine & _
"Este mensaje es confidencial, puede contener información privilegiada y no puede ser usado ni divulgado por personas distintas de su destinatario. Si obtiene esta transmisión por error," & vbNewLine & _
"por favor destruya su contenido y avise al remitente. Está prohibida su retención, grabación, utilización o divulgación con cualquier propósito." & vbNewLine & _
"Este mensaje ha sido sometido a programas antivirus. No obstante, UTM MEDIPOL. No asume ninguna responsabilidad por eventuales daños generados por el recibo y uso de este material," & vbNewLine & _
"siendo responsabilidad del destinatario verificar con sus propios medios la existencia de virus u otros defectos."
With iMsg
Set .Configuration = iConf
 'dest = "cgonzalez@duarquint.com"
 dest = "calandate@gmail.com,aleon@duarquint.com,anoriega@duarquint.com,facturacion.wendy@duarquint.com,cgarcia@duarquint.com,compras1.bogota@duarquint.com,yaneth.duarte@duarquint.com"
 If dest <> vbNullString Then
.To = dest
End If
.CC = ""
.BCC = ""
.From = "soporte.sistemas@medmfen.com"
.Subject = "RADICADOS NO CARGADOS INTERFAZ " & fec
.TextBody = strbody
 diradjun = "C:\Users\SOP1\Documents\errores interfaz\enviados\error_interfaz_envio_" & fec & ".xlsx"
 ActiveWorkbook.Save
 ActiveWindow.Close
  If diradjun <> vbNullString Then
.AddAttachment (diradjun)
 End If
        'antes de enviar actualizamos los datos:
        .Configuration.Fields.Update
        'colocamos un capturador de errores, por las dudas:
        On Error Resume Next
        'enviamos el mail
        .Send
        'si el numero de error es 0 (o sea, no existieron errores en el proceso),
        'hago que la función retorne Verdadero
        If err.Number = 0 Then
            respuesta1 = objShell.Popup("Correo enviado", 1, "Mensaje")
            Set objShell = Nothing
        Else
            respuesta1 = objShell.Popup("Correo no enviado", 1, "Mensaje")
            Set objShell = Nothing
        End If
        'destruyo el objeto, para liberar los recursos del sistema
        If Not Email Is Nothing Then
            Set Email = Nothing
        End If
        'libero posibles errores
        On Error GoTo 0
'.Send
End With
 Application.CutCopyMode = False
End Sub