Sub error_interfaz_cargue()
'
' error_interfaz_cargue Macro
'

'
    arch = ActiveWorkbook.Name
    n_arch = Split(arch, ".")
    fec = n_arch(0)
    ult = Range("A1048576").End(xlUp).Row
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Range("B2").Select
    Selection.End(xlToRight).Select
    Range("AE2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-30],""@"",RC[-29],""@"",RC[-28],""@"",RC[-27],""@"",RC[-26],""@"",RC[-25],""@"",RC[-24],""@"",RC[-23],""@"",RC[-22],""@"",RC[-21],""@"",RC[-20],""@"",RC[-19],""@"",RC[-18],""@"",RC[-17],""@"",RC[-16],""@"",RC[-15],""@"",RC[-14],""@"",RC[-13],""@"",RC[-12],""@"",RC[-11],""@"",RC[-10],""@"",RC[-9],""@"",RC[-8],""@"",RC[-7],""@"",RC[-6],""@"",RC[-5],""@"",RC[-4],""@"",RC[-3],""@"",RC[-2],""@"",RC[-1])"
    Range("A2").Select
    Selection.End(xlToRight).Select
    Range("AE2").Select
    Selection.Copy
    Range("AE1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("AE1:AE" & ult).Select
    Selection.FillDown
    Selection.Copy
    Workbooks.Add
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs Filename:="C:\Users\SOP1\Documents\errores interfaz\cargados\error_interfaz_" & fec & ".txt", _
        FileFormat:=xlText, CreateBackup:=False
    Application.DisplayAlerts = False
    ActiveWindow.Close
    ActiveWindow.Close
    Application.DisplayAlerts = False

End Sub