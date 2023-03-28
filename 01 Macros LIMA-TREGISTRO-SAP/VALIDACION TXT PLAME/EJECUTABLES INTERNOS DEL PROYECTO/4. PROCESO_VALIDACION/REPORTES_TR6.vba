Attribute VB_Name = "REPORTES_TR6"
Sub Borrar_Filtro_De_Tabla()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

  Dim ws As Worksheet
  Dim sTable As String
  Dim loTable As ListObject
  sTable = "TR6_PARAMETRIZADA"
  Set ws = ActiveSheet
  Set loTable = ws.ListObjects(sTable)
  loTable.AutoFilter.ShowAllData
  
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub
Sub Borrar_Filtro_De_Tabla_tr5()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

  Dim ws As Worksheet
  Dim sTable As String
  Dim loTable As ListObject
  sTable = "TR5_PARAMETRIZADA"
  Set ws = ActiveSheet
  Set loTable = ws.ListObjects(sTable)
  loTable.AutoFilter.ShowAllData
  
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub REPORTE_TR6()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

    Call REPORTE_CUSSP
    Call REPORTE_REGIMEN
    Call REPORTE_REGIMEN_SALUD
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


Sub REPORTE_CUSSP()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Sheets("TR6_PARAMETRIZADA").Select
columnas = "A:ZZ"

With Sheets("TR6_PARAMETRIZADA").Range(columnas)
    Set DOC_TR = .Find("NUMERO DOCUMENTO TR6")
    Set CUSSP_TR = .Find("CUSSP TR")
    Set CUSSP_SAP = .Find("CUSSP SAP")
    Set VALIDACION_CUSSP = .Find("VALIDACIÓN CUSSP TR-SAP")
    
    Sheets("TR6_PARAMETRIZADA").Select

    'Range("TR6_PARAMETRIZADA[VALIDACIÓN CUSSP TR-SAP]").Select
    'Range("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA").AutoFilter Field:=VALIDACION_CUSSP.Column, Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"
    
    'Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_HORARIO_NOCTURNO_SAP.Column, _
            Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"
    
    Range("TR6_PARAMETRIZADA[[NUMERO DOCUMENTO TR6]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("A10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    Range("TR6_PARAMETRIZADA[[CUSSP TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("B10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Range("TR6_PARAMETRIZADA[[CUSSP SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("C10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Range("TR6_PARAMETRIZADA[[VALIDACIÓN CUSSP TR-SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("D10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
End With
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[[VALIDACIÓN CUSSP TR-SAP]]").Select
ActiveSheet.ShowAllData
Range("A1").Select
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("A9").Select
ActiveCell.Value = DOC_TR.Value
Range("B9").Select
ActiveCell.Value = CUSSP_TR.Value
Range("C9").Select
ActiveCell.Value = CUSSP_SAP.Value
Range("D9").Select
ActiveCell.Value = VALIDACION_CUSSP.Value

Set tbl = Range("A9").CurrentRegion
Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_CUSSP"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub




Sub REPORTE_REGIMEN()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Sheets("TR6_PARAMETRIZADA").Select
columnas = "A10:AE10"
Dim Value_column As Integer


With Sheets("TR6_PARAMETRIZADA").Range(columnas)
    Set DOC_TR = .Find("NUMERO DOCUMENTO TR6")
    Set TIPO_REGIMEN_TR = .Find("TIPO DE REGIMEN TR")
    Set TIPO_REGIMEN_SAP = .Find("TIPO DE REGIMEN SAP")
    Set VALIDACION_TIPO_REGIMEN = .Find("VALIDACIÓN TIPO DE REGIMEN")
    
    Sheets("TR6_PARAMETRIZADA").Select

    'Range("TR6_PARAMETRIZADA[VALIDACIÓN CUSSP TR-SAP]").Select
    'Range("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA").AutoFilter Field:=VALIDACION_TIPO_REGIMEN.Column, Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"
    
    'Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_HORARIO_NOCTURNO_SAP.Column, _
            Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"
    
    Range("TR6_PARAMETRIZADA[[NUMERO DOCUMENTO TR6]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("F10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    Range("TR6_PARAMETRIZADA[[TIPO DE REGIMEN TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("G10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Range("TR6_PARAMETRIZADA[[TIPO DE REGIMEN SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("H10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Range("TR6_PARAMETRIZADA[[VALIDACIÓN TIPO DE REGIMEN]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("I10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
End With
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[[VALIDACIÓN TIPO DE REGIMEN]]").Select
ActiveSheet.ShowAllData
Range("A1").Select
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("F9").Select
ActiveCell.Value = DOC_TR.Value
Range("G9").Select
ActiveCell.Value = TIPO_REGIMEN_TR.Value
Range("H9").Select
ActiveCell.Value = TIPO_REGIMEN_SAP.Value
Range("I9").Select
ActiveCell.Value = VALIDACION_TIPO_REGIMEN.Value

Set tbl = Range("F9").CurrentRegion
Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_REGIMENX"


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub filtro()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[[#Headers],[VALIDACIÓN TIPO DE REGIMEN TR-SAP]]").Select
    ActiveSheet.ListObjects("TR6_PARAMETRIZADA").Range.AutoFilter Field:=31, _
        Criteria1:="REGISTRAR EPS"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


Sub REPORTE_REGIMEN_SALUD()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


On Error Resume Next

Dim Value_column As Integer
Sheets("TR6_PARAMETRIZADA").Select

columnas = "A10:AE10"
With Sheets("TR6_PARAMETRIZADA").Range(columnas)
    Set DOC_TR = .Find("NUMERO DOCUMENTO TR6")
    Set TIPO_REGIMEN_SALUD_TR = .Find("TIPO DE REGIMEN SALUD TR")
    Set TIPO_REGIMEN_SALUD_SAP = .Find("TIPO DE REGIMEN SALUD SAP")
    Set VALIDACION_TIPO_REGIMEN_SALUD = .Find("VALIDACIÓN TIPO DE REGIMEN TR-SAP")

    
    Range("TR6_PARAMETRIZADA").AutoFilter Field:=VALIDACION_TIPO_REGIMEN_SALUD.Column, Criteria1:="REGISTRAR EPS", Operator:=xlOr, Criteria2:="=#N/D"
    
    'Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_HORARIO_NOCTURNO_SAP.Column, _
            Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"
    
    Range("TR6_PARAMETRIZADA[[NUMERO DOCUMENTO TR6]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("K10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    Range("TR6_PARAMETRIZADA[[TIPO DE REGIMEN SALUD TR]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("L10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Range("TR6_PARAMETRIZADA[[TIPO DE REGIMEN SALUD SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("M10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Range("TR6_PARAMETRIZADA[[VALIDACIÓN TIPO DE REGIMEN TR-SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("N10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
End With
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[[VALIDACIÓN TIPO DE REGIMEN TR-SAP]]").Select
ActiveSheet.ShowAllData

Range("A1").Select
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("K9").Select
ActiveCell.Value = DOC_TR.Value
Range("L9").Select
ActiveCell.Value = TIPO_REGIMEN_SALUD_TR.Value
Range("M9").Select
ActiveCell.Value = TIPO_REGIMEN_SALUD_SAP.Value
Range("N9").Select
ActiveCell.Value = VALIDACION_TIPO_REGIMEN_SALUD.Value

Set tbl = Range("K9").CurrentRegion
Set ws = ActiveSheet

ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_SALUD"


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


