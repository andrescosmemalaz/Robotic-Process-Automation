Attribute VB_Name = "MOVIMIENTO_REPORTE_SUELDO"
Sub MOVIMIENTO_SAP_REPORTE_SUELDO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Call LIMPIAR_TABLA_SUELDO
Call COLUMNA_CODIGO_PERSONAL
Call COLUMNA_IMPORTE

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub COLUMNA_CODIGO_PERSONAL()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_codigo_sap As String
Sheets("REPORTE_SUELDO_PARAMETRIZADA").Select
columnas = "A:B"
On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SUELDO_PARAMETRIZADA").Range(columnas)
            Set c = .Find("Número de personal")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_codigo_sap = Range(NOMBRE).Value
            Range("DATA_SUELDO[" + CStr(Valor_codigo_sap) + "]").Select
            
            Selection.Copy
            'Range("B10").Select
            Sheets("REPORTE_SUELDO_PARAMETRIZADA").Select
            Range("REPORTE_SUELDO_BUSCAR[Número de personal]").Select
            ActiveSheet.Paste
    End With
    
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
    
Exit Sub
ErrorHandler:
    Sheets("REPORTE_SUELDO_PARAMETRIZADA").Select
    Range("REPORTE_SUELDO_BUSCAR[Número de personal]").Select
    Range("REPORTE_SUELDO_BUSCAR[Número de personal]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub

Sub COLUMNA_IMPORTE()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_codigo_sap As String
Sheets("REPORTE_SUELDO_PARAMETRIZADA").Select
columnas = "A:AQ"
On Error GoTo ErrorHandler
    
    
    With Sheets("REPORTE_SUELDO_PARAMETRIZADA").Range(columnas)
            Set c = .Find("Importe")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_codigo_sap = Range(NOMBRE).Value
            Range("DATA_SUELDO[" + CStr(Valor_codigo_sap) + "]").Select
            
            Selection.Copy
            'Range("B10").Select
            Sheets("REPORTE_SUELDO_PARAMETRIZADA").Select
            Range("REPORTE_SUELDO_BUSCAR[Importe]").Select
            ActiveSheet.Paste
    End With
 
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

 
Exit Sub
ErrorHandler:
    Sheets("REPORTE_SUELDO_PARAMETRIZADA").Select
    Range("REPORTE_SUELDO_BUSCAR[Importe]").Select
    Range("REPORTE_SUELDO_BUSCAR[Importe]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub

Sub LIMPIAR_TABLA_SUELDO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Sheets("REPORTE_SUELDO_PARAMETRIZADA").Select
Range("REPORTE_SUELDO_BUSCAR[[Número de personal]:[Importe]]").Select
Selection.ClearContents

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

