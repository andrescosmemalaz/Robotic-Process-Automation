Attribute VB_Name = "FUNCION_SUELDO"

Sub AGREGAR_FUNCION_AL_SUELDO()
'' MODELADO_SUELDO Macro
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
    Range("BW2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP([@Codigo],REPORTE_SUELDO_BUSCAR[[Número de personal]:[Importe]],2,0)"
    Range("BW2").Select
    Selection.AutoFill Destination:=Range("DATA_SAP[SUELDO]")
    Range("DATA_SAP[SUELDO]").Select
    


Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

