Attribute VB_Name = "MODULOS_MOVIMIENTO_PRINCIPALES"
Sub MOVIMIENTO_DATA_GRANDE()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Workbooks.Open "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\PROCESO_VALIDACION.xlsm"
Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Call EXPORT_DATA_SAP
Call MOVIMIENTO_SAP_REPORTE_SUELDO
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Workbooks("PROCESO_VALIDACION.xlsm").Save
Workbooks("PROCESO_VALIDACION.xlsm").Close SaveChanges:=False

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub

Sub IMPORTACION_DATA_GENERAL()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Workbooks.Open "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\SAP_REPORTES_MAESTRA.xlsm"
Workbooks.Open "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\SAP_REPORTES_SUELDOS.xlsm"

'Set DETALLE = Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("SAP_PARAMETRIZADA")
Call ELIMINAR_TABLAS_PREVIA_IMPORTANCION
Call IMPORTACION_DATA_SAP
Call IMPORTACION_DATA_SUELDO
Workbooks("CENTRAL_DATA_SAP.xlsm").Save

'Workbooks("SAP_REPORTES_MAESTRA.xlsm").Close SaveChanges:=False
'Workbooks("SAP_REPORTES_SUELDOS.xlsm").Close SaveChanges:=False

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub ELIMINAR_TABLAS_PREVIA_IMPORTANCION()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Workbooks.Open "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\CENTRAL_DATA_SAP.xlsm"

    Dim Tabla1 As ListObject
    Dim Tabla2 As ListObject
    'Dim ws_sueldos As Worksheets
    'Dim ws_maestra As Worksheets
    
    'Asigna las tablas a las variables correspondientes'
    Set Tabla1 = Worksheets("REPORTE_SUELDOS").ListObjects("DATA_SUELDO")
    Set Tabla2 = Worksheets("REPORTE_SAP").ListObjects("DATA_SAP_REPORTE")
    'Elimina las tablas'
    Tabla1.Delete
    Tabla2.Delete
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
