Attribute VB_Name = "IMPORTACION_DATA_OTROS_LIBROS"
Sub IMPORTACION_DATA_SAP()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next

Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Worksheets("REPORTE_SAP").ListObjects("DATA_SAP_REPORTE").Delete

Workbooks.Open "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\SAP_REPORTES_MAESTRA.xlsm"
Dim NOMBRE As String
Dim Valor_codigo_sap As String
Workbooks("SAP_REPORTES_MAESTRA.xlsm").Worksheets("SAP").Activate
Sheets("SAP").Select
columnas = "A1:ZZ1"

Worksheets("SAP").ListObjects("DATA_SAP_REPORTE").Range.Copy Destination:=Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Range("A1")

Workbooks("SAP_REPORTES_MAESTRA.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("SAP_REPORTES_MAESTRA.xlsm").Close SaveChanges:=False
'MsgBox ("Se ha movido la data de SAP")

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub
Sub IMPORTACION_DATA_SUELDO()
On Error Resume Next
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Workbooks("SAP_REPORTES_SUELDOS.xlsm").Worksheets("REPORTE_SUELDOS").Activate
Worksheets("REPORTE_SUELDOS").ListObjects("DATA_SUELDO").Delete

On Error Resume Next
Workbooks.Open "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\SAP_REPORTES_SUELDOS.xlsm"
Dim NOMBRE As String
Dim Valor_codigo_sap As String
Workbooks("SAP_REPORTES_SUELDOS.xlsm").Worksheets("REPORTE SUELDO").Activate
Sheets("REPORTE SUELDO").Select
columnas = "A1:ZZ1"

Worksheets("REPORTE SUELDO").ListObjects("DATA_SUELDO").Range.Copy Destination:=Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SUELDOS").Range("A1")

Workbooks("SAP_REPORTES_SUELDOS.xlsm").Worksheets("REPORTE_SUELDOS").Activate
Workbooks("SAP_REPORTES_SUELDOS.xlsm").Close SaveChanges:=True

'MsgBox ("Se ha movido la data de SUELDO")

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

