Attribute VB_Name = "MODULO_PRINCIPAL"
Sub MACRO_PRINCIPAL_DEL_PROYECTO()
'Esta l�nea desactiva las alertas y mensajes que se muestran en Excel durante la ejecuci�n de la macro, para evitar interrupciones innecesarias.
Application.DisplayAlerts = False
' Esta l�nea llama a otra subrutina llamada "EXTRACION_DE_REPORTES_SAP_VALIDADOR"
Call EXTRACION_DE_REPORTES_SAP_VALIDADOR
'Esta l�nea llama a otra subrutina llamada "EXTRACION_DE_REPORTES_SUELDO_VALIDADOR
Call EXTRACION_DE_REPORTES_SUELDO_VALIDADOR
' Esta l�nea llama a otra subrutina llamada "ANALISIS_DOCUMENTOS_TRS
Call ANALISIS_DOCUMENTOS_TRS
'Esta l�nea llama a otra subrutina llamada "MOVIMIENTO_ARCHIVOS_CENTRAL
Call MOVIMIENTO_ARCHIVOS_CENTRAL
'Esta l�nea llama a otra subrutina llamada "PASO_FINAL", que presumiblemente realiza
Call PASO_FINAL
End Sub

Sub EXTRACION_DE_REPORTES_SAP_VALIDADOR()
'MACROS DEL PRIMER PASO
' CON ESTA MACRO SE ABRE AUTOMATICAMENTE EL REPORTE DE EXTRACION DE DATA MAESTRA , SE PEGA EN EL EXCEL Y SE CIERRA EL ARCHIVO

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
'Macros para llevar valores a otro archivo

Call TRANSPORTE_VALORES_A_OTRO_LIBRO_SAP
Workbooks("SAP_REPORTES_MAESTRA.xlsm").Close saveChanges:=False
Call Extraccion_data_sap_paralelo
'Call CIERRE_ARCHIVO_SAP
'Call Pegado_data_sap_paralelo

'MsgBox ("PROCESO FINALIZADO CORRECTAMENTE")
'MsgBox ("IR A PASO 2")

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
End Sub

Sub TRANSPORTE_VALORES_A_OTRO_LIBRO_SAP()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Dim origen, destino As Worksheet
Workbooks.Open ""

'
'HOJA DE CREDENCIALES
'
'MOVIMIENTO DE USUARIO
Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("CREDENCIALES SAP").Activate
    'Sheets("CREDENCIALES SAP").Activate
    Range("Selectusuario").Select
    Selection.Copy
Workbooks("SAP_REPORTES_MAESTRA.xlsm").Worksheets("CREDENCIALES SAP").Activate
    'Sheets("CREDENCIALES SAP").Activate
    Range("Selectusuario_Maestra").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    

'MOVIMIENTO DE PASSWORD
Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("CREDENCIALES SAP").Activate
    'Sheets("CREDENCIALES SAP").Activate
    Range("Selectedpassword").Select
    Selection.Copy
Workbooks("SAP_REPORTES_MAESTRA.xlsm").Worksheets("CREDENCIALES SAP").Activate
    'Sheets("CREDENCIALES SAP").Activate
    Range("Selectedpassword_Maestra").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
'MOVIMIENTO DE CELDA AMBIENTE
Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("CREDENCIALES SAP").Activate
    'Sheets("CREDENCIALES SAP").Activate
    Range("Environment").Select
    Selection.Copy
Workbooks("SAP_REPORTES_MAESTRA.xlsm").Worksheets("CREDENCIALES SAP").Activate
    'Sheets("CREDENCIALES SAP").Activate
    Range("Environment_Maestra").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
'HOJA DE TRANSACCI�N
'
    
'MOVIMIENTO DE FECHA INICIAL

Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("PRINCIPAL").Activate
    'Sheets("PRINCIPAL").Activate
    Range("FECHA_1").Select
    Selection.Copy
Workbooks("SAP_REPORTES_MAESTRA.xlsm").Worksheets("SAP").Activate
    'Sheets("SAP").Activate
    Range("SAP_FECHA1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False


'MOVIMIENTO DE FECHA FINAL

Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("PRINCIPAL").Activate
    'Sheets("PRINCIPAL").Activate
    Range("FECHA_2").Select
    Selection.Copy
Workbooks("SAP_REPORTES_MAESTRA.xlsm").Worksheets("SAP").Activate
    'Sheets("SAP").Activate
    Range("SAP_FECHA2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'MOVIMIENTO DE FECHA INICIAL

Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("PRINCIPAL").Activate
    'Sheets("PRINCIPAL").Activate
    Range("FECHA_3").Select
    Selection.Copy
Workbooks("SAP_REPORTES_MAESTRA.xlsm").Worksheets("SAP").Activate
    'Sheets("SAP").Activate
    Range("SAP_FECHA3").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'MOVIMIENTO DE FECHA FINAL

Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("PRINCIPAL").Activate
    'Sheets("PRINCIPAL").Activate
    Range("FECHA_4").Select
    Selection.Copy
Workbooks("SAP_REPORTES_MAESTRA.xlsm").Worksheets("SAP").Activate
    'Sheets("SAP").Activate
    Range("SAP_FECHA4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
            
   
'MOVIMIENTO DE CELDA UNIDAD
   
Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("PRINCIPAL").Activate
    'Sheets("PRINCIPAL").Activate
    Range("CELDA_UNIDAD_SELECCIONADA").Select
    Selection.Copy
Workbooks("SAP_REPORTES.xlsm").Worksheets("SAP").Activate
    'Sheets("SAP").Activate
    Range("CELDA_UNIDAD_SELECCIONADA").Select
    ActiveSheet.Paste

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
    
End Sub

Sub Extraccion_data_sap_paralelo()
Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

LIBRO = ""
ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
ExcelApp.Run "SAP_extract_DataMaestra_Reporte"
'ExcelApp.Wait (Now + TimeValue("0:00:03"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub CIERRE_ARCHIVO_SAP()


Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

LIBRO = "C:\SAP_REPORTES_MAESTRA.xlsm"
ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
ExcelApp.Run "cerrar_archivo_export_maestra"
'ExcelApp.Wait (Now + TimeValue("0:00:03"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub



Sub Pegado_data_sap_paralelo()
On Error Resume Next

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application

LIBRO = "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\SAP_REPORTES_MAESTRA.xlsm"

ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
ExcelApp.Run "PEGAR_EXCEL"
'ExcelApp.Wait (Now + TimeValue("0:00:05"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub EXTRACION_DE_REPORTES_SUELDO_VALIDADOR()

'MACROS DEL SEGUNDO  PASO
' CON ESTA MACRO SE ABRE AUTOMATICAMENTE EL REPORTE DE EXTRACION DE DATA SUELDOS , SE PEGA EN EL EXCEL Y SE CIERRA EL ARCHIVO
Application.DisplayAlerts = False

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Call TRANSPORTE_VALORES_A_OTRO_LIBRO_SUELDO
Workbooks("SAP_REPORTES_SUELDOS.xlsm").Close saveChanges:=False
Call Extraccion_data_sueldo_paralelo
'Call CIERRE_ARCHIVO_SUELDO
'Call Pegado_data_sueldo_paralelo


'MsgBox ("PROCESO FINALIZADO CORRECTAMENTE")
'MsgBox ("IR A PASO 3")

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub TRANSPORTE_VALORES_A_OTRO_LIBRO_SUELDO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
On Error Resume Next
Dim origen, destino As Worksheet
Workbooks.Open "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\SAP_REPORTES_SUELDOS.xlsm"

'MOVIMIENTO DE CREDENCIALES_USUARIO
Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("CREDENCIALES SAP").Activate
    'Sheets("CREDENCIALES SAP").Activate
    Range("Selectusuario").Select
    Selection.Copy
Workbooks("SAP_REPORTES_SUELDOS.xlsm").Worksheets("CREDENCIALES SAP").Activate
    'Sheets("CREDENCIALES SAP").Activate
    Range("Selectusuario_Sueldo").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
'MOVIMIENTO DE CREDENCIALES_PASSWORD
Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("CREDENCIALES SAP").Activate
    'Sheets("CREDENCIALES SAP").Activate
    Range("Selectedpassword").Select
    Selection.Copy
Workbooks("SAP_REPORTES_SUELDOS.xlsm").Worksheets("CREDENCIALES SAP").Activate
    'Sheets("CREDENCIALES SAP").Select
    Range("Selectedpassword_Sueldo").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'MOVIMIENTO DE AMBIENTE
Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("CREDENCIALES SAP").Activate
    'Sheets("CREDENCIALES SAP").Select
    Range("Environment").Select
    Selection.Copy
Workbooks("SAP_REPORTES_SUELDOS.xlsm").Worksheets("CREDENCIALES SAP").Activate
    'Sheets("CREDENCIALES SAP").Select
    Range("Environment_Sueldo").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'MOVIMIENTO DE FECHA_1
Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("PRINCIPAL").Activate
    'Sheets("PRINCIPAL").Select
    Range("FECHA_1").Select
    Selection.Copy
Workbooks("SAP_REPORTES_SUELDOS.xlsm").Worksheets("REPORTE SUELDO").Activate
    'Sheets("REPORTE SUELDO").Select
    Range("SUELDO_FECHA1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'MOVIMIENTO DE FECHA_2
Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("PRINCIPAL").Activate
    'Sheets("PRINCIPAL").Select
    Range("FECHA_2").Select
    Selection.Copy
Workbooks("SAP_REPORTES_SUELDOS.xlsm").Worksheets("REPORTE SUELDO").Activate
    'Sheets("REPORTE SUELDO").Select
    Range("SUELDO_FECHA2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'MOVIMIENTO DE FECHA_3
Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("PRINCIPAL").Activate
    'Sheets("PRINCIPAL").Select
    Range("FECHA_3").Select
    Selection.Copy
Workbooks("SAP_REPORTES_SUELDOS.xlsm").Worksheets("REPORTE SUELDO").Activate
    'Sheets("REPORTE SUELDO").Select
    Range("SUELDO_FECHA3").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

'MOVIMIENTO DE FECHA_4
Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("PRINCIPAL").Activate
    'Sheets("PRINCIPAL").Select
    Range("FECHA_4").Select
    Selection.Copy
Workbooks("SAP_REPORTES_SUELDOS.xlsm").Worksheets("REPORTE SUELDO").Activate
    'Sheets("REPORTE SUELDO").Select
    Range("SUELDO_FECHA4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
        
   
'MOVIMIENTO DE CELDA_UNIDAD
Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Worksheets("PRINCIPAL").Activate
    'Sheets("PRINCIPAL").Select
    Range("CELDA_UNIDAD_SELECCIONADA").Select
    Selection.Copy
Workbooks("SAP_REPORTES_SUELDOS.xlsm").Worksheets("REPORTE SUELDO").Activate
    'Sheets("REPORTE SUELDO").Select
    Range("CELDA_UNIDAD_SELECCIONADA").Select
    ActiveSheet.Paste

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub

Sub Extraccion_data_sueldo_paralelo()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application


LIBRO = "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\SAP_REPORTES_SUELDOS.xlsm"

ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
ExcelApp.Run "SAP_extract_SUELDO"
ExcelApp.Wait (Now + TimeValue("0:00:01"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub CIERRE_ARCHIVO_SUELDO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application

LIBRO = "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\SAP_REPORTES_SUELDOS.xlsm"

ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
ExcelApp.Run "cerrar_archivo_export_sueldo"
'ExcelApp.Wait (Now + TimeValue("0:00:30"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


Sub Pegado_data_sueldo_paralelo()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application

LIBRO = "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\SAP_REPORTES_SUELDOS.xlsm"

ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
ExcelApp.Run "PEGAR_DATA_SUELDO"
'ExcelApp.Wait (Now + TimeValue("0:00:30"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


Sub ANALISIS_DOCUMENTOS_TRS()

'MACROS DEL TERCER  PASO
' CON ESTA MACRO SE ABRE AUTOMATICAMENTE EL REPORTE DE EXTRACION DE TR�S, SE PEGA EN EL EXCEL Y SE CIERRA EL ARCHIVO

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Call Extraccion_TRS_validador_part1
Call Extraccion_TRS_validador_part2
Call Extraccion_TRS_validador_part3

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


Sub Extraccion_TRS_validador_part1()

Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application
'Movimiento de desarrollo
'Movimiento caso plus

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


LIBRO = "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\Extractor Documentos TR�s.xlsm"

ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
ExcelApp.Run "Modelo_TR5_ARCHIVO"
'ExcelApp.Run "Modelo_TR6_ARCHIVO"
'ExcelApp.Run "MOVIMIENTO_TRS_VALIDADOR"
'ExcelApp.Wait (Now + TimeValue("0:00:01"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub Extraccion_TRS_validador_part2()

'Movimiento de desarrollo
'Movimiento caso plus

Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False



LIBRO = "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\Extractor Documentos TR�s.xlsm"

ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
'ExcelApp.Run "Modelo_TR5_ARCHIVO"
ExcelApp.Run "Modelo_TR6_ARCHIVO"
'ExcelApp.Run "MOVIMIENTO_TRS_VALIDADOR"
ExcelApp.Wait (Now + TimeValue("0:00:01"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub Extraccion_TRS_validador_part3()

'Movimiento de desarrollo
'Movimiento caso plus
Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


LIBRO = "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\Extractor Documentos TR�s.xlsm"

ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
'ExcelApp.Run "Modelo_TR5_ARCHIVO"
'ExcelApp.Run "Modelo_TR6_ARCHIVO"
ExcelApp.Run "MOVIMIENTO_TRS_VALIDADOR"
'ExcelApp.Wait (Now + TimeValue("0:00:01"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub MOVIMIENTO_ARCHIVOS_CENTRAL()

'MACROS DEL CUARTO  PASO
' CON ESTA MACRO SE ABRE AUTOMATICAMENTE EL VALIDADOR Y LOS ARCHIVOS CON LA DATA MAESTRA, SUELDOS Y DE TR�S, SE PEGA EN EL EXCEL DEL VALIDADOR Y SE CIERRA EL ARCHIVO

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


Call IMPORTACION_GENERAL
'Call MOVIMIENTO_ARCHIVOS_SAP

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


Sub IMPORTACION_GENERAL()

Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


LIBRO = "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\CENTRAL_DATA_SAP.xlsm"
ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
ExcelApp.Run "IMPORTACION_DATA_GENERAL"
ExcelApp.Run "MOVIMIENTO_DATA_GRANDE"
'ExcelApp.Wait (Now + TimeValue("0:00:10"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing
'MsgBox ("PROCESO FINALIZADO CORRECTAMENTE")
'MsgBox ("IR A PASO 4")


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub

Sub MOVIMIENTO_ARCHIVOS_SAP()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application


LIBRO = "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\CENTRAL_DATA_SAP.xlsm"
ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
'ExcelApp.Run "IMPORTACION_DATA_GENERAL"
ExcelApp.Run "MOVIMIENTO_DATA_GRANDE"
'ExcelApp.Wait (Now + TimeValue("0:00:10"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing
'MsgBox ("PROCESO FINALIZADO CORRECTAMENTE")
'MsgBox ("IR A PASO 5")
'
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False



End Sub


Sub PASO_FINAL()

'MACROS DEL QUINTO  PASO
' CON ESTA MACRO SE ABRE AUTOMATICAMENTE EL VALIDADOR Y SE EJECUTA LAS MACROS PARA PROCESAR LOS EXCEL , VALIDAR INFORMACI�N Y GENERAR UN REPORTE CON ERRORES, TARER EL REPORTE Y CERRAR EL VALIDADOR

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Call VALIDACION
Call EXTRACION_DOCUMENTO

MsgBox ("PROCESO FINALIZADO CORRECTAMENTE")


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


Sub VALIDACION()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application



LIBRO = "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\PROCESO_VALIDACION.xlsm"

ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
ExcelApp.Run "AUTOMATION_TR"
'ExcelApp.Run "PEGADO_REPORTE_FINAL"
'ExcelApp.Wait (Now + TimeValue("0:00:10"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing
'Call EXPULSION_DOCUMENTO_FINAL




Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub EXTRACION_DOCUMENTO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

     Set closedBook = Workbooks.Open("C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\PROCESO_VALIDACION.xlsm")
    'closedBook.Sheets("REPORTE").Copy Before:=ThisWorkbook.Sheets(3)
    closedBook.Sheets("REPORTE").Copy After:=Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Sheets(Workbooks("PRINCIPAL_REPORTE_VALIDACION.xlsm").Sheets.Count)
    closedBook.Close saveChanges:=False

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


Sub Extraccion_data_validador()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim ExcelApp  As Excel.Application
Set ExcelApp = New Excel.Application


LIBRO = "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\REPORTES_VALIDACION_SAP_TRs.xlsm"

ExcelApp.Workbooks.Open (LIBRO)
ExcelApp.Visible = True
ExcelApp.Run "MOVIMIENTO_DATA_GRANDE"
'ExcelApp.Wait (Now + TimeValue("0:00:10"))
ExcelApp.ActiveWorkbook.Save
ExcelApp.ActiveWorkbook.Close (0)
ExcelApp.Quit
Set ExcelApp = Nothing


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub














