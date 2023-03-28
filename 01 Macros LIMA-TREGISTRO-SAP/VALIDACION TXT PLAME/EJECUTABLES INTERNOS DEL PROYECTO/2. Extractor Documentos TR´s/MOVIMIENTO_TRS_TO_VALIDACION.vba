Attribute VB_Name = "MOVIMIENTO_TRS_TO_VALIDACION"
Sub MOVIMIENTO_TR6()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Call LIMPIAR_TABLA_TR6
Call COLUMNA_TIPO
Call COLUMNA_NUMERO
Call COLUMNA_APELLIDO_PATERNO
Call COLUMNA_APELLIDO_MATERNO
Call COLUMNA_NOMBRES
Call COLUMNA_SITUACION_TRABAJADOR
Call COLUMNA_TIPO_DE_TRABAJADOR
Call COLUMNA_REGIMEN_DEL_SALUD
Call COLUMNA_FECHA_INICIO
Call COLUMNA_EPS_SERV_PROPIO
Call COLUMNA_TIPO_DE_REGIMEN
Call COLUMNA_FECHA_INICIO_2
Call COLUMNA_CUSSP
Call COLUMNA_NINGUNO
Call COLUMNA_ESSALUD
Call COLUMNA_EPS
Call COLUMNA_NINGUNA_2
Call COLUMNA_ONP
Call COLUMNA_SEGURO_PRIVADO
Call COLUMNA_IND_RENTAS_LIR
Call COLUMNA_CONV_EVIT_DOBLE_IMP

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub

Sub COLUMNA_TIPO()

Application.DisplayAlerts = False

Dim NOMBRE As String
Dim Valor_codigo_personal As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"


On Error GoTo ErrorHandler
    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("Tipo")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_codigo_personal = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_codigo_personal) + "]").Select
            Selection.Copy
            
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Tipo]").Select
            ActiveSheet.Paste
    End With
    
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
    Sheets("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA[Tipo]").Select
    Range("TR6_PARAMETRIZADA[Tipo]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub

Sub COLUMNA_NUMERO()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")


Dim NOMBRE As String
Dim Valor_numero As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler

    With Sheets("TR6").Range(columnas)
            Set c = .Find("Número")
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_codigo_personal = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_codigo_personal) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Número]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
    Sheets("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA[Número]").Select
    Range("TR6_PARAMETRIZADA[Número]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_APELLIDO_PATERNO()


Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")


Dim NOMBRE As String
Dim Valor_apellido_paterno As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"

On Error GoTo ErrorHandler
    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("Apellido paterno")
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_apellido_paterno = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_apellido_paterno) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Apellido paterno]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
    Sheets("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA[Apellido paterno]").Select
    Range("TR6_PARAMETRIZADA[Apellido paterno]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_APELLIDO_MATERNO()


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Application.DisplayAlerts = False

Dim NOMBRE As String
Dim Valor_apellido_materno As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler


    With Sheets("TR6").Range(columnas)
            Set c = .Find("Apellido materno")
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_apellido_materno = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_apellido_materno) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Apellido materno]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
    Sheets("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA[Apellido materno]").Select
    Range("TR6_PARAMETRIZADA[Apellido materno]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_NOMBRES()


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")
Application.DisplayAlerts = False

Dim NOMBRE As String
Dim Valor_nombre As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"


On Error GoTo ErrorHandler


    With Sheets("TR6").Range(columnas)
            Set c = .Find("Nombres")
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_nombre = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_nombre) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Nombres]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
    Sheets("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA[Nombres]").Select
    Range("TR6_PARAMETRIZADA[Nombres]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_SITUACION_TRABAJADOR()


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Application.DisplayAlerts = False
Dim NOMBRE As String
Dim Valor_situacion_trabajador As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler


    With Sheets("TR6").Range(columnas)
            Set c = .Find("Situación")
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_situacion_trabajador = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_situacion_trabajador) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Situación]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
    Sheets("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA[Situación]").Select
    Range("TR6_PARAMETRIZADA[Situación]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_TIPO_DE_TRABAJADOR()
Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")


Dim NOMBRE As String
Dim Valor_tipo_de_trabajador As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler
    
    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("Tipo de Trabajador")
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_tipo_de_trabajador = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_tipo_de_trabajador) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Tipo de Trabajador]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
    Sheets("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA[Tipo de Trabajador]").Select
    Range("TR6_PARAMETRIZADA[Tipo de Trabajador]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_REGIMEN_DEL_SALUD()

Application.DisplayAlerts = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")


Dim NOMBRE As String
Dim Valor_regimen_del_salud As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler
    
    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("Tipo de Régimen")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_regimen_del_salud = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_regimen_del_salud) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Tipo de Régimen]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
    Sheets("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA[Tipo de Régimen]").Select
    Range("TR6_PARAMETRIZADA[Tipo de Régimen]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_FECHA_INICIO()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Application.DisplayAlerts = False
Dim NOMBRE As String
Dim Valor_fecha_inicio As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"

On Error GoTo ErrorHandler
    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("Fecha Inicio")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_fecha_inicio = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_fecha_inicio) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Fecha Inicio]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
    Sheets("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA[Fecha Inicio]").Select
    Range("TR6_PARAMETRIZADA[Fecha Inicio]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_EPS_SERV_PROPIO()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Application.DisplayAlerts = False
Dim NOMBRE As String
Dim Valor_eps_servi_propio As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler
    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("EPS/Serv. Propio")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_eps_servi_propio = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_eps_servi_propio) + "]").Select
            
            
            Selection.Copy
            
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[EPS/Serv. Propio]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
    Sheets("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA[EPS/Serv. Propio]").Select
    Range("TR6_PARAMETRIZADA[EPS/Serv. Propio]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_TIPO_DE_REGIMEN()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")
Application.DisplayAlerts = False
Dim NOMBRE As String
Dim Valor_tipo_regimen As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"


On Error GoTo ErrorHandler

    With Sheets("TR6").Range(columnas)
            Set c = .Find("Tipo de Régimen2")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_tipo_regimen = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_tipo_regimen) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Tipo de Régimen2]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
    Sheets("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA[Tipo de Régimen2]").Select
    Range("TR6_PARAMETRIZADA[Tipo de Régimen2]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_FECHA_INICIO_2()
Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_fecha_inicio As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"


On Error GoTo ErrorHandler
    
    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("Fecha Inicio3")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_fecha_inicio = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_fecha_inicio) + "]").Select
            
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Fecha Inicio3]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
    Sheets("TR6_PARAMETRIZADA").Select
    Range("TR6_PARAMETRIZADA[Fecha Inicio3]").Select
    Range("TR6_PARAMETRIZADA[Fecha Inicio3]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_CUSSP()

Application.DisplayAlerts = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_cussp As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler
    
    
    With Sheets("TR6").Range(columnas)
              Set c = .Find("CUSPP")
              
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_cussp = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_cussp) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[CUSPP]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[CUSPP]").Select
Range("TR6_PARAMETRIZADA[CUSPP]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_NINGUNO()
Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_ninguno As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler
    
    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("Ninguno")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_ninguno = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_ninguno) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Ninguno]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[Ninguno]").Select
Range("TR6_PARAMETRIZADA[Ninguno]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_ESSALUD()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_essalud As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler

    With Sheets("TR6").Range(columnas)
            Set c = .Find("ESSALUD")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_essalud = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_essalud) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[ESSALUD]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[ESSALUD]").Select
Range("TR6_PARAMETRIZADA[ESSALUD]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_EPS()

Application.DisplayAlerts = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_eps As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"

On Error GoTo ErrorHandler
    
    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("EPS")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_eps = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_eps) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[EPS]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[EPS]").Select
Range("TR6_PARAMETRIZADA[EPS]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_NINGUNA_2()

Application.DisplayAlerts = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_ninguna_2 As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler
    

    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("Ninguno4")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_ninguna_2 = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_ninguna_2) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Ninguno4]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[Ninguno4]").Select
Range("TR6_PARAMETRIZADA[Ninguno4]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_ONP()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_onp As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"


On Error GoTo ErrorHandler
    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("ONP")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_onp = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_onp) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[ONP]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[ONP]").Select
Range("TR6_PARAMETRIZADA[ONP]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_SEGURO_PRIVADO()
Application.DisplayAlerts = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_seguro_privado As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler

    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("Seg. Privado")
        
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_seguro_privado = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_seguro_privado) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Seg. Privado]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[Seg. Privado]").Select
Range("TR6_PARAMETRIZADA[Seg. Privado]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_IND_RENTAS_LIR()
Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_ind_rentas_lir As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler
    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("Ind. rentas LIR")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_ind_rentas_lir = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_ind_rentas_lir) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Ind. rentas LIR]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[Ind. rentas LIR]").Select
Range("TR6_PARAMETRIZADA[Ind. rentas LIR]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_CONV_EVIT_DOBLE_IMP()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_conv_evit_doble_imp As String
ThisWorkbook.Sheets("TR6").Activate
Sheets("TR6").Select
columnas = "A:W"



On Error GoTo ErrorHandler
    
    
    With Sheets("TR6").Range(columnas)
            Set c = .Find("Conv. evit. doble imp.")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_conv_evit_doble_imp = Range(NOMBRE).Value
            Range("MODELO_TR6[" + CStr(Valor_conv_evit_doble_imp) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR6_PARAMETRIZADA").Select
            Range("TR6_PARAMETRIZADA[Conv. evit. doble imp.]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[Conv. evit. doble imp.]").Select
Range("TR6_PARAMETRIZADA[Conv. evit. doble imp.]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub

Sub LIMPIAR_TABLA_TR6()
Application.DisplayAlerts = False

Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA").Activate
Sheets("TR6_PARAMETRIZADA").Select
Range("TR6_PARAMETRIZADA[[Tipo]:[Conv. evit. doble imp.]]").Select
Selection.ClearContents


End Sub

Sub MOVIMIENTO_TR5()
Application.DisplayAlerts = False

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Call LIMPIAR_TABLA_TR5
Call COLUMNA_TIPO_TR5
Call COLUMNA_NUMERO_TR5
Call COLUMNA_APELLIDO_PATERNO_TR5
Call COLUMNA_APELLIDO_MATERNO_TR5
Call COLUMNA_NOMBRES_TR5
Call COLUMNA_FECHA_INICIO_TR5
Call COLUMNA_TIPO_TRABAJADOR_TR5
Call COLUMNA_REGIMEN_LABORAL_TR5
Call COLUMNA_CATEGORIA_OCUPACIONAL_TR5
Call COLUMNA_OCUPACION_TR5
Call COLUMNA_NIVEL_EDUCATIVO_TR5
Call COLUMNA_DISCAPACIDAD_TR5
Call COLUMNA_SINDICALIZADO_TR5
Call COLUMNA_REGIMEN_ACUMULATIVO_TR5
Call COLUMNA_MAXIMA_TR5
Call COLUMNA_HORARIO_NOCTURNO_TR5
Call COLUMNA_SITUACION_ESPECIAL_TRABAJADOR_TR5
Call COLUMNA_ESTABLECIMIENTO_TR5
Call COLUMNA_TIPO_DE_CONTRATO_TR5
Call COLUMNA_TIPO_DE_PAGO_TR5
Call COLUMNA_PERIODICIDAD_TR5
Call COLUMNA_ENTIDAD_FINANCIERA_TR5
Call COLUMNA_NRO_CUENTA_TR5
Call COLUMNA_REMUNERACION_BASICA_TR5
Call COLUMNA_SITUACION_TR5

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub



Sub COLUMNA_TIPO_TR5()


Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")


Dim NOMBRE As String
Dim Valor_tipo_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler
    
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Tipo")
            
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_tipo_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_tipo_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Tipo]").Select
            ActiveSheet.Paste
    End With
    
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
    Sheets("TR5_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA[Tipo]").Select
    Range("TR5_PARAMETRIZADA[Tipo]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_NUMERO_TR5()
Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_numero_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Número")
             
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_numero_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_numero_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Número]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
    Sheets("TR5_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA[Número]").Select
    Range("TR5_PARAMETRIZADA[Número]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_APELLIDO_PATERNO_TR5()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_apellido_paterno_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler

    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Apellido paterno")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_apellido_paterno_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_apellido_paterno_tr5) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Apellido paterno]").Select
            ActiveSheet.Paste
    End With
    
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
    Sheets("TR5_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA[Apellido paterno]").Select
    Range("TR5_PARAMETRIZADA[Apellido paterno]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_APELLIDO_MATERNO_TR5()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_apellido_materno_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler

    With Sheets("TR5").Range(columnas)
            Set c = .Find("Apellido materno")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_apellido_materno_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_apellido_materno_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Apellido materno]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
    Sheets("TR5_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA[Apellido materno]").Select
    Range("TR5_PARAMETRIZADA[Apellido materno]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_NOMBRES_TR5()
Application.DisplayAlerts = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_nombres_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler
    
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Nombres")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_apellido_materno_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_apellido_materno_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Nombres]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
    Sheets("TR5_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA[Nombres]").Select
    Range("TR5_PARAMETRIZADA[Nombres]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_FECHA_INICIO_TR5()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_fecha_inicio_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler
    
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Fec. Inicio")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_fecha_inicio_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_fecha_inicio_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Fec. Inicio]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
    Sheets("TR5_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA[Fec. Inicio]").Select
    Range("TR5_PARAMETRIZADA[Fec. Inicio]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_TIPO_TRABAJADOR_TR5()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_tipo_trabajador_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"
Application.DisplayAlerts = False

On Error GoTo ErrorHandler
    
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Tipo de Trabajador")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_tipo_trabajador_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_tipo_trabajador_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Tipo de Trabajador]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
    Sheets("TR5_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA[Tipo de Trabajador]").Select
    Range("TR5_PARAMETRIZADA[Tipo de Trabajador]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_REGIMEN_LABORAL_TR5()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_regimen_laboral_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"



On Error GoTo ErrorHandler
    
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Régimen Laboral")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_regimen_laboral_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_regimen_laboral_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Régimen Laboral]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
    Sheets("TR5_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA[Régimen Laboral]").Select
    Range("TR5_PARAMETRIZADA[Régimen Laboral]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_CATEGORIA_OCUPACIONAL_TR5()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_categoria_ocupacional_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler
    
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Cat. Ocupacional")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_categoria_ocupacional_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_categoria_ocupacional_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Cat. Ocupacional]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
    Sheets("TR5_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA[Cat. Ocupacional]").Select
    Range("TR5_PARAMETRIZADA[Cat. Ocupacional]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_OCUPACION_TR5()
Application.DisplayAlerts = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_ocupacion_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"

On Error GoTo ErrorHandler
    
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Ocupación")
            
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            
            Valor_ocupacion_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_ocupacion_tr5) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Ocupación]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
    Sheets("TR5_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA[Ocupación]").Select
    Range("TR5_PARAMETRIZADA[Ocupación]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_NIVEL_EDUCATIVO_TR5()
Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_nivel_educativo_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"



On Error GoTo ErrorHandler
    
    
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Nivel Educativo")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_nivel_educativo_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_nivel_educativo_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Nivel Educativo]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
    Sheets("TR5_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA[Nivel Educativo]").Select
    Range("TR5_PARAMETRIZADA[Nivel Educativo]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_DISCAPACIDAD_TR5()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_discapacidad_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler
    
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Discapacidad")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_discapacidad_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_discapacidad_tr5) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Discapacidad]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
    Sheets("TR5_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA[Discapacidad]").Select
    Range("TR5_PARAMETRIZADA[Discapacidad]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_SINDICALIZADO_TR5()

Application.DisplayAlerts = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_sindicalizado_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler

    
    With Sheets("TR5").Range(columnas)
           Set c = .Find("Sindicalizado")
           
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_sindicalizado_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_sindicalizado_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Sindicalizado]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Sindicalizado]").Select
Range("TR5_PARAMETRIZADA[Sindicalizado]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_REGIMEN_ACUMULATIVO_TR5()

Application.DisplayAlerts = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_regimen_acumulativo_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler


    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Reg. Acumulativo")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_regimen_acumulativo_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_regimen_acumulativo_tr5) + "]").Select
            
        
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Reg. Acumulativo]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Reg. Acumulativo]").Select
Range("TR5_PARAMETRIZADA[Reg. Acumulativo]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_MAXIMA_TR5()

Application.DisplayAlerts = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_maxima_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"



On Error GoTo ErrorHandler
    
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Máxima")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_maxima_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_maxima_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Máxima]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Máxima]").Select
Range("TR5_PARAMETRIZADA[Máxima]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_HORARIO_NOCTURNO_TR5()


Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_horario_nocturno_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler
    
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Horario Nocturno")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_horario_nocturno_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_horario_nocturno_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Horario Nocturno]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Horario Nocturno]").Select
Range("TR5_PARAMETRIZADA[Horario Nocturno]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_SITUACION_ESPECIAL_TRABAJADOR_TR5()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_situacion_especial_trabajador_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"

On Error GoTo ErrorHandler
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Situación Especial del Trabajador")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_situacion_especial_trabajador_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_situacion_especial_trabajador_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Situación Especial del Trabajador]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Situación Especial del Trabajador]").Select
Range("TR5_PARAMETRIZADA[Situación Especial del Trabajador]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_ESTABLECIMIENTO_TR5()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")
Application.DisplayAlerts = False

Dim NOMBRE As String
Dim Valor_establecimiento_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"

On Error GoTo ErrorHandler

    With Sheets("TR5").Range(columnas)
            Set c = .Find("Establecimiento en el que labora")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_establecimiento_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_establecimiento_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Establecimiento en el que labora]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Establecimiento en el que labora]").Select
Range("TR5_PARAMETRIZADA[Establecimiento en el que labora]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_TIPO_DE_CONTRATO_TR5()

Application.DisplayAlerts = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_tipo_de_contrato_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler

    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Tipo de Contrato")
           
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_tipo_de_contrato_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_tipo_de_contrato_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Tipo de Contrato]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Tipo de Contrato]").Select
Range("TR5_PARAMETRIZADA[Tipo de Contrato]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_TIPO_DE_PAGO_TR5()

Application.DisplayAlerts = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_tipo_de_pago_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Tipo de pago")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_tipo_de_pago_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_tipo_de_pago_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Tipo de pago]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Tipo de pago]").Select
Range("TR5_PARAMETRIZADA[Tipo de pago]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_PERIODICIDAD_TR5()

Application.DisplayAlerts = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_periodicidad_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"


On Error GoTo ErrorHandler

    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Periodicidad")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_periodicidad_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_periodicidad_tr5) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Periodicidad]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Periodicidad]").Select
Range("TR5_PARAMETRIZADA[Periodicidad]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub

Sub COLUMNA_ENTIDAD_FINANCIERA_TR5()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")
Application.DisplayAlerts = False

Dim NOMBRE As String
Dim Valor_entidad_financiera_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"

On Error GoTo ErrorHandler
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Entidad Financiera")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_entidad_financiera_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_entidad_financiera_tr5) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Entidad Financiera]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Entidad Financiera]").Select
Range("TR5_PARAMETRIZADA[Entidad Financiera]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_NRO_CUENTA_TR5()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_numero_cuenta_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "A:W"

On Error GoTo ErrorHandler
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Nro de Cuenta")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_numero_cuenta_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_numero_cuenta_tr5) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Nro de Cuenta]").Select
            ActiveSheet.Paste
            Range("TR5_PARAMETRIZADA[Nro de Cuenta]").Select
            Selection.NumberFormat = "0.00"
            Selection.NumberFormat = "0.0"
            Selection.NumberFormat = "0"
            
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Nro de Cuenta]").Select
Range("TR5_PARAMETRIZADA[Nro de Cuenta]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_REMUNERACION_BASICA_TR5()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_remuneracion_basica_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "X:Y"


On Error GoTo ErrorHandler
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Remun Bas.")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_remuneracion_basica_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_remuneracion_basica_tr5) + "]").Select
            Set c = .Find("Nro de Cuenta")
            
            Selection.Copy
            
            
            
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Remun Bas.]").Select
            ActiveSheet.Paste
            Range("TR5_PARAMETRIZADA[Remun Bas.]").Select
            Selection.NumberFormat = "0.00"
            Selection.NumberFormat = "0.0"
            Selection.NumberFormat = "0"
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Remun Bas.]").Select
Range("TR5_PARAMETRIZADA[Remun Bas.]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_SITUACION_TR5()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_situacion_tr5 As String
ThisWorkbook.Sheets("TR5").Activate
Sheets("TR5").Select
columnas = "X:Y"


On Error GoTo ErrorHandler
    
    With Sheets("TR5").Range(columnas)
            Set c = .Find("Situación")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_situacion_tr5 = Range(NOMBRE).Value
            Range("HOJA_TR5[" + CStr(Valor_situacion_tr5) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("TR5_PARAMETRIZADA").Select
            Range("TR5_PARAMETRIZADA[Situación]").Select
            ActiveSheet.Paste
    End With
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[Situación]").Select
Range("TR5_PARAMETRIZADA[Situación]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub

Sub LIMPIAR_TABLA_TR5()
Application.DisplayAlerts = False

Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR5_PARAMETRIZADA").Activate
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[Tipo]:[Situación]]").Select
Selection.ClearContents

End Sub




Sub MOVIMIENTO_TRS_VALIDADOR()
Application.DisplayAlerts = False

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Workbooks.Open "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\PROCESO_VALIDACION.xlsm"
Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("TR6_PARAMETRIZADA")
Call MOVIMIENTO_TR5
Call MOVIMIENTO_TR6

Workbooks("PROCESO_VALIDACION.xlsm").Close SaveChanges:=True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
'Application.Quit

End Sub





