Attribute VB_Name = "EXPORTACION_DATA_MAESTRA"
Sub EXPORT_DATA_SAP()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Call LIMPIAR_TABLA_SAP
Call COLUMNA_CODIGO
Call COLUMNA_TIPO_DOCUMENTO
Call COLUMNA_NUMERO_DOCUMENTO
Call COLUMNA_NOMBRE
Call COLUMNA_APELLIDO_PATERNO
Call COLUMNA_APELLIDO_MATERNO
Call COLUMNA_DESCRIPCION_TIPO
Call COLUMNA_EMPRESA
Call COLUMNA_DIVISION
Call COLUMNA_AREA
Call COLUMNA_DEPARTAMENTO
Call COLUMNA_CARGO
Call COLUMNA_NOMBRE_CARGO
Call COLUMNA_FECHA_NACIMIENTO
Call COLUMNA_FECHA_INGRESO
Call COLUMNA_FECHA_CESE
Call COLUMNA_MOTIVO_CESE
Call COLUMNA_RELACION_LABORAL
Call COLUMNA_TIPO_TRABAJADOR
Call COLUMNA_JORNADA_TRAB
Call COLUMNA_TIPO_DE_CONTRATO
Call COLUMNA_INICIO_DE_CONTRATO
Call COLUMNA_FIN_DE_CONTRATO
Call COLUMNA_SINDICATO
Call COLUMNA_LOCACION
Call COLUMNA_EPS_SITUACION
Call COLUMNA_TIPO_PLAN
Call COLUMNA_SCTR_SALUD
Call COLUMNA_SCTR_PENSION
Call COLUMNA_VIDA_LEY
Call COLUMNA_ACCIDENTES_PERSONALES
Call COLUMNA_SISTEMA_DE_PENSIONES
Call COLUMNA_CUSSP
Call COLUMNA_TIPO_DE_COMISION
Call COLUMNA_BANCO_SUELDO
Call COLUMNA_NRO_CUENTA_SUELDO
Call COLUMNA_MONEDA_CTA_SUELDO
Call COLUMNA_TIPO_CUENTA_SUELDO
Call COLUMNA_NRO_CUENTA_CTS
Call COLUMNA_NRO_CUENTA_CTS_CCI
Call COLUMNA_TIPO_MONEDA_CTS
Call COLUMNA_SEXO
Call COLUMNA_ESTADO_CIVIL
Call COLUMNA_UBIGEO
Call COLUMNA_LUGAR_DE_ORIGEN
Call COLUMNA_FUNCION_CUMPLE
Call COLUMNA_DIRECCIÓN
Call COLUMNA_DISTRITO
Call COLUMNA_PROVINCIA
Call COLUMNA_DEPARTAMENTO_1
Call COLUMNA_TELEFONO
Call COLUMNA_EDAD
Call COLUMNA_PROFESION
Call COLUMNA_GRUPO_SANGUINEO
Call COLUMNA_GRADO_DE_INSTRUCCION
Call COLUMNA_LUG_NACIMIENTO
Call COLUMNA_C_COSTO
Call COLUMNA_DISCAPACIDAD
Call COLUMNA_NRO_DE_HIJOS
Call COLUMNA_GRADO_SALARIAL
Call COLUMNA_TIEMPO_DE_SERVICIO
Call COLUMNA_CODIGO_UNIDAD_ORGANIZATIVA
Call COLUMNA_UNIDAD_ORGANIZATIVA
Call COLUMNA_CODIGO_RESPONSABLE_UNIDAD_ORGANIZATIVA
Call COLUMNA_UNIDAD_ORGANIZATIVA_PADRE
Call COLUMNA_MEDIDA_DE_CONTRATACION
Call COLUMNA_TIPO_DE_DEPOSITO
Call COLUMNA_GRUPO_DE_PERSONAL
Call COLUMNA_CORREO_ELECTRONICO
Call COLUMNA_USUARIO_DE_RED
Call COLUMNA_USUARIO_SAP
Call COLUMNA_SUELDO

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub COLUMNA_CODIGO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_codigo_sap As String
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Codigo")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_codigo_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_codigo_sap) + "]").Select
            Selection.Copy
            
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Codigo]").Select
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[Codigo]").Select
    Range("DATA_SAP[Codigo]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub

Sub COLUMNA_TIPO_DOCUMENTO()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


Dim NOMBRE As String
Dim Valor_tipo_documento_sap As String
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Tipo de documento")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_tipo_documento_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_tipo_documento_sap) + "]").Select
            Selection.Copy

            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Tipo de documento]").Select
            ActiveSheet.Paste
    End With

Application.DisplayAlerts = False
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[Tipo de documento]").Select
    Range("DATA_SAP[Tipo de documento]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_NUMERO_DOCUMENTO()
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


Dim NOMBRE As String
Dim Valor_numero_documento_sap As String
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Número de documento")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_numero_documento_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_numero_documento_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Número de documento]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[Número de documento]").Select
    Range("DATA_SAP[Número de documento]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_NOMBRE()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_nombre_sap As String

Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Nombre")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_nombre_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_nombre_sap) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Nombre]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[Nombre]").Select
    Range("DATA_SAP[Nombre]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_APELLIDO_PATERNO()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_apellido_paterno_sap As String
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Apellido paterno")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_apellido_paterno_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_apellido_paterno_sap) + "]").Select
            
            Selection.Copy
            
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Apellido paterno]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[Apellido paterno]").Select
    Range("DATA_SAP[Apellido paterno]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_APELLIDO_MATERNO()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


Dim NOMBRE As String
Dim Valor_apellido_materno_sap As String
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler


    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Apellido materno")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_apellido_materno_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_apellido_materno_sap) + "]").Select
            
            
            Selection.Copy
            
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Apellido materno]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[Apellido materno]").Select
    Range("DATA_SAP[Apellido materno]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_DESCRIPCION_TIPO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_descripcion_tipo_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler
    
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Descripción tipo")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_descripcion_tipo_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_descripcion_tipo_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Descripción tipo]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[Descripción tipo]").Select
    Range("DATA_SAP[Descripción tipo]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_EMPRESA()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_empresa_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler

    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Empresa")
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_empresa_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_empresa_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Empresa]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[Empresa]").Select
    Range("DATA_SAP[Empresa]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_DIVISION()
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_division_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("División")
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_division_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_division_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[División]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[División]").Select
    Range("DATA_SAP[División]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_AREA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_area_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate

Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate

Sheets("REPORTE_SAP").Select

columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Area")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_area_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_area_sap) + "]").Select
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Area]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[Area]").Select
    Range("DATA_SAP[Area]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_DEPARTAMENTO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_departamento_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler

    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Departamento")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_departamento_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_departamento_sap) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Departamento]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[Departamento]").Select
    Range("DATA_SAP[Departamento]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_CARGO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_cargo_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Cargo")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_cargo_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_cargo_sap) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Cargo]").Select
            ActiveSheet.Paste
    End With
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[Cargo]").Select
    Range("DATA_SAP[Cargo]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_NOMBRE_CARGO()
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_nombre_cargo_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Nombre cargo")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_nombre_cargo_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_nombre_cargo_sap) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Nombre cargo]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
    Sheets("SAP_PARAMETRIZADA").Select
    Range("DATA_SAP[Nombre cargo]").Select
    Range("DATA_SAP[Nombre cargo]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_FECHA_NACIMIENTO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_fecha_nacimiento_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("F. Nacimiento")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_fecha_nacimiento_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_fecha_nacimiento_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[F. Nacimiento]").Select
            ActiveSheet.Paste
            Range("DATA_SAP[F. Nacimiento]").Select
            Selection.NumberFormat = "m/d/yyyy"
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
        Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
        Sheets("SAP_PARAMETRIZADA").Select
        Range("DATA_SAP[F. Nacimiento]").Select
        Range("DATA_SAP[F. Nacimiento]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_FECHA_INGRESO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_fecha_ingreso_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler

    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("F. Ingreso")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_fecha_ingreso_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_fecha_ingreso_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[F. Ingreso]").Select
            ActiveSheet.Paste
            Range("DATA_SAP[F. Ingreso]").Select
            Selection.NumberFormat = "m/d/yyyy"
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[F. Ingreso]").Select
Range("DATA_SAP[F. Ingreso]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_FECHA_CESE()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_fecha_cese_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("F. Cese")
              
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_fecha_cese_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_fecha_cese_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[F. Cese]").Select
            ActiveSheet.Paste
            Range("DATA_SAP[F. Cese]").Select
            Selection.NumberFormat = "m/d/yyyy"
            
    End With
    
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[F. Cese]").Select
Range("DATA_SAP[F. Cese]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_MOTIVO_CESE()


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_motivo_cese_sap As String
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"
On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
    
            Set c = .Find("Motivo cese")
             
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_motivo_cese_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_motivo_cese_sap) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Motivo cese]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[Motivo cese]").Select
Range("DATA_SAP[Motivo cese]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_RELACION_LABORAL()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_relacion_laboral_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"
On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Relación laboral")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_relacion_laboral_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_relacion_laboral_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Relación laboral]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[Relación laboral]").Select
Range("DATA_SAP[Relación laboral]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_TIPO_TRABAJADOR()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_tipo_trabajador_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Tipo de trabajador")
            
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_tipo_trabajador_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_tipo_trabajador_sap) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Tipo de trabajador]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[Tipo de trabajador]").Select
Range("DATA_SAP[Tipo de trabajador]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_JORNADA_TRAB()


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


Dim NOMBRE As String
Dim Valor_jornada_trab_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Jornada trab.")
              
              
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_jornada_trab_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_jornada_trab_sap) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Jornada trab.]").Select
            ActiveSheet.Paste
    End With
      
      
   
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

    
    
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[Jornada trab.]").Select
Range("DATA_SAP[Jornada trab.]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_TIPO_DE_CONTRATO()

Application.DisplayAlerts = False
'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_tipo_contrato_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Tipo contrato")
              
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_tipo_contrato_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_tipo_contrato_sap) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Tipo contrato]").Select
            ActiveSheet.Paste
            
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[Tipo contrato]").Select
Range("DATA_SAP[Tipo contrato]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_INICIO_DE_CONTRATO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_inicio_contrato_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler


    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Inicio contrato")
              
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_inicio_contrato_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_inicio_contrato_sap) + "]").Select
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Inicio contrato]").Select
            ActiveSheet.Paste
            Range("DATA_SAP[Inicio contrato]").Select
            Selection.NumberFormat = "m/d/yyyy"
            
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[Inicio contrato]").Select
Range("DATA_SAP[Inicio contrato]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_FIN_DE_CONTRATO()

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


Dim NOMBRE As String
Dim Valor_fin_contrato_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler
    
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Fin contrato")
             
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_fin_contrato_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_fin_contrato_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Fin contrato]").Select
            ActiveSheet.Paste
            Range("DATA_SAP[Fin contrato]").Select
            Selection.NumberFormat = "m/d/yyyy"
    End With

Application.DisplayAlerts = False
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[Fin contrato]").Select
Range("DATA_SAP[Fin contrato]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_SINDICATO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")

Dim NOMBRE As String
Dim Valor_sindicato_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Sindicato")
             
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_sindicato_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_sindicato_sap) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Sindicato]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[Sindicato]").Select
Range("DATA_SAP[Sindicato]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_LOCACION()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_locacion_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("Locación")
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_locacion_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_locacion_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[Locación]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[Locación]").Select
Range("DATA_SAP[Locación]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_EPS_SITUACION()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_eps_situacion_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("EPS SITUACIÓN")
              
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_eps_situacion_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_eps_situacion_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[EPS SITUACIÓN]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[EPS SITUACIÓN]").Select
Range("DATA_SAP[EPS SITUACIÓN]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_TIPO_PLAN()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_tipo_plan_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler

    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("TIPO DE PLAN")
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_tipo_plan_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_tipo_plan_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[TIPO DE PLAN]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False



Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[TIPO DE PLAN]").Select
Range("DATA_SAP[TIPO DE PLAN]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_SCTR_SALUD()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_sctr_salud_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("S.C.T.R. SALUD")
             
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_sctr_salud_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_sctr_salud_sap) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[S.C.T.R. SALUD]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[S.C.T.R. SALUD]").Select
Range("DATA_SAP[S.C.T.R. SALUD]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_SCTR_PENSION()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_sctr_pension_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("S.C.T.R. PENSION")
             
               
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_sctr_pension_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_sctr_pension_sap) + "]").Select
            
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[S.C.T.R. PENSION]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[S.C.T.R. PENSION]").Select
Range("DATA_SAP[S.C.T.R. PENSION]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_VIDA_LEY()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_vida_ley_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("VIDA LEY")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_vida_ley_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_vida_ley_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[VIDA LEY]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False



Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[VIDA LEY]").Select
Range("DATA_SAP[VIDA LEY]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_ACCIDENTES_PERSONALES()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_accidentes_personales As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("ACCIDENTES PERSONALES")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_accidentes_personales = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_accidentes_personales) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[ACCIDENTES PERSONALES]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[ACCIDENTES PERSONALES]").Select
Range("DATA_SAP[ACCIDENTES PERSONALES]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_SISTEMA_DE_PENSIONES()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_sistema_de_pensiones As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"




On Error GoTo ErrorHandler

    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("SISTEMA DE PENSIONES")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_sistema_de_pensiones = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_sistema_de_pensiones) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[SISTEMA DE PENSIONES]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[SISTEMA DE PENSIONES]").Select
Range("DATA_SAP[SISTEMA DE PENSIONES]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_CUSSP()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_cussp As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("CUSSP")
             
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_cussp = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_cussp) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[CUSSP]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[CUSSP]").Select
Range("DATA_SAP[CUSSP]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_TIPO_DE_COMISION()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_tipo_de_comision As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("TIPO DE COMISIÓN AFP (MIXTA O FLUJO) Y %")
             
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_tipo_de_comision = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_tipo_de_comision) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[TIPO DE COMISIÓN AFP (MIXTA O FLUJO) Y %]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
 Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[TIPO DE COMISIÓN AFP (MIXTA O FLUJO) Y %]").Select
Range("DATA_SAP[TIPO DE COMISIÓN AFP (MIXTA O FLUJO) Y %]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_BANCO_SUELDO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_banco_sueldo As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("BANCO SUELDO")
            '
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_banco_sueldo = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_banco_sueldo) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[BANCO SUELDO]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[BANCO SUELDO]").Select
Range("DATA_SAP[BANCO SUELDO]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_NRO_CUENTA_SUELDO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_nro_cuenta_sueldo As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("NRO CUENTA SUELDO")
             
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_nro_cuenta_sueldo = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_nro_cuenta_sueldo) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[NRO CUENTA SUELDO]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[NRO CUENTA SUELDO]").Select
Range("DATA_SAP[NRO CUENTA SUELDO]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_MONEDA_CTA_SUELDO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_moneda_cta_sueldo As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("MONEDA CTA SUELDO")
             
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_moneda_cta_sueldo = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_moneda_cta_sueldo) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[MONEDA CTA SUELDO]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[MONEDA CTA SUELDO]").Select
Range("DATA_SAP[MONEDA CTA SUELDO]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_TIPO_CUENTA_SUELDO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_tipo_cuenta_sueldo As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("TIPO CUENTA SUELDO")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_tipo_cuenta_sueldo = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_tipo_cuenta_sueldo) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[TIPO CUENTA SUELDO]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[TIPO CUENTA SUELDO]").Select
Range("DATA_SAP[TIPO CUENTA SUELDO]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_NRO_CUENTA_CTS()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_nro_cuenta_cts As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("NRO CUENTA CTS")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_nro_cuenta_cts = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_nro_cuenta_cts) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[NRO CUENTA CTS]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[NRO CUENTA CTS]").Select
Range("DATA_SAP[NRO CUENTA CTS]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_NRO_CUENTA_CTS_CCI()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_nro_cuenta_cts_cci As String
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("NRO CUENTA CTS (CCI)")
             
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_nro_cuenta_cts_cci = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_nro_cuenta_cts_cci) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[NRO CUENTA CTS (CCI)]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[NRO CUENTA CTS (CCI)]").Select
Range("DATA_SAP[NRO CUENTA CTS (CCI)]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_BANCO_CTS()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_banco_cts As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("BANCO CTS")
             
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_banco_cts = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_banco_cts) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[BANCO CTS]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[BANCO CTS]").Select
Range("DATA_SAP[BANCO CTS]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_TIPO_MONEDA_CTS()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_tipo_moneda_cts As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("TIPO MONEDA CTS")
             
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_tipo_moneda_cts = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_tipo_moneda_cts) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[TIPO MONEDA CTS]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[TIPO MONEDA CTS]").Select
Range("DATA_SAP[TIPO MONEDA CTS]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_SEXO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_sexo As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("SEXO")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_sexo = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_sexo) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[SEXO]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[SEXO]").Select
Range("DATA_SAP[SEXO]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_ESTADO_CIVIL()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_estado_civil As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("ESTADO CIVIL")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_estado_civil = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_estado_civil) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[ESTADO CIVIL]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[ESTADO CIVIL]").Select
Range("DATA_SAP[ESTADO CIVIL]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_UBIGEO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_ubigeo As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("UBIGEO")

            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_ubigeo = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_ubigeo) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[UBIGEO]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[UBIGEO]").Select
Range("DATA_SAP[UBIGEO]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_LUGAR_DE_ORIGEN()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_lugar_de_origen As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("LUGAR DE ORIGEN")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_lugar_de_origen = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_lugar_de_origen) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[LUGAR DE ORIGEN]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[LUGAR DE ORIGEN]").Select
Range("DATA_SAP[LUGAR DE ORIGEN]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_FUNCION_CUMPLE()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_funcion_cumple As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("FUNCION QUE DESEMPEÑA")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_funcion_cumple = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_funcion_cumple) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[FUNCION QUE DESEMPEÑA]").Select
            ActiveSheet.Paste
    End With
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[FUNCION QUE DESEMPEÑA]").Select
Range("DATA_SAP[FUNCION QUE DESEMPEÑA]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_DIRECCIÓN()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_direccion As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("DIRECCIÓN")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_direccion = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_direccion) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[DIRECCIÓN]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[DIRECCIÓN]").Select
Range("DATA_SAP[DIRECCIÓN]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_DISTRITO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_direccion As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("DISTRITO")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_direccion = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_direccion) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[DISTRITO]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[DISTRITO]").Select
Range("DATA_SAP[DISTRITO]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_PROVINCIA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_provincia As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("PROVINCIA")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_provincia = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_provincia) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[PROVINCIA]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[PROVINCIA]").Select
Range("DATA_SAP[PROVINCIA]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub
Sub COLUMNA_DEPARTAMENTO_1()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_departamento As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler


    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("DEPARTAMENTO.1")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_departamento = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_departamento) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[DEPARTAMENTO2]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[DEPARTAMENTO2]").Select
Range("DATA_SAP[DEPARTAMENTO2]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_TELEFONO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_telefono As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("TELEFONO")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_telefono = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_telefono) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[TELEFONO]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[TELEFONO]").Select
Range("DATA_SAP[TELEFONO]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_EDAD()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_edad As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("EDAD")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_edad = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_edad) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[EDAD]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[EDAD]").Select
Range("DATA_SAP[EDAD]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_PROFESION()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_profesion As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("PROFESION")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_profesion = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_profesion) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[PROFESION]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[PROFESION]").Select
Range("DATA_SAP[PROFESION]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_GRUPO_SANGUINEO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_grupo_sanguineo As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("GRUPO SANGUINEO")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_grupo_sanguineo = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_grupo_sanguineo) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[GRUPO SANGUINEO]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[GRUPO SANGUINEO]").Select
Range("DATA_SAP[GRUPO SANGUINEO]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_GRADO_DE_INSTRUCCION()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_grado_de_instruccion As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("GRADO DE INSTRUCCION")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_grado_de_instruccion = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_grado_de_instruccion) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[GRADO DE INSTRUCCION]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[GRADO DE INSTRUCCION]").Select
Range("DATA_SAP[GRADO DE INSTRUCCION]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_LUG_NACIMIENTO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_lugar_de_nacimiento As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("LUG.NACIMIENTO")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_lugar_de_nacimiento = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_lugar_de_nacimiento) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[LUG.NACIMIENTO]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[LUG.NACIMIENTO]").Select
Range("DATA_SAP[LUG.NACIMIENTO]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_C_COSTO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_c_costo As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("C.COSTO")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_c_costo = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_c_costo) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[C.COSTO]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[C.COSTO]").Select
Range("DATA_SAP[C.COSTO]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_DISCAPACIDAD()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_discapacidad As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("DISCAPACIDAD")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_discapacidad = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_discapacidad) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[DISCAPACIDAD]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[DISCAPACIDAD]").Select
Range("DATA_SAP[DISCAPACIDAD]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_NRO_DE_HIJOS()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_nro_hijos As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("NRO DE HIJOS")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_nro_hijos = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_nro_hijos) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[NRO DE HIJOS]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


    
Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[NRO DE HIJOS]").Select
Range("DATA_SAP[NRO DE HIJOS]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_GRADO_SALARIAL()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_grado_salarial As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("GRADO SALARIAL")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_grado_salarial = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_grado_salarial) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[GRADO SALARIAL]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[GRADO SALARIAL]").Select
Range("DATA_SAP[GRADO SALARIAL]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_TIEMPO_DE_SERVICIO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_tiempo_de_servicio As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("TIEMPO DE SERVICIO")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_tiempo_de_servicio = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_tiempo_de_servicio) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[TIEMPO DE SERVICIO]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[TIEMPO DE SERVICIO]").Select
Range("DATA_SAP[TIEMPO DE SERVICIO]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_CODIGO_UNIDAD_ORGANIZATIVA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_codigo_unidad_organizativa As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("CÓDIGO UNIDAD ORGANIZATIVA")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_codigo_unidad_organizativa = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_codigo_unidad_organizativa) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[CÓDIGO UNIDAD ORGANIZATIVA]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False



Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[CÓDIGO UNIDAD ORGANIZATIVA]").Select
Range("DATA_SAP[CÓDIGO UNIDAD ORGANIZATIVA]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_UNIDAD_ORGANIZATIVA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_unidad_organizativa As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"



On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("UNIDAD ORGANIZATIVA")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_unidad_organizativa = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_unidad_organizativa) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[UNIDAD ORGANIZATIVA]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[UNIDAD ORGANIZATIVA]").Select
Range("DATA_SAP[UNIDAD ORGANIZATIVA]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_CODIGO_RESPONSABLE_UNIDAD_ORGANIZATIVA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_codigo_responsable_unidad_organizativa As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("CÓDIGO RESPONSABLE DE UNIDAD ORGANIZATIV")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_codigo_responsable_unidad_organizativa = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_codigo_responsable_unidad_organizativa) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[CÓDIGO RESPONSABLE DE UNIDAD ORGANIZATIV]").Select
            ActiveSheet.Paste
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[CÓDIGO RESPONSABLE DE UNIDAD ORGANIZATIV]").Select
Range("DATA_SAP[CÓDIGO RESPONSABLE DE UNIDAD ORGANIZATIV]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_RESPONSABLE_DE_UNIDAD_ORGANIZATIVA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_responsable_unidad_organizativa As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("RESPONSABLE DE UNIDAD ORGANIZATIVA")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_responsable_unidad_organizativa = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_responsable_unidad_organizativa) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[RESPONSABLE DE UNIDAD ORGANIZATIVA]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[RESPONSABLE DE UNIDAD ORGANIZATIVA]").Select
Range("DATA_SAP[RESPONSABLE DE UNIDAD ORGANIZATIVA]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_UNIDAD_ORGANIZATIVA_PADRE()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_unidad_organizativa_padre As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler

    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("UNIDAD ORGANIZATIVA PADRE")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_unidad_organizativa_padre = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_unidad_organizativa_padre) + "]").Select


            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[UNIDAD ORGANIZATIVA PADRE]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[UNIDAD ORGANIZATIVA PADRE]").Select
Range("DATA_SAP[UNIDAD ORGANIZATIVA PADRE]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_MEDIDA_DE_CONTRATACION()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_medida_contratacion As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("MEDIDA DE CONTRATACIÓN")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_medida_contratacion = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_medida_contratacion) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[MEDIDA DE CONTRATACIÓN]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[MEDIDA DE CONTRATACIÓN]").Select
Range("DATA_SAP[MEDIDA DE CONTRATACIÓN]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_TIPO_DE_DEPOSITO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_tipo_de_deposito As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler

    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("TIPO DE DEPOSITO")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_tipo_de_deposito = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_tipo_de_deposito) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[TIPO DE DEPOSITO]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[TIPO DE DEPOSITO]").Select
Range("DATA_SAP[TIPO DE DEPOSITO]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"
End Sub
Sub COLUMNA_GRUPO_DE_PERSONAL()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_grupo_de_personal As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("GRUPO DE PERSONAL")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_grupo_de_personal = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_grupo_de_personal) + "]").Select
            
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[GRUPO DE PERSONAL]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[GRUPO DE PERSONAL]").Select
Range("DATA_SAP[GRUPO DE PERSONAL]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_CORREO_ELECTRONICO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_correo_electronico As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("CORREO ELECTRONICO")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_correo_electronico = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_correo_electronico) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[CORREO ELECTRONICO]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[CORREO ELECTRONICO]").Select
Range("DATA_SAP[CORREO ELECTRONICO]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_USUARIO_DE_RED()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_usuario_de_red As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("USUARIO DE RED")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_usuario_de_red = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_usuario_de_red) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[USUARIO DE RED]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[USUARIO DE RED]").Select
Range("DATA_SAP[USUARIO DE RED]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_USUARIO_SAP()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_usuario_sap As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("USUARIO SAP")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_usuario_sap = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_usuario_sap) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[USUARIO SAP]").Select
            ActiveSheet.Paste
    End With
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[USUARIO SAP]").Select
Range("DATA_SAP[USUARIO SAP]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub
Sub COLUMNA_COMUNIDAD()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_comunidad As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"


On Error GoTo ErrorHandler
    
    With Sheets("REPORTE_SAP").Range(columnas)
            Set c = .Find("COMUNIDAD")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_comunidad = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_comunidad) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[COMUNIDAD]").Select
            ActiveSheet.Paste
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[COMUNIDAD]").Select
Range("DATA_SAP[COMUNIDAD]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub


Sub COLUMNA_SUELDO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'Workbooks.Open "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\3.Indicadores DHO-Sofía\Recursos\VBA_SAP\MC PROYECTO\PROCESO_VALIDACION.xlsm"
'Set DETALLE = Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA")
Dim NOMBRE As String
Dim Valor_sueldo As String
'Workbooks("REPORTES_VALIDACION_SAP_TRs.xlsm").Worksheets("REPORTE_SAP").Activate
Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Select
columnas = "A1:ZZ1"

On Error GoTo ErrorHandler

    With Sheets("REPORTE_SAP").Range(columnas)
            
            Set c = .Find("SUELDO_REPORTE")
            
            letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(c.Row)
            Valor_sueldo = Range(NOMBRE).Value
            Range("DATA_SAP_REPORTE[" + CStr(Valor_sueldo) + "]").Select
            
            Selection.Copy
            Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
            Range("B10").Select
            Sheets("SAP_PARAMETRIZADA").Select
            Range("DATA_SAP[SUELDO]").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[SUELDO]").Select
Range("DATA_SAP[SUELDO]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"
End Sub


Sub LIMPIAR_TABLA_SAP()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Workbooks("PROCESO_VALIDACION.xlsm").Worksheets("SAP_PARAMETRIZADA").Activate
Sheets("SAP_PARAMETRIZADA").Select
Range("DATA_SAP[[Codigo]:[SUELDO]]").Select
Selection.ClearContents

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub












