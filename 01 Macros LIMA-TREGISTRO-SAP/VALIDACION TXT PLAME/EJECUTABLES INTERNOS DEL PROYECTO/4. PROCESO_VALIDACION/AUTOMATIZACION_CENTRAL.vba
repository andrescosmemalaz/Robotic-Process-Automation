Attribute VB_Name = "AUTOMATIZACION_CENTRAL"
Sub AUTOMATION_TR()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
    Call LIMPIAR_TRS
    Call AGREGAR_FUNCION_AL_SUELDO
    Call ANALIZAR_TR5
    Call ANALIZAR_TR6
    Call modelo_eliminacion_celdas_vacias
    Call REPORTE_TR6
    Call REPORTE_TR5
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub LIMPIAR_TRS()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
    Call LIMPIAR_TR_6
    Call LIMPIAR_TR_5
    Call LIMPIAR_REPORTES_FINALES
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub
Sub SACAR_INFO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

    Call MOVIMIENTO_SAP
    Call MOVIMIENTO_TR5
    Call MOVIMIENTO_TR6
    Call MOVIMIENTO_SAP_REPORTE_SUELDO
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub ANALIZAR_TR5()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

    Call INSERTAR_CODIGO_TR_5
    Call VALIDAR_FECHA_INGRESO
    Call VALIDAR_TIPO_TRABAJADOR
    Call VALIDAR_CATEGORIA_OCUPACIONAL
    Call VALIDAR_NIVEL_EDUCATIVO
    Call VALIDACION_DISCAPACIDAD
    Call VALIDAR_SINDICALIZADO
    Call VALIDACION_REGIMEN_ACUMULATIVO
    Call VALIDAR_JORNADA_MAXIMA
    Call VALIDACION_HORARIO_NOCTURNO
    Call VALIDACION_DOBLE_SITUACION_FISCALIZACION
    Call VALIDACION_TIPO_DE_CONTRATO
    Call VALIDACION_NRO_CUENTA
    Call VALIDACION_REMUNERACION_BASICA
    Call VALIDACION_ENTIDAD_FINANCIERA_EMPLEADO
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub
Sub ANALIZAR_TR6()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

    Call INSERTAR_CODIGO_TR_6
    Call VALIDAR_CUSSP_TR_6
    Call VALIDAR_TIPO_REGIMEN_TR_6
    Call VALIDAR_REGIMEN_SALUD_TR_6
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False



End Sub
Sub LIMPIAR_TR_6()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
    Dim DOC_TR_6 As Range
    Dim CUSSP_TR_6 As Range
    Dim CUSSP_SAP As Range
    Dim VALIDACION_CUSSP As Range
    Dim TIPO_REGIMEN_TR As Range
    Dim TIPO_REGIMEN_SAP As Range
    Dim VALIDACION_TIPO_REGIMEN As Range
    Dim TIPO_REGIMEN_SALUD_TR As Range
    Dim TIPO_REGIMEN_SALUD_SAP As Range
    Dim VALIDACION_TIPO_REGIMEN_SALUD As Range
    Dim columnas As String
    Dim hoja As String
    Dim letra_columna As String
    Dim ws As Worksheet
    
    Sheets("TR6_PARAMETRIZADA").Select
    columnas = "A:ZZ"
    
    With Sheets("TR6_PARAMETRIZADA").Range(columnas)
          Set DOC_TR_6 = .Find("NUMERO DOCUMENTO TR6")
          Set CUSSP_TR_6 = .Find("CUSSP TR")
          Set CUSSP_SAP = .Find("CUSSP SAP")
          Set VALIDACION_CUSSP = .Find("VALIDACIÓN CUSSP TR-SAP")
          Set TIPO_REGIMEN_TR = .Find("TIPO DE REGIMEN TR")
          Set TIPO_REGIMEN_SAP = .Find("TIPO DE REGIMEN SAP")
          Set VALIDACION_TIPO_REGIMEN = .Find("VALIDACIÓN TIPO DE REGIMEN")
          Set TIPO_REGIMEN_SALUD_TR = .Find("TIPO DE REGIMEN SALUD TR")
          Set TIPO_REGIMEN_SALUD_SAP = .Find("TIPO DE REGIMEN SALUD SAP")
          Set VALIDACION_TIPO_REGIMEN_SALUD = .Find("VALIDACIÓN TIPO DE REGIMEN TR-SAP")

          If Not DOC_TR_6 Is Nothing Then
            Columns(DOC_TR_6.Column).Delete
          End If
          If Not CUSSP_TR_6 Is Nothing Then
            Columns(CUSSP_TR_6.Column).Delete
          End If
          If Not CUSSP_SAP Is Nothing Then
            Columns(CUSSP_SAP.Column).Delete
          End If
          If Not VALIDACION_CUSSP Is Nothing Then
            Columns(VALIDACION_CUSSP.Column).Delete
          End If
          If Not TIPO_REGIMEN_TR Is Nothing Then
            Columns(TIPO_REGIMEN_TR.Column).Delete
          End If
          If Not TIPO_REGIMEN_SAP Is Nothing Then
            Columns(TIPO_REGIMEN_SAP.Column).Delete
          End If
          If Not VALIDACION_TIPO_REGIMEN Is Nothing Then
            Columns(VALIDACION_TIPO_REGIMEN.Column).Delete
          End If
          If Not TIPO_REGIMEN_SALUD_TR Is Nothing Then
            Columns(TIPO_REGIMEN_SALUD_TR.Column).Delete
          End If
          If Not TIPO_REGIMEN_SALUD_SAP Is Nothing Then
            Columns(TIPO_REGIMEN_SALUD_SAP.Column).Delete
          End If
          If Not VALIDACION_TIPO_REGIMEN_SALUD Is Nothing Then
            Columns(VALIDACION_TIPO_REGIMEN_SALUD.Column).Delete
          End If
          Range("C9").Select
    End With
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub LIMPIAR_TR_5()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
    Dim DOC_TR_5 As Range
    Dim FECHA_INGRESO_TR5 As Range
    Dim FECHA_INGRESO_SAP As Range
    Dim VALIDACION_FECHA_INGRESO As Range
    Dim TIPO_TRABAJADOR_TR5 As Range
    Dim TIPO_TRABAJADOR_SAP As Range
    Dim VALIDACION_TIPO_DE_TRABAJADOR As Range
    Dim CATEGORIA_OCUPACIONAL_TR As Range
    Dim CATEGORIA_OCUPACIONAL_SAP As Range
    Dim VALIDACION_CATEGORIA_OCUPACIONAL As Range
    Dim NIVEL_EDUCATIVO_TR As Range
    Dim NIVEL_EDUCATIVO_SAP As Range
    Dim VALIDACION_NIVEL_EDUCATIVO As Range
    Dim DISCAPACIDAD_TR As Range
    Dim DISCAPACIDAD_SAP As Range
    Dim VALIDACION_DE_DISCAPACIDAD As Range
    Dim SINDICALIZADO_TR As Range
    Dim SINDICALIZADO_SAP As Range
    Dim VALIDACION_DE_SINDICALIZADO As Range
    Dim REGIMEN_ACUMULATIVO_TR As Range
    Dim REGIMEN_ACUMULATIVO_SAP As Range
    Dim VALIDACION_REGIMEN_ACUMULATIVO As Range
    Dim JORNADA_MAX_TR As Range
    Dim JORNADA_MAX_SAP_REGIMEN_ESPECIAL_DE_TRABAJO As Range
    Dim JORNADA_MAX_JORNADA_DE_TRABAJO As Range
    Dim VALIDACION_JORNADA_MAXIMA_SAP_ERROR As Range
    Dim HORARIO_NOCTURNO_TR As Range
    Dim VALIDACION_HORARIO_NOCTURNO_SAP As Range
    Dim SITUACION_ESPECIAL_DEL_TRABAJADOR As Range
    Dim VAL_FISCA As Range
    Dim VALIDACION_DOBLE_SITUACION_FISCALIZADO As Range
    Dim TIPO_DE_CONTRATO_TR As Range
    Dim TIPO_DE_CONTRATO_SAP As Range
    Dim VALIDACION_DEL_TIPO_DE_CONTRATO As Range
    Dim NUMERO_DE_CUENTA_TR As Range
    Dim NUMERO_DE_CUENTA_SAP As Range
    Dim VALIDACION_DE_NUMERO_DE_CUENTA As Range
    Dim REMUNERACION_BASICA_TR As Range
    Dim REMUNERACION_BASICA_SAP As Range
    Dim VALIDACION_DE_REMUNERACION_BASICA As Range
    Dim ENTIDAD_FINANCIERA_TR As Range
    Dim ENTIDAD_FINANCIERA_SAP As Range
    Dim VALIDACION_ENTIDAD_FINANCIERA As Range
    
    Dim columnas As String
    Dim hoja As String
    Dim letra_columna As String
    Dim ws As Worksheet
    
    Sheets("TR5_PARAMETRIZADA").Select
    columnas = "A:ZZ"
    
    With Sheets("TR5_PARAMETRIZADA").Range(columnas)
          Set DOC_TR_5 = .Find("NUMERO DOCUMENTO TR5")
          Set FECHA_INGRESO_TR5 = .Find("FECHA INICIO TR")
          Set FECHA_INGRESO_SAP = .Find("FECHA INICIO SAP")
          Set VALIDACION_FECHA_INGRESO = .Find("VALIDACIÓN FECHA INICIO")
          Set TIPO_TRABAJADOR_TR5 = .Find("TIPO DE TRABAJADOR TR")
          Set TIPO_TRABAJADOR_SAP = .Find("TIPO DE TRABAJADOR SAP")
          Set VALIDACION_TIPO_DE_TRABAJADOR = .Find("VALIDACIÓN TIPO DE TRABAJADOR")
          Set CATEGORIA_OCUPACIONAL_TR = .Find("CATEGORIA OCUPACIONAL TR")
          Set CATEGORIA_OCUPACIONAL_SAP = .Find("CATEGORIA OCUPACIONAL SAP")
          Set VALIDACION_CATEGORIA_OCUPACIONAL = .Find("VALIDACIÓN CATEGORIA OCUPACIONAL")
          Set NIVEL_EDUCATIVO_TR = .Find("NIVEL EDUCATIVO TR")
          Set NIVEL_EDUCATIVO_SAP = .Find("NIVEL EDUCATIVO SAP")
          Set VALIDACION_NIVEL_EDUCATIVO = .Find("VALIDACIÓN NIVEL EDUCATIVO")
          Set DISCAPACIDAD_TR = .Find("DISCAPACIDAD TR")
          Set DISCAPACIDAD_SAP = .Find("DISCAPACIDAD SAP")
          Set VALIDACION_DE_DISCAPACIDAD = .Find("VALIDACIÓN DE DISCAPACIDAD")
          Set SINDICALIZADO_TR = .Find("SINDICALIZADO TR")
          Set SINDICALIZADO_SAP = .Find("SINDICALIZADO SAP")
          Set VALIDACION_DE_SINDICALIZADO = .Find("VALIDACIÓN DE SINDICALIZADO")
          Set REGIMEN_ACUMULATIVO_TR = .Find("REGIMEN ACUMULATIVO TR")
          Set REGIMEN_ACUMULATIVO_SAP = .Find("REGIMEN ACUMULATIVO SAP")
          Set VALIDACION_REGIMEN_ACUMULATIVO = .Find("VALIDACIÓN REGIMEN ACUMULATIVO")
          Set JORNADA_MAX_TR = .Find("JORNADA MAXIMA TR")
          Set JORNADA_MAX_SAP_REGIMEN_ESPECIAL_DE_TRABAJO = .Find("JORNADA MAX SAP - REGIMEN ESPECIAL DE TRABAJO")
          Set JORNADA_MAX_JORNADA_DE_TRABAJO = .Find("JORNADA MAX -  JORNADA DE TRABAJO")
          Set VALIDACION_JORNADA_MAXIMA_SAP_ERROR = .Find("VALIDACION JORNADA MAXIMA SAP_ERROR")
          Set HORARIO_NOCTURNO_TR = .Find("HORARIO NOCTURNO TR")
          Set VALIDACION_HORARIO_NOCTURNO_SAP = .Find("VALIDACION HORARIO NOCTURNO SAP")
          Set SITUACION_ESPECIAL_DEL_TRABAJADOR = .Find("SITUACION ESPECIAL DEL TRABAJADOR")
          Set VAL_FISCA = .Find("VAL FISCA")
          Set VALIDACION_DOBLE_SITUACION_FISCALIZADO = .Find("VALIDACION DOBLE SITUACION FISCALIZADO")
          Set TIPO_DE_CONTRATO_TR = .Find("TIPO DE CONTRATO TR")
          Set TIPO_DE_CONTRATO_SAP = .Find("TIPO DE CONTRATO SAP")
          Set VALIDACION_DEL_TIPO_DE_CONTRATO = .Find("VALIDACION DEL TIPO DE CONTRATO")
          Set NUMERO_DE_CUENTA_TR = .Find("NUMERO DE CUENTA TR")
          Set NUMERO_DE_CUENTA_SAP = .Find("NUMERO DE CUENTA SAP")
          Set VALIDACION_DE_NUMERO_DE_CUENTA = .Find("VALIDACION DE NUMERO DE CUENTA")
          Set REMUNERACION_BASICA_TR = .Find("REMUNERACION BASICA TR")
          Set REMUNERACION_BASICA_SAP = .Find("REMUNERACION BASICA SAP")
          Set VALIDACION_DE_REMUNERACION_BASICA = .Find("VALIDACION DE REMUNERACION BASICA")
          Set ENTIDAD_FINANCIERA_TR = .Find("ENTIDAD FINANCIERA TR")
          Set ENTIDAD_FINANCIERA_SAP = .Find("ENTIDAD FINANCIERA SAP")
          Set VALIDACION_ENTIDAD_FINANCIERA = .Find("VALIDACION ENTIDAD FINANCIERA")
         
          
          If Not DOC_TR_5 Is Nothing Then
            Columns(DOC_TR_5.Column).Delete
          End If
          
          If Not FECHA_INGRESO_TR5 Is Nothing Then
            Columns(FECHA_INGRESO_TR5.Column).Delete
          End If
          
          If Not FECHA_INGRESO_SAP Is Nothing Then
            Columns(FECHA_INGRESO_SAP.Column).Delete
          End If
          
          If Not VALIDACION_FECHA_INGRESO Is Nothing Then
            Columns(VALIDACION_FECHA_INGRESO.Column).Delete
          End If
          
          If Not TIPO_TRABAJADOR_TR5 Is Nothing Then
            Columns(TIPO_TRABAJADOR_TR5.Column).Delete
          End If
          
          If Not TIPO_TRABAJADOR_SAP Is Nothing Then
            Columns(TIPO_TRABAJADOR_SAP.Column).Delete
          End If
          
          If Not VALIDACION_TIPO_DE_TRABAJADOR Is Nothing Then
            Columns(VALIDACION_TIPO_DE_TRABAJADOR.Column).Delete
          End If
          
          If Not CATEGORIA_OCUPACIONAL_TR Is Nothing Then
            Columns(CATEGORIA_OCUPACIONAL_TR.Column).Delete
          End If
          
          If Not CATEGORIA_OCUPACIONAL_SAP Is Nothing Then
            Columns(CATEGORIA_OCUPACIONAL_SAP.Column).Delete
          End If
          
          If Not VALIDACION_CATEGORIA_OCUPACIONAL Is Nothing Then
            Columns(VALIDACION_CATEGORIA_OCUPACIONAL.Column).Delete
          End If
          
          If Not NIVEL_EDUCATIVO_TR Is Nothing Then
            Columns(NIVEL_EDUCATIVO_TR.Column).Delete
          End If
          If Not NIVEL_EDUCATIVO_SAP Is Nothing Then
            Columns(NIVEL_EDUCATIVO_SAP.Column).Delete
          End If
          If Not VALIDACION_NIVEL_EDUCATIVO Is Nothing Then
            Columns(VALIDACION_NIVEL_EDUCATIVO.Column).Delete
          End If
          If Not DISCAPACIDAD_TR Is Nothing Then
            Columns(DISCAPACIDAD_TR.Column).Delete
          End If
          If Not DISCAPACIDAD_SAP Is Nothing Then
            Columns(DISCAPACIDAD_SAP.Column).Delete
          End If
          If Not VALIDACION_DE_DISCAPACIDAD Is Nothing Then
            Columns(VALIDACION_DE_DISCAPACIDAD.Column).Delete
          End If
          If Not SINDICALIZADO_TR Is Nothing Then
            Columns(SINDICALIZADO_TR.Column).Delete
          End If
          If Not SINDICALIZADO_SAP Is Nothing Then
            Columns(SINDICALIZADO_SAP.Column).Delete
          End If
          If Not VALIDACION_DE_SINDICALIZADO Is Nothing Then
            Columns(VALIDACION_DE_SINDICALIZADO.Column).Delete
          End If
          If Not REGIMEN_ACUMULATIVO_TR Is Nothing Then
            Columns(REGIMEN_ACUMULATIVO_TR.Column).Delete
          End If
          If Not REGIMEN_ACUMULATIVO_SAP Is Nothing Then
            Columns(REGIMEN_ACUMULATIVO_SAP.Column).Delete
          End If
          If Not VALIDACION_REGIMEN_ACUMULATIVO Is Nothing Then
            Columns(VALIDACION_REGIMEN_ACUMULATIVO.Column).Delete
          End If
          If Not JORNADA_MAX_TR Is Nothing Then
            Columns(JORNADA_MAX_TR.Column).Delete
          End If
          If Not JORNADA_MAX_SAP_REGIMEN_ESPECIAL_DE_TRABAJO Is Nothing Then
            Columns(JORNADA_MAX_SAP_REGIMEN_ESPECIAL_DE_TRABAJO.Column).Delete
          End If
          If Not JORNADA_MAX_JORNADA_DE_TRABAJO Is Nothing Then
            Columns(JORNADA_MAX_JORNADA_DE_TRABAJO.Column).Delete
          End If
          If Not VALIDACION_JORNADA_MAXIMA_SAP_ERROR Is Nothing Then
            Columns(VALIDACION_JORNADA_MAXIMA_SAP_ERROR.Column).Delete
          End If
          If Not HORARIO_NOCTURNO_TR Is Nothing Then
            Columns(HORARIO_NOCTURNO_TR.Column).Delete
          End If
          If Not VALIDACION_HORARIO_NOCTURNO_SAP Is Nothing Then
            Columns(VALIDACION_HORARIO_NOCTURNO_SAP.Column).Delete
          End If
          If Not SITUACION_ESPECIAL_DEL_TRABAJADOR Is Nothing Then
            Columns(SITUACION_ESPECIAL_DEL_TRABAJADOR.Column).Delete
          End If
          If Not VAL_FISCA Is Nothing Then
            Columns(VAL_FISCA.Column).Delete
          End If
          If Not VALIDACION_DOBLE_SITUACION_FISCALIZADO Is Nothing Then
            Columns(VALIDACION_DOBLE_SITUACION_FISCALIZADO.Column).Delete
          End If
          If Not TIPO_DE_CONTRATO_TR Is Nothing Then
            Columns(TIPO_DE_CONTRATO_TR.Column).Delete
          End If
          If Not TIPO_DE_CONTRATO_SAP Is Nothing Then
            Columns(TIPO_DE_CONTRATO_SAP.Column).Delete
          End If
          If Not VALIDACION_DEL_TIPO_DE_CONTRATO Is Nothing Then
            Columns(VALIDACION_DEL_TIPO_DE_CONTRATO.Column).Delete
          End If
          If Not NUMERO_DE_CUENTA_TR Is Nothing Then
            Columns(NUMERO_DE_CUENTA_TR.Column).Delete
          End If
          If Not NUMERO_DE_CUENTA_SAP Is Nothing Then
            Columns(NUMERO_DE_CUENTA_SAP.Column).Delete
          End If
          If Not VALIDACION_DE_NUMERO_DE_CUENTA Is Nothing Then
            Columns(VALIDACION_DE_NUMERO_DE_CUENTA.Column).Delete
          End If
          If Not REMUNERACION_BASICA_TR Is Nothing Then
            Columns(REMUNERACION_BASICA_TR.Column).Delete
          End If
          If Not REMUNERACION_BASICA_SAP Is Nothing Then
            Columns(REMUNERACION_BASICA_SAP.Column).Delete
          End If
          If Not VALIDACION_DE_REMUNERACION_BASICA Is Nothing Then
            Columns(VALIDACION_DE_REMUNERACION_BASICA.Column).Delete
          End If
          If Not ENTIDAD_FINANCIERA_TR Is Nothing Then
            Columns(ENTIDAD_FINANCIERA_TR.Column).Delete
          End If
          If Not ENTIDAD_FINANCIERA_SAP Is Nothing Then
            Columns(ENTIDAD_FINANCIERA_SAP.Column).Delete
          End If
          If Not VALIDACION_ENTIDAD_FINANCIERA Is Nothing Then
            Columns(VALIDACION_ENTIDAD_FINANCIERA.Column).Delete
          End If
          Range("C9").Select
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub LIMPIAR_REPORTES_FINALES()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

    On Error Resume Next
    
    Dim REPORTE_CUSSP As Range
    Dim REPORTE_REGIMENX As Range
    Dim REPORTE_SALUD As Range
    Dim REPORTE_FECHA_TR5 As Range
    
    Dim REPORTE_TIPO_DE_TRABAJADOR As Range
    Dim REPORTE_CATEGORIA_OCUPACIONAL As Range
    Dim REPORTE_NIVEL_EDUCATIVO As Range
    Dim REPORTE_DISCAPACIDAD As Range
    Dim REPORTE_SINDICALIZADO As Range
    Dim REPORTE_REGIMEN_ACUMULATIVO As Range
    Dim REPORTE_JORNADA_MAXIMA As Range
    Dim REPORTE_HORARIO_NOCTURNO As Range
    Dim REPORTE_SOBRE_DOBLE_FISCALIZACION  As Range
    Dim REPORTE_TIPO_CONTRATO As Range
    Dim REPORTE_NUMERO_DE_CUENTA As Range
    Dim REPORTE_REMUNERACION_BASICA As Range
    Dim REPORTE_ENTIDAD_FINANCIERA As Range
    
    Dim columnas As String
    Dim hoja As String
    Dim letra_columna As String
    Dim ws As Worksheet
    Sheets("REPORTE").Select
    columnas = "A10:ZZ10"
    
    With Sheets("REPORTE").Range(columnas)
          Range("REPORTE_CUSSP[#All]").Select
          Selection.Clear
          Range("REPORTE_REGIMENX[#All]").Select
          Selection.Clear
          Range("REPORTE_SALUD[#All]").Select
          Selection.Clear
          Range("REPORTE_FECHA_INICIO_TR5[#All]").Select
          Selection.Clear
          Range("REPORTE_TIPO_DE_TRABAJADOR[#All]").Select
          Selection.Clear
          Range("REPORTE_CATEGORIA_OCUPACIONAL[#All]").Select
          Selection.Clear
          Range("REPORTE_NIVEL_EDUCATIVO[#All]").Select
          Selection.Clear
          Range("REPORTE_DISCAPACIDAD[#All]").Select
          Selection.Clear
          Range("REPORTE_SINDICALIZADO[#All]").Select
          Selection.Clear
          Range("REPORTE_REGIMEN_ACUMULATIVO[#All]").Select
          Selection.Clear
          Range("REPORTE_JORNADA_MAXIMA[#All]").Select
          Selection.Clear
          Range("REPORTE_HORARIO_NOCTURNO[#All]").Select
          Selection.Clear
          Range("REPORTE_SOBRE_DOBLE_FISCALIZACION[#All]").Select
          Selection.Clear
          Range("REPORTE_TIPO_CONTRATO[#All]").Select
          Selection.Clear
          Range("REPORTE_NUMERO_DE_CUENTA[#All]").Select
          Selection.Clear
          Range("REPORTE_REMUNERACION_BASICA[#All]").Select
          Selection.Clear
          Range("REPORTE_ENTIDAD_FINANCIERA[#All]").Select
          Selection.Clear
          
          
          
          If Not REPORTE_CUSSP Is Nothing Then
            Range("REPORTE_CUSSP[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_REGIMENX Is Nothing Then
            Range("REPORTE_REGIMENX[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_SALUD Is Nothing Then
            Range("REPORTE_SALUD[#All]").Select
          Selection.Clear
          End If

          If Not REPORTE_FECHA_TR5 Is Nothing Then
            Range("REPORTE_FECHA_INICIO_TR5[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_TIPO_DE_TRABAJADOR Is Nothing Then
            Range("REPORTE_TIPO_DE_TRABAJADOR[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_CATEGORIA_OCUPACIONAL Is Nothing Then
            Range("REPORTE_CATEGORIA_OCUPACIONAL[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_NIVEL_EDUCATIVO Is Nothing Then
            Range("REPORTE_NIVEL_EDUCATIVO[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_DISCAPACIDAD Is Nothing Then
            Range("REPORTE_DISCAPACIDAD[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_SINDICALIZADO Is Nothing Then
            Range("REPORTE_SINDICALIZADO[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_REGIMEN_ACUMULATIVO Is Nothing Then
            Range("REPORTE_REGIMEN_ACUMULATIVO[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_JORNADA_MAXIMA Is Nothing Then
            Range("REPORTE_JORNADA_MAXIMA[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_HORARIO_NOCTURNO Is Nothing Then
            Range("REPORTE_HORARIO_NOCTURNO[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_SOBRE_DOBLE_FISCALIZACION Is Nothing Then
            Range("REPORTE_SOBRE_DOBLE_FISCALIZACION[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_TIPO_CONTRATO Is Nothing Then
            Range("REPORTE_TIPO_CONTRATO[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_NUMERO_DE_CUENTA Is Nothing Then
            Range("REPORTE_NUMERO_DE_CUENTA[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_REMUNERACION_BASICA Is Nothing Then
            Range("REPORTE_REMUNERACION_BASICA[#All]").Select
          Selection.Clear
          End If
          
          If Not REPORTE_ENTIDAD_FINANCIERA Is Nothing Then
            Range("REPORTE_ENTIDAD_FINANCIERA[#All]").Select
          Selection.Clear
          End If

          Range("C9").Select
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub
ErrorHandler:
    MsgBox "Ya no se encuentra el reporte de validación"
    
End Sub
Sub INSERTAR_CODIGO_TR_5()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

    Dim B As Range
    Dim c As Range
    Dim columnas As String
    Dim hoja As String
    Dim letra_columna As String
    Dim ws As Worksheet
    columnas = "A:Z"
    Sheets("TR5_PARAMETRIZADA").Select
    With Sheets("TR5_PARAMETRIZADA").Range(columnas)
          Set B = .Find("Número")
          If Not B Is Nothing Then
          B.EntireColumn.Insert
          End If
          Range("B10").Value = "NUMERO DOCUMENTO TR5"
          Range("TR5_PARAMETRIZADA[NUMERO DOCUMENTO TR5]").FormulaR1C1 = "=TRIM([@Número])"
    End With
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub
Sub INSERTAR_CODIGO_TR_6()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

    Dim B As Range
    Dim c As Range
    Dim columnas As String
    Dim hoja As String
    Dim letra_columna As String
    
    Sheets("TR6_PARAMETRIZADA").Select
    columnas = "A:Z"
    With Sheets("TR6_PARAMETRIZADA").Range(columnas)
          Set B = .Find("Número")
          
          If Not B Is Nothing Then
          B.EntireColumn.Insert
          End If
          
          Range("B10").Value = "NUMERO DOCUMENTO TR6"
          Range("TR6_PARAMETRIZADA[NUMERO DOCUMENTO TR6]").FormulaR1C1 = "=TRIM([@Número])"
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub VALIDAR_FECHA_INGRESO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


' VALIDAR_FECHA_INGRESO Macro

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_fecha_tr = tablaDatos.ListColumns.Add
    Set columna_fecha_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN DE FECHA DE INGRESO TR
    
    With columna_fecha_tr
        .Name = "FECHA INICIO TR"
        .DataBodyRange.Formula = "= [@[Fec. Inicio]]"
        Selection.NumberFormat = "m/d/yyyy"
    End With
    
    ' VALIDACIÓN DE FECHA DE INGRESO SAP
    
    With columna_fecha_sap
        .Name = "FECHA INICIO SAP"
        Selection.NumberFormat = "m/d/yyyy"
        .DataBodyRange.Formula = "=VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],13,0)"
    End With
    
    
    Set columna_validacion_fecha = tablaDatos.ListColumns.Add
    With columna_validacion_fecha
        .Name = "VALIDACIÓN FECHA INICIO"
        .DataBodyRange.Formula = "= ([@[FECHA INICIO TR]]-[@[FECHA INICIO SAP]])"
        Selection.NumberFormat = "0"
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
        End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
    End With
    End With


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub VALIDAR_TIPO_TRABAJADOR()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'VALIDAR_TIPO_TRABAJADOR
    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_tipo_trabajador_tr = tablaDatos.ListColumns.Add
    Set columna_tipo_trabajador_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN DE TIPO DE TRABAJADOR TR
    With columna_tipo_trabajador_tr
        .Name = "TIPO DE TRABAJADOR TR"
        .DataBodyRange.Formula = "= [@[Tipo de Trabajador]]"
    End With
    
    
    ' VALIDACIÓN DE TIPO DE TRABAJADOR SAP
    With columna_tipo_trabajador_sap
        .Name = "TIPO DE TRABAJADOR SAP"
        .DataBodyRange.Formula = "=TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" & Chr(10) & "VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],17,0),1,""""),2,""""),3,""""),4,""""),5,""""),6,""""),7,""""),8,""""),9,""""),0,""""),""-"",""""),""MINA"",""""))"
    End With
    
    
    Set columna_validacion_tipo_trabajador = tablaDatos.ListColumns.Add
    With columna_validacion_tipo_trabajador
        .Name = "VALIDACIÓN TIPO DE TRABAJADOR"
        .DataBodyRange.Formula = "=TRIM([@[TIPO DE TRABAJADOR TR]])=TRIM([TIPO DE TRABAJADOR SAP])"
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
        End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
    End With
    End With
   
 
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
 
 
End Sub

Sub VALIDAR_CATEGORIA_OCUPACIONAL()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


'
' VALIDAR_CATEGORIA_OCUPACIONAL

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_categoria_ocupacional_tr = tablaDatos.ListColumns.Add
    Set columna_categoria_ocupacional_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN DE CATEGORIA OCUPACIONAL TR
    With columna_categoria_ocupacional_tr
        .Name = "CATEGORIA OCUPACIONAL TR"
        .DataBodyRange.Formula = "= [@[Cat. Ocupacional]]"
    End With
    
    
    ' VALIDACIÓN DE CATEGORIA OCUPACIONAL SAP
    With columna_categoria_ocupacional_sap
        .Name = "CATEGORIA OCUPACIONAL SAP"
        .DataBodyRange.Formula = "=VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],5,0)"
    End With
    
    Set columna_validacion_tipo_trabajador_sap = tablaDatos.ListColumns.Add
    With columna_validacion_tipo_trabajador_sap
        .Name = "VALIDACIÓN CATEGORIA OCUPACIONAL"
        .DataBodyRange.Formula = "=TRIM([@[CATEGORIA OCUPACIONAL TR]])=TRIM([CATEGORIA OCUPACIONAL SAP])"
         Range(Selection, Selection.End(xlDown)).Select
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
        End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
    End With
    
    End With
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
End Sub
Sub VALIDAR_NIVEL_EDUCATIVO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

' VALIDAR_NIVEL_EDUCATIVO

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_nivel_educativo_tr = tablaDatos.ListColumns.Add
    Set columna_nivel_educativo_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN DE NIVEL EDUCATIVO TR
    With columna_nivel_educativo_tr
        .Name = "NIVEL EDUCATIVO TR"
        .DataBodyRange.Formula = "= [@[Nivel Educativo]]"
    End With
    
    
    ' VALIDACIÓN DE NIVEL EDUCATIVO SAP
    With columna_nivel_educativo_sap
        .Name = "NIVEL EDUCATIVO SAP"
        .DataBodyRange.Formula = "=VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],54,0)"
    End With
    
    
    Set columna_validacion_nivel_educativo_sap = tablaDatos.ListColumns.Add
    With columna_validacion_nivel_educativo_sap
        .Name = "VALIDACIÓN NIVEL EDUCATIVO"
        .DataBodyRange.Formula = "=TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE([@[NIVEL EDUCATIVO TR]],""EDUCACIÓN"",""""),""TÉCNICA"",""TECNICA""),""MAESTRÍA"",""MAESTRIA""),""ESTUDIOS DE "",""""))=TRIM([@[NIVEL EDUCATIVO SAP]])"
        
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
        End With
        
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.149998474074526
        End With
        
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub VALIDACION_DISCAPACIDAD()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


' VALIDAR_DISCAPACIDAD

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_discapacidad_tr = tablaDatos.ListColumns.Add
    Set columna_discapacidad_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN DE DISCAPACIDAD TR
    With columna_discapacidad_tr
        .Name = "DISCAPACIDAD TR"
        .DataBodyRange.Formula = "= [@[Discapacidad]]"
    End With
    
    ' VALIDACIÓN DE DISCAPACIDAD SAP
    With columna_discapacidad_sap
        .Name = "DISCAPACIDAD SAP"
        .DataBodyRange.Formula = "=IF(LEN(VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],57,0))=0,""NO"",""TIENE DISCAPACIDAD"")"
    End With
    
    
    Set columna_validacion_discapacidad_sap = tablaDatos.ListColumns.Add
    With columna_validacion_discapacidad_sap
        .Name = "VALIDACIÓN DE DISCAPACIDAD"
        .DataBodyRange.Formula = "= TRIM([@[DISCAPACIDAD TR]])=TRIM([DISCAPACIDAD SAP])"
        
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
        End With
        
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.149998474074526
        End With
        
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub
Sub VALIDAR_SINDICALIZADO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False



' VALIDAR_SINDICALIZADO

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_discapacidad_tr = tablaDatos.ListColumns.Add
    Set columna_discapacidad_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN DE SINDICALIZADO TR
    With columna_discapacidad_tr
        .Name = "SINDICALIZADO TR"
        .DataBodyRange.Formula = "= [@[Sindicalizado]]"
    End With
    
    ' VALIDACIÓN DE SINDICALIZADO SAP
    With columna_discapacidad_sap
        .Name = "SINDICALIZADO SAP"
        .DataBodyRange.Formula = "=IF(TRIM(MID(SUBSTITUTE(VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],22,0),""LIMA - GENERAL"",""NO""),1,3))=""NO"",""NO"",""SI"")"
    End With
    
    
    Set columna_validacion_discapacidad_sap = tablaDatos.ListColumns.Add
    With columna_validacion_discapacidad_sap
        .Name = "VALIDACIÓN DE SINDICALIZADO"
        .DataBodyRange.Formula = "=TRIM([@[SINDICALIZADO TR]])=TRIM([SINDICALIZADO SAP])"
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
        End With
        
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.149998474074526
        End With
        
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub
Sub VALIDACION_REGIMEN_ACUMULATIVO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


' VALIDAR_REGIMEN_ACUMULATIVO

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_discapacidad_tr = tablaDatos.ListColumns.Add
    Set columna_discapacidad_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN DE REGIMEN_ACUMULATIVO TR
    With columna_discapacidad_tr
        .Name = "REGIMEN ACUMULATIVO TR"
        .DataBodyRange.Formula = "= [@[Reg. Acumulativo]]"
    End With
    
    ' VALIDACIÓN DE REGIMEN_ACUMULATIVO SAP
    With columna_discapacidad_sap
        .Name = "REGIMEN ACUMULATIVO SAP"
        .DataBodyRange.Formula = "=IF(VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],18,0)=""VMP 5 X 2"",""NO"",""SI"")"
    End With
    
    
    Set columna_validacion_discapacidad_sap = tablaDatos.ListColumns.Add
    With columna_validacion_discapacidad_sap
        .Name = "VALIDACIÓN REGIMEN ACUMULATIVO"
        .DataBodyRange.Formula = "=TRIM([@[REGIMEN ACUMULATIVO TR]])=TRIM([REGIMEN ACUMULATIVO SAP])"
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.149998474074526
        End With
        
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub VALIDAR_JORNADA_MAXIMA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


' VALIDAR_JORNADA_MAXIMA

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_jornada_maxima_tr = tablaDatos.ListColumns.Add
    Set columna_jornada_maxima_sap_regimen_especial = tablaDatos.ListColumns.Add
    Set columna_jornada_de_trabajo = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN DE JORNADA MAXIMA TR
    With columna_jornada_maxima_tr
        .Name = "JORNADA MAXIMA TR"
        .DataBodyRange.Formula = "= [@[Reg. Acumulativo]]"
    End With
    
    
    ' VALIDACIÓN DE JORNADA MAX SAP - REGIMEN ESPECIAL DE TRABAJO"
    With columna_jornada_maxima_sap_regimen_especial
        .Name = "JORNADA MAX SAP - REGIMEN ESPECIAL DE TRABAJO"
        .DataBodyRange.Formula = "=VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],16,0)"
    End With
    
     ' VALIDACIÓN DE JORNADA MAXIMA SAP
    With columna_jornada_de_trabajo
        .Name = "JORNADA MAX -  JORNADA DE TRABAJO"
        .DataBodyRange.Formula = "=VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],18,0)"
    End With
    
     Set columna_jornada_maxima_sap_error = tablaDatos.ListColumns.Add
    ' VALIDACIÓN JORNADA MAX - ERROR JORNADA MAXIMA
    With columna_jornada_maxima_sap_error
        .Name = "VALIDACION JORNADA MAXIMA SAP_ERROR"
        .DataBodyRange.Formula = "=AND(TEXT(TRIM([@[JORNADA MAX -  JORNADA DE TRABAJO]]),)=""VMP 5 X 2"",TEXT(TRIM([@[JORNADA MAX SAP - REGIMEN ESPECIAL DE TRABAJO]]),)=""FISCALIZADO"")"
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.149998474074526
        End With

    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub VALIDACION_HORARIO_NOCTURNO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


    ' VALIDAR_HORARIO_NOCTURNO
    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_horario_nocturno_tr = tablaDatos.ListColumns.Add
    Set columna_horario_nocturno_sap = tablaDatos.ListColumns.Add
   
    
    ' VALIDACIÓN DE HORARIO NOCTURNO TR
    With columna_horario_nocturno_tr
        .Name = "HORARIO NOCTURNO TR"
        .DataBodyRange.Formula = "= [@[Horario Nocturno]]"
    End With
    
    
    ' VALIDACIÓN DE HORARIO NOCTURNO SAP
    With columna_horario_nocturno_sap
        .Name = "VALIDACION HORARIO NOCTURNO SAP"
        .DataBodyRange.Formula = "=IF(TRIM([@[HORARIO NOCTURNO TR]])=""NO"",""SIN ERROR"", ""OBSERVACIÓN"")"
    End With
    
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.149998474074526
        End With
        
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub VALIDACION_DOBLE_SITUACION_FISCALIZACION()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

 ' VALIDAR DOBLE SITUACION FISCALIZACION
 
    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_doble_situacion_fiscalizacion_tr = tablaDatos.ListColumns.Add
    
    Set columna_validacion_fiscalizado = tablaDatos.ListColumns.Add
    
   
    
    ' VALIDACIÓN DOBLE SITUACION FISCALIZACION TR
    With columna_doble_situacion_fiscalizacion_tr
        .Name = "SITUACION ESPECIAL DEL TRABAJADOR"
        .DataBodyRange.Formula = "= [@[Situación Especial del Trabajador]]"
    End With
    
    ' VALIDACIÓN FISCALIZADO
    With columna_validacion_fiscalizado
        .Name = "VAL FISCA"
        .DataBodyRange.Formula = "= VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],16,0)"
    End With
    
    Set columna_doble_situacion_fiscalizacion_sap = tablaDatos.ListColumns.Add
    ' VALIDACIÓN DOBLE SITUACION FISCALIZACION SAP
    With columna_doble_situacion_fiscalizacion_sap
        .Name = "VALIDACION DOBLE SITUACION FISCALIZADO"
        .DataBodyRange.Formula = "=OR(AND(OR(IF(TRIM([@[SITUACION ESPECIAL DEL TRABAJADOR]])=""TELETRABAJO COMPLETO"",TRUE,FALSE),IF(TRIM([@[SITUACION ESPECIAL DEL TRABAJADOR]])=""NINGUNA"",TRUE,FALSE)),IF([@[VAL FISCA]]=""FISCALIZADO"",TRUE,FALSE)),OR(IFERROR(AND(TRIM(MID([@[VAL FISCA]],FIND(""CONF"",[@[VAL FISCA]]),4)) = ""CONF"",TRIM(MID([@[SITUACION ESPECIAL DEL TRABAJADOR]],FIND(""CONFIANZA"",[@[SITUACION ESPECIAL DEL TRABAJADOR]]),10)) =""CONFIANZA""),""NO COINCIDE""),2=3))" & _
        ""
    End With
    
       With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub VALIDACION_TIPO_DE_CONTRATO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


 ' VALIDAR TIPO DE CONTRATO
 
    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_doble_situacion_fiscalizacion_tr = tablaDatos.ListColumns.Add
    Set columna_validacion_fiscalizado = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN TIPO DE CONTRATO TR
    With columna_doble_situacion_fiscalizacion_tr
        .Name = "TIPO DE CONTRATO TR"
        .DataBodyRange.Formula = "= [@[Tipo de Contrato]]"
    End With
    
    ' VALIDACIÓN TIPO DE CONTRATO SAP
    With columna_validacion_fiscalizado
        .Name = "TIPO DE CONTRATO SAP"
        .DataBodyRange.Formula = "= VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],19,0)"
    End With
    
    Set columna_doble_situacion_fiscalizacion_sap = tablaDatos.ListColumns.Add
    ' VALIDACIÓN DOBLE SITUACION FISCALIZACION SAP
    With columna_doble_situacion_fiscalizacion_sap
        .Name = "VALIDACION DEL TIPO DE CONTRATO"
        .DataBodyRange.Formula = "=OR(IF(ISNUMBER(FIND(728,[@[TIPO DE CONTRATO TR]],1))*ISNUMBER(FIND(728,[@[TIPO DE CONTRATO SAP]],1))=0,FALSE,TRUE),ISNUMBER(FIND(""OBRA"",[@[TIPO DE CONTRATO TR]],1)*FIND(""OBRA"",[@[TIPO DE CONTRATO SAP]],1)),ISNUMBER(FIND(""SUPLENCIA"",TRIM([@[TIPO DE CONTRATO TR]]),1)*FIND(""SUPLENCIA"",TRIM([@[TIPO DE CONTRATO SAP]]),1)),ISNUMBER(FIND(""INCREM"",TRIM([@[TIPO DE CONTRATO TR]]),1)*FI" & _
        "ND(""INCREM"",TRIM([@[TIPO DE CONTRATO SAP]]),1)),ISNUMBER(FIND(""NECES"",TRIM([@[TIPO DE CONTRATO TR]]),1)*FIND(""NECES"",TRIM([@[TIPO DE CONTRATO SAP]]),1)))" & _
        ""
        
    End With
    
        With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub VALIDACION_NRO_CUENTA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


' VALIDACION_NRO_CUENTA

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_numero_cuenta_tr = tablaDatos.ListColumns.Add
    Set columna_numero_cuenta_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN NUMERO DE CUENTA TR
    With columna_numero_cuenta_tr
        .Name = "NUMERO DE CUENTA TR"
        .DataBodyRange.Formula = "= [@[Nro de Cuenta]]"
    End With
    
    ' VALIDACIÓN NUMERO DE CUENTA SAP
    With columna_numero_cuenta_sap
        .Name = "NUMERO DE CUENTA SAP"
        .DataBodyRange.Formula = "= VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],34,0)"
    End With
    
    Set columna_numero_cuenta_fiscalizacion_sap = tablaDatos.ListColumns.Add
    ' VALIDACIÓN DOBLE SITUACION FISCALIZACION SAP
    With columna_numero_cuenta_fiscalizacion_sap
        .Name = "VALIDACION DE NUMERO DE CUENTA"
        .DataBodyRange.Formula = "=IF(LEN(TRIM([@NUMERO DE CUENTA SAP]))=LEN(TRIM([@[NUMERO DE CUENTA TR]])),TRIM([@NUMERO DE CUENTA SAP])=TRIM([@[NUMERO DE CUENTA TR]]),""NO ESTAN IGUALES LOS CARACTERES"")"
    End With
    
        With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub VALIDACION_REMUNERACION_BASICA()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


' VALIDACION_REMUNERACION_BASICA

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_numero_cuenta_tr = tablaDatos.ListColumns.Add
    Set columna_numero_cuenta_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN REMUNERACION BASICA TR
    With columna_numero_cuenta_tr
        .Name = "REMUNERACION BASICA TR"
        .DataBodyRange.Formula = "= [@[Remun Bas.]]"
    End With
    
    ' VALIDACIÓN REMUNERACION BASICA SAP
    With columna_numero_cuenta_sap
        .Name = "REMUNERACION BASICA SAP"
        .DataBodyRange.Formula = "= VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],73,0)"
    End With
    
    Set columna_numero_cuenta_fiscalizacion_sap = tablaDatos.ListColumns.Add
    ' VALIDACIÓN REMUNERACION BASICA TR SAP
    With columna_numero_cuenta_fiscalizacion_sap
        .Name = "VALIDACION DE REMUNERACION BASICA"
        .DataBodyRange.Formula = "= [@[REMUNERACION BASICA TR]]-[@[REMUNERACION BASICA SAP]]"
    End With
    
        With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub

Sub VALIDACION_ENTIDAD_FINANCIERA_EMPLEADO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

    ' VALIDACION_ENTIDAD_FINANCIERA

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR5_PARAMETRIZADA").ListObjects("TR5_PARAMETRIZADA")
    Set columna_entidad_financiera_tr = tablaDatos.ListColumns.Add
    Set columna_entidad_financiera_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN ENTIDAD FINANCIERA TR
    With columna_entidad_financiera_tr
        .Name = "ENTIDAD FINANCIERA TR"
        .DataBodyRange.Formula = "= [@[Entidad Financiera]]"
    End With
    
    ' VALIDACIÓN ENTIDAD FINANCIERA SAP
    With columna_entidad_financiera_sap
        .Name = "ENTIDAD FINANCIERA SAP"
        .DataBodyRange.Formula = "=VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],33,0)"
    End With
    
    Set columna_numero_entidad_financiera_sap_tr = tablaDatos.ListColumns.Add
    ' VALIDACIÓN REMUNERACION BASICA TR SAP
    With columna_numero_entidad_financiera_sap_tr
        .Name = "VALIDACION ENTIDAD FINANCIERA"
        .DataBodyRange.Formula = "=OR(IF(ISNUMBER(FIND(""CREDITO"",[@[ENTIDAD FINANCIERA TR]],1))* ISNUMBER(FIND(""CREDITO"",[@[ENTIDAD FINANCIERA SAP]],1))=1,""VERDADERO"",""FALSO""),IF(ISNUMBER(FIND(""BBVA"",[@[ENTIDAD FINANCIERA TR]],1))* ISNUMBER(FIND(""BBVA"",[@[ENTIDAD FINANCIERA SAP]],1))=1,""VERDADERO"",""FALSO""),IF(ISNUMBER(FIND(""SCOTIABANK"",[@[ENTIDAD FINANCIERA TR]],1))* ISNUMBER(FIND(""SCOTIABANK"",[@[ENTIDAD FINANCIERA SAP]],1))=1,""VERDADERO"",""FALSO""),IF(ISNUMBER(FIND(""INTER"",[@[ENTIDAD FINANCIERA TR]],1))* ISNUMBER(FIND(""INTER"",[@[ENTIDAD FINANCIERA SAP]],1))=1,""VERDADERO"",""FALSO""),IF(ISNUMBER(FIND(""FALABELLA"",[@[ENTIDAD FINANCIERA TR]],1))* ISNUMBER(FIND(""FALABELLA"",[@[ENTIDAD FINANCIERA SAP]],1))=1,""VERDADERO"",""FALSO""))" & _
        ""
    End With
    
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
    End With
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub VALIDAR_CUSSP_TR_6()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


' VALIDAR_CUSSP Macro

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR6_PARAMETRIZADA").ListObjects("TR6_PARAMETRIZADA")
    Set columna_cussp_tr = tablaDatos.ListColumns.Add
    Set columna_cussp_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN CUSSP TR
    
    With columna_cussp_tr
        .Name = "CUSSP TR"
        .DataBodyRange.Formula = "= [@[CUSPP]]"
    End With
    
    ' VALIDACIÓN CUSSP SAP
    
    With columna_cussp_sap
        .Name = "CUSSP SAP"
        .DataBodyRange.Formula = "=VLOOKUP([@[NUMERO DOCUMENTO TR6]],DATA_SAP[[Número de documento]:[SUELDO]],31,0)"
    End With
    
    Set columna_validacion_fecha = tablaDatos.ListColumns.Add
    With columna_validacion_fecha
        .Name = "VALIDACIÓN CUSSP TR-SAP"
        .DataBodyRange.Formula = "= (TRIM([@[CUSSP TR]])= TRIM([@[CUSSP SAP]]))"
        
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
        End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
    End With
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub VALIDAR_TIPO_REGIMEN_TR_6()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

' TIPO_REGIMEN_TR_6 Macro

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR6_PARAMETRIZADA").ListObjects("TR6_PARAMETRIZADA")
    Set columna_tipo_regimen_tr = tablaDatos.ListColumns.Add
    Set columna_tipo_regimen_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN TIPO_REGIMEN_TR_6 TR
    
    With columna_tipo_regimen_tr
        .Name = "TIPO DE REGIMEN TR"
        .DataBodyRange.Formula = "= [@[Tipo de Régimen2]]"
    End With
    
    ' VALIDACIÓN TIPO_REGIMEN_TR_6 SAP
    
    With columna_tipo_regimen_sap
        .Name = "TIPO DE REGIMEN SAP"
        .DataBodyRange.Formula = "=VLOOKUP([@[NUMERO DOCUMENTO TR6]],DATA_SAP[[Número de documento]:[SUELDO]],30,0)"
    End With
    
    Set columna_validacion_tipo_regimen = tablaDatos.ListColumns.Add
    With columna_validacion_tipo_regimen
        .Name = "VALIDACIÓN TIPO DE REGIMEN"
        .DataBodyRange.Formula = "=OR(AND(ISNUMBER(FIND(""SIST"",[@[TIPO DE REGIMEN TR]],1)),ISNUMBER(FIND(""SIST"",[@[TIPO DE REGIMEN SAP]],1))),AND(ISNUMBER(FIND(""HABITAT"",[@[TIPO DE REGIMEN TR]],1)),ISNUMBER(FIND(""HABITAT"",[@[TIPO DE REGIMEN SAP]],1))),AND(ISNUMBER(FIND(""INTEGRA"",[@[TIPO DE REGIMEN TR]],1)),ISNUMBER(FIND(""INTEGRA"",[@[TIPO DE REGIMEN SAP]],1))),AND(ISNUMBER(FIND(""PROFUTURO"",[@[TIPO DE REGIMEN TR]],1)),ISNUMBER(FIND(""PROFUTURO"",[@[TIPO DE REGIMEN SAP]],1))),AND(ISNUMBER(FIND(""PRIMA"",[@[TIPO DE REGIMEN TR]],1)),ISNUMBER(FIND(""PRIMA"",[@[TIPO DE REGIMEN SAP]],1))))" & _
        ""
        
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
        End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
    End With
    End With
   
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

 
End Sub
Sub VALIDAR_REGIMEN_SALUD_TR_6()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

' TIPO_REGIMEN_SALUD_6 Macro

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("TR6_PARAMETRIZADA").ListObjects("TR6_PARAMETRIZADA")
    Set columna_tipo_regimen_salud_tr = tablaDatos.ListColumns.Add
    Set columna_tipo_regimen_salud_sap = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN TIPO_REGIMEN_SALUD_TR_6
    
    With columna_tipo_regimen_salud_tr
        .Name = "TIPO DE REGIMEN SALUD TR"
        .DataBodyRange.Formula = "= [@[Tipo de Régimen]]"
    End With
    
    ' VALIDACIÓN TIPO_REGIMEN_SALUD_TR_6 SAP
    
    With columna_tipo_regimen_salud_sap
        .Name = "TIPO DE REGIMEN SALUD SAP"
        .DataBodyRange.Formula = "=VLOOKUP([@NUMERO DOCUMENTO TR6],DATA_SAP[[Número de documento]:[SUELDO]],24,0)"
    End With
    
    Set columna_validacion_tipo_regimen_salud = tablaDatos.ListColumns.Add
    With columna_validacion_tipo_regimen_salud
        .Name = "VALIDACIÓN TIPO DE REGIMEN TR-SAP"
        .DataBodyRange.Formula = "=IF(AND(IF([@TIPO DE REGIMEN SALUD SAP]=""AFILIADO"",TRUE,FALSE),IF(TRIM([@[TIPO DE REGIMEN SALUD TR]])=""ESSALUD REGULAR"",TRUE,FALSE))=TRUE,""REGISTRAR EPS"","""")"
        
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
        End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.149998474074526
    End With
    End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub











