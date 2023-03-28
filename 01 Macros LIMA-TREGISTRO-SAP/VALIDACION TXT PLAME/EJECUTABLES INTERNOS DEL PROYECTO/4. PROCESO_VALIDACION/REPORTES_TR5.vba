Attribute VB_Name = "REPORTES_TR5"
Sub Borrar_Filtro_De_Tabla_tr5()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

  Dim ws As Worksheet
  Dim sTable As String
  Dim loTable As ListObject
  Sheets("TR5_PARAMETRIZADA").Select
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
Sub REPORTE_TR5()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Call REPORTE_FECHA_INICIO

Call REPORTE_TIPO_DE_TRABAJADOR
Call REPORTE_CATEGORIA_OCUPACIONAL
Call REPORTE_NIVEL_EDUCATIVO
Call REPORTE_DISCAPACIDAD
Call REPORTE_DE_SINDICALIZADO
Call REPORTE_REGIMEN_ACUMULATIVO
Call REPORTE_JORNADA_MAXIMA
Call REPORTE_HORARIO_NOCTURNO
Call REPORTE_DOBLE_FISCALIZACION
Call REPORTE_TIPO_CONTRATO
Call REPORTE_NUMERO_DE_CUENTA
Call REPORTE_REMUNERACION_BASICA
Call REPORTE_ENTIDAD_FINANCIERA


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
    
    'Sheets("20100123500_TRA_26042022_110813").Select
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

Sub REPORTE_FECHA_INICIO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


On Error Resume Next
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A:ZZ"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)
    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set FECHA_INGRESO_TR5 = .Find("FECHA INICIO TR")
    Set FECHA_INGRESO_SAP = .Find("FECHA INICIO SAP")
    Set VALIDACION_FECHA_INGRESO = .Find("VALIDACIÓN FECHA INICIO")
    
    Sheets("TR5_PARAMETRIZADA").Select

    'Range("TR6_PARAMETRIZADA[VALIDACIÓN CUSSP TR-SAP]").Select
    'Range("TR6_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_FECHA_INGRESO.Column, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="=#N/D"
    'Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_FECHA_INGRESO.Column, Criteria1:="<>0", Operator:=xlAnd, Criteria2:="<>#N/D"
    
    
    'Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_HORARIO_NOCTURNO_SAP.Column, _
            Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"
    
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("P10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    Range("TR5_PARAMETRIZADA[[FECHA INICIO TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("Q10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.NumberFormat = "m/d/yyyy"
    
    Range("TR5_PARAMETRIZADA[[FECHA INICIO SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("R10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.NumberFormat = "m/d/yyyy"
    Range("TR5_PARAMETRIZADA[[VALIDACIÓN FECHA INICIO]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("S10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
End With
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACIÓN FECHA INICIO]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("P9").Select
ActiveCell.Value = DOC_TR.Value
Range("Q9").Select
ActiveCell.Value = FECHA_INGRESO_TR5.Value
Range("R9").Select
ActiveCell.Value = FECHA_INGRESO_SAP.Value
Range("S9").Select
ActiveCell.Value = VALIDACION_FECHA_INGRESO.Value

Set tbl = Range("P9").CurrentRegion
Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_FECHA_INICIO_TR5"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub REPORTE_TIPO_DE_TRABAJADOR()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)
    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set TIPO_TRABAJADOR_TR5 = .Find("TIPO DE TRABAJADOR TR")
    Set TIPO_TRABAJADOR_SAP = .Find("TIPO DE TRABAJADOR SAP")
    Set VALIDACION_TIPO_DE_TRABAJADOR = .Find("VALIDACIÓN TIPO DE TRABAJADOR")
    
    Sheets("TR5_PARAMETRIZADA").Select

    'Range("TR6_PARAMETRIZADA[VALIDACIÓN CUSSP TR-SAP]").Select
    'Range("TR6_PARAMETRIZADA").Select
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_TIPO_DE_TRABAJADOR.Column, Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"
    
    'Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_HORARIO_NOCTURNO_SAP.Column, _
            Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"
    
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("U10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    Range("TR5_PARAMETRIZADA[[TIPO DE TRABAJADOR TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("V10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Selection.NumberFormat = "m/d/yyyy"
    
    Range("TR5_PARAMETRIZADA[[TIPO DE TRABAJADOR SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("W10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Selection.NumberFormat = "m/d/yyyy"
    Range("TR5_PARAMETRIZADA[[VALIDACIÓN TIPO DE TRABAJADOR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("X10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
End With
Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACIÓN TIPO DE TRABAJADOR]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("U9").Select
ActiveCell.Value = DOC_TR.Value
Range("V9").Select
ActiveCell.Value = TIPO_TRABAJADOR_TR5.Value
Range("W9").Select
ActiveCell.Value = TIPO_TRABAJADOR_SAP.Value
Range("X9").Select
ActiveCell.Value = VALIDACION_TIPO_DE_TRABAJADOR.Value

Set tbl = Range("U9").CurrentRegion
Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_TIPO_DE_TRABAJADOR"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub

Sub REPORTE_CATEGORIA_OCUPACIONAL()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)
    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set TIPO_TRABAJADOR_TR5 = .Find("CATEGORIA OCUPACIONAL TR")
    Set TIPO_TRABAJADOR_SAP = .Find("CATEGORIA OCUPACIONAL SAP")
    Set VALIDACION_TIPO_DE_TRABAJADOR = .Find("VALIDACIÓN CATEGORIA OCUPACIONAL")
    
    Sheets("TR5_PARAMETRIZADA").Select

    
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_TIPO_DE_TRABAJADOR.Column, _
        Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"
    
    
    'CODIGO
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AA10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    
    'COLUMNA TR

    Range("TR5_PARAMETRIZADA[[CATEGORIA OCUPACIONAL TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AB10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Selection.NumberFormat = "m/d/yyyy"
    
    'COLUMNA SAP
    Range("TR5_PARAMETRIZADA[[CATEGORIA OCUPACIONAL SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AC10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    'COLUMNA VALIDACION TR-SAP
    
    Range("TR5_PARAMETRIZADA[[VALIDACIÓN CATEGORIA OCUPACIONAL]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AD10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
End With

Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACIÓN CATEGORIA OCUPACIONAL]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("AA9").Select
ActiveCell.Value = DOC_TR.Value
Range("AB9").Select
ActiveCell.Value = TIPO_TRABAJADOR_TR5.Value
Range("AC9").Select
ActiveCell.Value = TIPO_TRABAJADOR_SAP.Value
Range("AD9").Select
ActiveCell.Value = VALIDACION_TIPO_DE_TRABAJADOR.Value

Set tbl = Range("AA9").CurrentRegion
Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_CATEGORIA_OCUPACIONAL"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub REPORTE_NIVEL_EDUCATIVO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)
    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set NIVEL_EDUCATIVO_TR = .Find("NIVEL EDUCATIVO TR")
    Set NIVEL_EDUCATIVO_SAP_DATOS = .Find("NIVEL EDUCATIVO SAP")
    'MsgBox (NIVEL_EDUCATIVO_SAP_DATOS.Value)
    
    Set VALIDACION_NIVEL_DE_EDUCATIVO = .Find("VALIDACIÓN NIVEL EDUCATIVO")
    
    Sheets("TR5_PARAMETRIZADA").Select

    
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_NIVEL_DE_EDUCATIVO.Column, _
        Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"
    
    'CODIGO
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AF10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    
    'COLUMNA TR

    Range("TR5_PARAMETRIZADA[[NIVEL EDUCATIVO TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AG10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Selection.NumberFormat = "m/d/yyyy"
    
    'COLUMNA SAP
    Range("TR5_PARAMETRIZADA[[NIVEL EDUCATIVO SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AH10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    'COLUMNA VALIDACION TR-SAP
    
    Range("TR5_PARAMETRIZADA[[VALIDACIÓN NIVEL EDUCATIVO]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AI10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
End With

Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACIÓN NIVEL EDUCATIVO]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("AF9").Select
ActiveCell.Value = DOC_TR.Value
Range("AG9").Select
ActiveCell.Value = NIVEL_EDUCATIVO_TR.Value
Range("AH9").Select
ActiveCell.Value = NIVEL_EDUCATIVO_SAP_DATOS.Value
Range("AI9").Select
ActiveCell.Value = VALIDACION_NIVEL_DE_EDUCATIVO.Value

Set tbl = Range("AF9").CurrentRegion
Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_NIVEL_EDUCATIVO"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub REPORTE_DISCAPACIDAD()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)

    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set DISCAPACIDAD_TR = .Find("DISCAPACIDAD TR")
    Set DISCAPACIDAD_SAP = .Find("DISCAPACIDAD SAP")
    Set VALIDACION_DE_DISCAPACIDAD = .Find("VALIDACIÓN DE DISCAPACIDAD")
    
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_DE_DISCAPACIDAD.Column, _
        Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"
        
    'CODIGO
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AK10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    
    'COLUMNA TR

    Range("TR5_PARAMETRIZADA[[DISCAPACIDAD TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AL10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Selection.NumberFormat = "m/d/yyyy"
    
    'COLUMNA SAP
    Range("TR5_PARAMETRIZADA[[DISCAPACIDAD SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AM10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    'COLUMNA VALIDACION TR-SAP
    
    Range("TR5_PARAMETRIZADA[[VALIDACIÓN DE DISCAPACIDAD]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AN10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
End With

Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACIÓN DE DISCAPACIDAD]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("AK9").Select
ActiveCell.Value = DOC_TR.Value
Range("AL9").Select
ActiveCell.Value = DISCAPACIDAD_TR.Value
Range("AM9").Select
ActiveCell.Value = DISCAPACIDAD_SAP.Value
Range("AN9").Select
ActiveCell.Value = VALIDACION_DE_DISCAPACIDAD.Value

Set tbl = Range("AK9").CurrentRegion
Set ws = ActiveSheet

ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_DISCAPACIDAD"


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub REPORTE_DE_SINDICALIZADO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)

    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set SINDICALIZADO_TR = .Find("SINDICALIZADO TR")
    Set SINDICALIZADO_SAP = .Find("SINDICALIZADO SAP")
    Set VALIDACION_DE_SINDICALIZADO = .Find("VALIDACIÓN DE SINDICALIZADO")
    
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_DE_SINDICALIZADO.Column, _
        Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"

    'CODIGO
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AP10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    
    'COLUMNA TR

    Range("TR5_PARAMETRIZADA[[SINDICALIZADO TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AQ10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Selection.NumberFormat = "m/d/yyyy"
    
    'COLUMNA SAP
    Range("TR5_PARAMETRIZADA[[SINDICALIZADO SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AR10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    'COLUMNA VALIDACION TR-SAP
    
    Range("TR5_PARAMETRIZADA[[VALIDACIÓN DE SINDICALIZADO]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AS10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
End With

Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACIÓN DE SINDICALIZADO]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("AP9").Select
ActiveCell.Value = DOC_TR.Value
Range("AQ9").Select
ActiveCell.Value = SINDICALIZADO_TR.Value
Range("AR9").Select
ActiveCell.Value = SINDICALIZADO_SAP.Value
Range("AS9").Select
ActiveCell.Value = VALIDACION_DE_SINDICALIZADO.Value
Set tbl = Range("AP9").CurrentRegion
Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_SINDICALIZADO"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub REPORTE_REGIMEN_ACUMULATIVO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)

    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set REGIMEN_ACUMULATIVO_TR = .Find("REGIMEN ACUMULATIVO TR")
    Set REGIMEN_ACUMULATIVO_SAP = .Find("REGIMEN ACUMULATIVO SAP")
    Set VALIDACION_REGIMEN_A = .Find("VALIDACIÓN REGIMEN ACUMULATIVO")
    
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_REGIMEN_A.Column, _
        Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"

    'CODIGO
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AV10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    
    'COLUMNA TR

    Range("TR5_PARAMETRIZADA[[REGIMEN ACUMULATIVO TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AW10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Selection.NumberFormat = "m/d/yyyy"
    
    'COLUMNA SAP
    Range("TR5_PARAMETRIZADA[[REGIMEN ACUMULATIVO SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AX10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    'COLUMNA VALIDACION TR-SAP
    
    Range("TR5_PARAMETRIZADA[[VALIDACIÓN REGIMEN ACUMULATIVO]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("AY10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
End With

Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACIÓN DE SINDICALIZADO]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("AV9").Select
ActiveCell.Value = DOC_TR.Value
Range("AW9").Select
ActiveCell.Value = REGIMEN_ACUMULATIVO_TR.Value
Range("AX9").Select
ActiveCell.Value = REGIMEN_ACUMULATIVO_SAP.Value
Range("AY9").Select
ActiveCell.Value = VALIDACION_REGIMEN_A.Value
Set tbl = Range("AV9").CurrentRegion
Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_REGIMEN_ACUMULATIVO"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub
Sub REPORTE_JORNADA_MAXIMA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)

    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set JORNADA_MAX_TR = .Find("JORNADA MAXIMA TR")
    Set JORNADA_MAX_SAP_REGIMEN_ESPECIAL_DE_TRABAJO = .Find("JORNADA MAX SAP - REGIMEN ESPECIAL DE TRABAJO")
    Set JORNADA_MAX_JORNADA_DE_TRABAJO = .Find("JORNADA MAX -  JORNADA DE TRABAJO")
    Set VALIDACION_JORNADA_MAXIMA_SAP_ERROR = .Find("VALIDACION JORNADA MAXIMA SAP_ERROR")
    
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_JORNADA_MAXIMA_SAP_ERROR.Column, _
        Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"


    'CODIGO
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BA10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    
    
    'COLUMNA SAP
    Range("TR5_PARAMETRIZADA[[JORNADA MAXIMA TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BB10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    'COLUMNA VALIDACION TR-SAP
    
    Range("TR5_PARAMETRIZADA[[JORNADA MAX SAP - REGIMEN ESPECIAL DE TRABAJO]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BC10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
    Range("TR5_PARAMETRIZADA[[JORNADA MAX -  JORNADA DE TRABAJO]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BD10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
    Range("TR5_PARAMETRIZADA[[VALIDACION JORNADA MAXIMA SAP_ERROR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BE10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Selection.NumberFormat = "m/d/yyyy"
End With


Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACION JORNADA MAXIMA SAP_ERROR]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("BA9").Select
ActiveCell.Value = DOC_TR.Value
Range("BB9").Select
ActiveCell.Value = JORNADA_MAX_TR.Value
Range("BC9").Select
ActiveCell.Value = JORNADA_MAX_SAP_REGIMEN_ESPECIAL_DE_TRABAJO.Value
Range("BD9").Select
ActiveCell.Value = JORNADA_MAX_JORNADA_DE_TRABAJO.Value
Range("BE9").Select
ActiveCell.Value = VALIDACION_JORNADA_MAXIMA_SAP_ERROR.Value

Set tbl = Range("BC9").CurrentRegion
Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_JORNADA_MAXIMA"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub REPORTE_HORARIO_NOCTURNO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)

    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set HORARIO_NOCTURNO_TR = .Find("HORARIO NOCTURNO TR")
    Set VALIDACION_HORARIO_NOCTURNO_SAP = .Find("VALIDACION HORARIO NOCTURNO SAP")
    
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_HORARIO_NOCTURNO_SAP.Column, _
        Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"


    'CODIGO
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BG10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    
    
    'COLUMNA SAP
    Range("TR5_PARAMETRIZADA[[HORARIO NOCTURNO TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BH10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    'COLUMNA VALIDACION TR-SAP
    
    Range("TR5_PARAMETRIZADA[[VALIDACION HORARIO NOCTURNO SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BI10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
End With

Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACION HORARIO NOCTURNO SAP]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("BG9").Select
ActiveCell.Value = DOC_TR.Value
Range("BH9").Select
ActiveCell.Value = HORARIO_NOCTURNO_TR.Value
Range("BI9").Select
ActiveCell.Value = VALIDACION_HORARIO_NOCTURNO_SAP.Value

Set tbl = Range("BG9").CurrentRegion
Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_HORARIO_NOCTURNO"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub REPORTE_DOBLE_FISCALIZACION()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)

    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set SITUACION_ESPECIAL_DEL_TRABAJADOR = .Find("SITUACION ESPECIAL DEL TRABAJADOR")
    Set VAL_FISCA = .Find("VAL FISCA")
    Set VALIDACION_DOBLE_SITUACION_FISCALIZADO = .Find("VALIDACION DOBLE SITUACION FISCALIZADO")
    
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_DOBLE_SITUACION_FISCALIZADO.Column, _
        Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"


    'CODIGO
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BL10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    
    
    'COLUMNA SAP
    Range("TR5_PARAMETRIZADA[[SITUACION ESPECIAL DEL TRABAJADOR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BM10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    'COLUMNA VALIDACION TR-SAP
    
    Range("TR5_PARAMETRIZADA[[VAL FISCA]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BN10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    Range("TR5_PARAMETRIZADA[[VALIDACION DOBLE SITUACION FISCALIZADO]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BO10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
End With

Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACION HORARIO NOCTURNO SAP]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("BL9").Select
ActiveCell.Value = DOC_TR.Value
Range("BM9").Select
ActiveCell.Value = SITUACION_ESPECIAL_DEL_TRABAJADOR.Value
Range("BN9").Select
ActiveCell.Value = VAL_FISCA.Value
Range("BO9").Select
ActiveCell.Value = VALIDACION_DOBLE_SITUACION_FISCALIZADO.Value
Set tbl = Range("BL9").CurrentRegion
Set ws = ActiveSheet

ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_SOBRE_DOBLE_FISCALIZACION"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub

Sub REPORTE_TIPO_CONTRATO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)

    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set TIPO_DE_CONTRATO_TR = .Find("TIPO DE CONTRATO TR")
    Set TIPO_DE_CONTRATO_SAP = .Find("TIPO DE CONTRATO SAP")
    Set VALIDACION_DEL_TIPO_DE_CONTRATO = .Find("VALIDACION DEL TIPO DE CONTRATO")
    
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_DEL_TIPO_DE_CONTRATO.Column, _
        Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"


    'CODIGO
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BR10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    'Range("TR6_PARAMETRIZADA").AutoFilter Field:=25, Criteria1:="FALSO"
    
    
    'COLUMNA SAP
    Range("TR5_PARAMETRIZADA[[TIPO DE CONTRATO TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BS10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    'COLUMNA VALIDACION TR-SAP
    
    Range("TR5_PARAMETRIZADA[[TIPO DE CONTRATO SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BT10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    Range("TR5_PARAMETRIZADA[[VALIDACION DEL TIPO DE CONTRATO]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BU10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
End With

Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACION DEL TIPO DE CONTRATO]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("BR9").Select
ActiveCell.Value = DOC_TR.Value
Range("BS9").Select
ActiveCell.Value = TIPO_DE_CONTRATO_TR.Value
Range("BT9").Select
ActiveCell.Value = TIPO_DE_CONTRATO_SAP.Value
Range("BU9").Select
ActiveCell.Value = VALIDACION_DEL_TIPO_DE_CONTRATO.Value
Set tbl = Range("BR9").CurrentRegion
Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_TIPO_CONTRATO"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub REPORTE_NUMERO_DE_CUENTA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)

    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set NUMERO_DE_CUENTA_TR = .Find("NUMERO DE CUENTA TR")
    Set NUMERO_DE_CUENTA_SAP = .Find("NUMERO DE CUENTA SAP")
    Set VALIDACION_DE_NUMERO_DE_CUENTA = .Find("VALIDACION DE NUMERO DE CUENTA")
    
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_DE_NUMERO_DE_CUENTA.Column, _
        Criteria1:="=NO ESTAN IGUALES LOS CARACTERES", Operator:=xlOr, Criteria2:="=#N/D"


    'CODIGO
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BW10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    
    'COLUMNA SAP
    Range("TR5_PARAMETRIZADA[[NUMERO DE CUENTA TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BX10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.NumberFormat = "0.00"
    Selection.NumberFormat = "0.0"
    Selection.NumberFormat = "0"
    
    'COLUMNA VALIDACION TR-SAP
    
    Range("TR5_PARAMETRIZADA[[NUMERO DE CUENTA SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BY10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    Range("TR5_PARAMETRIZADA[[VALIDACION DE NUMERO DE CUENTA]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("BZ10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
End With

Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACION DE NUMERO DE CUENTA]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("BW9").Select
ActiveCell.Value = DOC_TR.Value
Range("BX9").Select
ActiveCell.Value = NUMERO_DE_CUENTA_TR.Value
Range("BY9").Select
ActiveCell.Value = NUMERO_DE_CUENTA_SAP.Value
Range("BZ9").Select
ActiveCell.Value = VALIDACION_DE_NUMERO_DE_CUENTA.Value
Set tbl = Range("BW9").CurrentRegion

Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_NUMERO_DE_CUENTA"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub REPORTE_REMUNERACION_BASICA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"

With Sheets("TR5_PARAMETRIZADA").Range(columnas)

    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set REMUNERACION_BASICA_TR = .Find("REMUNERACION BASICA TR")
    Set REMUNERACION_BASICA_SAP = .Find("REMUNERACION BASICA SAP")
    Set VALIDACION_DE_REMUNERACION_BASICA = .Find("VALIDACION DE REMUNERACION BASICA")
    
    
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_DE_REMUNERACION_BASICA.Column, _
        Criteria1:="<>0", Operator:=xlOr, Criteria2:="=#N/D"

    
    
        'CODIGO
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("CB10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    
    'COLUMNA SAP
    Range("TR5_PARAMETRIZADA[[REMUNERACION BASICA TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("CC10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.NumberFormat = "0.00"
    Selection.NumberFormat = "0.0"
    Selection.NumberFormat = "0"
    
    'COLUMNA VALIDACION TR-SAP
    
    Range("TR5_PARAMETRIZADA[[REMUNERACION BASICA SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("CD10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.NumberFormat = "0.00"
    Selection.NumberFormat = "0.0"
    Selection.NumberFormat = "0"
    
    Range("TR5_PARAMETRIZADA[[VALIDACION DE REMUNERACION BASICA]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("CE10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
End With


Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACION DE NUMERO DE CUENTA]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("CB9").Select
ActiveCell.Value = DOC_TR.Value
Range("CC9").Select
ActiveCell.Value = REMUNERACION_BASICA_TR.Value
Range("CD9").Select
ActiveCell.Value = REMUNERACION_BASICA_SAP.Value
Range("CE9").Select
ActiveCell.Value = VALIDACION_DE_REMUNERACION_BASICA.Value
Set tbl = Range("CB9").CurrentRegion
Set ws = ActiveSheet
ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_REMUNERACION_BASICA"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub REPORTE_ENTIDAD_FINANCIERA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
Dim Value_column As Integer
Sheets("TR5_PARAMETRIZADA").Select
columnas = "A10:BQ10"


With Sheets("TR5_PARAMETRIZADA").Range(columnas)

    Set DOC_TR = .Find("NUMERO DOCUMENTO TR5")
    Set ENTIDAD_FINANCIERA_TR = .Find("ENTIDAD FINANCIERA TR")
    Set ENTIDAD_FINANCIERA_SAP = .Find("ENTIDAD FINANCIERA SAP")
    Set VALIDACION_ENTIDAD_FINANCIERA = .Find("VALIDACION ENTIDAD FINANCIERA")
    
    Range("TR5_PARAMETRIZADA").AutoFilter Field:=VALIDACION_ENTIDAD_FINANCIERA.Column, _
        Criteria1:="=FALSO", Operator:=xlOr, Criteria2:="=#N/D"

    
    
        'CODIGO
    Range("TR5_PARAMETRIZADA[[NUMERO DOCUMENTO TR5]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("CG10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    
    'COLUMNA SAP
    Range("TR5_PARAMETRIZADA[[ENTIDAD FINANCIERA TR]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("CH10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.NumberFormat = "0.00"
    Selection.NumberFormat = "0.0"
    Selection.NumberFormat = "0"
    
    'COLUMNA VALIDACION TR-SAP
    
    Range("TR5_PARAMETRIZADA[[ENTIDAD FINANCIERA SAP]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("CI10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Selection.NumberFormat = "0.00"
    Selection.NumberFormat = "0.0"
    Selection.NumberFormat = "0"
    
    Range("TR5_PARAMETRIZADA[[VALIDACION ENTIDAD FINANCIERA]]").SpecialCells(xlCellTypeVisible).Copy
    Sheets("REPORTE").Select
    Range("CJ10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
End With


Sheets("TR5_PARAMETRIZADA").Select
Range("TR5_PARAMETRIZADA[[VALIDACION ENTIDAD FINANCIERA]]").Select
ActiveSheet.ShowAllData
Dim tbl As Range
Dim ws As Worksheet
Sheets("REPORTE").Select
Range("CG9").Select
ActiveCell.Value = DOC_TR.Value
Range("CH9").Select
ActiveCell.Value = ENTIDAD_FINANCIERA_TR.Value
Range("CI9").Select
ActiveCell.Value = ENTIDAD_FINANCIERA_SAP.Value
Range("CJ9").Select
ActiveCell.Value = VALIDACION_ENTIDAD_FINANCIERA.Value
Set tbl = Range("CG9").CurrentRegion
Set ws = ActiveSheet

ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "REPORTE_ENTIDAD_FINANCIERA"


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub






