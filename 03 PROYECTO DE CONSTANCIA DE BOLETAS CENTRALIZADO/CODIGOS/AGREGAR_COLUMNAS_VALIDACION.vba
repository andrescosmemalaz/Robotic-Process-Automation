Attribute VB_Name = "AGREGAR_COLUMNAS_VALIDACION"
Sub CREACION_COLUMNAS_VALIDACION_FINAL()
Call ELIMINAR_COLUMNAS
Call VALIDAR_RUTA_PDF
Call AGREGAR_BANCO
Call COLUMNA_UNIDAD
'Call VALIDAR_REFERENCIA
'Call VALIDAR_FECHA
'Call VALIDAR_NOMBRE_CONSTANCIA_PDF
'Call VALIDAR_MONTO_PDF
Call VALIDAR_CONCILIACION_FINAL
End Sub

Sub AGREGAR_BANCO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


' VALIDAR_NOMBRE_CONSTANCIA

    'Funcion = "AGREGAR BANCO"
    'escribirLog Funcion, "Inicio del proceso de busqueda del banco  del  PDF"

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    Set columna_nombre_banco = tablaDatos.ListColumns.Add
    'Set columna_referencia_pdf = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN REFERENCIA SAP
    With columna_nombre_banco
        .Name = "BANCO DE PROCEDENCIA CONSTANCIA"
        .DataBodyRange.Formula = "=IFERROR(VLOOKUP([@Texto],BASE_DE_DATOS_CONSTANCIAS_PDF,10,0),""NO FUE ENCONTRADO"")"
        'Selection.NumberFormat = "m/d/yyyy"
    End With
    
    'escribirLog Funcion, "Fin del proceso de busqueda del banco del PDF"
    
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub


Sub VALIDAR_RUTA_PDF()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


' VALIDAR_NOMBRE_CONSTANCIA

    'Funcion = "CREAR RUTA PDF"
    'escribirLog Funcion, "Inicio del proceso de busqueda de la  RUTA del  PDF"

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    Set columna_nombre_constancia_pdf = tablaDatos.ListColumns.Add
    'Set columna_referencia_pdf = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN REFERENCIA SAP
    With columna_nombre_constancia_pdf
        .Name = "RUTA PDF"
        .DataBodyRange.Formula = "=IFERROR(VLOOKUP([@Texto],BASE_DE_DATOS_CONSTANCIAS_PDF,5,0),""NO FUE ENCONTRADO"")"
        'Selection.NumberFormat = "m/d/yyyy"
    End With
    
    'escribirLog Funcion, "Fin del proceso de busqueda de la  RUTA del  PDF"
    
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub COLUMNA_UNIDAD()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

' AGREGAR COLUMNA DE LA UNIDAD
    'Funcion = "CREAR COLUMNAS DE UNIDAD"
    'escribirLog Funcion, "Inicio del proceso de validación de UNIDAD"
    
    Dim tablaDatos As ListObject
    Dim COLUMNA_UNIDAD As ListColumn, columna_ls_comparativo As ListColumn
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    
    Set string_unidad = tablaDatos.ListColumns.Add
    With string_unidad
        .Name = "NOMBRE DE UNIDAD"
        .DataBodyRange.Formula = "=IF([@Sociedad]&[@División]=""70107101"",""NEXA PERU_CERRO LINDO"",IF([@Sociedad]&[@División]=""70107104"",""NEXA PERU_LIMA"",IF([@Sociedad]=""7022"",""ATACOCHA"",IF([@Sociedad]=""7042"",""CAJAMARQUILLA"",IF([@Sociedad]=""7053"",""EL PORVENIR"",IF([@Sociedad]=""7056"",""PAMPA COBRE"",""OTROS""))))))"
        'Selection.NumberFormat = "mm/dd/yyyy"
    End With

    'escribirLog Funcion, "Creación columna validación UNIDAD"
    

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
End Sub



Sub VALIDAR_REFERENCIA()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Funcion = "CREAR VALIDACION REFERENCIA"
'escribirLog Funcion, "Inicio del proceso de validación de referencia"

' VALIDAR_REFERENCIA

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    Set columna_referencia_Sap = tablaDatos.ListColumns.Add
    Set columna_referencia_pdf = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN REFERENCIA SAP
    With columna_referencia_Sap
        .Name = "REFERENCIA REPORTE SAP"
        .DataBodyRange.Formula = "= [@[Texto]]"
        'Selection.NumberFormat = "m/d/yyyy"
        
    End With
    escribirLog Funcion, "Creación de columna de Referencia Reporte SAP"
    
    ' VALIDACIÓN REFERENCIA PDF
    With columna_referencia_pdf
        .Name = "REFERENCIA REPORTE PDF"
        .DataBodyRange.Formula = "=IFERROR(VLOOKUP([@Texto],BASE_DE_DATOS_CONSTANCIAS_PDF,1,0),""NO FUE ENCONTRADO"")"
    End With
    
    escribirLog Funcion, "Creación de columna de Referencia Reporte PDF"
    
    Set columna_validacion_referencia = tablaDatos.ListColumns.Add
    With columna_validacion_referencia
        .Name = "VALIDACIÓN REFERENCIA"
        .DataBodyRange.Formula = "=[@[REFERENCIA REPORTE SAP]]=[@[REFERENCIA REPORTE PDF]]"
    End With
    escribirLog Funcion, "Creación de columna de Validacion de Referencia"

    'escribirLog Funcion, "Inicio del proceso de validación de referencia"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


Sub VALIDAR_NOMBRE_CONSTANCIA_PDF()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Funcion = "CREAR VALIDACION DE CONSTANCIA PDF"
escribirLog Funcion, "Inicio del proceso de validación de constancia PDF"

' VALIDAR_NOMBRE_CONSTANCIA

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    Set columna_nombre_constancia_pdf = tablaDatos.ListColumns.Add
    'Set columna_referencia_pdf = tablaDatos.ListColumns.Add
    
    ' VALIDACIÓN REFERENCIA SAP
    With columna_nombre_constancia_pdf
        .Name = "NOMBRE DE LA CONSTANCIA"
        .DataBodyRange.Formula = "=IFERROR(VLOOKUP([@Texto],BASE_DE_DATOS_CONSTANCIAS_PDF,2,0),""NO FUE ENCONTRADO"")"
        'Selection.NumberFormat = "m/d/yyyy"
    End With
    
    escribirLog Funcion, "FIn del proceso de validación de constancia PDF"
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub Nombre_unidad()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

    Funcion = "CREAR NOMBRE DETALLE UNIDAD "
    escribirLog Funcion, "Inicio del proceso de CREACION DE DETALLE DE UNIDAD"
    
    Dim tablaDatos As ListObject
    Dim COLUMNA_UNIDAD As ListColumn, columna_ls_comparativo As ListColumn
    
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    Set columna_valor_unidad = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA").ListColumns(3).DataBodyRange
    Set COLUMNA_UNIDAD = tablaDatos.ListColumns.Add
    
    With COLUMNA_UNIDAD
        .Name = "NOMBRE_UNIDAD"
        If columna_valor_unidad.Value = "7010" Then
        .DataBodyRange.Formula = "NEXA PERÚ"
        ElseIf performance = "7053" Then
        .DataBodyRange.Formula = "EL PORVENIR"
        ElseIf performance = "7042" Then
        .DataBodyRange.Formula = "CAJAMARQUILLA"
        ElseIf performance = "7022" Then
        .DataBodyRange.Formula = "ATACOCHA"
        ElseIf performance = "7056" Then
        .DataBodyRange.Formula = "PAMPA COBRE"
        Else
        .DataBodyRange.Formula = "NO DETALLA"
        End If
        'Selection.NumberFormat = "m/d/yyyy"
    End With
    
    escribirLog Funcion, "Fin del proceso de CREACION DE DETALLE DE UNIDAD"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
End Sub



Sub VALIDAR_FECHA()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Funcion = "CREAR COLUMNA FECHAS "
escribirLog Funcion, "Inicio del proceso de añadir detalle de columna de fechas"


' VALIDAR_COLUMNA_FECHA
    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    
    'Set columna_ejercicio_mes_sap = tablaDatos.ListColumns.Add
''     VALIDACIÓN REFERENCIA SAP
''    With columna_ejercicio_mes_sap
''        .Name = "FECHA CONTABILIZACION SAP"
''        .DataBodyRange.Formula = "=DATE(YEAR([@Fecha compensación]),MONTH([@Fecha compensación]),DAY([@Fecha compensación]))"
''        Selection.NumberFormat = "mm/dd/yyyy"
''    End With
''
    
    ' VALIDACIÓN REFERENCIA SAP

    Set columna_ejercicio_mes_pdf = tablaDatos.ListColumns.Add
    With columna_ejercicio_mes_pdf
        .Name = "FECHA PROCESO PDF"
        .DataBodyRange.Formula = "=IFERROR(VLOOKUP([@Texto],BASE_DE_DATOS_CONSTANCIAS_PDF,5,0),""NO FUE ENCONTRADO"")"
        Selection.NumberFormat = "mm/dd/yyyy"
    End With

    'Funcion = "CREAR COLUMNA FECHAS "
    

'    Set validacion_ejercicio = tablaDatos.ListColumns.Add
'    ' VALIDACIÓN REFERENCIA SAP
'    With validacion_ejercicio
'        .Name = "VALIDACION FECHAS"
'        .DataBodyRange.Formula = "=[@[FECHA CONTABILIZACION SAP]]=[@[FECHA PROCESO PDF]]"
'        'Selection.NumberFormat = "mm/dd/yyyy"
'    End With
    
    
   escribirLog Funcion, "Fin del proceso de añadir detalle de columna de fechas"
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


Sub VALIDAR_MONTO_PDF()

Application.CutCopyMode = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
    
    Funcion = "CREAR COLUMNA MONTOS PDF "
    escribirLog Funcion, "Inicio del proceso de validar monto de pdf"

    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    Set columna_monto_sap = tablaDatos.ListColumns.Add
    Set columna_monto_pdf = tablaDatos.ListColumns.Add
    
    With columna_monto_sap
        .Name = "MONTO SAP"
        .DataBodyRange.Formula = "= [@[Importe en moneda local]]"
        Selection.NumberFormat = "0.00"
    End With
    
    escribirLog Funcion, "añadir columna de monto de sap"
    
    With columna_monto_pdf
        .Name = "MONTO  PDF"
        .DataBodyRange.Formula = "=IFERROR(VLOOKUP([@Texto],BASE_DE_DATOS_CONSTANCIAS_PDF,7,0),""NO FUE ENCONTRADO"")"
        '.Selection.NumberFormat = "0.00"
        
    End With
    
    escribirLog Funcion, "añadir columna de monto de pdf"
    
    Set columna_validacion_montos = tablaDatos.ListColumns.Add
    With columna_validacion_montos
        .Name = "VALIDACION MONTOS"
        .DataBodyRange.Formula = "=IFERROR(ABS([@[MONTO SAP]])=ABS([@[MONTO  PDF]]),""NO FUE ENCONTRADO"")"
        
    End With
    
    escribirLog Funcion, "añadir columna de validación de montos pdf y sap"
    
    Set columna_validacion_constancia = tablaDatos.ListColumns.Add
    With columna_validacion_constancia
        .Name = "VALIDACION CONSTANCIA FINAL"
        .DataBodyRange.Formula = "=IF([@[VALIDACION MONTOS]]=TRUE,""CONFORME"",IF([@[VALIDACION MONTOS]]=FALSE,""MONTOS NO CUADRA"",""NO EXISTE DOCUMENTO EN COMPARTIDO""))"
        
    End With
    
    escribirLog Funcion, "añadir columna de validación de constancia final"
    
    escribirLog Funcion, "Fin del proceso de validar monto de pdf"
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


Sub VALIDAR_CONCILIACION_FINAL()

Application.CutCopyMode = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
    
    'Funcion = "CREAR COLUMNA DE CONCILIACION FINAL"
    'escribirLog Funcion, "Inicio del proceso de concialiación final"
    
    Dim tablaDatos As ListObject
    Dim columna_hb_comparativo As ListColumn, columna_ls_comparativo As ListColumn
    
    Dim myRow As ListRow
    Dim intRows As Integer
    Set tablaDatos = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    Set columna_validacion_constancia_final = tablaDatos.ListColumns.Add
    Set columna_validacion_conciliacion = tablaDatos.ListColumns.Add
   
    'Set columna_monto_pdf = tablaDatos.ListColumns.Add
    
     With columna_validacion_constancia_final
        .Name = "VALIDACION CONSTANCIA FINAL"
        .DataBodyRange.Formula = "=IF(IFERROR(ABS([@[Importe en moneda local]])=ABS(IFERROR(VLOOKUP([@Texto],'PROYECTO VALIDACION CONSTANCIA.xlsm'!BASE_DE_DATOS_CONSTANCIAS_PDF[#Data],9,0),""NO FUE ENCONTRADO"")),""NO FUE ENCONTRADO"")=TRUE,""CONFORME"",IF(IFERROR(ABS([@[Importe en moneda local]])=ABS(IFERROR(VLOOKUP([@Texto],'PROYECTO VALIDACION CONSTANCIA.xlsm'!BASE_DE_DATOS_CONSTANCIAS_PDF[#Data]" & _
        ",9,0),""NO FUE ENCONTRADO"")),""NO FUE ENCONTRADO"")=FALSE,""CONFORME"",""NO EXISTE DOCUMENTO EN COMPARTIDO""))" & _
        "" 'Selection.NumberFormat = "0.00"
    End With
    
    With columna_validacion_conciliacion
        .Name = "VALIDACION CONCILIACION FINAL"
        .DataBodyRange.Formula = "=IF([@[VALIDACION DE CONSTANCIA]]=0,""CONFORME"",""PENDIENTE DE CONCILIACION"")"
        'Selection.NumberFormat = "0.00"
    End With
    'escribirLog Funcion, "Fin del proceso de concialiación final"
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub



Sub ELIMINAR_COLUMNAS()
    Application.DisplayAlerts = False
    
    Funcion = "ELIMINAR COLUMNAS VALIDADAS EXISTENTES"
    escribirLog Funcion, "Inicio del proceso de columnas de validación"
    
    
    Dim tabla As ListObject
    Dim nombresColumnas As Variant
    
    Set tabla = ActiveSheet.ListObjects("VALIDACION_CONSTANCIA")
    nombresColumnas = Array("RUTA PDF", "REFERENCIA REPORTE SAP", "REFERENCIA REPORTE PDF", _
                            "VALIDACIÓN REFERENCIA", "FECHA PROCESO PDF", "NOMBRE DE LA CONSTANCIA", _
                            "NOMBRE DE UNIDAD", "CODIGO UNIDAD PDF", "CODIGO UNIDAD SAP", _
                            "VALIDACION UNIDAD", "MONTO SAP", "MONTO  PDF", "VALIDACION MONTOS", _
                            "VALIDACION CONSTANCIA FINAL", "VALIDACION CONCILIACION FINAL")
    
    For Each nombreColumna In nombresColumnas
        EliminarColumna tabla, nombreColumna
    Next
    
    escribirLog Funcion, "Fin del proceso de columnas de validación"
    
End Sub


Private Sub EliminarColumna(ByRef tabla As ListObject, ByVal nombreColumna As String)
    On Error Resume Next

    tabla.ListColumns(nombreColumna).Delete
End Sub



