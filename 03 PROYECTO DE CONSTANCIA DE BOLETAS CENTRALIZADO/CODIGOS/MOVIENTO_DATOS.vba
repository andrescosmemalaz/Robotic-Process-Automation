Attribute VB_Name = "MOVIENTO_dATOS"

Function limpiar_tabla(tabla As ListObject)
tabla.DataBodyRange.Clear
End Function


Sub LIMPIAR_RANGO_COLUMNAS()
Funcion = "ELIMINAR REGISTROS DE PDF EN LA TABLA DEL SISTEMA"
escribirLog Funcion, "Inicio del proceso de eliminar datos de tabla"

On Error Resume Next
'Declara la variable de la tabla
Dim tabla As ListObject
Dim ws_data_pdf As Worksheet
Set ws_data_pdf = Sheets("VALIDACION")

'Establece la tabla
Set tabla = ws_data_pdf.ListObjects("VALIDACION_CONSTANCIA")

'Elimina todos los registros de la tabla
tabla.DataBodyRange.ClearContents

escribirLog Funcion, "Fin del proceso de eliminar datos de tabla"

End Sub
Sub clearColumnFill(sheetName As String, tableName As String, columnName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim tbl As ListObject
    Set tbl = ws.ListObjects(tableName)

    Dim col As ListColumn
    Set col = tbl.ListColumns(columnName)

    col.Range.Interior.ColorIndex = xlNone
End Sub

Sub MOVIMIENTO_COLUMNAS_MASIVAS()

    Funcion = "MOVIMIENTO DE COLUMNAS MASIVAS"
    escribirLog Funcion, "Inicio del movimiento de columnas masivas"

    'Declara las variables
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim rngOrigen As Range
    Dim rngDestino As Range
    Dim ws_reporte_sap As Worksheet
    Set ws_reporte_sap = Sheets("REPORTE_SAP")
    Dim ws_validacion As Worksheet
    Set ws_validacion = Sheets("VALIDACION")
    
    
    'LIMPIAR_RANGO_COLUMNAS "VALIDACION", "VALIDACION_CONSTANCIA"
    
    'Asigna las variables a los objetos correspondientes
    Set wsOrigen = Worksheets("REPORTE_SAP") 'Cambia "Hoja1" por el nombre de la hoja de origen
    Set wsDestino = Worksheets("VALIDACION") 'Cambia "Hoja2" por el nombre de la hoja de destino
    wsOrigen.Select
    'Establece el rango de origen y destino iniciales
    Set rngOrigen = wsOrigen.ListObjects("DATA_SAP_FBLN").ListColumns("Cuenta").DataBodyRange 'Cambia "Tabla1" y "Columna1" por el nombre de tu tabla y columna de origen, respectivamente
    Set rngDestino = wsDestino.ListObjects("VALIDACION_CONSTANCIA").ListColumns("Cuenta").DataBodyRange 'Cambia "Tabla2" y "Columna2" por el nombre de tu tabla y columna de destino, respectivamente
    
    'Recorre las primeras 5 columnas de la tabla de origen
    For i = 1 To 19
        'Copia la columna desde la tabla de origen a la tabla de destino
        rngOrigen.Copy
        
        rngDestino.PasteSpecial 'xlPasteValues

        'Incrementa el rango de origen y destino para copiar la siguiente columna
        Set rngOrigen = rngOrigen.Offset(0, 1)
        Set rngDestino = rngDestino.Offset(0, 1)
    Next i
    ws_validacion.Select
    
    escribirLog Funcion, "Final del movimiento de columnas masivas"
End Sub

