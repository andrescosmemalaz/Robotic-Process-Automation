Attribute VB_Name = "VALIDACION_CONSTANCIA"
Function FiltrarTablaPorColumnaYValor(tbl As ListObject, columna As Integer, valor As Double)
   tbl.Range.AutoFilter Field:=columna, Criteria1:=valor
   
End Function
Sub DesfiltrarTablaPorColumna(tbl As ListObject, columna As Integer)
    ' Desactiva el filtro de la columna especificada de la tabla especificada
    tbl.Range.AutoFilter Field:=columna
End Sub
Sub VALIDACION_DE_CONSTANCIA_AVANZADA()
Application.DisplayAlerts = False

Dim columna_valor As Integer
Dim celda_recorrida, celda_nueva, rango_origen, rango_final As Range
Dim variable As Range
Dim valor_negativo As Double
Dim valor_positivo As Double
Dim valor_final As Double
Dim contador As Integer
Dim ws As Worksheet
Dim ws_proceso As Worksheet
Set ws = Sheets("VALIDACION")
Set ws_proceso = Sheets("PROCESO")

Dim tbl As ListObject
Set tbl = ws.ListObjects("VALIDACION_CONSTANCIA")

'On Error Resume Next
Dim valor As Double
Dim Filtro  As String

Funcion = "INICIO DE VALIDACION DE CONSTANCIA AVANZADA"
escribirLog Funcion, "Inicio del proceso de validación de constancia avanzada"

origen = Range("VALIDACION_CONSTANCIA[Doc.compensación]")
Set Dic_objeto = CreateObject("Scripting.Dictionary")
For fila = 1 To UBound(origen)
    Dic_objeto(origen(fila, 1)) = 0
Next

ws_proceso.Select

Range("E3").Resize(Dic_objeto.Count) = WorksheetFunction.Transpose(Dic_objeto.Keys)
'---FIN ÚNICOS---

fila_inicial = Range("E4").Row
ultimo = Range("E4").End(xlDown).Row
contador = 0

Columns("E:E").Select
Selection.NumberFormat = "General"

For i = fila_inicial To ultimo
    Sheets("PROCESO").Activate
    Range("E" + CStr(i)).Select
    Filtro = Range("E" + CStr(i)).Value
    
    ws.Select
    columna_valor = 13
    valor = Filtro
    FiltrarTablaPorColumnaYValor tbl, columna_valor, valor
    ws.Select
    fila_inicial = Range("E4").Row
    fila_final = Range("E4").End(xlDown).Row

    Set rango_origen = Range("K" + CStr(fila_inicial) + ":K" + CStr(fila_final)).SpecialCells(xlCellTypeVisible)
    Set rango_validacion = Range("U" + CStr(fila_inicial) + ":U" + CStr(fila_final)).SpecialCells(xlCellTypeVisible)
      
    For Each celda_recorrida In rango_origen
        If celda_recorrida.Value > 0 Then
            valor_positivo = celda_recorrida.Value
        Else
            valor_negativo = celda_recorrida.Value
        End If
    Next
    For Each celda_recorrida In rango_validacion
        'contador = 0
        valor_final = Abs(valor_positivo) - Abs(valor_negativo)
        Range("U" + CStr(celda_recorrida.Row)).Value = valor_final
        'contador = contador + 1
    Next
    DesfiltrarTablaPorColumna ActiveSheet.ListObjects("VALIDACION_CONSTANCIA"), columna_valor
    'contador = contador + 1
    'escribirLog Funcion, "Validación # " & contador & " Realizada"

Next


escribirLog Funcion, "Final del proceso de validación de constancia avanzada"

End Sub

Sub ELIMINAR_DATOS_PROCESO_VALIDACION_AVANZADA()


Funcion = "LIMPIEZA DE DATOS DEL PROCESO DE VALIDACIÓN AVANZADA"
escribirLog Funcion, "Inicio del proceso de limpieza de datos de hoja proceso"

Dim ws As Worksheet
Set ws = Sheets("PROCESO")
Dim ws_final As Worksheet
Set ws_final = Sheets("VALIDACION")
ws.Select

'Seleccionar la columna completa
ws.Columns("E:E").Select

'Eliminar los datos de la columna
Selection.ClearContents


escribirLog Funcion, "Final del proceso de limpieza de datos de hoja proceso"

ws_final.Select
'ws_final.Cells("G").Activate


End Sub




'Function escribirLog(Proceso, mensaje1 As String)
'     Establece la ruta del archivo de log
'    Dim logFile As String
'
'    logFile = "C:\Macros\PROTOTIPO CONSTANCIAS\ARCHIVO LOG\log_" & Format(Now(), "yyyy-mm-dd_hh") & ".txt"
'
'    Call crear_file(logFile)
'
'     Abre el archivo de log en modo de escritura
'    Open logFile For Append As #1
'
'    Print #1, Format(Now(), "yyyy-mm-dd") & "|" & Format(Now(), "hh:mm:ss") & "|" & Proceso & "|" & mensaje1 & "|" & "SATISFACTORIO"
'
'     ... código aquí ...
'     Cierra el archivo de log
'    Close #1
'
'
'
'End Function
'
'Sub crear_file(ruta As String)
'
'    If Dir(ruta, vbDirectory) = "" Then
'
'         Establece la ruta del archivo de log
'        logFile = ruta
'
'         Abre el archivo de log en modo de escritura
'        Open logFile For Append As #1
'        Print #1, "DIA" & "|" & "HORA |PROCESO | COMENTARIO" & "|" & "ESTADO"
'
'        Close #1
'    Else
'        MsgBox "Existe la ruta"
'
'    End If
'
'
'End Sub


