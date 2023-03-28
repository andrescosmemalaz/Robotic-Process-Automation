Attribute VB_Name = "PINTAR_VALIDACION"

Sub PINTAR_VALORES_VALIDACION()
Call FILTRAR_TABLA_VAL_CONSTANCIA
Call FILTRAR_TABLA_VAL_CONSTANCIA_FINAL_NEGATIVO
Call FILTRAR_TABLA_VAL_CONSTANCIA_FINAL_POSITIVO
Call FILTRAR_TABLA_VAL_CONCILIACION_FINAL_NEGATIVO
Call FILTRAR_TABLA_VAL_CONCILIACION_FINAL_POSITIVO
End Sub
Sub DesfiltrarTablaPorColumna_v2(tbl As ListObject, columna As Integer)
    ' Desactiva el filtro de la columna especificada de la tabla especificada
    tbl.Range.AutoFilter Field:=columna
End Sub


Sub FILTRAR_TABLA_VAL_CONSTANCIA()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("VALIDACION")
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    Dim columna_valor As Integer

    
    columna_valor = GetColumnNumber(tbl, "VALIDACION DE CONSTANCIA")
    'MsgBox (datetype(columna_valor))
    
    With ws
        'Filtrar los valores diferentes a 0 en la columna "VALIDACION DE CONSTANCIA"
        .ListObjects("VALIDACION_CONSTANCIA").Range.AutoFilter Field:=20, Criteria1:="<>0"
        'Pintar de rojo los valores filtrados
        .ListObjects("VALIDACION_CONSTANCIA").ListColumns("VALIDACION DE CONSTANCIA").DataBodyRange.SpecialCells(xlCellTypeVisible).Interior.Color = RGB(255, 0, 0)
    End With
    
    DesfiltrarTablaPorColumna_v2 tbl, columna_valor
    Exit Sub
    
ErrorHandler:
    'MsgBox "Ha ocurrido un error en el proceso de filtrado de la tabla. Por favor, inténtelo de nuevo más tarde.", vbCritical
    DesfiltrarTablaPorColumna_v2 tbl, columna_valor
End Sub

Sub FILTRAR_TABLA_VAL_CONSTANCIA_FINAL_NEGATIVO()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("VALIDACION")
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    Dim columna_valor As Integer
    
    columna_valor = GetColumnNumber(tbl, "VALIDACION CONSTANCIA FINAL")
    
    With ws
        'Filtrar los valores diferentes a 0 en la columna "VALIDACION DE CONSTANCIA"
        .ListObjects("VALIDACION_CONSTANCIA").Range.AutoFilter Field:=columna_valor, Criteria1:="NO EXISTE DOCUMENTO EN COMPARTIDO"
        'Pintar de rojo los valores filtrados
        .ListObjects("VALIDACION_CONSTANCIA").ListColumns(columna_valor).DataBodyRange.SpecialCells(xlCellTypeVisible).Interior.Color = RGB(255, 0, 0)
        .ListObjects("VALIDACION_CONSTANCIA").Range.AutoFilter Field:=columna_valor, Criteria1:="MONTOS NO CUADRA"
        'Pintar de rojo los valores filtrados
        .ListObjects("VALIDACION_CONSTANCIA").ListColumns(columna_valor).DataBodyRange.SpecialCells(xlCellTypeVisible).Interior.Color = RGB(255, 0, 0)
        
    End With
    
    DesfiltrarTablaPorColumna_v2 tbl, columna_valor
    
    Exit Sub
    
ErrorHandler:
    'MsgBox "Ha ocurrido un error en el proceso de filtrado de la tabla. Por favor, inténtelo de nuevo más tarde.", vbCritical
    DesfiltrarTablaPorColumna_v2 tbl, columna_valor
End Sub
Sub FILTRAR_TABLA_VAL_CONSTANCIA_FINAL_POSITIVO()
On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("VALIDACION")
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    Dim columna_valor As Integer
    
    columna_valor = GetColumnNumber(tbl, "VALIDACION CONSTANCIA FINAL")
    
    With ws
        'Filtrar los valores diferentes a 0 en la columna "VALIDACION DE CONSTANCIA"
        .ListObjects("VALIDACION_CONSTANCIA").Range.AutoFilter Field:=columna_valor, Criteria1:="CONFORME"
        'Pintar de rojo los valores filtrados
        .ListObjects("VALIDACION_CONSTANCIA").ListColumns(columna_valor).DataBodyRange.SpecialCells(xlCellTypeVisible).Interior.Color = RGB(0, 255, 0)
    End With
    
    DesfiltrarTablaPorColumna_v2 tbl, columna_valor
    Exit Sub
    
ErrorHandler:
    'MsgBox "Ha ocurrido un error en el proceso de filtrado de la tabla. Por favor, inténtelo de nuevo más tarde.", vbCritical
    DesfiltrarTablaPorColumna_v2 tbl, columna_valor
End Sub
Sub FILTRAR_TABLA_VAL_CONCILIACION_FINAL_NEGATIVO()
On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("VALIDACION")
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    Dim columna_valor As Integer
    
    columna_valor = GetColumnNumber(tbl, "VALIDACION CONCILIACION FINAL")
    
    With ws
        'Filtrar los valores diferentes a 0 en la columna "VALIDACION DE CONSTANCIA"
        .ListObjects("VALIDACION_CONSTANCIA").Range.AutoFilter Field:=columna_valor, Criteria1:="PENDIENTE DE CONCILIACION"
        'Pintar de rojo los valores filtrados
        .ListObjects("VALIDACION_CONSTANCIA").ListColumns(columna_valor).DataBodyRange.SpecialCells(xlCellTypeVisible).Interior.Color = RGB(255, 0, 0)
    End With
    
    DesfiltrarTablaPorColumna_v2 tbl, columna_valor
    Exit Sub
    
ErrorHandler:
    'MsgBox "Ha ocurrido un error en el proceso de filtrado de la tabla. Por favor, inténtelo de nuevo más tarde.", vbCritical
    DesfiltrarTablaPorColumna_v2 tbl, columna_valor
End Sub
Sub FILTRAR_TABLA_VAL_CONCILIACION_FINAL_POSITIVO()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("VALIDACION")
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("VALIDACION").ListObjects("VALIDACION_CONSTANCIA")
    Dim columna_valor As Integer
    
    columna_valor = GetColumnNumber(tbl, "VALIDACION CONCILIACION FINAL")
    
    With ws
        'Filtrar los valores diferentes a 0 en la columna "VALIDACION DE CONSTANCIA"
        .ListObjects("VALIDACION_CONSTANCIA").Range.AutoFilter Field:=columna_valor, Criteria1:="CONFORME"
        'Pintar de rojo los valores filtrados
        .ListObjects("VALIDACION_CONSTANCIA").ListColumns(columna_valor).DataBodyRange.SpecialCells(xlCellTypeVisible).Interior.Color = RGB(0, 255, 0)
    End With
    
    DesfiltrarTablaPorColumna_v2 tbl, columna_valor
    Exit Sub
    
ErrorHandler:
    'MsgBox "Ha ocurrido un error en el proceso de filtrado de la tabla. Por favor, inténtelo de nuevo más tarde.", vbCritical
    DesfiltrarTablaPorColumna_v2 tbl, columna_valor
End Sub

Function GetColumnNumber(table As ListObject, columnName As String) As Long
    Dim col As ListColumn
    For Each col In table.ListColumns
        If col.Name = columnName Then
            GetColumnNumber = col.Index
            Exit Function
        End If
    Next col
End Function
