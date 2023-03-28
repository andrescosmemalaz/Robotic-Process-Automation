Attribute VB_Name = "COPIAR_DESDE_DATA_PRINCIPAL"
Sub COPIAR_DATA_PRINCIPAL_TRANSACCIONES()

Dim valorCelda As String
Dim filaActual As Integer
Dim ws As Worksheet
Dim tbl As ListObject
Dim col As ListColumn
Dim rng As Range

' Define la hoja de trabajo "Proceso" como la hoja de trabajo activa
Sheets("PROCESO").Activate
Dim tabla As ListObject
Dim nuevaCelda As Range
Set tabla = Sheets("PL preliminar").ListObjects("MODELADO_PL_PRELIMINAR") 'Cambiar "Tabla1" por el nombre de tu tabla
'Agrega una celda después de la última fila de la tabla
'Set nuevaCelda = tabla.ListRows.Add.Range.Cells(1, 1)

' Recorre todas las filas de la columna A en "Proceso"
For filaActual = 1 To Cells(Rows.Count, 1).End(xlUp).Row
    ' Define la hoja de trabajo "Proceso" como la hoja de trabajo activa
    Sheets("PROCESO").Activate
    ' Copia el valor de la celda actual en la columna A
    valorCelda = Cells(filaActual, 1).Copy
    Sheets("PL preliminar").Activate
    Range("MODELADO_PL_PRELIMINAR[MONTOS]").Cells(filaActual, 1).Select
    'ActiveSheet.ListObjects("MODELADO_PL_PRELIMINAR").ListColumns("MONTOS").DataBodyRange.Cells(1, 1).Select
    ActiveSheet.Paste
    
    Sheets("PROCESO").Activate
    ' Copia el valor de la celda actual en la columna A
    valorCelda = Cells(filaActual, 2).Copy
    Sheets("PL preliminar").Activate
    Range("MODELADO_PL_PRELIMINAR[SOLES]").Cells(filaActual, 1).Select
    'ActiveSheet.ListObjects("MODELADO_PL_PRELIMINAR").ListColumns("SOLES").DataBodyRange.Cells(1, 2).Select
    ActiveSheet.Paste
    
    Sheets("PROCESO").Activate
    ' Copia el valor de la celda actual en la columna A
    valorCelda = Cells(filaActual, 5).Copy
    Sheets("PL preliminar").Activate
    Range("MODELADO_PL_PRELIMINAR[DOLARES]").Cells(filaActual, 1).Select
    'ActiveSheet.ListObjects("MODELADO_PL_PRELIMINAR").ListColumns("SOLES").DataBodyRange.Cells(1, 3).Select
    ActiveSheet.Paste
    
    Range("MODELADO_PL_PRELIMINAR[CAMBIO]").Cells(filaActual, 1).Select
    ActiveCell.FormulaR1C1 = "=IFERROR([@SOLES]/[@DOLARES],)"
    
    'Agrega una celda después de la última fila de la tabla
    Set nuevaCelda = tabla.ListRows.Add.Range.Cells(1, 1)

Next filaActual

End Sub
Sub MODELADO_LIMPIEZA()

Sheets("BORRADOR").Select
'Esta línea borra el contenido de todas las celdas en la hoja activa. Esto significa que, si hay información en la hoja "BORRADOR", este código la borra.
Cells.Clear

End Sub

