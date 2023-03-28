Attribute VB_Name = "PRUEBAS"
Sub modulo_prueba()
'Workbooks.Open "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\REPORTES\EXPORTABLE.xlsx"

    Call ELIMINAR_TABLA_SAP
    Call ELIMINAR_TABLA_SAP_PRUEBA

    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet
    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets("PRUEBA") '<~~ Sheet where you want to import
    Set wbO = Workbooks.Open("C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\REPORTES\EXPORTABLE.XLSX")
    Dim tbl As Range
    Dim ws As Worksheet
    
    Dim C As Range
    Dim columnas As String
    Dim hoja As String
    Dim letra_columna As String
    Dim nombre_libro As String
    
    wbO.Sheets(1).Cells.Copy wsI.Cells
    wbO.Close SaveChanges:=False
    'wbI.Sheets("PRUEBA").Range("A10").Select
    wbI.Sheets("PRUEBA").Activate
    wbI.Sheets("PRUEBA").Range("A1").Select
    
    Set tbl = Range("A1").CurrentRegion
    
    wsI.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "DATA_SAP_REPORTE"
    'wbI.Sheets("PRUEBA").Activate
'    wbI.ActiveSheet.ListObjects("DATA_SAP_REPORTE").DataBodyRange.Select
'    wbI.Sheets("SAP").Activate
'    wbI.Sheets("SAP").Range("A10").Select
'    wbI.Range.PasteSpecial xlPasteValue
    'Call ELIMINAR_TABLA_SAP_PRUEBA
    'wsI.Range("DATA_SAP_REPORTE[#All]").Cut
    'Destination:=Worksheet("SAP").Range("A10")
    wsI.ListObjects("DATA_SAP_REPORTE").Range.Cut _
        Destination:=Worksheets("SAP").Range("A10")
    wbI.Sheets("SAP").Activate
    
    wsI.ListObjects ("DATA_SAP_REPORTE[]")
    
End Sub


