Attribute VB_Name = "PRUEBAS"
Sub PEGAR_DATA_SUELDO2()

    Call ELIMINAR_REPORTE_SUELDO
    Call ELIMINAR_REPORTE_SUELDO_PROCESO
    
    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet
    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets("PROCESO") '<~~ Sheet where you want to import
    Set wbO = Workbooks.Open("C:\Macros LIMA\VALIDACION TXT PLAME\MC._IT0008.XLS")
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
    wbI.Sheets("PROCESO").Activate
    wbI.Sheets("PROCESO").Range("A1").Select
    
    Set tbl = Range("A1").CurrentRegion
    
    wsI.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "DATA_SUELDO"
    'wbI.Sheets("PRUEBA").Activate
'    wbI.ActiveSheet.ListObjects("DATA_SAP_REPORTE").DataBodyRange.Select
'    wbI.Sheets("SAP").Activate
'    wbI.Sheets("SAP").Range("A10").Select
'    wbI.Range.PasteSpecial xlPasteValue
    'Call ELIMINAR_TABLA_SAP_PRUEBA
    'wsI.Range("DATA_SAP_REPORTE[#All]").Cut
    'Destination:=Worksheet("SAP").Range("A10")
    wsI.ListObjects("DATA_SUELDO").Range.Cut _
        Destination:=Worksheets("REPORTE SUELDO").Range("A10")
    wbI.Sheets("REPORTE SUELDO").Activate


End Sub

