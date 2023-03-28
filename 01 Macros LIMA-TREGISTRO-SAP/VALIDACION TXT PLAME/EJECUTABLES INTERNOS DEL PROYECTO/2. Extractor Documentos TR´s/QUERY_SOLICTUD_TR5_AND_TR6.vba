Attribute VB_Name = "QUERY_SOLICTUD_TR5_AND_TR6"
Sub Modelo_TR5_ARCHIVO()

Application.DisplayAlerts = False

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False



Call ELIMINAR_TABLA_TR5


    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet

    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets("TR5") '<~~ Sheet where you want to import

    Set wbO = Workbooks.Open("C:\Macros LIMA\DOCUMENTOS TRS\TR5.txt")

    wbO.Sheets(1).Cells.Copy wsI.Cells

    wbO.Close SaveChanges:=False
    wbI.Sheets("TR5").Select
    wbI.Sheets("TR5").Range("A10").Select
    'wbI.Sheets("Modelo").Range("A10").Select
    wbI.Sheets("TR5").Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A10"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array( _
        25, 1)), TrailingMinusNumbers:=True
    
    wbI.Sheets("TR5").Range("A7").ClearContents
    wbI.Sheets("TR5").Range("A8").ClearContents
    wbI.Sheets("TR5").Range("A9").ClearContents
    'wbI.Sheets("TR5").Range("A10:U791").Select
    Call RemoveSpace_tr5
    
    nombre_libro = ActiveWorkbook.Name
    columnas = "A7:Z11"
    With Sheets("TR5").Range(columnas)
          Set c = .Find("===============================================================================================================================================================================================================================================================")
          If Not c Is Nothing Then
             c.Select
             letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
             Celda_final = letra_columna & c.Row
          End If
          
          Set inicio2 = .Find("Tipo")
          If Not inicio2 Is Nothing Then
          
             inicio2.Select
             letra_columna1 = Split(Cells(1, inicio2.Column).Address, "$")(1)
             Celda_de_inicio = letra_columna1 & inicio2.Row
          End If
          
          
    End With

    Range(Celda_final).Select
    Selection.EntireRow.Delete
    
    Set tbl = Range(Celda_de_inicio).CurrentRegion
    Set ws = ActiveSheet
    ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "HOJA_TR5"
    


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
    
End Sub


Sub RemoveSpace_tr5()
'Sheets("TR5").Select
Dim c As Range
'adjust range to suit
For Each c In Worksheets("TR5").Range("A10:AA10")
c = Trim(c)
Next c
End Sub
Sub ELIMINAR_TABLA_TR5()
    Dim NOMBRE As String
    On Error Resume Next
    Dim c As Range
     'Define Variables
    sSheetName = "TR5"
    sTableName = "HOJA_TR5"
    'Delete Table
     Sheets(sSheetName).ListObjects(sTableName).Delete
End Sub

Sub Modelo_TR6_ARCHIVO()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


    
    Call ELIMINAR_TABLA_TR6
    

    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet
    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets("TR6") '<~~ Sheet where you want to import
    Set wbO = Workbooks.Open("C:\Macros LIMA\DOCUMENTOS TRS\TR6.txt")
    
    Dim c As Range
    Dim inicio As Range
    Dim columnas As String
    Dim hoja As String
    Dim letra_columna As String
    Dim nombre_libro As String
    
    wbO.Sheets(1).Cells.Copy wsI.Cells
    wbO.Close SaveChanges:=False
    wbI.Sheets("TR6").Select
    wbI.Sheets("TR6").Range("A10").Select

    wbI.Sheets("TR6").Range("A10").Select
    wbI.Sheets("TR6").Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A10"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array( _
        25, 1)), TrailingMinusNumbers:=True
    'wbI.Sheets("Modelo2").Range("A11").ClearContents
    wbI.Sheets("TR6").Range("A7").ClearContents
    wbI.Sheets("TR6").Range("A8").ClearContents
    wbI.Sheets("TR6").Range("A9").ClearContents
    wbI.Sheets("TR6").Range("A10:U791").Select
    
    Call RemoveSpace_tr6
    
    nombre_libro = ActiveWorkbook.Name
    columnas = "A10:A100"
    With Sheets("TR6").Range(columnas)
          Set c = .Find("=")
          If Not c Is Nothing Then
          
             c.Select
             letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
             Celda_final = letra_columna & c.Row
             
             
          End If
          
          Set inicio2 = .Find("Tipo")
          If Not inicio2 Is Nothing Then
          
             inicio2.Select
             letra_columna1 = Split(Cells(1, inicio2.Column).Address, "$")(1)
             Celda_de_inicio = letra_columna1 & inicio2.Row
          End If
          
    End With
    
    Range(Celda_final).Select
    Selection.EntireRow.Delete
    
    Set tbl = Range(Celda_de_inicio).CurrentRegion
    Set ws = ActiveSheet
    ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "MODELO_TR6"


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub
Sub RemoveSpace_tr6()
'Sheets("TR5").Select
Dim c As Range
'adjust range to suit
For Each c In Worksheets("TR6").Range("A10:AB10")
c = Trim(c)
Next c
End Sub

Sub ELIMINAR_TABLA_TR6()
    Dim NOMBRE As String
    On Error Resume Next
    Dim c As Range
    columnas = "A:ZZ"
     'Define Variables
    sSheetName = "TR6"
    sTableName = "MODELO_TR6"
    'Delete Table
     Sheets(sSheetName).ListObjects(sTableName).Delete
End Sub
