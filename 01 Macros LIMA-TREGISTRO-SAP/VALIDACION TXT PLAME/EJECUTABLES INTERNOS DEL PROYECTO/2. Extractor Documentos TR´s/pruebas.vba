Attribute VB_Name = "pruebas"
Sub aumentar_TR5()
Range("D1").Value = Range("D1").Value + 1
Range("C1").Value = Range("C1").Value + 1
End Sub



Sub Modelo_TR5_ARCHIVO1()
Call ELIMINAR_TABLA_TR5

Sheets("TR5").Select
Dim fileToOpen As Variant
Dim fileFilter As String
Dim title As String
On Error GoTo ErrorTablaExistente

Dim wsMaster As Worksheet
Dim wbTextImport As Workbook

Dim nombre_hoja As String
nombre_hoja = Trim(Range("B1").Value)
Dim numero_hoja As String
numero_hoja = Trim(Range("C1").Value)
Dim numero_iteracion As String
numero_iteracion = Trim(Range("D1").Value)



'Application.ScreenUpdating = False
'fileFilterPattern = "Text Files (*.txt; *.csv), *.txt; *.csv"
'title = "Select Multiple Files"
'fileToOpen = Application.GetOpenFilename(fileFilter:="Txt Files (*.txt), *.xls", title:="Por favor, seleccione el txt del TR5")

    'If fileToOpen = False Then
     '   MsgBox "No se ha seleccionado ningún archivo"
    'Else
    
    Call aumentar_TR5
    
    ActiveWorkbook.Queries.Add Name:="HOJA TR5" + nombre_hoja + "-" + CStr(numero_hoja) + CStr(numero_iteracion), Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Origen = Csv.Document(File.Contents(""C:\Macros LIMA\DOCUMENTOS TRS\TR5.txt""),[Delimiter=""|"", Columns=25, Encoding=1252, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""Tipo cambiado"" = Table.TransformColumnTypes(Origen,{{""Column1"", type text}, {""Column2"", type text}, {""Column3"", type text}, {""Column4"", type text}, {""Column5"", type text}, {""Co" & _
        "lumn6"", type text}, {""Column7"", type text}, {""Column8"", type text}, {""Column9"", type text}, {""Column10"", type text}, {""Column11"", type text}, {""Column12"", type text}, {""Column13"", type text}, {""Column14"", type text}, {""Column15"", type text}, {""Column16"", type text}, {""Column17"", type text}, {""Column18"", type text}, {""Column19"", type text}," & _
        " {""Column20"", type text}, {""Column21"", type text}, {""Column22"", type text}, {""Column23"", type text}, {""Column24"", type text}, {""Column25"", type text}})," & Chr(13) & "" & Chr(10) & "    #""Filas superiores quitadas"" = Table.Skip(#""Tipo cambiado"",9)," & Chr(13) & "" & Chr(10) & "    #""Encabezados promovidos"" = Table.PromoteHeaders(#""Filas superiores quitadas"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Tipo cam" & _
        "biado1"" = Table.TransformColumnTypes(#""Encabezados promovidos"",{{""       Tipo       "", type text}, {""    Número    "", Int64.Type}, {""  Apellido paterno  "", type text}, {""  Apellido materno  "", type text}, {""         Nombres         "", type text}, {""   Fec. Inicio   "", type date}, {""                     Tipo de Trabajador                     "", type " & _
        "text}, {""       Régimen Laboral        "", type text}, {""   Cat. Ocupacional   "", type text}, {""                    Ocupación                     "", type text}, {""           Nivel Educativo            "", type text}, {""Discapacidad"", type text}, {""Sindicalizado"", type text}, {""Reg. Acumulativo"", type text}, {""Máxima"", type text}, {""Horario Nocturno""," & _
        " type text}, {""          Situación Especial del Trabajador                 "", type text}, {""                                                           Establecimiento en el que labora                                                           "", type text}, {""        Tipo de Contrato         "", type text}, {""     Tipo de pago "", type text}, {"" Periodicidad " & _
        """, type text}, {""Entidad Financiera              "", type text}, {""Nro de Cuenta       "", Int64.Type}, {"" Remun Bas."", type number}, {""Situación"", type text}})," & Chr(13) & "" & Chr(10) & "    #""Filas superiores quitadas1"" = Table.Skip(#""Tipo cambiado1"",1)" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Filas superiores quitadas1"""
    
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" + "HOJA TR5" + CStr(nombre_hoja) + "-" + CStr(numero_hoja) + CStr(numero_iteracion) + ";Extended Properties=""""", Destination:=Range("$A$8")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" + "HOJA TR5" + CStr(nombre_hoja) + "-" + CStr(numero_hoja) + CStr(numero_iteracion) + "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.Name = "HOJA_TR5"
        .Refresh BackgroundQuery:=False
    End With
    
    Call RemoveSpace_tr5
 
Exit Sub
ErrorTablaExistente:
 MsgBox "Debe Eliminar tabla Existente"
End Sub

'Sub RemoveSpace()
'Dim C As Range
'adjust range to suit
'For Each C In Range("A8:ZZ8")
'C = Trim(C)
'Next C
'End Sub
'
'Sub RemoveSpace_tr()
'Sheets("TR5").Select
'Dim C As Range
'adjust range to suit
'For Each C In Worksheets("TR5").Range("A8:ZZ8")
'C = Trim(C)
'Next C
'End Sub
'Sub Remove6()
'Sheets("TR6").Select
'Dim C As Range
'adjust range to suit
'For Each C In Range("A9:ZZ9")
'For Each C In Worksheets("TR6").Range("A8:ZZ8")
'C = Trim(C)
'Next C
'End Sub
'
'Sub aumentar_TR6()
'Range("C1").Value = Range("C1").Value + 1
'Range("D1").Value = Range("D1").Value + 1
'End Sub


Sub Modelo_TR6_ARCHIVO1()

    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet
    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets("Modelo2") '<~~ Sheet where you want to import
    Set wbO = Workbooks.Open("C:\Macros LIMA\DOCUMENTOS TRS\TR6.txt")
    
    Dim c As Range
    Dim columnas As String
    Dim hoja As String
    Dim letra_columna As String
    Dim nombre_libro As String
    
    wbO.Sheets(1).Cells.Copy wsI.Cells
    wbO.Close SaveChanges:=False
    wbI.Sheets("Modelo2").Range("A10").Select

    wbI.Sheets("Modelo2").Range("A10").Select
    wbI.Sheets("Modelo2").Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A10"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array( _
        25, 1)), TrailingMinusNumbers:=True
    'wbI.Sheets("Modelo2").Range("A11").ClearContents
    wbI.Sheets("Modelo2").Range("A7").ClearContents
    wbI.Sheets("Modelo2").Range("A8").ClearContents
    wbI.Sheets("Modelo2").Range("A9").ClearContents
    wbI.Sheets("Modelo2").Range("A10:U791").Select
    Call RemoveSpace_tr6
    nombre_libro = ActiveWorkbook.Name
    columnas = "A11:Z11"
    With Sheets("Modelo2").Range(columnas)
          Set c = .Find("=")
          If Not c Is Nothing Then
          
             c.Select
             letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
             Celda_final = letra_columna & c.Row
             
             
          End If
    End With

    Range(Celda_final).Select
    Selection.EntireRow.Delete
    
    Set tbl = Range("A10").CurrentRegion
    Set ws = ActiveSheet
    ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "MODELO_TR6"
    
End Sub

Sub Model()



Call ELIMINAR_TABLA_TR6

Sheets("TR6").Select

Dim fileToOpen As Variant
Dim fileFilter As String
Dim title As String
On Error GoTo ErrorTablaExistente

Dim wsMaster As Worksheet
Dim wbTextImport As Workbook

Dim nombre_hoja As String
nombre_hoja = Trim(Range("B1").Value)
Dim numero_hoja As String
numero_hoja = Trim(Range("C1").Value)
Dim numero_iteracion As String
numero_iteracion = Trim(Range("D1").Value)



'Application.ScreenUpdating = False
'fileFilterPattern = "Text Files (*.txt; *.csv,), *.txt; *.csv"
'title = "Select Multiple Files"
'fileToOpen = Application.GetOpenFilename(fileFilter:="Txt Files (*.txt), *.xls", title:="Por favor, seleccione el txt del TR6")


    'If fileToOpen = False Then
      '  MsgBox "No se ha seleccionado ningún archivo"
    'Else

    Call aumentar_TR6
    
    ActiveWorkbook.Queries.Add Name:="MODELO TR6" + nombre_hoja + "-" + CStr(numero_hoja) + CStr(numero_iteracion), Formula _
        := _
        "let" & Chr(13) & "" & Chr(10) & "    Origen = Csv.Document(File.Contents(""C:\Macros LIMA\DOCUMENTOS TRS\TR6.txt""),[Delimiter="","", Columns=2, Encoding=1252, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    #""Tipo cambiado"" = Table.TransformColumnTypes(Origen,{{""Column1"", type text}, {""Column2"", type text}})," & Chr(13) & "" & Chr(10) & "    #""Filas superiores" & _
        " quitadas"" = Table.Skip(#""Tipo cambiado"",9)," & Chr(13) & "" & Chr(10) & "    #""Dividir columna por delimitador"" = Table.SplitColumn(#""Filas superiores quitadas"", ""Column1"", Splitter.SplitTextByDelimiter(""|"", QuoteStyle.Csv), {""Column1.1"", ""Column1.2"", ""Column1.3"", ""Column1.4"", ""Column1.5"", ""Column1.6"", ""Column1.7"", ""Column1.8"", ""Column1.9"", ""Column1.10"", ""Colum" & _
        "n1.11"", ""Column1.12"", ""Column1.13"", ""Column1.14"", ""Column1.15"", ""Column1.16"", ""Column1.17"", ""Column1.18"", ""Column1.19"", ""Column1.20"", ""Column1.21""})," & Chr(13) & "" & Chr(10) & "    #""Tipo cambiado1"" = Table.TransformColumnTypes(#""Dividir columna por delimitador"",{{""Column1.1"", type text}, {""Column1.2"", type text}, {""Column1.3"", type text}, {""Column1.4"", type " & _
        "text}, {""Column1.5"", type text}, {""Column1.6"", type text}, {""Column1.7"", type text}, {""Column1.8"", type text}, {""Column1.9"", type text}, {""Column1.10"", type text}, {""Column1.11"", type text}, {""Column1.12"", type text}, {""Column1.13"", type text}, {""Column1.14"", type text}, {""Column1.15"", type text}, {""Column1.16"", type text}, {""Column1.17"", t" & _
        "ype text}, {""Column1.18"", type text}, {""Column1.19"", type text}, {""Column1.20"", type text}, {""Column1.21"", type text}})," & Chr(13) & "" & Chr(10) & "    #""Encabezados promovidos"" = Table.PromoteHeaders(#""Tipo cambiado1"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Tipo cambiado2"" = Table.TransformColumnTypes(#""Encabezados promovidos"",{{""       Tipo       "", type text}, {""    Número  " & _
        "  "", Int64.Type}, {""  Apellido paterno  "", type text}, {""  Apellido materno  "", type text}, {""         Nombres          "", type text}, {""Situación"", type text}, {""                     Tipo de Trabajador                     "", type text}, {""          Tipo de Régimen          "", type text}, {"" Fecha Inicio "", type date}, {""         EPS/Serv. Propio    " & _
        "      "", type text}, {""         Tipo de Régimen         "", type text}, {"" Fecha Inicio _1"", type date}, {""     CUSPP     "", type text}, {""Ninguno"", type text}, {""ESSALUD"", type text}, {""  EPS  "", type text}, {""Ninguno_2"", type text}, {"" ONP "", type text}, {""Seg. Privado"", type text}, {""Ind. rentas LIR"", type text}, {""Conv. evit. doble imp."", t" & _
        "ype text}, {"""", type text}})," & Chr(13) & "" & Chr(10) & "    #""Filas superiores quitadas1"" = Table.Skip(#""Tipo cambiado2"",1)" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Filas superiores quitadas1"""
    
    
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" + "MODELO TR6" + CStr(nombre_hoja) + "-" + CStr(numero_hoja) + CStr(numero_iteracion) + ";Extended Properties=""""" _
        , Destination:=Range("$A$8")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" + "MODELO TR6" + CStr(nombre_hoja) + "-" + CStr(numero_hoja) + CStr(numero_iteracion) + "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "MODELO_TR6"
        .Refresh BackgroundQuery:=False
    
    End With
    Call RemoveSpace_tr6

Exit Sub
ErrorTablaExistente:
 MsgBox "Debe Eliminar Tabla Existente"
 
End Sub


'
'    Sub RemoveSpace_other11()
'    Dim C As Range
'    adjust range to suit
'    For Each C In Range("A8:ZZ8")
'    C = Trim(C)
'    Next C
'    End Sub
'
'    Sub ELIMINAR_TABLA_TR611()
'    Dim NOMBRE As String
'    On Error Resume Next
'    Dim C As Range
'    columnas = "A:ZZ"
'     Define Variables
'    sSheetName = "TR6"
'    sTableName = "MODELO_TR6"
'    Delete Table
'     Sheets(sSheetName).ListObjects(sTableName).Delete
'    End Sub
'    Sub ELIMINAR_TABLA_TR511()
'    Dim NOMBRE As String
'    On Error Resume Next
'    Dim C As Range
'    columnas = "A:ZZ"
'     Define Variables
'    sSheetName = "TR5"
'    sTableName = "HOJA_TR5"
'    Delete Table
'     Sheets(sSheetName).ListObjects(sTableName).Delete
'    End Sub


Sub OpenAndImportTxtFile()
    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet

    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets("Modelo") '<~~ Sheet where you want to import

    Set wbO = Workbooks.Open("C:\Macros LIMA\DOCUMENTOS TRS\TR5.txt")

    wbO.Sheets(1).Cells.Copy wsI.Cells

    wbO.Close SaveChanges:=False
    
    wbI.Sheets("Modelo").Range("A10").Select
    'wbI.Sheets("Modelo").Range("A10").Select
    wbI.Sheets("Modelo").Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A10"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array( _
        25, 1)), TrailingMinusNumbers:=True
    
    wbI.Sheets("Modelo").Range("A7").ClearContents
    wbI.Sheets("Modelo").Range("A8").ClearContents
    wbI.Sheets("Modelo").Range("A9").ClearContents
    wbI.Sheets("Modelo").Range("A10:U791").Select
    Call RemoveSpace_tr6
    nombre_libro = ActiveWorkbook.Name
    columnas = "A7:Z11"
    With Sheets("Modelo").Range(columnas)
          Set c = .Find("=")
          If Not c Is Nothing Then
          
             c.Select
             letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
             Celda_final = letra_columna & c.Row
             
             
          End If
    End With

    Range(Celda_final).Select
    Selection.EntireRow.Delete
    
    Set tbl = Range("A10").CurrentRegion
    Set ws = ActiveSheet
    ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "HOJA_TR5"
    
    
    
End Sub

Sub OpenAndImportTxtFile_2()
    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet
    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets("Modelo2") '<~~ Sheet where you want to import
    Set wbO = Workbooks.Open("C:\Macros LIMA\DOCUMENTOS TRS\TR6.txt")
    
    Dim c As Range
    Dim columnas As String
    Dim hoja As String
    Dim letra_columna As String
    Dim nombre_libro As String
    
    wbO.Sheets(1).Cells.Copy wsI.Cells
    wbO.Close SaveChanges:=False
    wbI.Sheets("Modelo2").Range("A10").Select
    'wbI.Sheets("Modelo2").Range("A10").Select
    wbI.Sheets("Modelo2").Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A10"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1), Array(24, 1), Array( _
        25, 1)), TrailingMinusNumbers:=True
    'wbI.Sheets("Modelo2").Range("A11").ClearContents
    wbI.Sheets("Modelo2").Range("A7").ClearContents
    wbI.Sheets("Modelo2").Range("A8").ClearContents
    wbI.Sheets("Modelo2").Range("A9").ClearContents
    wbI.Sheets("Modelo2").Range("A10:U791").Select
    Call RemoveSpace_tr6
    nombre_libro = ActiveWorkbook.Name
    columnas = "A11:Z11"
    With Sheets("Modelo2").Range(columnas)
          Set c = .Find("=")
          If Not c Is Nothing Then
          
             c.Select
             letra_columna = Split(Cells(1, c.Column).Address, "$")(1)
             Celda_final = letra_columna & c.Row
             
             
          End If
    End With

    Range(Celda_final).Select
    Selection.EntireRow.Delete
    
    Set tbl = Range("A10").CurrentRegion
    Set ws = ActiveSheet
    ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "MODELO_TR6"
    
End Sub

Sub RemoveSpace_tr5()
'Sheets("TR6").Select
Dim c As Range
'adjust range to suit
'For Each c In Range("A9:ZZ9")
For Each c In Worksheets("Modelo").Range("A10:Z10")
c = Trim(c)
Next c
End Sub


Sub RemoveSpace_tr6()
'Sheets("TR6").Select
Dim c As Range
'adjust range to suit
'For Each c In Range("A9:ZZ9")
For Each c In Worksheets("Modelo2").Range("A10:Z10")
c = Trim(c)
Next c
End Sub




