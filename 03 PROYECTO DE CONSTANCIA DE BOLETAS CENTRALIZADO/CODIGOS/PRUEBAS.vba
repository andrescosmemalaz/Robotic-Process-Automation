Attribute VB_Name = "PRUEBAS"
Function FiltrarTablaPorColumnaYValor_v2(tbl As ListObject, columna As String, valor As Variant)
   Dim col As Range
   Set col = tbl.ListColumns(columna).Range
   tbl.Range.AutoFilter Field:=col.Column, Criteria1:=valor
   tbl.Range.AutoFilter
End Function


Sub OBTENER_VALORES_UNICOS()

'Dim k As String
'Dim celda As Range
Dim tbl As ListObject
Set tbl = ActiveSheet.ListObjects("VALIDACION_CONSTANCIA")
Dim columna_valor As Integer
origen = Range("VALIDACION_CONSTANCIA[Doc.compensación]")
Set Dic_objeto = CreateObject("Scripting.Dictionary")
For fila = 1 To UBound(origen)
    Dic_objeto(origen(fila, 1)) = 0
Next
Sheets("PROCESO").Activate
Range("E2").Resize(Dic_objeto.Count) = WorksheetFunction.Transpose(Dic_objeto.Keys)
'---FIN ÚNICOS---

fila_inicial = Range("E2").Row
ultimo = Range("E2").End(xlDown).Row

For i = fila_inicial To ultimo
    Sheets("PROCESO").Activate
    Range("E" + CStr(i)).Select
    Filtro = Range("E" + CStr(i)).Value
    
    MsgBox (Filtro)
    Sheets("VALIDACION").Activate
    'ActiveSheet.ListObjects("VALIDACION_CONSTANCIA").Range.AutoFilter Field:=13, Criteria1:=Filtro
    
    columna_valor = 13
    valor = Filtro
    FiltrarTablaPorColumnaYValor tbl, columna_valor, valor
    fila_inicio = Range("N2").Row
    fila_final = Range("N2").End(xlDown).Row
    Call FILTRO_AVANZADO
Next
MsgBox ("Proceso Finalizado")


End Sub

Sub FILTRO_AVANZADO()
Dim tbl As ListObject
Set tbl = ActiveSheet.ListObjects("VALIDACION_CONSTANCIA")

Dim celda_recorrida, rango_origen, rango_final As Range
Dim variable As Range
Dim valor_negativo As Double
Dim valor_positivo As Double
Dim valor_final As Integer
Dim ws As Worksheet
Set ws = Sheets("VALIDACION")

ws.Select

fila_inicial = Range("K3").Row
fila_final = Range("K3").End(xlDown).Row

'REALIZAR EL FILTRO
 
'ALMACENAS TODO EL RANGO
Set rango_origen = Range("K" + CStr(fila_inicial) + ":K" + CStr(fila_final)).SpecialCells(xlCellTypeVisible)
Set rango_validacion = Range("U" + CStr(fila_inicial) + ":U" + CStr(fila_final)).SpecialCells(xlCellTypeVisible)
  
For Each celda_recorrida In rango_origen
    MsgBox (celda_recorrida.Row)
    If celda_recorrida.Value > 0 Then
        valor_positivo = celda_recorrida.Value
        MsgBox (celda_recorrida.Value)
    Else
        valor_negativo = celda_recorrida.Value
        MsgBox (celda_recorrida.Value)
    'Range("V" + CStr(celda.Row)).Value = "SOLTERO"
    End If
Next

For Each celda_recorrida In rango_validacion
    'MsgBox (celda.Row)
    valor_final = Abs(valor_positivo) - Abs(valor_negativo)
    Range("U" + CStr(celda_recorrida.Row)).Value = valor_final
Next


End Sub



Sub CopiarColumna()
    'Declara las variables
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim rngOrigen As Range
    Dim rngDestino As Range

    'Asigna las variables a los objetos correspondientes
    Set wsOrigen = Worksheets("REPORTE_SAP") 'Cambia "Hoja1" por el nombre de la hoja de origen
    Set wsDestino = Worksheets("VALIDACION") 'Cambia "Hoja2" por el nombre de la hoja de destino
    
    
    Set rngOrigen = wsOrigen.ListObjects("DATA_SAP_FBLN").ListColumns("Cuenta").DataBodyRange 'Cambia "Tabla1" y "Columna1" por el nombre de tu tabla y columna de origen, respectivamente
    Set rngDestino = wsDestino.ListObjects("VALIDACION_CONSTANCIA").ListColumns("Cuenta").DataBodyRange 'Cambia "Tabla2" y "Columna2" por el nombre de tu tabla y columna de destino, respectivamente
    
    
    
    'Copia la columna desde la tabla de origen a la tabla de destino
    rngOrigen.Copy
    rngDestino.PasteSpecial xlPasteValues
End Sub



Sub COLUMNA_DOBLE()

Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_Cuenta As String
Dim Valor_ejercicio_mes As String

On Error GoTo ErrorHandler
With Sheets("REPORTE_SAP")
    columna = Application.Match("Cuenta", .Range("A10:O10"), 0)
    letra_columna = Split(.Cells(1, columna).Address, "$")(1)
    NOMBRE = letra_columna + CStr(10)
    Valor_Cuenta = .Range(NOMBRE).Value
    .Range("DATA_SAP_FBLN[Valor_Cuenta]").Copy
    
    columnas = "A10:O10"
    columna = Application.Match("Ejercicio / mes", .Range(columnas).Cells, 1)
    letra_columna = Split(.Cells(1, columna).Address, "$")(1)
    NOMBRE = CStr(letra_columna) + CStr(10)
    Valor_ejercicio_mes = .Range(NOMBRE).Value
    Range("DATA_SAP_FBLN[" & "Ejercicio / mes" & "]").Copy
End With

Sheets("VALIDACION").Range("VALIDACION_CONSTANCIA[Cuenta]").PasteSpecial xlPasteValues
Sheets("VALIDACION").Range("VALIDACION_CONSTANCIA[Ejercicio / mes]").PasteSpecial xlPasteValues

Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True

Exit Sub
ErrorHandler:
Sheets("VALIDACION").Range("VALIDACION_CONSTANCIA[Cta.contrapartida]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"

Sheets("VALIDACION").Range("VALIDACION_CONSTANCIA[Ejercicio / mes]").Value = "NO SE HA PODIDO ENCONTRAR EL VALOR"


End Sub



Sub DESAROLLO_PDF_PRUEBAS()

Application.DisplayAlerts = False

    Dim Nombre_archivo As String
    Dim nombre_pdf As String
    Dim myPath_new As String
    Dim fso As New FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ruta As String
    ruta = "C:\Macros\PROTOTIPO CONSTANCIAS\PRUEBA\"
    Dim libro_sistema As String
    libro_sistema = ThisWorkbook.Name

    myPath = Dir(ruta + "*.pdf")
            Do While myPath <> ""
                nombre_pdf = myPath
                myPath = ruta + myPath
                Set openPDF = CreateObject("Shell.Application")
                openPDF.Open (myPath)
                SendKeys "^a"
                SendKeys "^c"
                Set ws = Sheets("PROCESO")
                ws.Range("A10").Select
                ActiveSheet.Paste
                'ENCONTRANDO INFORMACIÓN EN EL PDF
                With Worksheets("PROCESO").Cells
                'Referencia de Planilla
                    Dim oFound_referencia As Range
                    Set oFound_referencia = .Find(What:="Referencia de planilla:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    If Not oFound_referencia Is Nothing Then
                        Dim fila_match_planilla As Long
                        fila_match_planilla = oFound_referencia.Row
                        Dim columna_match_planilla As Long
                        columna_match_planilla = oFound_referencia.Column
                        Dim referencia_planilla As String
                        referencia_planilla = Cells(fila_match_planilla + 1, columna_match_planilla).Value
                    Else
                        referencia_planilla = "Nombre no encontrado"
                    End If
        
                    'Fecha de Proceso
                    Dim oFound_fecha As Range
                    Set oFound_fecha = .Find(What:="Fecha de proceso:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    If Not oFound_fecha Is Nothing Then
                        Dim fila_match_fecha As Long
                        fila_match_fecha = oFound_fecha.Row
                        Dim columna_match_fecha As Long
                        columna_match_fecha = oFound_fecha.Column
                        Dim fecha_proceso As String
                        fecha_proceso = Cells(fila_match_fecha + 1, columna_match_fecha).Value
                    Else
                        fecha_proceso = "Nombre no encontrado"
                    End If
            
                    'Cuenta de origen
                    Dim oFound_cuenta_origen As Range
                    Set oFound_cuenta_origen = .Find(What:="Cuenta deorigen:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    If Not oFound_cuenta_origen Is Nothing Then
                        Dim fila_match_cuenta_origen As Long
                        fila_match_cuenta_origen = oFound_cuenta_origen.Row
                        
                    Dim columna_match_cuenta_origen As Long
                    columna_match_cuenta_origen = oFound_cuenta_origen.Column
                    Dim cuenta_origen As String
                    cuenta_origen = Cells(fila_match_cuenta_origen + 1, columna_match_cuenta_origen).Value
                    Else
                    cuenta_origen = "Nombre no encontrado"
                    End If
                    
                End With
                
                    'Aqui podria ir el codigo para guardar la información encontrada en una tabla o en una hoja específica
                
                    'Cerrar y mover el pdf
                    openPDF.Quit
                    myPath_new = historial_pdf + nombre_pdf
                    fso.MoveFile myPath, myPath_new
                    myPath = Dir()
            Loop
End Sub


Sub EXPORTAR_CELDA_ORIGEN()

'Declarar variables
Dim origen As Workbook
Dim destino As Workbook
Dim origen_hoja As Worksheet
Dim destino_hoja As Worksheet
Dim origen_rango As Range
Dim destino_rango As Range
Dim transformada As String

'Definir el libro de origen y la hoja de origen
Set origen = Workbooks.Open("C:\Macros\PROTOTIPO CONSTANCIAS\ATACOCHA.xlsx")
Set origen_hoja = origen.Sheets("SISTEMA")

'Definir el rango de celdas de origen
Set origen_rango_fecha_inicio = origen_hoja.Range("J8")
Set origen_rango_fecha_fin = origen_hoja.Range("J9")

'Transformar el contenido de la celda de origen
'transformada = "=CONCAT(SI(DIA(" & origen_rango.Address & ")>9,DIA(" & origen_rango.Address & "),""0""&DIA(" & origen_rango.Address & ")),""."",SI(MES(" & origen_rango.Address & ")>9,MES(" & origen_rango.Address & "),""0""&MES(" & origen_rango.Address & ")),""."",AÑO(" & origen_rango.Address & "))"

'Definir el libro de destino y la hoja de destino
Set destino = Workbooks.Open("C:\Macros\PROTOTIPO CONSTANCIAS\\PROYECTO VALIDACION CONSTANCIA.xlsm")
Set destino_hoja = destino.Sheets("REPORTE_SAP")

'Definir el rango de celdas de destino
Set destino_rango_fecha_inicio = destino_hoja.Range("B2")
Set destino_rango_fecha_fin = destino_hoja.Range("D2")


'Exportar el contenido transformado a la celda de destino
destino_rango_fecha_inicio.Value = origen_rango_fecha_inicio.Value
destino_rango_fecha_fin.Value = origen_rango_fecha_fin.Value

'Guardar y cerrar el libro de destino
origen.Save
origen.Close
destino.Save
'destino.Close


End Sub



Sub MOVIMIENTO_COLUMNAS_FBLN()

Application.DisplayAlerts = False

Dim logFile As String

Call LIMPIAR_COLUMNAS
Call COLUMNA_CUENTA
Call COLUMNA_EJERCICIO_MES
Call COLUMNA_SOCIEDAD
Call COLUMNA_DIVISION
Call COLUMNA_TEXTO
Call COLUMNA_REFERENIA
Call COLUMNA_CLASE_DE_DOCUMENTO
Call COLUMNA_N_DOCUMENTO
Call COLUMNA_MONEDA_DEL_DOCUMENTO
Call COLUMNA_IMPORTE_EN_MONEDA_LOCAL
Call COLUMNA_IMPORTE_EN_MONEDA_DOC
Call COLUMNA_FE_CONTABILIZACION
Call COLUMNA_DOC_COMPENSACION
Call COLUMNA_FECHA_COMPENSACION
Call COLUMNA_NOMBRE_USUARIO
Call COLUMNA_CUENTA_DE_MAYOR
Call COLUMNA_CODIGO_TRANSACCION
Call COLUMNA_CTA_CONTRAPARTIDA
Call COLUMNA_CLAVE_REFERENCIA

End Sub


Sub COLUMNA_CUENTA()
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_Cuenta As String
Dim ws As Worksheet
Set ws = Sheets("VALIDACION")

With Sheets("REPORTE_SAP")
    On Error GoTo ErrorHandler
    columna = Application.Match("Cuenta", .Range("A10:O10"), 0)
    letra_columna = Split(.Cells(1, columna).Address, "$")(1)
    NOMBRE = letra_columna & 10
    Valor_Cuenta = .Range(NOMBRE).Value
    .Range("DATA_SAP_FBLN[" & Valor_Cuenta & "]").Copy
End With

ws.Range("VALIDACION_CONSTANCIA[Cuenta]").PasteSpecial xlPasteValues

Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
Exit Sub
ErrorHandler:
Sheets("VALIDACION").Range("VALIDACION_CONSTANCIA[Cta.contrapartida]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub


Sub COLUMNA_EJERCICIO_MES()

Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_ejercicio_mes As String
Dim ws As Worksheet
Set ws = Sheets("VALIDACION")

columnas = "A10:C10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
        columna = Application.Match("Ejercicio / mes", .Cells, 1)
        letra_columna = Split(.Cells(1, columna).Address, "$")(1)
        NOMBRE = CStr(letra_columna) + CStr(10)
        Valor_ejercicio_mes = .Range(NOMBRE).Value
        Range("DATA_SAP_FBLN[" & "Ejercicio / mes" & "]").Copy
    End With
    
ws.Range("VALIDACION_CONSTANCIA[Ejercicio / mes]").PasteSpecial xlPasteValues

Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True

Exit Sub

ErrorHandler:
ws.Range("VALIDACION_CONSTANCIA[Ejercicio / mes]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub

Sub COLUMNA_SOCIEDAD()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_sociedad As String
'Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Activate

columnas = "A10:O10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Sociedad")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_sociedad = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_sociedad) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            'Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Sociedad]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Sociedad]").Activate
    Range("VALIDACION_CONSTANCIA[Sociedad]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub


Sub COLUMNA_DIVISION()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_division As String
'Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Activate

columnas = "A10:O10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("División")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_division = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_division) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            'Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[División]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[División]").Activate
    Range("VALIDACION_CONSTANCIA[División]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub

Sub COLUMNA_TEXTO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_texto As String
'Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Activate

columnas = "A10:O10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Texto")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_texto = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_texto) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            'Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Texto]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Texto]").Activate
    Range("VALIDACION_CONSTANCIA[Texto]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub

Sub COLUMNA_REFERENIA()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_referencia As String
'Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Activate

columnas = "A10:O10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Referencia")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_referencia = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_referencia) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Referencia]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Referencia]").Activate
    Range("VALIDACION_CONSTANCIA[Referencia]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub

Sub COLUMNA_CLASE_DE_DOCUMENTO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_clase_documento As String
'Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Activate

columnas = "A10:O10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Clase de documento")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_clase_documento = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_clase_documento) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Clase de documento]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Clase de documento]").Activate
    Range("VALIDACION_CONSTANCIA[Clase de documento]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub



Sub COLUMNA_N_DOCUMENTO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_n_documento As String
'Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Activate

columnas = "A10:O10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Nº documento")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_n_documento = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_n_documento) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Nº documento]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Nº documento]").Activate
    Range("VALIDACION_CONSTANCIA[Nº documento]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub


Sub COLUMNA_MONEDA_DEL_DOCUMENTO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_moneda_del_documento As String
'Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Activate

columnas = "A10:O10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Moneda del documento")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_moneda_del_documento = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_moneda_del_documento) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Moneda del documento]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Moneda del documento]").Activate
    Range("VALIDACION_CONSTANCIA[Moneda del documento]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub


Sub COLUMNA_IMPORTE_EN_MONEDA_LOCAL()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_importe_moneda_local As String
'Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Activate

columnas = "J10:K10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Importe en moneda local")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_importe_moneda_local = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_importe_moneda_local) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Importe en moneda local]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Importe en moneda local]").Activate
    Range("VALIDACION_CONSTANCIA[Importe en moneda local]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub


Sub COLUMNA_IMPORTE_EN_MONEDA_DOC()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_importe_moneda_doc As String
'Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Activate

columnas = "A10:O10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Importe en moneda doc.")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_importe_moneda_local = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_importe_moneda_local) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Importe en moneda doc.]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Importe en moneda doc.]").Activate
    Range("VALIDACION_CONSTANCIA[Importe en moneda doc.]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub


Sub COLUMNA_FE_CONTABILIZACION()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_fe_contabilizacion As String
'Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Activate

columnas = "A10:O10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Fe.contabilización")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_fe_contabilizacion = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_fe_contabilizacion) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Fe.contabilización]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Fe.contabilización]").Activate
    Range("VALIDACION_CONSTANCIA[Fe.contabilización]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub





Sub COLUMNA_DOC_COMPENSACION()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_doc_compensacion As String
'Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Activate

columnas = "A10:O10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Doc.compensación")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_doc_compensacion = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_doc_compensacion) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Doc.compensación]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Doc.compensación]").Activate
    Range("VALIDACION_CONSTANCIA[Doc.compensación]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub

Sub COLUMNA_FECHA_COMPENSACION()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_fecha_compensacion As String
'Workbooks("CENTRAL_DATA_SAP.xlsm").Worksheets("REPORTE_SAP").Activate
Sheets("REPORTE_SAP").Activate

columnas = "A10:O10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Fecha compensación")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_fecha_compensacion = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_fecha_compensacion) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Fecha compensación]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Fecha compensación]").Activate
    Range("VALIDACION_CONSTANCIA[Fecha compensación]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub

Sub COLUMNA_NOMBRE_USUARIO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_nombre_usuario As String

Sheets("REPORTE_SAP").Activate

columnas = "O10:Z10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Nombre del usuario")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_nombre_usuario = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_nombre_usuario) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Nombre del usuario]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Nombre del usuario]").Activate
    Range("VALIDACION_CONSTANCIA[Nombre del usuario]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub


Sub COLUMNA_CUENTA_DE_MAYOR()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_cuenta_mayor As String

Sheets("REPORTE_SAP").Activate

columnas = "O10:Z10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Cuenta de mayor")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_cuenta_mayor = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_cuenta_mayor) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Cuenta de mayor]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Cuenta de mayor]").Activate
    Range("VALIDACION_CONSTANCIA[Cuenta de mayor]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub


Sub COLUMNA_CODIGO_TRANSACCION()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_codigo_transaccion As String

Sheets("REPORTE_SAP").Activate

columnas = "O10:Z10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Código transacción")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_codigo_transaccion = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_codigo_transaccion) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Código transacción]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Código transacción]").Activate
    Range("VALIDACION_CONSTANCIA[Código transacción]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub


Sub COLUMNA_CTA_CONTRAPARTIDA()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_cta_contrapartida As String
Dim Valor_cta_contrapartida_validacion As String
Dim ws As Worksheet

Set ws = Sheets("VALIDACION")

With Sheets("REPORTE_SAP")
'    On Error GoTo ErrorHandler
    Set C = .Range("O10:Z10").Find("Cta.contrapartida")
    
    letra_columna = Split(.Cells(1, C.Column).Address, "$")(1)
    NOMBRE = letra_columna & C.Row
    Valor_cta_contrapartida = .Range(NOMBRE).Value
    .Range("DATA_SAP_FBLN[" & Valor_cta_contrapartida & "]").Copy
End With

ws.Range("VALIDACION_CONSTANCIA[Cta.contrapartida]").Select
ActiveSheet.Paste

    Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Range("VALIDACION_CONSTANCIA[Cta.contrapartida]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub


Sub COLUMNA_CLAVE_REFERENCIA()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Dim NOMBRE As String
Dim Valor_clave_referencia As String

Sheets("REPORTE_SAP").Activate

columnas = "O10:Z10"
On Error GoTo ErrorHandler
    With Sheets("REPORTE_SAP").Range(columnas)
            Set C = .Find("Clave de referencia")
            
            letra_columna = Split(Cells(1, C.Column).Address, "$")(1)
            NOMBRE = CStr(letra_columna) + CStr(C.Row)
            Valor_clave_referencia = Range(NOMBRE).Value
            Range("DATA_SAP_FBLN[" + CStr(Valor_clave_referencia) + "]").Activate
            Selection.Copy
            
            Sheets("VALIDACION").Activate
            Range("B1").Select
            'Sheets("SAP_PARAMETRIZADA").Select
            Range("VALIDACION_CONSTANCIA[Clave de referencia]").Activate
            ActiveSheet.Paste
    End With
   
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
   
   
Exit Sub
ErrorHandler:
    Sheets("VALIDACION").Activate
     Range("VALIDACION_CONSTANCIA[Clave de referencia]").Activate
    Range("VALIDACION_CONSTANCIA[Clave de referencia]").Value = "NO SE HA PODIDO ENCONTRAR DATOS"

End Sub

