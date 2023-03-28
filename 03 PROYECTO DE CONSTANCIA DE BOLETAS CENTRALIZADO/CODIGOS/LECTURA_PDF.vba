Attribute VB_Name = "LECTURA_PDF"
Sub mostrar_credenciales()
Attribute mostrar_credenciales.VB_ProcData.VB_Invoke_Func = " \n14"
' ocultar_desocultar Macro
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("CREDENCIALES SAP").Visible = True
End Sub


Sub ciopiar_pdf_75()

ruta_pdf = Worksheets("REPORTE_SAP").Range("B3").Value
sourcePath = ruta_pdf

ruta_pdf2 = Worksheets("REPORTE_SAP").Range("B4").Value
sourcePath2 = ruta_pdf2

ruta_pdf3 = Worksheets("REPORTE_SAP").Range("B5").Value
sourcePath3 = ruta_pdf3

ruta_pdf4 = Worksheets("REPORTE_SAP").Range("B6").Value
sourcePath4 = ruta_pdf4

ruta_pdf5 = Worksheets("REPORTE_SAP").Range("B7").Value
sourcePath5 = ruta_pdf5

ruta_pdf6 = Worksheets("REPORTE_SAP").Range("B8").Value
sourcePath6 = ruta_pdf6

Dim rutas() As String
rutas = Array("Z:\VARIOS\CONSTANCIAS CSC-NEXA\Atacocha\2023\01.2023\", "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Cajamarquilla\01.2023\", "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Cerro Lindo\01.2023\", "Z:\VARIOS\CONSTANCIAS CSC-NEXA\El Porvenir\01.2023\", "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Lima\01.2023\", "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Pampa\01.2023\")

Dim destino As String
destino = "C:\Macros\PROTOTIPO CONSTANCIAS\CONSTANCIAS\"

For i = 0 To 5
    Dim archivos() As String
    archivos = IO.Directory.GetFiles(rutas(i), "*.pdf")
    
    For Each archivo In archivos
        FileCopy archivo, destino & IO.Path.GetFileName(archivo)
    Next archivo
Next i

End Sub



 Sub COPIAR_PDF_DESDE_RUTA_DHO()

'Funcion = "COPIA PDF´S DESDE RUTA DE DHO"
'escribirLog Funcion, "Inicio del proceso de importación de los pdf´s de la ruta de DHO a carpeta de procesamiento"
' Declare variables for the source and destination paths.
' Asegurate de cambiar las rutas por las correspondientes a tus archivos y carpetas

Dim sourcePath As String
Dim ruta_pdf As String


ruta_pdf = Worksheets("REPORTE_SAP").Range("B3").Value

sourcePath = ruta_pdf

Dim ruta_historial_pdf As String

ruta_historial_pdf = "C:\Macros\PROTOTIPO CONSTANCIAS\HISTORIAL PDF\"

Dim destinationPath As String
destinationPath = "C:\Macros\PROTOTIPO CONSTANCIAS\CONSTANCIAS\"


'Archivos_pdf = Dir(sourcePath + "*.pdf")
Dim contador As Integer
contador = 0

myPath = Dir(sourcePath + "*.pdf")

Do While myPath <> ""
        'contador = contador + 1
        
        'Construir la ruta del archivo
        Dim src As String
        src = sourcePath & myPath
        
        'Copiar archivos a la carpeta en el escritorio
        
        FileCopy src, destinationPath & myPath
       
       'Mover el archivo
        myPath = Dir()
        'escribirLog Funcion, "PDF #" & contador & "ha sido copiado"
Loop

'MsgBox ("Hola")


End Sub
Sub ciopiar_pdf_6()

ruta_pdf = Worksheets("REPORTE_SAP").Range("B3").Value
sourcePath = ruta_pdf

ruta_pdf2 = Worksheets("REPORTE_SAP").Range("B4").Value
sourcePath2 = ruta_pdf2

ruta_pdf3 = Worksheets("REPORTE_SAP").Range("B5").Value
sourcePath3 = ruta_pdf3

ruta_pdf4 = Worksheets("REPORTE_SAP").Range("B6").Value
sourcePath4 = ruta_pdf4

ruta_pdf5 = Worksheets("REPORTE_SAP").Range("B7").Value
sourcePath5 = ruta_pdf5

ruta_pdf6 = Worksheets("REPORTE_SAP").Range("B8").Value
sourcePath6 = ruta_pdf6

Dim rutas() As String
rutas = Array(sourcePath, "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Cajamarquilla\2023\01.2023\", "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Cerro Lindo\2023\01.2023\", "Z:\VARIOS\CONSTANCIAS CSC-NEXA\El Porvenir\2023\01.2023\", "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Lima\2023\01.2023\", "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Pampa\2023\01.2023\")

Dim destino As String
destino = "C:\Macros\PROTOTIPO CONSTANCIAS\CONSTANCIAS\"

For i = 0 To 5
    Dim archivos() As String
    archivos = IO.Directory.GetFiles(rutas(i), "*.pdf")
    
    For Each archivo In archivos
        FileCopy archivo, destino & IO.Path.GetFileName(archivo)
    Next archivo
Next i



End Sub

Sub COPIAR_PDF_DESDE_RUTA_DHO_2()

'Funcion = "COPIA PDF´S DESDE RUTA DE DHO"
'escribirLog Funcion, "Inicio del proceso de importación de los pdf´s de la ruta de DHO a carpeta de procesamiento"
' Declare variables for the source and destination paths.
' Asegurate de cambiar las rutas por las correspondientes a tus archivos y carpetas




Dim sourcePath As String
Dim ruta_pdf As String
Dim sourcePath2 As String
Dim sourcePath3 As String
Dim sourcePath4 As String
Dim sourcePath5 As String
Dim sourcePath6 As String
Dim ruta_ano As String
Dim ruta_mes As String

Dim ruta_atacocha As String
Dim ruta_cajamarquilla As String
Dim ruta_Cerro_Lindo As String
Dim ruta_El_Porvenir As String
Dim ruta_Lima As String
Dim ruta_Pampa As String




ruta_ano = Worksheets("REPORTE_SAP").Range("B4").Value
ruta_mes = Worksheets("REPORTE_SAP").Range("B6").Value


ruta_atacocha = "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Atacocha\"
ruta_cajamarquilla = "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Cajamarquilla\"
ruta_Cerro_Lindo = "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Cerro Lindo\"
ruta_El_Porvenir = "Z:\VARIOS\CONSTANCIAS CSC-NEXA\El Porvenir\"
ruta_Lima = "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Lima\"
ruta_Pampa = "Z:\VARIOS\CONSTANCIAS CSC-NEXA\Pampa\"




ruta_pdf = ruta_atacocha & ruta_ano & "\" & ruta_mes & "\"
MsgBox (ruta_pdf)
sourcePath = ruta_pdf

ruta_pdf2 = ruta_cajamarquilla & ruta_ano & "\" & ruta_mes & "\"
MsgBox (ruta_pdf2)
sourcePath2 = ruta_pdf2

ruta_pdf3 = ruta_Cerro_Lindo & ruta_ano & "\" & ruta_mes & "\"
sourcePath3 = ruta_pdf3

ruta_pdf4 = ruta_El_Porvenir & ruta_ano & "\" & ruta_mes & "\"
sourcePath4 = ruta_pdf4

ruta_pdf5 = ruta_Lima & ruta_ano & "\" & ruta_mes & "\"
sourcePath5 = ruta_pdf5

ruta_pdf6 = ruta_Pampa & ruta_ano & "\" & ruta_mes & "\"
sourcePath6 = ruta_pdf6



ruta_historial_pdf = "C:\Macros\PROTOTIPO CONSTANCIAS\HISTORIAL PDF\"

Dim destinationPath As String
destinationPath = "C:\Macros\PROTOTIPO CONSTANCIAS\CONSTANCIAS\"


'Archivos_pdf = Dir(sourcePath + "*.pdf")

myPath = Dir(sourcePath + "*.pdf")

Do While myPath <> ""
        'contador = contador + 1
        
        'Construir la ruta del archivo
        Dim src As String
        src = sourcePath & myPath
        
        'Copiar archivos a la carpeta en el escritorio
        
        FileCopy src, destinationPath & myPath
       
       'Mover el archivo
        myPath = Dir()
        'escribirLog Funcion, "PDF #" & contador & "ha sido copiado
Loop

myPath2 = Dir(sourcePath2 + "*.pdf")

Do While myPath2 <> ""
        'contador = contador + 1
        
        'Construir la ruta del archivo
        Dim src2 As String
        src2 = sourcePath2 & myPath2
        
        'Copiar archivos a la carpeta en el escritorio
        
        FileCopy src2, destinationPath & myPath2
       
       'Mover el archivo
        myPath2 = Dir()
        'escribirLog Funcion, "PDF #" & contador & "ha sido copiado
Loop


myPath3 = Dir(sourcePath3 + "*.pdf")

Do While myPath3 <> ""
        'contador = contador + 1
        
        'Construir la ruta del archivo
        Dim src3 As String
        src3 = sourcePath3 & myPath3
        
        'Copiar archivos a la carpeta en el escritorio
        
        FileCopy src3, destinationPath & myPath3
       
       'Mover el archivo
        myPath3 = Dir()
        'escribirLog Funcion, "PDF #" & contador & "ha sido copiado
Loop


myPath4 = Dir(sourcePath4 + "*.pdf")

Do While myPath4 <> ""
        'contador = contador + 1
        
        'Construir la ruta del archivo
        Dim src4 As String
        src4 = sourcePath4 & myPath4
        
        'Copiar archivos a la carpeta en el escritorio
        
        FileCopy src4, destinationPath & myPath4
       
       'Mover el archivo
        myPath4 = Dir()
        'escribirLog Funcion, "PDF #" & contador & "ha sido copiado
Loop


myPath5 = Dir(sourcePath5 + "*.pdf")

Do While myPath5 <> ""
        'contador = contador + 1
        
        'Construir la ruta del archivo
        Dim src5 As String
        src5 = sourcePath5 & myPath5
        
        'Copiar archivos a la carpeta en el escritorio
        
        FileCopy src5, destinationPath & myPath5
       
       'Mover el archivo
        myPath5 = Dir()
        'escribirLog Funcion, "PDF #" & contador & "ha sido copiado
Loop


myPath6 = Dir(sourcePath6 + "*.pdf")

Do While myPath6 <> ""
        'contador = contador + 1
        
        'Construir la ruta del archivo
        Dim src6 As String
        src6 = sourcePath6 & myPath6
        
        'Copiar archivos a la carpeta en el escritorio
        
        FileCopy src6, destinationPath & myPath6
       
       'Mover el archivo
        myPath6 = Dir()
        'escribirLog Funcion, "PDF #" & contador & "ha sido copiado
Loop


End Sub

Function MOVER_PDFS_AL_TERMINAR()

'Declare variables
Dim sourceFolder As String, targetFolder As String, histFolder As String
Dim currentDate As String, currentTime As String
Dim file As String

'Set the source folder path
sourceFolder = "C:\Macros\PROTOTIPO CONSTANCIAS\CONSTANCIAS\"

'Set the hist folder path
currentDate = Format(Now(), "yyyy-mm-dd")
currentTime = Format(Now(), "hh-mm-ss")
histFolder = "C:\Macros\PROTOTIPO CONSTANCIAS\HISTORIAL PDF\" & currentDate & "_" & currentTime

'Create the hist folder
MkDir histFolder

'Loop through all files in the source folder
file = Dir(sourceFolder & "*.pdf")
Do While file <> ""
    'Move the file to the hist folder
    Name sourceFolder & file As histFolder & "\" & file
    'Get the next file
    file = Dir()
Loop

End Function

Sub DESARROLLO_PROCESO_PDF()

Funcion = "PROCESO DE LECTURA DE PDF"
escribirLog Funcion, "Inicio del proceso de LECTURA DE PDF´S"

Application.DisplayAlerts = False

Call ELIMINAR_REGISTROS_PDF


On Error Resume Next
Dim carpeta_reporte As String
Dim contador As Integer
Dim Nombre_archivo As String
Dim nombre_pdf As String
Dim Ruta_constancia As String
Dim Nueva_ruta As String
Dim myPath_new As String
Dim ws As Worksheet
Set ws = Sheets("PROCESO")
Dim ws_data_pdf As Worksheet
Set ws_data_pdf = Sheets("DATA_PDF")

Dim ruta_unidad As String
Dim modelo_unidad As String
Dim unidad As String
Dim MODELO As String
Dim ruta_ano As String
Dim ruta_ano_mes As String
Dim Nombre_sin_procesar As String

'Dim ruta_Pampa As String

ruta_ano = Worksheets("REPORTE_SAP").Range("B4").Value
ruta_ano_mes = Worksheets("REPORTE_SAP").Range("B5").Value & "." & Worksheets("REPORTE_SAP").Range("B4").Value

ruta_unidad = "Z:\VARIOS\CONSTANCIAS CSC-NEXA\"

Dim fso As New FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

ruta = "C:\Macros\PROTOTIPO CONSTANCIAS\CONSTANCIAS\"
Ruta_constancia = "C:\Macros\PROTOTIPO CONSTANCIAS\CONSTANCIAS\"
libro_sistema = ActiveWorkbook.Name
Ruta_mover_pdf = "C:\Macros\PROTOTIPO CONSTANCIAS\HISTORIAL PDF\"
historial_pdf = "C:\Macros\PROTOTIPO CONSTANCIAS\HISTORIAL PDF\"


'Workbooks(libro_sistema).Activate



Archivos_pdf = Dir(ruta + "*.pdf")
myPath = Dir(ruta + "*.pdf")

contador = 0
Do While myPath <> ""

        contador = contador + 1
        If contador = 30 Then
            Set oServ = GetObject("winmgmts:")
            Set cProc = oServ.ExecQuery("Select * from Win32_Process")
        
            For Each oProc In cProc
                If oProc.Name = "AcroRd32.exe" Then
                    oProc.Terminate
                    contador = 0
                Exit For
                End If
            Next
            Set openPDF = CreateObject("Shell.Application")
            openPDF.Open (myPath)
        End If
        
        nombre_pdf = myPath
        myPath = ruta + myPath
        Set openPDF = CreateObject("Shell.Application")
        'Set pdf = GetObject(myPath)
        openPDF.Open (myPath)
        'copiedContent = pdf.Content.Text
        Application.Wait Now + TimeValue("00:00:05")
        SendKeys "^a"
        Application.Wait Now + TimeValue("00:00:05")
        SendKeys "^c"
        Application.Wait Now + TimeValue("00:00:05")
        
        ws.Select
        Workbooks(libro_sistema).Activate
        ws.Select
        
        ' Convert the date format to match Excel sheet format
        
'        copiedContent = Application.GetClipboard
'        copiedContent = DateValue(copiedContent)
'        copiedContent = Format(copiedContent, "dd/mm/yyyy")
        
        ' Copy the modified date format back to clipboard
'        Application.SetClipboard copiedContent
'
'        copiedContent = Replace(copiedContent, "\/", "-")
'        copiedContent = Replace(copiedContent, "-", "/")
        
        ws.Columns("C").Select
        Selection.ClearFormats
        ws.Range("C10:C1000").Value = DateValue(ws.Range("A10:A1000").Value)
        ws.Range("C10").Activate
        ActiveSheet.Paste
        
         
        'ENCONTRANDO NOMBRE DE EMPLEADO
         With ws.Cells
            'Referencia de Planilla
    
            Set oFound_referencia_planilla = .Find(What:="Referencia de planilla:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
            If Not oFound_referencia_planilla Is Nothing Then
            
                Application.GoTo oFound_referencia_planilla, True
                fila_match_referencia_planilla = ActiveCell.Row
                columna_match_referencia_planilla = ActiveCell.Column
                'nombre_empleado = Range("A" + CStr(fila_match + 3)).Value
                'referencia_planilla = Cells(fila_match_referencia_planilla + 1, columna_match_referencia_planilla).Value
                referencia_planilla = CStr(Cells(fila_match_referencia_planilla + 1, columna_match_referencia_planilla).Value)
            ElseIf Not oFound_referencia = .Find(What:="Referencia:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) Is Nothing Then
                Application.GoTo oFound_referencia, True
                fila_match_planilla = ActiveCell.Row
                columna_match_planilla = ActiveCell.Column
                'referencia_planilla = Cells(fila_match_planilla + 1, columna_match_planilla).Value
                referencia_planilla = CStr(Cells(fila_match_referencia_planilla + 1, columna_match_referencia_planilla).Value)
            Else
                 referencia_planilla = "Referencia no encontrada"
            End If
            
            
             
             'escribirLog Funcion, "PDF # " & contador & " : Extracción de numero de referencia del pdf"
            'Fecha de Proceso
            
            Set oFound_fecha_proceso = .Find(What:="Fecha de proceso:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
            If Not oFound_fecha_proceso Is Nothing Then
                Application.GoTo oFound_fecha_proceso, True
                fila_match_fecha_proceso = ActiveCell.Row
                columna_match_fecha_proceso = ActiveCell.Column
                fecha_proceso = Cells(fila_match_fecha_proceso + 1, columna_match_fecha_proceso).Value
                fecha_proceso = Format(fecha_proceso, "dd/mm/yyyy")
                
            ElseIf Not oFound_fecha = .Find(What:="Fecha:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) Is Nothing Then
                Application.GoTo oFound_fecha, True
                fila_match_fecha = ActiveCell.Row
                columna_match_fecha = ActiveCell.Column
                fecha_proceso = Cells(fila_match_fecha + 1, columna_match_fecha).Value
                fecha_proceso = Format(fecha_proceso, "dd/mm/yyyy")
            Else
                 fecha_proceso = "Fecha no encontrada"
            End If
            
            
             'escribirLog Funcion, "PDF # " & contador & " : Extracción de fecha del proceso del pdf"
            
            'Cuenta de Origen
            '
            Set oFound_cuenta_origen = .Find(What:="Cuenta deorigen:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
            If Not oFound_cuenta_origen Is Nothing Then
                Application.GoTo oFound_cuenta_origen, True
                fila_match_cuenta_origen = ActiveCell.Row
                columna_match_cuenta_origen = ActiveCell.Column
                cuenta_origen = CStr(Cells(fila_match_cuenta_origen + 1, columna_match_cuenta_origen).Value)
            
            ElseIf Not oFound_cuenta_origen_junto = .Find(What:="Cuenta de origen:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) Is Nothing Then
                Application.GoTo oFound_cuenta_origen_junto, True
                fila_match_cuenta_junto = ActiveCell.Row
                columna_match_cuenta_junto = ActiveCell.Column
                cuenta_origen = CStr(Cells(fila_match_cuenta_junto + 1, columna_match_cuenta_junto).Value)
            ElseIf Not oFound_cuenta = .Find(What:="Cuenta:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) Is Nothing Then
                Application.GoTo oFound_cuenta, True
                fila_match_cuenta = ActiveCell.Row
                columna_match_cuenta = ActiveCell.Column
                cuenta_origen = Cells(fila_match_cuenta + 1, columna_match_cuenta).Value
                cuenta_origen = CStr(Cells(fila_match_cuenta_junto + 1, columna_match_cuenta_junto).Value)
            Else
                 cuenta_origen = " Número de cuenta no encontrada"
            End If
             
            'escribirLog Funcion, "PDF # " & contador & " : Extracción de Cuenta de Origen del pdf"
            
            'Monto Total
            
            Set oFound_monto_total = .Find(What:="Monto total:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
            If Not oFound_monto_total Is Nothing Then
                Application.GoTo oFound_monto_total, True
                fila_match_monto_total = ActiveCell.Row
                columna_match_monto_total = ActiveCell.Column
                monto_total = CInt(Cells(fila_match_monto_total + 1, columna_match_monto_total).Value)
            
            ElseIf Not oFound_monto = .Find(What:="Monto:", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) Is Nothing Then
                Application.GoTo oFound_monto, True
                fila_match_monto = ActiveCell.Row
                columna_match_monto = ActiveCell.Column
                monto_total = CInt(Cells(fila_match_monto + 1, columna_match_monto).Value)
            Else
                 monto_total = "Monto no encontrado"
            End If
            
            
            ws_data_pdf.Select
            'Obtener el nombre del archivo
            Nombre_archivo = fso.GetFileName(myPath)
            'Obtener el nombre del archivo sin extensión
            FileName_extraido = ExtraerNombre(Nombre_archivo)
            Nombre_sin_procesar = FileName_extraido & ".pdf"
            unidad = Mid(FileName_extraido, 1, 6)
            'ruta = """",IF(unidad=""7053_M"",""El Porvenir"",IF([@UNIDAD]=""7010_L"",""Lima"",IF([@UNIDAD]="",""Pampa"")))))))
            MODELO = ObtenerModeloUnidad(unidad)
            Nueva_ruta = ConcatenarRuta(ruta_unidad, MODELO, ruta_ano, ruta_ano_mes, Nombre_sin_procesar)
            
            
            
            
            'FileName_extraido_v1 = Left(Nombre_archivo, InStr(Nombre_archivo, "") - 1)
            'MsgBox (myPath & FileName_extraido & referencia_planilla & fecha_proceso & cuenta_origen & monto_total)
            
            'MsgBox (ruta_unidad & "" \ "" & ruta_ano & "" \ "" & ruta_ano_mes)
            
            Set HojaDatos = ThisWorkbook.Sheets("DATA_PDF")
            Set tabla = HojaDatos.ListObjects("BASE_DE_DATOS_CONSTANCIAS_PDF")
            Set NuevaFila = tabla.ListRows.Add

            ws_data_pdf.Select
            

            With NuevaFila
                .Range(2) = FileName_extraido
                .Range(3) = "=MID([@[NOMBRE DE LA CONSTANCIA]],1,6)"
                .Range(4) = "=IF([@UNIDAD]=""7022_A"",""Atacocha"",IF([@UNIDAD]=""7042_C"",""Cajamarquilla"",IF([@UNIDAD]=""7010_C"",""Cerro Lindo"",IF([@UNIDAD]=""7053_M"",""El Porvenir"",IF([@UNIDAD]=""7010_L"",""Lima"",IF([@UNIDAD]=""7056_P"",""Pampa""))))))"
                .Range(5) = Nueva_ruta
                .Range(6) = referencia_planilla
                '.Range(7).Select
                'Selection.ClearFormats
                .Range(7) = fecha_proceso
                .Range(8) = cuenta_origen
                .Range(9) = monto_total
                .Range(10) = "=IF(AND([@Referencia]=""Nombre no encontrado"",[@[Numero de Cuenta]]=""Nombre no encontrado"",[@Monto])=TRUE,""CONSTANCIA BBVA"",""CONSTANCIA BCP"")"
                .Range(1) = "=IF([@Referencia]=""Nombre no encontrado"",""NO TIENE REFERENCIA VALIDA EN EL DOCUMENTO"",IFERROR(TEXT(CONCAT(""T"",MID([@Referencia],FIND(""-"",[@Referencia],1)+1,(FIND(""P"",[@Referencia],1)-FIND(""-"",[@Referencia],1)-1)),MID([@Referencia],FIND(""P"",[@Referencia],1),1),RIGHT(LEFT([@Referencia],FIND(""-"",[@Referencia],1)-1),2),MID(LEFT([@Referencia],FIND(""-"",[" & _
        "@Referencia],1)-1),5,2),LEFT(LEFT([@Referencia],FIND(""-"",[@Referencia],1)-1),4)),),[@Referencia]))" & _
        ""

            End With
            
            
            'escribirLog Funcion, "PDF # " & contador & " pegado en tabla"
            
            '**CERRAR ADOBE ACROBAT**
            Set oServ = GetObject("winmgmts:")
            Set cProc = oServ.ExecQuery("Select * from Win32_Process")
        
            For Each oProc In cProc
                If oProc.Name = "Acrobat.exe" Then
                    oProc.Terminate
                    contador = 0
                End If
            Exit For
                
            Next
            
         End With
        
    
        '
        Workbooks(libro_sistema).Activate
        ws.Select
        Cells.ClearContents
        
'        ws.Delete
'        Worksheets.Add.Name = "PROCESO"
        
    
        '***** FIN LIMPIEZA *****
        
        myPath_new = FileName_extraido
        myPath = Dir
        referencia_planilla = ""
        fecha_proceso = ""
        cuenta_origen = ""
        monto_total = ""
        
        
        'Workbooks(myPath).Close SaveChanges:=True
        
         '***** FIN LIMPIEZA *****
         'escribirLog Funcion, "PDF # " & contador & " terminado de procesar"
         
Loop

    '**CERRAR ADOBE ACROBAT**
    Set oServ = GetObject("winmgmts:")
    Set cProc = oServ.ExecQuery("Select * from Win32_Process")
    
    For Each oProc In cProc
        If oProc.Name = "Acrobat.exe" Then
            oProc.Terminate
            contador = 0
        Exit For
        End If
    Next


'MsgBox ("Proceso culminado con éxito")

End Sub
Function ObtenerModeloUnidad(unidad As String) As String
If unidad = "7022_A" Then
ObtenerModeloUnidad = "Atacocha"
ElseIf unidad = "7022_A" Then
ObtenerModeloUnidad = "Cajamarquilla"
ElseIf unidad = "7010_C" Then
ObtenerModeloUnidad = "Cerro Lindo"
ElseIf unidad = "7010_L" Then
ObtenerModeloUnidad = "Lima"
ElseIf unidad = "7053_M" Then
ObtenerModeloUnidad = "El Porvenir"
ElseIf unidad = "7056_P" Then
ObtenerModeloUnidad = "Pampa"
Else
ObtenerModeloUnidad = "No se ha encontrado la unidad"
End If
End Function

Sub modelo_225()

MODELO = ObtenerModeloUnidad("7022_A")
MsgBox (MODELO)


End Sub

Function ConcatenarRuta(ruta_unidad As String, modelo_unidad As String, ruta_ano As String, ruta_ano_mes As String, ruta_documento As String) As String
ConcatenarRuta = ruta_unidad & modelo_unidad & "\" & ruta_ano & "\" & ruta_ano_mes & "\" & ruta_documento
End Function

Function ExtraerNombre(ruta As String) As String
    Dim pos As Integer
    pos = InStr(ruta, "\")
    Do While pos > 0
        ruta = Right(ruta, Len(ruta) - pos)
        pos = InStr(ruta, "\")
    Loop
    pos = InStr(ruta, ".pdf")
    ExtraerNombre = Left(ruta, pos - 1)
End Function

Sub ELIMINAR_REGISTROS_PDF()
Funcion = "ELIMINAR REGISTROS DE PDF EN LA TABLA DEL SISTEMA"
escribirLog Funcion, "Inicio del proceso de eliminar datos de tabla"

On Error Resume Next
'Declara la variable de la tabla
Dim tabla As ListObject
Dim ws_data_pdf As Worksheet
Set ws_data_pdf = Sheets("DATA_PDF")

'Establece la tabla
Set tabla = ws_data_pdf.ListObjects("BASE_DE_DATOS_CONSTANCIAS_PDF")

'Elimina todos los registros de la tabla
tabla.DataBodyRange.Delete
escribirLog Funcion, "Fin del proceso de eliminar datos de tabla"

End Sub
