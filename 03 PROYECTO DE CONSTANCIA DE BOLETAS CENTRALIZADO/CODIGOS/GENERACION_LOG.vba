Attribute VB_Name = "GENERACION_LOG"

Public Sub ErrorHandler(Error As Long)
' Obtiene el código de error
'lngErrorCode = Err.Number

' Abre el archivo de texto en modo escritura
Open ActiveWorkbook.Path & "\Log\error_" & Format(Now(), "yyyy-mm-dd") & ".txt" For Append As #1

' Escribe el código de error y la descripción del error en el archivo
Print #1, Format(Now(), "yyyy-mm-dd") & "|" & Format(Now(), "HH:mm:ss") & Right(Format(Timer, "0.000"), 4) & "|" & "Error: " & Error & "|" & Err.Description & "|" & Err.Source

' Cierra el archivo
Close #1

logFile = ActiveWorkbook.Path & "\Log\log_" & Format(Now(), "yyyy-mm-dd") & ".txt"
Open logFile For Append As #1
    Print #1, Format(Now(), "yyyy-mm-dd") & "|" & Format(Now(), "HH:mm:ss") & Right(Format(Timer, "0.000"), 4) & "|" & BOT & "| |" & Err.Description & "|" & "ERROR"
Close #1
MsgBox "¡Error! " & Err.Description, vbCritical
End Sub

Function escribirLog(Proceso, mensaje1 As String)
    Dim logFile As String
    
    ' Establece la ruta del archivo de log- Podemos determinar la forma en como definir la ruta
    
    logFile = ActiveWorkbook.Path & "\Log\log_" & Format(Now(), "yyyy-mm-dd") & ".txt"
    Call crear_file(logFile)
    
    ' Abre el archivo de log en modo de escritura
    Open logFile For Append As #1
    Print #1, Format(Now(), "yyyy-mm-dd") & "|" & Format(Now(), "HH:mm:ss") & Right(Format(Timer, "0.000"), 4) & "|" & BOT & "|" & Proceso & "|" & mensaje1 & "|" & "SATISFACTORIO"
    Close #1
    
End Function

Sub crear_file(ruta As String)
'MsgBox (ruta)
If Dir(ruta, vbDirectory) = "" Then
'MsgBox (ruta)
    ' Establece la ruta del archivo de log
    logFile = ruta
    ' Abre el archivo de log en modo de escritura para colocar la cabecera
    Open logFile For Append As #1
    Print #1, "DIA" & "|" & "HORA|BOT|PROCESO|COMENTARIO" & "|" & "ESTADO"
    Close #1
Else
    Exit Sub
End If
    
End Sub

Public Sub crear_carpetalog(ruta As String)

ExisteCarpeta = Dir(ruta & "\Log", vbDirectory)

If ExisteCarpeta = "" Then
    MkDir ruta & "\Log"
End If

End Sub







'Sub CrearLog_file()
'    Dim logFile As String
'
'     Establece la ruta del archivo de log utilizando la hora actual
'    logFile = "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\P&P\2. PROYECTOS DE PRACTICANTE\PROYECTO DE VALIDACIÓN DE CONSTANCIAS DE PAGO Y COMPENSACIONES\PROTOTIPO CONSTANCIAS\ARCHIVO LOG\log_" & Format(Now(), "yyyy-mm-dd_hh-mm-ss") & ".txt"
'
'     Abre el archivo de log en modo de escritura
'    Open logFile For Output As #1
'
'     Escribe el mensaje de bienvenida en el archivo de log
'    Print #1, "Bienvenido al archivo de log"
'
'     Cierra el archivo de log
'    Close #1
'End Sub
'
'Function DescargarConLogSAP(reporte As String) As Boolean
'    Dim logFile As String
'    Dim resultado As Boolean
'     Obtiene la ruta del archivo de log de la celda A1
'    logFile = "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\P&P\2. PROYECTOS DE PRACTICANTE\PROYECTO DE VALIDACIÓN DE CONSTANCIAS DE PAGO Y COMPENSACIONES\PROTOTIPO CONSTANCIAS\ARCHIVO LOG\log_" & Format(Now(), "yyyy-mm-dd_hh-mm-ss") & ".txt"
'
'     Abre el archivo de log en modo de escritura
'    Open logFile For Append As #1
'
'     Escribe la hora actual y el nombre del reporte en el archivo de log
'    Print #1, Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Inicio de la descarga del reporte " & reporte
'
'     Llama a la función de descarga
'    resultado = SAP_EXTRACT_DATA_FBLN_REPORT()
'
'     Escribe la hora actual y el resultado de la descarga en el archivo de log
'    If resultado Then
'        Print #1, Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Descarga del reporte " & reporte & " completada"
'    Else
'        Print #1, Format(Now(), "yyyy-mm-dd hh:mm:ss") & " - Error al descargar el reporte " & reporte
'    End If
'
'     Cierra el archivo de log
'    Close #1
'
'     Devuelve el resultado de la descarga
'    DescargarConLogSAP = resultado
'End Function


