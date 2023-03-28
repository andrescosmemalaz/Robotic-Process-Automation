Attribute VB_Name = "GENERAL"
Sub MODELO_GENERAL()

'Existe la ruta
Dim logFile As String
logFile = ActiveWorkbook.Path
Call crear_carpetalog(logFile)

Funcion = "VALIDACION DE EXISTENCIA DE CORREO"
escribirLog Funcion, "Inicio del proceso principal de la Automatización de Boletas y Constancias de Pago"
    'Definir la ruta de origen y destino
    Dim rutaOrigen As String
    Dim rutaDestino As String
    Dim file_name_comp As String
    
    rutaOrigen = "\\vidpe00239\Users\LUISI\Desktop\Robots\Recursos\08_COE_DHO_Automtizacion_boletas_de_pago_PE\input\"
    rutaDestino = "C:\Macros\PROTOTIPO CONSTANCIAS\REPORTE PARA PROCESAR\"
    myPath = Dir(rutaOrigen + "*.xlsx")
    
    'Comprobar la existencia del archivo
    If (Dir(rutaOrigen + "*.xlsx") <> "") Then
        'Mover el archivo
        Name rutaOrigen + myPath As rutaDestino + myPath
        escribirLog Funcion, "Transporte de valores desde el compartido al sistema "
        Exportar_valores_sistema rutaDestino + myPath, "C:\Macros\PROTOTIPO CONSTANCIAS\PROYECTO VALIDACION CONSTANCIA.xlsm"
       
        file_name_comp = Split(myPath, ".")(0)
        'Call SAP_EXTRACT_DATA_FBLN_REPORT
        Call PEGAR_EXCEL_FBLN
        Call ELIMINAR_REGISTROS_PDF
        Call COPIAR_PDF_DESDE_RUTA_DHO
        Call DESARROLLO_PROCESO_PDF
        Call MOVIMIENTO_COLUMNAS_MASIVAS
        Call VALIDACION_DE_CONSTANCIA_AVANZADA
        Call CREACION_COLUMNAS_VALIDACION_FINAL
        Call ELIMINAR_DATOS_PROCESO_VALIDACION_AVANZADA
       
       CopiarHoja_guardar_enviar_por_correo file_name_comp
       
       escribirLog Funcion, "Envío de correo al Cliente con el reporte procesado realizado"
       
       
    Else
        'Enviar correo electrónico
        Dim Mail_Object, Mail_Single As Variant
        Set Mail_Object = CreateObject("Outlook.Application")
        Set Mail_Single = Mail_Object.CreateItem(0)
        With Mail_Single
            .Subject = "Archivo Excel no encontrado"
            .To = "andres.maximo.cosme.malasquez@gmail.com"
            .Body = "No se ha encontrado archivos para Procesar"
            .Send
        End With
        escribirLog Funcion, "ENVÍO DE CORREO AL NO ENCONTRAR ARCHIVOS EN EL COMPARTIDO"
        
    End If
    
    'Clean up memory
    Set Mail_Single = Nothing
    Set Mail_Object = Nothing
    escribirLog Funcion, "Fin del proceso principal de la Automatización de Boletas y Constancias de Pago"
    
End Sub
Sub MODELO_GENERAL_V2()

'Existe la ruta
Dim logFile As String
logFile = ActiveWorkbook.Path
Call crear_carpetalog(logFile)

Funcion = "VALIDACION DE EXISTENCIA DE CORREO"
escribirLog Funcion, "Inicio del proceso principal de la Automatización de Boletas y Constancias de Pago"
    'Definir la ruta de origen y destino
    Dim rutaOrigen As String
    Dim rutaDestino As String
    Dim file_name_comp As String
    
    rutaOrigen = "\\vidpe00239\Users\LUISI\Desktop\Robots\Recursos\08_COE_DHO_Automtizacion_boletas_de_pago_PE\input\"
    rutaDestino = "C:\Macros\PROTOTIPO CONSTANCIAS\REPORTE PARA PROCESAR\"
    myPath = Dir(rutaOrigen + "*.xlsx")
    
    'Ciclo para recorrer todos los archivos excel en la ruta de origen
    Do While myPath <> ""
        'Comprobar la existencia del archivo
        If (Dir(rutaOrigen + myPath) <> "") Then
            'Mover el archivo
            Name rutaOrigen + myPath As rutaDestino + myPath
            escribirLog Funcion, "Transporte de valores desde el compartido al sistema "
            Exportar_valores_sistema rutaDestino + myPath, "C:\Macros\PROTOTIPO CONSTANCIAS\PROYECTO VALIDACION CONSTANCIA.xlsm"
           
            file_name_comp = Split(myPath, ".")(0)
            'Call SAP_EXTRACT_DATA_FBLN_REPORT
            Call PEGAR_EXCEL_FBLN
            Call ELIMINAR_REGISTROS_PDF
            Call COPIAR_PDF_DESDE_RUTA_DHO
            Call DESARROLLO_PROCESO_PDF
            Call MOVIMIENTO_COLUMNAS_MASIVAS
            Call VALIDACION_DE_CONSTANCIA_AVANZADA
            Call CREACION_COLUMNAS_VALIDACION_FINAL
            Call ELIMINAR_DATOS_PROCESO_VALIDACION_AVANZADA
           
           CopiarHoja_guardar_enviar_por_correo file_name_comp
           
           escribirLog Funcion, "Envío de correo al Cliente con el reporte procesado realizado"
        Else
            'Enviar correo electrónico
            Dim Mail_Object, Mail_Single As Variant
            Set Mail_Object = CreateObject("Outlook.Application")
            Set Mail_Single = Mail_Object.CreateItem(0)
            With Mail_Single
            .Subject = "Archivo Excel no encontrado"
            .To = "andres.maximo.cosme.malasquez@gmail.com"
            .Body = "No se ha encontrado el archivo " & myPath & " para procesar"
            .Send
            End With
            escribirLog Funcion, "ENVÍO DE CORREO AL NO ENCONTRAR ARCHIVOS EN EL COMPARTIDO"
        End If
        'Obtener el siguiente archivo en la ruta de origen
        myPath = Dir()
        Loop
            
        'Clean up memory
        Set Mail_Single = Nothing
        Set Mail_Object = Nothing
            
escribirLog Funcion, "Fin del proceso principal de la Automatización de Boletas y Constancias de Pago"
End Sub



Sub MOSTRAR_NOMBRE_DE_CARPETAS()

'Declarar variables
Dim ruta As String
Dim fso As Object
Dim carpeta As Object
Dim hoja As Worksheet
Dim nueva_cadena As String
Dim posicion As String
Dim primer_point As String
Dim primer_filtro As String
Dim second_point As String
Dim second_filtro As String
Dim last_filtro As String



Dim i As Integer

'Establecer ruta de la carpeta a recorrer
ruta = "Z:\CSC\SERVICIOS DHO\10. Nomina\01. Medios de pago\"

'Crear objeto de archivos
Set fso = CreateObject("Scripting.FileSystemObject")

'Crear objeto de la carpeta
Set carpeta = fso.GetFolder(ruta)

'Establecer hoja de Excel para escribir los resultados
Set hoja = ThisWorkbook.Sheets("PROCESO")

'Iniciar recorrido de las carpetas
For Each carpeta In carpeta.SubFolders
    nueva_cadena = ""
    posicion = ""
    primer_point = ""
    primer_filtro = ""
    second_point = ""
    second_filtro = ""
    last_filtro = ""
    'Recorrer cada caracter de la cadena
    For i = 1 To Len(carpeta.Name)
        'Verificar si el caracter no es un número
        If Not IsNumeric(Mid(carpeta.Name, i, 1)) Then
            'Agregar el caracter a la nueva cadena
            nueva_cadena = nueva_cadena & Mid(carpeta.Name, i, 1)
            posicion = InStr(1, nueva_cadena, ".")
            primer_point = Left(nueva_cadena, posicion)
            primer_filtro = Trim(EliminarPalabra(nueva_cadena, primer_point))
            second_point = InStr(1, primer_filtro, ".")
            second_filtro = Trim(Left(primer_filtro, second_point))
            last_filtro = Trim(EliminarPalabra(primer_filtro, second_filtro))
        End If
    Next i
    hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Offset(1, 0) = nueva_cadena
    hoja.Cells(hoja.Rows.Count, 2).End(xlUp).Offset(1, 0) = posicion
    hoja.Cells(hoja.Rows.Count, 3).End(xlUp).Offset(1, 0) = primer_point
    hoja.Cells(hoja.Rows.Count, 4).End(xlUp).Offset(1, 0) = primer_filtro
    hoja.Cells(hoja.Rows.Count, 5).End(xlUp).Offset(1, 0) = second_point
    hoja.Cells(hoja.Rows.Count, 6).End(xlUp).Offset(1, 0) = second_filtro
    hoja.Cells(hoja.Rows.Count, 7).End(xlUp).Offset(1, 0) = last_filtro
    
    
Next

'Mensaje de finalización
MsgBox "La lista de carpetas sin números ha sido escrita en la hoja Hoja1."

End Sub

Function EliminarPalabra(cadena As String, palabra As String) As String
    Dim pos As Integer
    pos = InStr(1, cadena, palabra, vbTextCompare)
    If pos > 0 Then
        EliminarPalabra = Left(cadena, pos - 1) & Right(cadena, Len(cadena) - pos - Len(palabra) + 1)
    Else
        EliminarPalabra = cadena
    End If
End Function
