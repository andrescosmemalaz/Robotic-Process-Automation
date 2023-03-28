Attribute VB_Name = "MOVIMIENTO_DE_ARCHIVOS"
Sub CopiarHoja_guardar_enviar_por_correo(Nombre_unidad As String)

Dim correo As Object
Dim adjunto As Object
Dim hoja As Worksheet
Dim libroNuevo As Workbook

'Nombre de la hoja que desea enviar
Set hoja = ThisWorkbook.Sheets("VALIDACION")

'Crear un nuevo libro de Excel y copiar la hoja
Set libroNuevo = Workbooks.Add
hoja.Copy After:=libroNuevo.Sheets(libroNuevo.Sheets.Count)
'libroNuevo.Sheets(1).Name = "REPORTE_VALIDACION" & Format(Now(), "yyyy-mm-dd")

'Guardar el libro con un nuevo nombre
libroNuevo.SaveAs "C:\Macros\PROTOTIPO CONSTANCIAS\HISTORIAL REPORTES ANALIZADOS\REPORTE_VALIDACION_" & Nombre_unidad & "_" & Format(Now(), "yyyy-mm-dd") & ".xlsx"
libroNuevo.Close

'Crear un nuevo correo electrónico
Set correo = CreateObject("Outlook.Application").CreateItem(0)

'Asunto, destinatario y mensaje
correo.Subject = "Estado de los reportes"
correo.To = "andres.maximo.cosme.malasquez@gmail.com"
correo.Body = "Buenas tardes, se envía el estado de los reportes en el archivo adjunto."

'Agregar el nuevo libro como un archivo adjunto
Set adjunto = correo.Attachments.Add("C:\Macros\PROTOTIPO CONSTANCIAS\HISTORIAL REPORTES ANALIZADOS\REPORTE_VALIDACION_" & Nombre_unidad & "_" & Format(Now(), "yyyy-mm-dd") & ".xlsx")

'Enviar correo
correo.Send

'Limpiar la memoria
Set correo = Nothing
Set adjunto = Nothing
Set hoja = Nothing

End Sub

'Nota: Recuerda cambiar el "NombreHoja" y "email@example.com" para adaptarlo a tu necesidad
'
'Es importante mencionar que este código requiere que tengas Outlook instalado en tu computadora y configurado con una cuenta de correo electrónico para poder enviar correos electrónicos.
Sub TRAER_EXCEL_DESDE_VOTOBOT_ONLINE()
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
    
       Exportar_valores_sistema rutaDestino + myPath, "C:\Macros\PROTOTIPO CONSTANCIAS\PROYECTO VALIDACION CONSTANCIA.xlsm"
       file_name_comp = Split(myPath, ".")(0)
       
       'Call SAP_EXTRACT_DATA_FBLN_REPORT
        Call PEGAR_EXCEL_FBLN
        Call ELIMINAR_REGISTROS_PDF
        
        Call DESARROLLO_PROCESO_PDF
        Call MOVIMIENTO_COLUMNAS_MASIVAS
        Call VALIDACION_DE_CONSTANCIA_AVANZADA
        Call CREACION_COLUMNAS_VALIDACION_FINAL
        Call ELIMINAR_DATOS_PROCESO
       
       
       
       CopiarHoja_guardar_enviar_por_correo file_name_comp
       
       
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
    End If
    
    'Clean up memory
    Set Mail_Single = Nothing
    Set Mail_Object = Nothing
End Sub

Function Exportar_valores_sistema(ByVal origen As String, ByVal destino As String)
    'Declarar variables
    Dim wb_origen As Workbook
    Dim wb_destino As Workbook
    Dim origen_hoja As Worksheet
    Dim destino_hoja As Worksheet
    Dim origen_rango_fecha_inicio As Range
    Dim origen_rango_fecha_fin As Range
    Dim destino_rango_fecha_inicio As Range
    Dim destino_rango_fecha_fin As Range
    Dim origen_nombre_carpeta_unidad As Range
    Dim destino_nombre_carpeta_unidad As Range
    Dim origen_unidad As Range
    Dim destino_unidad As Range
    Dim origen_unidades As Range
    Dim destino_unidades As Range
    Dim origen_ano As Range
    Dim destino_ano As Range
    Dim origen_meses As Range
    Dim destino_meses As Range


    'Definir el libro de origen y la hoja de origen
    Set wb_origen = Workbooks.Open(origen)
    Set origen_hoja = wb_origen.Sheets("SISTEMA")

    'Definir el rango de celdas de origen
    Set origen_rango_fecha_inicio = origen_hoja.Range("J8")
    Set origen_rango_fecha_fin = origen_hoja.Range("J9")
    Set origen_nombre_carpeta_unidad = origen_hoja.Range("H16")
    Set origen_unidad = origen_hoja.Range("H18")
    Set origen_unidades = origen_hoja.Range("H22")
    Set origen_ano = origen_hoja.Range("H20")
    Set origen_meses = origen_hoja.Range("H20")
    

    'Definir el libro de destino y la hoja de destino
    Set wb_destino = Workbooks.Open(destino)
    Set destino_hoja = wb_destino.Sheets("REPORTE_SAP")

    'Definir el rango de celdas de destino
    Set destino_rango_fecha_inicio = destino_hoja.Range("B2")
    Set destino_rango_fecha_fin = destino_hoja.Range("D2")
    Set destino_nombre_carpeta_unidad = destino_hoja.Range("B4")
    
    Set destino_unidad = destino_hoja.Range("B5")
    Set destino_unidades = destino_hoja.Range("B6")
    Set destino_ano = destino_hoja.Range("B7")
    Set destino_meses = destino_hoja.Range("B8")
    
    'Exportar el contenido transformado a la celda de destino
    destino_rango_fecha_inicio.Value = origen_rango_fecha_inicio.Value
    destino_rango_fecha_fin.Value = origen_rango_fecha_fin.Value
    destino_nombre_carpeta_unidad.Value = origen_nombre_carpeta_unidad.Value
    destino_unidad.Value = origen_unidad.Value
    destino_unidades.Value = origen_unidades.Value
    destino_meses.Value = origen_meses.Value
    
    

    'Guardar y cerrar el libro de origen y destino
    wb_origen.Save
    wb_origen.Close
    wb_destino.Save
    'wb_destino.Close

End Function
