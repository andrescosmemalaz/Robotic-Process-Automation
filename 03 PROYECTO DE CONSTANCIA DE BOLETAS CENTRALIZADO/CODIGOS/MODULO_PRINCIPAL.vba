Attribute VB_Name = "MODULO_PRINCIPAL"
Sub AUTOMATIZACION_CONSTANCIAS()

Call EXTRACCION_SAP
Call PEGAR_EXCEL_FBLN
Call ELIMINAR_REGISTROS_PDF
Call COPIAR_PDF_DESDE_RUTA_DHO_2
Call LIMPIAR_RANGO_COLUMNAS
Call DESARROLLO_PROCESO_PDF
Call MOVIMIENTO_COLUMNAS_MASIVAS
Call VALIDACION_DE_CONSTANCIA_AVANZADA
Call CREACION_COLUMNAS_VALIDACION_FINAL
Call ELIMINAR_DATOS_PROCESO_VALIDACION_AVANZADA
Call PINTAR_VALORES_VALIDACION

MOVER_PDFS_AL_TERMINAR

MsgBox ("PROCESO FINALIZADO")
End Sub

Sub AUTOMATIZAR_DESCARGA_REPORTE()
Call ELIMINAR_TABLA_SAP
'Call EXTRACCION_SAP
Call PEGAR_EXCEL_FBLN
End Sub
Sub AUTOMATIZAR_LECTURA_PDF()
Call ELIMINAR_REGISTROS_PDF
Call DESARROLLO_PROCESO_PDF
Call MOVIMIENTO_COLUMNAS_MASIVAS
End Sub
Sub VALIDACION()
Call VALIDACION_DE_CONSTANCIA_AVANZADA
Call CREACION_COLUMNAS_VALIDACION_FINAL
MsgBox ("PROCESO FINALIZADO")
End Sub

Sub VALIDACION_FINAL_CONSTANCIAS()
Call AUTOMATIZAR_DESCARGA_REPORTE
Call AUTOMATIZAR_LECTURA_PDF
Call VALIDACION
End Sub