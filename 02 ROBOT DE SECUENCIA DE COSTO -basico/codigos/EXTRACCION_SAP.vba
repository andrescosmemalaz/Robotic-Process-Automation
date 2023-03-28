Attribute VB_Name = "EXTRACCION_SAP"
Option Explicit

Dim FSO As Object

'Para versión de Excel de 32bit
'Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'Para versión de Excel de 64bit
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10

Private Const VK_SNAPSHOT = &H2C



Sub EXTRACCION_SAP()

'Codigo para extraer información de SAP

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Call ELIMINAR_TABLA_SAP
'On Error GoTo ErrorHandler

'DECLARACION DE OBJETOS, STRING, DE,AS
Dim Appl As Object
Dim Connection As Object
Dim session As Object
Dim WshShell As Object
Dim SapGui As Object
Dim SelectedTransaccion As Object
Dim SelectedFecha1 As String
Dim SelectedFecha2 As String
Dim SelectedFecha3 As String
Dim SelectedFecha4 As String
Dim SelectedSociedad As String
Dim Selectusuario As String
Dim Selectedpassword As String
Dim Selectedexepcion_one As String
Dim Selectedexepcion_two As String
Dim Selectedexcepcion_environment As String
Dim answer As String
Dim SAP As Object
Dim ws As Worksheet



'MsgBox ("Proceso " + nombre_proceso + " " + "completado con éxito.")
'Application.DisplayAlerts = False

' **** CONFIGURACIÓN DE RUTAS Y CAMPOS A UTILIZAR (IMP) ****

Dim ejecutar As String
Dim strFile, strFileName, strFolderName, strFolderExists As String
Dim contador As Integer
Dim contador_lista As Integer

'Declarar contadores
contador_lista = 1
contador = 1

'declaras rutas
Dim ruta_proyecto As String
Dim folder_proceso As String
Dim ruta_login As String
Dim libro_reporte, libro_sistema, nombre_reporte_sap As String
'Dim ws As Worksheet

'Sistema actual
libro_sistema = ActiveWorkbook.Name

Workbooks(libro_sistema).Activate
'Seleccionar una hoja especifica
Sheets("PROCESO").Select
'Limpiar Celdas
Cells.Clear
'Seleccionar Celda A 1
Range("A1").Select

Sheets("PROCESO").Select
'Definir sociedad, año, periodo_ini, periodo_fin
Dim Select_ano, Selected_unidad, Selected_mes As String

'Definir sociedad , año, periodo_ini, periodo_fin

'sociedad = Trim(Range("G18").Value)
'año = Trim(Range("G19").Value)
'periodo_ini = Trim(Range("F22").Value)
'periodo_fin = Trim(Range("G22").Value)

Select_ano = Trim(Worksheets("SISTEMA").Range("G10").Value)
Selected_unidad = Trim(Worksheets("SISTEMA").Range("G11").Value)
Selected_mes = Trim(Worksheets("SISTEMA").Range("G12").Value)


'ruta login y ruta proyecto
'ruta_login = "C:\Rpa\COE_TRI_Robot_Exter_PE\LOGIN\LOGIN.xlsx"
ruta_proyecto = "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\P&P\2. PROYECTOS DE PRACTICANTE\ROBOT DE SECUENCIA DE COSTO\"

'Declarar nombre del proceso
Dim nombre_proceso As String
nombre_proceso = "PL preliminar"

'folder del proceso
folder_proceso = "Proceso " + nombre_proceso + " (" + CStr(Selected_unidad) + ")" + " " + Format(Now(), "DD-MMM-YYYY") + "\"
' libro de reporte capturas
libro_reporte = "Capturas - " + nombre_proceso + ".xlsx"
'nombre de reporte sap
nombre_reporte_sap = "Reporte " + nombre_proceso + " " + CStr(Selected_unidad) + " " + Mid(Format(Now(), "DD-MMM-YYYY"), 4, Len(Format(Now(), "DD-MMM-YYYY"))) + ".XLSX"

'Seleccionar libro sistema por hoja borrador
Set ws = Workbooks(libro_sistema).Worksheets("PROCESO")

'La primera matriz se llama "lista_rangos" y se declara con la palabra clave "Dim", que indica que se está creando una nueva variable en el programa. La variable es una matriz de cadenas de texto, y se especifica que tiene 6 elementos (del índice 1 al índice 6) utilizando la sintaxis "1 To 6". Esto significa que la matriz puede contener hasta 6 cadenas de texto, que se pueden acceder utilizando los índices numéricos de la matriz (por ejemplo, "lista_rangos(1)", "lista_rangos(2)", etc.).

Dim lista_rangos(1 To 6) As String

'La segunda matriz se llama "lista_fotos_f01" y se declara de la misma manera que la primera. La única diferencia es el nombre de la matriz. También es una matriz de cadenas de texto con 6 elementos, y se puede acceder a ellos utilizando los índices numéricos de la matriz.
'Dim lista_fotos_f01(1 To 6) As String

'Estas matrices se utilizan para almacenar una lista de valores de cadenas de texto. Los valores se pueden asignar a cada elemento de la matriz utilizando la sintaxis "lista_rangos(1) = 'valor'" o "lista_fotos_f01(1) = 'valor'". Luego, los valores se pueden recuperar de la matriz en cualquier momento utilizando la sintaxis "valor = lista_rangos(1)" o "valor = lista_fotos_f01(1)".

'Por ejemplo, "lista_rangos(1) = 'A3:Q49'" asigna la cadena de texto "A3:Q49" al primer elemento de la matriz "lista_rangos". Esto significa que el rango de celdas A3:Q49 en la hoja de cálculo se puede hacer referencia a través de la variable "lista_rangos(1)" en el código.
'De manera similar, los valores asignados a los elementos de la matriz "lista_rangos"

lista_rangos(1) = "A1:H40"
lista_rangos(2) = "I3:Q40"
lista_rangos(3) = "Q3:AF42"
lista_rangos(4) = "AF3:AV42"
lista_rangos(5) = "AV3:BJ42"
lista_rangos(6) = "BK3:CA42"

strFolderName = ruta_proyecto + folder_proceso + "\"
strFolderExists = Dir(strFolderName, vbDirectory)

If strFolderExists = "" Then
    MkDir ruta_proyecto + folder_proceso
End If
' **** FIN DE CONFIGURACIÓN ****



'DEFINIR HOJA SISTEMA
Set ws = ThisWorkbook.Worksheets("SISTEMA")

'EXTRACCION DE VALOR DE ENVIRONMENT
'DEFINIR NOMBRE DEL REPORTE 1 Y 2
Selectedexcepcion_environment = Worksheets("CREDENCIALES SAP").Range("B3").Value
Dim nombre_reporte_1 As String
nombre_reporte_1 = "REPORTE-Y_e01_23000001"
Dim nombre_reporte_2 As String
nombre_reporte_2 = "REPORTE-Y_e01_23000002"


'carpeta directorio

'ABRIR SAP

Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", 4
Set WshShell = CreateObject("WScript.Shell")

'GENERAR BUCLE APERTURA  Y ESPERA DE UN SEGUNDO

Do Until WshShell.AppActivate("SAP Logon ")
    Application.Wait Now + TimeValue("0:00:01")
Loop

Set SapGui = GetObject("SAPGUI")
Set Appl = SapGui.GetScriptingEngine

'CONEXION SAP
Set Connection = Appl.OpenConnection(Selectedexcepcion_environment, _
    True)
Set session = Connection.Children.Item(0)

'SelectedFecha1 = Worksheets("REPORTE_SAP").Range("B2").Value
'SelectedFecha2 = Worksheets("REPORTE_SAP").Range("D2").Value

'SELECCIONAR USUARIO
'SELECCIONAR PASSWORD

Selectusuario = Worksheets("CREDENCIALES SAP").Range("B1").Value
Selectedpassword = Worksheets("CREDENCIALES SAP").Range("B2").Value

'INGRESAR SESSION NUMERO DE OBJECT , USUARIO , PASSWORD Y LANGUAGE
With session
    .findById("wnd[0]/usr/txtRSYST-MANDT").Text = 150
    .findById("wnd[0]/usr/txtRSYST-BNAME").Text = Selectusuario
    .findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Selectedpassword
    .findById("wnd[0]/usr/txtRSYST-LANGU").Text = "ES"
End With


'MAXIMIZA PANTALLA
session.findById("wnd[0]").maximize
'INGRESA A TRANSACCION
session.findById("wnd[0]/tbar[0]/okcd").Text = "Y_e01_23000001"
'DALE ENTER
session.findById("wnd[0]").sendVKey 0
'INGRESAR AÑO
session.findById("wnd[0]/usr/txt$0F-RY00").Text = Select_ano
'INGRESAR TRANSACCION
session.findById("wnd[0]/usr/ctxt$0FRBUKR").Text = Selected_unidad
'INGRESAR MES
session.findById("wnd[0]/usr/txt$Z_PER_F").Text = Selected_mes

': establece el enfoque en el cuadro de texto "0FBAGRP" en la ventana SAP actual.
session.findById("wnd[0]/usr/ctxt$0FBAGRP").SetFocus

'Establece la posición del cursor en el campo de texto "ctxt$0FBAGRP" a la posición 0 (es decir, al principio del campo).
session.findById("wnd[0]/usr/ctxt$0FBAGRP").caretPosition = 0
'Hace clic en el botón "Enter" de la barra de herramientas de SAP (tbar[1]) para enviar la información ingresada en el campo de texto anterior.
session.findById("wnd[0]/tbar[1]/btn[8]").press
'Selecciona el nodo "000002" en el árbol de navegación de SAP, ubicado en la ventana activa
session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell").selectedNode = "000002"
'Establece la posición de la barra de desplazamiento vertical de la ventana activa a 11.
session.findById("wnd[0]/usr").verticalScrollbar.Position = 11
'Establece la posición de la barra de desplazamiento vertical de la ventana activa a 15.
session.findById("wnd[0]/usr").verticalScrollbar.Position = 15
'Establece la posición de la barra de desplazamiento vertical de la ventana activa a 22
session.findById("wnd[0]/usr").verticalScrollbar.Position = 22
'Establece la posición de la barra de desplazamiento vertical de la ventana activa a 25
session.findById("wnd[0]/usr").verticalScrollbar.Position = 25
'Establece el foco en la etiqueta "lbl[97,11]" de la ventana activa..
session.findById("wnd[0]/usr/lbl[97,11]").SetFocus

' Establece la posición del cursor en la etiqueta "lbl[97,11]" de la ventana activa a la posición 12 (es decir, después de los primeros 12 caracteres de la etiqueta).
session.findById("wnd[0]/usr/lbl[97,11]").caretPosition = 12

' ============= PROCESO DE CAPTURA DE PANTALLAS =============

'**** Configuración SAP previa a Screenshot #01 ****

contador = 1
'**** Z = session.findById("wnd[0]").Text
'**** AppActivate Z
'**** Application.Wait (Now + TimeValue("0:00:03"))
'**** session.findById("wnd[0]/tbar[0]/okcd").Text = "ZVMFI060"
'**** session.findById("wnd[0]").sendVKey 0
'**** Application.Wait (Now + TimeValue("0:00:02"))

'========== SCREENSHOT #01 ==========


Dim Z As String
'Ficha Técnica
Z = session.findById("wnd[0]").Text
SetCursorPos 1663, 1013 'x and y position
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
Application.Wait (Now + TimeValue("0:00:03"))
'Fin Ficha Técnica

Z = session.findById("wnd[0]").Text
Application.Wait (Now + TimeValue("0:00:03"))
Application.SendKeys "({1068})", True
Workbooks(libro_sistema).Activate
Sheets("PROCESO").Select
Application.Wait (Now + TimeValue("0:00:02"))
ActiveSheet.Paste
ejecutar = Guardar_Imagen(libro_sistema, "PROCESO", contador, ruta_proyecto + folder_proceso)
contador = contador + 1

'========== FIN SCREENSHOT #01 ==========


'Abre el menú "Gestión de inventario" de la barra de menú de SAP, y selecciona el submenú "SISTEMA.
session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").Select
'Selecciona el botón de opción "Individual" en la ventana de selección dE SISTEMA .
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
'Establece el foco en el botón de opción "Individual" en la ventana de selección de lista de materiales.
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
'Hace clic en el botón "Aceptar" de la barra de herramientas de la ventana activa para confirmar la selección realizada.
session.findById("wnd[1]/tbar[0]/btn[0]").press

'Establece la ruta de acceso de destino para guardar el archivo de informe en el campo "DY_PATH" de la ventana secundaria.
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ruta_proyecto + folder_proceso
'Establece el nombre del archivo de informe que se guardará en el campo "DY_FILENAME" de la ventana secundaria. El nombre del archivo se compone del valor de la variable "nombre_reporte_1" concatenado con la fecha actual en formato "yyyy-mm-dd" y con la extensión ".xls".
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = nombre_reporte_1 & "-" & Format(Now(), "yyyy-mm-dd") & ".xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[1]/usr/btnBUTTON_YES").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/usr/btnSTARTBUTTON").press
session.findById("wnd[0]/tbar[0]/okcd").Text = "Y_e01_23000002"

session.findById("wnd[0]").sendVKey 0

'Ingresa el código de transacción "Y_e01_23000002" en la barra de comando de la ventana principal.
'session.findById("wnd[0]/tbar[0]/okcd").Text = "Y_e01_23000002"
'Presiona la tecla "Enter" para ejecutar la transacción ingresada en la barra de comando.
'session.findById("wnd[0]").sendVKey 0

' establece el valor "2022" en el cuadro de texto "txt$0F-RY00".
session.findById("wnd[0]/usr/txt$0F-RY00").Text = Select_ano
'ESTABLECE EL VALOR DE UNIDAD
session.findById("wnd[0]/usr/ctxt$0FRBUKR").Text = Selected_unidad
' busca el campo de texto en la ventana activa de SAP con el identificador "wnd[0]/usr/txt$Z_PER_F" y establece su valor en "1".
session.findById("wnd[0]/usr/txt$Z_PER_F").Text = Selected_mes
'establece el foco en el campo de texto "wnd[0]/usr/txt$Z_PER_F".
session.findById("wnd[0]/usr/txt$Z_PER_F").SetFocus
'establece la posición del cursor en el segundo carácter del campo de texto "wnd[0]/usr/txt$Z_PER_F".
session.findById("wnd[0]/usr/txt$Z_PER_F").caretPosition = 2
'presiona el botón con el identificador "wnd[0]/tbar[1]/btn[8]" en la barra de herramientas de la ventana activa de SAP.
session.findById("wnd[0]/tbar[1]/btn[8]").press
'establece el nodo seleccionado en la ventana activa de SAP con el identificador "wnd[0]/shellcont/shell/shellcont[0]/shell" en "000002".


session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell").selectedNode = "000002"
session.findById("wnd[0]/shellcont/shell/shellcont[0]/shell").selectedNode = "000001"
session.findById("wnd[0]/usr/lbl[0,17]").SetFocus
session.findById("wnd[0]/usr/lbl[0,17]").caretPosition = 0
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr").verticalScrollbar.Position = 2
session.findById("wnd[0]/usr").verticalScrollbar.Position = 3
session.findById("wnd[0]/usr").verticalScrollbar.Position = 4
session.findById("wnd[0]/usr").verticalScrollbar.Position = 5
session.findById("wnd[0]/usr/lbl[0,31]").SetFocus
session.findById("wnd[0]/usr/lbl[0,31]").caretPosition = 0
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr").verticalScrollbar.Position = 6
session.findById("wnd[0]/usr").verticalScrollbar.Position = 7
session.findById("wnd[0]/usr").verticalScrollbar.Position = 8
session.findById("wnd[0]/usr").verticalScrollbar.Position = 9
session.findById("wnd[0]/usr").verticalScrollbar.Position = 8
session.findById("wnd[0]/usr").horizontalScrollbar.Position = 22
session.findById("wnd[0]/usr").horizontalScrollbar.Position = 35

' ============= PROCESO DE CAPTURA DE PANTALLAS =============

'**** Configuración SAP previa a Screenshot #02 ****

'contador = 1


'========== SCREENSHOT #02 ==========


'Dim Z As String
'Ficha Técnica
Z = session.findById("wnd[0]").Text
SetCursorPos 1663, 1013 'x and y position
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
Application.Wait (Now + TimeValue("0:00:03"))
'Fin Ficha Técnica

Z = session.findById("wnd[0]").Text
Application.Wait (Now + TimeValue("0:00:03"))
Application.SendKeys "({1068})", True
Workbooks(libro_sistema).Activate
Sheets("PROCESO").Select
Application.Wait (Now + TimeValue("0:00:02"))
ActiveSheet.Paste
ejecutar = Guardar_Imagen(libro_sistema, "PROCESO", contador, ruta_proyecto + folder_proceso)
contador = contador + 1

'========== FIN SCREENSHOT #02 ==========






session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").Select

'La primera línea selecciona un botón de opción (radio button) en una pantalla de selección de opciones en SAP.
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
'La segunda línea establece el foco en el botón de opción seleccionado.
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
'La tercera línea presiona un botón en la barra de herramientas.
session.findById("wnd[1]/tbar[0]/btn[0]").press
'La cuarta línea escribe una ruta de archivo en el campo de texto correspondiente en una ventana de diálogo.
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = ruta_proyecto + folder_proceso
'La quinta línea escribe un nombre de archivo en el campo de texto correspondiente en la misma ventana de diálogo.
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = nombre_reporte_2 & "-" & Format(Now(), "yyyy-mm-dd") & ".xls"
'La sexta línea establece el cursor en la posición 18 en el campo de texto de nombre de archivo.
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
'La séptima línea presiona un botón en la barra de herramientas.
session.findById("wnd[1]/tbar[0]/btn[0]").press
'La octava línea presiona un botón en una ventana de diálogo.
session.findById("wnd[0]/tbar[0]/btn[3]").press
'La novena línea presiona un botón en la barra de herramientas.
session.findById("wnd[1]/usr/btnBUTTON_YES").press
'La novena línea presiona un botón en la barra de herramientas.
session.findById("wnd[0]/tbar[0]/btn[3]").press




Dim wb As Workbook
Set wb = Workbooks.Open("C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\P&P\2. PROYECTOS DE PRACTICANTE\ROBOT DE SECUENCIA DE COSTO\PROYECTOS DE SECUENCIA DE COSTOS.xlsm") ' Cambiar la ruta y el nombre del archivo existente

With wb
    .Worksheets("PL preliminar").Activate ' Cambiar el nombre de la hoja existente
End With

Dim lista_fotos_f01(2) As String ' Cambiar el tamaño del array según la cantidad de imágenes
'Dim contador_lista As Integer
contador_lista = 0

'Dim strFile As String
strFile = Dir(ruta_proyecto + folder_proceso + "\*.png") ' Cambiar la ruta de la carpeta con las imágenes

Do While strFile <> ""
    lista_fotos_f01(contador_lista) = strFile
    contador_lista = contador_lista + 1
    strFile = Dir()
Loop

contador_lista = 0

Do While contador_lista < UBound(lista_fotos_f01)
    ejecutar = InsertarImagenEnRango(ruta_proyecto + folder_proceso & lista_fotos_f01(contador_lista), Range(lista_rangos(contador_lista + 1)))
    ' Cambiar la ubicación de la imagen
    contador_lista = contador_lista + 1
Loop











'Call PEGAR_EXCEL
'MsgBox "Proceso realizado con éxito"

On Error Resume Next

'session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'While Connection.Children.Count > 0
'    Set session = Connection.Children(0)
'    session.findById("wnd[0]").Close
'    On Error Resume Next
'    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'    On Error GoTo 0
'Wend

Call CERRAR_PROCESO_SAP
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

'Exit Sub
'ErrorHandler:
'MsgBox "No se ha conectado a la VPN o la contraseña debe ser cambiada"

End Sub

Function InsertarImagenEnRango(PictureFileName As String, TargetCells As Range)

' inserts a picture and resizes it to fit the TargetCells range
Dim p As Object, t As Double, l As Double, w As Double, h As Double

If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function
If Dir(PictureFileName) = "" Then Exit Function

' import picture
'Set p = ActiveSheet.Pictures.Insert(PictureFileName), linktofile:=msoFalse, savewithdocument:=msoCTrue

Set p = ActiveSheet.Shapes.AddPicture(Filename:=PictureFileName, _
LinkToFile:=False, SaveWithDocument:=True, Left:=l, Top:=t, Width:=w, Height:=h)
p.Select


' determine positions
With TargetCells
t = .Top
l = .Left
w = .Offset(0, .Columns.Count).Left - .Left
h = .Offset(.Rows.Count, 0).Top - .Top
End With

' position picture

With p
.Top = t
.Left = l
.Width = w
.Height = h

Dim Pic As Shape
For Each Pic In ActiveSheet.Shapes
    Pic.Select
    Pic.LockAspectRatio = msoTrue
Next Pic
    
End With

Set p = Nothing

End Function

Function Guardar_Imagen(Book_Name, Sheet_Name As String, indice As Integer, Folder_Name As String)

Dim cht As ChartObject
Dim shp As Shape

Workbooks(Book_Name).Activate
Sheets(Sheet_Name).Select

For Each shp In ActiveSheet.Shapes

    'MsgBox (shp.Name)
    shp.Select
    'Create a temporary chart object (same size as shape)
    Set cht = ActiveSheet.ChartObjects.Add( _
    Left:=shp.Left, _
    Width:=shp.Width, _
    Top:=shp.Top, _
    Height:=shp.Height)

    'Format temporary chart to have a transparent background
    cht.ShapeRange.Fill.Visible = msoFalse
    cht.ShapeRange.Line.Visible = msoFalse
    
    'Copy/Paste Shape inside temporary chart
    shp.Copy
    cht.Activate
    ActiveChart.Paste
  
    'Save chart to User's Desktop as PNG File
    cht.Chart.Export Folder_Name & "Imagen" + " " + CStr(indice) & ".png"

    'Delete temporary Chart
    cht.Delete

    'Re-Select Shape (appears like nothing happened!)
    shp.Select
    shp.Delete
    Exit Function
Next

End Function




Sub CERRAR_PROCESO_SAP()
'
' Declarar una variable para almacenar el proceso de SAP Logon
Dim sapLogonProcess As Object
Dim p As Object


' Obtener el proceso de SAP Logon
Set sapLogonProcess = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("select * from Win32_Process where name='saplogon.exe'")

' Cerrar la aplicación de SAP Logon
For Each p In sapLogonProcess
    p.Terminate
Next

End Sub

Sub EliminarEspacios()
Application.DisplayAlerts = False

    Dim archivo1 As Workbook, archivo2 As Workbook
    Dim celda As Range
    
    'Abrir archivo 1
    Set archivo1 = Workbooks.Open("C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\P&P\2. PROYECTOS DE PRACTICANTE\ROBOT DE SECUENCIA DE COSTO\REPORTES SAP Y_e01\REPORTE-Y_e01_23000001-2023-02-27.xls") 'Cambiar "ruta\archivo.xlsx" por la ruta y nombre del archivo que desees abrir
    
    'Iterar sobre columna en archivo 1
    For Each celda In archivo1.Sheets(1).Range("B1:B200") 'Cambiar "A1:A200" por el rango de la columna que desees iterar
        celda.Value = Trim(celda.Value) 'Eliminar espacios al principio y al final de la celda
        celda.Value = Replace(celda.Value, " ", "") 'Eliminar espacios en medio de la celda
    Next celda
    
    'Guardar cambios en archivo 1 y cerrarlo
    archivo1.Save
    archivo1.Close
    
    'Abrir archivo 2
    Set archivo2 = Workbooks.Open("C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\P&P\2. PROYECTOS DE PRACTICANTE\ROBOT DE SECUENCIA DE COSTO\REPORTES SAP Y_e02\REPORTE-Y_e01_23000002-2023-02-27.xls") 'Cambiar "ruta\archivo2.xlsx" por la ruta y nombre del archivo 2 que desees abrir
    
    'Iterar sobre columna en archivo 2
    For Each celda In archivo2.Sheets(1).Range("B1:B200") 'Cambiar "A1:A200" por el rango de la columna que desees iterar
        celda.Value = Trim(celda.Value) 'Eliminar espacios al principio y al final de la celda
        celda.Value = Replace(celda.Value, " ", "") 'Eliminar espacios en medio de la celda
    Next celda
    
    'Guardar cambios en archivo 2 y cerrarlo
    archivo2.Save
    archivo2.Close
    
End Sub
'Este código abre dos archivos de Excel, elimina los espacios en la columna "A" desde la celda 1 hasta la 200 de ambos archivos, guarda los cambios y cierra los archivos. Puedes modificar la ruta y el nombre de los archivos y el rango de la columna según tus necesidades.


Sub BUSQUEDA_VALORES()

Application.DisplayAlerts = False

    Dim archivoOrigen As Workbook
    Dim archivoDestino As Workbook
    Dim archivoOrigen2 As Workbook
    Dim archivoDestino2 As Workbook
    
    Dim valor1 As String
    Dim valor2 As String
    Dim valor3 As String
    Dim valor4 As String
    Dim rangoCopia As Range
    Dim rangoCopia2 As Range
    
    
    'Abrir archivo origen y asignarlo a la variable "archivoOrigen"
    Set archivoOrigen = Workbooks.Open("C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\P&P\2. PROYECTOS DE PRACTICANTE\ROBOT DE SECUENCIA DE COSTO\REPORTES SAP Y_e01\REPORTE-Y_e01_23000001-2023-02-27.xls")
    
    'Buscar primera palabra y guardar valor de la celda
    valor1 = archivoOrigen.Sheets(1).Cells.Find("COSTO,GASTOS").Row + 1
    'Buscar segunda palabra y guardar valor de la celda
    valor2 = archivoOrigen.Sheets(1).Cells.Find("INVERSIONES").Row - 1
    
    

    'Unir los dos valores en un rango de 3 columnas
    Set rangoCopia = archivoOrigen.Sheets(1).Range("A" + valor1 + ":" + "B" + valor2)
    rangoCopia.Select
    rangoCopia.Copy
    'Abrir archivo destino y pegar los valores en la hoja activa
    Set archivoDestino = Workbooks.Open("C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\P&P\2. PROYECTOS DE PRACTICANTE\ROBOT DE SECUENCIA DE COSTO\PROYECTOS DE SECUENCIA DE COSTOS.xlsm")
    archivoDestino.Activate
    archivoDestino.Sheets("PROCESO").Activate
    Range("A1").PasteSpecial xlPasteValues
    
    
    archivoOrigen.Close
    
    Set archivoOrigen2 = Workbooks.Open("C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\P&P\2. PROYECTOS DE PRACTICANTE\ROBOT DE SECUENCIA DE COSTO\REPORTES SAP Y_e02\REPORTE-Y_e01_23000002-2023-02-27.xls")
    
    'Buscar primera palabra y guardar valor de la celda
    valor3 = archivoOrigen2.Sheets(1).Cells.Find("COSTO,GASTOS").Row + 1
    'Buscar segunda palabra y guardar valor de la celda
    valor4 = archivoOrigen2.Sheets(1).Cells.Find("INVERSIONES").Row - 1
    
    Set rangoCopia2 = archivoOrigen2.Sheets(1).Range("A" + valor3 + ":" + "B" + valor4)
    rangoCopia2.Select
    rangoCopia2.Copy
    'Abrir archivo destino y pegar los valores en la hoja activa
    Set archivoDestino = Workbooks.Open("C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\P&P\2. PROYECTOS DE PRACTICANTE\ROBOT DE SECUENCIA DE COSTO\PROYECTOS DE SECUENCIA DE COSTOS.xlsm")
    archivoDestino.Activate
    archivoDestino.Sheets("PROCESO").Activate
    Range("D1").PasteSpecial xlPasteValues
    
    
    'Cerrar archivos
    'archivoOrigen.Close
    archivoOrigen2.Close
    'archivoDestino.Close SaveChanges:=True
    
End Sub

Sub copiarDatos()
    Dim wsProceso As Worksheet
    Dim wsPLPreliminar As Worksheet
    Dim tblModelado As ListObject
    Dim i As Long
    
    ' Definir las hojas de trabajo
    Set wsProceso = ThisWorkbook.Worksheets("Proceso")
    Set wsPLPreliminar = ThisWorkbook.Worksheets("PL Preliminar")
    
    ' Definir la tabla de destino
    Set tblModelado = wsPLPreliminar.ListObjects("MODELADO_PL_PRELIMINAR")
    Set rango = wsProceso.Range("A1:A10")
    
    ' Copiar los datos celda por celda a la tabla
    With tblModelado.DataBodyRange
        .Cells(i, 1).Value = wsProceso.Cells(i, 1).Value
        .Cells(i, 2).Value = wsProceso.Cells(i, 2).Value
        .Cells(i, 3).Value = wsProceso.Cells(i, 5).Value
        .Cells(i, 4).Value = wsProceso.Cells(i, 6).Value
    End With

End Sub


Sub copiarDatos_v2()
    Dim wsProceso As Worksheet
    Dim wsPLPreliminar As Worksheet
    Dim tblModelado As ListObject
    Dim i As Long
    
    ' Definir las hojas de trabajo
    Set wsProceso = ThisWorkbook.Worksheets("Proceso")
    Set wsPLPreliminar = ThisWorkbook.Worksheets("PL Preliminar")
    
    ' Definir la tabla de destino
    Set tblModelado = wsPLPreliminar.ListObjects("MODELADO_PL_PRELIMINAR")
    Set rango = wsProceso.Range("A1:A10")
    
    ' Copiar los datos celda por celda a la tabla
    For i = 1 To rango.Rows.Count
        If Trim(wsProceso.Cells(i, 1).Value) <> "" Then
            tblModelado.DataBodyRange.Cells(i, 1).Value = Trim(wsProceso.Cells(i, 1).Value)
        Else
            tblModelado.DataBodyRange.Cells(i, 1).Value = wsProceso.Cells(i, 1).Value
        End If
        
        If Trim(wsProceso.Cells(i, 2).Value) <> "" Then
            tblModelado.DataBodyRange.Cells(i, 2).Value = Trim(wsProceso.Cells(i, 2).Value)
        Else
            tblModelado.DataBodyRange.Cells(i, 2).Value = wsProceso.Cells(i, 2).Value
        End If
        
        If Trim(wsProceso.Cells(i, 5).Value) <> "" Then
            tblModelado.DataBodyRange.Cells(i, 3).Value = Trim(wsProceso.Cells(i, 5).Value)
        Else
            tblModelado.DataBodyRange.Cells(i, 3).Value = wsProceso.Cells(i, 5).Value
        End If
        
        ' Agregar más líneas si hay más columnas en la tabla
    Next i
End Sub


Sub PEGAR_EXCEL_FBLN()
    Application.EnableCancelKey = xlDisabled

    
    Dim wbI As Workbook, wbO As Workbook
    Set wbI = ThisWorkbook
    Dim wsI As Worksheet
    Set wsI = wbI.Sheets("PROCESO")
    'Dim wbO As Workbook
    Dim ruta As String
    
    ruta = "C:\Users\est.andresc1\OneDrive - Votorantim\Área de Trabalho\Project\P&P\2. PROYECTOS DE PRACTICANTE\ROBOT DE SECUENCIA DE COSTO\REPORTES SAP Y_e01\" & nombre_reporte_1 & "-" & Format(Now(), "yyyy-mm-dd") & ".xls"
    Set wbO = Workbooks.Open(ruta)
    
    
    Dim ws As Worksheet
    
    Dim C As Range
    Dim columnas As String
    Dim hoja As String
    Dim letra_columna As String
    Dim nombre_libro As String
    
    Dim nombre_reporte_1 As String
    nombre_reporte_1 = "REPORTE-Y_e01_23000001"
    Dim nombre_reporte_2 As String
    nombre_reporte_2 = "REPORTE-Y_e01_23000002"
    
    'Call ELIMINAR_TABLA_SAP
    
    wbO.Sheets(1).Cells.Copy wsI.Cells
    wbO.Close SaveChanges:=False
    wbI.Sheets("PROCESO").Activate
    wbI.Sheets("PROCESO").Range("A10").Activate
    
    Dim tbl As Range
    Set tbl = wsI.Range("A10").CurrentRegion
    
    wsI.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "DATA_SAP_FBLN"
    wsI.ListObjects("DATA_SAP_FBLN").Range.Cut _
        Destination:=Worksheets("REPORTE_SAP").Range("A10")
    wbI.Sheets("REPORTE_SAP").Activate
    
End Sub



Sub ELIMINAR_TABLA_SAP()

On Error Resume Next
Application.EnableCancelKey = xlDisabled

    Dim NOMBRE As String
    Dim columnas As String
    Dim sSheetName As String
    Dim sTableName As String


    Dim ws As Worksheet
    

    On Error Resume Next
    Dim C As Range
    Sheets("SAP").Select
    columnas = "A10:Z10"
     'Define Variables
    sSheetName = "REPORTE_SAP"
    sTableName = "DATA_SAP_FBLN"
    'Delete Table
     Set ws = Sheets(sSheetName)
     ws.ListObjects(sTableName).Delete

'Application.Wait (Now + TimeValue("0:00:05"))
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub


