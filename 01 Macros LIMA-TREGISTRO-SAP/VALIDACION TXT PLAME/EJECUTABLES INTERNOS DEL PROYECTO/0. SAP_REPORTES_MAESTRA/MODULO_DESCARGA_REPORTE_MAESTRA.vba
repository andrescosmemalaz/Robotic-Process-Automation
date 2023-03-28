Attribute VB_Name = "MODULO_DESCARGA_REPORTE_MAESTRA"
Option Explicit
Public SapGuiAuto
Public objGui As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

Sub SAP_extract_DataMaestra_Reporte()

'Codigo para extraer información de SAP
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Call ELIMINAR_TABLA_SAP_FINAL
On Error GoTo ErrorHandler

Dim Appl As Object
Dim connection As Object
Dim session As Object
Dim WshShell As Object
Dim SapGui As Object
Dim SelectedTransaccion As String
Dim SelectedFecha1 As String
Dim SelectedFecha2 As String
Dim SelectedFecha3 As String
Dim SelectedFecha4 As String
Dim SelectedSociedad As String
Dim Selectusuario As String
Dim Selectedpassword As String
Dim Selectedexepcion_one As String
Dim Selectedexepcion_two As String
Dim environment As String


Dim answer As String
Dim SAP As Object

Sheets("CREDENCIALES SAP").Activate
environment = Worksheets("CREDENCIALES SAP").Range("B3").Value

Sheets("SAP").Activate
'carpeta directorio
Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", 4
Set WshShell = CreateObject("WScript.Shell")

Do Until WshShell.AppActivate("SAP Logon ")
    Application.Wait Now + TimeValue("0:00:01")
Loop

Set SapGui = GetObject("SAPGUI")
Set Appl = SapGui.GetScriptingEngine

Set connection = Appl.OpenConnection(environment, _
    True)
Set session = connection.Children.Item(0)

'declaracion de variables por operacion

SelectedTransaccion = ActiveWorkbook.ActiveSheet.Range("B1").Value
SelectedFecha1 = ActiveWorkbook.ActiveSheet.Range("B2").Value
SelectedFecha2 = ActiveWorkbook.ActiveSheet.Range("D2").Value
SelectedFecha3 = ActiveWorkbook.ActiveSheet.Range("B3").Value
SelectedFecha4 = ActiveWorkbook.ActiveSheet.Range("D3").Value
SelectedSociedad = ActiveWorkbook.ActiveSheet.Range("B5").Value
Selectedexepcion_one = ActiveWorkbook.ActiveSheet.Range("B6").Value
Selectedexepcion_two = ActiveWorkbook.ActiveSheet.Range("D6").Value


Sheets("CREDENCIALES SAP").Activate
Selectusuario = ActiveWorkbook.ActiveSheet.Range("B1").Value
Selectedpassword = ActiveWorkbook.ActiveSheet.Range("B2").Value
'usuario y contraseña

session.FindById("wnd[0]/usr/txtRSYST-MANDT").Text = 150
session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = Selectusuario
session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = Selectedpassword
session.FindById("wnd[0]/usr/txtRSYST-LANGU").Text = "ES"

Sheets("SAP").Activate
'Generacion de reporte en SAP


session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "ZVMHR047"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/ctxtPNPBEGDA").Text = SelectedFecha1
session.FindById("wnd[0]/usr/ctxtPNPENDDA").Text = SelectedFecha2
session.FindById("wnd[0]/usr/ctxtPNPBEGPS").Text = SelectedFecha3
session.FindById("wnd[0]/usr/ctxtPNPENDPS").Text = SelectedFecha4
session.FindById("wnd[0]/usr/ctxtPNPBUKRS-LOW").Text = SelectedSociedad
session.FindById("wnd[0]/usr/ctxtPNPSTAT2-LOW").SetFocus
session.FindById("wnd[0]/usr/ctxtPNPSTAT2-LOW").CaretPosition = 0
session.FindById("wnd[0]/usr/btn%_PNPSTAT2_%_APP_%-VALU_PUSH").Press
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "0"
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").CaretPosition = 1
session.FindById("wnd[1]/tbar[0]/btn[8]").Press
session.FindById("wnd[0]/usr/ctxtPNPABKRS-LOW").SetFocus
session.FindById("wnd[0]/usr/ctxtPNPABKRS-LOW").CaretPosition = 0
session.FindById("wnd[0]/usr/btn%_PNPABKRS_%_APP_%-VALU_PUSH").Press
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
session.FindById("wnd[1]").SendVKey 4
session.FindById("wnd[2]").Close
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = Selectedexepcion_one
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").Text = Selectedexepcion_two
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").SetFocus
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").CaretPosition = 2
session.FindById("wnd[1]/tbar[0]/btn[8]").Press
session.FindById("wnd[0]/tbar[1]/btn[8]").Press
session.FindById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\REPORTES"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORTABLE.XLSX"
session.FindById("wnd[1]/usr/ctxtDY_PATH").SetFocus
session.FindById("wnd[1]/usr/ctxtDY_PATH").CaretPosition = 111
session.FindById("wnd[1]/tbar[0]/btn[11]").Press
'session.FindById("wnd[0]/tbar[0]/btn[3]").Press
'session.FindById("wnd[0]/tbar[0]/btn[3]").Press
'session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[12]").Press
session.FindById("wnd[0]/tbar[0]/btn[12]").Press
session.FindById("wnd[0]/tbar[0]/btn[12]").Press

Call PEGAR_EXCEL
Call CERRAR_SAP_LOGON

 
'Cerrando sesión en SAP

'On Error Resume Next
'session.FindById("wnd[1]/usr/btnSPOP-OPTION1").Press



'While connection.Children.Count > 0
 '   Set session = connection.Children(0)
  '  session.FindById("wnd[0]").Close
  '  On Error Resume Next
   ' session.FindById("wnd[1]/usr/btnSPOP-OPTION1").Press
   ' On Error GoTo 0
'Wend


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub

ErrorHandler:
MsgBox "No se ha conectado a la VPN o la contraseña debe ser cambiada"

End Sub
Sub CERRAR_SAP_LOGON()

    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "TASKKILL /F /IM saplogon.exe", , True

End Sub


Sub aumentar_D4()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Sheets("FILTROS").Select
Range("M17").Value = Range("M17").Value + 1
Range("N17").Value = Range("N17").Value + 1


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub cerrar_archivo_export_maestra()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
    Workbooks.Open "C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\REPORTES\EXPORTABLE.XLSX"
    Workbooks("EXPORTABLE.XLSX").Worksheets("Sheet1").Activate
    'Workbooks("MC_IT0008.XLS").Close SaveChanges:=False
    Workbooks("EXPORTABLE.XLSX").Save
    Workbooks("EXPORTABLE.XLSX").Close
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub
Sub PEGAR_EXCEL()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Codigo para aumentar digito de extraccion de reporte
   'Call cerrar_archivo_export_maestra
   Call ELIMINAR_TABLA_SAP_FINAL
    
   
   Call aumentar_D4
   Sheets("FILTROS").Select
   Dim nombre_hoja As String
   nombre_hoja = Trim(Range("L17").Value)
   Dim numero_hoja As String
   numero_hoja = Trim(Range("M17").Value)
   Dim numero_iteracion As String
   numero_iteracion = Trim(Range("N17").Value)
   Sheets("SAP").Select
   
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    'Extraccion de tabla para determinar las ubicaciones de tableta
    ActiveWorkbook.Queries.Add Name:="DATA SAP" + nombre_hoja + "-" + CStr(numero_hoja) + CStr(numero_iteracion), Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Origen = Excel.Workbook(File.Contents(""C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\REPORTES\EXPORTABLE.XLSX""), null, true)," & Chr(13) & "" & Chr(10) & "    Sheet1_Sheet = Origen{[Item=""Sheet1"",Kind=""Sheet""]}[Data]," & Chr(13) & "" & Chr(10) & "    #""Encabezados promovidos"" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    #""Tip" & _
        "o cambiado"" = Table.TransformColumnTypes(#""Encabezados promovidos"",{{""Codigo"", type text}, {""Tipo de documento"", Int64.Type}, {""Número de documento"", Int64.Type}, {""Nombre"", type text}, {""Apellido paterno"", type text}, {""Apellido materno"", type text}, {""Descripción tipo"", type text}, {""Empresa"", type text}, {""División"", type text}, {""Area"", t" & _
        "ype text}, {""Departamento"", type text}, {""Cargo"", Int64.Type}, {""Nombre cargo"", type text}, {""F. Nacimiento"", type date}, {""F. Ingreso"", type date}, {""F. Cese"", type date}, {""Motivo cese"", type text}, {""Relación laboral"", type text}, {""Tipo de trabajador"", type text}, {""Jornada trab."", type text}, {""Tipo contrato"", type text}, {""Inicio contrat" & _
        "o"", type date}, {""Fin contrato"", type any}, {""Sindicato"", type text}, {""Locación"", type text}, {""EPS SITUACIÓN"", type text}, {""TIPO DE PLAN"", type text}, {""S.C.T.R. SALUD"", type text}, {""S.C.T.R. PENSION"", type text}, {""VIDA LEY"", type text}, {""ACCIDENTES PERSONALES"", type text}, {""SISTEMA DE PENSIONES"", type text}, {""CUSSP"", type text}, {""TI" & _
        "PO DE COMISIÓN AFP (MIXTA O FLUJO) Y %"", type text}, {""BANCO SUELDO"", type text}, {""NRO CUENTA SUELDO"", type number}, {""MONEDA CTA SUELDO"", type text}, {""TIPO CUENTA SUELDO"", type text}, {""NRO CUENTA CTS"", Int64.Type}, {""NRO CUENTA CTS (CCI)"", type text}, {""BANCO CTS"", type text}, {""TIPO MONEDA CTS"", type text}, {""SEXO"", type text}, {""ESTADO CIVI" & _
        "L"", type text}, {""UBIGEO"", Int64.Type}, {""LUGAR DE ORIGEN"", type text}, {""FUNCION QUE DESEMPEÑA"", type text}, {""DIRECCIÓN"", type text}, {""DISTRITO"", type text}, {""PROVINCIA"", type text}, {""DEPARTAMENTO"", type text}, {""TELEFONO"", type text}, {""EDAD"", Int64.Type}, {""PROFESION"", type text}, {""GRUPO SANGUINEO"", type text}, {""GRADO DE INSTRUCCION" & _
        """, type text}, {""LUG.NACIMIENTO"", type text}, {""C.COSTO"", type text}, {""DISCAPACIDAD"", type text}, {""NRO DE HIJOS"", Int64.Type}, {""GRADO SALARIAL"", type text}, {""TIEMPO DE SERVICIO"", type text}, {""CÓDIGO UNIDAD ORGANIZATIVA"", Int64.Type}, {""UNIDAD ORGANIZATIVA"", type text}, {""CÓDIGO RESPONSABLE DE UNIDAD ORGANIZATIV"", Int64.Type}, {""RESPONSABLE D" & _
        "E UNIDAD ORGANIZATIVA"", type text}, {""UNIDAD ORGANIZATIVA PADRE"", type text}, {""MEDIDA DE CONTRATACIÓN"", type text}, {""TIPO DE DEPOSITO"", type text}, {""GRUPO DE PERSONAL"", type text}, {""CORREO ELECTRONICO"", type text}, {""USUARIO DE RED"", type text}, {""USUARIO SAP"", type text}, {""COMUNIDAD"", type text}})," & Chr(13) & "" & Chr(10) & "    #""Columnas con nombre cambiado"" = Tabl" & _
        "e.RenameColumns(#""Tipo cambiado"",{{""DEPARTAMENTO"", ""DEPARTAMENTO.1""}})," & Chr(13) & "" & Chr(10) & "    #""Tipo cambiado1"" = Table.TransformColumnTypes(#""Columnas con nombre cambiado"",{{""Número de documento"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Tipo cambiado1"""
          
    'Codigo para inserar y agregar reccursos
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location= " + "DATA SAP" + CStr(nombre_hoja) + "-" + CStr(numero_hoja) + CStr(numero_iteracion) + ";Extended Properties=""""" _
        , Destination:=Range("$A$10")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" + "DATA SAP" + CStr(nombre_hoja) + "-" + CStr(numero_hoja) + CStr(numero_iteracion) + "]")
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
        .ListObject.DisplayName = "DATA_SAP_REPORTE"
        .Refresh BackgroundQuery:=False
    End With
    
'Application.Wait (Now + TimeValue("0:00:15"))
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub PEGAR_EXCEL1854()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Call cerrar_archivo_export_maestra
    Call ELIMINAR_TABLA_SAP_FINAL
    
    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet
    
    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets("PRUEBA")
    Set wbO = Workbooks.Open("C:\Macros LIMA\VALIDACION TXT PLAME\MC PROYECTO\REPORTES\EXPORTABLE.XLSX")
    
    ' Pegar datos en hoja SAP
    wbO.Sheets(1).UsedRange.Copy Destination:=ThisWorkbook.Worksheets("SAP").Range("A10")
    
    ' Convertir datos pegados en tabla
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("SAP").ListObjects.Add(SourceType:=xlSrcRange, _
        Source:=ThisWorkbook.Worksheets("SAP").Range("A10").CurrentRegion, _
        TableStyleName:="TableStyleLight9")
    tbl.Name = "DATA_SAP_REPORTE"
    
    wbO.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub ELIMINAR_TABLA_SAP_FINAL()

On Error Resume Next
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

    Dim NOMBRE As String
    Dim columnas As String
    Dim sSheetName As String
    Dim sTableName As String

    On Error Resume Next
    Dim C As Range
    Sheets("SAP").Select
    columnas = "A11:ZZ"
     'Define Variables
    sSheetName = "SAP"
    sTableName = "DATA_SAP_REPORTE"
    'Delete Table
     Sheets(sSheetName).ListObjects(sTableName).Delete

'Application.Wait (Now + TimeValue("0:00:05"))
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
     
End Sub




Sub ELIMINAR_TABLA_SAP18564()
On Error Resume Next

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

    Dim NOMBRE As String
    Dim columnas As String
    Dim sSheetName As String
    Dim sTableName As String

    On Error Resume Next
    Dim C As Range
    Sheets("SAP").Select
    columnas = "A11:ZZ"
     'Define Variables
    sSheetName = "SAP"
    sTableName = "DATA_SAP_REPORTE"
    'Delete Table
     Sheets(sSheetName).ListObjects(sTableName).Delete

'Application.Wait (Now + TimeValue("0:00:05"))
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False
     
End Sub








