Attribute VB_Name = "MODULO_DESCARGA_REPORTES_SUELDO"
Option Explicit
Public SapGuiAuto
Public objGui As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession
Sub SAP_extract_SUELDO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Call ELIMINAR_REPORTE_SUELDO
On Error GoTo ErrorHandler
Dim Appl As Object
Dim Connection As Object
Dim session As Object
Dim WshShell As Object
Dim SapGui As Object
Dim SelectedTransaccion As String
Dim Selected_sueldo_Fecha1 As String
Dim Selected_sueldo_Fecha2 As String
Dim Selected_sueldo_Fecha3 As String
Dim Selected_sueldo_Fecha4 As String
Dim Selected_sueldo_Sociedad As String
Dim Selected_sueldo_division_personal As String
Dim Selected_exepcion_one As String
Dim Selected_exepcion_two As String
Dim Selectusuario1 As String
Dim Selectedpassword1 As String
Dim answer As String
Dim Environment_sueldo As String


Sheets("CREDENCIALES SAP").Activate
Environment_sueldo = Worksheets("CREDENCIALES SAP").Range("B3").Value


Sheets("REPORTE SUELDO").Activate

'carpeta directorio
Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", 4
Set WshShell = CreateObject("WScript.Shell")

Do Until WshShell.AppActivate("SAP Logon ")
    Application.Wait Now + TimeValue("0:00:01")
Loop


Set SapGui = GetObject("SAPGUI")
Set Appl = SapGui.GetScriptingEngine
Set Connection = Appl.OpenConnection(Environment_sueldo, _
    True)
Set session = Connection.Children.Item(0)

Sheets("REPORTE SUELDO").Activate
'declaracion de variables por operacion
SelectedTransaccion = ActiveWorkbook.ActiveSheet.Range("B1").Value
Selected_sueldo_Fecha1 = ActiveWorkbook.ActiveSheet.Range("B2").Value
Selected_sueldo_Fecha2 = ActiveWorkbook.ActiveSheet.Range("D2").Value
Selected_sueldo_Fecha3 = ActiveWorkbook.ActiveSheet.Range("B3").Value
Selected_sueldo_Fecha4 = ActiveWorkbook.ActiveSheet.Range("D3").Value
Selected_sueldo_Sociedad = ActiveWorkbook.ActiveSheet.Range("B4").Value
Selected_sueldo_division_personal = ActiveWorkbook.ActiveSheet.Range("D4").Value
Selected_exepcion_one = ActiveWorkbook.ActiveSheet.Range("B5").Value
Selected_exepcion_two = ActiveWorkbook.ActiveSheet.Range("D5").Value

Sheets("CREDENCIALES SAP").Activate
Selectusuario1 = ActiveWorkbook.ActiveSheet.Range("B1").Value
Selectedpassword1 = ActiveWorkbook.ActiveSheet.Range("B2").Value

'usuario y contraseña
session.FindById("wnd[0]/usr/txtRSYST-MANDT").Text = 150
session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = Selectusuario1
session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = Selectedpassword1
session.FindById("wnd[0]/usr/txtRSYST-LANGU").Text = "ES"

Sheets("REPORTE SUELDO").Select
'Generacion de reporte en SAP
session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "ZGLHR051"
session.FindById("wnd[0]").SendVKey 0


session.FindById("wnd[0]/usr/chkP_0008").Selected = True
session.FindById("wnd[0]/usr/ctxtPNPBEGDA").Text = Selected_sueldo_Fecha1
session.FindById("wnd[0]/usr/ctxtPNPENDDA").Text = Selected_sueldo_Fecha2
session.FindById("wnd[0]/usr/ctxtPNPBEGPS").Text = Selected_sueldo_Fecha3
session.FindById("wnd[0]/usr/ctxtPNPENDPS").Text = Selected_sueldo_Fecha4

session.FindById("wnd[0]/usr/ctxtPNPBUKRS-LOW").Text = Selected_sueldo_Sociedad

session.FindById("wnd[0]/usr/ctxtP_FILE").Text = "C:\Macros LIMA\VALIDACION TXT PLAME\MC.XLS"
session.FindById("wnd[0]/usr/ctxtPNPSTAT2-LOW").SetFocus
session.FindById("wnd[0]/usr/ctxtPNPSTAT2-LOW").CaretPosition = 0
session.FindById("wnd[0]/usr/btn%_PNPSTAT2_%_APP_%-VALU_PUSH").Press


session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "3"
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").CaretPosition = 1
session.FindById("wnd[1]").SendVKey 0
session.FindById("wnd[1]/tbar[0]/btn[8]").Press
session.FindById("wnd[0]/usr/ctxtPNPABKRS-LOW").SetFocus
session.FindById("wnd[0]/usr/ctxtPNPABKRS-LOW").CaretPosition = 0
session.FindById("wnd[0]/usr/btn%_PNPABKRS_%_APP_%-VALU_PUSH").Press
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = Selected_exepcion_one
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").Text = Selected_exepcion_two
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").SetFocus
session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").CaretPosition = 2
session.FindById("wnd[1]/tbar[0]/btn[8]").Press
session.FindById("wnd[0]/tbar[1]/btn[8]").Press
Call PEGAR_DATA_SUELDO
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
Call CERRAR_SAP_LOGON


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


Exit Sub
ErrorHandler:
MsgBox "No se ha conectado a la VPN o la contraseña debe ser cambiada"

End Sub
Sub aumentar_D4()
Sheets("FILTROS").Select
Range("M20").Value = Range("M20").Value + 1
Range("N20").Value = Range("N20").Value + 1
End Sub
Sub cerrar_archivo_export_sueldo()
On Error Resume Next
    Workbooks.Open "C:\Macros LIMA\VALIDACION TXT PLAME\MC_IT0008.xls"
    Workbooks("MC_IT0008.xls").Worksheets("P0008").Activate
    'Workbooks("MC_IT0008.XLS").Close SaveChanges:=False
    Workbooks("MC_IT0008.xls").Save
    Workbooks("MC_IT0008.xls").Close
End Sub
Sub CERRAR_SAP_LOGON()
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "TASKKILL /F /IM saplogon.exe", , True
End Sub

Sub PEGAR_DATA_SUELDO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Call ELIMINAR_REPORTE_SUELDO
'Codigo para aumentar digito de extraccion de reporte
   
Call aumentar_D4
Sheets("FILTROS").Select
Dim nombre_hoja As String
nombre_hoja = Trim(Range("L20").Value)
Dim numero_hoja As String
numero_hoja = Trim(Range("M20").Value)
Dim numero_iteracion As String
numero_iteracion = Trim(Range("N20").Value)
Sheets("REPORTE SUELDO").Select

'Extraccion de tabla para determinar las ubicaciones de tableta
ActiveWorkbook.Queries.Add Name:="DATA SUELDO" + CStr(nombre_hoja) + "-" + CStr(numero_hoja) + CStr(numero_iteracion), Formula:= _
"let" & Chr(13) & "" & Chr(10) & "    Origen = Excel.Workbook(File.Contents(""C:\Macros LIMA\VALIDACION TXT PLAME\MC._IT0008.XLS""), null, true)," & Chr(13) & "" & Chr(10) & "    P1 = Origen{[Name=""P0008""]}[Data]," & Chr(13) & "" & Chr(10) & "    #""Encabezados promovidos"" = Table.PromoteHeaders(P1, [PromoteAllScalars=true])" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Encabezados promovidos"""
'Codigo para inserar y agregar reccursos
With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location= " + "DATA SUELDO" + CStr(nombre_hoja) + "-" + CStr(numero_hoja) + CStr(numero_iteracion) + ";Extended Properties=""""" _
    , Destination:=Range("$A$10")).QueryTable
    .CommandType = xlCmdSql
    .CommandText = Array("SELECT * FROM [" + "DATA SUELDO" + CStr(nombre_hoja) + "-" + CStr(numero_hoja) + CStr(numero_iteracion) + "]")
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
    .ListObject.DisplayName = "DATA_SUELDO"
    .Refresh BackgroundQuery:=False
End With

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub
Sub PEGAR_EXCEL_SUELDO18()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Call cerrar_archivo_export_maestra
    Call ELIMINAR_REPORTE_SUELDO
    
    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet
    
    Set wbI = ThisWorkbook
   ' Set wsI = wbI.Sheets("PRUEBA")
    Set wbO = Workbooks.Open("C:\Macros LIMA\VALIDACION TXT PLAME\MC._IT0008.XLS")
    
    ' Pegar datos en hoja SAP
    wbO.Sheets(1).UsedRange.Copy Destination:=ThisWorkbook.Worksheets("REPORTE SUELDO").Range("A10")
    
    ' Convertir datos pegados en tabla
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("REPORTE SUELDO").ListObjects.Add(SourceType:=xlSrcRange, _
        Source:=ThisWorkbook.Worksheets("REPORTE SUELDO").Range("A10").CurrentRegion, _
        TableStyleName:="TableStyleLight9")
    tbl.Name = "DATA_SUELDO"
    
    wbO.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub ELIMINAR_REPORTE_SUELDO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
    Dim NOMBRE As String
    Dim columnas As String
    Dim sSheetName As String
    Dim sTableName As String
    Dim C As Range
    
    Sheets("REPORTE SUELDO").Select
    columnas = "A:ZZ"
     'Define Variables
    sSheetName = "REPORTE SUELDO"
    sTableName = "DATA_SUELDO"
    'Delete Table
     Sheets(sSheetName).ListObjects(sTableName).Delete
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub ELIMINAR_REPORTE_SUELDO_PROCESO()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

On Error Resume Next
    Dim NOMBRE As String
    Dim columnas As String
    Dim sSheetName As String
    Dim sTableName As String
    Dim C As Range
    
    Sheets("PROCESO").Select
    columnas = "A:ZZ"
     'Define Variables
    sSheetName = "PROCESO"
    sTableName = "DATA_SUELDO"
    'Delete Table
     Sheets(sSheetName).ListObjects(sTableName).Delete
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub




