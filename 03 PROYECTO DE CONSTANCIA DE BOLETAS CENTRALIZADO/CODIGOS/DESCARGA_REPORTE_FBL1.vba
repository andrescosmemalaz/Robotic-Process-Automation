Attribute VB_Name = "DESCARGA_REPORTE_FBL1"
Sub EXTRACCION_SAP()


'Codigo para extraer información de SAP

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Funcion = "DESCARGA DE REPORTE DE SAP"
escribirLog Funcion, "Inicio del proceso de descarga de Reporte de Sap"

'Call ELIMINAR_TABLA_SAP
On Error GoTo ErrorHandler

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
Set ws = ThisWorkbook.Worksheets("REPORTE_SAP")

Selectedexcepcion_environment = Worksheets("CREDENCIALES SAP").Range("B3").Value


'carpeta directorio
Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", 4
Set WshShell = CreateObject("WScript.Shell")

Do Until WshShell.AppActivate("SAP Logon ")
    Application.Wait Now + TimeValue("0:00:01")
Loop

Set SapGui = GetObject("SAPGUI")
Set Appl = SapGui.GetScriptingEngine

Set Connection = Appl.OpenConnection(Selectedexcepcion_environment, _
    True)
Set session = Connection.Children.Item(0)


SelectedFecha1 = Worksheets("REPORTE_SAP").Range("B2").Value

SelectedFecha2 = Worksheets("REPORTE_SAP").Range("D2").Value


Selectusuario = Worksheets("CREDENCIALES SAP").Range("B1").Value
Selectedpassword = Worksheets("CREDENCIALES SAP").Range("B2").Value


With session
    .findById("wnd[0]/usr/txtRSYST-MANDT").Text = 150
    .findById("wnd[0]/usr/txtRSYST-BNAME").Text = Selectusuario
    .findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Selectedpassword
    .findById("wnd[0]/usr/txtRSYST-LANGU").Text = "ES"
End With


session.findById("wnd[0]").Maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "fbl1n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").Text = "SYOSSIGH"
session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = SelectedFecha1
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").Text = SelectedFecha2
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").SetFocus
session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").caretPosition = 2
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Macros\PROTOTIPO CONSTANCIAS\REPORTE CONSTANCIA"
'session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Macros\PROTOTIPO CONSTANCIAS\REPORTE CONSTANCIA\PRUEBA"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "EXPORTABLE_CONSTANCIA.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press



'Call PEGAR_EXCEL
'MsgBox "Proceso realizado con éxito"

On Error Resume Next
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press



While Connection.Children.Count > 0
    Set session = Connection.Children(0)
    session.findById("wnd[0]").Close
    On Error Resume Next
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    On Error GoTo 0
Wend
escribirLog Funcion, "Descarga Correcta del Reporte Sap"
escribirLog Funcion, "Final de descarga del Reporte Sap"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

Exit Sub

ErrorHandler:
MsgBox "No se ha conectado a la VPN o la contraseña debe ser cambiada"

End Sub

Sub PEGAR_EXCEL_FBLN()

    Funcion = "PEGADO DE REPORTE DE SAP"
    escribirLog Funcion, "Inicio del proceso de Pegado de Reporte de Sap en Excel"
    
    Application.EnableCancelKey = xlDisabled
    Call ELIMINAR_TABLA_SAP
    
    Dim wbI As Workbook, wbO As Workbook
    Set wbI = ThisWorkbook
    Dim wsI As Worksheet
    Set wsI = wbI.Sheets("PROCESO")
    'Dim wbO As Workbook
    Set wbO = Workbooks.Open("C:\Macros\PROTOTIPO CONSTANCIAS\REPORTE CONSTANCIA\EXPORTABLE_CONSTANCIA.XLSX")
    
    
    Dim ws As Worksheet
    
    Dim C As Range
    Dim columnas As String
    Dim hoja As String
    Dim letra_columna As String
    Dim nombre_libro As String
    
    
    wbO.Sheets(1).Cells.Copy wsI.Cells
    wbO.Close SaveChanges:=False
    wbI.Sheets("PROCESO").Activate
    wbI.Sheets("PROCESO").Range("A15").Activate
    
    Dim tbl As Range
    Set tbl = wsI.Range("A15").CurrentRegion
    
    wsI.ListObjects.Add(SourceType:=xlSrcRange, Source:=tbl).Name = "DATA_SAP_FBLN"
    wsI.ListObjects("DATA_SAP_FBLN").Range.Cut _
        Destination:=Worksheets("REPORTE_SAP").Range("A15")
    wbI.Sheets("REPORTE_SAP").Activate
    
    'escribirLog Funcion, "Final del proceso de Pegado de Reporte de Sap en Excel"
    escribirLog Funcion, "Proceso correcto de pegado de reporte Sap"
    escribirLog Funcion, "Final del proceso de Pegado de Reporte de Sap en Excel"
    
End Sub


Sub ELIMINAR_TABLA_SAP()

Funcion = "ELIMINAR  REPORTE DE SAP"
escribirLog Funcion, "Inicio del proceso de eliminar el reporte existente de Sap en Excel"
    
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
    columnas = "A15:Z15"
     'Define Variables
    sSheetName = "REPORTE_SAP"
    sTableName = "DATA_SAP_FBLN"
    'Delete Table
     Set ws = Sheets(sSheetName)
     ws.ListObjects(sTableName).Delete
     escribirLog Funcion, "Final del proceso de eliminar el reporte existente de Sap en Excel"

'Application.Wait (Now + TimeValue("0:00:05"))
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False



End Sub
