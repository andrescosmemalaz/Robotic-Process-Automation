Attribute VB_Name = "LIMPIEZA_DE_DATA"
Sub ELIMINAR_DATA_RESTANTE_TR5()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False


Sheets("TR5_PARAMETRIZADA").Select

     Dim rng As Range
     Dim i As Integer, counter As Integer
     
     Set rng = Range("TR5_PARAMETRIZADA[Tipo]")

     i = 1
    
     For counter = 1 To rng.Rows.Count
            If rng.Cells(i) = "" Then
                rng.Cells(i).EntireRow.Delete
            Else
                i = i + 1
            End If
    
     Next

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False


End Sub

Sub ELIMINAR_DATA_RESTANTE_TR6()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Sheets("TR6_PARAMETRIZADA").Select

     Dim rng As Range
     Dim i As Integer, counter As Integer
     Set rng = Range("TR6_PARAMETRIZADA[Tipo]")
    
     i = 1
    
     For counter = 1 To rng.Rows.Count
    
            If rng.Cells(i) = "" Then
                rng.Cells(i).EntireRow.Delete
            Else
                i = i + 1
            End If
    
     Next
     
     
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

Sub modelo_eliminacion_celdas_vacias()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

Call ELIMINAR_DATA_RESTANTE_TR5
Call ELIMINAR_DATA_RESTANTE_TR6

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
Application.CutCopyMode = False

End Sub

