Attribute VB_Name = "modulo_rib"
Sub CALMDOWM()
Attribute CALMDOWM.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CALMDOWM Macro
'

'
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(IF([@[TIPO DE REGIMEN SALUD SAP]]=""AFILIADO"",TRUE,FALSE),IF(TRIM([@[TIPO DE REGIMEN SALUD TR]])=""ESSALUD REGULAR"",TRUE,FALSE))=TRUE,""REGISTRAR EPS"","""")"
    Range("AE11").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(IF([@[TIPO DE REGIMEN SALUD SAP]]=""AFILIADO"",TRUE,FALSE),IF(TRIM([@[TIPO DE REGIMEN SALUD TR]])=""ESSALUD REGULAR"",TRUE,FALSE))=TRUE,""REGISTRAR EPS"","""")"
    Range("Y4").Select
End Sub
Sub RONDAMON()
Attribute RONDAMON.VB_ProcData.VB_Invoke_Func = " \n14"
'
' RONDAMON Macro
'

'
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP([@[NUMERO DOCUMENTO TR5]],DATA_SAP[[Número de documento]:[SUELDO]],13,0)"
    Range("AB12").Select
End Sub
