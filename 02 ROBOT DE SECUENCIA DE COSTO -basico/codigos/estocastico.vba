Attribute VB_Name = "M�dulo3"
Sub MODELADO_ESTOCASTICO()
Range("MODELADO_PL_PRELIMINAR[CAMBIO]").Select
ActiveCell.FormulaR1C1 = "=IFERROR([@SOLES]/[@DOLARES],)"

End Sub

