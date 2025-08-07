Attribute VB_Name = "Module1"
Sub MLIstatistics()
Attribute MLIstatistics.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MLIstatistics Macro
'

'
    Range("A1:A4").Select
    Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Total Average"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Total Standard Deviation"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Total CV"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "average"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "standard deviation"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "CV"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[3]C[2]:R[996]C[2])"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "=STDEV.S(R[3]C[1]:R[996]C[1])"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-2]*100"
    Range("A3:D4").Select
    Range("A4").Activate
    Selection.Copy
    Rows("3:4").Select
    ActiveSheet.Paste
    Range("A2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=AVERAGEIF(R[2], ""<>""&0)"
    Range("A2").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=AVERAGEIF(R[1], ""average"", R[2])"
    Range("A3").Select
    ActiveWindow.SmallScroll ToRight:=0
    ActiveWindow.SmallScroll Down:=0
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=AVERAGEIFS(R[2], R[1], ""average"",R[2], "">0"")"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    Range("C2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=AVERAGEIFS(R[2], R[1], ""CV"",R[2], "">0"")"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = _
        "=AVERAGEIFS(R[2], R[1], ""standard deviation"",R[2], "">0"")"
    Range("B3").Select
End Sub
