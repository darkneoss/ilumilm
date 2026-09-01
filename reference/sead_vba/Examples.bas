Attribute VB_Name = "Examples"
Option Explicit


Sub LoadExample()

Dim Range As Variant
Dim ResultsRange() As Variant
Dim ExNumber As Integer


Application.Calculation = xlCalculationManual
    
    Range = Sheet16.Range("EX_RG_Range").Value
    Sheet5.Range("IN_RG_Brange").Value = Range
    
    Range = Sheet16.Range("EX_RG_Range").Value
    Sheet5.Range("IN_RG_Urange").Value = Range
    
    Range = Sheet16.Range("EX_LS_Range").Value
    Sheet7.Range("IN_LS_range").Value = Range
    
    Range = Sheet16.Range("EX_In_Range1").Value
    Sheet6.Range("IN_In_range1").Value = Range
        
    Range = Sheet16.Range("EX_In_Range2").Value
    Sheet6.Range("IN_In_range2").Value = Range
    
    Range = Sheet16.Range("EX_In_Range3").Value
    Sheet6.Range("IN_In_range3").Value = Range
    
    ReDim ResultsRange(45, 42)
    ExNumber = Sheet16.Range("ExNumber")
    ResultsRange = Sheet16.Range("Ex" & ExNumber & "Results")
    Sheet10.Range("a4:Ap48") = ResultsRange

Application.Calculation = xlCalculationAutomatic

End Sub
