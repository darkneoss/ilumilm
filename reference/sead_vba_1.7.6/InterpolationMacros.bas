Attribute VB_Name = "InterpolationMacros"
' these not used anywhere as of now
Option Explicit
Option Compare Text


Sub RefreshIllCalcs()
Dim Break As Boolean

    'Sheet3.EnableCalculation = True
    'Sheet3.EnableCalculation = False

'Breakpoint
Break = True
End Sub
    Function RealEqual(x, y) As Boolean
        RealEqual = Abs(x - y) <= 0.00000001
        End Function
Function LinearInterp(XVals, YVals, TargetVal)
    Dim MatchVal
    On Error GoTo ErrXit
    With Application.WorksheetFunction
    MatchVal = .Match(TargetVal, XVals, 1)
    If MatchVal = XVals.Cells.count _
            And RealEqual(TargetVal, .Index(XVals, MatchVal)) Then
        LinearInterp = .Index(YVals, MatchVal)
    Else
        LinearInterp = .Index(YVals, MatchVal) _
            + (.Index(YVals, MatchVal + 1) - .Index(YVals, MatchVal)) _
                / (.Index(XVals, MatchVal + 1) _
                    - .Index(XVals, MatchVal)) _
                * (TargetVal - .Index(XVals, MatchVal))
        End If
        End With
    Exit Function
ErrXit:
    With Err
    LinearInterp = .Description & "(Number= " & .Number & ")"
        End With
    End Function

Option Explicit
Option Base 0
    Function CellAreaDecode(aRng, ByVal i As Long) As Range
        Dim AreaI As Long
        For AreaI = 1 To aRng.Areas.count
            If i <= aRng.Areas(AreaI).Cells.count Then
                Set CellAreaDecode = aRng.Areas(AreaI).Cells(i)
                Exit Function
            Else
                i = i - aRng.Areas(AreaI).Cells.count
                End If
            Next AreaI
        End Function
    Sub MapIn(InVal, ByRef Where)
        Dim i As Integer, HowMany As Integer
        If Not (TypeOf InVal Is Range) Then
            Where = InVal
        ElseIf InVal.Areas.count = 1 Then
            If InVal.Cells.count = 1 Then
                Where = InVal.Value
            ElseIf InVal.Columns.count = 1 Then
                Where = Application.WorksheetFunction.Transpose(InVal.Value)
            Else
                Where = Application.WorksheetFunction.Transpose( _
                    Application.WorksheetFunction.Transpose(InVal.Value))
                End If
        Else
            HowMany = InVal.Cells.count
            ReDim Where(HowMany - 1)
            For i = 0 To HowMany - 1
                Where(i) = CellAreaDecode(InVal, i + 1).Value
                Next i
            End If
        End Sub
Function Interpolate2D(InF, InX, InY, InX2, InY2)
    'InX contains two values, x0 and x1 _
     InY contains two values, y0 and y1 _
     InF contains 4 values, defined at (x0,y0), (x0,y1), _
                                       (x1,y0), (x1,y1) _
     InX2 and InY2 define the point at which the value of _
     the function is required
    'tests to ensure x0<x2<x1 and 'y0<y2<y1 needed
    Dim F, x, y, _
        x2 As Double, y2 As Double
    Dim NoXVals(1)
    MapIn InF, F
    MapIn InX, x
    MapIn InY, y
    MapIn InX2, x2
    MapIn InY2, y2
    NoXVals(0) = (F(2) - F(0)) / (x(1) - x(0)) * (x2 - x(0)) + F(0)
    NoXVals(1) = (F(3) - F(1)) / (x(1) - x(0)) * (x2 - x(0)) + F(1)
    Interpolate2D = _
        (NoXVals(1) - NoXVals(0)) / (y(1) - y(0)) * (y2 - y(0)) _
        + NoXVals(0)
    End Function
Function Interpolate2DArray(InF, InX, InY, InX2, InY2)
'Interpolate2DArray (Z-values, X-values, Y-values, X2, Y2)

'Z-values: Known z-values.  A single 2D range or a 2D array.
'X-values: Known x-values.  A single 1D range (or 1D array) with the same number of rows as in Z-values.
'Y-values: Known y-values.  A single 1D range (or 1D array) with the same number of columns as in Z-values.
'X2, Y2: The function returns the z value corresponding to this (x,y) pair.
    
    
    'Arguments should be in the following format.  However, currently _
     there is no validation of the arguments. _
     Each of the arguments can be either a range or an array. _
     InX is a single dimension array of x values sorted ascending. _
     InY is a single dimension array of y values sorted ascending. _
     InF is a 2D array with 1 entry for each (X, Y) pair of values in _
     the InX and InY arrays. _
     InX2 is a single value. _
     InY2 is a single value.
    Dim F, x, y, _
        x2 As Double, y2 As Double, _
        XIdx As Long, YIdx As Long
    Dim NoXVals(1)
    MapIn InF, F
    MapIn InX, x
    MapIn InY, y
    MapIn InX2, x2
    MapIn InY2, y2
    On Error GoTo ErrXit
    XIdx = Application.WorksheetFunction.Match(x2, x, 1)
    YIdx = Application.WorksheetFunction.Match(y2, y, 1)
    If XIdx = UBound(x) And RealEqual(x2, x(XIdx)) Then
        If YIdx = UBound(y) And RealEqual(y2, y(YIdx)) Then
            Interpolate2DArray = F(XIdx, YIdx)
        Else
            Interpolate2DArray = F(XIdx, YIdx) _
                + (F(XIdx, YIdx + 1) - F(XIdx, YIdx)) _
                    / (y(YIdx + 1) - y(YIdx)) * (y2 - y(YIdx))
            End If
    ElseIf YIdx = UBound(y) And RealEqual(y2, y(YIdx)) Then
        Interpolate2DArray = F(XIdx, YIdx) _
            + (F(XIdx + 1, YIdx) - F(XIdx, YIdx)) _
                / (x(XIdx + 1) - x(XIdx)) * (x2 - x(XIdx))
    Else
        NoXVals(0) = F(XIdx, YIdx) _
            + (F(XIdx + 1, YIdx) - F(XIdx, YIdx)) _
                / (x(XIdx + 1) - x(XIdx)) * (x2 - x(XIdx))
        NoXVals(1) = F(XIdx, YIdx + 1) _
            + (F(XIdx + 1, YIdx + 1) - F(XIdx, YIdx + 1)) _
                / (x(XIdx + 1) - x(XIdx)) * (x2 - x(XIdx))
        Interpolate2DArray = NoXVals(0) _
            + (NoXVals(1) - NoXVals(0)) _
                / (y(YIdx + 1) - y(YIdx)) * (y2 - y(YIdx))
        End If
    Exit Function
ErrXit:
    With Err
    Interpolate2DArray = .Description & "(Number= " & .Number & ")"
        End With
    End Function
 

