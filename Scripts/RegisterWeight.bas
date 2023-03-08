Attribute VB_Name = "RegisterWeight"
Option Explicit

Dim Database As Worksheet
Dim Interface As Worksheet
Dim BlankCell As Range
Dim weight As Double

Sub Main()
    Call Init
    Call GetWeight
    Call GetBlankCell
    If Not BlankCell Is Nothing Then
        BlankCell.Value = weight
        Call SetDate
        Call LinearRegresseion
    Else
        MsgBox "There is no more blank space in the database."
    End If
    
End Sub

Sub Init()
    Set Database = Sheets("Database")
    Set Interface = Sheets("Interface")
End Sub
    
Sub GetWeight()
    weight = Interface.Range("E13").Value
End Sub

Sub GetBlankCell()
    Set BlankCell = Database.Range("B2:B199").Find("", LookIn:=xlValues, lookat:=xlWhole)
End Sub

Sub SetDate()
    Dim dateCell As Range
    Set dateCell = Database.Range("A" & BlankCell.Row)
    dateCell.Value = Date
End Sub

Sub LinearRegresseion()
    Dim xRange As Range, yRange As Range, x() As Variant, y() As Variant, n As Long
    
    Set xRange = Database.Range("A2:A" & BlankCell.Row)
    Set yRange = Database.Range("B2:B" & BlankCell.Row)
    
    x = xRange.Value
    y = yRange.Value
    n = xRange.Count
    
    Dim xy() As Double, x_2() As Double, CovP() As Double, yErr() As Double, xErr2() As Double, yErr2() As Double
    Dim xMean As Double, yMean As Double, i As Long
    
    Dim m As Double, b As Double, r_num As Double, r_den As Double, r_2 As Double
    ReDim xy(1 To UBound(x)), x_2(1 To UBound(x))
    ReDim CovP(1 To UBound(x)), xErr2(1 To UBound(x)), yErr2(1 To UBound(x))
    
    xMean = WorksheetFunction.Average(xRange)
    yMean = WorksheetFunction.Average(yRange)
    
    For i = 1 To n - 1:
        xy(i) = x(i, 1) * y(i, 1)
        x_2(i) = x(i, 1) ^ 2
        CovP(i) = (x(i, 1) - xMean) * (y(i, 1) - yMean)
        xErr2(i) = (x(i, 1) - xMean) ^ 2
        yErr2(i) = (y(i, 1) - yMean) ^ 2
    Next i
    
    m = (n * WorksheetFunction.Sum(xy) - (WorksheetFunction.Sum(x) * WorksheetFunction.Sum(y))) / _
        (n * (WorksheetFunction.Sum(x_2) - WorksheetFunction.Sum(x) ^ 2))
    b = (WorksheetFunction.Sum(y) - m * WorksheetFunction.Sum(x)) / n
    

    
    r_num = WorksheetFunction.Sum(CovP) / (n - 1)
    r_den = Sqr(WorksheetFunction.Sum(xErr2) / (n - 1)) * Sqr(WorksheetFunction.Sum(yErr2) / (n - 1))
    r_2 = r_num / r_den
    
    Database.Range("J3").Value = m
    Database.Range("J4").Value = b
    Database.Range("J5").Value = r_2
    
End Sub
