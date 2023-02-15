Attribute VB_Name = "Module1"
Sub BMI_calculate()
Dim Weight, Height As Double

Dim ws As Worksheet
Set ws = ActiveWorkbook.Worksheets("Interface")

ws.Unprotect Password:="123"

Weight = Range("F14").Value
Height = Range("F15").Value
If Weight > 0 And Height > 0 Then
    Dim BMI As Double
    BMI = Weight / ((Height / 100) ^ 2)
    Range("F17").Value = Round(BMI, 2)
    If BMI > 0 And BMI <= 18.5 Then
        Range("C19").Value = "Underweight"
        Range("C19").Interior.Color = colorPalette(3)
        Range("C19").Font.Color = colorPalette(4)
        
    ElseIf BMI > 18.5 And BMI <= 25 Then
        Range("C19").Value = "Healthy Weight"
        Range("C19").Interior.Color = colorPalette(1)
        Range("C19").Font.Color = colorPalette(2)
    ElseIf BMI > 25 And BMI <= 30 Then
        Range("C19").Value = "Overweight"
        Range("C19").Interior.Color = colorPalette(3)
        Range("C19").Font.Color = colorPalette(4)
    ElseIf BMI > 30 Then
        Range("C19").Value = "Obese"
        Range("C19").Interior.Color = colorPalette(5)
        Range("C19").Font.Color = colorPalette(6)
        
    End If
Else
    Range("F17").Value = "Error"
    Range("C19").Value = "Check Weight and Height Values"
End If

ws.Protect Password:="123"
End Sub


Private Function colorPalette(index As Long) As Long
    Dim col(1 To 6) As Long
    col(1) = RGB(233, 255, 233)
    col(2) = RGB(0, 102, 0)
    col(3) = RGB(255, 233, 210)
    col(4) = RGB(152, 8, 8)
    col(5) = RGB(255, 204, 204)
    col(6) = RGB(152, 8, 8)
    colorPalette = col(index)
End Function
