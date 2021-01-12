Attribute VB_Name = "Module1"
Sub multiplestock()

Dim t As Double
Dim tsv As Double
Dim op As Double
Dim cp As Double

For Each ws In Worksheets

t = 2
tsv = 0

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Year Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Columns("I:L").AutoFit

For i = 2 To ws.Range("A1", ws.Range("A1").End(xlDown)).Count
    If ws.Cells(i, 2).Value = "20160101" Then
        op = ws.Cells(i, 3).Value
    ElseIf ws.Cells(i, 2).Value = ws.Range("B2").End(xlDown).Value Then
        cp = ws.Cells(i, 6).Value
    End If
    
    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1) Then
        tsv = tsv + ws.Cells(i, 7).Value
    Else
        tsv = tsv + ws.Cells(i, 7).Value
        ws.Cells(t, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(t, 10).Value = cp - op
            If ws.Cells(t, 10).Value < 0 Then
                ws.Cells(t, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(t, 10).Interior.ColorIndex = 4
            End If
        ws.Cells(t, 11).Value = cp / op - 1
        ws.Cells(t, 11).NumberFormat = "0.0000%"
        ws.Cells(t, 12).Value = tsv
        tsv = 0
        t = t + 1
    End If
    
Next i
Next

End Sub

Sub clear1()

ws.Range("I:L").ClearContents

Next

End Sub

Sub Bonus()

Dim t As Double
Dim lv As Range
Dim tv As Range
Dim min As Double
Dim max1 As Double
Dim max2 As Double
Dim tick As Range

For Each ws In Worksheets
    
t = 2
    
ws.Range("N2") = "Greatest % Increase"
ws.Range("N3") = "Greatest % Decrease"
ws.Range("N4") = "Greatest Total Value"
ws.Range("O1") = "Ticker"
ws.Range("P1") = "Value"
ws.Columns("N:P").AutoFit
    
    Set tick = ws.Range("I2", ws.Range("I2").End(xlDown))
    
    Set lv = ws.Range("K2", ws.Range("K2").End(xlDown))
    max1 = Application.WorksheetFunction.Max(lv)
    Cells(t, 16).Value = max1
    Cells(t, 16).NumberFormat = "0.0000%"
    ws.Range("O2").Formula = "=XLOOKUP(P2,$K:$K,$I:$I)"
    ws.Range("O2").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    t = t + 1
    
    min = Application.WorksheetFunction.min(lv)
    Cells(t, 16).Value = min
    Cells(t, 16).NumberFormat = "0.0000%"
    ws.Range("O3").Formula = "=XLOOKUP(P3,$K:$K,$I:$I)"
    ws.Range("O3").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    t = t + 1
    
    Set tv = Range("L2", ws.Range("L2").End(xlDown))
    max2 = Application.WorksheetFunction.Max(tv)
    Cells(t, 16).Value = max2
    ws.Range("O4").Formula = "=XLOOKUP(P4,$L:$L,$I:$I)"
    ws.Range("O4").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Next
End Sub

Sub clr()

ws.Columns("N:P").ClearContents

Next

End Sub

