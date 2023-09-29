VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Sub Stockdata_1()

Dim i, j, openIdx As Integer
Dim vol, rowCount As Long

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        vol = 0
        j = j + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If

Next i


End Sub


Sub Stockdata_2()

Dim i, j, openIdx As Integer
Dim vol, rowCount As Long
Dim Delta As Double

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openIdx = 2

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        Delta = Cells(i, 6).Value - Cells(openIdx, 3).Value
        Range("J" & 2 + j).Value = Delta
        
        vol = 0
        Delta = 0
        j = j + 1
        openIdx = openIdx + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If

Next i


End Sub


Sub Stockdata_3()

Dim i, j, openIdx As Integer
Dim vol, rowCount As Long
Dim Delta, percChange As Double

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openIdx = 2

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        Delta = Cells(i, 6).Value - Cells(openIdx, 3).Value
        Range("J" & 2 + j).Value = Delta
        
        percChange = Delta / Cells(openIdx, 3).Value
        Range("K" & 2 + j).Value = percChange
        
        vol = 0
        Delta = 0
        percChange = 0
        j = j + 1
        openIdx = openIdx + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If

Next i


End Sub


Sub Stockdata_4()

Dim i, j, openIdx As Integer
Dim vol, rowCount As Long
Dim Delta, percChange As Double

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openIdx = 2

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        Delta = Cells(i, 6).Value - Cells(openIdx, 3).Value
        Range("J" & 2 + j).Value = Delta
        Range("J" & 2 + j).NumberFormat = "0.00"
        
        If Delta > 0 Then
            Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf Delta < 0 Then
            Range("J" & 2 + j).Interior.ColorIndex = 3
        Else
            Range("J" & 2 + j).Interior.ColorIndex = 0
        End If
            
        
        percChange = Delta / Cells(openIdx, 3).Value
        Range("K" & 2 + j).Value = percChange
        
        vol = 0
        Delta = 0
        percChange = 0
        j = j + 1
        openIdx = openIdx + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If

Next i

End Sub


Sub Stockdata_5()

Dim i, j, openIdx As Integer
Dim vol, rowCount, rowCountK, rowCountL As Long
Dim Delta, percChange As Double
Dim bestIdx, worstIdx, mostIdx As Integer

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openIdx = 2

rowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        vol = vol + Cells(i, 7).Value
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("L" & 2 + j).Value = vol
        
        Delta = Cells(i, 6).Value - Cells(openIdx, 3).Value
        Range("J" & 2 + j).Value = Delta
        
        If Delta > 0 Then
            Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf Delta < 0 Then
            Range("J" & 2 + j).Interior.ColorIndex = 3
        Else
            Range("J" & 2 + j).Interior.ColorIndex = 0
        End If
            
        
        percChange = Delta / Cells(openIdx, 3).Value
        Range("K" & 2 + j).Value = percChange
        
        vol = 0
        Delta = 0
        percChange = 0
        j = j + 1
        openIdx = openIdx + 1
        
    Else
        
        vol = vol + Cells(i, 7).Value
        
    End If
    
Next i

rowCountK = Cells(Rows.Count, "K").End(xlUp).Row
rowCountL = Cells(Rows.Count, "L").End(xlUp).Row

Range("Q2") = WorksheetFunction.Max(Range("K2:K" & rowCountK).Value)
Range("Q3") = WorksheetFunction.Min(Range("K2:K" & rowCountK).Value)
Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCountL).Value)

bestIdx = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCountK).Value), Range("K2:K" & rowCountK).Value, 0)
worstIdx = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCountK).Value), Range("K2:K" & rowCountK).Value, 0)
mostIdx = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCountL).Value), Range("L2:L" & rowCountL).Value, 0)

Range("P2").Value = Cells(bestIdx + 1, 9).Value
Range("P3").Value = Cells(worstIdx + 1, 9).Value
Range("P4").Value = Cells(mostIdx + 1, 9).Value

End Sub

Sub Stockdata_6()

Dim ws As Object

For Each ws In Worksheets

Dim i, j, openIdx As Integer
Dim vol, rowCount, rowCountK, rowCountL As Long
Dim Delta, percChange As Double
Dim bestIdx, worstIdx, mostIdx As Integer

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

vol = 0
j = 0
openIdx = 2


rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowCount

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        vol = vol + ws.Cells(i, 7).Value
        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("L" & 2 + j).Value = vol
        
        Delta = ws.Cells(i, 6).Value - ws.Cells(openIdx, 3).Value
        ws.Range("J" & 2 + j).Value = Delta
        
        If Delta > 0 Then
            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf Delta < 0 Then
            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
        Else
            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
        End If
            
        
        percChange = Delta / ws.Cells(openIdx, 3).Value
        ws.Range("K" & 2 + j).Value = percChange
        
        vol = 0
        Delta = 0
        percChange = 0
        j = j + 1
        openIdx = openIdx + 1
        
    Else
        
        vol = vol + ws.Cells(i, 7).Value
        
    End If
    
Next i

rowCountK = ws.Cells(Rows.Count, "K").End(xlUp).Row
rowCountL = ws.Cells(Rows.Count, "L").End(xlUp).Row

ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & rowCountK).Value)
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & rowCountK).Value)
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & rowCountL).Value)


bestIdx = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCountK).Value), ws.Range("K2:K" & rowCountK).Value, 0)
worstIdx = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCountK).Value), ws.Range("K2:K" & rowCountK).Value, 0)
mostIdx = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCountL).Value), ws.Range("L2:L" & rowCountL).Value, 0)

ws.Range("K2:K" & rowCountK).NumberFormat = "0.00%"

ws.Range("P2").Value = ws.Cells(bestIdx + 1, 9).Value
ws.Range("P3").Value = ws.Cells(worstIdx + 1, 9).Value
ws.Range("P4").Value = ws.Cells(mostIdx + 1, 9).Value


Next ws

End Sub
