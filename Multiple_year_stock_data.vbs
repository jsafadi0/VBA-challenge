Sub VBA_Challenge()
For Each ws In Worksheets
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


Dim WorksheetName As String
Dim A As Long
Dim B As Long
Dim Tickercount As Long
Dim LastrowA As Long
Dim LastrowI As Long
Dim Percentchange As Double
Dim Greatincrease As Double
Dim Greatdecrease As Double
Dim Greatvolume As Double
WorksheetName = ws.Name


Tickercount = 2
B = 2
LastrowA = ws.Cells(Rows.Count, 1).End(xlUp).Row


For A = 2 To LastrowA
If ws.Cells(A + 1, 1).Value <> ws.Cells(A, 1).Value Then
ws.Cells(Tickercount, 9).Value = ws.Cells(A, 1).Value
ws.Cells(Tickercount, 10).Value = ws.Cells(A, 6).Value - ws.Cells(B, 3).Value
If ws.Cells(Tickercount, 10).Value < 0 Then
ws.Cells(Tickercount, 10).Interior.ColorIndex = 3
Else
ws.Cells(Tickercount, 10).Interior.ColorIndex = 4
End If


If ws.Cells(B, 3).Value <> 0 Then
Percentchange = ((ws.Cells(A, 6).Value - ws.Cells(B, 3).Value) / ws.Cells(B, 3).Value)
ws.Cells(Tickercount, 11).Value = Format(Percentchange, "Percent")
Else
ws.Cells(Tickerount, 11).Value = Format(0, "Percent")
End If
ws.Cells(Tickercount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(B, 7), ws.Cells(A, 7)))
Tickercount = Tickercount + 1
B = A + 1
End If


Next A


LastrowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
Greatvolume = ws.Cells(2, 12).Value
Greatincrease = ws.Cells(2, 11).Value
Greatdecrease = ws.Cells(2, 11).Value
For A = 2 To LastrowI
If ws.Cells(A, 12).Value > Greatvolume Then
Greatvolume = ws.Cells(A, 12).Value
ws.Cells(4, 16).Value = ws.Cells(A, 9).Value
Else
Greatvolume = Greatvolume
End If
If ws.Cells(A, 11).Value > Greatincrease Then
Greatincrease = ws.Cells(A, 11).Value
ws.Cells(2, 16).Value = ws.Cells(A, 9).Value
Else
Greatincrease = Greatincrease
End If
If ws.Cells(A, 11).Value < Greatdecrease Then
Greatdecrease = ws.Cells(A, 11).Value
ws.Cells(3, 16).Value = ws.Cells(A, 9).Value
Else
Greatdecrease = Greatdecrease
End If


ws.Cells(2, 17).Value = Format(Greatincrease, "Percent")
ws.Cells(3, 17).Value = Format(Greatdecrease, "Percent")
ws.Cells(4, 17).Value = Format(Greatvolume, "Scientific")


Next A


Worksheets(WorksheetName).Columns("A:Z").AutoFit


Next ws
End Sub


