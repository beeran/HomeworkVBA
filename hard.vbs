
Sub hard()

For Each ws In Worksheets

Dim i As Long
Dim k As Long
Dim Ticker As String
Dim OpenIndex As Long
Dim CloseIndex As Long
Dim Volume As Double
OpenIndex = 2
k = 1
ws.Cells(k + 1, 11).Value = OpenIndex
Volume = 0

'Define the header of summary table
ws.Cells(k, 10).Value = "Ticker"
ws.Cells(k, 11).Value = "Yearly Change"
ws.Cells(k, 12).Value = "Percent Change"
ws.Cells(k, 13).Value = "Total stock volume"
ws.Cells(k, 16).Value = "Ticker"
ws.Cells(k, 17).Value = "Value"

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Total Volume"

RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To RowCount
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            Volume = Volume + ws.Cells(i, 7).Value
            
        Else
            k = k + 1
            OpenIndex = ws.Cells(k, 11).Value
            CloseIndex = i
            Ticker = ws.Cells(i, 1).Value
            Volume = Volume + ws.Cells(i, 7).Value
            YearlyChange = ws.Cells(CloseIndex, 6).Value - ws.Cells(OpenIndex, 3).Value
            If ws.Cells(OpenIndex, 3).Value = 0 Then
                PercentageChange = YearlyChange
            Else
               PercentChange = (YearlyChange / ws.Cells(OpenIndex, 3).Value)
            End If
            
            ws.Cells(k, 10).Value = Ticker
            ws.Cells(k, 11).Value = YearlyChange
            ws.Cells(k, 12).Value = PercentChange
            ws.Cells(k, 13).Value = Volume
            
            If ws.Cells(k, 11).Value < 0 Then
                ws.Cells(k, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(k, 11).Interior.ColorIndex = 4
            End If
            
            Volume = 0
            OpenIndex = i + 1
            ws.Cells(k + 1, 11).Value = OpenIndex
        End If

    Next i


Set myRange = ws.Range("L1:L" & RowCount)
MaxPercentage = Application.WorksheetFunction.Max(myRange)
MinPercentage = Application.WorksheetFunction.Min(myRange)
ws.Cells(2, 17).Value = MaxPercentage
ws.Cells(3, 17).Value = MinPercentage
Set myRange = ws.Range("M1:M" & RowCount)
MaxVolume = Application.WorksheetFunction.Max(myRange)
ws.Cells(4, 17).Value = MaxVolume
'To find the corresponding ticker
    For j = 2 To k
        If ws.Cells(j, 12).Value = ws.Range("Q2").Value Then
            ws.Cells(2, 16).Value = ws.Cells(j, 1).Value
        ElseIf ws.Cells(j, 12).Value = ws.Range("Q3").Value Then
            ws.Cells(3, 16).Value = ws.Cells(j, 1).Value
        ElseIf ws.Cells(j, 13).Value = ws.Range("Q4").Value Then
            ws.Cells(4, 16).Value = ws.Cells(j, 1).Value
        End If
    Next j
    
'Formatting
ws.Range("Q2:Q3").NumberFormat = "0.00%"
ws.Range("L1:L" & RowCount).NumberFormat = "0.00%"
ws.Cells.EntireColumn.AutoFit

Next ws

End Sub
