Attribute VB_Name = "Module2"
Sub Percetage_Values()
For Each ws In Worksheets
'To find greatest percentage Increase
Greatest_Decrease = Application.WorksheetFunction.Max(ws.Range("K2" & ":" & "K" & ws.Cells(Rows.Count, 11).End(xlUp).Row))
For i = 2 To ws.Cells(Rows.Count, 11).End(xlUp).Row
If ws.Cells(i, 11).Value = Greatest_Decrease Then
 Lowest_ticker = ws.Cells(i, 9).Value
 ws.Cells(2, 17).Value = FormatPercent(Greatest_Decrease)
End If
Next i
ws.Cells(2, 16) = Lowest_ticker

'To find greatest percentage Decrease
Greatest_Perc_Increase = Application.WorksheetFunction.Min(ws.Range("K2" & ":" & "K" & ws.Cells(Rows.Count, 11).End(xlUp).Row))
For i = 2 To ws.Cells(Rows.Count, 11).End(xlUp).Row
If ws.Cells(i, 11).Value = Greatest_Perc_Increase Then
 Greatest_ticker = ws.Cells(i, 9).Value
 ws.Cells(3, 17).Value = FormatPercent(Greatest_Perc_Increase)
End If
Next i
ws.Cells(3, 16) = Greatest_ticker

'To find greatest total volume
ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L2" & ":" & "L" & ws.Cells(Rows.Count, 12).End(xlUp).Row))
For i = 2 To ws.Cells(Rows.Count, 12).End(xlUp).Row
If ws.Cells(i, 12).Value = ws.Cells(4, 17).Value Then
 Greatest_Total_Volume = ws.Cells(i, 9).Value
End If
Next i
ws.Cells(1, 17) = "Value"
ws.Cells(1, 16) = "ticker"
ws.Cells(4, 15) = "Greatest Total Volume"
ws.Cells(4, 16) = Greatest_Total_Volume
ws.Cells(2, 15) = " Greatest % increase"
ws.Cells(3, 15) = "Greatest % decrease"


ws.Cells.EntireColumn.AutoFit
Next ws
End Sub

