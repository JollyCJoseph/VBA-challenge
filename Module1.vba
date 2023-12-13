Attribute VB_Name = "Module1"
Sub Stock_Yearly()
For Each ws In Worksheets
'Initialysing variables
Dim ticker As String
Dim Greatest_ticker As String
Dim Lowest_ticker As String
Dim open_price As Double
Dim close_price As Double
Dim Brand_Total As Double
Dim yearly_change As Double
Dim Total_Stock_Volume As Double
Dim Greatest_Total_Volume As String
Dim Greatest_Perc_Increase As Double
Dim Greatest_Perc_Decrease As Double
Dim percentage_change As Double
Dim Summary_Table_Row As Integer
Total_Stock_Volume = 0
Greatest_Perc_Increase = 0
Lowest_Perc_Increase = 0
Summary_Table_Row = 2
open_price = ws.Cells(2, 3).Value
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
'ws.Cells(1, 17) = "Value"
'ws.Cells(1, 16) = "ticker"
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To (LastRow - 1)
 Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   ticker = ws.Cells(i, 1).Value
   ws.Range("I" & Summary_Table_Row).Value = ticker
   close_price = ws.Cells(i, 6).Value
   yearly_change = close_price - open_price
   percentage_change = (yearly_change / open_price)
   If yearly_change < 0 Then
   ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
   Else
   ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
   End If
   ws.Range("J" & Summary_Table_Row).Value = yearly_change
   ws.Range("K" & Summary_Table_Row).Value = FormatPercent(percentage_change)
   ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
   Summary_Table_Row = Summary_Table_Row + 1
   open_price = ws.Cells(i + 1, 3).Value
   Total_Stock_Volume = 0
 End If
Next i
ws.Cells.EntireColumn.AutoFit
Next ws

End Sub
