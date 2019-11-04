Below is the code used in case it didn't save correctly. 
Attached are both the test worksheet and the final multiple year sheet. There are slight differences in the code to let the test worksheet figure out the greatest changes across the different alphabetical letters/ sheets. 
Also attached are screenshots of the multiple year worksheet. 



# VBA-Challenge
Code Used for VBA Homework- Multiple year sheet:

Sub Homework()

For Each ws In Worksheets

ws.Cells(1, 9).Value = " Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

Dim R As Double
Dim Total_Volume As Double
Dim Ticker_number As Double
Dim number_rows As Double
Dim I As Double
Dim Biggest_Percentage_Change As Double
Dim Smallest_Percentage_Change As Double
Dim Greatest_Volume As Double


Total_Volume = 0
number_rows = 0
Ticker_number = 0


For R = 2 To 100000
Total_Volume = Total_Volume + ws.Cells(R, 7).Value
number_rows = number_rows + 1

If ws.Cells(R, 1).Value <> ws.Cells(R + 1, 1) Then
Ticker_number = Ticker_number + 1
ws.Cells(Ticker_number + 1, 9).Value = ws.Cells(R, 1).Value
ws.Cells(Ticker_number + 1, 10).Value = ws.Cells(R, 6).Value - ws.Cells(R - number_rows + 1, 3).Value
ws.Cells(Ticker_number + 1, 11).Value = ws.Cells(Ticker_number + 1, 10).Value / ws.Cells(R - number_rows + 1, 3).Value
ws.Cells(Ticker_number + 1, 11).NumberFormat = "0.0%"
If ws.Cells(Ticker_number + 1, 11).Value > 0 Then
ws.Cells(Ticker_number + 1, 11).Interior.Color = RGB(0, 255, 0)
Else
ws.Cells(Ticker_number + 1, 11).Interior.Color = RGB(255, 0, 0)
End If

ws.Cells(Ticker_number + 1, 12).Value = Total_Volume
number_rows = 0
Total_Volume = 0

End If

Next R
Biggest_Percentage_Change = ws.Cells(2, 11).Value
Smallest_Percentage_Change = ws.Cells(2, 11).Value
Greatest_Volume = ws.Cells(2, 12).Value

For I = 2 To 100

If ws.Cells(I, 11).Value > Biggest_Percentage_Change Then
Biggest_Percentage_Change = ws.Cells(I, 11).Value
ws.Cells(2, 15).Value = ws.Cells(I, 9).Value
ws.Cells(2, 16).Value = ws.Cells(I, 11).Value
ws.Cells(2, 16).NumberFormat = "0.0%"
End If

If ws.Cells(I, 11).Value < Smallest_Percentage_Change Then
Smallest_Percentage_Change = ws.Cells(I, 11).Value
ws.Cells(3, 15).Value = ws.Cells(I, 9).Value
ws.Cells(3, 16).Value = ws.Cells(I, 11).Value
ws.Cells(3, 16).NumberFormat = "0.0%"

End If

If ws.Cells(I, 12).Value > Greatest_Volume Then
Greatest_Volume = ws.Cells(I, 12).Value
ws.Cells(4, 15).Value = ws.Cells(I, 9).Value
ws.Cells(4, 16).Value = ws.Cells(I, 12).Value
End If


Next I


Next ws
End Sub



Code used for test sheet: 

Sub Homework()

For Each ws In Worksheets

ws.Cells(1, 9).Value = " Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

Dim R As Double
Dim Total_Volume As Double
Dim Ticker_number As Double
Dim number_rows As Double
Dim I As Double
Dim Biggest_Percentage_Change As Double
Dim Smallest_Percentage_Change As Double
Dim Greatest_Volume As Double


Total_Volume = 0
number_rows = 0
Ticker_number = 0


For R = 2 To 100000
Total_Volume = Total_Volume + ws.Cells(R, 7).Value
number_rows = number_rows + 1

If ws.Cells(R, 1).Value <> ws.Cells(R + 1, 1) Then
Ticker_number = Ticker_number + 1
ws.Cells(Ticker_number + 1, 9).Value = ws.Cells(R, 1).Value
ws.Cells(Ticker_number + 1, 10).Value = ws.Cells(R, 6).Value - ws.Cells(R - number_rows + 1, 3).Value
ws.Cells(Ticker_number + 1, 11).Value = ws.Cells(Ticker_number + 1, 10).Value / ws.Cells(R - number_rows + 1, 3).Value
ws.Cells(Ticker_number + 1, 11).NumberFormat = "0.0%"
If ws.Cells(Ticker_number + 1, 11).Value > 0 Then
ws.Cells(Ticker_number + 1, 11).Interior.Color = RGB(0, 255, 0)
Else
ws.Cells(Ticker_number + 1, 11).Interior.Color = RGB(255, 0, 0)
End If

ws.Cells(Ticker_number + 1, 12).Value = Total_Volume
number_rows = 0
Total_Volume = 0

End If

Next R
Biggest_Percentage_Change = ws.Cells(2, 11).Value
Smallest_Percentage_Change = ws.Cells(2, 11).Value
Greatest_Volume = ws.Cells(2, 12).Value

For Each ss In Worksheets
For I = 2 To 100

If ss.Cells(I, 11).Value > Biggest_Percentage_Change Then
Biggest_Percentage_Change = ss.Cells(I, 11).Value
ws.Cells(2, 15).Value = ss.Cells(I, 9).Value
ws.Cells(2, 16).Value = ss.Cells(I, 11).Value
ws.Cells(2, 16).NumberFormat = "0.0%"
End If

If ss.Cells(I, 11).Value < Smallest_Percentage_Change Then
Smallest_Percentage_Change = ss.Cells(I, 11).Value
ws.Cells(3, 15).Value = ss.Cells(I, 9).Value
ws.Cells(3, 16).Value = ss.Cells(I, 11).Value
ws.Cells(3, 16).NumberFormat = "0.0%"

End If

If ss.Cells(I, 12).Value > Greatest_Volume Then
Greatest_Volume = ss.Cells(I, 12).Value
ws.Cells(4, 15).Value = ss.Cells(I, 9).Value
ws.Cells(4, 16).Value = ss.Cells(I, 12).Value
End If


Next I
Next ss


Next ws
End Sub


