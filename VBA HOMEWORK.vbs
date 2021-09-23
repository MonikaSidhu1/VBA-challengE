Attribute VB_Name = "Module1"
Sub Totalstock()
Dim Ticker_Name As String
Ticker_Name = " "
Dim Total_Ticker_Volume As Long
Total_Ticker_Volume = 0
Dim Open_Price As Long
Open_Price = 0
Dim Close_Price As Long
Close_Price = 0
Dim Delta_Price As Long
Delta_Price = 0
Dim Delta_Percent As Long
Delta_Percent = 0
Dim Summary_Table_Row As Long
Summary_Table_Row = 0
Dim Lastrow As Long
Dim i As Long


Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Range("H1").Value = "Ticker"
Range("I1").Value = "Yearly Change"
Range("J1").Value = "Percent Change"
Range("K1").Value = "Total Stock Volume"
    
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

Open_Price = Cells(2, 3).Value

For i = 2 To Lastrow


Close_Price = Cells(i, 6).Value

If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

Ticker_Name = Cells(i, 1).Value


Delta_Price = Close_Price - Open_Price

If Open_Price <> 0 Then
End If
Total_Ticker_Volume = Total_Ticker_Volume + Cells(i, 7).Value

Range("H" & Summary_Table_Row).Value = Ticker_Name
Range("I" & Summary_Table_Row).Value = Delta_Price
 
 If (Delta_Price > 0) Then
 Range("I" & Summary_Table_Row).Interior.ColorIndex = 4
 ElseIf (Delta_Price <= 0) Then
 
 Range("I" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                Range("K" & Summary_Table_Row).Value = (CStr(Delta_Percent) & "%")
                Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                Summary_Table_Row = Summary_Table_Row + 1
                Delta_Price = 0
                
                Close_Price = 0
                Open_Price = Cells(i + 1, 3).Value
 
 End If
 Next
 

















End Sub

