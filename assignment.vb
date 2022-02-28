Sub Worksheet_VBA()

' Set CurrentWs as active worksheet
Dim CurrentWs As Worksheet

' Loop through all of the worksheets in the active workbook
For Each CurrentWs In Worksheets

' Set initial variable for ticker
Dim Ticker_Name As String
Ticker_Name = " "

' Set an initial variable for total of each ticker name
Dim Total_Ticker_Volume As Double
Total_Ticker_Volume = 0

' Set variables for the rest
Dim Open_Price As Double
Open_Price = 0
Dim Close_Price As Double
Close_Price = 0
Dim Delta_Price As Double
Delta_Price = 0
Dim Delta_Percent As Double
Delta_Percent = 0
' Set variables for the bonus part
Dim MAX_TICKER_NAME As String
MAX_TICKER_NAME = " "
Dim MIN_TICKER_NAME As String
MIN_TICKER_NAME = " "
Dim MAX_PERCENT As Double
MAX_PERCENT = 0
Dim MIN_PERCENT As Double
MIN_PERCENT = 0
Dim MAX_VOLUME_TICKER As String
MAX_VOLUME_TICKER = " "
Dim MAX_VOLUME As Double
MAX_VOLUME = 0
' set variables for summary table
Dim Summary_Table_Row As Long
Summary_Table_Row = 2
Dim Lastrow As Long
Dim i As Long
        
Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

CurrentWs.Range("I1").Value = "Ticker"
CurrentWs.Range("J1").Value = "Yearly Change"
CurrentWs.Range("K1").Value = "Percent Change"
CurrentWs.Range("L1").Value = "Total Stock Volume"
' for bonus part on the summary greatest incraese and decrease
CurrentWs.Range("O2").Value = "Greatest % Increase"
CurrentWs.Range("O3").Value = "Greatest % Decrease"
CurrentWs.Range("O4").Value = "Greatest Total Volume"
CurrentWs.Range("P1").Value = "Ticker"
CurrentWs.Range("Q1").Value = "Value"
      
' Initial value of Open Price for the first Ticker of CurrentWs
Open_Price = CurrentWs.Cells(2, 3).Value

' Loop from row 2 till last row on each worksheet
For i = 2 To Lastrow
             
' this to determine if still within the same ticker name,otherwise add in result to summary table
If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
    
    ' Set the ticker name, we are ready to insert this ticker name data
    Ticker_Name = CurrentWs.Cells(i, 1).Value
    
    ' Calculate different price its percentage (open price and close price)
    Close_Price = CurrentWs.Cells(i, 6).Value
    different_Price = Close_Price - Open_Price
    
        If Open_Price <> 0 Then
            different_Percent = (different_Price / Open_Price) * 100
        End If
    
    ' for total volume on each ticker
    Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
    ' Print the Ticker Name in the Summary Table, Column I
    CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
    ' Print the Ticker Name in the Summary Table, Column I
    CurrentWs.Range("J" & Summary_Table_Row).Value = different_Price
    ' Fill "Yearly Change", i.e. Delta_Price with Green and Red colors
    If (different_Price > 0) Then
        'Fill column with GREEN color - good
        CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        ElseIf (different_Price <= 0) Then
            'Fill column with RED color - bad
            CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
                
    ' Print the Ticker Name in the Summary Table, Column I convert from value to string
  CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(different_Percent) & "%")
  ' Print the Ticker Name in the Summary Table, Column J
  CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
  
  Summary_Table_Row = Summary_Table_Row + 1
  ' to reset difference price to zero after each ticker
  different_Price = 0
  Close_Price = 0
  Open_Price = CurrentWs.Cells(i + 1, 3).Value

 'for bonus part
  If (different_Percent > MAX_PERCENT) Then
      MAX_PERCENT = different_Percent
      MAX_TICKER_NAME = Ticker_Name
  ElseIf (different_Percent < MIN_PERCENT) Then
      MIN_PERCENT = different_Percent
      MIN_TICKER_NAME = Ticker_Name
  End If
         
  If (Total_Ticker_Volume > MAX_VOLUME) Then
      MAX_VOLUME = Total_Ticker_Volume
      MAX_VOLUME_TICKER = Ticker_Name
  End If
  
  'to reset each of ticker percentage and volume to zero
  different_Percent = 0
  Total_Ticker_Volume = 0
                
            
            'Else for each ticker, add in all volume and move to next one
            Else
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
            End If
      ' for next row
        Next i
                'Print the bonus part from column O to Q
                CurrentWs.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
                CurrentWs.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
                CurrentWs.Range("P2").Value = MAX_TICKER_NAME
                CurrentWs.Range("P3").Value = MIN_TICKER_NAME
                CurrentWs.Range("Q4").Value = MAX_VOLUME
                CurrentWs.Range("P4").Value = MAX_VOLUME_TICKER
                
    'for next active worksheet
     Next CurrentWs
End Sub
