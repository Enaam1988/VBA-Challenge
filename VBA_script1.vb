Sub stock_analysis()

          ' Set an initial variable for holding the ticker name
          Dim Ticker_Name As String
        
          ' Set an initial variable for holding the Ticker Total Volume per Ticker
          Dim Ticker_Total As Double
          Ticker_Total_Volume = 0
        
          ' Keep track of the location for each ticker in the summary table
          Dim Summary_Table_Row As Integer
          Summary_Table_Row = 2
          
          ' Set headers for summary table
           Range("I1").Value = "Ticker"
           Range("J1").Value = "Yearly Change"
           Range("K1").Value = "Percent Change"
           Range("L1").Value = "Total Stock Volume"
          
          ' get total number of rows in sheet
          Dim Row_Count As Long
          Row_Count = Range("A1").End(xlDown).Row
        
          ' Loop through all ticker daily records
          For i = 2 To Row_Count
        
            ' Check if we are still within the same ticker, if it is not...
            Dim val1 As String
            Dim val2 As String
            Dim First_Open As Double
            Dim Last_Close As Double
            val1 = Cells(i + 1, 1).Value
            val2 = Cells(i, 1).Value
            
            ' detect if this is first row
            If i = 2 Then
                ' set first open for first ticker in data set
                First_Open = Cells(i, 3).Value
            End If
            
            If val1 <> val2 Then
        
              ' Set the Brand name
              Ticker_Name = Cells(i, 1).Value
        
              ' Add to the Brand Total
              Ticker_Total_Volume = Ticker_Total_Volume + Cells(i, 7).Value
              
              ' Last close to calc yearly and perc change
              Last_Close = Cells(i, 6).Value
        
              ' Print the Ticker name in the Summary Table
              Range("I" & Summary_Table_Row).Value = Ticker_Name
        
              ' Calculate yrly change
              Range("J" & Summary_Table_Row).Value = Last_Close - First_Open
              
              ' Calculate percentage change only if starting value was not 0
              ' 0 starting value would simultaneously cause divide by zero error
              ' and be meaningless
              If (First_Open <> 0) Then
                Range("K" & Summary_Table_Row).Value = ((Last_Close - First_Open) / First_Open) * 100
               Range("K" & Summary_Table_Row).Style = "Percent"
              Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
              End If
        
              ' Print the Ticker Total Volume to the Summary Table
              Range("L" & Summary_Table_Row).Value = Ticker_Total_Volume
        
              ' Add one to the summary table row
              Summary_Table_Row = Summary_Table_Row + 1
              
              ' Reset the Ticker Total Volume
              Ticker_Total_Volume = 0
              
              ' Get first open price for next ticker
              First_Open = Cells(i + 1, 3)
        
            ' If the cell immediately following a row is the same ticker...
            Else
        
              ' Add to the Ticker Total Volume
              Ticker_Total_Volume = Ticker_Total_Volume + Cells(i, 7).Value
        
            End If
        
          Next i
          
          'Conditional formatting that will highlight positive change in green and negative change in red
          Dim Fmt_Range As Range
          Set Fmt_Range = Range("J2:J" & CStr(Summary_Table_Row - 1))
          Fmt_Range.FormatConditions.Delete
          Fmt_Range.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
          Fmt_Range.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
          Fmt_Range.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
            Formula1:="=0"
          Fmt_Range.FormatConditions(2).Interior.Color = RGB(0, 255, 0)
         'script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
   
   
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double

GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

For i = 2 To Row_Count


    If Cells(i, 11).Value > GreatestIncrease Then
        GreatestIncrease = Cells(i, 11).Value
        Range("Q2").Value = GreatestIncrease
        Range("Q2").Style = "Percent"
        Range("Q2").NumberFormat = "0.00%"
        Range("P2").Value = Cells(i, 9).Value
    End If

    Next i

For i = 2 To Row_Count
    
    If Cells(i, 11).Value < GreatestDecrease Then
        GreatestDecrease = Cells(i, 11).Value
        Range("Q3").Value = GreatestDecrease
        Range("Q3").Style = "Percent"
        Range("Q3").NumberFormat = "0.00%"
        Range("P3").Value = Cells(i, 9).Value
    End If
    
   Next i

For i = 2 To Row_Count
    
    If Cells(i, 12).Value > GreatestVolume Then
        GreatestVolume = Cells(i, 12).Value
        Range("Q4").Value = GreatestVolume
        Range("P4").Value = Cells(i, 9).Value
End If
  Next i
  
End Sub
