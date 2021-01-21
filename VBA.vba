Sub VBAHomework()

       
    'Headers
    Range("A1").Value = "Ticker"
    Range("B1").Value = "Date"
    Range("C1").Value = "Open"
    Range("D1").Value = "High"
    Range("E1").Value = "Low"
    Range("F1").Value = "Close"
    Range("G1").Value = "Volume"
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Variable to hold Ticker name
    Dim Ticker_Name As String
    Ticker_Name = " "
    Dim Total_Ticker_Volume As Double
    Total_Ticker_Volume = 0
    Dim Open_price As Double
    Open_price = 0
    Dim Close_Price As Double
    Close_Price = 0
    Dim Yearly_Price_Change As Double
    Yearly_Price_Change = 0
    Dim Max_Ticker_Name As String
    Max_Ticker_Name = " "
    Dim Min_Ticker_Name As String
    Min_Ticker_Name = " "
    Dim Max_Percent As Double
    Max_Percent = 0
    Dim Min_Percent As Double
    Min_Percent = 0
    Dim Max_Volume_Ticker_Name As String
    Max_Volume_Ticker_Name = " "
    Dim Max_Volume As Double
    Max_Volume = 0
    
    
    Total_value = 0
    Table_row = 2
    
    'Value of Beginning Stock Value
    Open_price = Cells(2, 30).Value
    
                           
    'Loop through all tickers
    For i = 2 To Lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker_Name = Cells(i, 1).Value
            Range("I" & Table_row).Value = Ticker_Name
            
    'Calculate
        Close_Price = Cells(i, 6).Value
        Yearly_Price_Change = Close_Price - Open_price
        
        If Open_price <> 0 Then
            Yearly_price_change_percent = (Yearly_Price_Change / Open_price) * 100
            
        Total_Ticker_Volume = Total_Ticker_Volume + Cells(i, 7).Value
        
        Range("I" & Table_row).Value = Ticker_Name
        Range("J" & Table_row).Value = Yearly_Price_Change
        End If
        
        'Condition for Color coding, Red for Negative, Green for Postitive changes
        If (Yearly_Price_Change > 0) Then
            Range("J" & Table_row).Interior.ColorIndex = 4
            
        ElseIf (Yearly_Price_Change <= 0) Then
            Range("J" & Table_row).Interior.ColorIndex = 3
            
        End If
        
        Range("K" & Table_row).Value = (CStr(Yearly_price_change_percent) & "%")
        
        Range("L" & Table - Row).Value = Total_Ticker_Volume
        
        Talbe_row = Table_row = 1
        Total_value = 0
        
        End If
        
        'Beginning Price
        Open_price = Cells(i + 1, 3).Value
        
        'Determine Last Row
        Lastrow = Cells(Rows.Count, 1).End(x1up).Row
        
        'Calculations
        If (Yearly_price_change_percent > Max_Percent) Then
            Max_Percent = Yearly_price_change_percent
            Max_Ticker_Name = Ticker_Name
        
        ElseIf (Yearly_price_change_percent < Min_Percent) Then
            Min_Percent = Yearly_price_change_percent
            Min_Ticker_Name = Ticker_Name
        
        End If
        
        If (Total_Ticker_Volume > Max_Volume) Then
            Max_Volume = Total_Ticker_Volume
            Max_Volume_Ticker_Name = Ticker_Name
            
        End If
        
            Yearly_price_change_percent = 0
            Total_Ticker_Volume = 0
            Total_Ticker_Volume = Total_Ticker_Volume + Cells(i, 7).Value
            
        Next i
        
        'Print values in cells
        Range("Q2").Value = (CStr(Max_Percent) & "%")
        Range("Q3").Value = (CStr(Min_Percent) & "%")
        Range("P2").Value = Max_Ticker_Name
        Range("P3").Value = Min_Ticker_Name
        Range("Q4").Value = Max_Volume
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        

        
    
End Sub
