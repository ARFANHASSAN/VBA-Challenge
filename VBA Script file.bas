
Sub WorksheetsLoops()

        ' Set object variables in the current worksheet.
        Dim CurrentWs As Worksheet
        Dim Summary_Table_Header As Boolean
        Dim COMMAND_SPREADSHEET As Boolean
        
        Summary_Table_Header = True
        COMMAND_SPREADSHEET = True
        
        ' Loop through all the worksheets.
        For Each CurrentWs In Worksheets
        
            ' Initial variable for holding the ticker name
            Dim Ticker_Name As String
            Ticker_Name = " "
            
            ' Initial variable for holding the total per ticker name
            Dim Total_Ticker_Volume As Double
            Total_Ticker_Volume = 0
            
            ' Set new variables
            Dim Open_Price As Double
            Open_Price = 0
            Dim Close_Price As Double
            Close_Price = 0
            Dim Delta_Price As Double
            Delta_Price = 0
            Dim Delta_Percent As Double
            Delta_Percent = 0
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
        
        
            ' For the summary table
            Dim Summary_Table_Row As Long
            Summary_Table_Row = 2
            
            ' Initial row count
            Dim Lastrow As Long
            Dim i As Long
            
            Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        
                ' Titles for the Summary Table
                CurrentWs.Range("I1").Value = "Ticker"
                CurrentWs.Range("J1").Value = "Yearly Change"
                CurrentWs.Range("K1").Value = "Percent Change"
                CurrentWs.Range("L1").Value = "Total Stock Volume"
                CurrentWs.Range("O2").Value = "Greatest % Increase"
                CurrentWs.Range("O3").Value = "Greatest % Decrease"
                CurrentWs.Range("O4").Value = "Greatest Total Volume"
                CurrentWs.Range("P1").Value = "Ticker"
                CurrentWs.Range("Q1").Value = "Value"
        
                
        
        
            
            '   Initial value of Open Price and within the for loop
            Open_Price = CurrentWs.Cells(2, 3).Value
            
            ' Loop
            For i = 2 To Lastrow
            
          
                ' Checking  if we are not within the same ticker, write results to summary table
                If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
                
                    ' Set ticker name
                    Ticker_Name = CurrentWs.Cells(i, 1).Value
                    
                    ' Delta_Price and Delta_Percent
                    Close_Price = CurrentWs.Cells(i, 6).Value
                    Delta_Price = Close_Price - Open_Price
                    ' Division by 0 condition
                    If Open_Price <> 0 Then
                        Delta_Percent = (Delta_Price / Open_Price) * 100
                    Else
                        MsgBox ("For " & Ticker_Name & ", Row " & CStr(i) & ": Open Price =" & Open_Price & ". Fix <open> field manually and save the spreadsheet.")
                    End If
                    
                    ' Add total volume
                    Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
                  
                    
                    '  Ticker Name in Column I
                    CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
                    '  Ticker Name in Column J
                    CurrentWs.Range("J" & Summary_Table_Row).Value = Delta_Price
                    ' Fill Yearly Change with Green and Red colors
                    If (Delta_Price > 0) Then
                        'column with GREEN color are positive
                        CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    ElseIf (Delta_Price <= 0) Then
                        'column with RED color are with negative value
                        CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                    
                     ' Print the Ticker name
                    CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Delta_Percent) & "%")
                    ' Print the Ticker Name
                    CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                    
                    ' Add 1 to the summary table row count
                    Summary_Table_Row = Summary_Table_Row + 1
                    ' Reset Delta_price
                    Delta_Price = 0
                    'Set Close Price
                    Close_Price = 0
                    ' Tickers Open Price
                    Open_Price = CurrentWs.Cells(i + 1, 3).Value
                  
                    
                    ' To Populate final Summary table
                    If (Delta_Percent > MAX_PERCENT) Then
                        MAX_PERCENT = Delta_Percent
                        MAX_TICKER_NAME = Ticker_Name
                    ElseIf (Delta_Percent < MIN_PERCENT) Then
                        MIN_PERCENT = Delta_Percent
                        MIN_TICKER_NAME = Ticker_Name
                    End If
                           
                    If (Total_Ticker_Volume > MAX_VOLUME) Then
                        MAX_VOLUME = Total_Ticker_Volume
                        MAX_VOLUME_TICKER = Ticker_Name
                    End If
                    
                    ' Assigning the value to delta percent,and total ticker volume
                    Delta_Percent = 0
                    Total_Ticker_Volume = 0
                    
                Else
                    ' Encrease Total Ticker Volume
                    Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
                End If
          
            Next i

                ' For final result
                
                
               CurrentWs.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
                    CurrentWs.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
                    CurrentWs.Range("P2").Value = MAX_TICKER_NAME
                    CurrentWs.Range("P3").Value = MIN_TICKER_NAME
                    CurrentWs.Range("Q4").Value = MAX_VOLUME
                    CurrentWs.Range("P4").Value = MAX_VOLUME_TICKER
                     
                     
                Next CurrentWs
            
        
End Sub


