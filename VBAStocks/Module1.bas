Attribute VB_Name = "Module1"
Sub analysis()
' used assistance from stack over flow and internet forums to help me complete this assignment
' Set CurrentWs as a worksheet object variable.
    
    Dim CurrentWs As Worksheet
    Dim Summary_Table_Header As Boolean
    Dim c_spreadsheet As Boolean
    
    Summary_Table_Header = False       'Set Header flag
    c_spreadsheet = True              'Hard part flag
    
    ' Loop through all of the worksheets in the active workbook.
    For Each CurrentWs In Worksheets
    
        ' Set initial variable for holding the ticker name
        Dim Ticker_Name As String
        Ticker_Name = " "
        
        ' Set an initial variable for holding the total per ticker name
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        
        ' Moderate Part Variables
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim d_price As Double
        d_price = 0
        Dim d_percent As Double
        d_percent = 0
        ' Hard Part Variables
        Dim max_ticker_name As String
        max_ticker_name = " "
        Dim min_ticker_name As String
        min_ticker_name = " "
        Dim max_percent As Double
        max_percent = 0
        Dim min_percent As Double
        min_percent = 0
        Dim max_volume_ticker As String
        max_volume_ticker = " "
        Dim max_volume As Double
        max_volume = 0
        '----------------------------------------------------------------
         
        ' Keep track of the location for each ticker name
        ' in the summary table for the current worksheet
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        ' Set initial row count for the current worksheet
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        ' For all worksheet except the first one, the Results
        If Summary_Table_Header Then
            ' Set Titles for the Summary Table for current worksheet
            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"
            ' Set Additional Titles for new Summary Table on the right for current worksheet
            CurrentWs.Range("O2").Value = "Greatest % Increase"
            CurrentWs.Range("O3").Value = "Greatest % Decrease"
            CurrentWs.Range("O4").Value = "Greatest Total Volume"
            CurrentWs.Range("P1").Value = "Ticker"
            CurrentWs.Range("Q1").Value = "Value"
        Else
            'This is the first, resulting worksheet, reset flag for the rest of worksheets
            Summary_Table_Header = True
        End If
        
        ' Set initial value of Open Price for the first Ticker of CurrentWs,
        Open_Price = CurrentWs.Cells(2, 3).Value
        
        ' Loop from the beginning till its last row
        For i = 2 To Lastrow
        
      
            ' Check if still within the same ticker name,
            ' if not - write results to summary table
            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            
                ' Set the ticker name, we are ready to insert this ticker name data
                Ticker_Name = CurrentWs.Cells(i, 1).Value
                
                ' Calculate d_price and d_percent
                Close_Price = CurrentWs.Cells(i, 6).Value
                d_price = Close_Price - Open_Price
                ' Check Division by 0 condition
                If Open_Price <> 0 Then
                    d_percent = (d_price / Open_Price) * 100
                    
                End If
                
                ' Add to the Ticker name total volume
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
              
                
                ' Print the Ticker Name in the Summary Table, Column I
                CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
                ' Print the Ticker Name in the Summary Table, Column I
                CurrentWs.Range("J" & Summary_Table_Row).Value = d_price
                ' Fill "Yearly Change", i.e. d_price with Green and Red colors
                If (d_price > 0) Then
                    'Fill column with GREEN color - good
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (d_price <= 0) Then
                    'Fill column with RED color - bad
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 ' Print the Ticker Name in the Summary Table, Column I
                CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(d_percent) & "%")
                ' Print the Ticker Name in the Summary Table, Column J
                CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                ' Add 1 to the summary table row count
                Summary_Table_Row = Summary_Table_Row + 1
                ' Reset Delta_rice and d_percent holders, as we will be working with new Ticker
                d_price = 0
                ' Hard part,do this in the beginning of the for loop d_percent = 0
                Close_Price = 0
                ' Capture next Ticker's Open_Price
                Open_Price = CurrentWs.Cells(i + 1, 3).Value
              
                
                ' Hard part : Populate new Summary table on the right for the current spreadsheet HERE
                ' Keep track of all extra hard counters and do calculations within the current spreadsheet
                If (d_percent > max_percent) Then
                    max_percent = d_percent
                    max_ticker_name = Ticker_Name
                ElseIf (d_percent < min_percent) Then
                    min_percent = d_percent
                    min_ticker_name = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > max_volume) Then
                    max_volume = Total_Ticker_Volume
                    max_volume_ticker = Ticker_Name
                End If
                
                d_percent = 0
                Total_Ticker_Volume = 0
                
            
            'Else - If the cell immediately following a row is still the same ticker name,
            'just add to Totl Ticker Volume
            Else
                ' Encrease the Total Ticker Volume
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
            End If
            
      
        Next i

            
            ' Record all new counts to the new summary table on the right of the current spreadsheet
            If Not c_spreadsheet Then
            
                CurrentWs.Range("Q2").Value = (CStr(max_percent) & "%")
                CurrentWs.Range("Q3").Value = (CStr(min_percent) & "%")
                CurrentWs.Range("P2").Value = max_ticker_name
                CurrentWs.Range("P3").Value = min_ticker_name
                CurrentWs.Range("Q4").Value = max_volume
                CurrentWs.Range("P4").Value = max_volume_ticker
                
            Else
                c_spreadsheet = False
            End If
        
     Next CurrentWs

End Sub
