VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub hard_solution()
    
    'Declare constant
    Const IN_TYPE_COL As Integer = 1
    
    'Declare worksheet
    Dim ws As Worksheet
    
    'Loop to print the code in all the worksheets
    For Each ws In ThisWorkbook.Worksheets
       
        'Declare variables
        Dim Ticker As String
        Dim Year_open As Double
             Year_open = 0
        Dim Year_close As Double
            Year_close = 0
        Dim percent_change As Double
            percent_change = 0
        Dim Yearly_change As Double
            Yearly_change = 0
        Dim Total_Stock_Volume As Double
        Dim max_ticker_name As String
            max_ticker_name = 0
        Dim min_ticker_name As String
            min_ticker_name = 0
        Dim max_percent As Double
            max_percent = 0
        Dim min_percent As Double
            min_percent = 0
        Dim max_volume As Double
            max_volume = 0
        Dim min_volume As Double
            min_volume = 0
        Dim max_vol_ticker_name As String
            max_vol_ticker_name = " "
        Dim input_row As Long
            input_row = 2
        Dim output_row As Long
            output_row = 2
        
        'Insert new columns by giving range
         ws.Range("I1:J1:K1:L1").EntireColumn.Insert
         Dim name: name = Split("Ticker,Yearly Change,Percent Change,Total Stock Volume", ",")
         ws.Range("I1").Resize(1, UBound(name) + 1) = name
         
         'Assigns the Year open ,Volume and ticker values if input row is 2
         If input_row = 2 Then
         Year_open = ws.Cells(input_row, 3).Value
         Total_Stock_Volume = ws.Cells(input_row, 7)
         Ticker = ws.Cells(input_row, IN_TYPE_COL).Value
         End If
         
         'Set row count as last_row
         Dim last_row As Long
         
         'Retrieves the last cells
         last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
               
            'Loop starts from value 2 To last row of the sheet
            For input_row = 3 To last_row
                
                'Calculating Total stock volume and collecting identical ticker value
                If ws.Cells(input_row, IN_TYPE_COL).Value = ws.Cells((input_row - 1), IN_TYPE_COL).Value Then
                     Total_Stock_Volume = Total_Stock_Volume + ws.Cells(input_row, 7).Value
                     Ticker = ws.Cells(input_row, IN_TYPE_COL).Value
    
               'Calculating yearly change and percent change
                ElseIf ws.Cells(input_row, IN_TYPE_COL).Value <> ws.Cells((input_row - 1), IN_TYPE_COL).Value Then
                      Year_close = ws.Cells(input_row - 1, 6).Value
                      Yearly_change = Year_close - Year_open
                      percent_change = (Yearly_change / Year_open)
                        
                        
                         'Prints Ticker in given cells
                         ws.Cells(output_row, 9).Value = Ticker
                         
                         'Prints Yearly change in given cells
                         ws.Cells(output_row, 10).Value = Yearly_change
              
                         'It gives the red and green colors to yearly change based on values
                        If (Yearly_change > 0) Then
                          ws.Cells(output_row, 10).Interior.ColorIndex = 4
                        
                        ElseIf (Yearly_change <= 0) Then
                          ws.Cells(output_row, 10).Interior.ColorIndex = 3
                        
                        End If
                          
                          'Prints percent change and total stock volume
                          ws.Cells(output_row, 11).Value = percent_change
                          ws.Cells(output_row, 11).NumberFormat = "0.00%"
                          ws.Cells(output_row, 12).Value = Total_Stock_Volume
                         
                          'Increment the output_row
                          output_row = output_row + 1
                     
                        'Compare and stores the max and min of percent change,volume and ticker name
                         If (percent_change > max_percent) Then
                            max_percent = percent_change
                            max_ticker_name = Ticker
                            
                         ElseIf (percent_change < min_percent) Then
                            min_percent = percent_change
                            min_ticker_name = Ticker
                            
                         End If
                         
                         If (Total_Stock_Volume > max_volume) Then
                            max_volume = Total_Stock_Volume
                            max_vol_ticker_name = Ticker
                          
                         End If
                     
                     'Reset the values
                     percent_change = 0
                     Total_Stock_Volume = 0
                     
                    'Assisgning the year open
                    Year_open = ws.Cells(input_row, 3).Value
             
                End If
          
            Next input_row
                    
                    'print the values in assigned cells
                    ws.Range("O2").Value = "Greatest % Increase"
                    ws.Range("O3").Value = "Greatest % Decrease"
                    ws.Range("O4").Value = "Greatest Total volume"
                    ws.Range("P1").Value = "Ticker"
                    ws.Range("P2").Value = max_ticker_name
                    ws.Range("P3").Value = min_ticker_name
                    ws.Range("P4").Value = max_vol_ticker_name
                    ws.Range("Q1").Value = "Value"
                    ws.Range("Q2").Value = CStr(max_percent)
                    ws.Range("Q2").NumberFormat = "0.00%"
                    ws.Range("Q3").Value = CStr(min_percent)
                    ws.Range("Q3").NumberFormat = "0.00%"
                    ws.Range("Q4").Value = max_volume
                   
             
    Next ws
                 
 End Sub
