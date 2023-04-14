Attribute VB_Name = "StockData"
Sub StockData():

'Variables
Dim Ticker As String
Ticker_Row = 2

Dim Max_Ticker As String

Dim YearOpen As Double

Dim YearClose As Double

Dim Yearly_Change As Double

Dim Percent_Change As Double

Dim Total_Volume As Double
Total_Volume = 0

Dim Max_Volume As Double
Max_Volume = 0

'Start of Loop for ALL Worksheets
For Each ws In Worksheets

'Last Row Variable
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Adding Column Titles
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
'Adding Values for Each Parameter
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
    
            Ticker = ws.Cells(i, 1)
            Total_Volume = Total_Volume + ws.Cells(i, 7)
                        
            ws.Cells(Ticker_Row, 9) = Ticker
            ws.Cells(Ticker_Row, 12) = Total_Volume
            
            Total_Volume = 0
            Ticker = ""
            
            YearClose = ws.Cells(i, 6)
            Yearly_Change = YearClose - YearOpen
            Percent_Change = (Yearly_Change / YearOpen)
                   
            ws.Cells(Ticker_Row, 10) = Yearly_Change
            ws.Cells(Ticker_Row, 11) = Percent_Change
            
            Ticker_Row = Ticker_Row + 1
            
        Else
            
            If ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
            
            YearOpen = ws.Cells(i, 3)
            
            End If
                            
            Total_Volume = Total_Volume + ws.Cells(i, 7)
            
        End If
        
    Next i

'Data Analysis
LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To LastRow2
        
        If ws.Cells(i, 12) > Max_Volume Then
        
            Max_Volume = ws.Cells(i, 12)
            Max_Ticker = ws.Cells(i, 9)
            
        End If
        
        If ws.Cells(i, 11) > Greatest_Increase Then
        
            Greatest_Increase = ws.Cells(i, 11)
            GIncTicker = ws.Cells(i, 9)
            
        End If
                       
        If ws.Cells(i, 11) < Greatest_Decrease Then
        
            Greatest_Decrease = ws.Cells(i, 11)
            GDecTicker = ws.Cells(i, 9)
            
        End If
    Next i
    
    ws.Range("Q4") = Max_Volume
    ws.Range("P4") = Max_Ticker
    ws.Range("Q2") = Greatest_Increase
    ws.Range("P2") = GIncTicker
    ws.Range("Q3") = Greatest_Decrease
    ws.Range("P3") = GDecTicker
    
'Reset for next ws
    Ticker_Row = 2
    Max_Volume = 0
    Max_Ticker = ""
    
'Formatting
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("Q4").NumberFormat = "0"
    ws.Range("I:L").Columns.AutoFit
    ws.Range("O:Q").Columns.AutoFit
    
    For i = 2 To LastRow2
    
        If ws.Cells(i, 10) < 0 Then
        
        ws.Cells(i, 10).Interior.ColorIndex = 3
        
        Else
        
        ws.Cells(i, 10).Interior.ColorIndex = 4
        
        End If
        
    Next i
    
Next ws

End Sub
