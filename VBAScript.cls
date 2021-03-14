VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_market():

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets
    
    
    'Create the column headings
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    'Define Ticker variable
    Dim Ticker As String
    Ticker = " "
    Dim Ticker_volume As Double
    Ticker_volume = 0
    Dim stock_volume As Double
    stock_volume = 0
    
    'Set initial and last row for worksheet
    Dim Lastrow As Long
    Dim i As Long
    Dim j As Integer
    
    j = 2
    
    'Define Lastrow of worksheet
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set new variables for prices and percent changes
    Dim open_price As Double
    open_price = 0
    Dim close_price As Double
    close_price = 0
    Dim price_change As Double
    price_change = 0
    Dim price_change_percent As Double
    price_change_percent = 0
    
    open_price = ws.Cells(2, 3).Value
    
    
    
    'Do loop of current worksheet to Lastrow
    For i = 2 To Lastrow
    
        'Ticker symbol output
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(j, 9).Value = Ticker
            Total_Ticker_Volume = ws.Cells(i, 7) + Total_Ticker_Volume
            ws.Cells(j, 12).Value = Total_Ticker_Volume
            Total_Ticker_Volume = 0
        
            
            'Calculate change in Price
            close_price = ws.Cells(i, 6).Value
            price_change = close_price - open_price
            ws.Cells(j, 10).Value = price_change
            
            If price_change > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf price_change < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
           
            If open_price <> 0 Then
                price_change_percent = (price_change / open_price)
            Else
                price_change_percent = 0
            End If
        
            ws.Cells(j, 11).Value = price_change_percent
            ws.Cells(j, 11).NumberFormat = "0.00%"
            open_price = ws.Cells(i + 1, 3).Value
            j = j + 1
            
        Else
            
            Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
        
        End If
        
    
    
    Next i

Next ws

End Sub

