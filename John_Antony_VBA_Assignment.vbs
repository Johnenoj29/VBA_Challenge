Sub stock_market()

'loop through all sheets
For Each ws In Worksheets

'create variable to hold file name, and column names
    Dim WorksheetName As String
    Dim Ticker_Name As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim LastRow As Long
    Dim j As Integer
    
    j = 2
   
   ' Set an initial variable for holding the Ticker_Name Open_price,Close_Price total Yearly_Change, Percent_Change and Total_Stock_Volume
    Ticker_Name = " "
    Open_Price = 0
    Close_Price = 0
    Yearly_Change = 0
    Percent_Change = 0
    Total_Stock_Volume = 0
  

'grabbed the worksheetname
    WorksheetName = ws.Name

'add new column

' Add a Column for all sheets
    ws.Range("I1").EntireColumn.Insert

' Add the word "Ticker, Yearly_Change, Percent_Change and Total_Stock_Volume to columns header starting from I1
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly_Change"
    ws.Cells(1, 11).Value = "Percent_Change"
    ws.Cells(1, 12).Value = "Total_Stock_Volume"
         
'auto fit column headers
    ws.Range("I1:L1").Columns.AutoFit
  
'Keep track of the location for each Ticker in the summary table
    Dim Stock_summary_table As Long
    Stock_summary_table = 2

'set initial row count and last row already declared above
    Dim i As Long

 'determine the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'set initial value of open_price
    Open_Price = ws.Cells(2, 3).Value

'to loop from begining of each row to last row
    For i = 2 To LastRow

 ' Check if we are still within the Ticker, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
   'set/assign ticker name
            Ticker_Name = ws.Cells(i, 1).Value
   
   'set and calculate close price, yearly change and change percentage
   
            Close_Price = ws.Cells(i, 6).Value
            Yearly_Change = Close_Price - Open_Price
            Percent_Change = (Yearly_Change / Open_Price) * 100
  
   
   'add ticker name to total stock volume
   
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
   
   'print ticker name,yearly change,percent change and total stock volume on summary sheet and conditional formatting
   
            ws.Range("I" & Stock_summary_table).Value = Ticker_Name
            ws.Range("J" & Stock_summary_table).Value = Yearly_Change
            ws.Range("K" & Stock_summary_table).Value = (CStr(Percent_Change) & "%")
            ws.Range("L" & Stock_summary_table).Value = Total_Stock_Volume
   
            If Yearly_Change > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
            
   
   'reset to next ticker
            Stock_summary_table = Stock_summary_table + 1
            Yearly_Change = 0
            Close_Price = 0
            Open_Price = ws.Cells(i + 1, 3).Value
            j = j + 1
   
        Else
   
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
   
        End If

    

    Next i


Next ws

End Sub
