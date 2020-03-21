Attribute VB_Name = "Module1"
Sub Multi_year_stock()

Dim ws As Worksheet

For Each ws In Worksheets

    Dim WorksheetName As String
    Dim Ticker As String
    Dim TickerTotal As Double
    Dim Summary_Table_Row As Integer
    Dim LastRowNumber As Long
    Dim LastColumnNumber As Integer
    Dim ColumnNum As Integer
    Dim RowNum As Integer
    Dim YearHeader As String
    Dim TheYear
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YrChg As Double
    Dim PctChg As Double
    
    
    WorksheetName = ws.Name
    MsgBox (WorksheetName)
       
       i = 0
       TickerTotal = 0
       Summary_Table_Row = 2
       LastRowNumber = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
       
    If ws.Cells(1, 9) <> "Ticker" Then
       ws.Cells(1, 9) = "Ticker"
       ws.Cells(1, 10) = "Yearly Change"
       ws.Cells(1, 11) = "Percent Change"
       ws.Cells(1, 12) = "Total Stock Volume"
    Else
        MsgBox "Ticker in place"
    End If
      
   ' Loop through all ticker symbols
    
    OpenPrice = ws.Cells(2, 3).Value
    
    MsgBox (LastRowNumber)
    
For i = 2 To LastRowNumber
    ' Check if we are still within the same ticker
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Collect the ticker name and calc the price difference..
      Ticker = ws.Cells(i, 1).Value
      
      ' Add to the Brand Total
      TickerTotal = TickerTotal + ws.Cells(i, 7).Value

      ' Print the ticker symbol in the Summary Table
       ws.Cells(Summary_Table_Row, 9) = Ticker
       
      ' Calc and print the yearly open to close price change and the percentage change..
        YrChg = (ClosePrice - OpenPrice)
        
        If OpenPrice <> 0 Then
            PctChg = (ClosePrice - OpenPrice) / OpenPrice
        Else
            PctChg = 100
        End If
        
        ws.Cells(Summary_Table_Row, 10) = YrChg
        
        If YrChg < 0 Then
           ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
        Else
           ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
        End If
                
        ws.Cells(Summary_Table_Row, 11) = PctChg
        ws.Cells(Summary_Table_Row, 11).NumberFormat = Percent
        ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
               
      ' Print the Ticker total volume to the Summary Table
       ws.Cells(Summary_Table_Row, 12) = TickerTotal

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      'Initialize ticker total
      TickerTotal = 0
      OpenPrice = ws.Cells(i + 1, 3).Value
      
     Else
        
      ' If the cell immediately following a row is the same ticker...
      ' add to the ticker total and keep saving the closing price..
      
        TickerTotal = TickerTotal + ws.Cells(i, 7).Value
        ClosePrice = ws.Cells(i + 1, 6).Value
      
    End If
    
  Next i
         
 Next ws
     
End Sub


