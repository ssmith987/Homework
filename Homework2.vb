Sub StockVolume()

Dim ws As Worksheet
For Each ws In Worksheets

  ' Set an initial variable for holding the Ticker name
  Dim Ticker As String

  ' Set an initial variable for holding the Total Stock Volume per Ticker
  Dim TotalStockVolume As Double
  TotalStockVolume = 0

  'Set an initial variable for holding the Stock Open
  Dim StockOpen As Double
  StockOpen = 0

  'Set initial Variable for holding the Stock Close
  Dim StockClose As Double
  StockClose = 0

  'Create summary table headers
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O3").Value = "Greatest % Decrease"
  ws.Range("O4").Value = "Greatest Total Volume"

  ' Keep track of the location for each Ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  'Find the last row of data
  Dim LRow As Long
  LRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through all stock data
    For I = 2 To LRow

        ' Check if we are still within the same ticker
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

        ' Set the Ticker name
        Ticker = ws.Cells(I, 1).Value

        ' Add to the Total Stock Volume
        TotalStockVolume = TotalStockVolume + ws.Cells(I, 7).Value

        ' Add to the Stock Open
        StockOpen = StockOpen + ws.Cells(I, 3).Value

        ' Add to the Stock Close
        StockClose = StockClose + ws.Cells(I, 6).Value

        ' Print the Ticker name in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker

        ' Print the Total Stock Volume to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = TotalStockVolume

        'Print the Yearly Change to the Summary Table
        ws.Range("J" & Summary_Table_Row).Value = StockClose - StockOpen

        'Print the Perent Change
            if StockOpen = 0 Then
                ws.Range("K" & Summary_Table_Row).Value = 0
            Else    
                ws.Range("K" & Summary_Table_Row).Value = (StockClose - StockOpen) / StockOpen
            end If

        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' Reset the Total Stock Value
        TotalStockVolume = 0

        ' Reset the Stock Open
        StockOpen = 0

        ' Reset the Stock Close
        StockClose = 0

        ' If the cell immediately following a row is the same Ticker...
        Else

        ' Add to the Total Stock Volume
        TotalStockVolume = TotalStockVolume + ws.Cells(I, 7).Value

        ' Add to the Stock Open
        StockOpen = StockOpen + ws.Cells(I, 3).Value

        ' Add to the Stock Close
        StockClose = StockClose + ws.Cells(I, 6).Value

        End If

    Next I
    
    ' Format cells
    ws.Range("K2:K" & Summary_Table_Row).NumberFormat = "0.00%"

    For j = 2 to Summary_Table_Row
        if ws.Cells(j,10).value = <0 Then
            ws.Cells(j,10).interior.colorindex = 3

        Else
            ws.Cells(j,10).interior.colorindex = 4

        End If

    Next J

  'Add Greatest Total data to table
  ws.Range("P2").value = "=INDEX(I:I,MATCH(Q2,K:K,0))"
  ws.Range("P3").value = "=INDEX(I:I,MATCH(Q3,K:K,0))"
  ws.Range("P4").value = "=INDEX(I:I,MATCH(Q4,L:L,0))"
  ws.Range("Q2").value = "=MAX(K:K)"
  ws.Range("Q2").NumberFormat = "0.00%"
  ws.Range("Q3").value = "=MIN(K:K)"
  ws.Range("Q3").NumberFormat = "0.00%"
  ws.Range("Q4").value = "=MAX(L:L)"

Next ws

End Sub

