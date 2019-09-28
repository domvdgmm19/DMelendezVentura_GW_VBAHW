Attribute VB_Name = "Module1"
Sub MYStockData()
 
Dim ticker As String
Dim ws As Worksheet
Dim vol As Integer
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Total  As Double
Total = 0

For Each ws In ThisWorkbook.Worksheets

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    



Dim Summary_Table_Row As Double
Summary_Table_Row = 2

Dim rangeRowsCount As Double
    rangeRowsCount = Cells(Rows.Count, "A").End(xlUp).Row
    

    
For i = 2 To rangeRowsCount
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ticker = Cells(i, 1).Value
    Total = Total + Cells(i, 7).Value
            year_open = Cells(i, 3).Value
            year_close = Cells(i, 6).Value
            yearly_change = year_close - year_open
            percent_change = (year_close - year_open) / (year_close)

    
    Range("I" & Summary_Table_Row).Value = ticker
    Range("J" & Summary_Table_Row).Value = yearly_change
    Range("K" & Summary_Table_Row).Value = percent_change
    Range("L" & Summary_Table_Row).Value = Total
    
    Summary_Table_Row = Summary_Table_Row + 1
    Total = 0
    Else

    Total = Total + Cells(i, 7).Value
    
    End If
 
 Next i
 ws.Columns("K").NumberFormat = "0.00%"

    
    Next ws
 
 Dim rg As Range
    Dim g As Long
    Dim c As Long
    Dim color_cell As Range
    
    Set rg = ws.Range("J2", Range("J2")).End(xlDown)
    c = rg.Cells.Count
    
    For g = 1 To c
    Set color_cell = rg(g)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next g
 
 End Sub
    


