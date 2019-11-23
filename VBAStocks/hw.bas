Attribute VB_Name = "Module1"

Sub VBASTOCK()



Dim work As Worksheet
Dim ticker As String
Dim volume As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer

'this prevents my overflow error
On Error Resume Next

'run through each worksheet
For Each work In ThisWorkbook.Worksheets
    'set headers
    work.Cells(1, 10).Value = "Stock_Ticker"
    work.Cells(1, 11).Value = "Year_Change"
    work.Cells(1, 12).Value = "Percent_Change"
    work.Cells(1, 13).Value = "Total_Volume"

    'setup integers for loop
    Summary_Table_Row = 2
     volume = 0

    'loop
        For i = 2 To work.UsedRange.Rows.Count
             If work.Cells(i + 1, 1).Value <> work.Cells(i, 1).Value Then
            
            ticker = work.Cells(i, 1).Value
            volume = work.Cells(i, 7).Value

            year_open = work.Cells(i, 3).Value
            year_close = work.Cells(i, 6).Value
            yearly_change = year_close - year_open

            
            percent_change = (year_close - year_open) / year_close

            work.Cells(Summary_Table_Row, 10).Value = ticker
            work.Cells(Summary_Table_Row, 11).Value = yearly_change
            work.Cells(Summary_Table_Row, 12).Value = percent_change
            work.Cells(Summary_Table_Row, 13).Value = volume
            Summary_Table_Row = Summary_Table_Row + 1
            volume = volume + Cells(i, 7).Value
            
            ElseIf work.Cells(i, 11).Value < 0 Then
            work.Cells(i, 11).Interior.color = vbRed
            work.Cells(i, 11).Font.ColorIndex = 1
            Else
            work.Cells(i, 11).Interior.color = vbGreen
            work.Cells(i, 11).Font.ColorIndex = 1
    
                
         End If
           
            
            
     Next i
    work.Columns("L").NumberFormat = "0.00%"
            
    
Next work

End Sub
       
     
     
         
   

