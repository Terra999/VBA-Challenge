Attribute VB_Name = "Module1"
Sub MultiYearStock():

    
    For Each ws In Worksheets

   ' Create variables to hold File Name, Ticker Name, Closing Price, Opening Price, Yearly Change
    
        Dim Ticker_Name As String
        Dim Closing_Price As Double
        Dim Opening_Price As Double
        Dim Yearly_Change As Double
        Dim Volume As Double
              

        ColorRed = 3
        ColorGreen = 4
        
        
        ' Keep track of the location for each ticker name in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
    
        ' number of rows in this worksheet (too bad the lc 'l' looks so much like the number '1')
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' MsgBox ("LastRow" + Str(LastRow))
        
        ' Putting headings on the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
    
        Ticker_Name = ws.Cells(2, 1).Value
        Opening_Price = ws.Cells(2, 3).Value
        Volume = 0
    
        For i = 2 To LastRow
        
            ' Adding the volume for each ticker name
            Volume = Volume + ws.Cells(i, 7).Value
            
            ' Identify where the ticker name changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Closing_Price = ws.Cells(i, 6).Value
                Yearly_Change = Closing_Price - Opening_Price
                
                ' Do not divide by 0 (zero) - got code from @SharonTemplin
                If Opening_Price <> 0 Then
                    
                    Percent_Change = (Closing_Price - Opening_Price) / Opening_Price
                        
                End If
                    
                    'Percent_Change = (Closing_Price - Opening_Price) / Opening_Price
                
                ' The stuff on the left is the target (it receives what is on the right)
                ' What is on the right is the value that you want to put in the target.
                
                ws.Cells(Summary_Table_Row, 9).Value = Ticker_Name
                ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
                ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
                ws.Cells(Summary_Table_Row, 12).Value = Volume
                ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                ws.Cells(Summary_Table_Row, 12).NumberFormat = "#,###"
                ' ws.Columns("A:L").EntireColumn.AutoFit
                
                
                    If ws.Cells(Summary_Table_Row, 10) > 0 Then
                    
                        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                        
                    ElseIf ws.Cells(Summary_Table_Row, 10) < 0 Then
                    
                        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    
                    End If
                    
                        
                    
                ' Getting ready for the next ticker
                
                Ticker_Name = ws.Cells(i + 1, 1).Value
                Opening_Price = ws.Cells(i + 1, 3).Value
                
                Volume = 0
                Summary_Table_Row = Summary_Table_Row + 1
              
                
            End If
        
        Next i
        
    Next ws
        
    

End Sub

