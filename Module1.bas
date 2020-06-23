Attribute VB_Name = "Module1"
Sub Test1()
    
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        
        Dim WorksheetName As String
        
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim Ticker As String
            
        Dim Yearly_Change As Double
        Yearly_Change = 0
            
        Dim Percent_Change As Double
        Percent_Change = 0
        
        Dim Summary_Index As Integer
        Summary_Index = 2
            
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Value"
            
        Closing_Price = 0
        Opening_Price = 0
            
        Opening_Price = Range("C2").Value
            
            
            For r = 2 To Last_Row
            
                If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
            
                Ticker = Cells(r, 1).Value
                    
                    
                Closing_Price = Cells(r, 6).Value
                Yearly_Change = (Closing_Price - Opening_Price)
                 
                
                
                Range("I" & Summary_Index).Value = Ticker
                Range("J" & Summary_Index).Value = Yearly_Change
                
                    If Yearly_Change >= 0 Then
                        Range("J" & Summary_Index).Interior.Color = vbGreen
                    Else
                        Range("J" & Summary_Index).Interior.Color = vbRed
                    End If
                    
                Range("K" & Summary_Index).NumberFormat = "0.00%"
                
                    If (Opening_Price = 0 And Closing_Price = 0) Then
                        Percent_Change = 0
                    ElseIf (Opening_Price = 0 And Closing_Price <> 0) Then
                        Percent_Change = 1
                    Else
                        Percent_Change = (Yearly_Change / Opening_Price)
                    End If
                Range("K" & Summary_Index).Value = Percent_Change
                
                
                
                Opening_Price = Cells(r + 1, 3).Value
                    
                Range("L" & Summary_Index).Value = Total_Stock_Volume
                Total_Stock_Volume = 0
                    
                Summary_Index = Summary_Index + 1
            Else
                
                Total_Stock_Volume = Total_Stock_Volume + Cells(r, 7).Value
                
       
               End If
                
                
            Next r
    
    Next ws

End Sub


