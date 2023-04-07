Attribute VB_Name = "Module1"
Sub Stock():

    For Each Stock_Data In Worksheets
    
        Dim Sheet_Name As String
        
        Dim i As Long
        
        Dim j As Long
        
        Dim Summary_Table_Row As Long
        
        Dim Last_Ticker_Row_Column_A As Long
        
        Dim Last_Ticker_Row_Column_I As Long
       
        Dim Percent_Change As Double
        
        Dim Greatest_Increase As Double
        
        Dim Greatest_Decrease As Double
       
        Dim Greatest_Total_Volume As Double
        
       
        
        
        Summary_Table_Row = 2
        
        
        j = 2
        
        
        Last_Ticker_Row_Column_A = Cells(Rows.Count, 1).End(xlUp).Row
        
        
            
            For i = 2 To Last_Ticker_Row_Column_A
            
               
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
             
                Cells(Summary_Table_Row, 9).Value = Cells(i, 1).Value
                
                
                Cells(Summary_Table_Row, 10).Value = Cells(i, 6).Value - Cells(j, 3).Value
                
                    
                    If Cells(j, 3).Value <> 0 Then
                    Percent_Change = ((Cells(i, 6).Value - Cells(j, 3).Value) / Cells(j, 3).Value)
                    
                    
                    Cells(Summary_Table_Row, 11).Value = Format(Percent_Change, "Percent")
                    
                    Else
                    
                    Cells(Summary_Table_Row, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                    
                
                Cells(Summary_Table_Row, 12).Value = WorksheetFunction.Sum(Range(Cells(j, 7), Cells(i, 7)))
                
            
                Summary_Table_Row = Summary_Table_Row + 1
                
                
                j = i + 1
                
                End If
            
            Next i
            
        
        Last_Ticker_Row_Column_I = Cells(Rows.Count, 9).End(xlUp).Row
    
        
        Greatest_Increase = Cells(2, 11).Value
        
        Greatest_Decrease = Cells(2, 11).Value
        
        Greatest_Total_Volume = Cells(2, 12).Value
        
        
            For i = 2 To Last_Ticker_Row_Column_I
            
            
                If Cells(i, 12).Value > Greatest_Total_Volume Then
                Greatest_Total_Volume = Cells(i, 12).Value
                Cells(4, 16).Value = Cells(i, 9).Value
                
                Else
                
                Greatest_Total_Volume = Greatest_Total_Volume
                
                End If
                
            
                If Cells(i, 11).Value > Greatest_Increase Then
                
                Cells(2, 16).Value = Cells(i, 9).Value
                Greatest_Increase = Cells(i, 11).Value
                
                
                Else
                
                Greatest_Increase = Greatest_Increase
                
                End If
                
            
                If Cells(i, 11).Value < Greatest_Decrease Then
                
                Cells(3, 16).Value = Cells(i, 9).Value
                
                Greatest_Decrease = Cells(i, 11).Value
            
                
                Else
                
                Greatest_Decrease = Greatest_Decrease
                
                End If
                
        
            Cells(2, 17).Value = Format(Greatest_Increase, "Percent")
            
            Cells(3, 17).Value = Format(Greatest_Decrease, "Percent")
            
            Cells(4, 17).Value = Format(Greatest_Total_Volume, "Scientific")
            
            Next i
            
    
            
    Next Stock_Data
        
End Sub
