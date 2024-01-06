Attribute VB_Name = "Module1"
Sub Stock_Data():

    Dim Stock_Ticker As String
    Dim lastRow As Long
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim Yearly_Change_Value As Double
    Dim Close_Date_Value As Double
    Dim Open_Date_Value As Double
    
    Dim Percent_Change_Value As Double
    
    Dim Open_Date_Tracker As Double
    Open_Date_Tracker = 0
    
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    
    
    For Each ws In Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        Last_Row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To Last_Row:
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Calculate yearly change, percent change, and update total stock volume
                Stock_Ticker = ws.Cells(i, 1).Value
                
                Close_Date_Value = ws.Cells(i, 6).Value
                
                Yearly_Change_Value = Close_Date_Value - Open_Date_Value
                
                Percent_Change_Value = (Yearly_Change_Value / Abs(Open_Date_Value)) * 100
                
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                ' Write data to summary table
                ws.Range("I" & Summary_Table_Row).Value = Stock_Ticker
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change_Value
                ws.Range("K" & Summary_Table_Row).Value = Format(Percent_Change_Value, "0.00%")
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset variables for the next stock
                Yearly_Change_Value = 0
                Close_Date_Value = 0
                Open_Date_Tracker = 0
                Open_Date_Value = 0
                Total_Stock_Volume = 0
            Else
                 ' Track open date if not already done
                If Open_Date_Tracker = 0 Then
                    Open_Date_Tracker = ws.Cells(i, 2).Value
                    Open_Date_Value = ws.Cells(i, 3).Value
                End If
                
                 ' Update total stock volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            End If
        
        Next i
        
        Yearly_Change_Last_Row = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
        
        ' Highlight positive and negative percentage changes
        For i = 2 To Yearly_Change_Last_Row:
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0) ' Highlight green
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0) ' Highlight red
            End If
        Next i
        
        ' Set Headers and row data
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Dim Greatest_Increase_Ticker As String
        Greatest_Increase_Ticker = ws.Range("I2").Value
        Dim Greatest_Decrease_Ticker As String
        Greatest_Decrease_Ticker = ws.Range("I2").Value
        Dim Greatest_Total_Volume_Ticker As String
        Greatest_Total_Volume_Ticker = ws.Range("I2").Value
        
        ' Initialize with the first values
        Greatest_Increase = ws.Range("K2").Value
        Greatest_Decrease = ws.Range("K2").Value
        Greatest_Total_Volume = ws.Range("L2").Value
        
        ' Loop through data to find greatest values
        For i = 2 To Yearly_Change_Last_Row:
            If ws.Cells(i, 11).Value > Greatest_Increase Then
                Greatest_Increase = ws.Cells(i, 11).Value
                Greatest_Increase_Ticker = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value < Greatest_Decrease Then
                Greatest_Decrease = ws.Cells(i, 11).Value
                Greatest_Decrease_Ticker = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 12).Value > Greatest_Total_Volume Then
                Greatest_Total_Volume = ws.Cells(i, 12).Value
                Greatest_Total_Volume_Ticker = ws.Cells(i, 9).Value
            End If
            
        Next i
        
        ' Write greatest values to the worksheet
        ws.Range("P2").Value = Greatest_Increase_Ticker
        ws.Range("Q2").Value = Format(Greatest_Increase, "0.00%")
        
        ws.Range("P3").Value = Greatest_Decrease_Ticker
        ws.Range("Q3").Value = Format(Greatest_Decrease, "0.00%")
        
        ws.Range("P4").Value = Greatest_Total_Volume_Ticker
        ws.Range("Q4").Value = Greatest_Total_Volume
        
        ' Reset summary table row counter for the next worksheet
        Summary_Table_Row = 2
        
    Next ws
    
End Sub
