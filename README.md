Sub Stock()
    For Each ws In Worksheets
    
    Dim WorksheetName As String
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    WorksheetName = ws.Name
    Dim Ticker As String
    Dim Total_Volume As Double
    Total_Volume = 0
    Dim Open_Count As Double
    Open_Count = 0
    Dim Close_Count As Double
    Close_Count = 0
    Dim Year_Change As Double
    yearly_change = 0
    Dim Yearly_Change_Per As Double
    Yearly_Change_Per = 0
    Dim i As Long
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim GIT, GDT As String
    Dim GIP, GDP As Double
    Dim GTV As Double
    Dim GTT As String
    Dim First_open As Double
    
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    ws.Range("p2").Value = "Greastest Percent Increase"
    ws.Range("p3").Value = "Greastest Percent Decrease"
    ws.Range("p4").Value = "Greastest Total Volume"
    ws.Range("q1").Value = "Ticker"
    ws.Range("r1").Value = "Value"
    
    GIT = ""
    GDT = ""
    GTT = ""
    GIP = 0
    GDP = 0
    GTV = 0
    
    'Sourcing from Class lecture
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        Open_Count = Open_Count + ws.Cells(i, 3).Value
        Close_Count = Close_Count + ws.Cells(i, 6).Value
        Year_Change = Close_Count - Open_Count
        
'Sourcing from https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.vlookup        
'First_open = Application.WorksheetFunction.VLookup(Cells(i, 10), Range("A2:G753001"), 3, "False")
                
        Yearly_Change_Per = Year_Change / Open_Count
        
        ws.Range("J" & Summary_Table_Row).Value = Ticker
        ws.Range("K" & Summary_Table_Row).Value = Year_Change
        ws.Range("L" & Summary_Table_Row).Value = Format(Yearly_Change_Per, "0.00")
        ws.Range("M" & Summary_Table_Row).Value = Total_Volume
        'ws.Range("N" & Summary_Table_Row).Value = Open_Count
        'ws.Range("O" & Summary_Table_Row).Value = Close_Count
                
        
        
        Summary_Table_Row = Summary_Table_Row + 1
        'Sourcing from class lecture
        'Reset Counter to 0 when new ticker is identified
        
        Total_Volume = 0
        Open_Count = 0
        Close_Count = 0
        
        
        Else
        
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        Open_Count = Open_Count + ws.Cells(i, 3).Value
        Close_Count = Close_Count + ws.Cells(i, 6).Value
        
        End If
        
               If Yearly_Change_Per > GIP Then
                    GIP = Yearly_Change_Per
                    GIT = ws.Cells(i, 1).Value
                    
                End If
                
                If Yearly_Change_Per < GDP Then
                    GDP = Yearly_Change_Per
                    GDT = ws.Cells(i, 1).Value
                End If
                
                If Total_Volume > GTV Then
                    GTV = Total_Volume
                    GTT = ws.Cells(i, 1).Value
                End If
                
                If ws.Cells(i, 11).Value > 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
            
        ElseIf ws.Cells(i, 11).Value < 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 3
        
        Else: ws.Cells(i, 11) = 0
            ws.Cells(i, 11).Interior.ColorIndex = 0
            
            
        End If
        
        If ws.Cells(i, 12).Value > 0 Then
            ws.Cells(i, 12).Interior.ColorIndex = 4
            
        ElseIf ws.Cells(i, 12).Value < 0 Then
            ws.Cells(i, 12).Interior.ColorIndex = 3
        
        Else
            ws.Cells(i, 12).Interior.ColorIndex = 0
                      
        End If
    
    Next i
    'Source Formatting from https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/format-function-visual-basic-for-applications
        ws.Cells(2, 17).Value = GIT
        ws.Cells(2, 18).Value = Format(GIP, "0.00%")
        
        ws.Cells(3, 17).Value = GDT
        ws.Cells(3, 18).Value = Format(GDP, "0.00%")
        
        ws.Cells(4, 17).Value = GTT
        ws.Cells(4, 18).Value = GTV
      
    Next ws
    
End Sub
