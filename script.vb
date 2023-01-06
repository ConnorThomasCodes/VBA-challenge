Sub stocks()
'This program will loop through all sheets on the workbook
'On each sheet it will loop through all rows,accumulating the volume value along the way, and pause whenever the ticker (column A) does not match the next entry
'It will then calculate the yearly change, percent change, and total volume for that stock and insert those values to the right of the table
'The loop will then continue the same process for the next stock, all the way to the end of the worksheet
'Additionally, the loop will also keep track of the greatest % increase, greatest % decrease, and greatest total volume for each sheet as it progresses
    
    'initialize variables
    Dim ticker As String
    Dim opening As Double
    Dim closing As Double
    Dim volume As Double
    Dim lastrow As Long
    Dim i As Long
    Dim outputRow As Long
    Dim ws As Worksheet
    Dim greatInc As Double
    Dim GIticker As String
    Dim greatDec As Double
    Dim GDticker As String
    Dim greatVol As Double
    Dim GVticker As String
    
    'loop through all worksheets in workbook
    For Each ws In Worksheets
        
        'grab last row from sheet (code from class)
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'initialize values before we begin our loop
        ticker = ws.Cells(2, 1).Value
        opening = ws.Cells(2, 3).Value
        outputRow = 2
        volume = 0
        
        greatInc = 0
        greatDec = 0
        greatVol = 0
        
        'insert row and column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        
        'loop through all rows in sheet
        For i = 2 To lastrow
            'accumulate volume in each row
            volume = volume + ws.Cells(i, 7).Value
            
            'check to see if next ticker string does NOT match current ticker string
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                'update closing value
                closing = ws.Cells(i, 6).Value
                
                'insert stored/calculated values into current outputRow
                ws.Cells(outputRow, 9).Value = ticker
                
                ws.Cells(outputRow, 10).Value = closing - opening
                If closing - opening < 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                ElseIf closing - opening > 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                End If
                ws.Cells(outputRow, 11).Value = (closing - opening) / opening
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                
                'check if current percent change is greater/less than stored value in greatInc/greatDec
                'update value if needed
                If (closing - opening) / opening > greatInc Then
                    greatInc = (closing - opening) / opening
                    GIticker = ticker
                End If
                If (closing - opening) / opening < greatDec Then
                    greatDec = (closing - opening) / opening
                    GDticker = ticker
                End If
                
                'set openining to next stock's value
                opening = ws.Cells(i + 1, 3).Value
                
                ws.Cells(outputRow, 12).Value = volume
                
                'check if current volume is greater than stored value in greatVol
                'update if needed
                If volume > greatVol Then
                    greatVol = volume
                    GVticker = ticker
                End If
                
                'reset volume, update ticker, and increment outputRow for next stock
                volume = 0
                ticker = ws.Cells(i + 1, 1).Value
                outputRow = outputRow + 1
            
            End If
        
        Next i
         
        'insert calculated greatest increase, decrease, and volume into table
        ws.Cells(2, 17).Value = GIticker
        ws.Cells(2, 18).Value = greatInc
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = GDticker
        ws.Cells(3, 18).Value = greatDec
        ws.Cells(3, 18).NumberFormat = "0.00%"
        ws.Cells(4, 17).Value = GVticker
        ws.Cells(4, 18).Value = greatVol
        
    Next ws

End Sub

