 Sub Stock_Analyzer()
'define items 
    For Each ws In Worksheets
        Dim Ticker As String
        Dim Volume As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim Percent As Double
        Dim Position As Long
        Dim Lastrow As Long
        Dim i As Long
'sets values equal to zero 
        Ticker = " "
        Volume = 0
        OpenPrice = 0
        ClosePrice = 0
        YearlyChange = 0
        Percent = 0
        Position = 2
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        OpenPrice = ws.Cells(2, 3).Value
'add titles to the columns 
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        For i = 2 To Lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker = ws.Cells(i, 1).Value
                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                
                If OpenPrice <> 0 Then
                    Percent = (YearlyChange / OpenPrice) * 100
                End If
    'updates the columns 
                Volume = Volume + ws.Cells(i, 7).Value
                ws.Range("I" & Position).Value = Ticker
                ws.Range("J" & Position).Value = YearlyChange
    'makes the colors 
            'green
                If (YearlyChange > 0) Then
                    ws.Range("J" & Position).Interior.ColorIndex = 4
            'red        
                ElseIf (YearlyChange <= 0) Then
                    ws.Range("J" & Position).Interior.ColorIndex = 3
                End If
    'CSrt creates a string fxn             
                ws.Range("K" & Position).Value = (CStr(Percent) & "%")
                ws.Range("L" & Position).Value = Volume
    'resets the values             
                Position = Position + 1
                YearlyChange = 0
                ClosePrice = 0
                OpenPrice = ws.Cells(i + 1, 3).Value
                Percent = 0
                Volume = 0
            Else
                Volume = Volume + ws.Cells(i, 7).Value
            End If      
        Next i
     Next ws
End Sub
