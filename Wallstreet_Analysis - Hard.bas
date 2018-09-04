Attribute VB_Name = "Module1"
Sub Wallstreet_Analysis()

    Dim Current As Worksheet
   
    ' Loop through all of the worksheets in the active workbook.
    For Each ws In Worksheets
   
        Dim TotalAmount, TotalRow As Long
        Dim BeginingAmount, EndAmount, YearlyChange As Double
        Dim OpeningYear As Boolean
        
        TotalAmount = 0
        TotalRow = 0
        OpeningYear = True
        
        ws.Activate
        
        'Setup the column name
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
            
            'If Ticker value change
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                
                TotalRow = TotalRow + 1
                TotalAmount = TotalAmount + ws.Cells(i, 7).Value
                
                'Closing amount Column F
                EndAmount = ws.Cells(i, 6).Value
                YearlyChange = EndAmount - BeginingAmount
                
                ' Populate the Summary
                ws.Cells(TotalRow + 1, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TotalRow + 1, 10).Value = YearlyChange

                ws.Cells(TotalRow + 1, 11).Value = YearlyChange / BeginingAmount
                ws.Cells(TotalRow + 1, 12).Value = TotalAmount
                
                ' Update the background color and format
                If YearlyChange >= 0 Then
                    ws.Cells(TotalRow + 1, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(TotalRow + 1, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(TotalRow + 1, 10).NumberFormat = "0.000000000"
                
                ws.Cells(TotalRow + 1, 11).Style = "Percent"
                ws.Cells(TotalRow + 1, 11).NumberFormat = "0.00%"
                                
                TotalAmount = 0
                OpeningYear = True

            Else
                'If Ticker value the same
                
                TotalAmount = TotalAmount + ws.Cells(i, 7).Value
                
                If OpeningYear And ws.Cells(i, 3).Value <> 0 Then
                    'Get the opening amount Column C
                    BeginingAmount = ws.Cells(i, 3).Value
                    OpeningYear = False
                End If
                
            End If
            
        Next i
        
        ' Adjusting Autofit to the column
        ws.Columns("J:J").EntireColumn.AutoFit
        ws.Columns("I:I").EntireColumn.AutoFit
        ws.Columns("K:K").EntireColumn.AutoFit
        ws.Columns("L:L").EntireColumn.AutoFit
        
        '---------- Below are Codes for Hard Option ----------
        
        ws.Range("N2").Value = "Greatest % increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest total volume"
        
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        SummaryLastRow = ws.Range("I1").End(xlDown).Row

        Dim RngInc, RngVol As Range
        Dim MaxInc, MaxDec, MaxVal As Double
        Dim PosMaxInc, PosMaxDec, PosMaxVal As Double
        
        PosMaxInc = 0
        PosMaxDec = 0
        PosMaxVal = 0
        
        Set RngInc = ws.Range("K2:K" & (SummaryLastRow + 1))
        Set RngVol = ws.Range("L2:L" & (SummaryLastRow + 1))
        
        MaxInc = WorksheetFunction.Max(RngInc)
        MaxDec = WorksheetFunction.Min(RngInc)
        MaxVal = WorksheetFunction.Max(RngVol)
        
        ws.Range("P2").Value = MaxInc
        ws.Range("P3").Value = MaxDec
        ws.Range("P4").Value = MaxVal
        
        ws.Range("P2").Style = "Percent"
        ws.Range("P2").NumberFormat = "0.00%"
        
        ws.Range("P3").Style = "Percent"
        ws.Range("P3").NumberFormat = "0.00%"
        
        ' Get the row position
        PosMaxInc = WorksheetFunction.Match(MaxInc, RngInc, 0)
        PosMaxDec = WorksheetFunction.Match(MaxDec, RngInc, 0)
        PosMaxVal = WorksheetFunction.Match(MaxVal, RngVol, 0)
        
        '
        ws.Range("O2").Value = ws.Range("I" & (PosMaxInc + 1)).Value
        ws.Range("O3").Value = ws.Range("I" & (PosMaxDec + 1)).Value
        ws.Range("O4").Value = ws.Range("I" & (PosMaxVal + 1)).Value
        
        ws.Columns("N:N").EntireColumn.AutoFit
        ws.Columns("O:O").EntireColumn.AutoFit
        ws.Columns("P:P").EntireColumn.AutoFit
        
    Next ws

End Sub

