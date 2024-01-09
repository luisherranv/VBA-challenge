# VBA-challenge
Module 2 Challenge - VBA Coding
```
Sub challenge2()
Dim sh As Worksheet

'Loop to ensure it runs through each of the sheets in the workbook
For Each sh In ThisWorkbook.Sheets

'Setting the new column headers
sh.Range("I1") = "Ticker"
sh.Range("J1") = "Yearly Change"
sh.Range("K1") = "Percent Change"
sh.Range("L1") = "Total Stock Volume"

'Find the number of rows in column A
n = sh.Range("A2", sh.Range("A2").End(xlDown)).Rows.Count

'Array including the first value of each group of tickers
Dim first() As Variant
ReDim first(1 To n)
first(1) = 2

'Define the new row index for the new calculated columns
Dim NewRow As Integer
NewRow = 2

'Yearly change
Dim YrCh As Variant
'ReDim YrCh(1 To n2)

'Percent Change
Dim PerCh As Variant
'ReDim PerCh(1 To n2)

'Define the total sum of stocks
Dim TotalSum As Double
TotalSum = 0



For i = 2 To n
        'If the row above is not the same..
        If sh.Cells(i + 1, 1) <> sh.Cells(i, 1) Then

        'Define the ticker
        sh.Cells(NewRow, 9) = sh.Cells(i, 1)

        'Update array with index of first of set of tickers
        first(NewRow) = i + 1
        
        'Yearly change calculation and color coding
        YrCh = sh.Cells(i, 6) - sh.Cells(first(NewRow - 1), 3)
        sh.Cells(NewRow, 10) = YrCh
        If YrCh < 0 Then
        sh.Cells(NewRow, 10).Interior.ColorIndex = 3
        Else: sh.Cells(NewRow, 10).Interior.ColorIndex = 4
        End If
            
        'Percent change calculation
        PerCh = YrCh / sh.Cells(first(NewRow - 1), 3)
        sh.Cells(NewRow, 11) = Format(PerCh, "Percent")
        If PerCh < 0 Then
        sh.Cells(NewRow, 11).Interior.ColorIndex = 3
        Else: sh.Cells(NewRow, 11).Interior.ColorIndex = 4
        End If
                
        'Total sum calculation
        TotalSum = TotalSum + sh.Cells(i, 7)
        sh.Cells(NewRow, 12) = TotalSum
                
        NewRow = NewRow + 1
        
        
        TotalSum = 0
        
        Else
        TotalSum = TotalSum + sh.Cells(i, 7)
        End If

Next i


'Setting the new column headers
sh.Range("O2") = "Greatest % increase"
sh.Range("O3") = "Greatest % decrease"
sh.Range("O4") = "Greatest total volume"
sh.Range("P1") = "Ticker"
sh.Range("Q1") = "Value"

'Find max and min for percent change as well as max for total stock volume
sh.Range("Q2") = Format(WorksheetFunction.Max(sh.Range("K2:K" & NewRow)), "Percent")
sh.Range("P2") = Range("I" & WorksheetFunction.Match(sh.Range("Q2"), sh.Range("K1:K" & NewRow), 0))
sh.Range("Q3") = Format(WorksheetFunction.Min(sh.Range("K2:K" & NewRow)), "Percent")
sh.Range("P3") = Range("I" & WorksheetFunction.Match(sh.Range("Q3"), sh.Range("K1:K" & NewRow), 0))
sh.Range("Q4") = WorksheetFunction.Max(sh.Range("L2:L" & NewRow))
sh.Range("P4") = Range("I" & WorksheetFunction.Match(sh.Range("Q4"), sh.Range("L1:L" & NewRow), 0))


Next sh
End Sub


```

