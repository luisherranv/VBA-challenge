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


'Create Ticker Symbol Column

'Find the number of rows in column A
n = sh.Range("A2", sh.Range("A2").End(xlDown)).Rows.Count

'Index for new ticker column
Dim k As Long

'Index for A colum
Dim i As Long

'Array including the first value of each group of tickers
Dim first() As Variant
ReDim first(1 To n)

'Setting up the first value of the new tickers column
k = 2
sh.Cells(k, 9) = sh.Cells(2, 1)
k = k + 1
first(1) = 2

'Loop to pull first ticker of a repetition into a new column
For i = 2 To n
    'If loop to ensure that if a ticker is already pulled, then move on to next
    If sh.Cells(i, 1) = sh.Cells(k - 1, 9) Then
    Else: sh.Cells(k, 9) = sh.Cells(i, 1)
            first(k - 1) = i
            k = k + 1
    End If
            
Next

'Find the number of rows in column I
n2 = sh.Range("I2", sh.Range("I2").End(xlDown)).Rows.Count

'Calculate Yearly change, Percent change and Total Stocks

'Yearly change
Dim YrCh As Variant
ReDim YrCh(1 To n2)

'Percent Change
Dim PerCh As Variant
ReDim PerCh(1 To n2)

'Total Stocks Volume
Dim TotSum As Variant
ReDim TotSum(1 To n2)

'Range of new columbs
Dim rng1 As String

'First index for set of same Tickers
Dim FirstInd As Variant
'First index for NEXT set of same Tickers
Dim FirstIndNext As Variant
'Last index for set of same Tickers
Dim LastInd As Variant


For j = 1 To n2

            'Index calculations for first and last value of a group of tickers
            FirstInd = first(j) '
            'If loop created to ensure that the last value of the sequence is considered
            If j = n2 Then
            LastInd = n
            Else: FirstIndNext = first(j + 1) '
            LastInd = FirstIndNext - 1
            End If
            
    'Yearly change calculation and color coding
    YrCh(j) = sh.Cells(LastInd, 6) - sh.Cells(FirstInd, 3)
    sh.Cells(j + 1, 10) = YrCh(j)
    If YrCh(j) < 0 Then
    sh.Cells(j + 1, 10).Interior.ColorIndex = 3
    Else: sh.Cells(j + 1, 10).Interior.ColorIndex = 4
    End If
    
    'Percent change calculation
    PerCh(j) = YrCh(j) / sh.Cells(FirstInd, 3)
    sh.Cells(j + 1, 11) = Format(PerCh(j), "Percent")
    If PerCh(j) < 0 Then
    sh.Cells(j + 1, 11).Interior.ColorIndex = 3
    Else: sh.Cells(j + 1, 11).Interior.ColorIndex = 4
    End If
    
    'Total stock calculation
    rng1 = "G" & FirstInd & ":G" & LastInd
    TotSum(j) = WorksheetFunction.Sum(sh.Range(rng1))
    sh.Cells(j + 1, 12) = TotSum(j)
Next

'Setting the new column headers
sh.Range("O2") = "Greatest % increase"
sh.Range("O3") = "Greatest % decrease"
sh.Range("O4") = "Greatest total volume"
sh.Range("P1") = "Ticker"
sh.Range("Q1") = "Value"

'Find max and min for percent change as well as max for total stock volume
sh.Range("Q2") = Format(WorksheetFunction.Max(PerCh), "Percent")
sh.Range("P2") = Range("I" & WorksheetFunction.Match(sh.Range("Q2"), sh.Range("K1:K" & n2), 0))
sh.Range("Q3") = Format(WorksheetFunction.Min(PerCh), "Percent")
sh.Range("P3") = Range("I" & WorksheetFunction.Match(sh.Range("Q3"), sh.Range("K1:K" & n2), 0))
sh.Range("Q4") = WorksheetFunction.Max(TotSum)
sh.Range("P4") = Range("I" & WorksheetFunction.Match(sh.Range("Q4"), sh.Range("L1:L" & n2), 0))


Next sh
End Sub
```

