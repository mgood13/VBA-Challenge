Attribute VB_Name = "Module1"
Sub stockAnalysis():
Dim finalrow As Long
Dim tickers() As String
Dim length As Integer
Dim summaryrow As Integer
Dim openval As Double
Dim closeval As Double
Dim yrchange As Double
Dim colorset As Integer
'Total Stock Volume is far too large for long
Dim tsv As Double







' Find the final row number
finalrow = Cells(Rows.Count, 1).End(xlUp).Row


'Set up the summary table and autofits to ensure you can read the titles
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Range("I1:L1").Columns.AutoFit


'Summary row iterates every time there is a new ticker value
summaryrow = 2
tsv = 0

'It turns out you can't re-size an array that was initialized with a length
ReDim Preserve tickers(0)

length = UBound(tickers, 1) - LBound(tickers, 1) + 1

For i = 2 To finalrow

    'If statement that sets up each array with the correct values initially
    If i = 2 Then
        tickers(0) = Cells(i, 1)
        openval = Cells(i, 3).Value
    End If

    'If statement that runs when we reach the end of a ticker value
    If Cells(i, 1).Value <> tickers(length - 1) Then
        closeval = Cells(i - 1, 6).Value
        yrchange = closeval - openval

        If yrchange > 0 Then
            colorset = 4

        'This seems unlikely but just in case
        ElseIf yrchange = 0 Then
            colorset = 6
        Else
            colorset = 3
        End If


        'Populate summary row
        Cells(summaryrow, 9).Value = tickers(length - 1)
        Cells(summaryrow, 10).Value = yrchange
        Cells(summaryrow, 10).Interior.ColorIndex = colorset
        Cells(summaryrow, 11).Value = Abs(yrchange) / openval
        'Format that cell as a percentage
        Range("K" & summaryrow).NumberFormat = "0.00%"
        Cells(summaryrow, 12).Value = tsv



        'Iterate things for the next ticker value
        'tsv is given the value of the first row for the new ticker value
        tsv = Cells(i, 7).Value
        openval = Cells(i, 3).Value
        ReDim Preserve tickers(length)
        summaryrow = summaryrow + 1

        'Update the length of the array and then add the next ticker value in it
        'This prevents an infinite loop caused by placing the new ticker in tickers(0), not that I would know...
        length = UBound(tickers, 1) - LBound(tickers, 1) + 1
        tickers(length - 1) = Cells(i, 1).Value


    Else
        'If a new ticker hasn't been reached yet keep adding to total stock volume
        tsv = tsv + Cells(i, 7).Value

    End If




Next i





End Sub
