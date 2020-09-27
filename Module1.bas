Attribute VB_Name = "Module1"
Sub stockAnalysis():
Dim finalrow As Long
Dim tickers() As String
Dim length As Integer
Dim summaryrow As Integer
Dim openval As Double
Dim closeval As Double
Dim yrchange As Double
Dim tsv As Double





' Find the final row number
finalrow = Cells(Rows.Count, 1).End(xlUp).Row


'Set up the summary table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

summaryrow = 2
tsv = 0
ReDim Preserve tickers(0)

length = UBound(tickers, 1) - LBound(tickers, 1) + 1

For i = 2 To finalrow
    If i = 2 Then
        tickers(0) = Cells(i, 1)
        openval = Cells(i, 3).Value
    End If
    
    
    If Cells(i, 1).Value <> tickers(length - 1) Then
        closeval = Cells(i - 1, 6).Value
        yrchange = closeval - openval
        
        'Populate summary row
        Cells(summaryrow, 9).Value = tickers(length - 1)
        Cells(summaryrow, 10).Value = yrchange
        Cells(summaryrow, 11).Value = Abs(yrchange) / openval
        Cells(summaryrow, 12).Value = tsv
    
    
    
        'Iterate things for the next ticker value
        tsv = 0
        openval = Cells(i, 3).Value
        ReDim Preserve tickers(length)
        summaryrow = summaryrow + 1
        
        
        length = UBound(tickers, 1) - LBound(tickers, 1) + 1
        tickers(length - 1) = Cells(i, 1).Value
    
    
    Else
        tsv = tsv + Cells(i, 7).Value
        
    End If
    



Next i












End Sub
