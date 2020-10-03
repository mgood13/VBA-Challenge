Attribute VB_Name = "Module11"

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

Dim sheetsnum As Integer

Dim greatincrease As Double

Dim greatdecrease As Double

Dim greattotal As Double

Dim percentchange As Double

Dim greatest(2) As String


'If you run the code multiple times this turns off the warning message
'that pops up when deleting the first sheet
Application.DisplayAlerts = False





If Sheets(1).Name = "Summary Sheet" Then
    Sheets("Summary Sheet").Delete
End If



'Obtain number of sheets for the current workbook

sheetsnum = Sheets.Count



'Cycle through all of the sheets in the workbook

For q = 1 To sheetsnum

    Sheets(q).Activate



    ' Find the final row number

    finalrow = Cells(Rows.Count, 1).End(xlUp).Row





    'Set up the summary table and autofits to ensure you can read the titles

    Cells(1, 9).Value = "Ticker"

    Cells(1, 10).Value = "Yearly Change"

    Cells(1, 11).Value = "Percent Change"

    Cells(1, 12).Value = "Total Stock Volume"

    Range("I1:L1").Columns.AutoFit

    'Can't forget about the greatest table
    Cells(1, 15).Value = "Tickers"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"





    'Summary row iterates every time there is a new ticker value

    summaryrow = 2

    tsv = 0

    greatincrease = 0

    greatdecrease = 0

    greattotal = 0

    greatest(0) = ""
    greatest(1) = ""
    greatest(2) = ""


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



            'Ensure that there are no issues with dividing by zero if openval and closeval = 0

            If openval <> 0 Then

                percentchange = yrchange / openval

            Else

                percentchange = 0

            End If

            Cells(summaryrow, 11).Value = percentchange


            'Format that cell as a percentage

            Range("K" & summaryrow).NumberFormat = "0.00%"

            Cells(summaryrow, 12).Value = tsv


            'The Greatest Comparisons

            'Greatest Increase
            If percentchange > greatincrease Then
                greatincrease = percentchange
                greatest(0) = tickers(length - 1)
            End If

            'Greatest Decrease
            If percentchange < greatdecrease Then
                greatdecrease = percentchange
                greatest(1) = tickers(length - 1)
            End If

            'Greatest Total
            If tsv > greattotal Then
                greattotal = tsv
                greatest(2) = tickers(length - 1)
            End If




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

    'Populate the Greatest Table
    Cells(2, 15).Value = greatest(0)
    Cells(2, 16).Value = greatincrease
    Range("P2").NumberFormat = "0.00%"

    Cells(3, 15).Value = greatest(1)
    Cells(3, 16).Value = greatdecrease
    Range("P3").NumberFormat = "0.00%"

    Cells(4, 15).Value = greatest(2)
    Cells(4, 16).Value = greattotal
    Range("N1:P4").Columns.AutoFit


Next q



'Reactivate the first sheet in the Workbook

Sheets(1).Activate

'Create the new summary Sheet before the first sheet of data
Sheets.Add Before:=Sheets(1)
Sheets(1).Activate

'Name the sheet Summary sheet
ActiveSheet.Name = "Summary Sheet"

'Loop through each sheet of data
'Place the name of the sheet into the given cell in the summary sheet
'Place the "Greatest Table" from each page below its sheetname title
'Bold each of the sheet names just for fun

For j = 1 To sheetsnum
    If j = 1 Then
        Cells(1, 1).Value = Sheets(j + 1).Name
        Range("A1:A1").Font.Bold = True
        Sheets(j + 1).Range("N1:P4").Copy Range("A2:C5")

    Else
        Cells(((j - 1) * 6) + 1, 1).Value = Sheets(j + 1).Name
        Range("A" & ((j - 1) * 6) + 1 & ":A" & ((j - 1) * 6) + 1 & "").Font.Bold = True
        Sheets(j + 1).Range("N1:P4").Copy Range("A" & (j - 1) * 6 + 2 & ":C" & (j - 1) * 6 + 2 & "")

     End If


Next j

'Autofit the columns to make sure that the data is visible
'Don't make the user click things like a barbarian

Range("A1:C" & sheetsnum * 6 & "").Columns.AutoFit






End Sub

