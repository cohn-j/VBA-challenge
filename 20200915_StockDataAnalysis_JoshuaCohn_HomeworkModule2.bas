Attribute VB_Name = "Module1"
Sub StockAnalysis()
    
'wks is the Worksheet type to meet the requirement of cycling through each tab of the workbook via the for each loop immediately following the declaration.
    Dim wks As Worksheet
        
   For Each wks In Worksheets

'All variables containing #s are LongLong based on errors I encountered as they can holder larger values than an Integer or Long type. The Tickers are String as they contain the Stock Tickers in the first column.
    Dim LastRow As LongLong
    Dim Ticker As String
    Dim Ticker2 As String
    Dim Vol As LongLong
    Dim Header(5) As String
    Dim UniqueTickers As LongLong
    Dim PriceTracker As LongLong
    Dim x As LongLong
    Dim y As LongLong

'LastRow is a variable to calculate the total # of rows in each tab of the workbook; this value is then used to determine how many times the for loop will run.
    LastRow = wks.Cells(Rows.Count, 1).End(xlUp).Row

'Establish the Column Headers for the table that contains the Stock Perfomance Evaluation
    Header(0) = "Ticker"
    Header(1) = "Yearly Change"
    Header(2) = "Percent Change"
    Header(3) = "Total Stock Volume"

    wks.Range("J1:M1") = Header()

'Bolds the column headers
    For Z = 10 To 13
    wks.Cells(1, Z).Font.Bold = True
    Next Z

'Initialize the variable values. UniqueTickers is used to determine the # of unique tickers in the initial data set. PriceTracker is used to determine the start of year price per stock.
    UniqueTickers = 1
    PriceTracker = 2

'This for loop contains all required items for the base homework:
    For x = 2 To LastRow
    'Ticker variables are to compare one row against the successive one:
        Ticker = wks.Cells(x, 1)
        Ticker2 = wks.Cells(x + 1, 1)
        'Only criteria where we are looking for the ticker on a row to be the same as the next is for calculating total volume so this If statement is to check that and if they are the same, add the number in the coordinate to the Vol (total volume) variable.
        If (Ticker = Ticker2) Then
            Vol = wks.Cells(x, 7) + Vol
        'If the tickers do not equal each other then the following items will occur:
        ElseIf (Ticker <> Ticker2) Then
        'Add the volume to the continuing volume variable:
            Vol = wks.Cells(x, 7) + Vol
            'Increase the # of unique tickers by one for our results table
            UniqueTickers = UniqueTickers + 1
            'Post the stock ticker in the Ticker column of the table:
            wks.Cells(UniqueTickers, 10).Value = Ticker
            'Post the volume in the Total Stock Volume column of the table:
            wks.Cells(UniqueTickers, 13).Value = Vol
            'Calculate the yearly price change and post it to the Yearly Change column of the table:
            wks.Cells(UniqueTickers, 11).Value = wks.Cells(x, 6).Value - wks.Cells(PriceTracker, 3).Value
            'Perform the conditional formatting check on the yearly change; if >0 the cell will be green, if <0 will be red, if not meeting those criteria it will remain white.
            If (wks.Cells(UniqueTickers, 11).Value > 0) Then
                wks.Cells(UniqueTickers, 11).Interior.ColorIndex = 4
            ElseIf (wks.Cells(UniqueTickers, 11).Value < 0) Then
                wks.Cells(UniqueTickers, 11).Interior.ColorIndex = 3
            Else
                wks.Cells(UniqueTickers, 11).Interior.ColorIndex = 2
            End If
            'change the cell formatting for the Percent Change results to be in 0.00% format.
            wks.Cells(UniqueTickers, 12).NumberFormat = "0.00%"
            'This If statement checks if the numerator and denominator are both 0 for the %age change calc:
            If wks.Cells(PriceTracker, 3).Value = 0 And wks.Cells(x, 6) = 0 Then
            'If both are 0, the %age change posted by 0.
                wks.Cells(UniqueTickers, 12).Value = 0
            ElseIf wks.Cells(PriceTracker, 3).Value = 0 And wks.Cells(x, 6) <> 0 Then
            'If the start of year price is 0 but the end of year price is not, set it to 1 (100%).
                wks.Cells(UniqueTickers, 12).Value = 1
            Else
            'If both the start of year and end of year prices are not 0, the posted %age change is (End Price - Start Price) / Start Price.
                wks.Cells(UniqueTickers, 12).Value = (wks.Cells(x, 6).Value - wks.Cells(PriceTracker, 3).Value) / wks.Cells(PriceTracker, 3).Value
            End If
            'Update the PriceTracker to the current value of x + 1 as that would be the start price for the next ticker.
            PriceTracker = x + 1
            'Reset volume to 0 as we are now going to start calculating the next stock ticker's volume.
            Vol = 0
        End If
    Next x

'This is to move the cursor back to the top after completion of each tab's review:
   Cells(1, 1).Select

'Start of challenge to calculate the greatest results within the base assignment results:
    Dim Greatest(3) As String
    Dim GVol As LongLong
    Dim GInc As Double
    Dim GDec As Double


    Header(4) = "Data"
    Greatest(0) = "Greatest % Increase"
    Greatest(1) = "Greatest % Decrease"
    Greatest(2) = "Greatest Total Volume"
    
'Assign the column headers and bold them:
    wks.Cells(1, 17) = Header(0)
    wks.Cells(1, 18) = Header(4)
    wks.Cells(1, 17).Font.Bold = True
    wks.Cells(1, 18).Font.Bold = True
    
'for loop to assign the row headers and bold them:
    For w = 2 To 4
    wks.Cells(w, 16).Value = Greatest(w - 2)
    wks.Cells(w, 16).Font.Bold = True
    Next w

'Variables to track the greatest volume # and ticker:
    GVol = wks.Cells(2, 13)
    GVol_T = wks.Cells(2, 10)

'Variables to track the greatest %age increase:
    GPerInc = wks.Cells(2, 12)
    GPerInc_T = wks.Cells(2, 10)

'Variables to track the greatest %age decrease:
    GPerDec = wks.Cells(2, 12)
    GPerDec_T = wks.Cells(2, 10)

'for loop to run through the results table:
 For y = 2 To UniqueTickers

 'If statement to compare a ticker's total volume to the one below it. If larger, the variables'
 'values are replaced, if not they stay the same and are compared to the next row's.
    If GVol < wks.Cells(y + 1, 13) Then
        GVol = wks.Cells(y + 1, 13)
        GVol_T = wks.Cells(y + 1, 10)
    Else
        GVol = GVol
        GVol_T = GVol_T
    End If
'If statement to compare a ticker's %age change to the one below it; if larger, the variables'
'values are replaced. If not they stay the same and are compared to the next row's.
    If GPerInc < wks.Cells(y + 1, 12) Then
        GPerInc = wks.Cells(y + 1, 12)
        GPerInc_T = wks.Cells(y + 1, 10)
    Else
        GPerInc = GPerInc
        GPerInc_T = GPerInc_T
    End If
'If statement to compare a ticker's %age change to the one below it; if smaller, the variables'
'values are replaced. If not they stay the same and are compared to the next row's.
If GPerDec > wks.Cells(y + 1, 12) Then
        GPerDec = wks.Cells(y + 1, 12)
        GPerDec_T = wks.Cells(y + 1, 10)
    Else
        GPerDec = GPerDec
        GPerDec_T = GPerDec_T
    End If

 Next y

'Print the greatest volume and its ticker:
    wks.Cells(4, 18).Value = GVol
    wks.Cells(4, 17).Value = GVol_T

'Print the greatest %age increase in % format and its ticker:
    wks.Cells(2, 18).NumberFormat = "0.00%"
    wks.Cells(2, 18).Value = GPerInc
    wks.Cells(2, 17).Value = GPerInc_T

'Prince the greatest %age decrease in % format and its ticker:
    wks.Cells(3, 18).NumberFormat = "0.00%"
    wks.Cells(3, 18).Value = GPerDec
    wks.Cells(3, 17).Value = GPerDec_T
 
  wks.Columns.AutoFit

 Next wks


    
End Sub
