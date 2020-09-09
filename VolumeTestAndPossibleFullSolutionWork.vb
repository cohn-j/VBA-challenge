Sub VolumeTest()
    Dim LastRow As LongLong
    LastRow = Range("A" & Rows.Count).End(xlUp).Row
    'MsgBox (LastRow)
    Dim Ticker As String
    Dim Ticker2 As String
    Dim Vol As LongLong
    Dim Header(4) As String
    Dim UniqueTickers As Integer

'Establish Table for Stock Perfomance Evaluation
    Header(0) = "Ticker"
    Header(1) = "Yearly Change"
    Header(2) = "Percent Change"
    Header(3) = "Total Stock Volume"

    Range("J1:M1") = Header()

    For Z = 10 To 13
        Cells(1, Z).Font.Bold = True
    Next Z
    
    UniqueTickers = 1

'This For loop is to determine the number of unique stock tickers and effectively paste them in column
'10 of the active worksheet.
    For x = 2 to LastRow
        Ticker = Cells(x,1)
        Ticker2 = Cells(x + 1, 1)
        If Ticker <> Ticker2 Then
            UniqueTickers = UniqueTickers + 1
            Cells(UniqueTickers, 10).Value = Ticker
        End If
    Next x

    UniqueTickers = 1

'This for loop calculates the total volume for each unique stock ticker.
    For i = 2 To LastRow
        Ticker = Cells(i, 1)
        Ticker2 = Cells(i + 1, 1)
        If Ticker = Ticker2 Then
            Vol = Cells(i, 7) + Vol
        ElseIf Ticker <> Ticker2 Then
    'Result posts the summation of all instances of a particular ticker:
            Vol = Cells(i, 7) + Vol
            UniqueTickers = UniqueTickers + 1
            Cells(UniqueTickers, 13).Value = Vol
            Vol = 0
        End If
    Next i
    'MsgBox (UniqueTickers)
End Sub
