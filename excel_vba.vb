
Sub Main()
'Main() executes a set of subroutines to analyze stock prices;
'finds tickers with the highest positive/negative price changes 
'and highest volume for each year presented in the Multi_year_stock_data.xlsm workbook

'Count the number of Worksheets in the Workbook
s = Application.Sheets.count
'Call ClearRange(s) subroutine to clear columns I-P of each worksheet
Call ClearRange(s)
'Call CreatHeadline(s) to create headlines for the new columns
Call CreateHeadline(s)
'Call GreatestValuesTable(s) subroutine to create the summary table in each worksheet
Call GreatestValuesTable(s)
'Call CountTickers(s) to create a list of unique tickers in each worksheet, 
'their respective year price changes and total volumes
Call CountTickers(s)
'Call Formatting(s) to set the cells' background color in "Percent change" column
'according to the value (negative = red, positive =green)
Call Formatting(s)

'Find the stock ticker with the highest volume for each year
Call FindMaxAllSheets(s, "L:L", "O4", "P4")
'Find the stock ticker with the highest price increase for each year
Call FindMaxAllSheets(s, "K:K", "O2", "P2")
'Find the stock ticker with the largest price decrease for each year
Call FindMinAllSheets(s, "K:K", "O3", "P3")


End Sub

Sub ClearRange(s)
'ClearRange() loops over to activate and clear all worksheets in the specified range
For i = 1 To s
    With Worksheets(i)
        .Activate
        .Range("I:L").Clear
    End With
Next i

End Sub

Sub CreateHeadline(s)

'CreateHeadline() Loops over all worksheets, activates them and creates new headlines
'"Ticker", "yearly change", "Percent Change" and "Total Stock Volume"
    For i = 1 To s
        Worksheets(i).Activate
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
    Next i

End Sub


Sub GreatestValuesTable(s)

' GretestValuesTable() creates an empty table for the greatest values first worksheet:
'-----------------------------------------------------------------
'                      | Ticker      Value
'----------------------|-----------------------------------------
'Greatest % Increse    |
'Greatest % Decrease   |
'Gretest total volume  |

For i = 1 To s
'Loop over ith worksheet
    Worksheets(i).Activate
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
Next i

End Sub

Sub CountTickers(s)

'CountTickers() fills out the new columns I and L ("ticker" and "total volume") in each sheet
'-Loops over all worksheets
'-creates a new column (column I) of unique tickers from column A
'-counts the total volume of each unique ticker and assigns it to the corresponding cells of column I

'Declare variables
Dim i As Double
Dim ws As Double
Dim count As Double
Dim lastrow As Double

For ws = 1 To s

'Activate worksheet(ws)
Worksheets(ws).Activate

'assign count to 1
count = 1

'Determine number of rows in column A
lastrow = Cells(Rows.count, 1).End(xlUp).Row

' Loop over all the rows in the worksheet
    For i = 2 To lastrow
    
        'If current ticker value (column A) is not equal to the previous ticker value,
        '-Assign this ticker value to a new cell in column I
        '-Assign volume corresponding to this ticker/row to the new cell in column L
        '-Calculate the year's opening price for the stock
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            count = count + 1
            Cells(count, 9).Value = Cells(i, 1)
            Cells(count, 12).Value = Cells(i, 7).Value
            yearstart = Cells(i, 3).Value
        
        'If current ticker value (column A) is not equal to the next ticker value,
        '-Extract the closing price of the stock a the end of the year
        '-Calculate the difference between the year's stock opening and closing prices
        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            yearend = Cells(i, 6).Value
            price_change = yearend - yearstart
                If yearstart = 0 Or price_change = 0 Then
                        Cells(count, 10).Value = 0
                        Cells(count, 11).Value = 0
                        
                Else
                        Cells(count, 10).Value = price_change
                        Cells(count, 11).Value = price_change / yearstart
                
                End If
            
            
        'If current ticker value is equal to the previous, add the volume from column G
        'to the total ticker volume in Column L
        ElseIf Cells(i, 1).Value = Cells(i - 1, 1).Value Then
            Cells(count, 12).Value = Cells(count, 12).Value + Cells(i, 7).Value
        
        End If
    
    Next i
    
Next ws


End Sub



Sub FindMaxAllSheets(s, rng, ticpos, valpos)

'FindMaxAllSheets() finds maximum values in every sheet,
'by looping over s number of worksheets
'rng - range where to look for the maximum value
'ticpos - is a destination cell for a ticker corresponding to the maximum value
'valpos - destination cell for the maximum value

'Declare variables
Dim i As Integer
Dim maxval As Double
Dim mx As Double
Dim pos As Double
Dim ticker As String

'Set the inital maximum value to 0
maxval = 0

'Loop over all worksheets
    For i = 1 To s
    
    'Activate i-th worksheet
    Worksheets(i).Activate
    'Find the maximum value in range rng
    mx = Application.WorksheetFunction.Max(Range(rng))
    
        'Determine if the maximum value from sth sheet is larger then the maxval
        If mx > maxval Then
            'If the condition is met, update the maximum value (maxval)
            maxval = mx
            'Find the row number corresponding to the maximum value
            pos = Application.WorksheetFunction.Match(maxval, Range(rng), 0)
            'Find the ticker correcsponding to the maximum value
            ticker = Cells(pos, 9).Value
            
        End If
    'Assign found maximum value and ticker to their specified locations in the summary table
    Range(ticpos).Value = ticker
    Range(valpos).Value = maxval
    Next i


End Sub

Sub FindMinAllSheets(s, rng, ticpos, valpos)

'FindMinAllSheets() finds minimum values in every sheet,
'by looping over s number of worksheets
'rng - range where to look for the minimum value
'ticpos - is a destination cell for a ticker corresponding to the minimum value
'valpos - destination cell for the minimum value

Dim minval As Double
Dim mn As Double
Dim pos As Double
Dim ticker As String

'Set the inital minimum value to 0
minval = 0

    'Loop over all worksheets
    For i = 1 To s
    
    'Activate i-th worksheet
    Worksheets(i).Activate
    'Find the minimum value in range rng
    mn = Application.WorksheetFunction.Min(Range(rng))
        
        'Determine if the maximum value from sth sheet is larger then the maxval
        If mn < minval Then
        'If the condition is met, update the minimum value (maxval)
            minval = mn
            'Find the row number corresponding to the minimum value
            pos = Application.WorksheetFunction.Match(minval, Range(rng), 0)
            'Find the ticker correcsponding to the minimum value
            ticker = Cells(pos, 9).Value
            
        End If
    
    Range(ticpos).Value = ticker
    Range(valpos).Value = minval
    
    Next i

End Sub




Sub Formatting(s)

'Formatting() subroutine formats cells in column K (Percent change) and the summary table in Worksheet(1):
'-Sets columns' I-P widths to AutoFit
'-Cells with positive change are colored green
'-Cells with the negative change are colored red
'-Cell number format in "Total Persent change" is set to %
Dim i As Double
Dim ws As Double
Dim ticlastrow As Double

' Formatting of summary table

'Looping over worksheets
For ws = 1 To s

    With Worksheets(ws)
        .Activate
        .Columns("I:P").AutoFit
        End With
    
    'Set format for the Greatest % decrease/encrease value cells
    Range("P2:P3").NumberFormat = "0%"
    'Set number format of the Total Volume value cell
    Range("P4").NumberFormat = "General"

    'Determine the number of unique ticker cells
    ticlastrow = Cells(Rows.count, 9).End(xlUp).Row
        'Loop over all unique tickers
        For i = 2 To ticlastrow
            'If cell value is positive, color it green
            If Cells(i, 11).Value > 0 Then
                Cells(i, 11).Interior.ColorIndex = 4
            'Otherwise, color it red
            Else: Cells(i, 11).Interior.ColorIndex = 3
            
            End If
        Next i
    Next ws
'
End Sub

