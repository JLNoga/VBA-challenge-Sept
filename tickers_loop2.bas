Attribute VB_Name = "Module1"
Sub tickers_loop2()
    
'loop through tabs
Dim number_of_worksheets As Integer
number_of_worksheets = ActiveWorkbook.Worksheets.Count
Dim worksheet_number As Integer
For worksheet_number = 1 To number_of_worksheets
    ActiveWorkbook.Worksheets(worksheet_number).Select
        
    'define variables
    Dim lastrow As LongLong
    Dim row As LongLong
    Dim tickerrow As Integer
    Dim ticker As String
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearchange As Double
    Dim percentchange As Double
    Dim volume As LongLong
    Dim lastrow2 As LongLong
    Dim row2 As LongLong
    Dim maxrow As Integer
    Dim minrow As Integer
    Dim maxvol As Integer
    
    'set values
    lastrow = ActiveWorkbook.Worksheets(worksheet_number).Cells(Rows.Count, 1).End(xlUp).row
    openprice = ActiveWorkbook.Worksheets(worksheet_number).Cells(2, 3).Value
    tickerrow = ActiveWorkbook.Worksheets(worksheet_number).Cells(2, 10).row
    
    
    'create table
    ActiveWorkbook.Worksheets(worksheet_number).Cells(1, 10).Value = "Ticker"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(1, 11).Value = "Yearly Change"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(1, 12).Value = "Percentage Change"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(1, 13).Value = "Total Stock Volume"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(1, 16).Value = "Ticker"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(1, 17).Value = "Value"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(2, 15).Value = "Greatest % Increase"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(3, 15).Value = "Greatest % Decrease"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(4, 15).Value = "Greatest Total Volume"
    
    'loop through sheet
    For row = 2 To lastrow
        'look for a change in ticker
        If ActiveWorkbook.Worksheets(worksheet_number).Cells(row + 1, 1) <> ActiveWorkbook.Worksheets(worksheet_number).Cells(row, 1) Then
            'store ticker
            ticker = ActiveWorkbook.Worksheets(worksheet_number).Cells(row, 1).Value
            'store close price
            closeprice = ActiveWorkbook.Worksheets(worksheet_number).Cells(row, 6).Value
            'calculate yearly change
            yearchange = closeprice - openprice
            'calculate percentage change
            percentchange = yearchange / openprice
            'add to volume
            volume = ActiveWorkbook.Worksheets(worksheet_number).Cells(row, 7).Value + volume
            'save ticker
            ActiveWorkbook.Worksheets(worksheet_number).Cells(tickerrow, 10).Value = ticker
            'save yearly change
            ActiveWorkbook.Worksheets(worksheet_number).Cells(tickerrow, 11).Value = yearchange
            'save percentage change
            ActiveWorkbook.Worksheets(worksheet_number).Cells(tickerrow, 12).Value = percentchange
            'save stock volume
            ActiveWorkbook.Worksheets(worksheet_number).Cells(tickerrow, 13).Value = volume
            'reset open price
            openprice = ActiveWorkbook.Worksheets(worksheet_number).Cells(row + 1, 3).Value
            'reset ticker row
            tickerrow = tickerrow + 1
            'reset volume
            volume = 0
        Else  'if ticker does not match
            'add to volume
            volume = ActiveWorkbook.Worksheets(worksheet_number).Cells(row, 7).Value + volume
        End If
    Next row
    'find last row 2
    lastrow2 = ActiveWorkbook.Worksheets(worksheet_number).Cells(Rows.Count, 12).End(xlUp).row
    
    'Format percentage change as percent
    ActiveWorkbook.Worksheets(worksheet_number).Range("L:L").NumberFormat = "0.00%"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(2, 17).NumberFormat = "0.00%"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(3, 17).NumberFormat = "0.00%"
    'Format yaerly change as green for positive and red for negative
    For row2 = 2 To lastrow2
        If ActiveWorkbook.Worksheets(worksheet_number).Cells(row2, 11).Value >= 0 Then
            ActiveWorkbook.Worksheets(worksheet_number).Cells(row2, 11).Interior.ColorIndex = 4
        Else: ActiveWorkbook.Worksheets(worksheet_number).Cells(row2, 11).Interior.ColorIndex = 3
        End If
    Next row2
    'Format volume as number
    ActiveWorkbook.Worksheets(worksheet_number).Range("M:M").NumberFormat = "0,000"
    ActiveWorkbook.Worksheets(worksheet_number).Cells(4, 17).NumberFormat = "0,000"
    'select max increase (formula help found here:https://stackoverflow.com/questions/45422688/vba-for-loop-to-find-maximum-value-in-a-column)
    ActiveWorkbook.Worksheets(worksheet_number).Cells(2, 17).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow2))
    maxrow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("L2:L" & lastrow2)), Range("L2:L" & lastrow2), 0)
    ActiveWorkbook.Worksheets(worksheet_number).Cells(2, 16).Value = Range("J2:J" & lastrow2)(maxrow)
    'select min increase
    ActiveWorkbook.Worksheets(worksheet_number).Cells(3, 17).Value = Application.WorksheetFunction.Min(Range("L2:L" & lastrow2))
    minrow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(Range("L2:L" & lastrow2)), Range("L2:L" & lastrow2), 0)
    ActiveWorkbook.Worksheets(worksheet_number).Cells(3, 16).Value = Range("J2:J" & lastrow2)(minrow)
    'select max volume
    ActiveWorkbook.Worksheets(worksheet_number).Cells(4, 17).Value = Application.WorksheetFunction.Max(Range("M2:M" & lastrow2))
    maxvol = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("M2:M" & lastrow2)), Range("M2:M" & lastrow2), 0)
    ActiveWorkbook.Worksheets(worksheet_number).Cells(4, 16).Value = Range("J2:J" & lastrow2)(maxvol)
Next worksheet_number
End Sub

