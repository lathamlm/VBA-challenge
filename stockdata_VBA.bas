Attribute VB_Name = "Module1"
Option Explicit

Sub stock_analysis_final():

    'RUN ALL WORKSHEETS AT ONCE
    'REFERENCED STACKOVERFLOW 43738802
    '---------------------------------------------
    Dim ws_count As Integer
    Dim w As Integer
    
    ws_count = ActiveWorkbook.Worksheets.Count
    
    For w = 1 To ws_count Step 1
        ActiveWorkbook.Worksheets(w).Select
    '---------------------------------------------
    
        Dim row_count As Long
        Dim row_count_percent As Long
        Dim row_count_vol As Long
        Dim i As Long
        Dim summary_row As Long
        Dim open_price As Double
        Dim close_price As Double
        Dim year_change As Double
        Dim percent_change As Double
        Dim vol_open As String
        Dim vol_close As String
        Dim vol_range As Range
        Dim ticker_name As String
        Dim percent_first As String
        Dim percent_last As String
        Dim percent_range As Range
        Dim vol_first As String
        Dim vol_last As String
        Dim sum_vol_range As Range
       
           
        'FINDS LAST ROW WITH VALUES
        'REFERENCED STACKOVERFLOW 18088729
        '---------------------------------
        row_count = Range("A1").End(xlDown).Row
        '---------------------------------
        
        'DISCUSSED CONCEPT IN CLASS WITH EXAMPLES
        summary_row = 2
        
        'LABELING HEADERS
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'FORMATTING
        Columns("J:L").EntireColumn.AutoFit
        
        
        For i = 2 To row_count
            
            'IDENFYING FIRST TICKER INSTANCE
            If Not Cells(i, 1).Value = Cells(i - 1, 1).Value Then
                ticker_name = Cells(i, 1).Value
                open_price = Cells(i, 3).Value
                
                'IDENTIFIES CELL RANGE FOR SUM FUNCTION
                'REFERENCED STACKOVERFLOW 6262743
                '---------------------------------
                vol_open = Cells(i, 7).Address(rowAbsolute:=False, columnAbsolute:=False)
                '---------------------------------
                
            'IDENTIFYING LAST TICKER INSTANCE
            ElseIf Not Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                close_price = Cells(i, 6).Value
                
                'IDENTIFIES CELL RANGE FOR SUM FUNCTION
                'REFERENCED STACKOVERFLOW 6262743
                '---------------------------------
                vol_close = Cells(i, 7).Address(rowAbsolute:=False, columnAbsolute:=False)
                '---------------------------------
                
                'CALCULATIONS - YEAR AND PERCENT CHANGE
                year_change = close_price - open_price
                percent_change = year_change / open_price
                
                'IDENTIFYING VOLUME RANGE FOR SUM
                Set vol_range = Range(vol_open + ":" + vol_close)
                
                'SUMMARY TABLE
                Cells(summary_row, 9).Value = ticker_name
                Cells(summary_row, 10).Value = year_change
                
                'FORMATS YEAR CHANGE TO RED OR GREEN
                Cells(summary_row, 10).Select
                If year_change > 0 Then
                    Selection.Interior.Color = RGB(0, 255, 0)
                ElseIf year_change < 0 Then
                    Selection.Interior.Color = RGB(255, 0, 0)
                End If
                    
                'SUMMARY TABLE
                Cells(summary_row, 11).Value = percent_change
                
                'FORMATS CELL TO PERCENT
                Cells(summary_row, 11).Select
                Selection.Style = "Percent"
                Selection.NumberFormat = "0.00%"
                
                'SUMMARY TABLE
                Cells(summary_row, 12).Value = Application.WorksheetFunction.Sum(vol_range)
                
                'CHANGE SUMMARY ROW FOR UNIQUE LIST
                summary_row = summary_row + 1
            End If
        Next i
        
        
        'LABELING HEADERS
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        'FINDS LAST ROW WITH VALUES
        'REFERENCED STACKOVERFLOW 18088729
        '---------------------------------
        row_count_percent = Range("K1").End(xlDown).Row
        row_count_vol = Range("L1").End(xlDown).Row
        '---------------------------------
        
        'IDENTIFIES RANGE FOR PERCENT SUMMARY
        'REFERENCED STACKOVERFLOW 6262743
        '---------------------------------
        percent_first = Cells(2, 11).Address(rowAbsolute:=False, columnAbsolute:=False)
        percent_last = Cells(row_count_percent, 11).Address(rowAbsolute:=False, columnAbsolute:=False)
        '---------------------------------
        
        'FINDS GREATEST PERCENT INCREASE & DECREASE
        Set percent_range = Range(percent_first + ":" + percent_last)
        Cells(2, 17).Value = Application.WorksheetFunction.Max(percent_range)
        Cells(3, 17).Value = Application.WorksheetFunction.Min(percent_range)

        'FORMATTING TO PERCENT
        Range("Q2:Q3").Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
                          
        'IDENTIFIES RANGE FOR VOLUME SUMMARY
        'REFERENCED STACKOVERFLOW 6262743
        '---------------------------------
        vol_first = Cells(2, 12).Address(rowAbsolute:=False, columnAbsolute:=False)
        vol_last = Cells(row_count_vol, 12).Address(rowAbsolute:=False, columnAbsolute:=False)
        '---------------------------------
        
        'FINDS GREATEST VOLUME
        Set sum_vol_range = Range(vol_first + ":" + vol_last)
        Cells(4, 17).Value = Application.WorksheetFunction.Max(sum_vol_range)
    
        'FORMATTING
        Columns("O:O").EntireColumn.AutoFit
        Columns("Q:Q").EntireColumn.AutoFit
        
        'FINDS CORRESPONDING TICKER
        For i = 2 To row_count_vol
            If Cells(i, 11).Value = Cells(2, 17).Value Then
                Cells(2, 16).Value = Cells(i, 9).Value
            ElseIf Cells(i, 11).Value = Cells(3, 17).Value Then
                Cells(3, 16).Value = Cells(i, 9).Value
            ElseIf Cells(i, 12).Value = Cells(4, 17).Value Then
                Cells(4, 16).Value = Cells(i, 9).Value
            End If
        Next i
    Next w
        
End Sub
