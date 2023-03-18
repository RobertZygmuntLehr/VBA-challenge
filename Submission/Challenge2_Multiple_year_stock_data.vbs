Attribute VB_Name = "Stock_Data"
' NOTES:
' ¥ LEGEND: In the comments, brackets enwrap the [modules], [sub-routines], [objects], [conditionals], [loops], [functions], and single quotes enwrap worksheet 'names'.

Sub Stock_Data()
    
    ' --------------------------------------------
    ' INSTANTIATING VARIABLES - START
    ' --------------------------------------------
    Dim row_Variable As Long ' This object stores the current row index.
    Dim counter_Variable As Long  ' This object stores the current index, separate from [row_Variable].
    Dim lastRow_Variable As Long ' This object stores the index of the last row with a non-null value from column "A".
    Dim openValue_Variable As Double ' This object stores the current <open> value from column "C".
    Dim closeValue_Variable As Double ' This object stores the current <close> value from column "F".
    Dim totalStockVolume_Variable As Double ' This object stores the current <vol> value summation from column "G" for the same <ticker> value.
    Dim ticker_Variable As String ' This object stores the current <ticker> value from column [A].
    Dim greatestPercentIncrease_Variable As Double ' This object stores the calculation of the greatest positive <Percent Change> value from column "K".
    Dim greatestPercentDecrease_Variable As Double ' This object stores the calculation of the greatest negative <Percent Change> value from column "K".
    Dim greatestTotalVolume_Variable As Double ' This object stores the calculation of the greatest <Total Stock Volume> value from column "L".
    Dim ws As Worksheet 'This object stores current Worksheet.
    Dim rng As Range
    ' --------------------------------------------
    ' INSTANTIATING VARIABLES - END
    ' --------------------------------------------
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS - START
    ' --------------------------------------------
    For Each ws In ThisWorkbook.Worksheets
         
        ws.Activate ' Activates the new worksheet in the workbook.

        ' -------------------------------------------
        ' INITIALIZING VARIABLES - START
        ' --------------------------------------------
        row_Variable = 2 ' [row_Variable] is initialized with the value of 2.
        counter_Variable = 2 ' [counter_Variable] is initialized with the value of 2.
        greatestTotalVolume_Variable = 0 ' [greatestTotalVolume_Variable] is initialized with the value of 0.
        greatestPercentIncrease_Variable = 0 ' [greatestNegativeChange_Variable] is initialized with the value of 0.
        greatestPercentDecrease_Variable = 0 ' [greatestPositiveChange_Variable] is initialized with the value of 0.
        ticker_Variable = Range("A" & row_Variable).Value ' [ticker_Variable] is initialized with the value of cell "A2".
        lastRow_Variable = Range("A" & Rows.Count).End(xlUp).Row ' [lastRow_Variable] is initialized to be the last non-null value from column "A".
        ' NOTE:
        ' Rows.Count returns the index of the last row in the current worksheet.
        ' .End(xlUp).Row returns the index of the last row with a non-null value for rows less than Rows.Count in the current worksheet
        ' --------------------------------------------
        ' INITIALIZING VARIABLES - END
        ' --------------------------------------------
        
        Set rng = Range("H1:Q" & lastRow_Variable) ' The output range.
        rng.ClearContents ' Clear cells in ouput range.
        
        ' Assigning values to header cells and identification cells.
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
            
            ' --------------------------------------------
            ' LOOP THROUGH THE DATA WITHIN ACTIVE SHEET- START
            ' --------------------------------------------
            For row_Variable = 2 To lastRow_Variable ' This [For] Loop iterates from row 2 to row [lastRow_Variable].
   
                ticker_Variable = Range("A" & row_Variable).Value ' [ticker_Variable] is assigned the value of the cell in column "A" at row [row_Variable].
                
                openValue_Variable = Range("C" & row_Variable).Value ' [openValue_Variable] is assigned the value of the cell in column "C" at row [row_Variable].

                Range("I" & counter_Variable).Value = Range("A" & row_Variable).Value ' Assigns the <ticker> value from column "A"  at row [row_Variable] to "I" column at row [counter_Variable].
   
                totalStockVolume_Variable = Range("G" & row_Variable).Value ' [totalStockVolume_Variable] is initialized with the value of the cell in column [G] and row [row_Variable].
   
                While Range("A" & row_Variable).Value = Range("A" & row_Variable + 1).Value ' This [While] loop iterates as long as the value of the cell in the next row is the same as the current row.
        
                    row_Variable = row_Variable + 1 ' Increases the value of [row_Variable] by 1.
            
                    totalStockVolume_Variable = totalStockVolume_Variable + Range("G" & row_Variable).Value ' Adds the value of the current cell in column "G" to the value of [totalStockVolume_Variable].
      
                    If greatestTotalVolume_Variable < totalStockVolume_Variable Then ' Checks to see if [greatestTotalVolume_Variable] is less than [totalStockVolume_Variable].
        
                        greatestTotalVolume_Variable = totalStockVolume_Variable ' If true, [greatestTotalVolume_Variable] is assigned the value of [totalStockVolume_Variable].
            
                        Range("P4").Value = Range("I" & counter_Variable).Value ' If true, the cell "P4" is assigned the <ticker> value in column "I" for the respective [greatestTotalVolume_Variable] <ticker>.
            
                    End If ' Closes the If statement.
            
                Wend ' Closes the while loop.
        
                closeValue_Variable = Range("F" & row_Variable).Value ' [openValue_Variable] is assigned the value of Range("C" & [row_Variable]).Value
        
                Range("J" & counter_Variable).Value = closeValue_Variable - openValue_Variable ' comment
        
                Range("K" & counter_Variable).Value = FormatPercent((closeValue_Variable - openValue_Variable) / openValue_Variable) ' Calculate Percent change
        
              ' Calcuations for change in [greatestPercentDecrease_Variable].
                If Range("K" & counter_Variable).Value < greatestPercentDecrease_Variable Then
                    greatestPercentDecrease_Variable = Range("K" & counter_Variable).Value
                    Range("P3").Value = Range("I" & counter_Variable).Value
                End If
        
                ' Calcuations for change in [greatestPercentIncrease_Variable].
                If Range("K" & counter_Variable).Value > greatestPercentIncrease_Variable Then
                    greatestPercentIncrease_Variable = Range("K" & counter_Variable).Value
                    Range("P2").Value = Range("I" & counter_Variable).Value
                End If
        
                If Range("K" & counter_Variable).Value < 0 Then ' Checks to see if the percent change is negative.
            
                    Range("J" & counter_Variable).Interior.ColorIndex = 3 ' 3 is the index for red color
        
                    ElseIf Range("K" & counter_Variable).Value > 0 Then ' Checks to see if the percent change is positive.

                        Range("J" & counter_Variable).Interior.ColorIndex = 4 ' 4 is the index for green color

                    Else ' Runs the following code if the percent change is zero

                        Range("J" & counter_Variable).Interior.ColorIndex = 6 ' 6 is the index for yellow color
                
                End If
        
                Range("L" & counter_Variable).Value = totalStockVolume_Variable

                counter_Variable = counter_Variable + 1  ' Advances the [counter_Variable] by 1. This ensures the value in the next iteration is recorded in the row index = [counter_Variable].
                
            Next row_Variable ' Ends the [For] Loop, increasing the value of [row_Variable] by 1.
    ' --------------------------------------------
    ' LOOP THROUGH THE DATA WITHIN ACTIVE SHEET - END
    ' --------------------------------------------
    
            ' Returns the the greatest value calculations.
            Range("Q2").Value = FormatPercent(greatestPercentIncrease_Variable)
            Range("Q3").Value = FormatPercent(greatestPercentDecrease_Variable)
            Range("Q4").Value = greatestTotalVolume_Variable
            
            rng.EntireColumn.AutoFit ' autofits column widths to data.
    
    Next ws
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS - END
    ' --------------------------------------------

' Ends the sub-routine
End Sub
