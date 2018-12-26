Attribute VB_Name = "Module1"
Sub ticker()

'Define and set variables

Dim ticker_name As String

Dim ws As Integer

Dim ws_name As Integer

ws_name = ThisWorkbook.Worksheets.Count

' Loop through sheets

    For ws = 1 To ws_name

    ThisWorkbook.Worksheets(ws).Activate

' Add headers to columns

    Cells(1, 9).Value = "Ticker"

    Cells(1, 10).Value = "Total Stock Volume"

' Set the tick variable to start at 0

Dim Ticker_Total As Double

Ticker_Total = 0

Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

' Define the last row

    lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

' For loop using the last row

    For I = 2 To lastrow

           If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
     

                ticker_name = Cells(I, 1).Value

    
                Ticker_Total = Ticker_Total + Cells(I, 7).Value

    
                Range("i" & Summary_Table_Row).Value = ticker_name

     

                Range("j" & Summary_Table_Row).Value = Ticker_Total

 

                Summary_Table_Row = Summary_Table_Row + 1

     

                Ticker_Total = 0

 

         Else
                            
                Ticker_Total = Ticker_Total + Cells(I, 7).Value

 
        End If

 
    Next I
 

Next ws
End Sub
