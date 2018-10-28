Attribute VB_Name = "Module1"
Sub SheetLoop()

Dim i As Integer
Dim ws_num As Integer

Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet
ws_num = ThisWorkbook.Worksheets.Count

For i = 1 To ws_num
    ThisWorkbook.Worksheets(i).Activate
    Call StockData
Next

starting_ws.Activate


End Sub

Sub StockData()
Dim lRow As Long
Dim lRow2 As Long
Dim VolSubtotal As Double
Dim lOpenDate As Long
Dim lCloseDate As Long
Dim OpenValue As Double
Dim CloseValue As Double
Dim YearlyChange As Double
Dim PercentChange As Double

Dim TickerCol, TotalStockVolumeCol, YearlyChangeCol, PercentChangeCol As String

Dim OpenValueAsiigned As Integer

TickerCol = "I"
YearlyChangeCol = "J"
PercentChangeCol = "K"
TotalStockVolumeCol = "L"

Cells(1, TickerCol) = "Ticker"
Cells(1, TotalStockVolumeCol) = "Total Stock Volume"
Cells(1, YearlyChangeCol) = "Yearly Change"
Cells(1, PercentChangeCol) = "Percent Change"



lRow = 2
lRow2 = 2
VolSubtotal = 0
OpenValueAsiigned = 0


Do While (Cells(lRow, 1) <> "")

    If Cells(lRow, "A") = Cells(lRow + 1, "A") Then
          VolSubtotal = VolSubtotal + Cells(lRow, "G")
          If OpenValueAsiigned = 0 Then
            OpenValue = Cells(lRow, "C")
            OpenValueAsiigned = 1
          End If
    Else
        

        Cells(lRow2, "I") = Cells(lRow, "A")
        VolSubtotal = VolSubtotal + Cells(lRow, "G")
        Cells(lRow2, TotalStockVolumeCol) = VolSubtotal
        VolSubtotal = 0

        CloseValue = Cells(lRow, "F")

        YearlyChange = CloseValue - OpenValue
        Cells(lRow2, YearlyChangeCol) = YearlyChange
        If OpenValue <> 0 Then
            PercentChange = (YearlyChange * 100) / OpenValue
            Cells(lRow2, PercentChangeCol) = PercentChange
        Else
            Cells(lRow2, PercentChangeCol) = 0
        End If

        OpenValueAsiigned = 0 'reset

        lRow2 = lRow2 + 1
    End If

    lRow = lRow + 1

Loop



lRow2 = 2
Do While (Cells(lRow2, YearlyChangeCol) <> "")
    If Cells(lRow2, YearlyChangeCol) > 0 Then
        Cells(lRow2, YearlyChangeCol).Interior.ColorIndex = 3
    Else
        Cells(lRow2, YearlyChangeCol).Interior.ColorIndex = 4
    End If
    lRow2 = lRow2 + 1
Loop


Dim GreatestLabelsCol, GreatestTickerCol, GreatestValueCol As String

GreatestLabelsCol = "O"
GreatestTickerCol = "P"
GreatestValueCol = "Q"

Cells(1, GreatestTickerCol) = "Ticker"
Cells(1, GreatestValueCol) = "Value"

Cells(2, GreatestLabelsCol) = "Greatest % Increase"
Cells(3, GreatestLabelsCol) = "Greatest % Decrease"
Cells(4, GreatestLabelsCol) = "Greatest Total Volume"


Cells(2, GreatestValueCol).Formula = "=MAX(K:K)"
Cells(3, GreatestValueCol).Formula = "=MIN(K:K)"
Cells(4, GreatestValueCol).Formula = "=MAX(L:L)"

Cells(2, GreatestTickerCol).Formula = "=INDEX(I:I,MATCH(MAX(K:K),K:K,0))"
Cells(3, GreatestTickerCol).Formula = "=INDEX(I:I,MATCH(MIN(K:K),K:K,0))"
Cells(4, GreatestTickerCol).Formula = "=INDEX(I:I,MATCH(MAX(L:L),L:L,0))"




 
End Sub


