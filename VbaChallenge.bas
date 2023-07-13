Attribute VB_Name = "Module1"
' Author: Roel De Los Santos
' Date: 7/13/2023
' Module 2 Challenge
' Sources Cited
' David Jaimes (Nov 22, 2019) yearly-stock-market-analysis  [vba]. https://github.com/davidjaimes/yearly-stock-market-analysis.

' Create control to loop though all worksheets
Sub VBA_Challenge()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        Call TickerCalc(ws)
    Next
End Sub
' Create control to loop within ws worksheet
Sub TickerCalc(ws As Worksheet)
    Dim i, Count, Lastrow As Long
    Dim yc As Double
    Dim MaxPT, MinPT, MVT As String
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    op = ws.Cells(2, 3).Value
    ' Declare counter
    Count = 2
    vt = 0
    ' Create and format headers
    With ws
        .Cells(1, 9).Value = "Ticker "
        .Cells(1, 10).Value = "Yearly Change "
        .Cells(1, 11).Value = "Percent Change "
        .Cells(1, 12).Value = "Total Stock Volume "
        .Cells(1, 16).Value = "Ticker "
        .Cells(1, 17).Value = "Values"
        .Cells(2, 15).Value = "Greatest % Increase "
        .Cells(3, 15).Value = "Greatest % Decrease "
        .Cells(4, 15).Value = "Greatest Total Volume "
        .Rows(1).Font.Bold = True
        .Rows(1).Columns.AutoFit
        .Columns("G").ColumnWidth = 10
        .Range("O:O, Q:Q").ColumnWidth = 21
        ' Set opening price for ticker
        op = ws.Cells(2, 3).Value
    
        For i = Count To Lastrow
            ' Get ticker yearly change
            If .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then
                .Cells(Count, 9).Value = .Cells(i, 1).Value
                cp = .Cells(i, 6).Value
                yc = cp - op
                .Cells(Count, 10).Value = yc
                ' Get percent change
                If op <> 0 Then
                    .Cells(Count, 11).Value = CStr((yc / op) * 100 & "%")
                End If
                ' Set conditionsl format
                If yc < 0 Then
                    .Cells(Count, 10).Interior.ColorIndex = 3
                    .Cells(Count, 11).Interior.ColorIndex = 3
                ElseIf yc > 0 Then
                    .Cells(Count, 10).Interior.ColorIndex = 4
                    .Cells(Count, 11).Interior.ColorIndex = 4
                End If
                ' Get max percentage
                If .Cells(Count, 11).Value > MaxP Then
                    MaxP = .Cells(Count, 11).Value
                    MaxPT = Cells(Count, 9).Value
                ' Get min percentage
                ElseIf .Cells(Count, 11).Value < MinP Then
                    MinP = .Cells(Count, 11).Value
                    MinPT = .Cells(Count, 9).Value
                End If
                ' Get max volume and ticker
                If .Cells(Count, 12).Value > MV Then
                    MV = .Cells(Count, 12).Value
                    MVT = .Cells(Count, 9).Value
                End If
                ' Set max volume
                vt = vt + .Cells(i, 7).Value
                .Cells(Count, 12).Value = vt
                ' Reset count and volume
                Count = Count + 1
                vt = 0
            Else
                vt = vt + ws.Cells(i, 7).Value
            End If
        
        Next i
        ' Print max, min percentage, volume and ticker
        .Cells(2, 17).Value = Format(MaxP, "#.##%")
        .Cells(2, 16).Value = MaxPT
        .Cells(3, 17).Value = Format(MinP, "#.##%")
        .Cells(3, 16).Value = MinPT
        .Cells(4, 17).Value = CStr(MV)
        .Cells(4, 16).Value = MVT
        
    End With

End Sub


