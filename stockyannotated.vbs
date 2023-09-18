Attribute VB_Name = "Module2"
' Create Macro
Sub stocky():
Dim ws As Worksheet
For Each ws In Worksheets
'Assign all necessary rows and columns and define variables
ws.Range("I1").Value = "TICKER"
ws.Range("J1").Value = "YEARLY CHANGE"
ws.Range("K1").Value = "PERCENT CHANGE"
ws.Range("L1").Value = "TOTAL STOCK VOLUME"
ws.Range("O2").Value = "GREATEST % INCREASE"
ws.Range("O3").Value = "GREATEST % DECREASE"
ws.Range("O4").Value = "GREATEST TOTAL VOLUME"
ws.Range("O5").Value = "LOWEST TOTAL VOLUME"
ws.Range("P1").Value = ws.Range("I1").Value
ws.Range("Q1").Value = "VALUE"
ws.Range("A2:G759001").Sort Key1:=ws.Range("A2"), Order1:=xlAscending, Header:=xlNo
Dim ticker As String
Dim yearchange As Double
Dim percentchange As Double
Dim totalvolume As LongLong
totalvolume = 0
volsum = 2
Dim start As Double
Dim finish As Double
Dim lastrow As Long
start = ws.Cells(2, 3).Value
lastrow = ws.Range("A" & Rows.Count).End(xlUp).Row
'Loop through all stock data
For i = 2 To lastrow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value
totalvolume = totalvolume + ws.Cells(i, 7).Value
finish = ws.Cells(i, 6).Value
yearchange = finish - start
percentchange = (yearchange / start)
ws.Range("I" & volsum).Value = ticker
ws.Range("J" & volsum).Value = yearchange
ws.Range("K" & volsum).Value = percentchange
ws.Range("K" & volsum).NumberFormat = "0.00%"
ws.Range("L" & volsum).Value = totalvolume
'Apply conditional formatting
If ws.Range("J" & volsum).Value >= 0 Then
ws.Range("J" & volsum).Interior.ColorIndex = 4
ElseIf ws.Range("J" & volsum).Value < 0 Then
ws.Range("J" & volsum).Interior.ColorIndex = 3
ElseIf ws.Range("K" & volsum).Value >= 0 Then
ws.Range("K" & volsum).Interior.ColorIndex = 4
ElseIf ws.Range("K" & volsum).Value < 0 Then
ws.Range("K" & volsum).Interior.ColorIndex = 3
End If
'Reset for each new loop iteration, the next stock 
volsum = volsum + 1
totalvolume = 0
start = ws.Cells(i + 1, 3).Value
Else
totalvolume = totalvolume + ws.Cells(i + 1, 7).Value
End If
Next i
'Get statistics
ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
ws.Range("Q5").Value = WorksheetFunction.Min(ws.Range("L2:L" & lastrow))
'Match each stock to the statistic that appears
maxup = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & lastrow), 0)
maxdown = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & lastrow), 0)
maxvol = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & lastrow), 0)
minvol = WorksheetFunction.Match(ws.Range("Q5").Value, ws.Range("L2:L" & lastrow), 0)
ws.Range("P2").Value = ws.Cells(maxup + 1, 9).Value
ws.Range("P3").Value = ws.Cells(maxdown + 1, 9).Value
ws.Range("P4").Value = ws.Cells(maxvol + 1, 9).Value
ws.Range("P5").Value = ws.Cells(minvol + 1, 9).Value
'Make the Macro run on all worksheets
Next ws
End Sub
