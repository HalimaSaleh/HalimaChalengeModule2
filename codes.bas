Attribute VB_Name = "Module2"
'Calling All Sheets
Sub CallingAllSheets()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

If IsNumeric(ws.Name) Then
Great ws
End If
Next ws
End Sub
'Header
Sub AddHeaders()
With ThisWorkbook.Sheets("2020") 'hange "Sheet1" to your sheet name
.Range("I1").Value = "Ticker"
.Range("J1").Value = "YearlyChange"
.Range("K1").Value = "Percent Change"
.Range("L1").Value = "TotalStockVolume"
.Range("O1").Value = "Ticker"
.Range("P1").Value = "Value"
End With
End Sub

Sub Great(ws As Worksheet)
Dim Ticker As String
Dim TotalStockVolume As Double
Dim openyr As Double
Dim closeyr As Double
Dim yearlyChange As Double
Dim PercentChange As Double
Dim lastRow As Long
Dim Summary_Table_Row As Long
'creating for the Greatest stuffs
Dim MPI As Double
Dim MPD As Double
Dim MTV As Double
Dim MPIT As String
Dim MPDT As String
Dim MTVT As String

lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
Summary_Table_Row = 2
TotalStockVolume = 0
MPI = 0
MPD = 0
MTV = 0

Dim i As Long
'Format For the colors
ws.Range("J2:J" & lastRow).FormatConditions.Delete

With ws.Range("J2:J" & lastRow).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
.Interior.Color = RGB(0, 255, 0)
End With
With ws.Range("J2:J" & lastRow).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
.Interior.Color = RGB(255, 0, 0)
End With

For i = 2 To lastRow
Ticker = ws.Cells(i, 1).Value
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

If ws.Cells(i, 2).Value = "20180102" Then
openyr = ws.Cells(i, 3).Value
ElseIf ws.Cells(i, 2).Value = "20181231" Then
closeyr = ws.Cells(i, 6).Value
End If

If closeyr <> 0 Then
yearlyChange = closeyr - openyr
PercentChange = (yearlyChange / openyr) * 100
Else
PercentChange = 0
End If

'for the greatest taincrease and decrease table
If PercentChange > MPI Then
MPI = PercentChange
MPIT = Ticker
End If

If PercentChange < MPD Then
MPD = PercentChange
MPDT = Ticker
End If

If TotalStockVolume > MTV Then
MTV = TotalStockVolume
MTVT = Ticker
End If
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ws.Cells(Summary_Table_Row, "I").Value = Ticker
ws.Cells(Summary_Table_Row, "J").Value = yearlyChange
ws.Cells(Summary_Table_Row, "K").NumberFormat = "0.0%"
ws.Cells(Summary_Table_Row, "K").Value = PercentChange
ws.Cells(Summary_Table_Row, "L").Value = TotalStockVolume

Summary_Table_Row = Summary_Table_Row + 1
TotalStockVolume = 0
End If

Next i

With ws
.Range("N2:N4").Value = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
.Range("O2").Value = MPIT
.Range("O3").Value = MPDT
.Range("O4").Value = MTVT
.Range("P2").Value = MPI
.Range("P3").Value = MPD
.Range("P4").Value = MTV
End With
End Sub

