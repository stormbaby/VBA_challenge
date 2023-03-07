Sub getData()

For k = 1 To 3

Dim Sheetname As String
Dim i As Double
Dim z As Integer
Dim percChange As Double
Dim length As Double
Dim highChange As Double
Dim highChangeticker As String
Dim lowchangeticker As String
Dim highvolticker As String
Dim lowChange As Double
Dim highVol As Double
Worksheets(k).Activate

Dim numrows As Double
numrows = (Range("A1").End(xlDown).Row)
'MsgBox (numrows)

Dim AllRows() As String

'Worksheets(1).Activate
Dim tickersym As String


Dim Result() As String
tickersym = ""


Dim resultrow
resultrow = 1

length = 0
highChange = 0
highChangeticker = ""
lowchangeticker = ""
highvolticker = ""
lowChange = 0
highVol = 0

Cells(1, 9).Value = "Ticker Symbol"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Dim begprice As Double
begprice = 0
Dim endprice As Double
endprice = 0
Dim stockvolume As Double
stockvolume = 0
For i = 2 To numrows

If tickersym <> Cells(i, 1).Value Then
            If begprice <> 0 Then
                Cells(resultrow, 10).Value = endprice - begprice
                percChange = (endprice - begprice) / begprice
                Cells(resultrow, 11).Value = Round(Cells(resultrow, 11).Value, 4)
                Cells(resultrow, 11).Value = FormatPercent(percChange)
  
  '              If Cells(resultrow, 11).Value < 0 Then Cells(i, j).Interior.Color = vbRed
   '                 Else: If Cells(resultrow, 11).Value > 0 Then Cells(i, j).Interior.Color = vbGreen
    '            End If
                Cells(resultrow, 12).Value = stockvolume
                
            End If
            begprice = Cells(i, 3).Value
            tickersym = Cells(i, 1).Value
            resultrow = resultrow + 1
            Cells(resultrow, 9) = tickersym
            stockvolume = Cells(i, 7).Value

Else
        If i = numrows Then
    Cells(resultrow, 10).Value = endprice - begprice
                percChange = (endprice - begprice) / begprice
                Cells(resultrow, 11).Value = Round(Cells(resultrow, 11).Value, 4)
                Cells(resultrow, 11).Value = FormatPercent(percChange)

                Cells(resultrow, 12).Value = stockvolume
  
        End If
            endprice = Cells(i, 6).Value
            stockvolume = stockvolume + Cells(i, 7).Value
        
End If

If percChange > highChange Then
      highChangeticker = Cells(resultrow - 1, 9).Value
        highChange = percChange
ElseIf percChange < lowChange Then
        lowChange = percChange
             lowchangeticker = Cells(resultrow - 1, 9).Value
End If
    If percChange < lowChange Then
        lowChange = percChange
    End If
    If stockvolume > highVol Then
        highVol = stockvolume
        highvolticker = Cells(resultrow, 9).Value
    End If
                

Next i

'Cells(2, 16).Value = Cells(highChangerow, 1)
                Cells(2, 17).Value = FormatPercent(Round(highChange, 4))
                Cells(3, 17).Value = FormatPercent(Round(lowChange, 4))
                Cells(4, 17).Value = highVol
                Cells(2, 16).Value = highChangeticker
                Cells(3, 16).Value = lowchangeticker
                Cells(4, 16).Value = highvolticker
 '               Cells(2, 17).Value = FormatPercent(percChange)
'        MsgBox (highChangerow)
Columns("L").ColumnWidth = 25
Columns("O").ColumnWidth = 25
Columns("Q").ColumnWidth = 25
Dim numrowsx As Integer
numrowsx = (Range("k1").End(xlDown).Row)
For z = 2 To numrowsx
    If Cells(z, 11).Value < 0 Then
        Cells(z, 11).Interior.Color = vbRed
    Else: Cells(z, 11).Interior.Color = vbGreen
    End If
Next z
    



Next k

End Sub
