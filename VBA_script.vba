Sub getData()

For k = 1 To 3

Dim Sheetname As String
Dim i As Double
Dim z As Integer
Dim percChange As Double
Dim length As Double
Worksheets(k).Activate

Dim numrows As Double
numrows = (Range("A1").End(xlDown).Row)
MsgBox (numrows)

Dim AllRows() As String

'Worksheets(1).Activate
Dim tickersym As String


Dim Result() As String
tickersym = ""


Dim resultrow
resultrow = 1



Cells(1, 9).Value = "Ticker Symbol"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
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


Next i

Columns("L").ColumnWidth = 25
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
