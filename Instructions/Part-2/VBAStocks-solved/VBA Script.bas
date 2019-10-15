Attribute VB_Name = "Module1"
Sub analyze()
    
Dim lastrow As Long
Dim nexttickerslot As Integer
Dim nextyearslot As Integer
Dim nextp100slot As Integer
Dim nextvolslot As Integer
Dim tempmin As Long
Dim tempmax As Long
Dim mindate As Long
Dim openvalue As Double
Dim closevalue As Double
Dim yearlychange As Double
Dim percentchange As Double

 
lastrow = Cells(Rows.Count, 2).End(xlUp).Row
nexttickerslot = Cells(Rows.Count, 9).End(xlUp).Row + 1
nextyearslot = Cells(Rows.Count, 10).End(xlUp).Row
nextp100slot = Cells(Rows.Count, 11).End(xlUp).Row
nextvolslot = Cells(Rows.Count, 12).End(xlUp).Row

'initial, temporary values for comparison with each date of one ticker
tempmin = 99999999
tempmax = 0

For i = 2 To lastrow

'if the ticker cell value changes, reset temp values for comparison with next ticker, increase index for yearly change slot cell
    If Cells(i, 1).Value <> Cells(i - 1, 1) Then
    
    Cells(nexttickerslot, 9).Value = Cells(i, 1)
        
    nexttickerslot = nexttickerslot + 1
    nextyearslot = nextyearslot + 1
    nextp100slot = nextp100slot + 1
    nextvolslot = nextvolslot + 1
    tempmin = 99999999
    tempmax = 0
    voltop = i
    
    End If
    
    If Cells(i, 2).Value < tempmin Then

        tempmin = Cells(i, 2)
        
        openvalue = (Cells(i, 3))

    End If

    If Cells(i, 2).Value > tempmax Then
    
        tempmax = Cells(i, 2)
        
        closevalue = (Cells(i, 6))
        
    End If
    
yearlychange = (closevalue - openvalue)

    If openvalue = 0 Then
    
    Cells(nextp100slot, 11).Value = "#DIV0"
    
    Else

    percentchange = (closevalue - openvalue) / (openvalue)

    End If
    
    If Cells(i, 1).Value <> Cells(i + 1, 1) Then
    
        volbottom = i
        
        volrange = Range((Cells(voltop, 7)), ((Cells(volbottom, 7))))
 
        Cells(nextyearslot, 10).Value = yearlychange
        
        Cells(nextp100slot, 11) = Format(percentchange, "Percent")
        
        Cells(nextvolslot, 12) = Application.Sum(volrange)
        
                    
        If Cells(nextyearslot, 10) >= 0 Then
       
            Cells(nextyearslot, 10).Interior.ColorIndex = 4
        
        Else
        
            Cells(nextyearslot, 10).Interior.ColorIndex = 3
        
        End If
        
    End If
    
Next i

End Sub

Sub reset()

Range("I2", "L500").Clear

End Sub

Sub bonus1()

Dim lastrow As Long
lastrow = Cells(Rows.Count, 11).End(xlUp).Row

Dim i As Integer

Range("o2").Value =
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'***************NOTES/DRAFTING********************
'For i = 2 To lastrow
'
'    If Cells(i, 11) >= temp100in Then
'
'        temp100in = Cells(i, 11)
'        tempintick = Cells(i, 9)
'
'    End If
'
'    If Cells(i, 11) <= temp100de Then
'
'        temp100de = Cells(i, 11)
'        tempdetick = Cells(i, 9)
'
'    End If
'
'Next i
'
'End
'
'Cells(2, 16) = temp100in
'Cells(3, 16) = temp100de


'while cells in ticker column are the same
'find max date close
'find min date open
'year change = close - open

'Dim lastrow As Long
'lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'make arrays for each ticker
'arraystartcell =  cell(i, 1) = arraystartcell
'arraystopcell = if cell(i, 1).Value <> cell(i + 1, 1).value then cell(i, 1) = arraystopcell
'if arraystart > 0 and arraystop > 0 then
'dim ranget as range
'ranget().address =


'open date = if ticker id matches that in the display section, if

'For i = 2 To lastrow
'
'    If rangestart > 0 And rangestop > 0 Then
'
'        Dim ranget As Range
'
'        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
'
'        rangestart = Cells(i, 1).Address
'
'        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
'
'        rangestop = Cells(i, 1).Address
'
'        Else
'
'        End If
'
'Next i
'
'for i = 2 to lastrow (in ticker set)
'
'mindate =
'
'for i = tempmindate to
'
'tempmindate: if (i, 2).value < (i + 1, 2).value
'
'             tempmindate.value = (i, 2).value

'Dim lastrow As Long
'Dim nexttickerslot As Integer
'dim rangestart as variant
'
'lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'nexttickerslot = Cells(Rows.Count, 9).End(xlUp).Row + 1
'
'
'For i = 2 To lastrow
'
'If Cells(i, 1).Value <> Cells(i - 1, 1) Then:
'
'    rangestart.address = cells(i, 1).address
    
'   Else
    
'   End If
    
'if cells(i, 1).value <> (i + 1, 1) then: '
'
'   rangestop.address
'    nexttickerslot = nexttickerslot + 1
'
'    Else
'
'    End If
'
''Next i
'
'Dim lastrow As Long
'Dim nexttickerslot As Integer
'Dim rangestart As Variant
'Dim workingrange As Range
'
'lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'lastrowt = Cells(Rows.Count, 9).End(xlUp).Row
'nexttickerslot = Cells(Rows.Count, 9).End(xlUp).Row + 1
'
'        For i = 2 To lastrow
'
'            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
'
'                rangestart = i
'
'            Else
'
'            End If
'
'        Next i
'
'        End
'
'        For k = 2 To lastrow
'
'            If Cells(k, 1).Value <> Cells(k + 1, 1).Value Then
'
'                rangestop = k
'
'            Else
'
'            End If
'
'        Next k
'
'        End
'
'End Sub







