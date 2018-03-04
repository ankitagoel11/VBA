Sub tickerdisplay():

'Finds the last non-blank cell in a single row or column

Dim j, i As Long

Dim sum, k, l As Double

Dim lRow As Long

j = 1

sum = 0

k = 1

    
  '  Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row

    
   ' MsgBox "Last Row: " & lRow
       
    For i = 2 To lRow
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
        Cells(j, 13).Value = Cells(i, 1).Value
    
        j = j + 1
     
    
        End If
  
    Next i


For l = 2 To lrow

    If Cells(l, 1).Value = Cells(l + 1, 1).Value Then

    sum = sum + Cells(l, 7)

    Else: sum = sum + Cells(l, 7)

    Cells(k, 14).Value = sum

    sum = 0

    k = k + 1


    End If

Next l

End Sub