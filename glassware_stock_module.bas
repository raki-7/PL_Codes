Sub Glassware_stock()
' creating integer i for number of rows, a for number of iteration,
' b for total inward, c for total outward, d for difference b/w c & b
Dim i As Integer
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer

    ActiveCell.Value = "=row(rc)"   ' to kown the total row count to run the iteration
    i = ActiveCell.Value
    
    b = 0  'initializing b and c to zero
    c = 0
    For a = 5 To i  'since the values of entry in the excel starts from 5th row
' checking if the glassware name and capacity are macthing or not
        If Cells(a, 3) = Cells(i, 3) And Cells(a, 4) = Cells(i, 4) Then
            If IsNumeric(Cells(a, 5)) = True Then  ' checking the glasware inward quantity entered is a numeric or not
            b = b + Cells(a, 5).Value ' if found to be numeric the value is added to "b" to keep track of all the inward
            End If
        End If
' checking if the glassware name and capacity are macthing or not
        If Cells(a, 3) = Cells(i, 3) And Cells(a, 4) = Cells(i, 4) Then
            If IsNumeric(Cells(a, 7)) = True Then  ' checking the glasware outward quantity entered is a numeric or not
            c = c + Cells(a, 7).Value  ' if found to be numeric the value is added to "c" to keep track of all the outward
            End If
        End If
    Next a ' iteration incresed by number 1
    
    d = b - c ' difference b/w inward and outward
    Cells(i, 9).Value = d
    
End Sub
