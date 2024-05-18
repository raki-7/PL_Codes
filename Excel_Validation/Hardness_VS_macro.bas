Private Sub Worksheet_Change(ByVal Target As Range)
    
'defining intergers, variables etc.
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim d As Integer
    Dim Answer As VbMsgBoxResult
    
'finding the column and row number in which the value is added.
    a = Target.Column
    b = Target.Row
    c = b - 1
    
'checking if there is any empty cells in the previous row.
        For d = 1 To 9
        If IsEmpty(Cells(c, d)) = True And Cells(c, d).Locked = False And IsEmpty(Cells(b, a)) = False Then
            Answer = MsgBox("Previsous cell(" & c & "," & d & ") is empty. If you proceed further this cell will be locked. Do you wish to proceed", vbYesNo + vbDefaultButton2 + vbCritical, "Error")
                If Answer = vbYes Then
                    ActiveSheet.Unprotect Password:="123"
                    Cells(c, d).Locked = True
                    ActiveSheet.Protect Password:="123"
                Else
                    Cells(b, a) = Cells(18, 1)
                    Exit Sub
                End If
        End If
        Next d
        

'locking the cell in which the data are being entered.
        If IsEmpty(Cells(b, a)) = True Then
            ActiveSheet.Unprotect Password:="123"
            Target.Locked = False
            ActiveSheet.Protect Password:="123"
        Else
            ActiveSheet.Unprotect Password:="123"
            Target.Locked = True
            ActiveSheet.Protect Password:="123"
        End If

End Sub

