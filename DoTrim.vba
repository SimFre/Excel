' Trim all selected cells from bad data.
Sub DoTrim()
    Dim cell As Range
    Dim str As String
    Dim nAscii As Integer
    For Each cell In Selection.Cells
        If cell.HasFormula = False Then
            str = CStr(cell)
            If Len(str) > 0 Then
                str = Trim(str)
                If Len(str) > 0 Then
                    nAscii = Asc(Left(str, 1))
                    If nAscii < 33 Or nAscii = 160 Then
                        If Len(str) > 1 Then
                            str = Right(str, Len(str) - 1)
                        Else
                            strl = ""
                        End If
                    End If
                End If
                cell = str
            End If
        End If
    Next
End Sub
