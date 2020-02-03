' Use this by adding function:
' =PERSONAL.XLSB!Luhn(<yournumber>)
Function Luhn(InVal As String) As Integer
    Dim factor As Integer: factor = 2
    Dim strLen As Integer: strLen = Len(InVal)
    Dim sum As Integer
    Dim digit As Integer
    
    Dim i As Integer
    For i = strLen To 1 Step -1
        digit = CInt(Mid(InVal, i, 1))
        digit = digit * factor
        If digit >= 10 Then
            digit = digit - 9
        End If
        sum = sum + digit
        If factor = 1 Then
            factor = 2
        Else
            factor = 1
        End If
    Next i

    sum = 10 - (sum Mod 10)
    If sum = 10 Then
        sum = 0
    End If
    Luhn = sum
End Function
   
