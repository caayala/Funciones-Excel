Function luhnSum(InVal As String) As Integer
    Dim evenSum As Integer
    Dim oddSum As Integer
         
    evenSum = 0
    oddSum = 0
     
    Dim strLen As Integer
    strLen = Len(InVal)
     
    Dim i As Integer
    For i = 1 To strLen Step 1
        Dim digit As Integer
        digit = CInt(Mid(InVal, strLen - i + 1, 1))
        
        ' Debug.Print "Valor " & digit & " de posicion " & i
        
        If ((i Mod 2) = 0) Then
                oddSum = oddSum + digit
                ' Debug.Print "Suma odd de " & digit & " de posicion " & i
            
            Else
                digit = digit * 2
                ' Debug.Print "Suma even de " & digit & " de posicion " & i
             
                If (digit > 9) Then
                    digit = digit - 9
                End If
                
                evenSum = evenSum + digit
        End If
    Next i
     
    luhnSum = (oddSum + evenSum)
End Function
 
' for the curious
Function luhnCheckSum(InVal As String)
    Dim vluhnSum As Integer
    vluhnSum = luhnSum(Left(InVal, Len(InVal) - 1))
    luhnCheckSum = (vluhnSum + CInt(Right(InVal, 1)))
End Function
 
' true/false check
Function luhnCheck(InVal As String)
    luhnCheck = (luhnCheckSum(InVal) Mod 10) = 0
End Function
 
' returns a number which, appended to the InVal, turns the composed number into a valid luhn number
Function luhnNext(InVal As String)
    Dim luhnCheckSumRes
    
    luhnCheckSumRes = luhnSum(InVal) Mod 10
     
    If (luhnCheckSumRes = 0) Then
        luhnNext = 0
    Else
        luhnNext = ((10 - luhnCheckSumRes))
    End If
End Function
