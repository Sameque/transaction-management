Attribute VB_Name = "Validations"

Public Function IsValidCardNumber(ByVal sCardNumber As String) As Boolean
    Dim sDigits As String
    Dim i As Integer
    Dim iDigit As Integer
    Dim iSum As Integer
    Dim bDouble As Boolean
    
    sDigits = ""
    For i = 1 To Len(sCardNumber)
        If IsNumeric(Mid(sCardNumber, i, 1)) Then
            sDigits = sDigits & Mid(sCardNumber, i, 1)
        End If
    Next i
    
    If Len(sDigits) < 13 Then
        IsValidCardNumber = False
        Exit Function
    End If
    
    bDouble = False
    iSum = 0
    
    For i = Len(sDigits) To 1 Step -1
        iDigit = CInt(Mid(sDigits, i, 1))
        
        If bDouble Then
            iDigit = iDigit * 2
            If iDigit > 9 Then iDigit = iDigit - 9
        End If
        
        iSum = iSum + iDigit
        bDouble = Not bDouble
    Next i
    
    IsValidCardNumber = (iSum Mod 10 = 0)
End Function

Public Function IsValidDescription(ByVal Descricao As String) As Boolean
On Error GoTo ErroHandler

   If Trim(Descricao) = "" Then
      IsValidDescription = False
      Exit Function
   End If
   
   If Len(Trim(Descricao)) > 255 Then
      IsValidDescription = False
      Exit Function
   End If

   IsValidDescription = True
   Exit Function

ErroHandler:
    IsValidDescription = False
End Function

Public Function IsValidAmount(ByVal strAmount As String) As Boolean
On Error GoTo ErrorHandler

   Dim dbAmaunt As Double
   
   If IsNull(strAmount) Or Trim(strAmount) = "" Then
      IsValidAmount = False
      Exit Function
   End If
      
   strAmount = Replace(Replace(strAmount, ".", ""), ",", ".")
   dbAmaunt = CDec(strAmount)
   
   If dbAmaunt <= 0 Then
      IsValidAmount = False
      Exit Function
   End If
      
   IsValidAmount = True
   
   Exit Function
ErrorHandler:
   IsValidAmount = False
End Function

Public Function IsValidStatus(ByVal intStatus As Integer) As Boolean
On Error GoTo ErrorHandler
   
   If intStatus = -1 Or intStatus = 0 Then
      IsValidStatus = False
      Exit Function
   End If
            
   IsValidStatus = True
   
   Exit Function
ErrorHandler:
   IsValidStatus = False
End Function

