Attribute VB_Name = "Helpers"
Option Explicit

Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Private Const MF_BYCOMMAND As Long = &H0&
Private Const SC_CLOSE     As Long = &HF060&

Public Function MaskCard(ByVal NumberCard As String) As String
On Error GoTo ErroHandler
    Dim digitos As String

    If Len(NumberCard) = 16 Then
        MaskCard = Left(NumberCard, 4) & " " & _
                        Mid(NumberCard, 5, 4) & " " & _
                        Mid(NumberCard, 9, 4) & " " & _
                        Right(NumberCard, 4)
    Else
        MaskCard = NumberCard
    End If

    Exit Function

ErroHandler:
    MaskCard = NumberCard
End Function

Public Function NumbersOnly_KeyPress(ByRef KeyAscii As Integer)
   If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Function

Public Function RemoveMaskForNumeric(ByVal value As String) As String
On Error GoTo ErroHandler
   Dim numbers(0 To 9) As String
   
   Dim varNumbers As Variant
   Dim newValue As String
   Dim character As String
   Dim strNumber As String
   
   Dim indexNumbers As Integer
   Dim indexValue As Integer
   
   If Trim(value) = "" Then
      RemoveMaskForNumeric = ""
      Exit Function
   End If
   
   varNumbers = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
   newValue = ""
   
   For indexValue = 1 To Len(value)
   
      character = Mid(value, indexValue, 1)

      For indexNumbers = LBound(varNumbers) To UBound(varNumbers)
         
         strNumber = varNumbers(indexNumbers)
         
         If character = strNumber Then
            newValue = newValue & character
            Exit For
         End If
      
      Next indexNumbers
   
   Next indexValue

   RemoveMaskForNumeric = newValue
   
   Exit Function
ErroHandler:
   RemoveMaskForNumeric = ""
End Function

Public Function DateBrlToSql(ByVal InputStr As String) As String
On Error GoTo ErrorHandler

   Dim dt As Date
   Dim result As String

   If IsNull(InputStr) Or Trim(InputStr) = "" Then
      DateBrlToSql = ""
      Exit Function
   End If
    
   dt = CDate(InputStr)
   DateBrlToSql = Format(dt, "yyyy-mm-dd")
   
   Exit Function
ErrorHandler:
    DateBrlToSql = ""
End Function

Public Function CashValueBrlToSql(ByVal InputStr As String) As String
   Dim newValue As String

   If IsNull(InputStr) Or Trim(InputStr) = "" Then
      CashValueBrlToSql = ""
      Exit Function
   End If
   
   newValue = Replace(Replace(InputStr, ".", ""), ",", ".")

   CashValueBrlToSql = newValue
   
End Function

Public Function FormatMonetary(ByVal Text) As String
   Dim strValue As String
   Dim lngValue As Long
   Dim lngPosition As Long
   Dim lngLen As Long

   strValue = Replace(Replace(Replace(Replace(Text, " ", ""), ",", ""), ".", ""), "-", "")
   strValue = IIf(strValue = "", "0", strValue)
   lngValue = CLng(strValue)
   strValue = CStr(lngValue)
   lngLen = Len(strValue)

   If lngValue <= 9 Then
      strValue = "0,0" & strValue
   ElseIf lngValue <= 99 Then
      strValue = "0," & Right(strValue, 2)
   ElseIf lngValue <= 99999 Then
      strValue = Mid(strValue, 1, lngLen - 2) & "," & Right(strValue, 2)
   ElseIf lngValue <= 99999999 Then
      strValue = Mid(strValue, 1, lngLen - 5) & "." & Mid(strValue, lngLen - 4, 3) & "," & Right(strValue, 2)
   ElseIf lngValue > 99999999 Then
      strValue = Mid(strValue, 1, lngLen - 8) & "." & Mid(strValue, lngLen - 7, 3) & "." & Mid(strValue, lngLen - 4, 3) & "," & Right(strValue, 2)
   End If

   FormatMonetary = strValue

End Function

Public Sub RemoveCloseButton(ByVal pForm As Form)
    
    Dim hMenu As Long
    
    hMenu = GetSystemMenu(pForm.hwnd, 0)
    
    If hMenu <> 0 Then
        RemoveMenu hMenu, SC_CLOSE, MF_BYCOMMAND
        DrawMenuBar pForm.hwnd
    End If
    
End Sub
