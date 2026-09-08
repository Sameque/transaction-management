Attribute VB_Name = "Messages"
Option Explicit

Public Function DisplayFriendlyMessage(MsgErro As String, Optional ByVal Detail As String = "")
   Dim strMessage As String

   If MsgErro <> "" Then strMessage = MsgErro & vbCrLf & vbCrLf

   strMessage = strMessage & "Não conseguimos completar sua solicitacao no momento." & vbCrLf & vbCrLf & _
                  "O que fazer:" & vbCrLf & _
                  "• Tente novamente em alguns instantes." & vbCrLf & _
                  "• Se persistir, contate o suporte."

   If Detail <> "" Then strMessage = strMessage & vbCrLf & vbCrLf & "Detalhe do erro: " & Detail

   MsgBox strMessage, vbExclamation + vbOKOnly, "Ops, algo deu errado"

End Function
