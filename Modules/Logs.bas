Attribute VB_Name = "Logs"
Option Explicit

Private m_strLogPath As String
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function NomeEstacao() As String
    Dim nome As String * 255
    Dim tamanho As Long
    tamanho = 255
    GetComputerName nome, tamanho
    NomeEstacao = Left(nome, tamanho)
End Function

Public Sub InicializarLog(Optional ByVal PastaLogs As String = "logs")
On Error GoTo ErroHandler
   
   Dim fso As Object
   Dim pastaCompleta As String
   Dim nomeArquivo As String
      
   If InStr(PastaLogs, ":") = 0 And InStr(PastaLogs, "\") <> 1 Then
      pastaCompleta = App.Path & "\" & PastaLogs
   Else
      pastaCompleta = PastaLogs
   End If
   
   Set fso = CreateObject("Scripting.FileSystemObject")
   If Not fso.FolderExists(pastaCompleta) Then
      fso.CreateFolder pastaCompleta
   End If
   
   ' Nome do arquivo: erros_YYYYMMDD_HHMM_NOME-MAQUINA.log
   nomeArquivo = "erros_" & Format(Date, "YYYYMMDD") & "_" & _
                 Format(Time, "HHNN") & "_" & NomeEstacao() & ".log"
   
   m_strLogPath = pastaCompleta & "\" & nomeArquivo
   
   If Not fso.FileExists(m_strLogPath) Then
      Dim f As Long
      f = FreeFile
      Open m_strLogPath For Append As #f
      Print #f, "================================================================"
      Print #f, "LOG - Gerenciamento de Transacoes"
      Print #f, "Estacao: " & NomeEstacao()
      Print #f, "Usuario: " & Environ$("USERNAME")
      Print #f, "Iniciado em: " & Format(Now, "DD/MM/YYYY HH:NN:SS")
      Print #f, "================================================================"
      Print #f, ""
      Close #f
   End If
   
   Exit Sub
   
ErroHandler:
   m_strLogPath = ""
End Sub

Public Sub SaveLogErr(ByVal Routine As String, Optional ByVal MsgErro As String = "", Optional ByVal Detail As String = "")
On Error Resume Next
   
   Dim free As Long
   Dim line As String

   If m_strLogPath = "" Then Exit Sub
   
   free = FreeFile
   Open m_strLogPath For Append As #free
      
   Err.Clear
          
   If Err.Number <> 0 Then
       On Error GoTo 0
       Exit Sub
   End If

   line = "[" & Format(Now, "DD/MM/YYYY HH:NN:SS") & "] " & Routine
   
   If MsgErro <> "" Then line = line & " | MsgErro:" & MsgErro
   If Detail <> "" Then line = line & " | Detail: " & Detail

   Print #free, line
   Close #free

   On Error GoTo 0
End Sub

