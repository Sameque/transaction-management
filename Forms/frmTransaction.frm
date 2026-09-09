VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTransaction 
   Caption         =   "Transaction Registration"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   2085
      TabIndex        =   8
      Top             =   3720
      Width           =   1635
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Salvar"
      Height          =   495
      Left            =   90
      TabIndex        =   7
      Top             =   3735
      Width           =   1635
   End
   Begin VB.Frame fraDados 
      Caption         =   "Transaction Data"
      Height          =   3525
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtDescription 
         Height          =   1050
         Left            =   90
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2385
         Width           =   3405
      End
      Begin VB.TextBox txtId 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1380
         TabIndex        =   1
         Top             =   240
         Width           =   2100
      End
      Begin VB.TextBox txtAmount 
         Height          =   315
         Left            =   1380
         TabIndex        =   3
         Top             =   990
         Width           =   2100
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   2100
      End
      Begin MSMask.MaskEdBox mskDate 
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Top             =   1335
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskCard 
         Height          =   360
         Left            =   1380
         TabIndex        =   2
         Top             =   600
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   19
         Mask            =   "#### #### #### ####"
         PromptChar      =   " "
      End
      Begin VB.Label lblDescription 
         Caption         =   "Descrição:"
         Height          =   240
         Left            =   135
         TabIndex        =   14
         Top             =   2085
         Width           =   1650
      End
      Begin VB.Label lblId 
         Caption         =   "ID:"
         Height          =   210
         Left            =   135
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCard 
         Caption         =   "Nr. do Cartão:*"
         Height          =   255
         Left            =   135
         TabIndex        =   10
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblAmount 
         Caption         =   "Valor:*"
         Height          =   255
         Left            =   135
         TabIndex        =   11
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label lblDate 
         Caption         =   "Data:*"
         Height          =   255
         Left            =   135
         TabIndex        =   12
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status:*"
         Height          =   255
         Left            =   135
         TabIndex        =   13
         Top             =   1740
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_modo As String
Private m_idTransacao As Long
Private m_statusOriginal As String

Public Property Let Modo(ByVal Valor As String)
    m_modo = Valor
End Property

Public Property Let IdTransacao(ByVal Valor As Long)
    m_idTransacao = Valor
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = 0 Then
      Cancel = 1
      MsgBox "Utilize os botoes da tela para sair.", vbExclamation, "Aviso"
   End If
End Sub

Private Sub Form_Activate()
   If m_statusOriginal = "Approved" Then
     Call BloquearEdicao
     Exit Sub
   End If
   
   mskCard.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ErroHandler
   
   RemoveCloseButton Me
   LoadStatusCombo cboStatus
    
   If m_modo = "NEW" Then
       mskDate.Text = Format(Now, "DD/MM/YYYY HH:NN")
       Me.Caption = "Nova Transação"
       txtId.Text = ""
   Else
       Me.Caption = "Editar Transação"
       Call LoadData
   End If
   
   Exit Sub

ErroHandler:
   Dim strRoutine As String
   
   strRoutine = Me.Name & " Form_Load"
   
   Call DisplayFriendlyMessage(Err.Description)
   Call SaveLogErr("Form_Load", Err.Description)
End Sub

Private Sub LoadData()
On Error GoTo ErroHandler
   
   Dim rs As ADODB.Recordset
   Dim sql As String
   
   sql = "SELECT TransactionId, CardNumber, TransactionAmount, " & _
         "       TransactionDate, Description, TransactionStatus " & _
         "FROM dbo.Transactions " & _
         "WHERE TransactionId = " & m_idTransacao
   
   Set rs = QueryExec(sql)
   If rs Is Nothing Or (rs.EOF And rs.BOF) Then
      MsgBox "Transaction not found.", vbExclamation, "Error"
      Me.Hide
      Exit Sub
   End If
   
   txtId.Text = rs.Fields("TransactionId").value
   mskCard.Text = MaskCard(rs.Fields("CardNumber").value)
   txtAmount.Text = Format(rs.Fields("TransactionAmount").value, "#,##0.00")
   mskDate.Text = Format(rs.Fields("TransactionDate").value, "MM/DD/YYYY HH:NN")
   txtDescription.Text = rs.Fields("Description").value
   
   m_statusOriginal = rs.Fields("TransactionStatus").value
   
   Dim i As Long
   For i = 0 To cboStatus.ListCount - 1
      If cboStatus.List(i) = GetStatusLabel(m_statusOriginal) Then
         cboStatus.ListIndex = i
         Exit For
      End If
   Next i
   
   rs.Close
   Set rs = Nothing
   
   Exit Sub
   
ErroHandler:
   Call SaveLogErr("LoadData", "Erro ao editar Transação", "ID: " & m_idTransacao)
   Call DisplayFriendlyMessage("LoadData", Err.Description)
End Sub

Private Sub BloquearEdicao()
   mskCard.enabled = False
   txtAmount.enabled = False
   mskDate.enabled = False
   txtDescription.enabled = False
   cboStatus.enabled = False
   cmdSave.enabled = False
   
   MsgBox "Esta transação tem o status 'Aprovado' e não pode ser editada.", _
          vbExclamation + vbOKOnly, "Edição Bloqueada"
End Sub

Private Function Validations() As Boolean
On Error GoTo ErroHandler

   Dim strMsg As String
      
   strMsg = ""

   If Trim(mskCard.Text) = "" Then
       strMsg = strMsg & "- Número do Cartão, Obrigatório." & vbCrLf
   ElseIf Not IsValidCardNumber(Replace(mskCard.Text, " ", "")) Then
       strMsg = strMsg & "- Numero do Cartão Inválido." & vbCrLf
   End If
   
   If Trim(txtAmount.Text) = "" Then
       strMsg = strMsg & "- Valor É Obrigatório." & vbCrLf
   ElseIf Not IsValidAmount(txtAmount.Text) Then
       strMsg = strMsg & "- Valor Inválido." & vbCrLf
   End If
   
   If Trim(mskDate.Text) = "" Then
       strMsg = strMsg & "- Data É Obrigatória." & vbCrLf
   ElseIf Not IsDate(mskDate.Text) Then
       strMsg = strMsg & "- Data Inválida." & vbCrLf
   End If
   
   If Not IsValidStatus(cboStatus.ListIndex) Then
     strMsg = strMsg & "- Status " & cboStatus.Text & " Inválido." & vbCrLf
   End If
   
   If Not IsValidDescription(txtDescription.Text) Then
       strMsg = strMsg & "- Descrição Inválida." & vbCrLf
   End If
   
   If strMsg <> "" Then
     MsgBox "Por favor, corrija os seguintes erros:" & vbCrLf & vbCrLf & strMsg, _
        vbExclamation + vbOKOnly, "Validation"
     Validations = False
   Else
     Validations = True
   End If
   
   Exit Function
   
ErroHandler:
   Validations = False
   Call SaveLogErr("ValidarFormulario", Err.Description, Err)
   Call DisplayFriendlyMessage("cmdSave_Click", Err.Description)
End Function

Private Sub cmdSave_Click()
On Error GoTo ErroHandler
   
   Dim sql As String
   Dim card As String
   Dim status As String
   Dim amount As String
   Dim dateTransaction As Date
   Dim affectedLines As Long
   
   If Not Validations Then Exit Sub
   
   card = RemoveMaskForNumeric(mskCard.Text)
   amount = CashValueBrlToSql(txtAmount.Text)
   dateTransaction = CDate(mskDate.Text)
   status = GetStatusKey(cboStatus.Text)
    
   If m_modo = "EDIT" Then
        If MsgBox("Confirmar alterações nesta transação?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Confirm") = vbNo Then
            Exit Sub
        End If
   End If
   
   If m_modo = "NEW" Then
        sql = "INSERT INTO dbo.Transactions (CardNumber, TransactionAmount, TransactionDate, Description, TransactionStatus) " & _
               "VALUES ('" & card & "', " & amount & ", " & _
               "'" & Format(dateTransaction, "YYYY-MM-DD HH:NN:SS") & "', " & _
               "'" & Replace(txtDescription.Text, "'", "''") & "', " & _
               "'" & status & "')"
   Else
        sql = "UPDATE dbo.Transactions SET " & _
              "CardNumber = '" & card & "', " & _
              "TransactionAmount = " & amount & ", " & _
              "TransactionDate = '" & Format(dateTransaction, "YYYY-MM-DD HH:NN:SS") & "', " & _
              "Description = '" & Replace(txtDescription.Text, "'", "''") & "', " & _
              "TransactionStatus = '" & status & "' " & _
              "WHERE TransactionStatus <> 'Approved' AND TransactionId = " & m_idTransacao
   End If
   
   affectedLines = CommandExec(sql)
   
   If affectedLines > 0 Then
      MsgBox "Transação Salva!", vbInformation, "Sucesso"
      Unload Me
   Else
        Call DisplayFriendlyMessage("cmdSave_Click", "Falha ao salvar a transação no banco de dados")
   End If
   
   Exit Sub
   
ErroHandler:
   Call SaveLogErr("cmdSave_Click", Err.Description, "Mode: " & m_modo & ", ID: " & m_idTransacao)
   Call DisplayFriendlyMessage(Err.Description)
End Sub

Private Sub cmdCancel_Click()
   If m_statusOriginal = "Approved" Then
      Unload Me
      Exit Sub
   End If

   If MsgBox("Cancelar? As alterações não salvas serão perdidas.", _
             vbQuestion + vbYesNo + vbDefaultButton2, "Cancel") = vbYes Then
      
      Unload Me
   End If
End Sub

Private Sub txtAmount_Change()
   txtAmount.Text = FormatMonetary(txtAmount.Text)
   txtAmount.SelStart = Len(txtAmount.Text)
End Sub

