VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTransactions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transações"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   9075
   Begin VB.CommandButton cmdExportExcel 
      Caption         =   "Exportar Excel (Últim Mês)"
      Height          =   795
      Left            =   5385
      TabIndex        =   14
      Top             =   8625
      Width           =   1530
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Editar"
      Height          =   795
      Left            =   3555
      TabIndex        =   13
      Top             =   8595
      Width           =   1635
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Apagar"
      Height          =   795
      Left            =   1830
      TabIndex        =   12
      Top             =   8595
      Width           =   1635
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Novo"
      Height          =   795
      Left            =   120
      TabIndex        =   11
      Top             =   8595
      Width           =   1635
   End
   Begin VB.CommandButton cmdFirstPage 
      Caption         =   "|<"
      Height          =   375
      Left            =   495
      TabIndex        =   17
      Top             =   7860
      Width           =   495
   End
   Begin VB.CommandButton cmdPrevPage 
      Caption         =   "<"
      Height          =   375
      Left            =   975
      TabIndex        =   16
      Top             =   7860
      Width           =   495
   End
   Begin VB.CommandButton cmdNextPage 
      Caption         =   ">"
      Height          =   375
      Left            =   7485
      TabIndex        =   18
      Top             =   7860
      Width           =   495
   End
   Begin VB.CommandButton cmdLastPage 
      Caption         =   ">|"
      Height          =   375
      Left            =   7965
      TabIndex        =   19
      Top             =   7860
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Sair"
      Height          =   795
      Left            =   7350
      TabIndex        =   15
      Top             =   8595
      Width           =   1635
   End
   Begin MSFlexGridLib.MSFlexGrid grdTransactions 
      Height          =   5160
      Left            =   75
      TabIndex        =   10
      Top             =   2610
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   9102
      _Version        =   393216
      Cols            =   9
   End
   Begin VB.Frame fraFilters 
      Caption         =   "Filtrros"
      Height          =   1815
      Left            =   90
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      Begin VB.ComboBox cboCategory 
         Height          =   315
         Left            =   5400
         TabIndex        =   3
         Text            =   "cboCategory"
         Top             =   300
         Width           =   1875
      End
      Begin VB.Frame fmrFilterValue 
         Caption         =   "Valor"
         Height          =   795
         Left            =   3345
         TabIndex        =   25
         Top             =   825
         Width           =   4005
         Begin VB.TextBox txtFilterAmountMax 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """R$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
            Height          =   315
            Left            =   2565
            MaxLength       =   13
            TabIndex        =   7
            Top             =   300
            Width           =   1200
         End
         Begin VB.TextBox txtFilterAmountMin 
            DataMember      =   "&"
            Height          =   315
            Left            =   690
            TabIndex        =   6
            Top             =   300
            Width           =   1200
         End
         Begin VB.Label lblFilterAmountMax 
            Caption         =   "Máximo:"
            Height          =   255
            Left            =   1965
            TabIndex        =   27
            Top             =   330
            Width           =   630
         End
         Begin VB.Label lblFilterAmountMin 
            Caption         =   "Mínimo:"
            Height          =   315
            Left            =   90
            TabIndex        =   26
            Top             =   330
            Width           =   630
         End
      End
      Begin MSMask.MaskEdBox mskFilterDateFrom 
         Height          =   315
         Left            =   525
         TabIndex        =   4
         Top             =   1125
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFilterDateTo 
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Top             =   1125
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskCard 
         Height          =   360
         Left            =   690
         TabIndex        =   1
         Top             =   285
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   19
         Mask            =   "#### #### #### ####"
         PromptChar      =   " "
      End
      Begin VB.CommandButton cmdClearFilters 
         Caption         =   "&Limpar"
         Height          =   600
         Left            =   7485
         TabIndex        =   9
         Top             =   930
         Width           =   1290
      End
      Begin VB.ComboBox cboFilterStatus 
         Height          =   315
         ItemData        =   "frmTransactions.frx":0000
         Left            =   3015
         List            =   "frmTransactions.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   285
         Width           =   1545
      End
      Begin VB.Frame frmFilterDate 
         Caption         =   "Data"
         Height          =   795
         Left            =   90
         TabIndex        =   22
         Top             =   825
         Width           =   3195
         Begin VB.Label lblFilterDateFrom 
            Caption         =   "Até:"
            Height          =   315
            Left            =   1635
            TabIndex        =   24
            Top             =   330
            Width           =   390
         End
         Begin VB.Label lblFilterDateTo 
            Caption         =   "De:"
            Height          =   210
            Left            =   165
            TabIndex        =   23
            Top             =   330
            Width           =   345
         End
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Filtrar"
         Height          =   600
         Left            =   7485
         TabIndex        =   8
         Top             =   210
         Width           =   1290
      End
      Begin VB.Label lblCategory 
         Caption         =   "Categoria:"
         Height          =   270
         Left            =   4620
         TabIndex        =   28
         Top             =   315
         Width           =   885
      End
      Begin VB.Label lblFilterStatus 
         Caption         =   "Status:"
         Height          =   255
         Left            =   2475
         TabIndex        =   21
         Top             =   315
         Width           =   570
      End
      Begin VB.Label lblFilterCard 
         Caption         =   "Cartão:"
         Height          =   255
         Left            =   135
         TabIndex        =   20
         Top             =   315
         Width           =   720
      End
   End
   Begin VB.Label lblPagination 
      Alignment       =   2  'Center
      Caption         =   "Page 1 of 1"
      Height          =   375
      Left            =   1530
      TabIndex        =   30
      Top             =   7860
      Width           =   5955
   End
   Begin VB.Label lblTransactions 
      Caption         =   "Transações:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   210
      TabIndex        =   29
      Top             =   2235
      Width           =   2940
   End
End
Attribute VB_Name = "frmTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_currentPage As Long
Private m_recordsPerPage As Long
Private m_totalRecords As Long
Private m_totalPages As Long

Private Sub cmdClearFilters_Click()
   Call ClearFields
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorHandler
   
   Dim id As Long
   Dim cardDisplay As String
   
   id = GetIdSelected

   If id < 0 Then
       MsgBox "Selecione um Id", vbExclamation, "Warning"
       Exit Sub
   End If
   
    If IsTransactionApproved Then
      MsgBox "Não é possível apagar transações com status 'Aprovado'.", vbExclamation, "Remoção Bloqueada"
       Exit Sub
   End If

   cardDisplay = grdTransactions.TextMatrix(grdTransactions.row, 1)

   If MsgBox("Confirme a exclusãoda transação, ID: " & id & ", Cartão: " & cardDisplay & "?", _
             vbQuestion + vbYesNo + vbDefaultButton2 + vbExclamation, "Confirm Deletion") = vbNo Then
       Exit Sub
   End If

   If Not Delete(id) Then
      Call SaveLogErr("cmdDelete_Click", "Nenhuma linha afetada na exclusão", "ID: " & id)
      Call DisplayFriendlyMessage("Falha ao excluir a transação.", vbCrLf & " - Verifique se o Status está como Aprovado." & vbCrLf & " - Verifique se a transação ainda existe.")
      Exit Sub
   End If
   
    MsgBox "Transação excluída com sucesso!", vbInformation, "Sucesso"
    Call LoadGrid

   Exit Sub

ErrorHandler:
   Call SaveLogErr("cmdDelete_Click", Err.Description)
   Call DisplayFriendlyMessage("Falha ao excluir a transação.")
End Sub

Private Function Delete(ByVal id As Long) As Boolean

   Dim cardDisplay As String
   Dim sql As String
   Dim rowsAffected As Long

   sql = "DELETE FROM dbo.Transactions WHERE TransactionStatus <> 'Approved' AND TransactionId = " & id
   
   rowsAffected = CommandExec(sql)

   Delete = rowsAffected > 0

End Function

Private Function GetIdSelected() As Long
On Error GoTo ErrorHandler

    If grdTransactions.row < 1 Then
        GetIdSelected = -1
        Exit Function
    End If

    GetIdSelected = CLng(grdTransactions.TextMatrix(grdTransactions.row, 0))

    Exit Function

ErrorHandler:
    GetIdSelected = -1
End Function


Private Sub cmdEdit_Click()
On Error GoTo ErrorHandler
   
   Dim id As Long
      
   id = GetIdSelected
   
   If id < 0 Then
     MsgBox "Selecione um Id", vbExclamation, "Warning"
      Exit Sub
   End If
      
   If IsTransactionApproved Then
       MsgBox "Não é possível editar transações com status 'Aprovado'.", vbExclamation, "Edição Bloqueada"
       Exit Sub
   End If
   
   Set frmTransaction = New frmTransaction
   frmTransaction.Modo = "EDIT"
   frmTransaction.IdTransacao = id
   frmTransaction.Show vbModal
   
   If frmTransaction.Visible = False Then
       Call LoadGrid
   End If
   
   Exit Sub
   
ErrorHandler:
   Call SaveLogErr("cmdEdit_Click", Err.Description)
   Call DisplayFriendlyMessage("cmdEdit_Click", Err.Description)
End Sub

Private Function IsTransactionApproved() As Boolean
    Dim statusSel As String
    
    statusSel = GetStatusSelected
    
    If LCase(statusSel) <> "aprovado" Then
        IsTransactionApproved = False
        Exit Function
    End If
    
    IsTransactionApproved = True

End Function
Private Function GetStatusSelected() As String
On Error GoTo ErrorHandler

    If grdTransactions.Rows <= 1 Or grdTransactions.row < 1 Then
        GetStatusSelected = ""
        Exit Function
    End If

    GetStatusSelected = grdTransactions.TextMatrix(grdTransactions.row, 4)

    Exit Function

ErrorHandler:
    GetStatusSelected = ""
End Function

Private Sub cmdExit_Click()
   Me.Hide
End Sub

Private Sub cmdExportExcel_Click()
On Error GoTo ErrorHandler
   Dim blnExported As Boolean
   Dim strMsg As String
   Dim strReportPath As String
   
   cmdExportExcel.enabled = False
   
   blnExported = ExportLastMonthToExcel(strMsg, strReportPath)
   
   cmdExportExcel.enabled = True
   
   If Not blnExported Then
      Call DisplayFriendlyMessage(strMsg, "Verifique se o Excel está instalado.")
      Exit Sub
   End If

   MsgBox strMsg, vbInformation, "Exportação"
   Call Shell("explorer.exe " & strReportPath, vbNormalFocus)

   Exit Sub
ErrorHandler:
    cmdExportExcel.enabled = True
    Call SaveLogErr("cmdExportExcel_Click", Err.Description)
    Call DisplayFriendlyMessage("Erro ao exportar relatório")
End Sub

Private Sub cmdFilter_Click()
   m_currentPage = 1
   Call EnableButons(False)
   Call LoadGrid

End Sub
Private Sub cmdNew_Click()
On Error GoTo ErrorHandler

   frmTransaction.Modo = "NEW"
   frmTransaction.Show vbModal

   If frmTransaction.Visible = False Then
      Call LoadGrid
   End If

   Exit Sub

ErrorHandler:
    Call SaveLogErr("cmdNew_Click", Err.Description)
    Call DisplayFriendlyMessage("cmdNew_Click", Err)
End Sub

Private Sub Form_Load()
   Call ClearFields
   Call ConfigureGrid
   LoadStatusCombo cboFilterStatus
   LoadCategoryCombo cboCategory

   m_recordsPerPage = 30
   m_currentPage = 1
   Call EnableButons(True)
   Call UpdatePagination

End Sub

Private Function ClearFields()
   Dim mask As String
   
   mask = mskCard.mask
   mskCard.mask = ""
   mskCard.Text = ""
   mskCard.mask = mask
   
   mask = mskFilterDateFrom.mask
   mskFilterDateFrom.mask = ""
   mskFilterDateFrom.Text = ""
   mskFilterDateFrom.mask = mask
   
   mask = mskFilterDateTo.mask
   mskFilterDateTo.mask = ""
   mskFilterDateTo.Text = ""
   mskFilterDateTo.mask = mask
      
   txtFilterAmountMin.Text = "0"
   txtFilterAmountMax.Text = "0"
         
   If cboFilterStatus.ListIndex >= 0 Then cboFilterStatus.ListIndex = 0
   If cboCategory.ListIndex >= 0 Then cboCategory.ListIndex = 0
   
End Function

Private Sub ConfigureGrid()
On Error GoTo ErrorHandler

   Dim i As Long

   With grdTransactions
        ' Columns: ID, Card, Amount, Date, Description, Status, Category
        .Cols = 7
        .Rows = 2  ' 1 header + 1 initial row
        .FixedRows = 1
        .FixedCols = 0

        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Cartão"
        .TextMatrix(0, 2) = "Valor"
        .TextMatrix(0, 3) = "Data"
        .TextMatrix(0, 4) = "Status"
        .TextMatrix(0, 5) = "Categoria"
        .TextMatrix(0, 6) = "Description"
        
        .ColWidth(0) = 600    ' ID
        .ColWidth(1) = 1900   ' Card
        .ColWidth(2) = 1400   ' Amount
        .ColWidth(3) = 1600   ' Date
        .ColWidth(4) = 1200   ' Status
        .ColWidth(5) = 1200   ' Category
        .ColWidth(6) = 3000   ' Description

        .row = 0
        .col = 0
        For i = 0 To .Cols - 1
            .col = i
            .CellFontBold = True
            .CellBackColor = RGB(0, 51, 102)
            .CellForeColor = RGB(255, 255, 255)
            .CellAlignment = flexAlignCenterCenter
        Next i

        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .BackColor = RGB(255, 255, 255)
        .BackColorFixed = RGB(0, 51, 102)
        .ForeColorFixed = RGB(255, 255, 255)
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        .WordWrap = False
   End With

   Exit Sub

ErrorHandler:
    Call SaveLogErr("ConfigurarGrid", Err.Description)
    Call DisplayFriendlyMessage("Erro ao Criar Grid")
End Sub

Private Sub LoadGrid()
On Error GoTo ErrorHandler
    
   Dim sql As String
   Dim whereClause As String
   Dim offset As Long
   Dim rs As ADODB.Recordset
   Dim i As Long
   Dim row As Long

   whereClause = SetUpCommandWhere

   m_totalRecords = CountRecords(whereClause)
   m_totalPages = IIf(m_totalRecords = 0, 1, -Int(-m_totalRecords / m_recordsPerPage))

   If m_currentPage > m_totalPages Then m_currentPage = m_totalPages
   If m_currentPage < 1 Then m_currentPage = 1

   offset = (m_currentPage - 1) * m_recordsPerPage

    sql = "SELECT t.TransactionId, t.CardNumber, t.TransactionAmount, " & _
          "       t.TransactionDate, t.Description, t.TransactionStatus, " & _
          "       dbo.fn_GetValueCategory(t.TransactionAmount) AS Category " & _
          "FROM dbo.Transactions t"

    If whereClause <> "" Then
        sql = sql & " WHERE " & whereClause
    End If

    sql = sql & " ORDER BY t.TransactionDate DESC, t.TransactionId DESC " & _
              "OFFSET " & offset & " ROWS FETCH NEXT " & m_recordsPerPage & " ROWS ONLY"

    Set rs = QueryExec(sql)
    If rs Is Nothing Then
        Call DisplayFriendlyMessage("LoadGrid", "Falha ao consultar transações")
        Call EnableButons(True)
        Exit Sub
    End If

    grdTransactions.Redraw = False
    grdTransactions.Rows = 2

    row = 1
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            If row >= grdTransactions.Rows Then
                grdTransactions.Rows = grdTransactions.Rows + 1
            End If

            grdTransactions.TextMatrix(row, 0) = rs.Fields("TransactionId").value
            grdTransactions.TextMatrix(row, 1) = MaskCard(rs.Fields("CardNumber").value)
            grdTransactions.TextMatrix(row, 2) = Format(rs.Fields("TransactionAmount").value, "#,##0.00")
            grdTransactions.TextMatrix(row, 3) = Format(rs.Fields("TransactionDate").value, "DD/MM/YYYY HH:NN")
            grdTransactions.TextMatrix(row, 4) = GetStatusLabel(rs.Fields("TransactionStatus").value)
            grdTransactions.TextMatrix(row, 5) = GetCategoryLabel(rs.Fields("Category").value)
            grdTransactions.TextMatrix(row, 6) = rs.Fields("Description").value

            Select Case LCase(rs.Fields("TransactionStatus").value)
                Case "approved"
                    grdTransactions.row = row
                    grdTransactions.col = 4
                    grdTransactions.CellForeColor = RGB(0, 100, 0)
                Case "cancelled"
                    grdTransactions.row = row
                                        grdTransactions.col = 4
                    grdTransactions.CellForeColor = RGB(180, 0, 0)
                Case "pending"
                    grdTransactions.row = row
                                        grdTransactions.col = 4
                    grdTransactions.CellForeColor = RGB(180, 120, 0)
            End Select


            row = row + 1
            rs.MoveNext
        Loop
    End If

    If row = 1 Then
        grdTransactions.Rows = 2
        For i = 0 To grdTransactions.Cols - 1
            grdTransactions.TextMatrix(1, i) = ""
        Next i
    Else
        grdTransactions.Rows = row
    End If

    grdTransactions.Redraw = True
    rs.Close
    Set rs = Nothing

    Call UpdatePagination
    Call UpdateStatusBar
    Call EnableButons(True)

    Exit Sub

ErrorHandler:
    grdTransactions.Redraw = True
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    Call SaveLogErr("LoadGrid", Err.Description)
    Call DisplayFriendlyMessage("LoadGrid", Err.Description)
    Call EnableButons(True)
End Sub

Private Function CountRecords(Optional ByVal whereClause As String = "") As Long
On Error GoTo ErrorHandler
    
   Dim sql As String
   Dim rs As ADODB.Recordset

   sql = "SELECT COUNT(*) AS Total FROM dbo.Transactions t"
   If whereClause <> "" Then
      sql = sql & " WHERE " & whereClause
   End If

   Set rs = QueryExec(sql)
   If Not rs Is Nothing And Not (rs.EOF And rs.BOF) Then
      CountRecords = rs.Fields("Total").value
   Else
      CountRecords = 0
   End If

   If Not rs Is Nothing Then
      rs.Close
      Set rs = Nothing
   End If

   Exit Function

ErrorHandler:
   CountRecords = 0
   If Not rs Is Nothing Then
      If rs.State = adStateOpen Then rs.Close
      Set rs = Nothing
   End If
   Call SaveLogErr("ContarRegistros", Err.Description)
End Function

Private Sub UpdateStatusBar()
On Error GoTo ErrorHandler

   Dim Msg As String

   If m_totalRecords = 0 Then
      Msg = "No transactions found"
   Else
      Msg = "Total: " & Format(m_totalRecords, "#,##0") & " transações(s)"
      Msg = Msg & " | Page " & m_currentPage & " of " & m_totalPages
      Msg = Msg & " | " & m_recordsPerPage & " per page"
    End If

    Me.Caption = "Transações- " & Msg

    Exit Sub

ErrorHandler:
    Call SaveLogErr("AtualizarStatusBar", Err.Description)
End Sub

Private Function SetUpCommandWhere() As String
On Error GoTo ErrorHandler
   
   Dim card As String
   Dim amountMin As String
   Dim amountMax As String
   Dim dateFrom As String
   Dim dateTo As String
   Dim status As String
   Dim category As String
   
   Dim where As String
   
   where = ""

   card = RemoveMaskForNumeric(mskCard.Text)
   If card <> "" Then
       where = where & " AND t.CardNumber LIKE '%" & card & "%'"
   End If

   dateFrom = DateBrlToSql(mskFilterDateFrom.Text)
   If dateFrom <> "" Then
       where = where & " AND t.TransactionDate >= '" & Format(dateFrom, "YYYY-MM-DD") & " 00:00:00'"
   End If

   dateTo = DateBrlToSql(mskFilterDateTo.Text)
   If dateTo <> "" Then
       where = where & " AND t.TransactionDate <= '" & Format(dateTo, "YYYY-MM-DD") & " 23:59:59'"
   End If

   amountMin = CashValueBrlToSql(txtFilterAmountMin.Text)
   If amountMin <> "" And amountMin <> "0.00" Then
        where = where & " AND t.TransactionAmount >= " & amountMin
   End If

   amountMax = CashValueBrlToSql(txtFilterAmountMax.Text)
   If amountMax <> "" And amountMax <> "0.00" Then
       where = where & " AND t.TransactionAmount <= " & amountMax
   End If

   If cboFilterStatus.ListIndex > 0 Then
       status = GetStatusKey(cboFilterStatus.Text)
       where = where & " AND t.TransactionStatus = '" & status & "'"
   End If

   If cboCategory.ListIndex > 0 Then
      
      category = GetCategoryKey(cboCategory.Text)
        where = where & " AND dbo.fn_GetValueCategory(t.TransactionAmount) = '" & category & "'"
   End If

   If where = "" Then
        SetUpCommandWhere = ""
   Else
        SetUpCommandWhere = Mid(where, 6)
   End If

   Exit Function

ErrorHandler:
    SetUpCommandWhere = ""
    Call SaveLogErr("SetUpCommandWhere", Err.Description)
End Function

Private Sub EnableButons(ByVal enabled As Boolean)
   cmdNew.enabled = enabled
   cmdEdit.enabled = enabled And (grdTransactions.Rows > 1)
   cmdDelete.enabled = enabled And (grdTransactions.Rows > 1)
   cmdFilter.enabled = enabled
   cmdClearFilters.enabled = enabled
   cmdExportExcel.enabled = enabled
End Sub

Private Sub UpdatePagination()
On Error GoTo ErrorHandler
   
   lblPagination.Caption = "Page " & m_currentPage & " of " & m_totalPages
   
   cmdFirstPage.enabled = (m_currentPage > 1)
   cmdPrevPage.enabled = (m_currentPage > 1)
   cmdNextPage.enabled = (m_currentPage < m_totalPages)
   cmdLastPage.enabled = (m_currentPage < m_totalPages)
   
   Exit Sub

ErrorHandler:
    Call SaveLogErr("AtualizarPaginacao", Err.Description)
End Sub

Private Sub txtFilterAmountMax_Change()
   txtFilterAmountMax.Text = FormatMonetary(txtFilterAmountMax.Text)
   txtFilterAmountMax.SelStart = Len(txtFilterAmountMax.Text)
End Sub


Private Sub txtFilterAmountMax_KeyPress(KeyAscii As Integer)
   NumbersOnly_KeyPress KeyAscii
End Sub

Private Sub txtFilterAmountMin_Change()
   txtFilterAmountMin.Text = FormatMonetary(txtFilterAmountMin.Text)
   txtFilterAmountMin.SelStart = Len(txtFilterAmountMin.Text)
End Sub

Private Sub txtFilterAmountMin_KeyPress(KeyAscii As Integer)
   NumbersOnly_KeyPress KeyAscii
End Sub

Private Sub cmdFirstPage_Click()
   If m_currentPage > 1 Then
      m_currentPage = 1
      Call LoadGrid
   End If
End Sub

Private Sub cmdPrevPage_Click()
   If m_currentPage > 1 Then
      m_currentPage = m_currentPage - 1
      Call LoadGrid
   End If
End Sub
Private Sub cmdNextPage_Click()
   If m_currentPage < m_totalPages Then
      m_currentPage = m_currentPage + 1
      Call LoadGrid
   End If
End Sub

Private Sub cmdLastPage_Click()
   If m_currentPage < m_totalPages Then
      m_currentPage = m_totalPages
      Call LoadGrid
   End If
End Sub

