VERSION 5.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Gerenciamento de Transações"
   ClientHeight    =   10755
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   20370
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuTransactions 
      Caption         =   "&Transações"
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Relatórios"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

   Call InitializeConnection
   Call InicializarLog
   
   If Not TestConnection Then
      Call DisplayFriendlyMessage("Form_Load", "Could not connect to database. Check connection string.")
      End
   End If
End Sub

Private Sub mnuExit_Click()
   End
End Sub

Private Sub mnuTransactions_Click()
   Dim frm As New frmTransactions
   frm.Show
End Sub
