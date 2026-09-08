Attribute VB_Name = "Translations"
Option Explicit

Private m_colStatusLabels As Collection
Private m_coStatuslKeys   As Collection

Private m_colCategoryLabels As Collection
Private m_colCategoryKeys   As Collection

Private m_bInitialized    As Boolean

Private Sub Initialize()
   If m_bInitialized Then Exit Sub
   
   Set m_colStatusLabels = New Collection
   Set m_coStatuslKeys = New Collection
   
   Set m_colCategoryLabels = New Collection
   Set m_colCategoryKeys = New Collection
   
   AddStatusPair "approved", "Aprovado"
   AddStatusPair "pending", "Pendente"
   AddStatusPair "canceled", "Cancelado"
   
   AddCategory "premium", "Premium"
   AddCategory "high", "Alta"
   AddCategory "medium", "Média"
   AddCategory "low", "Baixa"
   
   m_bInitialized = True
End Sub

Private Sub AddStatusPair(ByVal sKey As String, ByVal sLabel As String)
   m_colStatusLabels.Add sLabel, sKey
   m_coStatuslKeys.Add sKey, sLabel
End Sub

Private Sub AddCategory(ByVal sKey As String, ByVal sLabel As String)
   m_colCategoryLabels.Add sLabel, sKey
   m_colCategoryKeys.Add sKey, sLabel
End Sub

Public Function GetStatusLabel(ByVal sStatus As String) As String
   Initialize
    
   On Error Resume Next
   GetStatusLabel = m_colStatusLabels(LCase(Trim(sStatus)))
   If Err.Number <> 0 Then
      Err.Clear
      GetStatusLabel = Capitalize(sStatus)
   End If
   On Error GoTo 0
End Function

Public Function GetStatusKey(ByVal sLabel As String) As String
   Initialize
    
   On Error Resume Next
   GetStatusKey = m_coStatuslKeys(sLabel)
   If Err.Number <> 0 Then
       Err.Clear
       GetStatusKey = LCase(Trim(sLabel))
   End If
   On Error GoTo 0
    
End Function

Public Sub LoadStatusCombo(ByRef oCombo As ComboBox)
   Initialize
   oCombo.Clear
   oCombo.AddItem ""
   
   Dim vLabel As Variant
   For Each vLabel In m_colStatusLabels
       oCombo.AddItem vLabel
   Next vLabel
End Sub


Public Function GetCategoryLabel(ByVal sCategory As String) As String
   On Error Resume Next
    
   Initialize
   
   GetCategoryLabel = m_colCategoryLabels(LCase(Trim(sCategory)))
   
   If Err.Number <> 0 Then Err.Clear: GetCategoryLabel = Capitalize(sCategory)
   
   On Error GoTo 0
End Function

Public Function GetCategoryKey(ByVal sLabel As String) As String
    Initialize
    On Error Resume Next
    
    GetCategoryKey = m_colCategoryKeys(sLabel)
    If Err.Number <> 0 Then Err.Clear: GetCategoryKey = LCase(Trim(sLabel))
    
    On Error GoTo 0
End Function

Public Sub LoadCategoryCombo(ByRef oCombo As ComboBox)
    Initialize
    oCombo.Clear
    oCombo.AddItem ""
    
    Dim vItem As Variant
    For Each vItem In m_colCategoryLabels
        oCombo.AddItem vItem
    Next vItem
End Sub


Private Function Capitalize(ByVal sText As String) As String
   If Len(sText) = 0 Then
      Capitalize = ""
   Else
      Capitalize = UCase(Left(sText, 1)) & LCase(Mid(sText, 2))
   End If
End Function
