Attribute VB_Name = "ExportacaoExcel"
Option Explicit
Private Const xlWorkbookDefault = 51
Private Const xlContinuous = 1
Private Const xlThin = 2
Private Const xlAutomatic = -4105
Private Const xlCenter = -4108
Private Const xlEdgeBottom = 9
Private Const xlEdgeLeft = 7
Private Const xlEdgeRight = 10
Private Const xlEdgeTop = 8
Private Const xlInsideHorizontal = 12
Private Const xlInsideVertical = 11

Public Function ExportLastMonthToExcel(ByRef Msg As String, ByRef ReportPath As String) As Boolean
On Error GoTo ErrorHandler

   Dim xlApp As Object
   Dim xlWb As Object
   Dim xlWs As Object
   Dim rs As ADODB.Recordset
   Dim sql As String
   Dim fileName As String
   Dim reportsFolder As String
   Dim fso As Object
   Dim col As Long
   Dim row As Long
   Dim campo As ADODB.Field

   sql = "SELECT " & _
         "    TransactionId, " & _
         "    CardNumber, " & _
         "    TransactionAmount, " & _
         "    TransactionDate, " & _
         "    Description, " & _
         "    TransactionStatus, " & _
         "    dbo.fn_GetValueCategory(TransactionAmount) AS Category " & _
         "FROM dbo.Transactions " & _
         "WHERE TransactionDate >= DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 1, 0) " & _
         "  AND TransactionDate < DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0) " & _
         "ORDER BY TransactionDate DESC, TransactionId DESC"
   
   Set rs = QueryExec(sql)
   If rs Is Nothing Then
       Msg = "Falha ao recuperar dados para exportação"
       ExportLastMonthToExcel = False
       Exit Function
   End If
   
   If rs.EOF And rs.BOF Then
      ExportLastMonthToExcel = False
       Msg = "Nenhuma transação de exportação encontrada no último mês."
       rs.Close
       Set rs = Nothing
       ExportLastMonthToExcel = False
       Exit Function
   End If
   
   reportsFolder = App.Path & "\Reports"
   Set fso = CreateObject("Scripting.FileSystemObject")
   If Not fso.FolderExists(reportsFolder) Then
       fso.CreateFolder reportsFolder
   End If
   
   fileName = "\Report_Transactions_" & _
              Format(DateAdd("m", -1, Date), "YYYYMM") & "_" & Format(Date + Time, "YYYYMMHHNN") & ".xlsx"
   
   ReportPath = reportsFolder & fileName
   
   Set xlApp = CreateObject("Excel.Application")
   If xlApp Is Nothing Then
         
      Msg = "Não foi possível criar uma instância do Excel. Verifique se o Excel está instalado."
      rs.Close
      Set rs = Nothing
      ExportLastMonthToExcel = False
       
      Exit Function
   End If
   
   xlApp.Visible = False
   xlApp.DisplayAlerts = False
   xlApp.ScreenUpdating = False
   
   Set xlWb = xlApp.Workbooks.Add
   Set xlWs = xlWb.ActiveSheet
   xlWs.Name = "Transactions " & Format(DateAdd("m", -1, Date), "MMM YYYY")

   Dim headers(1 To 7) As String
   headers(1) = "ID"
   headers(2) = "Cartão"
   headers(3) = "Valor"
   headers(4) = "Data Hora"
   headers(5) = "Descrição"
   headers(6) = "Status"
   headers(7) = "Categorya"
   
   For col = 1 To 7
       xlWs.Cells(1, col).value = headers(col)
   Next col
   
   With xlWs.Range("A1:G1")
       .Font.Bold = True
       .Font.Size = 11
       .Interior.Color = RGB(0, 51, 102)
       .Font.Color = RGB(255, 255, 255)
       .HorizontalAlignment = xlCenter
       .Borders(xlEdgeBottom).LineStyle = xlContinuous
       .Borders(xlEdgeBottom).Weight = xlThin
   End With
   
   row = 2
   rs.MoveFirst
   
   Do While Not rs.EOF
       xlWs.Cells(row, 1).value = rs.Fields("TransactionId").value
       xlWs.Cells(row, 2).value = "'" & MaskCard(rs.Fields("CardNumber").value)
       xlWs.Cells(row, 3).value = CDbl(rs.Fields("TransactionAmount").value)
       xlWs.Cells(row, 4).value = CDate(rs.Fields("TransactionDate").value)
       xlWs.Cells(row, 5).value = rs.Fields("Description").value
       xlWs.Cells(row, 6).value = GetStatusLabel(rs.Fields("TransactionStatus").value)
       xlWs.Cells(row, 7).value = GetCategoryLabel(rs.Fields("Category").value)
   
       row = row + 1
       rs.MoveNext
   Loop
   
   rs.Close
   Set rs = Nothing
   
   xlWs.Columns("A").HorizontalAlignment = xlCenter
   xlWs.Columns("A").ColumnWidth = 8
   
   xlWs.Columns("B").ColumnWidth = 22
   
   xlWs.Columns("C").NumberFormat = "$#,##0.00"
   xlWs.Columns("C").HorizontalAlignment = xlCenter
   xlWs.Columns("C").ColumnWidth = 16
   
   xlWs.Columns("D").NumberFormat = "MM/DD/YYYY HH:MM"
   xlWs.Columns("D").HorizontalAlignment = xlCenter
   xlWs.Columns("D").ColumnWidth = 18
   
   xlWs.Columns("E").ColumnWidth = 35
   
   xlWs.Columns("F").HorizontalAlignment = xlCenter
   xlWs.Columns("F").ColumnWidth = 12
   
   xlWs.Columns("G").HorizontalAlignment = xlCenter
   xlWs.Columns("G").ColumnWidth = 12
   
   Dim lastRow As Long
   lastRow = row - 1
   
   If lastRow >= 2 Then
       With xlWs.Range("A2:G" & lastRow)
           .Borders(xlInsideHorizontal).LineStyle = xlContinuous
           .Borders(xlInsideHorizontal).Weight = xlThin
           .Borders(xlInsideHorizontal).Color = RGB(200, 200, 200)
           .Borders(xlInsideVertical).LineStyle = xlContinuous
           .Borders(xlInsideVertical).Weight = xlThin
           .Borders(xlInsideVertical).Color = RGB(200, 200, 200)
           .Borders(xlEdgeBottom).LineStyle = xlContinuous
           .Borders(xlEdgeBottom).Weight = xlThin
           .Borders(xlEdgeLeft).LineStyle = xlContinuous
           .Borders(xlEdgeLeft).Weight = xlThin
           .Borders(xlEdgeRight).LineStyle = xlContinuous
           .Borders(xlEdgeRight).Weight = xlThin
           .Borders(xlEdgeTop).LineStyle = xlContinuous
           .Borders(xlEdgeTop).Weight = xlThin
       End With
   
       Dim r As Long
       For r = 2 To lastRow
           If r Mod 2 = 0 Then
               xlWs.Rows(r).Interior.Color = RGB(240, 245, 250)
           End If
       Next r
   End If
   
   If lastRow >= 2 Then
       row = lastRow + 1
       xlWs.Cells(row, 1).value = ""
       xlWs.Cells(row, 2).value = "TOTAL"
       xlWs.Cells(row, 2).Font.Bold = True
       xlWs.Cells(row, 2).HorizontalAlignment = xlCenter
       xlWs.Cells(row, 3).Formula = "=SUM(C2:C" & lastRow & ")"
       xlWs.Cells(row, 3).NumberFormat = "$#,##0.00"
       xlWs.Cells(row, 3).Font.Bold = True
       xlWs.Cells(row, 3).HorizontalAlignment = xlCenter
       xlWs.Cells(row, 4).value = ""
       xlWs.Cells(row, 5).value = ""
       xlWs.Cells(row, 6).value = ""
       xlWs.Cells(row, 7).value = ""
   
       With xlWs.Range("A" & row & ":G" & row)
           .Borders(xlEdgeTop).LineStyle = xlContinuous
           .Borders(xlEdgeTop).Weight = 3
           .Interior.Color = RGB(230, 230, 230)
       End With
   End If

   xlApp.ActiveWindow.SplitRow = 1
   xlApp.ActiveWindow.FreezePanes = True

   xlWb.SaveAs ReportPath, xlWorkbookDefault

   xlWb.Close False
   xlApp.Quit

   Set xlWs = Nothing
   Set xlWb = Nothing
   Set xlApp = Nothing
   Set fso = Nothing
   
   Msg = "Relatório exportado com sucesso!"
   ExportLastMonthToExcel = True
   
   Exit Function

ErrorHandler:
   Call SaveLogErr("ExportLastMonthToExcel", Err.Description, "SQL: " & sql)
   ExportLastMonthToExcel = False
   
   On Error Resume Next
   If Not xlWb Is Nothing Then xlWb.Close False
   If Not xlApp Is Nothing Then xlApp.Quit
   Set xlWs = Nothing
   Set xlWb = Nothing
   Set xlApp = Nothing
   If Not rs Is Nothing Then
       If rs.State = adStateOpen Then rs.Close
       Set rs = Nothing
   End If
   On Error GoTo 0
End Function
