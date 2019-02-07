Attribute VB_Name = "AutoTunnel"
Option Explicit
Const MaxSectionNumber As Integer = 1000   '隧道最多断面数
Private Const IndexSheetName As String = "首页"

Dim ExcelSourceFileName As String 'excel源文件
Dim ReportTemplateFileName As String '"报告模板.docx"
Dim ReportFileName As String '自动生成的报告文件

Public OnMonitorSectionName(1 To MaxSectionNumber) As String    '在测隧道断面名称
Public OnMonitorSectionCount As Integer
Public SortedOnMonitorSectionSheetNumber(1 To MaxSectionNumber)    '排序后的断面标签序号

Public OnMonitorSectionSheet(1 To MaxSectionNumber, 1 To 2) As Integer
'在测隧道断面对应sheet的序号以及对应数值桩号
Public SortedOnMonitorSectionMiles(1 To MaxSectionNumber) As Integer    '排序后的截面公里数
Public wk    'workbook

Public Sub AutoTunnel()
    ExcelSourceFileName = Trim(CStr(Sheets(IndexSheetName).Cells(1, 2)))
    ReportTemplateFileName = Trim(CStr(Sheets(IndexSheetName).Cells(2, 2))) '"报告模板.docx"
    ReportFileName = Trim(CStr(Sheets(IndexSheetName).Cells(3, 2)))
    
    Dim orderType As Integer
    orderType = CInt(Sheets(IndexSheetName).Cells(4, 2))
    Dim picWidth As Double
    Dim picHeight As Double
    Dim autoPicSize As Boolean    '是否自动调整图片尺寸
    autoPicSize = True
    If Len(CStr(Sheets(IndexSheetName).Cells(5, 2))) <> 0 And Len(CStr(Sheets(IndexSheetName).Cells(5, 2))) <> 0 Then
        autoPicSize = False
        picWidth = CDbl(Sheets(IndexSheetName).Cells(5, 2))
        picHeight = CDbl(Sheets(IndexSheetName).Cells(6, 2))
    End If

    Dim wordApp As Word.Application
    Dim excelApp As Excel.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    Dim tbl As Table
    
    Dim ts As TunnelService
    Set ts = New TunnelService
    
    Dim i, j As Integer
    Dim resultFlag As Boolean   'True表示导出成功
    resultFlag = False
   
    Dim s As Shape
   
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & ReportTemplateFileName, ThisWorkbook.Path & "\AutoTunnelSource.docx"
    wordApp.Documents.Open ThisWorkbook.Path & "\AutoTunnelSource.docx"
    wordApp.Visible = False
   
    OnMonitorSectionCount = 0
    
    Set excelApp = New Excel.Application
    Set wk = excelApp.Workbooks.Open(ThisWorkbook.Path & "\" & ExcelSourceFileName, , True)   '只读方式打开
    excelApp.Visible = False
    
    For i = 1 To wk.Sheets.Count
        If ts.IsSheetOnMonitor(wk.Sheets(i)) Then
            OnMonitorSectionCount = OnMonitorSectionCount + 1
            SortedOnMonitorSectionSheetNumber(OnMonitorSectionCount) = i
            OnMonitorSectionSheet(OnMonitorSectionCount, 1) = i    '序号
            OnMonitorSectionSheet(OnMonitorSectionCount, 2) = ts.GetSheetSectionMile(wk.Sheets(i))   '数值桩号
            SortedOnMonitorSectionMiles(OnMonitorSectionCount) = OnMonitorSectionSheet(OnMonitorSectionCount, 2)
        End If
    Next i
    ts.ChooseSort SortedOnMonitorSectionMiles, 1, OnMonitorSectionCount, orderType
    
    'Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(Trim("table1")), NumRows:=10, NumColumns:=2)
    Set tbl = wordApp.ActiveDocument.Tables(1)

    Dim insertedCounts As Integer   '已插入的图片数量
    insertedCounts = 0
    
    '参考：http://club.excelhome.net/forum.php?mod=viewthread&tid=1278295&extra=page%3D1&page=1&
    '参考：opiona，北极狐工作室QQ：14885553
    FormProgressBar.Show 0   '//显示窗体
    
    FormProgressBar.Caption = "自动报告生成中，请稍候......"         '//根据实际情况改一下显示内容"

    
    '先粘帖"沉降累计"的图片
    For i = 1 To OnMonitorSectionCount
        For j = 1 To wk.Sheets(ts.FindOnMonitorSectionSheetNo(OnMonitorSectionSheet, SortedOnMonitorSectionMiles(i))).Shapes.Count
            Set s = wk.Sheets(ts.FindOnMonitorSectionSheetNo(OnMonitorSectionSheet, SortedOnMonitorSectionMiles(i))).Shapes(j)
            If autoPicSize = False Then
                s.Width = picWidth * (72 / 2.54)
                s.Height = picHeight * (72 / 2.54)
            End If
            's.Height = 142
            's.Width = 225.9
            If ts.IsSubsidence(s) = True Then
                insertedCounts = insertedCounts + 1
                s.CopyPicture
                tbl.Cell(ts.GetTableRow(insertedCounts), ts.GetTableCol(insertedCounts)).Range.Paste
                
                Set r = tbl.Cell(ts.GetTableRow(insertedCounts) + 1, ts.GetTableCol(insertedCounts)).Range
                r.MoveEnd , -1
                r.Text = "SEQ 附图 \* ARABIC"
                r.Fields.Add r, wdFieldEmpty, , False
                
                r.InsertBefore "附图"
                
                'Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldSequence, Text:="SEQ 附图 \* ARABIC"  ', PreserveFormatting:=True
                'tbl.Cell(ts.GetTableRow(insertedCounts) + 1, ts.GetTableCol(insertedCounts)).Range.Fields.Add tbl.Cell(ts.GetTableRow(insertedCounts) + 1, ts.GetTableCol(insertedCounts)).Range, -1, "SEQ 附图"
            End If
        Next j
            FormProgressBar.Label1.Width = Int(i / (2 * OnMonitorSectionCount) * 324) '//窗体设置：底色为红色，宽度自动代码设置
            FormProgressBar.Frame1.Caption = CStr(Round((i / (2 * OnMonitorSectionCount) * 100), 4)) & "%" '//显示内容，窗体设置：长度设为：324
            DoEvents
    Next i
    
    '再粘帖"收敛累计"的图片
    For i = 1 To OnMonitorSectionCount
        For j = 1 To wk.Sheets(ts.FindOnMonitorSectionSheetNo(OnMonitorSectionSheet, SortedOnMonitorSectionMiles(i))).Shapes.Count
            Set s = wk.Sheets(ts.FindOnMonitorSectionSheetNo(OnMonitorSectionSheet, SortedOnMonitorSectionMiles(i))).Shapes(j)
            If autoPicSize = False Then
                s.Width = picWidth
                s.Height = picHeight
            End If
            If ts.IsConvergence(s) = True Then
                insertedCounts = insertedCounts + 1
                s.CopyPicture
                tbl.Cell(ts.GetTableRow(insertedCounts), ts.GetTableCol(insertedCounts)).Range.Paste
                
                Set r = tbl.Cell(ts.GetTableRow(insertedCounts) + 1, ts.GetTableCol(insertedCounts)).Range
                r.MoveEnd , -1
                r.Text = "SEQ 附图 \* ARABIC"
                r.Fields.Add r, wdFieldEmpty, , False
                
                r.InsertBefore "附图"
            End If
        Next j
        FormProgressBar.Label1.Width = Int((OnMonitorSectionCount + i) / (2 * OnMonitorSectionCount) * 324) '//窗体设置：底色为红色，宽度自动代码设置
        FormProgressBar.Frame1.Caption = CStr(Round(((OnMonitorSectionCount + i) / (2 * OnMonitorSectionCount) * 100), 4)) & "%" '//显示内容，窗体设置：长度设为：324
        DoEvents
    Next i
        
    Unload FormProgressBar  '//关闭窗体
        
    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\" & ReportFileName
    resultFlag = True
CloseWord:
    wordApp.Documents.Close
    wordApp.Quit
    wk.Close SaveChanges:=False
    excelApp.Quit
    Set wordApp = Nothing
    Set wk = Nothing
    Set excelApp = Nothing
    Set tbl = Nothing
    Set s = Nothing
    Set ts = Nothing
    
    If resultFlag = True Then
        MsgBox "报告导出完成！"
    Else
        MsgBox "报告导出失败！"
    End If
End Sub


'对在测隧道断面进行排序
'从大桩号到小桩号进行排序
'截取后5位
'有1数组变量存储各个标签名称
Public Sub test()
  
    Dim i, j, k As Integer
    Dim ts As TunnelService
    Set ts = New TunnelService
    OnMonitorSectionCount = 0
    
    For i = 1 To ThisWorkbook.Sheets.Count
        If ts.IsSheetOnMonitor(ThisWorkbook.Sheets(i)) Then
            OnMonitorSectionCount = OnMonitorSectionCount + 1
            SortedOnMonitorSectionSheetNumber(OnMonitorSectionCount) = i
            Debug.Print SortedOnMonitorSectionSheetNumber(OnMonitorSectionCount)
            'Debug.Print ts.GetSheetSectionMile(ThisWorkbook.Sheets(i))
        End If
    Next i
    
    'Debug.Print OnMonitorSectionCount
    Set ts = Nothing
End Sub

