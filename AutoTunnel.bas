Attribute VB_Name = "AutoTunnel"
Option Explicit
Const MaxSectionNumber As Integer = 1000   '�����������
Private Const IndexSheetName As String = "��ҳ"

Dim ExcelSourceFileName As String 'excelԴ�ļ�
Dim ReportTemplateFileName As String '"����ģ��.docx"
Dim ReportFileName As String '�Զ����ɵı����ļ�

Public OnMonitorSectionName(1 To MaxSectionNumber) As String    '�ڲ������������
Public OnMonitorSectionCount As Integer
Public SortedOnMonitorSectionSheetNumber(1 To MaxSectionNumber)    '�����Ķ����ǩ���

Public OnMonitorSectionSheet(1 To MaxSectionNumber, 1 To 2) As Integer
'�ڲ���������Ӧsheet������Լ���Ӧ��ֵ׮��
Public SortedOnMonitorSectionMiles(1 To MaxSectionNumber) As Integer    '�����Ľ��湫����
Public wk    'workbook

Public Sub AutoTunnel()
    ExcelSourceFileName = Trim(CStr(Sheets(IndexSheetName).Cells(1, 2)))
    ReportTemplateFileName = Trim(CStr(Sheets(IndexSheetName).Cells(2, 2))) '"����ģ��.docx"
    ReportFileName = Trim(CStr(Sheets(IndexSheetName).Cells(3, 2)))
    
    Dim orderType As Integer
    orderType = CInt(Sheets(IndexSheetName).Cells(4, 2))
    Dim picWidth As Double
    Dim picHeight As Double
    Dim autoPicSize As Boolean    '�Ƿ��Զ�����ͼƬ�ߴ�
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
    Dim resultFlag As Boolean   'True��ʾ�����ɹ�
    resultFlag = False
   
    Dim s As Shape
   
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & ReportTemplateFileName, ThisWorkbook.Path & "\AutoTunnelSource.docx"
    wordApp.Documents.Open ThisWorkbook.Path & "\AutoTunnelSource.docx"
    wordApp.Visible = False
   
    OnMonitorSectionCount = 0
    
    Set excelApp = New Excel.Application
    Set wk = excelApp.Workbooks.Open(ThisWorkbook.Path & "\" & ExcelSourceFileName, , True)   'ֻ����ʽ��
    excelApp.Visible = False
    
    For i = 1 To wk.Sheets.Count
        If ts.IsSheetOnMonitor(wk.Sheets(i)) Then
            OnMonitorSectionCount = OnMonitorSectionCount + 1
            SortedOnMonitorSectionSheetNumber(OnMonitorSectionCount) = i
            OnMonitorSectionSheet(OnMonitorSectionCount, 1) = i    '���
            OnMonitorSectionSheet(OnMonitorSectionCount, 2) = ts.GetSheetSectionMile(wk.Sheets(i))   '��ֵ׮��
            SortedOnMonitorSectionMiles(OnMonitorSectionCount) = OnMonitorSectionSheet(OnMonitorSectionCount, 2)
        End If
    Next i
    ts.ChooseSort SortedOnMonitorSectionMiles, 1, OnMonitorSectionCount, orderType
    
    'Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(Trim("table1")), NumRows:=10, NumColumns:=2)
    Set tbl = wordApp.ActiveDocument.Tables(1)

    Dim insertedCounts As Integer   '�Ѳ����ͼƬ����
    insertedCounts = 0
    
    '�ο���http://club.excelhome.net/forum.php?mod=viewthread&tid=1278295&extra=page%3D1&page=1&
    '�ο���opiona��������������QQ��14885553
    FormProgressBar.Show 0   '//��ʾ����
    
    FormProgressBar.Caption = "�Զ����������У����Ժ�......"         '//����ʵ�������һ����ʾ����"

    
    '��ճ��"�����ۼ�"��ͼƬ
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
                r.Text = "SEQ ��ͼ \* ARABIC"
                r.Fields.Add r, wdFieldEmpty, , False
                
                r.InsertBefore "��ͼ"
                
                'Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldSequence, Text:="SEQ ��ͼ \* ARABIC"  ', PreserveFormatting:=True
                'tbl.Cell(ts.GetTableRow(insertedCounts) + 1, ts.GetTableCol(insertedCounts)).Range.Fields.Add tbl.Cell(ts.GetTableRow(insertedCounts) + 1, ts.GetTableCol(insertedCounts)).Range, -1, "SEQ ��ͼ"
            End If
        Next j
            FormProgressBar.Label1.Width = Int(i / (2 * OnMonitorSectionCount) * 324) '//�������ã���ɫΪ��ɫ������Զ���������
            FormProgressBar.Frame1.Caption = CStr(Round((i / (2 * OnMonitorSectionCount) * 100), 4)) & "%" '//��ʾ���ݣ��������ã�������Ϊ��324
            DoEvents
    Next i
    
    '��ճ��"�����ۼ�"��ͼƬ
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
                r.Text = "SEQ ��ͼ \* ARABIC"
                r.Fields.Add r, wdFieldEmpty, , False
                
                r.InsertBefore "��ͼ"
            End If
        Next j
        FormProgressBar.Label1.Width = Int((OnMonitorSectionCount + i) / (2 * OnMonitorSectionCount) * 324) '//�������ã���ɫΪ��ɫ������Զ���������
        FormProgressBar.Frame1.Caption = CStr(Round(((OnMonitorSectionCount + i) / (2 * OnMonitorSectionCount) * 100), 4)) & "%" '//��ʾ���ݣ��������ã�������Ϊ��324
        DoEvents
    Next i
        
    Unload FormProgressBar  '//�رմ���
        
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
        MsgBox "���浼����ɣ�"
    Else
        MsgBox "���浼��ʧ�ܣ�"
    End If
End Sub


'���ڲ���������������
'�Ӵ�׮�ŵ�С׮�Ž�������
'��ȡ��5λ
'��1��������洢������ǩ����
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

