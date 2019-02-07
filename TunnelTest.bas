Attribute VB_Name = "TunnelTest"
Option Explicit

Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Dim excelApp As Excel.Application
Dim testExcel As String
Dim wk As Excel.Workbook
Private Const TestConfigSheet As String = "开发者专区"

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

Public Sub LoadTestFile()
    Set excelApp = New Excel.Application
    testExcel = CStr(ThisWorkbook.Sheets(TestConfigSheet).Cells(1, 2))
    'testExcel = "5、1倪家山-左洞 - 0124 - 测试.xls"
    Set wk = excelApp.Workbooks.Open(ThisWorkbook.Path & "\" & testExcel, , True)    '只读方式打开
    excelApp.Visible = False
End Sub

Public Sub CloseTestFile()
    wk.Close SaveChanges:=False
    excelApp.Quit
    Set wk = Nothing
    Set excelApp = Nothing
End Sub

'@TestMethod
Public Sub GetSheetSectionMile_Tests()
    On Error GoTo TestFail
    
    'Arrange:
    LoadTestFile
    Dim ts As TunnelService
    Dim m As Integer
    Set ts = New TunnelService
    'Act:
    m = ts.GetSheetSectionMile(wk.Sheets("M14、ZK8+618"))
    'Assert:
    Assert.AreEqual m, 8618, "Not Equal"
    Set ts = Nothing
    CloseTestFile
    Exit Sub
TestFail:
    Set ts = Nothing
    CloseTestFile
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
'Assert.Istrue:从右往左数第4个是"+",无色
'Assert.False:红色的
'Assert.False:从右往左数第4个不是"+"
Public Sub IsSheetOnMonitor_Tests()
     On Error GoTo TestFail
    'Arrange:

    LoadTestFile
    Dim ts As TunnelService
    Dim b1, b2 As Boolean
    Set ts = New TunnelService
    'Act:
    
    'Assert:
    Assert.IsTrue ts.IsSheetOnMonitor(wk.Sheets("M14、ZK8+618"))
    Assert.IsFalse ts.IsSheetOnMonitor(wk.Sheets("M13、ZK8+648"))
    Assert.IsFalse ts.IsSheetOnMonitor(wk.Sheets("18下"))
    Set ts = Nothing
    CloseTestFile
    Exit Sub
TestFail:
    Set ts = Nothing
    CloseTestFile
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub GetTableRow_Tests()
     On Error GoTo TestFail
    'Arrange:
    Dim ts As TunnelService
    Set ts = New TunnelService
    'Act:
    
    'Assert:
    Assert.AreEqual ts.GetTableRow(5), 5
    Assert.AreEqual ts.GetTableRow(6), 5
    Set ts = Nothing
    Exit Sub
TestFail:
    Set ts = Nothing
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub GetTableCol_Tests()
     On Error GoTo TestFail
    'Arrange:
    Dim ts As TunnelService
    Set ts = New TunnelService
    'Act:
    
    'Assert:
    Assert.AreEqual ts.GetTableCol(5), 1
    Assert.AreEqual ts.GetTableCol(6), 2
    Set ts = Nothing
    Exit Sub
TestFail:
    Set ts = Nothing
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
'奇数个排序测试
'偶数个排序测试
'降序测试
Public Sub ChooseSort_Tests()
     On Error GoTo TestFail
    'Arrange:
    Dim ts As TunnelService
    Set ts = New TunnelService
    Dim a1(1 To 5) As Integer    '奇数个排序测试
    Dim a2(1 To 6) As Integer    '偶数个排序测试
    Dim a3(1 To 6) As Integer    '降序测试
    
    'Act:
    a1(1) = 3: a1(2) = 2: a1(3) = 1: a1(4) = 5: a1(5) = 4
    ts.ChooseSort a1, 1, 5, 0
    a2(1) = 3: a2(2) = 2: a2(3) = 1: a2(4) = 6: a2(5) = 4: a2(6) = 5
    ts.ChooseSort a2, 1, 6, 0
    a3(1) = 3: a3(2) = 2: a3(3) = 1: a3(4) = 6: a3(5) = 4: a3(6) = 5
    ts.ChooseSort a3, 1, 6, 1
    'Assert:
    Assert.AreEqual a1(1), 1
    Assert.AreEqual a1(2), 2
    Assert.AreEqual a1(3), 3
    Assert.AreEqual a1(4), 4
    Assert.AreEqual a1(5), 5
    
    Assert.AreEqual a2(1), 1
    Assert.AreEqual a2(2), 2
    Assert.AreEqual a2(3), 3
    Assert.AreEqual a2(4), 4
    Assert.AreEqual a2(5), 5
    Assert.AreEqual a2(6), 6
    
    Assert.AreEqual a3(1), 6
    Assert.AreEqual a3(2), 5
    Assert.AreEqual a3(3), 4
    Assert.AreEqual a3(4), 3
    Assert.AreEqual a3(5), 2
    Assert.AreEqual a3(6), 1
    Set ts = Nothing
    Exit Sub
TestFail:
    Set ts = Nothing
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
'查找到，返回正确的下标，否则返回0
Public Sub FindOnMonitorSectionSheetNo_Tests()
     On Error GoTo TestFail
    'Arrange:
    Dim ts As TunnelService
    Set ts = New TunnelService
    Dim OnMonitorSectionSheet(1 To 1000, 2) As Integer
    OnMonitorSectionSheet(1, 1) = 3: OnMonitorSectionSheet(1, 2) = 8618
    OnMonitorSectionSheet(2, 1) = 5: OnMonitorSectionSheet(2, 2) = 8588
    OnMonitorSectionSheet(3, 1) = 7: OnMonitorSectionSheet(3, 2) = 8558
    OnMonitorSectionSheet(4, 1) = 9: OnMonitorSectionSheet(4, 2) = 8528
    OnMonitorSectionSheet(5, 1) = 10: OnMonitorSectionSheet(5, 2) = 8498
    OnMonitorSectionSheet(6, 1) = 13: OnMonitorSectionSheet(6, 2) = 8478
    
    'Act:
     
    'Assert:
    Assert.AreEqual ts.FindOnMonitorSectionSheetNo(OnMonitorSectionSheet, 8588), 5
    Assert.AreEqual ts.FindOnMonitorSectionSheetNo(OnMonitorSectionSheet, 8528), 9
    Assert.AreEqual ts.FindOnMonitorSectionSheetNo(OnMonitorSectionSheet, 8000), 0
    
    Set ts = Nothing
    Exit Sub
TestFail:
    Set ts = Nothing
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


