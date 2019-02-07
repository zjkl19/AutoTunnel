VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TunnelService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'标签是否为在测断面对应的标签
'判断依据
'标签名称从右往左数第4个字符为"+"
'标签没有颜色
Public Function IsSheetOnMonitor(ByRef s As Object) As Boolean
    If Left(Right(Trim(s.Name), 4), 1) = "+" And s.Tab.Color = False Then
        IsSheetOnMonitor = True
    Else
        IsSheetOnMonitor = False
    End If
End Function

'获取标签里程
'前提：必须是里程标签
Public Function GetSheetSectionMile(ByRef s As Object) As Integer
    GetSheetSectionMile = Int(Left(Right(Trim(s.Name), 5), 1)) * 1000 + Int(Right(Trim(s.Name), 3))
End Function

'计算第n张图片对应表格的行号
Public Function GetTableRow(ByVal n As Integer) As Integer
    If n Mod 2 = 0 Then
        GetTableRow = n - 1
    Else
        GetTableRow = n
    End If
End Function

'计算第n张图片对应表格的列号
Public Function GetTableCol(ByVal n As Integer) As Integer
    If n Mod 2 = 0 Then
        GetTableCol = 2
    Else
        GetTableCol = 1
    End If
End Function

'判断是否为沉降累计
Public Function IsSubsidence(ByRef s As Object) As Boolean
    If InStr(s.Chart.ChartTitle.Caption, "沉降累计") > 0 Then
        IsSubsidence = True
    Else
        IsSubsidence = False
    End If
    
End Function

'判断是否为沉降累计
Public Function IsConvergence(ByRef s As Object) As Boolean
    If InStr(s.Chart.ChartTitle.Caption, "收敛累计") > 0 Then
        IsConvergence = True
    Else
        IsConvergence = False
    End If
    
End Function

'选择排序
'仅支持1维数组
'数组元素为Integer
'lb:下界
'ub:上界
'orderType:0,从小到大;1,从大到小
Public Sub ChooseSort(a() As Integer, ByVal lb As Integer, ByVal ub As Integer, ByVal orderType As Integer)
   Dim i As Integer, j As Integer, temp As Integer
    For i = 1 To ub - 1
       For j = i + 1 To ub
          If orderType = 0 Then
                If a(i) > a(j) Then
                  temp = a(i)
                  a(i) = a(j)
                  a(j) = temp
                End If
         ElseIf orderType = 1 Then
                If a(i) < a(j) Then
                   temp = a(i)
                   a(i) = a(j)
                   a(j) = temp
                End If
         End If
       Next j
      Next i
End Sub

'查找公里数所对应的表标签
'查找到则返回对应下标，否则返回0
'OnMonitorSectionSheet二维数组
'milesMatch:欲匹配的公里数
'只匹配第一个查找到的
Public Function FindOnMonitorSectionSheetNo(OnMonitorSectionSheet() As Integer, ByVal milesMatch As Integer) As Integer
     Dim i, retNo As Integer
     retNo = 0
     For i = 1 To UBound(OnMonitorSectionSheet)
            If OnMonitorSectionSheet(i, 2) = milesMatch Then
                retNo = OnMonitorSectionSheet(i, 1)
            End If
            If retNo <> 0 Then
                Exit For
            End If
     Next
     FindOnMonitorSectionSheetNo = retNo
End Function

