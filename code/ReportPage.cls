VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear argb(255, 0, 0, 0) '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Page.Writes "摇号统计", RGW / 2, 70, 48, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    Page.Writes "绝对公平公正", RGW / 2, 140, 28, argb(150, 255, 255, 255), align:=StringAlignmentCenter
    
    Dim pro As Single

    pro = (GetTickCount - ReportTime) / 2000
    pro = Cubic(pro, 0, 1, 1, 1)
    
    If pro > 1 Then
        pro = 1
    End If
    
    Dim max As Single, Pitch As Single, Total As Long
    For I = 1 To 62
        If RCount(I) > max Then max = RCount(I)
        Pitch = (Pitch + RCount(I)) / 2
        Total = Total + RCount(I)
    Next
    
    Dim w As Long
    w = Int(RGW / 62)
    Pitch = Pitch * pro
    
    For I = 1 To 62
        Page.Paint 0, w * (I - 1), RGH - 200, w + 1, RCount(I) * pro / max * (RGH / 2), argb(RCount(I) * pro / max * 255, 51, 203, 186), pos:=posOnBottom
        If CheckMouse2 = mMouseUp Then
            SetWindowPos GameWindow.Hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
            MsgBox Student(0, I - 1) & " 摇号报告" & vbCrLf & _
                   "性别：" & Student(3, I - 1) & vbCrLf & _
                   "座号：" & Student(1, I - 1) & vbCrLf & _
                   "闪避加成：" & Val(Student(2, I - 1)) * 100 & "%" & vbCrLf & _
                   "等级：" & Int(RCount(I) / 5) & vbCrLf & _
                   "累计抽中次数：" & RCount(I) & vbCrLf & _
                   "抽中次数与平均抽中次数的比值：" & Int(RCount(I) / Pitch * 1000) / 1000 & "%" & vbCrLf & _
                   "本轮是否已经被抽中：" & IIf(Ignored(I), "是，在剩余的人抽完之前不会再被抽中。", "否，当心。你可能会被抽中。") & vbCrLf & _
                   "抽中概率：" & Int(RCount(I) / Total * 1000) / 1000 & "%", 48, "报告"
            SetWindowPos GameWindow.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        End If
        Page.Writes Student(0, I - 1), w * (I - 1), RGH - 200, 24, argb(RCount(I) * pro / max * 255, 255, 255, 255), w + 1, 200, StringAlignmentCenter, FontStyleBold
        Page.Writes Int(RCount(I) * pro), w * (I - 1), RGH - 200 - RCount(I) / max * (RGH / 2) * pro - 30, 16, argb(255, 255, 255, 255), w + 1, 30, align:=StringAlignmentCenter, style:=FontStyleBold
    Next
    Page.Paint 0, 0, RGH - 200 - Pitch / max * (RGH / 2), RGW + 1, 10, argb(100, 255, 255, 255)
    Page.Writes "平均" & Int(Pitch) & "次", RGW / 2, RGH - 200 - Pitch / max * (RGH / 2) + 10, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    
    If Mouse.state = 2 Then ECore.NewTransform , , "MainPage"
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(ID As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
    ReportTime = GetTickCount
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(ID As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.path & "\assets"
    '创建页面
    ECore.Add Page, "ReportPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

