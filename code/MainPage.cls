VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim Clicked As Boolean, ClickTime As Long
    Dim Touched As Boolean
    Dim RandomStyle As Integer
    Dim RandomPerson As Integer
    Dim RandomMax As Integer
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear 0 '清空画布

    Page.DrawImage "ui.png", RGW / 2, RGH / 2, alpha:=1, pos:=posOnCenter
    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    Page.Writes "高一五班专属摇号机", RGW / 2, RGH / 2 - 90, 24, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    Page.Writes "Version 3.0.0", RGW / 2, RGH / 2 - 50, 18, argb(150, 255, 255, 255), align:=StringAlignmentCenter
    
    Page.Writes "摇号方式", RGW / 2 - 250, RGH / 2 + 50, 18, argb(255, 0, 0, 0)
    Dim Ops(2) As String
    Ops(0) = "默认": Ops(1) = "排除男生": Ops(2) = "排除女生"
    For I = 0 To 2
        Page.Paint 0, RGW / 2 - 250 + 120 + I * 120 + 50 - 2, RGH / 2 + 63, 100, 30, IIf(RandomStyle = I, argb(255, 51, 203, 186), argb(255, 232, 232, 232)), pos:=posOnCenter
        If CheckMouse2 = mMouseUp Then RandomStyle = I
        Page.Writes Ops(I), RGW / 2 - 250 + 120 + I * 120 + 50, RGH / 2 + 50, 18, IIf(RandomStyle = I, argb(255, 255, 255, 255), argb(255, 128, 128, 128)), align:=StringAlignmentCenter
    Next
    
    Page.Writes "摇号人数", RGW / 2 - 250, RGH / 2 + 100, 18, argb(255, 0, 0, 0)
    Page.Writes RandomPerson + 1, RGW / 2 - 100 + 9 + 4, RGH / 2 + 100, 18, argb(255, 51, 203, 186), align:=StringAlignmentCenter
    Page.Paint 1, RGW / 2 - 140 + 4, RGH / 2 + 100 - 2, 30, 30, IIf(RandomPerson > 0, argb(255, 51, 203, 186), argb(255, 232, 232, 232))
    If CheckMouse2 = mMouseUp And RandomPerson > 0 Then RandomPerson = RandomPerson - 1
    Page.Writes "-", RGW / 2 - 140 + 8 + 4, RGH / 2 + 100, 18, IIf(RandomPerson > 0, argb(255, 255, 255, 255), argb(255, 32, 32, 32))
    Page.Paint 1, RGW / 2 - 72 + 4, RGH / 2 + 100 - 2, 30, 30, IIf(RandomPerson < RandomMax - 1, argb(255, 51, 203, 186), argb(255, 232, 232, 232))
    If CheckMouse2 = mMouseUp And RandomPerson < RandomMax - 1 Then RandomPerson = RandomPerson + 1
    Page.Writes "+", RGW / 2 - 72 + 5 + 4, RGH / 2 + 100, 18, IIf(RandomPerson < RandomMax - 1, argb(255, 255, 255, 255), argb(255, 32, 32, 32))
    
    Page.Writes "只凭运气：" & IIf(JustRandom, "开", "关"), RGW / 2 + 98, RGH / 2 + 100, 18, IIf(JustRandom, argb(255, 51, 203, 186), argb(255, 0, 0, 0)), 150, 25
    If CheckMouse2 = mMouseUp Then
        MusicList.Play "Option.wav"
        JustRandom = Not JustRandom
    End If
    Page.Writes "关闭音效：" & IIf(CloseSnd, "开", "关"), RGW / 2 + 98, RGH / 2 + 130, 18, IIf(CloseSnd, argb(255, 51, 203, 186), argb(255, 0, 0, 0)), 150, 25
    If CheckMouse2 = mMouseUp Then
        MusicList.Play "Option.wav"
        CloseSnd = Not CloseSnd
    End If
    
    Page.Writes "About", RGW / 2 - 280, RGH / 2 - 275, 18, argb(150, 255, 255, 255), 100, 25, StringAlignmentCenter
    If CheckMouse2 = mMouseUp Then
        MsgBox "高一五班专属摇号机" & vbCrLf & "版本号：3.0.0" & vbCrLf & vbCrLf & "Copyright 2019 CZY&LSM all rights reserved .", 48, "关于"
    End If
    
    Page.Writes "Hide", RGW / 2 + 185, RGH / 2 - 275, 18, argb(150, 255, 255, 255), 100, 25, StringAlignmentCenter
    If CheckMouse2 = mMouseUp Then ECore.NewTransform transFadeIn, , "FlyPage"
    
    Page.Writes "Report", RGW / 2 + 105, RGH / 2 - 275, 18, argb(150, 255, 255, 255), 80, 25, StringAlignmentCenter
    If CheckMouse2 = mMouseUp Then ECore.NewTransform transFadeIn, , "ReportPage"
    
    Page.DrawImage "go.png", RGW / 2, RGH / 2 + 200, alpha:=1, pos:=posOnCenter
    If CheckMouse2 = mMouseUp Then
        RandomTime = GetTickCount
        Erase Beeps
        Call StartRandom
        ClickTime = GetTickCount: Clicked = True
        ReDim Person(RandomPerson)
        For I = 0 To UBound(Person)
            Person(I) = GetRandom(RandomStyle)
        Next
        ECore.NewTransform transFadeIn, , "PersonPage"
    End If

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
    Page.Res.NewImages App.path & "\assets", 0.75
    '创建页面
    ECore.Add Page, "MainPage"
    Dim WCount As Long, HCount As Long
    WCount = Int(RGW / (400 + 50))
    HCount = Int(RGH / (220 + 50))
    RandomMax = WCount * HCount
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
