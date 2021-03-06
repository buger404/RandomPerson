VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlyPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim IX As Long, IY As Long
    Dim sx As Long, sy As Long, OX As Long, OY As Long
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear 0 '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Page.DrawImage "image.png", IX, IY, alpha:=1, pos:=posOnCenter

    If CheckMouse2 >= mMouseDown Then
        Dim p As POINTAPI
        GetCursorPos p
        If sx = -1 Then
            sx = p.X - IX: sy = p.y - IY: OX = IX: OY = IY
            Do While Mouse.state <> 2 And Mouse.state <> 0
                GetCursorPos p
                IX = p.X - sx: IY = p.y - sy
                ECore.Display: DoEvents
            Loop
            If Abs(OX - IX) < 10 And Abs(OY - IY) < 10 Then
                MusicList.Play "Option.wav"
                ECore.NewTransform transFadeIn, , "MainPage"
            End If
            sx = -1
        End If
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
    Page.Res.NewImages App.path & "\assets\window", 0.7
    '创建页面
    ECore.Add Page, "FlyPage"
    IX = RGW - 100
    IY = RGH - 150
    sx = -1
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
