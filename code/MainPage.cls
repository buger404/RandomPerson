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
    Dim Beeps(3) As Boolean
    Dim Touched As Boolean
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear 0 '清空画布

    Page.DrawImage "ui.png", RGW / 2, RGH / 2, alpha:=1, pos:=posOnCenter
    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    Page.Writes "高一五班专属摇号机", RGW / 2 - 100 + 2, 60, 18, argb(100, 0, 0, 0)
    If Index <> -1 Then
        Page.Writes "号数 " & Student(1, RIndex), RGW / 2 - 100, RGH / 2 - 48 * 2.5 + 75, 24, argb(150, 0, 0, 0)
        If Clicked Then
            Dim Name As String, j As Integer
            If GetTickCount - ClickTime <= 1000 Then
                Name = "？？？"
            Else
                For j = 1 To 3
                    If GetTickCount - ClickTime >= 1000 + j * 333 Then
                        Name = Name & Mid(Student(0, RIndex), j, 1)
                        If Not Beeps(j) Then Beeps(j) = True: win32api.Beep 800 + j * 400, 200
                    Else
                        Name = Name & "？"
                    End If
                Next
            End If
            Page.Writes "幸运小朋友 " & Name, RGW / 2 - 100 - 3, RGH / 2 - 48 * 1.8 + 75, 36, argb(255, 32, 32, 32)
        Else
            Page.Writes "幸运小朋友 " & Student(0, RIndex), RGW / 2 - 100 - 3, RGH / 2 - 48 * 1.8 + 75, 36, argb(255, 32, 32, 32)
        End If
    End If
    
    If Clicked Then
        If GetTickCount - ClickTime <= 1000 Then Call GetRandom
        If GetTickCount - ClickTime >= 2000 Then
            Call DoneRandom
            Clicked = False
        End If
    End If
    
    Page.DrawImage "close.png", RGW / 2, RGH - 40, alpha:=1, pos:=posOnCenter
    If CheckMouse2 = mMouseUp Then GameWindow.WindowState = 1
    
    If CheckMouse(0, 0, GW, GH) = mMouseUp Then
        Erase Beeps
        Call StartRandom
        ClickTime = GetTickCount: Clicked = True
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
    Page.Res.NewImages App.path & "\assets", 0.65
    '创建页面
    ECore.Add Page, "MainPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
