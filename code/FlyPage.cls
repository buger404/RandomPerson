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
'   ҳ�������
    Dim Page As GPage
    Dim IX As Long, IY As Long
    Dim sx As Long, sy As Long, OX As Long, OY As Long
'======================================

Public Sub Update()
    '��Ϸÿһ֡�Ĺ��̣����ƹ�������ڴ˴���
    
    Page.Clear 0 '��ջ���

    '================================================================
    '   ʾ������������
    '   Page.Writes "��������", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Page.DrawImage "image.png", IX, IY, alpha:=1, pos:=posOnCenter

    If CheckMouse2 >= mMouseDown Then
        MusicList.Play "Option.wav"
        Dim p As POINTAPI
        GetCursorPos p
        If sx = -1 Then
            sx = p.X - IX: sy = p.y - IY: OX = IX: OY = IY
            Do While Mouse.state <> 2 And Mouse.state <> 0
                GetCursorPos p
                IX = p.X - sx: IY = p.y - sy
                ECore.Display: DoEvents
            Loop
            If Abs(OX - IX) < 10 And Abs(OY - IY) < 10 Then ECore.NewTransform transFadeIn, , "MainPage"
            sx = -1
        End If
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '�������¼�
    'Direction:����Depth�����
End Sub
Public Sub AnimationMsg(ID As String, msg As String)
    '������Ϣ����
End Sub
Public Sub Enter()
    'ҳ������¼�
End Sub
Public Sub Leave()
    'ҳ���뿪�¼�
End Sub
Public Sub Gotfocus()
    '����ȡ�ý����¼�
End Sub
Public Sub Lostfocus()
    '����ʧȥ�����¼�
End Sub
Public Sub AnimationDone(ID As Integer)
    '�������Ž����¼�
    'id�������Ķ���ID
End Sub
Private Sub Class_Initialize()
    '����ҳ�������
    Set Page = New GPage
    Page.Create Me
    '������Ϸ��Դ
    Page.Res.NewImages App.path & "\assets\window", 0.7
    '����ҳ��
    ECore.Add Page, "FlyPage"
    IX = RGW - 90
    IY = RGH - 90
    sx = -1
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub