VERSION 5.00
Begin VB.Form GameWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "�齱"
   ClientHeight    =   6684
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   557
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9000
      Top             =   240
   End
End
Attribute VB_Name = "GameWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   ҳ�������
    Dim EC As GMan
'==================================================
'   �ڴ˴��������ҳ����ģ������
    Dim MainPage As MainPage
    Dim FlyPage As FlyPage
    Dim PersonPage As PersonPage
    Dim ReportPage As ReportPage
'==================================================

Private Sub DrawTimer_Timer()
    '����
    EC.Display
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ�����
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub

Private Sub Form_Load()
    '��ʼ��Emerald���ڴ˴������޸Ĵ��ڴ�СӴ~��
    StartEmerald Me.Hwnd, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY
    'ScaleGame 805 / 1326, ScaleSuitable
    '��������
    MakeFont "΢���ź�"
    '����ҳ�������
    Set EC = New GMan
    EC.Layered False
    '�����浵����ѡ��
    'Set ESave = New GSaving
    'ESave.Create "emerald.test", "Emerald.test"
    
    '���������б�
    Set MusicList = New GMusicList
    MusicList.Create App.path & "\music"

    '��ʼ��ʾ
    Me.Show
    DrawTimer.Enabled = True
    
    '�ڴ˴���ʼ�����ҳ��
    '=============================================
    'ʾ����TestPage.cls
    '     Set TestPage = New TestPage
    '�������֣�Dim TestPage As TestPage
        Set MainPage = New MainPage
        Set FlyPage = New FlyPage
        Set PersonPage = New PersonPage
        Set ReportPage = New ReportPage
    '=============================================

    HideLOGO = 1
    DisableLOGO = 1
    
    '���ûҳ��
    EC.ActivePage = "FlyPage"
    
    If App.LogMode <> 0 Then SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Student = ReadExcel(App.path & "\Person.xls")
    Call Start
    Piano.Init
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    UpdateMouse X, Y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    If Mouse.state = 0 Then
        UpdateMouse X, Y, 0, button
    Else
        Mouse.X = X: Mouse.Y = Y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    '���������Ϣ
    UpdateMouse X, Y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '��ֹ����
    Piano.Dispose
    DrawTimer.Enabled = False
    '�ͷ�Emerald��Դ
    EndEmerald
End Sub
