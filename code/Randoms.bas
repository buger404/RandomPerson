Attribute VB_Name = "Randoms"
'=====================================================================================
'   ҡ�ź���ģ��
'   Maker��LSM/CZY
'=====================================================================================
'   ͨ�ò���
    Public Student() As String
    '����ͬѧ��Student(x,y)
    '��һ��(x=1)Ϊ���֣��ڶ���(x=2)Ϊ���ţ�������(x=3)Ϊ���ܸ���
    Public Person() As Integer      '���
    Public Ignored(62) As Boolean   '���ӱ��
    Public RCount(62) As Long       '���д���
    Public Sticks() As Integer      '���Գ�ȡ��ͬѧ����ӦStudent��y
    Public RIndex As Integer        '��ǰ�鵽�ĺ���
    Public Voice As Object          '��������
    Public RandomTime As Long       '�齱ʱ��
    Public RandomDone As Boolean
    Public JustRandom As Boolean
    Public Beeps(3) As Boolean
    Public CloseSnd As Boolean
    Public ReportTime As Long
'=====================================================================================
Public Sub Speak(ByVal Content As String)
    If CloseSnd Then Exit Sub
    If Not Voice Is Nothing Then Voice.Speak Content, 1          '����1��ʾ�첽����
End Sub
Public Sub Start()
    '��������
    
    '���ļ�ȡ�ú��Ӽ�¼
    Open App.path & "\ignored.stulist" For Binary As #1
    Get #1, , Ignored
    Close #1
    Open App.path & "\count.stulist" For Binary As #1
    Get #1, , RCount
    Close #1
    
    '������������
    On Error Resume Next
    Set Voice = CreateObject("SAPI.SpVoice")
    Voice.Volume = 100
End Sub
Public Sub StartRandom()
    RandomDone = False
    MusicList.Play "Done.mp3"
End Sub
Public Function GetRandom(Filter As Integer) As Integer
    'ҡ��׼��
    ReDim Sticks(0)
    '�������п��Ա����ѧ��
    For I = 1 To 62
        If I <> 39 Then '�����
            'û�б�����
            If (Not Ignored(I)) Or JustRandom Then
                If Filter = 1 And Student(3, I - 1) = "��" Then GoTo SkipThis
                If Filter = 2 And Student(3, I - 1) = "Ů" Then GoTo SkipThis
                ReDim Preserve Sticks(UBound(Sticks) + 1)
                Sticks(UBound(Sticks)) = I
SkipThis:
            End If
        End If
    Next
    'ҡ����
    Randomize
Miss:
    '+1����ΪSticks(0)=0��˳���ֹ�鵽62�ŵļ��ʹ�С
    Dim index As Integer
    index = Int(Rnd * UBound(Sticks) + 1)
    '��ֹ�±�Խ��
    If index > UBound(Sticks) Then index = UBound(Sticks)
    RIndex = Sticks(index) - 1
    
    If Rnd < Val(Student(2, RIndex)) Then
        '���ܳɹ�
        IgnoredSomebody RIndex + 1
        GoTo Miss
    End If
    
    IgnoredSomebody RIndex + 1
    GetRandom = RIndex
End Function
Public Sub DoneRandom()
    'ҡ�����
    MusicList.Play "Papa.mp3"
    '��������С����
    Dim Ret As String, Bing As Boolean
    For I = 0 To UBound(Person)
        Ret = Ret & Student(0, Person(I)) & " "
        If RCount(Person(I) + 1) Mod 5 = 0 Then Bing = True
    Next
    If Bing Then MusicList.Play "LevelUp.mp3"
    Speak "��ϲ" & Ret
End Sub
Public Sub IgnoredSomebody(index As Integer)
    '����ĳ��
    Ignored(index) = True
    RCount(index) = RCount(index) + 1
    '�ж��Ƿ��Ѿ�ȫ������
    Dim AllIgnoreBoy As Boolean, AllIgnoreGirl As Boolean
    AllIgnoreBoy = True: AllIgnoreGirl = True
    For I = 1 To 62
        If I <> 39 Then '�����
            If (Not Ignored(I)) And Student(3, I - 1) = "��" Then AllIgnoreBoy = False: Exit For
            If (Not Ignored(I)) And Student(3, I - 1) = "Ů" Then AllIgnoreGirl = False: Exit For
        End If
    Next
    '����Ѿ�ȫ������
    If AllIgnoreBoy Or AllIgnoreGirl Then Erase Ignored()
    '�����ļ�
    Open App.path & "\ignored.stulist" For Binary As #1
    Put #1, , Ignored
    Close #1
    Open App.path & "\count.stulist" For Binary As #1
    Put #1, , RCount
    Close #1
End Sub
