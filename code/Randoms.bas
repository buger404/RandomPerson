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
    Public AntiDouble(62) As Boolean
'=====================================================================================
Public Function IsInClass(TimeShift As Long) As Boolean
    Dim DTime As Long, Ret As Boolean, NowTime As Long
    DTime = TimeShift
    'DTime = Hour(Now) * 60 + Minute(Now)
    
    '���
    NowTime = (7 * 60) + 20
    If DTime >= NowTime - 3 And DTime < NowTime + 40 Then Ret = True
    '��һ�ڿ�
    NowTime = NowTime + 5 + 20
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    '�ڶ��ڿ�
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    '�����ڿΣ���μ䣩
    NowTime = NowTime + IIf(Weekday(Now) = 6, 10, 25) + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    '���Ľڿ�
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    '����ڿΣ���ʱѵ����
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    
    '����
    '����
    NowTime = (14 * 60) + 0
    If DTime >= NowTime And DTime < NowTime + 20 Then Ret = True
    '��һ�ڿ�
    NowTime = (14 * 60) + 20
    If DTime >= NowTime - 3 And DTime < NowTime + 40 Then Ret = True
    '�ڶ��ڿ�
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    '�����ڿ�
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    '���ĽڿΣ���ʱѵ����
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    
    '����ϰ
    '��һ�ڿ�
    NowTime = (18 * 60) + 20
    If DTime >= NowTime - 3 And DTime < NowTime + 40 Then Ret = True
    '�ڶ��ڿ�
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 50 Then Ret = True
    '�����ڿ�
    NowTime = NowTime + 10 + 50
    If DTime >= NowTime - 2 And DTime < NowTime + 50 Then Ret = True
    '���ĽڿΣ�������ϰ��
    NowTime = NowTime + 10 + 50
    If DTime >= NowTime - 2 And DTime < NowTime + 50 Then Ret = True
    '����ϰ�¿�
    NowTime = NowTime + 50
    If DTime >= NowTime And DTime < NowTime + 20 Then Ret = True
    
    IsInClass = Ret Or (App.LogMode = 0)
End Function
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
    Erase AntiDouble
End Sub
Public Function GetRandom(Filter As Integer) As Integer
    'ҡ��׼��
    ReDim Sticks(0)
    '�������п��Ա����ѧ��
    For I = 1 To 62
        If I <> 39 Then '�����
            'û�б�����
            If ((Not Ignored(I)) Or JustRandom) And AntiDouble(I) = False Then
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
    Dim Index As Integer
    Index = Int(Rnd * UBound(Sticks) + 1)
    '��ֹ�±�Խ��
    If Index > UBound(Sticks) Then Index = UBound(Sticks)
    RIndex = Sticks(Index) - 1
    
    If Rnd < Val(Student(2, RIndex)) Then
        '���ܳɹ�
        IgnoredSomebody RIndex + 1
        GoTo Miss
    End If
    
    IgnoredSomebody RIndex + 1
    AntiDouble(RIndex + 1) = True
    
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
Public Sub CheckIgnored(Filter As Integer, Needed As Long)
    Dim Check As Boolean, Count As Long
    Count = Needed
    For I = 1 To 62
        If I <> 39 Then '�����
            If Not Ignored(I) Then
                Check = True
                If Filter = 1 Then Check = Check And (Student(3, I - 1) = "Ů")
                If Filter = 2 Then Check = Check And (Student(3, I - 1) = "��")
                If Check Then Count = Count - 1
            End If
        End If
    Next
    If Count > 0 Then
        '�����飬��Ҫ����
        For I = 1 To 62
            If I <> 39 Then '�����
                Check = True
                If Filter = 1 Then Check = Check And (Student(3, I - 1) = "Ů")
                If Filter = 2 Then Check = Check And (Student(3, I - 1) = "��")
                If Check Then Ignored(I) = False
            End If
        Next
        Open App.path & "\ignored.stulist" For Binary As #1
        Put #1, , Ignored
        Close #1
    End If
End Sub
Public Sub IgnoredSomebody(Index As Integer)
    '����ĳ��
    Ignored(Index) = True
    RCount(Index) = RCount(Index) + 1
    '�����ļ�
    Open App.path & "\ignored.stulist" For Binary As #1
    Put #1, , Ignored
    Close #1
    Open App.path & "\count.stulist" For Binary As #1
    Put #1, , RCount
    Close #1
End Sub
