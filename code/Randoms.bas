Attribute VB_Name = "Randoms"
'=====================================================================================
'   ҡ�ź���ģ��
'   Maker��LSM/CZY
'=====================================================================================
'   ͨ�ò���
    Public Student() As String
    '����ͬѧ��Student(x,y)
    '��һ��(x=1)Ϊ���֣��ڶ���(x=2)Ϊ���ţ�������(x=3)Ϊ���ܸ���
    
    Public Ignored(62) As Boolean   '���ӱ��
    Public Sticks() As Integer      '���Գ�ȡ��ͬѧ����ӦStudent��y
    Public RIndex As Integer        '��ǰ�鵽�ĺ���
    Public Voice As Object          '��������
'=====================================================================================
Public Sub Speak(ByVal Content As String)
    Voice.Speak Content, 1          '����1��ʾ�첽����
End Sub
Public Sub Start()
    '��������
    
    '���ļ�ȡ�ú��Ӽ�¼
    Open App.path & "\ignored.stulist" For Binary As #1
    Get #1, , Ignored
    Close #1
    
    '������������
    Set Voice = CreateObject("SAPI.SpVoice")
    Voice.Volume = 100
End Sub
Public Sub StartRandom()
    'ҡ��׼��
    ReDim Sticks(0)
    '�������п��Ա����ѧ��
    For I = 1 To 62
        If I <> 39 Then '�����
            'û�б�����
            If Not Ignored(I) Then
                ReDim Preserve Sticks(UBound(Sticks) + 1)
                Sticks(UBound(Sticks)) = I
            End If
        End If
    Next
End Sub
Public Sub GetRandom()
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
        IgnoredSomebody RIndex
        GoTo Miss
    End If
End Sub
Public Sub DoneRandom()
    'ҡ�����
    IgnoredSomebody RIndex
    '��������С����
    Speak "��ϲ" & Student(0, RIndex)
End Sub
Public Sub IgnoredSomebody(Index As Integer)
    '����ĳ��
    Ignored(Index) = True
    '�ж��Ƿ��Ѿ�ȫ������
    Dim AllIgnore As Boolean
    AllIgnore = True
    For I = 1 To 62
        If I <> 39 Then '�����
            If Not Ignored(I) Then AllIgnore = False: Exit For
        End If
    Next
    '����Ѿ�ȫ������
    Erase Ignored()
    '�����ļ�
    Open App.path & "\ignored.stulist" For Binary As #1
    Put #1, , Ignored
    Close #1
End Sub
