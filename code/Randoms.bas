Attribute VB_Name = "Randoms"
'=====================================================================================
'   摇号核心模块
'   Maker：LSM/CZY
'=====================================================================================
'   通用部分
    Public Student() As String
    '所有同学：Student(x,y)
    '第一列(x=1)为名字，第二列(x=2)为座号，第三列(x=3)为闪避概率
    Public Person() As Integer      '多抽
    Public Ignored(62) As Boolean   '忽视标记
    Public RCount(62) As Long       '抽中次数
    Public Sticks() As Integer      '可以抽取的同学，对应Student的y
    Public RIndex As Integer        '当前抽到的号数
    Public Voice As Object          '语音对象
    Public RandomTime As Long       '抽奖时间
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
    
    '早读
    NowTime = (7 * 60) + 20
    If DTime >= NowTime - 3 And DTime < NowTime + 40 Then Ret = True
    '第一节课
    NowTime = NowTime + 5 + 20
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    '第二节课
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    '第三节课（大课间）
    NowTime = NowTime + IIf(Weekday(Now) = 6, 10, 25) + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    '第四节课
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    '第五节课（限时训练）
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    
    '下午
    '午休
    NowTime = (14 * 60) + 0
    If DTime >= NowTime And DTime < NowTime + 20 Then Ret = True
    '第一节课
    NowTime = (14 * 60) + 20
    If DTime >= NowTime - 3 And DTime < NowTime + 40 Then Ret = True
    '第二节课
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    '第三节课
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    '第四节课（限时训练）
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 40 Then Ret = True
    
    '晚自习
    '第一节课
    NowTime = (18 * 60) + 20
    If DTime >= NowTime - 3 And DTime < NowTime + 40 Then Ret = True
    '第二节课
    NowTime = NowTime + 10 + 40
    If DTime >= NowTime - 2 And DTime < NowTime + 50 Then Ret = True
    '第三节课
    NowTime = NowTime + 10 + 50
    If DTime >= NowTime - 2 And DTime < NowTime + 50 Then Ret = True
    '第四节课（公共自习）
    NowTime = NowTime + 10 + 50
    If DTime >= NowTime - 2 And DTime < NowTime + 50 Then Ret = True
    '晚自习下课
    NowTime = NowTime + 50
    If DTime >= NowTime And DTime < NowTime + 20 Then Ret = True
    
    IsInClass = Ret Or (App.LogMode = 0)
End Function
Public Sub Speak(ByVal Content As String)
    If CloseSnd Then Exit Sub
    If Not Voice Is Nothing Then Voice.Speak Content, 1          '参数1表示异步播放
End Sub
Public Sub Start()
    '程序启动
    
    '从文件取得忽视记录
    Open App.path & "\ignored.stulist" For Binary As #1
    Get #1, , Ignored
    Close #1
    Open App.path & "\count.stulist" For Binary As #1
    Get #1, , RCount
    Close #1
    
    '创建语音对象
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
    '摇号准备
    ReDim Sticks(0)
    '加载所有可以被抽的学生
    For I = 1 To 62
        If I <> 39 Then '张亦佳
            '没有被忽略
            If ((Not Ignored(I)) Or JustRandom) And AntiDouble(I) = False Then
                If Filter = 1 And Student(3, I - 1) = "男" Then GoTo SkipThis
                If Filter = 2 And Student(3, I - 1) = "女" Then GoTo SkipThis
                ReDim Preserve Sticks(UBound(Sticks) + 1)
                Sticks(UBound(Sticks)) = I
SkipThis:
            End If
        End If
    Next
    '摇号中
    Randomize
Miss:
    '+1是因为Sticks(0)=0，顺便防止抽到62号的几率过小
    Dim Index As Integer
    Index = Int(Rnd * UBound(Sticks) + 1)
    '防止下标越界
    If Index > UBound(Sticks) Then Index = UBound(Sticks)
    RIndex = Sticks(Index) - 1
    
    If Rnd < Val(Student(2, RIndex)) Then
        '闪避成功
        IgnoredSomebody RIndex + 1
        GoTo Miss
    End If
    
    IgnoredSomebody RIndex + 1
    AntiDouble(RIndex + 1) = True
    
    GetRandom = RIndex
End Function
Public Sub DoneRandom()
    '摇号完毕
    MusicList.Play "Papa.mp3"
    '读出幸运小朋友
    Dim Ret As String, Bing As Boolean
    For I = 0 To UBound(Person)
        Ret = Ret & Student(0, Person(I)) & " "
        If RCount(Person(I) + 1) Mod 5 = 0 Then Bing = True
    Next
    If Bing Then MusicList.Play "LevelUp.mp3"
    Speak "恭喜" & Ret
End Sub
Public Sub CheckIgnored(Filter As Integer, Needed As Long)
    Dim Check As Boolean, Count As Long
    Count = Needed
    For I = 1 To 62
        If I <> 39 Then '张亦佳
            If Not Ignored(I) Then
                Check = True
                If Filter = 1 Then Check = Check And (Student(3, I - 1) = "女")
                If Filter = 2 Then Check = Check And (Student(3, I - 1) = "男")
                If Check Then Count = Count - 1
            End If
        End If
    Next
    If Count > 0 Then
        '不够抽，需要擦除
        For I = 1 To 62
            If I <> 39 Then '张亦佳
                Check = True
                If Filter = 1 Then Check = Check And (Student(3, I - 1) = "女")
                If Filter = 2 Then Check = Check And (Student(3, I - 1) = "男")
                If Check Then Ignored(I) = False
            End If
        Next
        Open App.path & "\ignored.stulist" For Binary As #1
        Put #1, , Ignored
        Close #1
    End If
End Sub
Public Sub IgnoredSomebody(Index As Integer)
    '忽略某人
    Ignored(Index) = True
    RCount(Index) = RCount(Index) + 1
    '存入文件
    Open App.path & "\ignored.stulist" For Binary As #1
    Put #1, , Ignored
    Close #1
    Open App.path & "\count.stulist" For Binary As #1
    Put #1, , RCount
    Close #1
End Sub
