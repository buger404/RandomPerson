Attribute VB_Name = "Randoms"
'=====================================================================================
'   摇号核心模块
'   Maker：LSM/CZY
'=====================================================================================
'   通用部分
    Public Student() As String
    '所有同学：Student(x,y)
    '第一列(x=1)为名字，第二列(x=2)为座号，第三列(x=3)为闪避概率
    
    Public Ignored(62) As Boolean   '忽视标记
    Public Sticks() As Integer      '可以抽取的同学，对应Student的y
    Public RIndex As Integer        '当前抽到的号数
    Public Voice As Object          '语音对象
'=====================================================================================
Public Sub Speak(ByVal Content As String)
    Voice.Speak Content, 1          '参数1表示异步播放
End Sub
Public Sub Start()
    '程序启动
    
    '从文件取得忽视记录
    Open App.path & "\ignored.stulist" For Binary As #1
    Get #1, , Ignored
    Close #1
    
    '创建语音对象
    Set Voice = CreateObject("SAPI.SpVoice")
    Voice.Volume = 100
End Sub
Public Sub StartRandom()
    '摇号准备
    ReDim Sticks(0)
    '加载所有可以被抽的学生
    For I = 1 To 62
        If I <> 39 Then '张亦佳
            '没有被忽略
            If Not Ignored(I) Then
                ReDim Preserve Sticks(UBound(Sticks) + 1)
                Sticks(UBound(Sticks)) = I
            End If
        End If
    Next
End Sub
Public Sub GetRandom()
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
        IgnoredSomebody RIndex
        GoTo Miss
    End If
End Sub
Public Sub DoneRandom()
    '摇号完毕
    IgnoredSomebody RIndex
    '读出幸运小朋友
    Speak "恭喜" & Student(0, RIndex)
End Sub
Public Sub IgnoredSomebody(Index As Integer)
    '忽略某人
    Ignored(Index) = True
    '判断是否已经全部忽略
    Dim AllIgnore As Boolean
    AllIgnore = True
    For I = 1 To 62
        If I <> 39 Then '张亦佳
            If Not Ignored(I) Then AllIgnore = False: Exit For
        End If
    Next
    '如果已经全部忽略
    Erase Ignored()
    '存入文件
    Open App.path & "\ignored.stulist" For Binary As #1
    Put #1, , Ignored
    Close #1
End Sub
