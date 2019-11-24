VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Piano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

Dim Hmidi As Long, RC As Long, Midiout As Long, mChannel As Integer, mVolume As Integer
Dim mBaseNote As Long

Public Property Get BaseNote() As Long
    BaseNote = mBaseNote
End Property
Public Property Let BaseNote(ByVal nValue As Long)
    mBaseNote = nValue
End Property
Public Property Get channel() As Integer
    channel = mChannel
End Property
Public Property Let channel(ByVal nValue As Integer)
    mChannel = nValue
End Property
Public Property Get Volume() As Integer
    Volume = mVolume
End Property
Public Property Let Volume(ByVal nValue As Integer)
    mVolume = nValue
End Property
Public Sub Play(ByVal index As Integer)
   If CloseSnd Then Exit Sub
   Midiout = &H90 + ((mBaseNote + index) * &H100) + (mVolume * &H10000) + mChannel
   midiOutShortMsg Hmidi, Midiout
End Sub
Public Sub PlayDone(ByVal index As Integer)
   Midiout = &H80 + ((mBaseNote + index) * &H100) + mChannel
   midiOutShortMsg Hmidi, Midiout
End Sub
Sub Init()
    RC = midiOutClose(Hmidi)
    RC = midiOutOpen(Hmidi, -1, 0, 0, 0)
    mChannel = 0: mVolume = 127: mBaseNote = 60
    If (RC <> 0) Then Debug.Print "Failed to open midi ."
End Sub
Sub Dispose()
   RC = midiOutClose(Hmidi)
End Sub
Private Sub Class_Terminate()
   Call Dispose
End Sub

