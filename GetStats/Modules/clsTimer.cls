VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public dtTime0 As Date
Public dtTime1 As Date
Public strSubOrFunc As String

Private Sub Class_Initialize(ByVal argSubOrFunc As String)
    dtTime0 = Timer
    strSubOrFunc = argSubOrFunc
End Sub

Private Sub Class_Terminate()
    dtTime1 = Timer
    Debug.Print "terminated " & suborfunc & " - " & (dtTime1 - dtTime0)
End Sub
'public property let strSubOrFunc as String


