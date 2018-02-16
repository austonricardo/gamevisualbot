Attribute VB_Name = "mouse"
Private Declare Sub mouse_event Lib "USER32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetMessageExtraInfo Lib "USER32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
Const MOUSEEVENTF_MOVE = &H1
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_RIGHTDOWN = &H8
Const MOUSEEVENTF_RIGHTUP = &H10
Const MOUSEEVENTF_MIDDLEDOWN = &H20
Const MOUSEEVENTF_MIDDLEUP = &H40
Const MOUSEEVENTF_ABSOLUTE = &H8000

Public Sub MouseLeftClick()
    'left click at current position
    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, GetMessageExtraInfo
End Sub

Private Sub MouseRightClick()
    'right click at current position
    mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0&, 0&, 0&, GetMessageExtraInfo
End Sub

Public Sub MouseMoveAbsolute(ByVal x As Long, ByVal Y As Long)
    'move the mouse to absolute screen position
    mouse_event MOUSEEVENTF_MOVE Or MOUSEEVENTF_ABSOLUTE, x, Y, 0&, GetMessageExtraInfo
End Sub

Private Sub MouseMoveRelative(ByVal x As Long, ByVal Y As Long)
    'move mouse relative to current position
    mouse_event MOUSEEVENTF_MOVE, x, Y, 0&, GetMessageExtraInfo
End Sub


Public Sub MouseLeftDown()
    'left click at current position
    mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, GetMessageExtraInfo
End Sub

Public Sub MouseLeftUp()
    'left click at current position
    mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, GetMessageExtraInfo
End Sub

