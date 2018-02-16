Attribute VB_Name = "tela"
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" ( _
    ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DeleteDC Lib "GDI32" ( _
    ByVal hDC As Long) As Long
Private Declare Function CreateDCAsNull Lib "GDI32" Alias "CreateDCA" ( _
    ByVal lpDriverName As String, lpDeviceName As Any, _
   lpOutput As Any, lpInitData As Any) As Long
   
Public Declare Function BitBlt Lib "GDI32" ( _
    ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
    
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Public Sub CopyDesktop(hwnd As Long, ByRef objTo As Object, Optional x = 0, Optional y = 0)
Dim tR As RECT
Dim hDC As Long
    
    ' Note: objTo must have hDC,Picture,Width and Height
    ' properties and should have AutoRedraw = True
    
    ' Get the size of the desktop window:
    GetWindowRect hwnd, tR
    tR.Top = tR.Top + (y / Screen.TwipsPerPixelY)
    tR.Left = tR.Left + (x / Screen.TwipsPerPixelX)
    tR.Bottom = tR.Top + (objTo.Height / Screen.TwipsPerPixelY)
    tR.Right = tR.Left + (objTo.Width / Screen.TwipsPerPixelX)
    
    
    ' Set the object to the relevant size:
    'objTo.Width = (tR.Right - tR.Left) * Screen.TwipsPerPixelX
    'objTo.Height = (tR.Bottom - tR.Top) * Screen.TwipsPerPixelY
    
    'Modo de teste
    If False Then
        Dim hwndTemp As Long
        'hwndTemp = GetActiveWindow()
        'Call SetForegroundWindow(hWnd)
        'Call SetActiveWindow(hwnd)
        hDC = GetDC(hwnd)
        
        'DoEvents
        ' Copy the contents of the desktop to the object:
        BitBlt objTo.hDC, 0, 0, _
            (tR.Right - tR.Left), (tR.Bottom - tR.Top), hDC, 0, 0, SRCCOPY
        'Call SetForegroundWindow(hwndTemp)
    Else
        ' Now get the desktop DC:
        hDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        
        ' Copy the contents of the desktop to the object:
        BitBlt objTo.hDC, 0, 0, _
            (tR.Right - tR.Left), (tR.Bottom - tR.Top), hDC, tR.Left, tR.Top, SRCCOPY
    End If
        
        ' Ensure we clear up DC GDI has given us:
    DeleteDC hDC
    
End Sub


