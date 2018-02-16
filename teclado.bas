Attribute VB_Name = "teclado"
Option Explicit

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
 
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_SETTEXT = &HC

Private Const WM_KEYDOWN = &H100
Private Const VK_RETURN = &HD
' Function Keys
'Private Const VK_F1 = &H70
'Private Const VK_F2 = &H71
'Private Const VK_F3 = &H72
'Private Const VK_F4 = &H73
'Private Const VK_F5 = &H74
'Private Const VK_F6 = &H75
'Private Const VK_F7 = &H76
'Private Const VK_F8 = &H77
'Private Const VK_F9 = &H78
'Private Const VK_F10 = &H79
'Private Const VK_F11 = &H7A
'Private Const VK_F12 = &H7B

Public Enum TeclaFuncao
    VK_F1 = &H70
    VK_F2 = &H71
    VK_F3 = &H72
    VK_F4 = &H73
    VK_F5 = &H74
    VK_F6 = &H75
    VK_F7 = &H76
    VK_F8 = &H77
    VK_F9 = &H78
    VK_F10 = &H79
    VK_F11 = &H7A
    VK_F12 = &H7B
End Enum

Public Function WindowClass(ByVal hwnd As Long) As String
    Const MAX_LEN As Byte = 255
    Dim strBuff As String, intLen As Integer
    strBuff = String(MAX_LEN, vbNullChar)
    intLen = GetClassName(hwnd, strBuff, MAX_LEN)
    WindowClass = Left(strBuff, intLen)
End Function

Public Function WindowTextGet(ByVal hwnd As Long) As String
    Dim strBuff As String, lngLen As Long
    lngLen = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
    If lngLen > 0 Then
        lngLen = lngLen + 1
        strBuff = String(lngLen, vbNullChar)
        lngLen = SendMessage(hwnd, WM_GETTEXT, lngLen, ByVal strBuff)
        WindowTextGet = Left(strBuff, lngLen)
    End If
End Function
Public Function WindowTextSet(ByVal hwnd As Long, ByVal strText As String) As Boolean
    WindowTextSet = (SendMessage(hwnd, WM_SETTEXT, Len(strText), ByVal strText) <> 0)
End Function

Public Sub clica_externo(Janela As String, Botao As String)
    Dim lngHandle  As Long, lngHandlePai  As Long
    Debug.Print "Dialogo:" + Janela + "-" + Botao
    
    lngHandlePai = FindWindow(vbNullString, Janela)
    If lngHandlePai > 0 Then
        lngHandle = achaFilho(lngHandlePai, Botao)
        If lngHandle > 0 Then
            Call ClickButton3(lngHandle)
        End If
    End If
End Sub

Public Function achaFilho(ByVal lngHandlePai As Long, nomeFilho As String) As Long
    Dim lngHandleNeto As Long, lngHandleFilho As Long, texto As String, resultFilho As Long
    lngHandleFilho = GetWindow(lngHandlePai, GW_CHILD)
        
    Do Until lngHandleFilho = 0
        'Testa o iten filho
        texto = WindowTextGet(lngHandleFilho)
        Debug.Print "classe:" & WindowClass(lngHandleFilho) & " Valor:" & texto
        If texto = nomeFilho Then
            achaFilho = lngHandleFilho
            Exit Function
        Else
            resultFilho = achaFilho(lngHandleFilho, nomeFilho)
            If resultFilho <> 0 Then
               achaFilho = resultFilho
               Exit Function
            End If
        End If
            
        lngHandleFilho = GetWindow(lngHandleFilho, GW_HWNDNEXT)
   Loop
   achaFilho = 0
End Function

Public Sub ClickButton3(ByVal hWndChild As Long)
    Dim lResult As Long
    lResult = PostMessage(hWndChild, WM_LBUTTONDOWN, 1, 0)
    Debug.Print lResult
    lResult = PostMessage(hWndChild, WM_LBUTTONUP, 1, 0)
    Debug.Print lResult
End Sub


Public Sub EnviarTecla(ByVal lngHandlePai As Long, Tecla As TeclaFuncao)
    If lngHandlePai > 0 Then
        Call PostMessage(lngHandlePai, WM_KEYDOWN, Tecla, 0)
    End If
End Sub

Public Function AcharJanela(Janela As String) As Long

    Dim lngHandlePai  As Long, classe As String
        
    lngHandlePai = FindWindow(vbNullString, Janela)
    classe = WindowClass(lngHandlePai)
    Debug.Print "Dialogo:" + Janela + "(" & classe & ")-Enter"
    
    If classe = "MozillaDialogClass" Then
        lngHandlePai = GetWindow(lngHandlePai, GW_CHILD)
    End If
    AcharJanela = lngHandlePai
End Function

