VERSION 5.00
Begin VB.Form frmTeste 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   960
      Top             =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" _
   (lpPoint As POINTAPI) As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal _
   xpoint As Long, ByVal ypoint As Long) As Long
   
Private Declare Function GetClassName Lib "user32" Alias _
   "GetClassNameA" (ByVal hWnd As Long, ByVal lpClass _
    As String, ByVal nMaxCount As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private gStop As Boolean
Private prevWindow As Long, curWindow As Long
Private X As Long, Y As Long
Private className As String
Private retValue As Long
Private mousePT As POINTAPI

Private Sub Timer1_Timer()

gStop = False
prevWindow = 0
Do
    If gStop = True Then Exit Do
    Call GetCursorPos(mousePT)
    X = mousePT.X
    Y = mousePT.Y
    curWindow = WindowFromPoint(X, Y)
    If curWindow <> prevWindow Then
        className = String$(256, " ")
        prevWindow = curWindow
        retValue = GetClassName(curWindow, className, 255)
        className = Left$(className, InStr(className, _
            vbNullChar) - 1)
            If className = "SysListView32" Then
             Label1.Caption = "the mouse is over the desktop. "
            Else
               Label1.Caption = "the mouse is over " & className
            End If
    End If
          DoEvents
 Loop
End Sub

