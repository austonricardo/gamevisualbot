VERSION 5.00
Begin VB.Form frmAux 
   Caption         =   "L2 - Captcha detect"
   ClientHeight    =   12030
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6270
   Icon            =   "frm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12030
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdl2net 
      Caption         =   "l2net stop"
      Height          =   375
      Left            =   4800
      TabIndex        =   24
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer tmrAlertaSonoro 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5400
      Top             =   2280
   End
   Begin VB.PictureBox picNum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   3
      Left            =   4440
      ScaleHeight     =   900
      ScaleWidth      =   1215
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox picNum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   2
      Left            =   3000
      ScaleHeight     =   900
      ScaleWidth      =   1215
      TabIndex        =   22
      Top             =   4440
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox picNum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   1
      Left            =   1560
      ScaleHeight     =   900
      ScaleWidth      =   1215
      TabIndex        =   21
      Top             =   4440
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox picNum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   930
      Index           =   0
      Left            =   120
      ScaleHeight     =   900
      ScaleWidth      =   1215
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   120
      ScaleHeight     =   900
      ScaleWidth      =   5415
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   5445
   End
   Begin VB.CommandButton cmdlimparPos 
      Caption         =   "LimparCordenadas"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton cmdProcuraJanela 
      Caption         =   "Janela L2.Net"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtFiltroJanela 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Text            =   "WindowsForm"
      Top             =   1890
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer tmrLerTela 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   2160
   End
   Begin VB.CommandButton cmdCalular 
      Caption         =   "Captura incial"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CheckBox chkCP 
      BackColor       =   &H000093F9&
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "870"
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox picTeste 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5370
      Left            =   120
      ScaleHeight     =   5340
      ScaleWidth      =   5415
      TabIndex        =   7
      Top             =   3480
      Width           =   5445
   End
   Begin VB.TextBox txtFiltroJanela 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Text            =   "L2U"
      Top             =   1410
      Width           =   1335
   End
   Begin VB.Timer tmrProcuraJanela 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   2160
   End
   Begin VB.CommandButton cmdProcuraJanela 
      Caption         =   "Janela L2"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   1320
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Testar"
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdParar 
      Caption         =   "Parar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton cmdIniciar 
      Caption         =   "Iniciar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2955
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblocr 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Filtro Janela"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblHwndSelNet 
      Caption         =   "#####"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   60
      Visible         =   0   'False
      X1              =   5640
      X2              =   2640
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblHwndSel 
      Caption         =   "#####"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Cor da captcha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblCP 
      AutoSize        =   -1  'True
      Caption         =   "CP: Não setado"
      Height          =   195
      Left            =   600
      TabIndex        =   8
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "Filtro Janela"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblMensagem 
      Caption         =   "Clique em janela e mova o mouse sobre a janela a ser controlada"
      Height          =   555
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   5295
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "USER32" _
   (lpPoint As POINTAPI) As Long

Private Declare Function WindowFromPoint Lib "USER32" (ByVal _
   xpoint As Long, ByVal ypoint As Long) As Long
   
Private Declare Function GetClassName Lib "USER32" Alias _
   "GetClassNameA" (ByVal hwnd As Long, ByVal lpClass _
    As String, ByVal nMaxCount As Long) As Long

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private gStop As Boolean
Private prevWindow As Long, curWindow As Long
Private x As Long, Y As Long
Private className As String
Private retValue As Long
Private mousePT As POINTAPI
Private windowHandled As Long
Private chkAtual As CheckBox
Private lbIniciouProcuraJanela As Boolean
Private lbLerAtributos As Boolean

Private windowHandledNet As Long

Private selecaoCor As Boolean, selecaoPonto1 As Boolean, selecaoPonto2 As Boolean
Private liAguardando As Long

Private liJanela As Integer


Private HPmin As Single
Private HPMax As Single
Private HPAtu As Single
Private HPY As Single
Private HPPC As Byte

Private CPmin As Single
Private CPMax As Single
Private CPAtu As Single
Private CPY As Single
Private CPPC As Byte

Private MPmin As Single
Private MPMax As Single
Private MPAtu As Single
Private MPY As Single
Private MPPC As Byte

Private HPartymin(2) As Single
Private HPartyMax(2) As Single
Private HPartyAtu(2) As Single
Private HPartyY(2) As Single
Private HPartyPC(2) As Byte

Dim liMarca(5) As Integer
Dim liMarcaTotal(5) As Integer
Dim ponto1 As POINTAPI, ponto2 As POINTAPI

Private Sub chkCP_Click()
    Set chkAtual = chkCP
    selecaoCor = True
    picTeste.MousePointer = vbCrosshair
End Sub


Private Sub cmdCalular_Click()
        Dim COR As Long, ix As Single
        CopyDesktop windowHandled, picTeste
        Dim yleitura  As Single
        yleitura = picTeste.Top + (picTeste.Height / 2)
        'call picTeste.Line vbred, (0;yleitura) (picTeste.Width;yleitura)
        Line1.Y1 = yleitura
        Line1.Y2 = yleitura
        
        cmdlimparPos_Click
End Sub

Private Sub cmdl2net_Click()
    menu.acionaMenu windowHandledNet, 3, 1
End Sub

Private Sub cmdlimparPos_Click()
    selecaoPonto1 = True
    lblMensagem.Caption = "Selecione dois pontos na imagem abaixo que delimitam a área do captcha."
End Sub

Private Sub cmdProcuraJanela_Click(Index As Integer)
    liJanela = Index
    lbIniciouProcuraJanela = Not lbIniciouProcuraJanela
    txtFiltroJanela(Index).Enabled = Not lbIniciouProcuraJanela
    tmrProcuraJanela.Enabled = lbIniciouProcuraJanela
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        picTeste.Height = Me.Height - (picTeste.Top + 600)
    End If
End Sub

Private Sub picTeste_Click()
    If selecaoCor Then
                selecaoCor = False
                picTeste.MousePointer = vbDefault
                chkAtual.value = 0
    ElseIf selecaoPonto1 Then
            selecaoPonto1 = False
            selecaoPonto2 = True
    ElseIf selecaoPonto2 Then
            selecaoPonto2 = False
            
            'desenha quadrado
            picTeste.Line (ponto1.x, ponto1.Y)-(ponto2.x, ponto1.Y), vbRed
            picTeste.Line (ponto1.x, ponto2.Y)-(ponto2.x, ponto2.Y), vbRed
            
            picTeste.Line (ponto1.x, ponto1.Y)-(ponto1.x, ponto2.Y), vbRed
            picTeste.Line (ponto2.x, ponto1.Y)-(ponto2.x, ponto2.Y), vbRed
            
    End If
End Sub

Private Sub picTeste_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If selecaoCor Then
        chkAtual.BackColor = picTeste.Point(x, Y)
        chkAtual.Tag = Y
    ElseIf selecaoPonto1 Then
        ponto1.x = x
        ponto1.Y = Y
    ElseIf selecaoPonto2 Then
        ponto2.x = x
        ponto2.Y = Y
    End If
End Sub

Private Sub tmrAlertaSonoro_Timer()
    media.tocarSom
End Sub

Private Sub tmrLerTela_Timer()
    Dim COR As Long
    'CopyDesktop windowHandled, picTeste
        
  Picture2.Width = (ponto2.x - ponto1.x)
  Picture2.Height = ponto2.Y - ponto1.Y
  
  tela.CopyDesktop windowHandled, Picture2, ponto1.x, ponto1.Y
        
    Dim yleitura  As Single
    yleitura = (picTeste.Height / 2)
    
    Dim ix As Single, iy As Single
    For iy = 0 To Picture2.Height Step 100
        For ix = 0 To Picture2.Width Step 100
            
                COR = Picture2.Point(ix, iy)
    'Debug.Print COR & ":" & chkCP.BackColor
    'Picture2.PSet (ix, iy), vbGreen
    
                If fgCoresProximas(COR, chkCP.BackColor) Then
                    alertaCaptcha
                    Exit For
                End If
        Next
    Next

End Sub

Private Function alertaCaptcha()
    'menu.acionaMenu windowHandledNet, 3, 1
  Picture2.Width = (ponto2.x - ponto1.x)
  Picture2.Height = ponto2.Y - ponto1.Y
  
  tela.CopyDesktop windowHandled, Picture2, ponto1.x, ponto1.Y

'  Dim lbTanaletra As Boolean, lbEstavaNaletra As Boolean
'  Dim COR As Long
'  Dim ix As Single, iy As Single
'  Dim cortes(4) As Integer, corteAtual As Integer
'  'separa letras
''    For ix = 0 To Picture2.Width Step 5
'        lbEstavaNaletra = lbTanaletra
'        lbTanaletra = False
'        For iy = 0 To Picture2.Height Step 5
'                COR = Picture2.Point(ix, iy)
'                'Picture2.p .PSet(ix, iy), vbGreen
'                If fgCoresProximas(COR, chkCP.BackColor) Then
'                    lbTanaletra = True
'                    Exit For
'                End If
'        Next
'        If lbEstavaNaletra And Not lbTanaletra Then
'            cortes(corteAtual) = ix
'            corteAtual = corteAtual + 1
'            If corteAtual >= 4 Then Exit For
'        'ElseIf lbTanaletra = True Then
'         '   Exit For
'        End If
'    Next
'
'    For ix = 0 To 4
'        Picture2.Line (cortes(ix), 0)-(cortes(ix), Picture2.Height), vbRed
'    Next
    
  'SavePicture Picture2.Image, App.path & "\001.bmp"
  
   'MsgBox OCRImage(ConvertToTif(App.path & "\001.bmp"))
    
    tmrAlertaSonoro.Enabled = True
    tmrLerTela.Enabled = False
    MsgBox "captcha detectada"
    tmrLerTela.Enabled = True
    tmrAlertaSonoro.Enabled = False
End Function

Private Function ConvertToTif(ImageName As String) As String
'    Dim imgFile As New ImageFile
'    Dim IP As New ImageProcess
'    Dim strFileName As String
    
    Dim imgFile As Object
    Dim IP As Object
    Dim strFileName As String
    Const wiaFormatTIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
    
    Set IP = CreateObject("WIA.ImageProcess")
    Set imgFile = CreateObject("WIA.ImageFile")
    imgFile.LoadFile ImageName
    
    IP.Filters.Add IP.FilterInfos("Convert").FilterID
    IP.Filters(1).Properties("FormatID").value = wiaFormatTIFF
    IP.Filters(1).Properties("Quality").value = 5
    
    Set imgFile = IP.Apply(imgFile)
     
    strFileName = Replace(ImageName, imgFile.FileExtension, ".tif")
    
    If Dir(strFileName) <> "" Then
        Kill strFileName
    End If
    
    imgFile.SaveFile strFileName
    Set imgFile = Nothing
    
    ConvertToTif = strFileName
End Function

Private Function OCRImage(strFileName As String) As String
    'Dim objDoc As MODI.Document
    'Dim objImg As MODI.Image
    Dim objDoc As Object
    Dim objImg As Object
    Const miLANG_ENGLISH = 9
    
    
    Set objDoc = CreateObject("MODI.Document")
    objDoc.Create (strFileName)
    Set objImg = objDoc.Images(0)
    objImg.OCR miLANG_ENGLISH
    
    
    OCRImage = objImg.Layout.Text
    
End Function


Private Sub tmrProcuraJanela_Timer()

gStop = False
prevWindow = 0
Do
    If gStop = True Then Exit Do
    Call GetCursorPos(mousePT)
    x = mousePT.x
    Y = mousePT.Y
    curWindow = WindowFromPoint(x, Y)
    If curWindow <> prevWindow Then
        className = String$(256, " ")
        prevWindow = curWindow
        retValue = GetClassName(curWindow, className, 255)
        className = Left$(className, InStr(className, _
            vbNullChar) - 1)
            If className = "SysListView32" Then
             lblMensagem.Caption = "the mouse is over the desktop. "
            Else
               lblMensagem.Caption = "the mouse is over " & className
               If InStr(className, txtFiltroJanela(liJanela).Text) > 0 Then
                    If liJanela = 0 Then
                        windowHandled = curWindow
                        lblHwndSel.Caption = CStr(windowHandled)
                    Else
                        windowHandledNet = curWindow
                        lblHwndSelNet.Caption = CStr(windowHandledNet)
                    End If
                    
                    tmrProcuraJanela.Enabled = False
                    cmdIniciar.Enabled = True
                                        cmdCalular.Enabled = True
                    lblMensagem.Caption = "Janela Escolhida: " & CStr(windowHandled) & "-" & className
                    lblHwndSel.Caption = CStr(windowHandled)
                    Exit Do
                Else
                    If lbIniciouProcuraJanela = False Then Exit Do
               End If
            End If
    End If
          DoEvents
 Loop
 
End Sub




Private Sub cmdIniciar_Click()
    'So inicia se foi inicializado
    tmrLerTela.Enabled = True
    
    Me.Caption = "Executando"
    cmdParar.Enabled = True
    cmdProcuraJanela(0).Enabled = False
    cmdProcuraJanela(1).Enabled = False
    cmdIniciar.Enabled = False
    cmdlimparPos.Enabled = False
    cmdCalular.Enabled = False
    picTeste.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdParar_Click()
    Dim li As Integer
    Me.Caption = "Parado"
    cmdParar.Enabled = False
    cmdIniciar.Enabled = True
    cmdProcuraJanela(0).Enabled = True
    cmdProcuraJanela(1).Enabled = True
    tmrLerTela.Enabled = False
    cmdlimparPos.Enabled = True
    cmdCalular.Enabled = True
    picTeste.Visible = True
    Picture2.Visible = False
    
End Sub

Private Sub SendKeysHWD(key As String)
    If windowHandled = 0 Then MsgBox "A Janela escolhida não esta mais disponivel"
        Dim Tecla As TeclaFuncao
        
        Select Case UCase(key)
            Case "F1": Tecla = VK_F1
            Case "F2": Tecla = VK_F2
            Case "F3": Tecla = VK_F3
            Case "F4": Tecla = VK_F4
            Case "F5": Tecla = VK_F5
            Case "F6": Tecla = VK_F6
            Case "F7": Tecla = VK_F7
            Case "F8": Tecla = VK_F8
            Case "F9": Tecla = VK_F9
            Case "F10": Tecla = VK_F10
            Case "F11": Tecla = VK_F11
            Case "F12": Tecla = VK_F12
            Case Else:
                'Exit Sub
                                If IsNumeric(key) Then
                                        liAguardando = CLng(key)
                                        Do
                                           Sleep 1000
                                           liAguardando = liAguardando - 1
                                           lblMensagem.Caption = "Aguardando ... " & CStr(liAguardando) & "s"
                                        Loop While (liAguardando > 0)
                                        lblMensagem.Caption = "Pronto..."
                                End If
                                Exit Sub
        End Select
        
        teclado.EnviarTecla windowHandled, Tecla
        
End Sub



