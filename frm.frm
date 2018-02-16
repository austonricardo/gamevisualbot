VERSION 5.00
Begin VB.Form frmAux 
   Caption         =   "L2 - Auxiliar"
   ClientHeight    =   10440
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3285
   Icon            =   "frm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTeclas 
      Height          =   285
      Index           =   5
      Left            =   0
      TabIndex        =   48
      Text            =   "F12"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtTempoAcao 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   47
      Text            =   "3"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txtCond 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   2280
      TabIndex        =   46
      Text            =   "HP<70"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtTeclas 
      Height          =   285
      Index           =   4
      Left            =   0
      TabIndex        =   45
      Text            =   "F11"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtTempoAcao 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   44
      Text            =   "2"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtCond 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   2280
      TabIndex        =   43
      Text            =   "CP<70"
      Top             =   1680
      Width           =   975
   End
   Begin VB.CheckBox chkH 
      BackColor       =   &H00111C79&
      Height          =   255
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6240
      Width           =   255
   End
   Begin VB.CheckBox chkH 
      BackColor       =   &H00111C79&
      Height          =   255
      Index           =   1
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6600
      Width           =   255
   End
   Begin VB.CheckBox chkH 
      BackColor       =   &H00111C79&
      Height          =   255
      Index           =   2
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6960
      Width           =   255
   End
   Begin VB.TextBox txtCond 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   2280
      TabIndex        =   30
      Text            =   "MP<70"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtCond 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   2280
      TabIndex        =   29
      Text            =   "HP<10"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtCond 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   2280
      TabIndex        =   28
      Text            =   "Sempre"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtCond 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   27
      Text            =   "Sempre"
      Top             =   240
      Width           =   975
   End
   Begin VB.Timer tmrObterAtributos 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   3840
   End
   Begin VB.CommandButton cmdCalular 
      Caption         =   "Full Status"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CheckBox chkMP 
      BackColor       =   &H00924007&
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "1290"
      Top             =   5640
      Width           =   255
   End
   Begin VB.CheckBox chkHP 
      BackColor       =   &H00111C79&
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   "1080"
      Top             =   5280
      Width           =   255
   End
   Begin VB.CheckBox chkCP 
      BackColor       =   &H00206485&
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "870"
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox picTeste 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   50
      ScaleHeight     =   3180
      ScaleWidth      =   3135
      TabIndex        =   19
      Top             =   7320
      Width           =   3165
   End
   Begin VB.TextBox txtFiltroJanela 
      Height          =   285
      Left            =   960
      TabIndex        =   18
      Text            =   "L2U"
      Top             =   3330
      Width           =   1335
   End
   Begin VB.Timer tmrProcuraJanela 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   3360
   End
   Begin VB.CommandButton cmdProcuraJanela 
      Caption         =   "Janela"
      Height          =   375
      Left            =   1340
      TabIndex        =   15
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdAbrirArquivo 
      Caption         =   "Abrir"
      Height          =   375
      Left            =   2000
      TabIndex        =   32
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdSalvarArquivo 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   2650
      TabIndex        =   33
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Testar"
      Height          =   375
      Left            =   8160
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtPosicoes 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Text            =   "35,4;35,60"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtTempoMov 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox txtTempoAcao 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   11
      Text            =   "2"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtTeclas 
      Height          =   285
      Index           =   3
      Left            =   0
      TabIndex        =   10
      Text            =   "F10"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Timer tmAcao 
      Index           =   3
      Left            =   2520
      Top             =   1200
   End
   Begin VB.TextBox txtTempoAcao 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   9
      Text            =   "1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtTeclas 
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Text            =   "F8;F9"
      Top             =   960
      Width           =   1695
   End
   Begin VB.Timer tmAcao 
      Index           =   2
      Left            =   2520
      Top             =   840
   End
   Begin VB.TextBox txtTempoAcao 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Text            =   "60"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtTeclas 
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Text            =   "F6;F7"
      Top             =   600
      Width           =   1695
   End
   Begin VB.Timer tmAcao 
      Index           =   1
      Left            =   2520
      Top             =   480
   End
   Begin VB.TextBox txtTempoAcao 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Text            =   "3"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtTeclas 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Text            =   "F2;F3;F1;F1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Timer tmMov 
      Left            =   2520
      Top             =   2520
   End
   Begin VB.CommandButton cmdParar 
      Caption         =   "Parar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2880
      Width           =   615
   End
   Begin VB.Timer tmAcao 
      Index           =   0
      Left            =   2520
      Top             =   120
   End
   Begin VB.CommandButton cmdIniciar 
      Caption         =   "Iniciar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   680
      TabIndex        =   0
      Top             =   2880
      Width           =   615
   End
   Begin VB.Timer tmAcao 
      Index           =   4
      Left            =   2520
      Top             =   1560
   End
   Begin VB.Timer tmAcao 
      Index           =   5
      Left            =   2520
      Top             =   1920
   End
   Begin VB.Label lblHwndSel 
      Caption         =   "#####"
      Height          =   255
      Left            =   2280
      TabIndex        =   42
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Party Status"
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
      Index           =   1
      Left            =   240
      TabIndex        =   41
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Char Status"
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
      TabIndex        =   40
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblH 
      AutoSize        =   -1  'True
      Caption         =   "H1: Não setado"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   39
      Top             =   6240
      Width           =   1125
   End
   Begin VB.Label lblH 
      AutoSize        =   -1  'True
      Caption         =   "H2:Não setado"
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   38
      Top             =   6600
      Width           =   1080
   End
   Begin VB.Label lblH 
      AutoSize        =   -1  'True
      Caption         =   "H3:Não setado"
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   37
      Top             =   6960
      Width           =   1080
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Condição"
      Height          =   255
      Left            =   2520
      TabIndex        =   31
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblMP 
      AutoSize        =   -1  'True
      Caption         =   "MP:Não setado"
      Height          =   195
      Left            =   600
      TabIndex        =   22
      Top             =   5640
      Width           =   1110
   End
   Begin VB.Label lblHP 
      AutoSize        =   -1  'True
      Caption         =   "HP:Não setado"
      Height          =   195
      Left            =   600
      TabIndex        =   21
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lblCP 
      AutoSize        =   -1  'True
      Caption         =   "CP: Não setado"
      Height          =   195
      Left            =   600
      TabIndex        =   20
      Top             =   4920
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "Filtro Janela"
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblMensagem 
      Caption         =   "Clique em janela e mova o mouse sobre a janela a ser controlada"
      Height          =   555
      Left            =   0
      TabIndex        =   16
      Top             =   3720
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tempo (s)"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Teclas"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmAux"
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

Private selecaoCor As Boolean
Private liAguardando As Long

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

Private Sub chkCP_Click()
    Set chkAtual = chkCP
    selecaoCor = True
    picTeste.MousePointer = vbCrosshair
End Sub

Private Sub chkH_Click(Index As Integer)
    Set chkAtual = chkH(Index)
    selecaoCor = True
    picTeste.MousePointer = vbCrosshair
End Sub

Private Sub chkHP_Click()
    Set chkAtual = chkHP
    selecaoCor = True
    picTeste.MousePointer = vbCrosshair
End Sub

Private Sub chkMP_Click()
    Set chkAtual = chkMP
    selecaoCor = True
    picTeste.MousePointer = vbCrosshair
End Sub

Private Sub cmdCalular_Click()
        Dim COR As Long, ix As Single
        CopyDesktop windowHandled, picTeste
        
    CPY = CSng(chkCP.Tag)
    HPY = CSng(chkHP.Tag)
    MPY = CSng(chkMP.Tag)
    
    Dim li As Integer
    For li = 0 To chkH.UBound
        If chkH(li).Tag <> "" Then
            HPartyY(li) = CSng(chkH(li).Tag)
        End If
    Next
    
    For ix = 1 To picTeste.Width Step 10
    
            COR = picTeste.Point(ix, CPY)
            If fgCoresProximas(COR, chkCP.BackColor) Then
                    If ix < CPmin Or CPmin = 0 Then CPmin = ix
                    If ix > CPMax Then CPMax = ix
            End If
            
            COR = picTeste.Point(ix, HPY)
            If fgCoresProximas(COR, chkHP.BackColor) Then
                    If ix < HPmin Or HPmin = 0 Then HPmin = ix
                    If ix > HPMax Then HPMax = ix
            End If
            
            COR = picTeste.Point(ix, MPY)
            If fgCoresProximas(COR, chkMP.BackColor) Then
                    If ix < MPmin Or MPmin = 0 Then MPmin = ix
                    If ix > MPMax Then MPMax = ix
            End If
            
            For li = 0 To chkH.UBound
                If chkH(li).Tag <> "" Then
                    COR = picTeste.Point(ix, HPartyY(li))
                    If fgCoresProximas(COR, chkH(li).BackColor) Then
                            If ix < HPartymin(li) Or HPartymin(li) = 0 Then HPartymin(li) = ix
                            If ix > HPartyMax(li) Then HPartyMax(li) = ix
                    End If
                End If
            Next
    Next
        
        CPAtu = (CPMax - CPmin) / 2
        HPAtu = (HPMax - HPmin) / 2
        MPAtu = (MPMax - MPmin) / 2
        
        For li = 0 To chkH.UBound
            If chkH(li).Tag <> "" Then
                HPartyAtu(li) = (HPartyMax(li) - HPartymin(li)) / 2
            End If
        Next
        
        'lblMsg.caption
        mostrarAtributos
        lbLerAtributos = True
End Sub

Private Sub cmdProcuraJanela_Click()
    lbIniciouProcuraJanela = Not lbIniciouProcuraJanela
    txtFiltroJanela.Enabled = Not lbIniciouProcuraJanela
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
        End If
End Sub

Private Sub picTeste_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If selecaoCor Then
        chkAtual.BackColor = picTeste.Point(x, Y)
        chkAtual.Tag = Y
    End If
End Sub

Private Sub tmrObterAtributos_Timer()
    Dim COR As Long
    CopyDesktop windowHandled, picTeste
    
    CPAtu = CPmin
    HPAtu = HPmin
    MPAtu = MPmin
    Dim li As Integer
    For li = 0 To chkH.UBound
        HPartyAtu(li) = HPartymin(li)
    Next
    
    Dim ix As Single
    For ix = HPmin To HPMax
            COR = picTeste.Point(ix, CPY)
            If fgCoresProximas(COR, chkCP.BackColor) Then
                    If ix > CPAtu Then CPAtu = ix
            End If
            
            COR = picTeste.Point(ix, HPY)
            If fgCoresProximas(COR, chkHP.BackColor) Then
                    If ix > HPAtu Then HPAtu = ix
            End If
            
            COR = picTeste.Point(ix, MPY)
            If fgCoresProximas(COR, chkMP.BackColor) Then
                    If ix > MPAtu Then MPAtu = ix
            End If
            
            For li = 0 To chkH.UBound
                If chkH(li).Tag <> "" Then
                    COR = picTeste.Point(ix, HPartyY(li))
                    If fgCoresProximas(COR, chkH(li).BackColor) Then
                            If ix > HPartyAtu(li) Then HPartyAtu(li) = ix
                    End If
                End If
            Next
            
    Next
        
    mostrarAtributos
End Sub

Private Sub mostrarAtributos()
        On Error Resume Next
    CPPC = Round(100 * (CPAtu - CPmin) / (CPMax - CPmin))
    HPPC = Round(100 * (HPAtu - HPmin) / (HPMax - HPmin))
    MPPC = Round(100 * (MPAtu - MPmin) / (MPMax - MPmin))
    
    lblCP.Caption = "CP:" & CPPC & "%"
    lblHP.Caption = "HP:" & HPPC & "%"
    lblMP.Caption = "MP:" & MPPC & "%"
    Dim li As Integer
    For li = 0 To chkH.UBound
        If chkH(li).Tag <> "" Then
            HPartyPC(li) = Round(100 * (HPartyAtu(li) - HPartymin(li)) / (HPartyMax(li) - HPartymin(li)))
            lblH(li).Caption = "H" & (li + 1) & ":" & HPartyPC(li) & "%"
        End If
    Next
    
    DoEvents
End Sub

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
               If InStr(className, txtFiltroJanela.Text) > 0 Then
                    windowHandled = curWindow
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
    Dim li As Integer
    For li = 0 To tmAcao.UBound
        If txtTempoAcao(li).Text <> "" Then
            Dim lTimer As Integer
            lTimer = CLng(txtTempoAcao(li).Text)
                     
            liMarca(li) = lTimer
            liMarcaTotal(li) = 0
            tmAcao(li).Interval = 1000
            tmAcao(li).Enabled = True
        Else
            tmAcao(li).Enabled = False
        End If
    Next
    
        'So inicia se foi inicializado
        tmrObterAtributos.Enabled = lbLerAtributos
        
    If txtTempoMov.Text <> "" Then
        tmMov.Interval = CInt(txtTempoMov.Text) * 1000
        tmMov.Enabled = True
    End If
    
    Me.Caption = "Executando"
    cmdParar.Enabled = True
    cmdProcuraJanela.Enabled = False
    cmdIniciar.Enabled = False
End Sub

Private Sub cmdParar_Click()
    Dim li As Integer
    For li = 0 To tmAcao.UBound
        tmAcao(li).Enabled = False
    Next
    tmMov.Enabled = False
    Me.Caption = "Parado"
    cmdParar.Enabled = False
    cmdIniciar.Enabled = True
    cmdProcuraJanela.Enabled = True
    tmrObterAtributos.Enabled = False
End Sub

Private Sub cmdSalvarArquivo_Click()
        saveFile
End Sub
Private Sub cmdAbrirArquivo_Click()
        OpenFile
End Sub

Private Sub tmAcao_Timer(Index As Integer)
    If condicaoSatisfeita(txtCond(Index).Text) Then
        If liMarcaTotal(Index) < liMarca(Index) Then
                        liMarcaTotal(Index) = liMarcaTotal(Index) + 1
        Else
                        lblMensagem.Caption = "Executando Comando " & (Index + 1)
                        Dim keys
                        keys = Split(txtTeclas(Index).Text, ";")
                        Dim conta As Integer
                        For conta = 0 To UBound(keys)
                                SendKeysHWD CStr(keys(conta))
                                Call Sleep(500)
                        Next
                        liMarcaTotal(Index) = 0
        End If
    End If
End Sub

Private Function condicaoSatisfeita(expressao As String) As Boolean
    Dim lval As Byte
        If liAguardando <> 0 Then
                condicaoSatisfeita = False
    ElseIf expressao = "Sempre" Or Not lbLerAtributos Then
        condicaoSatisfeita = True
    Else
       lval = CByte(Mid(expressao, 4))
       Dim lvatributo As Byte
       Select Case Left(expressao, 2)
        Case "CP"
                lvatributo = CPPC
        Case "HP"
                lvatributo = HPPC
        Case "MP"
                lvatributo = MPPC
        Case "H1"
                lvatributo = HPartyPC(0)
        Case "H2"
                lvatributo = HPartyPC(1)
        Case "H3"
                lvatributo = HPartyPC(2)
       End Select
       condicaoSatisfeita = IIf(Mid(expressao, 3, 1) = "<", lvatributo < lval, lvatributo > lval)
    End If
End Function

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

Private Sub tmMov_Timer()
    Dim lmovY As Long, lmovX As Long
        
    Dim posicoes, nums
    posicoes = Split(txtPosicoes.Text, ";")
    Dim conta As Integer
    For conta = 0 To UBound(posicoes)
        nums = Split(posicoes(conta), ",")
        lmovX = CLng(nums(0)) * 1000
        lmovY = CLng(nums(1)) * 1000
        
        mouse.MouseMoveAbsolute lmovX, lmovY
        DoEvents
        Sleep 1000
        mouse.MouseLeftClick
        Sleep 500
        mouse.MouseLeftDown
        Sleep 500
        mouse.MouseLeftUp
    Next
End Sub

Private Sub OpenFile()
        Dim strFilter As String
        Dim strInputFileName As String, strLinha As String

        strFilter = ahtAddFilterItem(strFilter, "Bots  (*.bot)", "*.bot")
        strInputFileName = ahtCommonFileOpenSave( _
                                        Filter:=strFilter, OpenFile:=True, _
                                        DialogTitle:="Please select an input file...", _
                                        Flags:=ahtOFN_HIDEREADONLY, hwnd:=Me.hwnd)

        If strInputFileName <> "" Then
                'Carrega o arquivo
                Dim liFile As Integer, liContador As Integer
                liFile = FreeFile
                Open strInputFileName For Input As #liFile
                        
                        For liContador = 0 To txtTeclas.UBound
                                Line Input #liFile, strLinha
                                txtTeclas(liContador) = strLinha

                                Line Input #liFile, strLinha
                                txtTempoAcao(liContador) = strLinha

                                Line Input #liFile, strLinha
                                txtCond(liContador) = strLinha
                        Next
                        Line Input #liFile, strLinha
                        chkCP.BackColor = CLng(strLinha)
                        
                        Line Input #liFile, strLinha
                        chkCP.Tag = strLinha
                        
                        Line Input #liFile, strLinha
                        chkHP.BackColor = CLng(strLinha)
                        
                        Line Input #liFile, strLinha
                        chkHP.Tag = strLinha
                        
                        Line Input #liFile, strLinha
                        chkMP.BackColor = CLng(strLinha)
                        
                        Line Input #liFile, strLinha
                        chkMP.Tag = strLinha
                        
                        If Not EOF(liFile) Then
                            For liContador = 0 To 2
                                Line Input #liFile, strLinha
                                chkH(liContador).BackColor = CLng(strLinha)
                                
                                Line Input #liFile, strLinha
                                chkH(liContador).Tag = strLinha
                            Next
                        End If
                        
                Close #liFile
        End If
End Sub

Private Sub saveFile()
'Ask for SaveFileName
        Dim strFilter As String
        Dim strSaveFileName As String, strLinha As String

        strFilter = ahtAddFilterItem(strFilter, "Bots  (*.bot)", "*.bot")
        strSaveFileName = ahtCommonFileOpenSave( _
                                    OpenFile:=False, _
                                    Filter:=strFilter, _
                    Flags:=ahtOFN_OVERWRITEPROMPT Or ahtOFN_READONLY, hwnd:=Me.hwnd)
        'Salva o arquivo

        If strSaveFileName <> "" Then
                Dim liFile As Integer, liContador As Integer
                liFile = FreeFile
                Open strSaveFileName For Output As #liFile
                        For liContador = 0 To txtTeclas.UBound
                        
                                                        strLinha = txtTeclas(liContador)
                                                        Print #liFile, strLinha

                                                        strLinha = txtTempoAcao(liContador)
                                                        Print #liFile, strLinha
                                                        
                                                        strLinha = txtCond(liContador)
                                                        Print #liFile, strLinha
                        Next
                        
                        strLinha = CStr(chkCP.BackColor)
                        Print #liFile, strLinha

                        strLinha = CStr(chkCP.Tag)
                        Print #liFile, strLinha

                        strLinha = CStr(chkHP.BackColor)
                        Print #liFile, strLinha

                        strLinha = CStr(chkHP.Tag)
                        Print #liFile, strLinha

                        strLinha = CStr(chkMP.BackColor)
                        Print #liFile, strLinha

                        strLinha = CStr(chkMP.Tag)
                        Print #liFile, strLinha
                        
                        For liContador = 0 To 2
                            strLinha = CStr(chkH(liContador).BackColor)
                            Print #liFile, strLinha
    
                            strLinha = CStr(chkH(liContador).Tag)
                            Print #liFile, strLinha
                        Next
                        
                        
                Close #liFile
        End If
End Sub
