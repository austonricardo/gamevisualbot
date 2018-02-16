Attribute VB_Name = "tela3"
Public Sub CopyDesktop2(hWnd As Long, _
        ByRef objTo As Object _
    )


SetForegroundWindow hWnd     ' ensure it's fully repainted
'Picture1.Width = IE.Width
'Picture1.Height = IE.Height
BitBlt objTo.hDC, 0, 0, objTo.Width, objTo.Height, hIEDC, 0, 0, &HCC0020
objTo.Refresh
'SetForegroundWindow Me.hWnd    ' restore myself
    End Sub
