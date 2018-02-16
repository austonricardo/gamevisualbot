Attribute VB_Name = "util"
Private Type RGBColor
    r As Byte
    G As Byte
    B As Byte
End Type

Private Function fgConvertColor(COR As Long) As RGBColor
    Dim tmpCor  As RGBColor, tmpCor2 As String, i As Integer
    tmpCor2 = Hex$(COR)
    If Len(tmpCor2) < 6 Then tmpCor2 = String(6 - Len(tmpCor2), "0") & tmpCor2
    
    tmpCor.r = CByte(CHex(Right(tmpCor2, 2)))
    tmpCor.G = CByte(CHex(Mid(tmpCor2, 3, 2)))
    tmpCor.B = CByte(CHex(Left(tmpCor2, 2)))
    fgConvertColor = tmpCor
End Function


Public Function fgCompareCores(cor1 As Long, cor2 As Long) As Long
    'Debug.Print cor1
    Dim RGBcor1 As RGBColor, RGBcor2 As RGBColor
    RGBcor1 = fgConvertColor(cor1)
    RGBcor2 = fgConvertColor(cor2)
    fgCompareCores = fgABSDiff(RGBcor1.r, RGBcor2.r) + fgABSDiff(RGBcor1.G, RGBcor2.G) + fgABSDiff(RGBcor1.B, RGBcor2.B)
End Function

Public Function fgABSDiff(ByVal val1 As Byte, ByVal val2 As Byte) As Long
    If val1 > val2 Then
        fgABSDiff = CLng(val1 - val2)
    Else
        fgABSDiff = CLng(val2 - val1)
    End If
End Function


Public Function fgCoresProximas(cor1 As Long, cor2 As Long) As Boolean
    fgCoresProximas = (fgCompareCores(cor1, cor2) <= 10)
End Function

'Purpose   :    Converts a Hex string value to a long
'Inputs    :    sHex                        The hex value to convert to a long. eg. "H1" or "&H1".
'Outputs   :    Returns the numeric value of a string containing a hex value.
'Notes     :

Public Function CHex(sHex As String) As Long
    Dim iNegative As Integer, sPrefixH As String
    
    On Error Resume Next
    iNegative = CBool(Left$(sHex, 1) = "-")
    sPrefixH = IIf(InStr(1, sHex, "H", vbTextCompare), "", "H")
    
    If iNegative Then
        'Negative number
        If Mid$(sHex, 2, 1) = "&" Then
            CHex = CLng("&" & sPrefixH & Mid$(sHex, 3)) * iNegative
        Else
            'Append the ampersand to enable CLng to convert the value
            CHex = CLng("&" & sPrefixH & Mid$(sHex, 2)) * iNegative
        End If
    Else
        'Positive number
        If Left$(sHex, 1) = "&" Then
            CHex = CLng("&" & sPrefixH & Mid$(sHex, 2))
        Else
            'Append the ampersand to enable CLng to convert the value
            CHex = CLng("&" & sPrefixH & sHex)
        End If
    End If
End Function
