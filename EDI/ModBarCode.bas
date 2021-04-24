Attribute VB_Name = "ModBarCode"
Sub DrawBarcode(ByVal bc_string As String, obj As Control)
    
    Dim xpos!, y1!, y2!, dw%, th!, tw, new_string$
    
    'define barcode patterns
    Dim bc(90) As String
    bc(1) = "1 1221"            'pre-amble
    bc(2) = "1 1221"            'post-amble
    bc(48) = "11 221"           'digits
    bc(49) = "21 112"
    bc(50) = "12 112"
    bc(51) = "22 111"
    bc(52) = "11 212"
    bc(53) = "21 211"
    bc(54) = "12 211"
    bc(55) = "11 122"
    bc(56) = "21 121"
    bc(57) = "12 121"
                                'capital letters
    bc(65) = "211 12"           'A
    bc(66) = "121 12"           'B
    bc(67) = "221 11"           'C
    bc(68) = "112 12"           'D
    bc(69) = "212 11"           'E
    bc(70) = "122 11"           'F
    bc(71) = "111 22"           'G
    bc(72) = "211 21"           'H
    bc(73) = "121 21"           'I
    bc(74) = "112 21"           'J
    bc(75) = "2111 2"           'K
    bc(76) = "1211 2"           'L
    bc(77) = "2211 1"           'M
    bc(78) = "1121 2"           'N
    bc(79) = "2121 1"           'O
    bc(80) = "1221 1"           'P
    bc(81) = "1112 2"           'Q
    bc(82) = "2112 1"           'R
    bc(83) = "1212 1"           'S
    bc(84) = "1122 1"           'T
    bc(85) = "2 1112"           'U
    bc(86) = "1 2112"           'V
    bc(87) = "2 2111"           'W
    bc(88) = "1 1212"           'X
    bc(89) = "2 1211"           'Y
    bc(90) = "1 2211"           'Z
                                'Misc
    bc(32) = "1 2121"           'space
    bc(35) = ""                 '# cannot do!
    bc(36) = "1 1 1 11"         '$
    bc(37) = "11 1 1 1"         '%
    bc(43) = "1 11 1 1"         '+
    bc(45) = "1 1122"           '-
    bc(47) = "1 1 11 1"         '/
    bc(46) = "2 1121"           '.
    bc(64) = ""                 '@ cannot do!
    bc(65) = "1 1221"           '*
    
    
    
    bc_string = UCase(bc_string)
    
    
    'dimensions
    obj.ScaleMode = 3                               'pixels
    obj.Cls
    obj.Picture = Nothing
    dw = CInt(obj.ScaleHeight / 40)                 'space between bars
    If dw < 1 Then dw = 1
    'Debug.Print dw
    th = obj.TextHeight(bc_string)                  'text height
    tw = obj.TextWidth(bc_string)                   'text width
    new_string = Chr$(1) & bc_string & Chr$(2)      'add pre-amble, post-amble
    
    y1 = obj.ScaleTop
    y2 = obj.ScaleTop + obj.ScaleHeight - 1.5 * th
    obj.Width = 1.1 * Len(new_string) * (15 * dw) * obj.Width / obj.ScaleWidth
    
    
    'draw each character in barcode string
    xpos = obj.ScaleLeft
    For n = 1 To Len(new_string)
        c = Asc(Mid$(new_string, n, 1))
        If c > 90 Then c = 0
        bc_pattern$ = bc(c)
        
        'draw each bar
        For i = 1 To Len(bc_pattern$)
            Select Case Mid$(bc_pattern$, i, 1)
                Case " "
                    'space
                    obj.Line (xpos, y1)-(xpos + 1 * dw, y2), &HFFFFFF, BF
                    xpos = xpos + dw
                    
                Case "1"
                    'space
                    obj.Line (xpos, y1)-(xpos + 1 * dw, y2), &HFFFFFF, BF
                    xpos = xpos + dw
                    'line
                    obj.Line (xpos, y1)-(xpos + 1 * dw, y2), &H0&, BF
                    xpos = xpos + dw
                
                Case "2"
                    'space
                    obj.Line (xpos, y1)-(xpos + 1 * dw, y2), &HFFFFFF, BF
                    xpos = xpos + dw
                    'wide line
                    obj.Line (xpos, y1)-(xpos + 2 * dw, y2), &H0&, BF
                    xpos = xpos + 2 * dw
            End Select
        Next
    Next
    
    '1 more space
    obj.Line (xpos, y1)-(xpos + 1 * dw, y2), &HFFFFFF, BF
    xpos = xpos + dw
    
    'final size and text
    obj.Width = (xpos + dw) * obj.Width / obj.ScaleWidth
    obj.CurrentX = (obj.ScaleWidth - tw) / 2
    obj.CurrentY = y2 + 0.25 * th
    obj.Print bc_string
    
    'copy to clipboard
    obj.Picture = obj.Image
    Clipboard.Clear
    Clipboard.SetData obj.Image, 2



End Sub

