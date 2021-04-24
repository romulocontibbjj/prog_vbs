Attribute VB_Name = "Decl"
  



  
Public Const WM_USER = &H400
Public Const WM_DESTROY = &H2
Public Const WM_CLOSE = &H10
Public Const WM_INITDIALOG = &H110
Public Const WM_SIZE = &H5
Public Const WM_SETREDRAW = &HB
Public Const WM_SIZING = &H214
Public Const WM_ACTIVATE = &H6
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Public Const WM_PAINT = &HF
Public Const WM_NCPAINT = &H85
Public Const WM_ERASEBKGND = &H14
Public Const WM_DRAWITEM = &H2B
Public Const WM_SETTEXT = &HC
Public Const WM_SETICON = &H80
Public Const WM_SETFONT = &H30
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEMOVE = &H200
Public Const WM_CHAR = &H102
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_NOTIFY = &H4E
Public Const WM_COMMAND = &H111
Public Const WM_VSCROLL = &H115
Public Const WM_HSCROLL = &H114
Public Const WM_INITMENU = &H116
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_CREATE = &H1
Public Const WM_MOUSELEAVE = &H2A3
Public Const WM_MOUSEHOVER = &H2A1

 
 
 Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

 Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type
 
 Public Type ToolTipText
    hdr As NMHDR
    lpszText As Long
    szText As String * 80
    hinst As Long
    uFlags As Long
End Type

Public Type ToolTipDraw
hdr As NMHDR
dwDrawStage As Long
hdc As Long
Rct As RECT
dwItemSpec As Long
uItemState As Integer
LItemParam As Long
End Type

Public Type ToolTipDrawState
ToolTipD As ToolTipDraw
uDrawFlags As Integer
End Type

Public Type PAINTSTRUCT
    hdc As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(31) As Byte
End Type

Public Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Public Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long


Public Const CDDS_ITEM = &H10000
Public Const CDDS_MAPPART = &H5
Public Const CDDS_POSTERASE = &H4
Public Const CDDS_POSTPAINT = &H2
Public Const CDDS_PREPAINT = &H1
Public Const CDDS_PREERASE = &H3
Public Const CDDS_ITEMPOSTERASE = (CDDS_ITEM Or CDDS_POSTERASE)
Public Const CDDS_ITEMPOSTPAINT = (CDDS_ITEM Or CDDS_POSTPAINT)
Public Const CDDS_ITEMPREERASE = (CDDS_ITEM Or CDDS_PREERASE)
Public Const CDDS_ITEMPREPAINT = (CDDS_ITEM Or CDDS_PREPAINT)

Public Const CDRF_DODEFAULT = &H0
Public Const CDRF_NEWFONT = &H2
Public Const CDRF_NOTIFYITEMDRAW = &H20
Public Const CDRF_NOTIFYPOSTERASE = &H40
Public Const CDRF_NOTIFYPOSTPAINT = &H10
Public Const CDRF_NOTIFYSUBITEMDRAW = &H20
Public Const CDRF_SKIPDEFAULT = &H4


Public Const HORZSIZE = 4
Public Const VERTSIZE = 6



 Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
 Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

 Declare Sub InitCommonControls Lib "comctl32.dll" ()
 Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
 Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
 Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long


Public Const CW_USEDEFAULT = &H80000000
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1


Public Const CCM_FIRST = &H2000
Public Const CCM_SETWINDOWTHEME = (CCM_FIRST + &HB)

Public Const TTN_FIRST = (-520)
Public Const TTN_GETDISPINFOA = (TTN_FIRST - 0)
Public Const TTN_GETDISPINFOW = (TTN_FIRST - 10)
Public Const TTN_LAST = (-549)
Public Const TTN_LINKCLICK = (TTN_FIRST - 3)
Public Const TTN_NEEDTEXTA = TTN_GETDISPINFOA
Public Const TTN_NEEDTEXTW = TTN_GETDISPINFOW
Public Const TTN_POP = (TTN_FIRST - 2)
Public Const TTN_SHOW = (TTN_FIRST - 1)


Public Const TTM_ACTIVATE = (WM_USER + 1)
Public Const TTM_ADDTOOLA = (WM_USER + 4)
Public Const TTM_ADDTOOLW = (WM_USER + 50)
Public Const TTM_ADJUSTRECT = (WM_USER + 31)
Public Const TTM_DELTOOLA = (WM_USER + 5)
Public Const TTM_DELTOOLW = (WM_USER + 51)
Public Const TTM_ENUMTOOLSA = (WM_USER + 14)
Public Const TTM_ENUMTOOLSW = (WM_USER + 58)
Public Const TTM_GETBUBBLESIZE = (WM_USER + 30)
Public Const TTM_GETCURRENTTOOLA = (WM_USER + 15)
Public Const TTM_GETCURRENTTOOLW = (WM_USER + 59)
Public Const TTM_GETDELAYTIME = (WM_USER + 21)
Public Const TTM_GETMARGIN = (WM_USER + 27)
Public Const TTM_GETMAXTIPWIDTH = (WM_USER + 25)
Public Const TTM_GETTEXTA = (WM_USER + 11)
Public Const TTM_GETTEXTW = (WM_USER + 56)
Public Const TTM_GETTIPBKCOLOR = (WM_USER + 22)
Public Const TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
Public Const TTM_GETTOOLCOUNT = (WM_USER + 13)
Public Const TTM_GETTOOLINFOA = (WM_USER + 8)
Public Const TTM_GETTOOLINFOW = (WM_USER + 53)
Public Const TTM_HITTESTA = (WM_USER + 10)
Public Const TTM_HITTESTW = (WM_USER + 55)
Public Const TTM_NEWTOOLRECTA = (WM_USER + 6)
Public Const TTM_NEWTOOLRECTW = (WM_USER + 52)
Public Const TTM_POP = (WM_USER + 28)
Public Const TTM_POPUP = (WM_USER + 34)
Public Const TTM_RELAYEVENT = (WM_USER + 7)
Public Const TTM_SETDELAYTIME = (WM_USER + 3)
Public Const TTM_SETMARGIN = (WM_USER + 26)
Public Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Public Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Public Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Public Const TTM_SETTITLEA = (WM_USER + 32)
Public Const TTM_SETTITLEW = (WM_USER + 33)
Public Const TTM_SETTOOLINFOA = (WM_USER + 9)
Public Const TTM_SETTOOLINFOW = (WM_USER + 54)
Public Const TTM_SETWINDOWTHEME = CCM_SETWINDOWTHEME
Public Const TTM_TRACKACTIVATE = (WM_USER + 17)
Public Const TTM_TRACKPOSITION = (WM_USER + 18)
Public Const TTM_UPDATE = (WM_USER + 29)
Public Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Public Const TTM_UPDATETIPTEXTW = (WM_USER + 57)
Public Const TTM_WINDOWFROMPOINT = (WM_USER + 16)


Public Const TBS_AUTOTICKS = &H1
Public Const TBS_BOTH = &H8
Public Const TBS_BOTTOM = &H0
Public Const TBS_ENABLESELRANGE = &H20
Public Const TBS_FIXEDLENGTH = &H40
Public Const TBS_HORZ = &H0
Public Const TBS_LEFT = &H4
Public Const TBS_NOTHUMB = &H80
Public Const TBS_NOTICKS = &H10
Public Const TBS_REVERSED = &H200
Public Const TBS_TOOLTIPS = &H100
Public Const TBS_RIGHT = &H0
Public Const TBS_TOP = &H4
Public Const TBS_VERT = &H2


Public Const TTF_ABSOLUTE = &H80
Public Const TTF_CENTERTIP = &H2
Public Const TTF_DI_SETITEM = &H8000
Public Const TTF_IDISHWND = &H1
Public Const TTF_RTLREADING = &H4
Public Const TTF_SUBCLASS = &H10
Public Const TTF_TRACK = &H20
Public Const TTF_TRANSPARENT = &H100

Public Enum FontWeights
FW_DONTCARE = 0
FW_THIN = 100
FW_EXTRALIGHT = 200
FW_LIGHT = 300
FW_NORMAL = 400
FW_MEDIUM = 500
FW_SEMIBOLD = 600
FW_BOLD = 700
FW_EXTRABOLD = 800
FW_HEAVY = 900
End Enum



Public Type SIZE
        cx As Long
        cy As Long
End Type
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long

Public Const TME_HOVER = &H1&
Public Const TME_LEAVE = &H2&
Public Const TME_QUERY = &H40000000
Public Const TME_CANCEL = &H80000000
Public Const HOVER_DEFAULT = &HFFFFFFFF

Public Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Public Declare Function TrackMouseEvent Lib "user32" _
   (lpEventTrack As tagTRACKMOUSEEVENT) As Long


Public Const NM_FIRST = 0
Public Const NM_CUSTOMDRAW = (NM_FIRST - 12)
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
 Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Const LOGPIXELSY = 90
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Function GetFont(ByVal nameX As String, nSize As Integer, ByVal charset As Long, ByVal italicX As Boolean, ByVal underlineX As Boolean, ByVal strikeX As Boolean, ByVal fontW As FontWeights) As Long
Dim DCX As Long
Dim FB As Long
DCX = GetDC(0)
GetFont = CreateFont(-MulDiv(nSize, GetDeviceCaps(DCX, LOGPIXELSY), 72), 0, 0, 0, fontW, italicX, underlineX, strikeX, charset, 0, 0, 2, 1, nameX)
ReleaseDC 0, DCX
End Function

Public Function MakeLong(ByVal Low As Long, ByVal High As Long) As Long
CopyMemory ByVal VarPtr(MakeLong), ByVal VarPtr(Low), 2
CopyMemory ByVal VarPtr(MakeLong) + 2, ByVal VarPtr(High), 2
End Function
Public Function GetHI(ByVal Value As Long) As Long
CopyMemory GetHI, ByVal (VarPtr(Value) + 2), 2
End Function
Public Function GetLO(ByVal Value As Long) As Long
CopyMemory GetLO, ByVal (VarPtr(Value)), 2
End Function
