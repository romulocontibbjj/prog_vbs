VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolTipEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Enum ToolTipExStyle
[Simple]
[StandAlways]
End Enum



Public Enum WindowsStyles
 WS_MAXIMIZEBOX = &H10000
 WS_MINIMIZEBOX = &H20000
 WS_THICKFRAME = &H40000
 WS_SYSMENU = &H80000
 ws_hscroll = &H100000
 ws_VSCROLL = &H200000
 WS_DLGFRAME = &H400000
 WS_BORDER = &H800000
 WS_MAXIMIZE = &H1000000
' WS_CLIPCHILDREN = &H2000000
' WS_CLIPSIBLINGS = &H4000000
 WS_DISABLED = &H8000000
 ws_VISIBLE = &H10000000
 WS_MINIMIZE = &H20000000
 WS_CHILD = &H40000000
 WS_POPUP = &H80000000
End Enum

Public Enum WindowsExStyles
WS_EX_DLGMODALFRAME = &H1&
WS_EX_NOPARENTNOTIFY = &H4&
WS_EX_TOPMOST = &H8&
WS_EX_ACCEPTFILES = &H10&
WS_EX_TRANSPARENT = &H20&
WS_EX_MDICHILD = &H40&
WS_EX_TOOLWINDOW = &H80&
WS_EX_WINDOWEDGE = &H100&
WS_EX_CLIENTEDGE = &H200&
WS_EX_CONTEXTHELP = &H400&
WS_EX_RIGHT = &H1000&
WS_EX_RTLREADING = &H2000&
WS_EX_LEFTSCROLLBAR = &H4000&
WS_EX_CONTROLPARENT = &H10000
WS_EX_STATICEDGE = &H20000
WS_EX_APPWINDOW = &H40000
End Enum


Public Enum TTS_Styles
TTS_ALWAYSTIP = &H1
 TTS_BALLOON = &H40
 TTS_NOANIMATE = &H10
 TTS_NOFADE = &H20
 TTS_NOPREFIX = &H2
End Enum

Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    lHwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type


Private HDL As Long 'Tool Tip Handle
Dim TT As TOOLINFO
Dim FH As Long 'Font Handle

Private OldProc As Long
Private OldProc2 As Long

Private ParentH As Long
Private ObjectProcedures() As FunctionSPointerS
Private HMEM As Long
Private HWNDPROC As Long

Private HMEM2 As Long
Private HWNDPROC2 As Long

Private TTitle As String
Private Ich As Long

Event MouseMove(ByVal X As Long, ByVal Y As Long)
Event Click()
Event RightClick()
Event DblClick()
Event RightDblClick()
Event MouseLeave()
Event BeginShow()

Private SUBC As ISubC
Implements ISubC

Public Function CreateToolTip(ByVal ParentHwnd As Long, ByVal Style As TTS_Styles, ByVal ToolTipText As String, ByVal ToolTipTitle As String, ByVal IconH As Long, ByVal ToolTipStyle As ToolTipExStyle) As Long
If HDL <> 0 Then DestroyToolTip
HDL = CreateWindowEx(0&, "tooltips_class32", "", Style, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, ParentHwnd, 0, App.hInstance, ByVal 0&)
CreateToolTip = HDL
SetWindowPos HDL, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
TT.lSize = Len(TT)
TT.lHwnd = ParentHwnd
TT.lId = 0
TT.hInstance = App.hInstance
TT.lpStr = ToolTipText

If ToolTipStyle = [Simple] Then
Dim RCX As RECT
GetClientRect ParentHwnd, RCX
TT.lpRect = RCX
TT.lFlags = TTF_SUBCLASS
ElseIf ToolTipStyle = [StandAlways] Then
TT.lFlags = TTF_ABSOLUTE Or TTF_TRACK Or TTF_IDISHWND Or TTF_SUBCLASS
End If

SendMessage HDL, TTM_ADDTOOLA, 0&, TT
SendMessage HDL, TTM_SETMAXTIPWIDTH, 0&, 0 'MultiLine
SendMessage HDL, TTM_SETTITLEA, IconH, ByVal ToolTipTitle
TTitle = ToolTipTitle
Ich = IconH
ParentH = ParentHwnd
OldProc = SetWindowLong(ParentHwnd, -4, HWNDPROC)
OldProc2 = SetWindowLong(HDL, -4, HWNDPROC2)
End Function
Public Sub DestroyToolTip()
If HDL <> 0 Then
DestroyWindow HDL
If OldProc <> 0 Then SetWindowLong ParentH, -4, OldProc
OldProc = 0
OldProc2 = 0
HDL = 0
End If
End Sub
Public Property Let Title(ByVal IconH As Long, ByVal NewTitle As String)
SendMessage HDL, TTM_SETTITLEA, IconH, ByVal NewTitle
SendMessage HDL, TTM_UPDATE, 0, ByVal 0&
TTitle = NewTitle
Ich = IconH
End Property

Public Property Let ForegroundColor(ByVal Bclr As Long)
SendMessage HDL, TTM_SETTIPTEXTCOLOR, Bclr, ByVal 0&
End Property
Public Property Get ForegroundColor() As Long
ForegroundColor = SendMessage(HDL, TTM_GETTIPTEXTCOLOR, 0&, ByVal 0&)
End Property
Public Property Let BackgroundColor(ByVal Fclr As Long)
SendMessage HDL, TTM_SETTIPBKCOLOR, Fclr, ByVal 0&
End Property
Public Property Get BackgroundColor() As Long
BackgroundColor = SendMessage(HDL, TTM_GETTIPBKCOLOR, 0&, ByVal 0&)
End Property


Public Sub ShowToolTip(ByVal hwnd As Long, ByVal ShowTip As Boolean)
SendMessage HDL, TTM_TRACKACTIVATE, ShowTip, TT
End Sub

Public Sub AtPosition(ByVal Left As Long, ByVal Top As Long)
SendMessage HDL, TTM_TRACKPOSITION, 0&, ByVal MakeLong(Left, Top)
End Sub

Public Sub SetToolTipFont(ByVal Fontname As String, ByVal FontSize As Integer, ByVal FontCharSet As Long, ByVal FItalic As Boolean, ByVal FUnderLine As Boolean, ByVal FStrike As Boolean, ByVal FBold As Boolean)
Dim FWX As FontWeights
If FBold Then FWX = FW_BOLD
FH = GetFont(Fontname, FontSize, FontCharSet, FItalic, FUnderLine, FStrike, FWX)
SendMessage HDL, TTM_SETTITLEA, Ich, ByVal TTitle
SendMessage HDL, TTM_UPDATE, 0, ByVal 0&
End Sub
Public Property Let Text(ByVal NewToolTipText As String)
TT.lpStr = NewToolTipText
SendMessage HDL, TTM_UPDATETIPTEXTA, 0&, TT
End Property
Public Property Get Text() As String
Text = TT.lpStr
End Property

Private Sub Class_Initialize()
InitCommonControls
Set SUBC = Me
ObjectProcedures = GetObjectFunctionsPointers(SUBC, 2)
HMEM = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
HMEM2 = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))

HWNDPROC = GlobalLock(HMEM)
HWNDPROC2 = GlobalLock(HMEM2)

DelegateFunction HWNDPROC, SUBC, ObjectProcedures(0).FunctionAddress, 4
DelegateFunction HWNDPROC2, SUBC, ObjectProcedures(1).FunctionAddress, 4

End Sub

Private Sub Class_Terminate()
If HMEM <> 0 Then
If OldProc <> 0 Then SetWindowLong ParentH, -4, OldProc
Call GlobalUnlock(HMEM)
Call GlobalFree(HMEM)
End If

If HMEM2 <> 0 Then
If OldProc2 <> 0 Then SetWindowLong HDL, -4, OldProc2
Call GlobalUnlock(HMEM2)
Call GlobalFree(HMEM2)
End If


If FH <> 0 Then
DeleteObject FH
End If
End Sub

Private Function ISubC_WndProc(ByVal hwnd As Long, ByVal umsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Dim NOTIFY As NMHDR
Select Case umsg


Case WM_NOTIFY
CopyMemory NOTIFY, ByVal lParam, Len(NOTIFY)


Select Case NOTIFY.code

Case TTN_SHOW
RaiseEvent BeginShow

Case NM_CUSTOMDRAW
Dim CDR As ToolTipDraw
CopyMemory CDR, ByVal lParam, Len(CDR)


Select Case CDR.dwDrawStage
Case CDDS_PREPAINT
SelectObject CDR.hdc, FH
ISubC_WndProc = CDRF_NEWFONT 'New Font
Exit Function



End Select
End Select
End Select



ISubC_WndProc = CallWindowProc(OldProc, hwnd, umsg, wParam, lParam)
End Function

Private Function ISubC_WndProc2(ByVal hwnd As Long, ByVal umsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case umsg

Case WM_LBUTTONDOWN
RaiseEvent Click
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
ReleaseCapture
SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
Dim RCX As RECT
GetWindowRect hwnd, RCX
AtPosition RCX.Left, RCX.Top

Case WM_RBUTTONDOWN
RaiseEvent RightClick

Case WM_LBUTTONDBLCLK
RaiseEvent DblClick

Case WM_RBUTTONDBLCLK
RaiseEvent RightDblClick


Case WM_MOUSEMOVE
Dim ML As tagTRACKMOUSEEVENT
ML.cbSize = Len(ML)
ML.hwndTrack = hwnd
ML.dwFlags = TME_LEAVE
ML.dwHoverTime = 1
TrackMouseEvent ML
RaiseEvent MouseMove(GetLO(lParam), GetHI(lParam))

Case WM_MOUSELEAVE
RaiseEvent MouseLeave




End Select
ISubC_WndProc2 = CallWindowProc(OldProc2, hwnd, umsg, wParam, lParam)
End Function
