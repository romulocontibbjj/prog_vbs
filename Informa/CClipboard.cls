VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Clipboard Manager Functions
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GetClipboardOwner Lib "user32" () As Long
Private Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClipboardViewer Lib "user32" () As Long
Private Declare Function ChangeClipboardChain Lib "user32" (ByVal hWnd As Long, ByVal hWndNext As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Function CountClipboardFormats Lib "user32" () As Long
Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetPriorityClipboardFormat Lib "user32" (lpPriorityList As Long, ByVal nCount As Long) As Long
Private Declare Function GetOpenClipboardWindow Lib "user32" () As Long

' Other useful APIs
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

' Predefined Clipboard Formats
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2
Private Const CF_METAFILEPICT = 3
Private Const CF_SYLK = 4
Private Const CF_DIF = 5
Private Const CF_TIFF = 6
Private Const CF_OEMTEXT = 7
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_RIFF = 11
Private Const CF_WAVE = 12
Private Const CF_UNICODETEXT = 13
Private Const CF_ENHMETAFILE = 14
Private Const CF_HDROP = 15
Private Const CF_LOCALE = 16
Private Const CF_MAX = 17

Private Const CF_OWNERDISPLAY = &H80
Private Const CF_DSPTEXT = &H81
Private Const CF_DSPBITMAP = &H82
Private Const CF_DSPMETAFILEPICT = &H83
Private Const CF_DSPENHMETAFILE = &H8E

Private Const GMEM_FIXED = &H0

' Private variables
Private m_Fmts As CClipFormats

' ************************************************
'  Initialization and Termination
' ************************************************
Private Sub Class_Initialize()
   Set m_Fmts = New CClipFormats
End Sub

Private Sub Class_Terminate()
   Set m_Fmts = Nothing
End Sub

' ************************************************
'  Properties
' ************************************************
Public Property Get Formats() As CClipFormats
   ' Expose "New and Improved!" enumeration object.
   m_Fmts.Refresh
   Set Formats = m_Fmts
End Property

Public Property Get Owner() As Long
   ' Return window handle of Clipboard owner.
   ' Useful(?) if conflicts occur.
   Owner = GetClipboardOwner()
End Property

' ************************************************
'  Public Methods
' ************************************************
Public Function Clear() As Boolean
   ' Clear the clipboard of all content. (Same
   ' behavior as standard VB Clipboard object.)
   If OpenClipboard(0&) Then
      Clear = CBool(EmptyClipboard)
      Call CloseClipboard
   End If
End Function

Public Function GetData(Optional ByVal Format As Long = CF_BITMAP) As Picture
   Dim nFmt As Long
   Dim hData As Long
   Dim dc As CMemoryDC
   Dim nRet As Boolean
   
   Select Case Format
      Case CF_BITMAP
         ' Check for desired format.
         nFmt = Me.GetPriorityFormat(CF_BITMAP, CF_DSPBITMAP, CF_DIB)
         
         ' -1=None requested, 0=Empty
         If nFmt > 0 Then
            ' Get handle of clipboard bitmap.
            If OpenClipboard(0&) Then
               hData = GetClipboardData(nFmt)
         
               ' Use class to create disconnected Picture.
               If hData Then
                  Set dc = New CMemoryDC
                  dc.hBitmap = hData
                  dc.TerminateKillsBitmap = False
                  Set GetData = dc.Picture
                  Set dc = Nothing
               End If
               
               ' Clean up
               Call CloseClipboard
            End If
         End If
      Case Else
         ' To be implemented...   :-(
   End Select
End Function

Public Function GetFormat(ByVal Format As Long) As Boolean
   ' Check if the requested format is available
   ' on the clipboard. (Same behavior as standard
   ' VB Clipboard object.)
   If OpenClipboard(0&) Then
      If IsClipboardFormatAvailable(Format) Then
         GetFormat = True
      End If
      Call CloseClipboard
   End If
End Function

Public Function GetPriorityFormat(ParamArray Formats()) As Long
   Dim Fmts() As Long
   Dim i As Long
   Dim nFmt As Long
   
   ' Bail, if no formats were requested.
   If UBound(Formats) < 0 Then Exit Function
   
   ' Transfer desired formats into a non-variant array.
   ReDim Fmts(0 To UBound(Formats)) As Long
   For i = 0 To UBound(Formats)
      ' Double conversion, to be safer.
      ' Could error trap, but that'd mean the
      ' user was a hoser, and we wouldn't want
      ' to insinuate *that*, would we? <g>
      Fmts(i) = CLng(Val(Formats(i)))
   Next i
   
   ' Try opening clipboard...
   If OpenClipboard(0&) Then
      ' Check to see which format is highest in list.
      nFmt = GetPriorityClipboardFormat(Fmts(0), UBound(Fmts) + 1)
      Call CloseClipboard
   Else
      ' Clipboard may already be open by another
      ' routine in same process, try anyway to see
      ' if we can get a successful result.  Not
      ' clean, but worth a shot.
      nFmt = GetPriorityClipboardFormat(Fmts(0), UBound(Fmts) + 1)
   End If
   
   ' Return results
   GetPriorityFormat = nFmt
End Function

Public Function GetText() As String
   Dim nFmt As Long
   Dim hData As Long
   Dim lpData As Long
   
   ' Check for desired format.
   nFmt = Me.GetPriorityFormat(CF_TEXT, CF_UNICODETEXT, CF_OEMTEXT, CF_DSPTEXT)
   
   ' -1=None requested, 0=Empty
   If nFmt > 0 Then
      ' Grab text from clipboard, if available.
      If OpenClipboard(0&) Then
         hData = GetClipboardData(nFmt)
         ' Slurp characters from global memory.
         If hData Then
            lpData = GlobalLock(hData)
            If nFmt = CF_UNICODETEXT Then
               GetText = PointerToStringW(lpData)
            Else
               GetText = PointerToStringA(lpData)
            End If
            Call GlobalUnlock(hData)
         End If
         Call CloseClipboard
      End If
   End If
End Function

Public Function SetData(ByVal NewVal As Picture, Optional ByVal Format As Long = CF_BITMAP) As Boolean
   Dim dc As CMemoryDC
   Dim hData As Long
   Dim nRet As Boolean
   
   Select Case Format
      Case CF_BITMAP
         ' Use class to disconnect image from object.
         Set dc = New CMemoryDC
         Set dc.Picture = NewVal
         dc.TerminateKillsBitmap = False
         hData = dc.hBitmap
         Set dc = Nothing
         
      Case Else
         ' To be implemented...   :-(
   End Select

   ' If we have the handle to data,
   ' attempt to place on clipboard.
   If hData Then
      If OpenClipboard(0&) Then
         nRet = CBool(SetClipboardData(Format, hData))
         Call CloseClipboard
      End If
      ' May need to destroy data object?
      If nRet Then
         SetData = True
      Else
         Call DeleteObject(hData)
      End If
   End If
End Function

Public Function SetText(ByVal NewVal As String) As Boolean
   Dim hData As Long
   Dim lpData As Long
   Dim Buffer() As Byte
   
   ' Try to set text onto clipboard.
   If OpenClipboard(0&) Then
      
      ' Convert data to ANSI byte array.
      Buffer = StrConv(NewVal & vbNullChar, vbFromUnicode)
      ' Allocate enough memory for buffer.
      hData = GlobalAlloc(GMEM_FIXED, UBound(Buffer) + 1)
      If hData Then
         ' Copy data to alloc'd memory.
         lpData = GlobalLock(hData)
         Call CopyMemory(ByVal lpData, Buffer(0), UBound(Buffer) + 1)
         Call GlobalUnlock(hData)
         ' Hand data off to clipboard
         SetText = CBool(SetClipboardData(CF_TEXT, hData))
      End If
      
      ' Place Unicode text on clipboard, too.
      ' Not strictly necessary, as Windows will
      ' convert by default, with above code.
      ' Already null-terminated, so just
      ' allocate sufficient space for copy.
      hData = GlobalAlloc(GMEM_FIXED, LenB(NewVal) + 2)
      If hData Then
         ' Copy data to alloc'd memory.
         lpData = GlobalLock(hData)
         Call CopyMemory(ByVal lpData, ByVal StrPtr(NewVal), LenB(NewVal) + 2)
         Call GlobalUnlock(hData)
         ' Hand data off to clipboard
         Call SetClipboardData(CF_UNICODETEXT, hData)
      End If
      
      Call CloseClipboard
   End If
End Function

' ************************************************
'  Private Methods
' ************************************************
Private Function PointerToStringA(ByVal lpStringA As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringA Then
      nLen = lstrlenA(ByVal lpStringA)
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringA, nLen
         PointerToStringA = StrConv(Buffer, vbUnicode)
      End If
   End If
End Function

Private Function PointerToStringW(ByVal lpStringW As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long

   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringW, nLen
         PointerToStringW = Buffer
      End If
   End If
End Function


