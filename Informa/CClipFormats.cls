VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClipFormats"
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

' Local variables
Private m_Fmts As Collection

' ************************************************
'  Initialization and Termination
' ************************************************
Private Sub Class_Initialize()
   Me.Refresh
End Sub

Private Sub Class_Terminate()
   Set m_Fmts = Nothing
End Sub

' ************************************************
'  Properties
' ************************************************
Public Property Get Item(ByVal nFmt As Long) As CClipFormat
   ' Guard against request for non-existant format.
   On Error Resume Next
   Set Item = m_Fmts.Item(CStr(nFmt))
End Property

Public Property Get Count() As Long
   Count = m_Fmts.Count
End Property

' ************************************************
'  Public Methods
' ************************************************
Public Property Get NewEnum() As IUnknown
   ' This property allows you to enumerate
   ' this collection with the For...Each syntax
   Set NewEnum = m_Fmts.[_NewEnum]
End Property

Public Function Present(ByVal nFmt As Long) As Boolean
   ' Test for availability of requested format.
   ' Could map directly to the standard object,
   ' but here's the API equivalent.
   If OpenClipboard(0&) Then
      Present = CBool(IsClipboardFormatAvailable(nFmt))
      Call CloseClipboard
   End If
End Function

Public Function GetName(ByVal nFmt As Long) As String
   Dim nRet
   Dim sFmt As String

   ' Translate numeric constant into textual description.
   Select Case nFmt
      Case CF_TEXT
         sFmt = "Text (TXT)"
      Case CF_BITMAP
         sFmt = "Bitmap (BMP)"
      Case CF_METAFILEPICT
         sFmt = "Metafile (WMF)"
      Case CF_SYLK
         sFmt = "Microsoft Symbolic Link (SYLK)"
      Case CF_DIF
         sFmt = "Data Interchange Format (DIF)"
      Case CF_TIFF
         sFmt = "Tagged Interchange File Format (TIF)"
      Case CF_OEMTEXT
         sFmt = "OEM Text"
      Case CF_DIB
         sFmt = "Device Independent Bitmap (DIB)"
      Case CF_PALETTE
         sFmt = "Palette"
      Case CF_PENDATA
         sFmt = "Pen Data"
      Case CF_RIFF
         sFmt = "RIFF"
      Case CF_WAVE
         sFmt = "Wave"
      Case CF_UNICODETEXT
         sFmt = "Unicode Text"
      Case CF_ENHMETAFILE
         sFmt = "Enhanced Metafile (EMF)"
      Case CF_HDROP
         sFmt = "Dropped Filelist"
      Case CF_LOCALE
         sFmt = "Locale Identifier"
      Case Else
         sFmt = Space(256)
         nRet = GetClipboardFormatName(nFmt, sFmt, Len(sFmt))
         If nRet = 0 Then
            sFmt = "Unknown: " & CStr(nFmt)
         Else
            sFmt = Left$(sFmt, nRet)
         End If
   End Select
   GetName = sFmt
End Function

' ************************************************
'  Friend Methods
' ************************************************
Friend Sub Refresh()
   Dim Fmt As CClipFormat
   Dim nFmt As Long

   ' Reset the collection object
   Set m_Fmts = New Collection

   ' Fill it with each format on the clipboard
   If OpenClipboard(0&) Then
      Do
         Set Fmt = New CClipFormat
         nFmt = EnumClipboardFormats(nFmt)
         If nFmt Then
            Fmt.Format = nFmt
            Fmt.Name = Me.GetName(nFmt)
            m_Fmts.Add Fmt, CStr(nFmt)
         End If
      Loop While nFmt
      Call CloseClipboard
   End If
End Sub




