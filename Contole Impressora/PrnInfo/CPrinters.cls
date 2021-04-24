VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©2001 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Win32 API declarations
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

' Member variables
Private m_prns As Collection
Private m_VB5ErrorFlag As Boolean

' Exposed events
Public Event PrinterAdded(ByVal DeviceName As String)

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()
   ' Initialize member objects
   Me.Refresh
End Sub

Private Sub Class_Terminate()
   ' Release member objects
   Set m_prns = Nothing
End Sub

' *********************************************
'  Public Properties (Read-Only)
' *********************************************
Public Property Get Count() As Long
   ' Just return collection size.
   Count = m_prns.Count
End Property

Public Property Get Item(ByVal DevName As String) As CPrinterInfo
Attribute Item.VB_UserMemId = 0
   ' Just allow errors to propogate up.
   Set Item = m_prns(DevName)
End Property

Public Property Get PrintersCollectionBad() As Boolean
   ' FIX: Printers Collection May Not Contain All Printers in the Printers Folder
   ' http://support.microsoft.com/support/kb/articles/Q253/6/12.ASP
   ' This property is meaningless *except* under VB5!!!
   PrintersCollectionBad = m_VB5ErrorFlag
End Property
' *********************************************
'  Public Methods
' *********************************************
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   ' Support enumeration of Item (default).
   Set NewEnum = m_prns.[_NewEnum]
End Function

Public Sub Refresh()
   Dim Buffer As String
   Dim BufSize As Long
   Dim nChars As Long
   Dim Devices() As String
   Dim inf As CPrinterInfo
   Dim i As Long
   
   ' VB5 fails because it never expands this value!
   ' Start with a reasonably sized buffer.
   BufSize = 512
   
   ' Attempt to get list of installed printers
   ' by looping until successful.
   Do
      Buffer = Space$(BufSize)
      nChars = GetProfileString("PrinterPorts", vbNullString, "", Buffer, BufSize)
      If nChars = (BufSize - 2) Then
         ' MSDN: If either lpAppName or lpKeyName is NULL and
         ' the supplied destination buffer is too small to hold
         ' all the strings, the last string is truncated and
         ' followed by two null characters. In this case, the
         ' return value is equal to nSize minus two.
         BufSize = BufSize * 2
      ElseIf nChars = 0 Then
         ' The call failed entirely.
         Exit Do
      Else
         ' We got a reasonable return.
         Exit Do
      End If
   Loop
   
   ' Build a fresh collection
   Set m_prns = New Collection
   Call ExtractStringZ(Buffer, Devices())
   For i = LBound(Devices) To UBound(Devices)
      Set inf = New CPrinterInfo
      inf.DeviceName = Devices(i)
      m_prns.Add inf, Devices(i)
      RaiseEvent PrinterAdded(Devices(i))
   Next i
   
   ' Set flag to indicate this system will
   ' choke if the Printers collection is
   ' queried from VB5.
   m_VB5ErrorFlag = (nChars > (1024 - 2))
End Sub

' *********************************************
'  Private Methods
' *********************************************
Private Function ExtractStringZ(Buffer As String, OutArray() As String) As Long
   Dim StartPos As Long
   Dim NullPos As Long
   Dim BuffLen As Long
   Dim Elements As Long
   
   ' Extract null terminated strings from large
   ' double-null terminated buffer.
   StartPos = 1
   Elements = 0
   BuffLen = Len(Buffer)
   
   ' Loop through buffer looking for nulls.
   Do While StartPos < BuffLen
      NullPos = InStr(StartPos, Buffer, vbNullChar)
      If NullPos = StartPos Then
         ' We've hit the double-null terminator.
         Exit Do
      Else
         ' Expand array, store new substring, and
         ' increment counters.
         ReDim Preserve OutArray(0 To Elements) As String
         OutArray(Elements) = Mid$(Buffer, StartPos, NullPos - StartPos)
         StartPos = NullPos + 1
         Elements = Elements + 1
      End If
   Loop
   
   ' Return number of substrings found.
   ExtractStringZ = Elements
End Function


