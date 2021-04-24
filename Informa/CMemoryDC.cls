VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemoryDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Win32 API Declarations, Structures, and Constants
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObj Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (Pic As PictDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

' Bitmap Header Definition
Private Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Type PictDesc
   Size As Long
   Type As Long
   hImage As Long
   Data1 As Long
   Data2 As Long
End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

' Member variables
Private m_bmp As BITMAP
Private m_hDC As Long
Private m_hBmp As Long
Private m_UserBmp As Boolean
Private m_hBmpTmp As Long
Private m_Width As Long
Private m_Height As Long
Private m_TermKills As Boolean

' Private constants
Private Const SRCCOPY As Long = &HCC0020

' ****************************************************
'  Initialization and Termination
' ****************************************************
Private Sub Class_Initialize()
   Dim hWndScn As Long
   Dim hDCScn As Long
   
   ' Get desktop DC, and create compatable DC.
   hWndScn = GetDesktopWindow()
   hDCScn = GetDC(hWndScn)
   m_hDC = CreateCompatibleDC(hDCScn)
   
   ' Clean-up.
   Call ReleaseDC(hWndScn, hDCScn)
   
   ' Create initial 1x1 bitmap
   m_Width = 1
   m_Height = 1
   Call RecreateBitmap
   
   ' By default, kill last created
   ' bitmap at Terminate.
   m_TermKills = True
End Sub

Private Sub Class_Terminate()
   ' Clean up resources!
   If m_hBmp Then
      ' Yank bitmap from memory.
      Call SelectObject(m_hDC, m_hBmpTmp)
      ' Only destroy bitmap if not one
      ' the user passed to us.
      If (m_UserBmp And m_TermKills) Xor (m_UserBmp = False) Then
         ' Destroy created bitmap, unless
         ' instructed not to. WARNING:
         ' This may induce a memory leak
         ' if not handled appropriately
         ' outside this class!!!
         Call DeleteObject(m_hBmp)
      End If
   End If
   Call DeleteDC(m_hDC)
End Sub

' ****************************************************
'  Public Properties
' ****************************************************
Public Property Let hBitmap(ByVal hBmp As Long)
   Dim bmp As BITMAP
   
   ' Make sure we have a valid bitmap handle,
   ' or else raise error and bail.
   If GetObj(hBmp, Len(bmp), bmp) Then
      Call UserBitmap(hBmp)
   Else
      Err.Raise Number:=vbObjectError + 1, _
         Source:="CMemoryDC.hBitmap", _
         Description:="Invalid handle."
   End If
End Property

Public Property Get hBitmap() As Long
   hBitmap = m_hBmp
End Property

Public Property Let Height(ByVal NewVal As Long)
   m_Height = NewVal
   Call RecreateBitmap
End Property

Public Property Get Height() As Long
   Height = m_bmp.bmHeight
End Property

Public Property Set Picture(ByVal NewPict As StdPicture)
   ' Create a temporary instance of
   ' this class to measure/blt from.
   Dim dc As New CMemoryDC
   dc.hBitmap = NewPict.Handle
   
   ' Redimension this DC to match.
   m_Height = dc.Height
   m_Width = dc.Width
   Call RecreateBitmap
   
   ' BitBlt from temporary DC.
   Call BitBlt(m_hDC, 0, 0, m_Width, m_Height, dc.hDC, 0, 0, SRCCOPY)
End Property

Public Property Get Picture() As StdPicture
   Dim dc As CMemoryDC
   Dim hBmp As Long
   Dim Pic As PictDesc
   Dim IPic As IPicture
   Dim IID_IDispatch As GUID
   Const PICTYPE_BITMAP = 1
   Const S_OK As Long = 0
   
   ' Create a temporary instance of
   ' this class to blt to.
   Set dc = New CMemoryDC
   dc.Height = m_Height
   dc.Width = m_Width
   dc.TerminateKillsBitmap = False
   
   ' BitBlt to temporary DC, grab handle,
   ' and release.
   Call BitBlt(dc.hDC, 0, 0, m_Width, m_Height, m_hDC, 0, 0, SRCCOPY)
   hBmp = dc.hBitmap
   Set dc = Nothing
   
   ' Fill in with IDispatch Interface ID
   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With
   
   ' Fill PicBmp struct with necessary parts
   With Pic
      .Size = Len(Pic)           ' Length of structure
      .Type = PICTYPE_BITMAP     ' Type of Picture
      .hImage = hBmp             ' Handle to bitmap
      .Data1 = 0                 ' Handle to palette (may be null)
   End With
   
   ' Create and return Picture object
   If OleCreatePictureIndirect(Pic, IID_IDispatch, True, IPic) = S_OK Then
      Set Picture = IPic
   End If
End Property

Public Property Let PixelColor(ByVal Row As Long, ByVal Col As Long, ByVal NewColor As OLE_COLOR)
   Call SetPixel(m_hDC, Row, Col, CheckSysColor(NewColor))
End Property

Public Property Get PixelColor(ByVal Row As Long, ByVal Col As Long) As OLE_COLOR
   PixelColor = GetPixel(m_hDC, Row, Col)
End Property

Public Property Let TerminateKillsBitmap(ByVal NewVal As Boolean)
   ' WARNING: *****************************************
   '  You may not want to destroy created bitmap at
   '  termination some cases, such as when passing
   '  handle to clipboard. However, if you set this
   '  property to False, and don't handle destruction
   '  yourself, you will induce a memory leak!
   ' **************************************************
   m_TermKills = NewVal
End Property

Public Property Get TerminateKillsBitmap() As Boolean
   TerminateKillsBitmap = m_TermKills
End Property

Public Property Let Width(ByVal NewVal As Long)
   m_Width = NewVal
   Call RecreateBitmap
End Property

Public Property Get Width() As Long
   Width = m_bmp.bmWidth
End Property

' ****************************************************
'  Public Properties (read-only)
' ****************************************************
Public Property Get hDC() As Long
   hDC = m_hDC
End Property

Public Property Get bmType() As Long
   bmType = m_bmp.bmType
End Property

Public Property Get WidthBytes() As Long
   WidthBytes = m_bmp.bmWidthBytes
End Property

Public Property Get Planes() As Integer
   Planes = m_bmp.bmPlanes
End Property

Public Property Get BitsPerPixel() As Integer
   BitsPerPixel = m_bmp.bmBitsPixel
End Property

Public Property Get Bits() As Long
   Bits = m_bmp.bmBits
End Property

' ****************************************************
'  Private Methods
' ****************************************************
Private Function CheckSysColor(ByVal Color As Long) As Long
   Const HighBit = &H80000000
   
   ' If high bit set, strip, and get system color.
   If Color And HighBit Then
      CheckSysColor = GetSysColor(Color And Not HighBit)
   Else
      CheckSysColor = Color
   End If
End Function

Private Sub RecreateBitmap()
   Dim hWndScn As Long
   Dim hDCScn As Long
   
   ' Ready to create bitmap yet?
   If m_Width <= 0 Or m_Height <= 0 Then
      Exit Sub
   End If
   
   ' Get current bitmap out of memory DC.
   If m_hBmp Then
      Call SelectObject(m_hDC, m_hBmpTmp)
      ' Destroy existing, if necessary.
      If Not m_UserBmp Then
         Call DeleteObject(m_hBmp)
      End If
   End If
   
   ' Create new bitmap, select it into
   ' memory DC, using desktop DC to insure
   ' we get full-color rather than mono.
   hWndScn = GetDesktopWindow()
   hDCScn = GetDC(hWndScn)
   m_hBmp = CreateCompatibleBitmap(hDCScn, m_Width, m_Height)
   m_hBmpTmp = SelectObject(m_hDC, m_hBmp)
   Call ReleaseDC(hWndScn, hDCScn)
   
   ' Update properties, especially set
   ' flag so we know this is our bitmap!
   Call GetObj(m_hBmp, Len(m_bmp), m_bmp)
   m_UserBmp = False
End Sub

Private Sub UserBitmap(ByVal hBmp As Long)
   ' Get current bitmap out of memory DC.
   Call SelectObject(m_hDC, m_hBmpTmp)
   
   ' Destroy existing, if necessary.
   If Not m_UserBmp Then
      Call DeleteObject(m_hBmp)
   End If
   
   ' Cache handle.
   m_hBmp = hBmp
   
   ' Select new bitmap, and update properties.
   m_hBmpTmp = SelectObject(m_hDC, m_hBmp)
   Call GetObj(m_hBmp, Len(m_bmp), m_bmp)
   
   ' Update tracking variables.
   m_Height = m_bmp.bmHeight
   m_Width = m_bmp.bmWidth
   
   ' Set flag so we know this is our bitmap!
   m_UserBmp = True
End Sub


