VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImgComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Object used to compare images.
'Uses DIBits API functions for fast image comparison.

Private Type BITMAP '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Function CompareImage(ByVal SourceDC1 As Long, ByVal SourceHandle1 As Long, ByVal SourceDC2 As Long, ByVal SourceHandle2 As Long) As Boolean
On Error GoTo ErrorHandler

Dim udtRGB1() As RGBQUAD, udtRGB2() As RGBQUAD
Dim lonLoop As Long

udtRGB1() = GetRGBQuad(SourceDC1, SourceHandle1)
udtRGB2() = GetRGBQuad(SourceDC2, SourceHandle2)

If Not UBound(udtRGB1) = UBound(udtRGB2) Then Exit Function

For lonLoop = 0 To UBound(udtRGB1)
    
    With udtRGB1(lonLoop)
        
        If Not .rgbBlue = udtRGB2(lonLoop).rgbBlue Or _
            Not .rgbGreen = udtRGB2(lonLoop).rgbGreen Or _
            Not .rgbRed = udtRGB2(lonLoop).rgbRed Then
            
                GoTo IMAGES_NOT_EQUAL
        End If
    
    End With

Next lonLoop

CompareImage = True

Exit Function

IMAGES_NOT_EQUAL:

Exit Function

ErrorHandler:
End Function

Private Function GetRGBQuad(ByVal SourceDC As Long, ByVal SourceHandle As Long) As RGBQUAD()
Dim lonDC As Long, udtBMP As BITMAP
Dim udtInfo As BITMAPINFO, udtData() As RGBQUAD

lonDC = CreateCompatibleDC(SourceDC)
GetObjectAPI SourceHandle, Len(udtBMP), udtBMP

With udtInfo.bmiHeader
    .biBitCount = 32
    .biPlanes = 1
    .biSize = 40
    .biWidth = udtBMP.bmWidth
    .biHeight = udtBMP.bmHeight
    ReDim udtData((.biWidth * .biHeight) - 1)
End With

GetDIBits lonDC, SourceHandle, 0, udtBMP.bmHeight, udtData(0), udtInfo, 0
DeleteDC lonDC

GetRGBQuad = udtData()
End Function

