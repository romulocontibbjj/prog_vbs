Attribute VB_Name = "MJobCount"
' *************************************************************************
'  Copyright ©2001 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Win32 API declares
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrn As Long, pDefault As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long

'   The data area passed to a system call is too small.
Private Const ERROR_INSUFFICIENT_BUFFER = 122&

' Values used to define DEVMODE structure
Private Const CCHDEVICENAME As Long = 32
Private Const CCHFORMNAME As Long = 32

Private Type ACL
   AclRevision As Byte
   Sbz1 As Byte
   AclSize As Integer
   AceCount As Integer
   Sbz2 As Integer
End Type

Private Type SECURITY_DESCRIPTOR
   Revision As Byte
   Sbz1 As Byte
   Control As Long
   Owner As Long
   Group As Long
   Sacl As ACL
   Dacl As ACL
End Type

Private Type DevMode
   dmDeviceName As String * CCHDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer               ' not exposed
   dmDriverExtra As Integer        ' not exposed
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCHFORMNAME
   dmLogPixels As Integer
   dmBitsPerPel As Long
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmNup As Long            ' union with dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

Private Type PRINTER_INFO_2
   pServerName As String
   pPrinterName As String
   pShareName As String
   pPortName As String
   pDriverName As String
   pComment As String
   pLocation As String
   pDevMode As Long 'DEVMODE
   pSepFile As String
   pPrintProcessor As String
   pDatatype As String
   pParameters As String
   pSecurityDescriptor As Long 'SECURITY_DESCRIPTOR
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   cJobs As Long
   AveragePPM As Long
End Type

Public Sub Main()
   Dim prn As Printer
   For Each prn In Printers
      Debug.Print JobCount(prn.DeviceName)
   Next prn
End Sub

Public Function JobCount(ByVal DevName As String) As Long
   Dim hPrn As Long
   Dim BytesNeeded As Long
   Dim BytesUsed As Long
   Dim pi2 As PRINTER_INFO_2
   Const StrSize As Long = 256

   ' init string elements
   pi2.pServerName = Space$(StrSize)
   pi2.pPrinterName = Space$(StrSize)
   pi2.pShareName = Space$(StrSize)
   pi2.pPortName = Space$(StrSize)
   pi2.pDriverName = Space$(StrSize)
   pi2.pComment = Space$(StrSize)
   pi2.pLocation = Space$(StrSize)
   pi2.pSepFile = Space$(StrSize)
   pi2.pPrintProcessor = Space$(StrSize)
   pi2.pDatatype = Space$(StrSize)
   pi2.pParameters = Space$(StrSize)

   Call OpenPrinter(DevName, hPrn, ByVal 0&)
   If hPrn Then
      If GetPrinter(hPrn, 2, pi2, LenB(pi2), BytesUsed) Then
         JobCount = pi2.cJobs
      Else
         Debug.Print Err.LastDllError
      End If
      Call ClosePrinter(hPrn)
   End If
End Function

'Public Function JobCount(ByVal DevName As String) As Long
'   Dim hPrn As Long
'   Dim BytesNeeded As Long
'   Dim BytesUsed As Long
'   Dim pi2 As PRINTER_INFO_2
'
'   Call OpenPrinter(DevName, hPrn, ByVal 0&)
'   If hPrn Then
'      ReDim Buffer(0 To 0) As Byte
'      ' call once to get proper buffer size
'      Call GetPrinter(hPrn, 2, ByVal 0&, 0, BytesNeeded)
'      If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
'         ReDim Buffer(0 To BytesNeeded - 1) As Byte
'         If GetPrinter(hPrn, 2, Buffer(0), BytesNeeded, BytesUsed) Then
'            'Debug.Print HexDump(VarPtr(Buffer(0)), BytesUsed)
'            pi2.pServerName = PointerToStringA(PointerToDWord(VarPtr(Buffer(0))))
'            pi2.pPrinterName = PointerToStringA(PointerToDWord(VarPtr(Buffer(4))))
'            pi2.pShareName = PointerToStringA(PointerToDWord(VarPtr(Buffer(8))))
'            pi2.pPortName = PointerToStringA(PointerToDWord(VarPtr(Buffer(12))))
'            pi2.pDriverName = PointerToStringA(PointerToDWord(VarPtr(Buffer(16))))
'            pi2.pComment = PointerToStringA(PointerToDWord(VarPtr(Buffer(20))))
'            pi2.pLocation = PointerToStringA(PointerToDWord(VarPtr(Buffer(24))))
'            pi2.pDevMode = PointerToDWord(VarPtr(Buffer(28)))
'            pi2.pSepFile = PointerToStringA(PointerToDWord(VarPtr(Buffer(32))))
'            pi2.pPrintProcessor = PointerToStringA(PointerToDWord(VarPtr(Buffer(36))))
'            pi2.pDatatype = PointerToStringA(PointerToDWord(VarPtr(Buffer(40))))
'            pi2.pParameters = PointerToStringA(PointerToDWord(VarPtr(Buffer(44))))
'            pi2.pSecurityDescriptor = PointerToDWord(VarPtr(Buffer(48)))
'            pi2.Attributes = PointerToDWord(VarPtr(Buffer(52)))
'            pi2.Priority = PointerToDWord(VarPtr(Buffer(56)))
'            pi2.DefaultPriority = PointerToDWord(VarPtr(Buffer(60)))
'            pi2.StartTime = PointerToDWord(VarPtr(Buffer(64)))
'            pi2.UntilTime = PointerToDWord(VarPtr(Buffer(68)))
'            pi2.Status = PointerToDWord(VarPtr(Buffer(72)))
'            pi2.cJobs = PointerToDWord(VarPtr(Buffer(76)))
'            pi2.AveragePPM = PointerToDWord(VarPtr(Buffer(80)))
'            JobCount = pi2.cJobs
'         End If
'      End If
'      Call ClosePrinter(hPrn)
'   End If
'End Function

