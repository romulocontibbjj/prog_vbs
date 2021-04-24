Attribute VB_Name = "MPrnDefault"
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
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal Flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function GetDefaultPrinter Lib "winspool.drv" Alias "GetDefaultPrinterA" (ByVal pszBuffer As String, pcchBuffer As Long) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

' Some calls need to know OS
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

' Structure returned by EnumPrinters, level 5
'   typedef struct _PRINTER_INFO_5 { // pri5
'       LPTSTR    pPrinterName;
'       LPTSTR    pPortName;
'       DWORD     Attributes;
'       DWORD     DeviceNotSelectedTimeout;
'       DWORD     TransmissionRetryTimeout;
'   } PRINTER_INFO_5;

' Platform ID constants
Private Const VER_PLATFORM_WIN32s As Long = &H0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = &H1
Private Const VER_PLATFORM_WIN32_NT As Long = &H2

' Used to indicate what to enumerate
Private Const PRINTER_ENUM_DEFAULT         As Long = &H1

Public Function DefaultPrinterName() As String
   ' HOWTO: Retrieve and Set the Default Printer in Windows
   ' http://support.microsoft.com/support/kb/articles/q246/7/72.asp
   ' HOWTO: Get and Set the Default Printer in Windows
   ' http://support.microsoft.com/support/kb/articles/q135/3/87.asp
   Dim Buffer() As Byte
   Dim BufSize As Long
   Dim pPrinterName As Long
   Dim Returned As Long
   Dim Result As String
   Dim os As OSVERSIONINFO

   ' Get OS version info, so we know which way to fork.
   os.dwOSVersionInfoSize = Len(os)
   Call GetVersionEx(os)

   If os.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then  '95/98/ME
      ' Determine how big the buffer needs to be.
      Call EnumPrinters(PRINTER_ENUM_DEFAULT, vbNullString, 5, ByVal 0&, 0, BufSize, Returned)
      If BufSize > 0 Then
         ' Size buffer accordingly, and call again.
         ReDim Buffer(0 To BufSize - 1) As Byte
         If EnumPrinters(PRINTER_ENUM_DEFAULT, vbNullString, 5, Buffer(0), BufSize, BufSize, Returned) Then
            ' A pointer to the default printer name is
            ' returned at the beginning of the buffer.
            Call CopyMemory(pPrinterName, Buffer(0), 4)
            Result = PointerToStringA(pPrinterName)
         End If
      End If

   ElseIf os.dwPlatformId = VER_PLATFORM_WIN32_NT Then
      ' Create satisfactory buffer.
      BufSize = 1024
      Result = Space$(BufSize)

      ' Use either GetDefaultPrinter (2k+) or WIN.INI (NT4).
      If os.dwMajorVersion >= 5 Then
         If GetDefaultPrinter(Result, BufSize) Then
            ' Truncate at first NULL
            Result = Left$(Result, InStr(Result, vbNullChar) - 1)
         End If
      Else 'NT4 or less
         ' Look for default printer in WIN.INI
         ' Returns: "printer name,driver name,port"
         If GetProfileString("Windows", ByVal "device", "", Result, BufSize) Then
            ' Truncate buffer at end of name.
            Result = Left$(Result, InStr(Result, ",") - 1)
         End If
      End If
   End If

   ' Return default printer name.
   DefaultPrinterName = Result
End Function

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

