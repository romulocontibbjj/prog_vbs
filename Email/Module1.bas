Attribute VB_Name = "Module1"
Public fMainForm As frmMain
Public FileChanged As Boolean
Public FileLocation As String
Public TableHtml As String
Public BackGroundColor As String
Public ColorBody As String
Public ColorLink As String
Public ColorText As String
Public ColorVisited As String
Public ColorActive As String
Public CommandLine As String
Public RecipFlg As Integer
Public Const OpenFilter = " All Web Files (*.asp, *.asa, *.pl, *.htm, *.html, *.css) | *.asp; *.asa; *.pl; *.htm; *.html; *.css; | Asp files (*.asp) | *.asp; | Asa files (*.asa) | *.asa; | Htm files (*.htm) | *.htm | Html files (*.html) | *.html; | Perl files (*.pl) | *.pl; | Style Sheets (*.css) | *.css; | All files (*.*)|*.*|"

Public BoldKey As Boolean
Public ItalicKey As Boolean
Public UnderlineKey As Boolean
Public Const REG_SZ = 1 ' Unicode nul terminated string
Public Const REG_BINARY = 3 ' Free form binary
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public bAuthLogin      As Boolean
Public bPopLogin       As Boolean
Public bHtml           As Boolean
Public MyEncodeType    As ENCODE_METHOD
Public etPriority      As MAIL_PRIORITY
Public bReceipt        As Boolean
Public ServerIP        As String



'Sub Main()
'    Set fMainForm = New frmMain
'    fMainForm.Show
'
'    ' Sets the bold etc keys to false to start out with
'
'End Sub

Public Sub ExtractRGB(ColorVal, r, G, B)
    r = ColorVal And &HFF
    G = (ColorVal \ &H100) And &HFF
    B = (ColorVal \ &H10000) And &HFF
End Sub
Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim ret
    'Open the key
    RegOpenKey hKey, strPath, ret
    'Get the key's content
    GetString = RegQueryStringValue(ret, strValue)
    'Close the key
    RegCloseKey ret
End Function
Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve nformation about the key
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function

Public Function EmailValid(EmailAd As String) As Boolean
Dim StSym As Integer
Dim NxtSym As Integer
Dim DotCheck As Integer
Dim EmailAddr
Dim i As Integer
EmailAddr = Split(EmailAd, ",")
For i = 0 To UBound(EmailAddr)
    StSym = 0
    NxtSym = 0
    DotCheck = 0
    StSym = InStr(1, EmailAddr(i), "@")
    If StSym > 0 Then
      NxtSym = InStr(StSym + 1, EmailAddr(i), "@")
    End If
    DotCheck = InStr(1, EmailAddr(i), ".")
    If StSym = 0 Or NxtSym <> 0 Or DotCheck = 0 Then
      EmailValid = False
      Exit Function
    End If
Next
EmailValid = True
End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
    
    Dim hCurKey As Long
    Dim lRegResult As Long
    
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, _
    ByVal strData, Len(strData))
    
    If lRegResult <> ERROR_SUCCESS Then
        ' Problem in Writing Registry Settings
    End If
    
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function KeyExists() As Boolean
    Dim hCurKey As Long
    Dim lValueType As Long
    Dim lDataBufferSize As Long
    Dim lRegResult As Long
    
    lRegResult = RegOpenKey(HKEY_LOCAL_MACHINE, strKEYPATH, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, "LegalNoticeAlert" & strLe, 0&, _
    lValueType, ByVal 0&, lDataBufferSize)
    
    If lRegResult = ERROR_SUCCESS Then
        KeyExists = True
    End If
    lRegResult = RegCloseKey(hCurKey)
End Function

Public Function SMTPKEY(SMTPADDR As String, SType As Integer) As Boolean
Dim strKEYPATH As String

strKEYPATH = "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system"
RegKeyExist = KeyExists
If Not RegKeyExist Then
    SaveSettingString HKEY_LOCAL_MACHINE, strKEYPATH, "SMTPIP", SMTPADDR & "%" & SType
End If
End Function

Public Function MAILFORMAT(MAILFT As Boolean, SType As ENCODE_METHOD) As Boolean
Dim strKEYPATH As String

strKEYPATH = "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system"
RegKeyExist = KeyExists
If Not RegKeyExist Then
    SaveSettingString HKEY_LOCAL_MACHINE, strKEYPATH, "MAILFORMAT", MAILFT & "%" & SType
End If
End Function


Public Function SenderEmail(SenderName As String, MailAddress As String) As Boolean
Dim strKEYPATH As String

strKEYPATH = "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system"
RegKeyExist = KeyExists
If Not RegKeyExist Then
    SaveSettingString HKEY_LOCAL_MACHINE, strKEYPATH, "SENDER", SenderName & "%" & MailAddress
End If
End Function
