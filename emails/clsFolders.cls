VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const REG_SZ = 1 ' Unicode nul terminated string
Private Const REG_BINARY = 3 ' Free form binary
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_USERS = &H80000003
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Enum HKEYS
    CURRENT_USER = HKEY_CURRENT_USER
    LOCAL_MACHINE = HKEY_LOCAL_MACHINE
    USERS = HKEY_USERS
    CLASSES_ROOT = HKEY_CLASSES_ROOT
    CURRENT_CONFIG = HKEY_CURRENT_CONFIG
    DYN_DATA = HKEY_DYN_DATA
    PERFORMANCE_DATA = HKEY_PERFORMANCE_DATA
End Enum

Public Enum SpecialFolders
    [Arquivos de Programas] = 1
    [Arquivos Comuns] = 2
    [Mídia] = 3
    [Raíz do Sistema] = 4
    [Config] = 5
    [Device] = 6
    [WallPaper] = 7
    [Desktop] = 8
    [Templates] = 9
    [App Data] = 10
    [Start Menu] = 11
    [Programs] = 12
    [Startup] = 13
    [Fonts] = 14
    [sendto] = 15
    [Recent] = 16
    [Favorites] = 17
    [Cache] = 18
    [Cookies] = 19
    [History] = 20
    [NetHood] = 21
    [Personal] = 22
    [PrintHood] = 23
    [Local AppData] = 24
    [My Pictures] = 25
    [Administrative Tools] = 26
End Enum

Private Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
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

Public Function GetSpecialFolder(ByVal PathName As SpecialFolders) As String
    Dim Ret, strFolder As String, strPath As String, tmpHKey As HKEYS
    'Open the key
    If PathName >= 8 Then
        strPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
        tmpHKey = CURRENT_USER
    ElseIf PathName <= 7 And PathName > 0 Then
        strPath = "Software\Microsoft\Windows\CurrentVersion"
        tmpHKey = LOCAL_MACHINE
    Else
        strPath = "Software\Microsoft\Windows\CurrentVersion"
        tmpHKey = LOCAL_MACHINE
        PathName = [Raíz do Sistema]
    End If
    RegOpenKey tmpHKey, strPath, Ret
    'Get the key's content
    strFolder = GetStringName(PathName)
    GetSpecialFolder = FixName(RegQueryStringValue(Ret, strFolder))
    'Close the key
    RegCloseKey Ret
End Function

Private Function GetStringName(ByVal intPath As Integer) As String
Dim s As String
Select Case intPath
    Case 1
        s = "ProgramFilesDir"
    Case 2
        s = "CommonFilesDir"
    Case 3
        s = "MediaPath"
    Case 4
        s = "SystemRoot"
    Case 5
        s = "ConfigPath"
    Case 6
        s = "DevicePath"
    Case 7
        s = "WallPaperDir"
    Case 8
        s = "Desktop"
    Case 9
        s = "Templates"
    Case 10
        s = "AppData"
    Case 11
        s = "Start Menu"
    Case 12
        s = "Programs"
    Case 13
        s = "Startup"
    Case 14
        s = "Fonts"
    Case 15
        s = "SendTo"
    Case 16
        s = "Recent"
    Case 17
        s = "Favorites"
    Case 18
        s = "Cache"
    Case 19
        s = "Cookies"
    Case 20
        s = "History"
    Case 21
        s = "NetHood"
    Case 22
        s = "Personal"
    Case 23
        s = "PrintHood"
    Case 24
        s = "Local AppData"
    Case 25
        s = "My Pictures"
    Case 26
        s = "Administrative Tools"
End Select
GetStringName = s
End Function

Private Function FixName(ByVal Name As String) As String
If Not Right(Name, 1) = "\" Then
    FixName = Name & "\"
Else
    FixName = Name
End If
End Function

Public Function Dexist(strDirName As String, Optional Criar As Boolean = False) As Boolean
Dim Dexist1 As Boolean
    Dim strDummy As String

    ' Check for Existence of Directory
    On Error Resume Next
    strDirName = FixName(strDirName)
    strDummy = Dir$(strDirName & "*.*", vbDirectory)
    Dexist1 = Not (strDummy = "")
    If Criar Then
        If Not Dexist1 Then
            MkDir strDirName
            Dexist1 = True
        End If
    End If
    Dexist = Dexist1
    Dir$ AppPath, vbDirectory
    Err = 0
End Function

Public Function Fexist(strFile As String) As Boolean
Dim intFhand As Integer
    ' Check for Existence of File
    On Error Resume Next
    intFhand = FreeFile
    Open strFile For Input As intFhand
    Fexist = IIf(Err = 0, True, False)
    Close intFhand
End Function

Public Function InAppPath(ByVal File As String, Optional Dir As String = "False") As String
Dim strDir As String
If Not Dir = "False" Then
    Dir = FixName(Dir)
    Dexist AppPath & Dir, True
    strDir = AppPath & Dir
Else
    strDir = AppPath()
End If
InAppPath = strDir & File
End Function

Public Function AppPath() As String
AppPath = FixName(App.Path)
End Function
