VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmailReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========================================================================
'Copyright © 2001 by Stan Schultes, All Rights Reserved
'
'   CEmailReg.cls
'   Email Registry Interface
'   EmailConfig - Email Send Configuration Utility
'   Date Created: 11-Nov-2000
'
'   Notes:
'The Registry configuration for the Email Send process looks like this:
'   HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ProgName\EmailList
'where: ProgName is the program name set through the ProgramName property,
'   and EmailList is the name of an Email send list set through the
'       EmailList property.
'Once the ProgramName (defaults to app.exename) and EmailList properties of this
'   class are set, accessing the property values reads/writes the registry settings
'   directly. The properties values for a ProgramName/EmailList include:
'   > SendTo - semicolon-delimited list of email addresses or distribution list names
'   > Importance - string value of High, Normal, Low
'   > Enabled - 0 to disable, 1 to enable
'
'==========================================================================
'member variables
Private m_sProgramName As String
Private m_sEmailList As String

Public Property Get ProgramName() As String
'Returns the current ProgramName, defaults to App.EXEName
    If Len(m_sProgramName) Then
        ProgramName = m_sProgramName
    Else
        ProgramName = App.EXEName
    End If
End Property

Public Property Let ProgramName(ByVal ProgName As String)
'Sets the ProgramName
    m_sProgramName = ProgName
End Property

Public Property Get EmailList() As String
'Returns the current EmailList, defaults to "Default"
    If Len(m_sEmailList) Then
        EmailList = m_sEmailList
    Else
        EmailList = "Default"
    End If
End Property

Public Property Let EmailList(ByVal ListName As String)
'Sets the EmailList
    m_sEmailList = ListName
End Property

Public Property Get SendTo() As String
'Returns the SendTo list (semicolon-delimited)
    SendTo = GetSetting(ProgramName, EmailList, "SendTo", "")
End Property

Public Property Let SendTo(ByVal SendToList As String)
'Sets the SendTo list (semicolon-delimited)
    If Len(SendToList) Then
        SaveSetting ProgramName, EmailList, "SendTo", SendToList
    End If
End Property

Public Property Get Importance() As String
'Returns the Importance setting
    Importance = GetSetting(ProgramName, EmailList, "Importance", "Normal")
End Property

Public Property Let Importance(ByVal ImportanceValue As String)
'Sets the Importance setting
    Select Case UCase$(ImportanceValue)
    Case "LOW", "NORMAL", "HIGH"
        SaveSetting ProgramName, EmailList, "Importance", ImportanceValue
    End Select
End Property

Public Property Get Enabled() As Boolean
'Returns the Enabled setting
    Enabled = CBool(GetSetting(ProgramName, EmailList, "Enabled", "0"))
End Property

Public Property Let Enabled(ByVal EnabledValue As Boolean)
'Sets the Enabled setting
    SaveSetting ProgramName, EmailList, "Enabled", IIf(EnabledValue, "1", "0")
End Property

