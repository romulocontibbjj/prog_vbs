Attribute VB_Name = "EmailTestBas"
Option Explicit
'==========================================================================
'Copyright © 2001 by Stan Schultes, All Rights Reserved
'
'   EmailTest.bas
'   General purpose functions
'   EmailTest Send Utility
'   Date Created: 11-Nov-2000
'
'   Notes:
'
'==========================================================================
'Global definitions
Public goReg As CEmailReg
Public goMailCDO As New CEmailCDO
Public goMailOL As New CEmailOL
Public goMailMAPI As New CEmailMAPI

Public Function CheckAttachment(ByVal Attachment As String, AttachPath As String, AttachName As String) As Boolean
'determines Attachment info
    If Len(Attachment) = 0 Then Exit Function
    CheckAttachment = True
    If InStr(Attachment, "\") Then
        AttachPath = Attachment
        AttachName = Mid$(Attachment, InStrRev(Attachment, "\") + 1)
    Else
        AttachPath = App.Path & "\" & Attachment
        AttachName = Attachment
    End If
End Function

Public Function GetImportance(ByVal Importance As String) As Long
'returns CDO Importance constant for CEmailReg Importance string
    Select Case UCase$(Importance)
    Case "LOW"
        GetImportance = CdoLow
    Case "NORMAL"
        GetImportance = CdoNormal
    Case "HIGH"
        GetImportance = CdoHigh
    Case Else
        'default if not valid value
        GetImportance = CdoNormal
    End Select
End Function


