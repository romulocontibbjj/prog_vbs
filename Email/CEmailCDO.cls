VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmailCDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========================================================================
'Copyright © 2001 by Stan Schultes, All Rights Reserved
'
'   CEmailCDO.cls
'   Sends Mail with CDO calls
'   EmailTest
'   Date Created: 11-Nov-2000
'
'   Notes:
'- Requires a VB reference to: Microsoft CDO 1.21 Library (CDO.DLL)
'==========================================================================
'Member variables
Private m_oSession As mapi.Session
Private m_oMessage As mapi.Message
Private m_oRecipient As mapi.Recipient
Private m_oAttachment As mapi.Attachment

Public Function Send(ByVal ListName As String, ByVal ClientActive As Boolean, ByVal Subject As String, ByVal Body As String, Optional ByVal Attachment As String) As Long
'send using CDO
Dim sAddresses() As String
Dim lRecip As Long
Dim sName As String, sPath As String
Dim bShow As Boolean
    On Error Resume Next
    If m_oSession Is Nothing Then Exit Function
    'check that email is enabled for this List
    If Len(ListName) = 0 Then Exit Function
    
    'Check email list Enabled flag
    goReg.EmailList = ListName
    If Not goReg.Enabled Then Exit Function
    bShow = True            'if email client is not running
    If ClientActive Then bShow = False
    m_oSession.Logon ShowDialog:=bShow, NewSession:=False
    
    'create the message
    Set m_oMessage = m_oSession.Outbox.Messages.Add
    With m_oMessage
        .Subject = Subject
        .Text = Body
        .Importance = GetImportance(goReg.Importance)
    End With
    'email addresses are semicolon-delimited
    sAddresses = Split(goReg.SendTo, ";")
    For lRecip = 0 To UBound(sAddresses, 1)
        With m_oMessage.Recipients.Add
            .Name = sAddresses(lRecip)
            .Type = CdoTo
            .Resolve
        End With
    Next
    'one Attachment in this example
    If CheckAttachment(Attachment, sPath, sName) Then
        With m_oMessage.Attachments.Add
            .Position = 1
            .Type = CdoFileData
            .Source = sPath
            .Name = sName
            .ReadFromFile sPath
        End With
    End If
    m_oMessage.Send
    'return any error code that occurs (0=success)
    Send = Err
End Function

Private Sub Class_Initialize()
'set up for CDO mail operations
    On Error GoTo CI_Error
    Set m_oSession = CreateObject("MAPI.Session")
    Exit Sub
    
CI_Error:
    Err.Raise Err, "CEmailCDO", "Startup/Logon failed"
End Sub

Private Sub Class_Terminate()
'destroy objects that might be in use
    Set m_oMessage = Nothing
    On Error Resume Next
    m_oSession.Logoff
    Set m_oSession = Nothing
End Sub

