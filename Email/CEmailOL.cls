VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmailOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========================================================================
'Copyright © 2001 by Stan Schultes, All Rights Reserved
'
'   CEmailOL.cls
'   Sends Email using Outlook objects
'   EmailTest Send Utility
'   Date Created: 11-Nov-2000
'
'   Notes:
'   This class requires a VB reference to: Microsoft Outlook 9.0 Object Library (MSOUTL9.OLB)
'==========================================================================
'Member variables
Dim m_oOLApp As Outlook.Application
Dim m_oOLMailItem As Outlook.MailItem

Public Function Send(ByVal ListName As String, ByVal ClientActive As Boolean, ByVal Subject As String, ByVal Body As String, Optional ByVal Attachment As String) As Long
'send using Outlook objects
Dim sAddresses() As String
Dim lRecip As Long
Dim sPath As String, sName As String
    On Error Resume Next
    'check that email is enabled for this List
    If Len(ListName) = 0 Then Exit Function
    'Check email list Enabled flag
    goReg.EmailList = ListName
    If Not goReg.Enabled Then Exit Function
    
    'Outlook shows the Logon form if needed, so the ClientActive parameter isn't used
    Set m_oOLMailItem = m_oOLApp.CreateItem(olMailItem)
    With m_oOLMailItem
        .Subject = Subject
        'Set list of recipients & resolve (list is semicolon-delimited)
        sAddresses = Split(goReg.SendTo, ";")
        For lRecip = 0 To UBound(sAddresses, 1)
            .Recipients.Add sAddresses(lRecip)
        Next
        .Recipients.ResolveAll
        'leave space after body text
        .Body = Body & vbCrLf & vbCrLf
        .Importance = GetImportance(goReg.Importance)
        'one attachment in this example
        If CheckAttachment(Attachment, sPath, sName) Then
            .Attachments.Add sPath, olByValue, Len(.Body) + 1, sName
        End If
        'Send the new message
        .Send
    End With
    'return any error code that occurs (0=success)
    Send = Err
End Function

Private Sub Class_Initialize()
'set up for Outlook mail operations
    On Error GoTo CI_Error
    Set m_oOLApp = CreateObject("Outlook.Application")
    Exit Sub
    
CI_Error:
    Err.Raise Err, "CEmailOL", "Startup/Logon failed"
End Sub

Private Sub Class_Terminate()
'destroy objects that might be in use
    Set m_oOLMailItem = Nothing
    Set m_oOLApp = Nothing
End Sub

