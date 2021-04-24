VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form CEmailMAPI 
   Caption         =   "MAPI Controls"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows Default
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1440
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   600
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   0   'False
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "CEmailMAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==========================================================================
'Copyright © 2001 by Stan Schultes, All Rights Reserved
'
'   CEmailMAPI.frm
'   Form "class" that contains VB MAPI controls
'   EmailTest - send message utility
'   Date Created: 11-Nov-2000
'
'   Notes:
'   This form is used instead of a class because the VB MAPI controls need
'   to be sited on a form. The form is named as a class and acts just like a
'   class, except that an instance is destroyed using the Unload statement
'   rather than setting the class variable to Nothing.
'
'   The MAPI controls are loaded from the Components dialog, Microsoft MAPI
'   Controls 6.0 (MSMAPI32.OCX).
'==========================================================================
'
Public Function Send(ByVal ListName As String, ByVal ClientActive As Boolean, ByVal Subject As String, ByVal Body As String, Optional ByVal Attachment As String) As Long
'sends using VB MAPI controls
Dim sAddresses() As String
Dim lRecip As Long
Dim sPath As String, sName As String
    On Error Resume Next
    'check that email is enabled for this List
    If Len(ListName) = 0 Then Exit Function
    'Check email list Enabled flag
    goReg.EmailList = ListName
    If Not goReg.Enabled Then Exit Function
    
    'if email client not running, prompt for Logon
    MAPISession1.LogonUI = Not ClientActive
    With MAPIMessages1
        If .SessionID = 0 Then MAPISession1.SignOn
        'MAPIMEssages.SessionID must be set to work!
        .SessionID = MAPISession1.SessionID
        .Compose
        'don't prompt if names don't resolve
        .AddressResolveUI = False
        .MsgSubject = Subject
        .MsgNoteText = Body
        'Importance isn't implemented in MAPI controls - this is a placeholder
        '.MsgImportance = GetImportance(goReg.Importance)
        
        'email addresses are semicolon-delimited
        sAddresses = Split(goReg.SendTo, ";")
        For lRecip = 0 To UBound(sAddresses, 1)
            .RecipIndex = lRecip
            .RecipDisplayName = sAddresses(lRecip)
            .RecipAddress = sAddresses(lRecip)
            .ResolveName
        Next
        'one attachment in this example
        If CheckAttachment(Attachment, sPath, sName) Then
            .AttachmentIndex = 0
            .AttachmentPathName = sPath
            .AttachmentName = sName
            .AttachmentType = mapData
        End If
        .Send
    End With
    'return any error code that occurs (0=success)
    Send = Err
End Function

'Form_QueryUnload isn't used because the resources are released when
'   the form is unloaded.
