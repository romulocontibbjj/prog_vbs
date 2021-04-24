VERSION 5.00
Begin VB.UserControl ctl_dynacom 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   ScaleHeight     =   390
   ScaleWidth      =   4365
   ToolboxBitmap   =   "ctl_dynacom.ctx":0000
   Begin VB.ComboBox com_dyna 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmd_clear 
      Height          =   315
      Left            =   1440
      Picture         =   "ctl_dynacom.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar scr_dyna 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "ctl_dynacom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**********************************
'INTRODUCTION
'**********************************
'NAME      : DYNACOM
'PROGRAMMER: NICK PATEMAN
'EMAIL     : np24@blueyonder.co.uk
'RELEASED  : 20 JULY 2001
'----------------------------------
'Thanks for downloading the source
'to dynacom.  I made this little
'addition to the standard combo box
'as I found my self needing these
'functions more and more in my apps
'if you like this code then please
'show appreciation for the sharing
'of it by giving me a vote on PSC
'and FEEDBACK of course ;)
'thanks again!!!

'**********************************
'EVENT DECLARATIONS
'**********************************
Public Event entryadded(entry As String)
Public Event entrynotadded(entry As String)
Public Event clearrequest()
Public Event entrychanged(entry As String)

'**********************************
'PRIVATE VARIABLES NO TO BE ACCESSED
'**********************************
Private ctl_loaded As Boolean

'**********************************
'USER DEFINEABLE VARIABLES
'**********************************
Private max_entries As Integer
Private autocomplete As Boolean
Private clearonreturn As Boolean

'**********************************
'USER DEFINEABLE VARIABLE INTERFACES
'**********************************

'USED TO TURN AUTOCOMPLETE ON/OFF
Public Sub Setautocomplete(enabled As Boolean)
    autocomplete = enabled
End Sub

'USED TO TURN CLEAR ON RETURN ON/OFF
Public Sub Setclearonreturn(enabled As Boolean)
    clearonreturn = enabled
End Sub

'USED TO SET THE MAXIMUM ENTRIES
Public Sub Setmaxentries(value As Integer)
    If value >= 0 Then
        max_entries = value
        If ctl_loaded Then UserControl_Resize
        Cropentries True
    End If
End Sub

'USED TO GET THE MAXIMUM ENTRIES
Public Function Getmaxentries() As Integer
    Getmaxentries = max_entries
End Function

'**********************************
'USER AXCESS FUNCTIONS
'**********************************

'USED TO ADD AN ENTRY IN CODE
Public Function Addentry(entry As String) As Boolean
    Addentry = add_history(com_dyna, entry, max_entries)
End Function

'USED TO CLEAR THE HISTORY ENTRIES
Public Sub Clearentries()
    com_dyna.Clear
    scr_dyna.Max = com_dyna.ListCount
End Sub

'USED TO SAVE THE HISTORY ENTRIES
Public Sub Cropentries(tomaxentries As Boolean, Optional value As Integer)
    If tomaxentries Then
        crop_history com_dyna, max_entries
    Else
        crop_history com_dyna, value
    End If
End Sub

'USED TO DELETE AN ENTRY IN CODE BY STRING
Public Function Delentrystring(entry As String) As Boolean
    If Findentry(entry, True) >= 0 Then
        Delentryindex Findentry(entry, True)
        Delentrystring = True
    End If
End Function

'USED TO DELETE AN ENTRY IN CODE BY INDEX
Public Function Delentryindex(Index As Integer) As Boolean
    If Getnoofentries Then If Index <= Getnoofentries And Index >= 0 Then com_dyna.RemoveItem Index
End Function

'USED TO FIND AN ENTRY BY STRING
Public Function Findentry(entry As String, completematch As Boolean) As Integer
    Dim f_entry As Integer
    For f_entry = 0 To Getnoofentries
        If completematch Then
            If com_dyna.List(f_entry) = entry Then
                Findentry = f_entry
                Exit For
            End If
        Else
            If InStr(1, com_dyna.List(f_entry), entry, vbTextCompare) Then
                Findentry = f_entry
                Exit For
            End If
        End If
        Findentry = -1
    Next f_entry
End Function

'USED TO GET AN ITEM FROM THE LIST
Public Function Getcurrententry() As String
    Getcurrententry = com_dyna.Text
End Function

'USED TO GET AN ITEM FROM THE LIST
Public Function Getentry(Index As Integer) As String
    If Index <= Getnoofentries And Index >= 0 Then Getentry = com_dyna.List(Index)
End Function

'USED TO GET THE NUMBER OF ENTRIES
Public Function Getnoofentries() As Integer
    Getnoofentries = com_dyna.ListCount
End Function

'USED TO SAVE THE HISTORY ENTRIES
Public Sub Loadentries(filename As String)
    If verify_file(filename) Then load_history com_dyna, filename
End Sub

'USED TO SET THE CLEAR BUTTON'S TOOLTIP
Public Sub Setcleartip(tip As String)
    cmd_clear.ToolTipText = tip
End Sub

'USED TO SET THE CLEAR BUTTON'S TOOLTIP
Public Sub Setcombotip(tip As String)
    com_dyna.ToolTipText = tip
End Sub

'USED TO GET AN ITEM FROM THE LIST
Public Sub Setcurrententry(entry As String)
    com_dyna.Text = entry
End Sub

'USED TO SAVE THE HISTORY ENTRIES
Public Sub Saveentries(filename As String)
    If filename <> "" Then save_history com_dyna, filename
End Sub

'**********************************
'COMBO BOX KEYDOWN ROUTINE
'**********************************
Private Sub com_dyna_KeyDown(KeyCode As Integer, Shift As Integer)
    Static btext As String
    Select Case KeyCode
        Case vbKeyReturn
            Dim history_item As String
            history_item = Getcurrententry
            If add_history(com_dyna, history_item, max_entries) Then
                scr_dyna.Max = com_dyna.ListCount
                RaiseEvent entryadded(history_item)
            Else
                RaiseEvent entrynotadded(history_item)
            End If
            If clearonreturn Then com_dyna.Text = ""
        Case vbKeyShift, vbKeyDelete, vbKeyBack
            DoEvents
            btext = Getcurrententry
            Exit Sub
        Case Else
            DoEvents
            RaiseEvent entrychanged(Getcurrententry)
            If Len(btext) < Len(Getcurrententry) Then
                btext = Getcurrententry
                If autocomplete Then autocompleteentry Getcurrententry
            Else
                btext = Getcurrententry
            End If
    End Select
End Sub

'MATCHES THE REST OF A GIVEN STRING TO EXISTING ENTRIES IN THE LIST
Private Sub autocompleteentry(entry As String)
    Dim f_entry As Integer
    For f_entry = 0 To Getnoofentries
        If Left(Getentry(f_entry), Len(entry)) = entry Then
            Setcurrententry entry & Right(Getentry(f_entry), Len(Getentry(f_entry)) - Len(entry))
            com_dyna.SelStart = Len(entry)
            com_dyna.SelLength = Len(Getcurrententry) - Len(entry)
        End If
    Next f_entry
End Sub

'**********************************
'CLEAR BUTTON FUNCTIONS
'**********************************
Private Sub cmd_clear_Click()
    RaiseEvent clearrequest
End Sub

'**********************************
'SCROLL BAR FUNCTIONS
'**********************************
Private Sub scr_dyna_Change()
    com_dyna.Text = com_dyna.List(scr_dyna.value - 1)
End Sub

'**********************************
'USER CONTROL SPECIFIC FUNCTIONS
'**********************************
'INITIALIZE
Private Sub UserControl_Initialize()
    Setmaxentries 0
    ctl_loaded = True
End Sub

'USER CONTROL RESIZE ROUTINE
Private Sub UserControl_Resize()
    If Getmaxentries > 0 Then
        scr_dyna.Visible = True
        cmd_clear.Visible = True
        scr_dyna.Left = UserControl.Width - scr_dyna.Width
        cmd_clear.Left = scr_dyna.Left - cmd_clear.Width
        com_dyna.Width = UserControl.Width - (scr_dyna.Width + cmd_clear.Width)
        UserControl.Height = com_dyna.Height
    ElseIf Getmaxentries = 0 Then
        scr_dyna.Visible = False
        cmd_clear.Visible = False
        com_dyna.Width = UserControl.Width
        UserControl.Height = com_dyna.Height
    End If
End Sub
