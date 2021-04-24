VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmContacts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contacts"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContacts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMain 
      Height          =   5790
      Left            =   30
      TabIndex        =   13
      Top             =   -45
      Width           =   8100
      Begin VB.CommandButton cmdSearchFile 
         Caption         =   "&From Other File"
         Height          =   375
         Left            =   1920
         TabIndex        =   27
         Top             =   5325
         Width           =   1815
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   4170
         TabIndex        =   5
         Top             =   5325
         Width           =   1200
      End
      Begin MSComctlLib.ListView lstContacts 
         Height          =   4815
         Left            =   75
         TabIndex        =   0
         Top             =   405
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2858
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Email"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Canc&el"
         Height          =   375
         Left            =   6765
         TabIndex        =   7
         Top             =   5325
         Width           =   1200
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear All"
         Height          =   375
         Left            =   5467
         TabIndex        =   6
         Top             =   5325
         Width           =   1200
      End
      Begin VB.CommandButton cmdNewContacts 
         Caption         =   "&Add New Contacts"
         Height          =   375
         Left            =   75
         TabIndex        =   4
         Top             =   5325
         Width           =   1815
      End
      Begin VB.ListBox lstBCC 
         Height          =   1230
         Left            =   4695
         TabIndex        =   20
         Top             =   3960
         Width           =   3255
      End
      Begin VB.ListBox lstCC 
         Height          =   1230
         Left            =   4695
         TabIndex        =   19
         Top             =   2085
         Width           =   3255
      End
      Begin VB.ListBox lstTo 
         Height          =   1230
         Left            =   4695
         TabIndex        =   18
         Top             =   405
         Width           =   3255
      End
      Begin VB.CommandButton cmbBCC 
         Caption         =   "BCC: >>"
         Height          =   325
         Left            =   3615
         TabIndex        =   3
         Top             =   4125
         Width           =   975
      End
      Begin VB.CommandButton cmdCC 
         Caption         =   "CC: >>"
         Height          =   325
         Left            =   3615
         TabIndex        =   2
         Top             =   2445
         Width           =   975
      End
      Begin VB.CommandButton cmdTo 
         Caption         =   "To: >>"
         Height          =   325
         Left            =   3615
         TabIndex        =   1
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BCC List"
         Height          =   195
         Left            =   4695
         TabIndex        =   24
         Top             =   3645
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CC List"
         Height          =   195
         Left            =   4695
         TabIndex        =   23
         Top             =   1845
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To List"
         Height          =   195
         Left            =   4695
         TabIndex        =   22
         Top             =   165
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contacts List"
         Height          =   195
         Left            =   75
         TabIndex        =   21
         Top             =   165
         Width           =   1095
      End
   End
   Begin VB.Frame fraNewContact 
      Height          =   5790
      Left            =   30
      TabIndex        =   15
      Top             =   -45
      Visible         =   0   'False
      Width           =   8100
      Begin VB.CommandButton cmdAddEdit 
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3097
         TabIndex        =   26
         Top             =   1170
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4289
         TabIndex        =   25
         Top             =   1170
         Width           =   1100
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshContacts 
         Height          =   4050
         Left            =   90
         TabIndex        =   14
         Top             =   1635
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   7144
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   -2147483628
         HighLight       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdAddClose 
         Caption         =   "Clos&e"
         Height          =   375
         Left            =   6675
         TabIndex        =   12
         Top             =   1170
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   5481
         TabIndex        =   11
         Top             =   1170
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1905
         TabIndex        =   10
         Top             =   1170
         Width           =   1100
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1890
         TabIndex        =   8
         Top             =   195
         Width           =   4215
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   1890
         TabIndex        =   9
         Top             =   735
         Width           =   4215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail Address"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   795
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   255
         Width           =   495
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selFlg As Boolean
Dim pathstr As String
Dim fs As New FileSystemObject
Dim textline
Private Sub cmbBCC_Click()
    If lstContacts.ListItems.Count = 0 Then
        MsgBox "There is Contacts to add", vbInformation, "Smart Easy E-Mail"
        lstContacts.SetFocus
        Exit Sub
    End If
    For i = 1 To lstContacts.ListItems.Count
        If lstContacts.ListItems(i).Checked = True Then
            selFlg = True
            Exit For
        End If
    Next
    
    If selFlg = False Then
        MsgBox "There is no Contact selected to add", vbInformation, "Smart Easy E-Mail"
        lstContacts.SetFocus
        Exit Sub
    End If
    
    For i = 1 To lstContacts.ListItems.Count
        If lstContacts.ListItems(i).Checked = True Then
            lstBCC.AddItem lstContacts.ListItems(i).ListSubItems(1)
            lstContacts.ListItems(i).Checked = False
        End If
    Next
End Sub

Private Sub cmdAdd_Click()
    If Trim(txtName) = "" Then
        MsgBox "Enter Name", vbInformation, "Smart Easy E-Mail"
        txtName.SetFocus
        Exit Sub
    End If
    If Trim(txtEmail) = "" Then
        MsgBox "Enter Email", vbInformation, "Smart Easy E-Mail"
        txtEmail.SetFocus
        Exit Sub
    End If
    
    If EmailValid(Trim(txtEmail.Text)) = False Then
        MsgBox "Enter Valid E-mail Address", vbInformation, "Smart Easy E-Mail"
        txtEmail.SetFocus
        Exit Sub
    End If
    
    '********* to find existence of the Email ID
    pathstr = Mid(App.Path, 1, 2) & Replace(App.Path & "\" & "Contacts.txt", "\\", "\", 3)
    If fs.FileExists(pathstr) = True Then
        Open pathstr For Input As #1
        Do While Not EOF(1)
            Line Input #1, textline
            If InStr(1, textline, Trim(txtEmail), vbTextCompare) Then
                MsgBox "This Email Address already exists in the Contact List.", vbInformation, "Smart Easy E-Mail"
                txtEmail.SetFocus
                SendKeys "{home}+{end}"
                Close #1
                Exit Sub
            End If
        Loop
        Close #1
    End If
    
    '******** Writing to the Text File

    If fs.FileExists(pathstr) = False Then
        Set a = fs.CreateTextFile(pathstr)
    Else
        Set a = fs.OpenTextFile(pathstr, ForAppending)
    End If
    a.WriteLine Trim(txtName) & "," & Trim(txtEmail)
    a.Close
    
    '********** Adding to the Grid
    mshContacts.TextMatrix(mshContacts.rows - 1, 0) = Trim(txtName)
    mshContacts.TextMatrix(mshContacts.rows - 1, 1) = Trim(txtEmail)
    mshContacts.rows = mshContacts.rows + 1
    
    '*********** Adding New contact to the Listview
    lstContacts.ListItems.Clear
    pathstr = Mid(App.Path, 1, 2) & Replace(App.Path & "\" & "Contacts.txt", "\\", "\", 3)
    Open pathstr For Input As #1
    i = 1
    Do While Not EOF(1)
        Line Input #1, textline
        fileline = Split(Trim(textline), ",")
        lstContacts.ListItems.Add , , fileline(0)
        lstContacts.ListItems(i).ListSubItems.Add , , fileline(1)
        i = i + 1
    Loop
    Close #1
    
    Call cmdAddClear_Click
    
End Sub

Private Sub cmdAddClear_Click()
    txtName.Text = ""
    txtEmail.Text = ""
    txtEmail.Tag = ""
    cmdAddEdit.Enabled = False
    cmdAdd.Enabled = True
End Sub

Private Sub cmdAddClose_Click()
    fraMain.Visible = True
    fraNewContact.Visible = False
End Sub

Private Sub cmdAddDelete_Click()
    If Trim(txtEmail.Tag) = "" Then
        MsgBox "Select the Contact from the List to Delete", vbInformation, "Smart Easy E-Mail"
        mshContacts.SetFocus
        Exit Sub
    End If
    If Trim(txtName.Text) = "" Or Trim(txtEmail.Text) = "" Then
        MsgBox "Select the Contact from the List to Delete", vbInformation, "Smart Easy E-Mail"
        mshContacts.SetFocus
        Exit Sub
    End If
    If MsgBox("Are you sure to Delete this Contact : " & Trim(txtName.Text), vbQuestion + vbYesNo, "Smart Easy E-Mail") = vbNo Then Exit Sub
    
    Dim filecontent As String
    pathstr = Mid(App.Path, 1, 2) & Replace(App.Path & "\" & "Contacts.txt", "\\", "\", 3)
    
    '******** Finding the Corresponding value
    Open pathstr For Input As #1
    Do While Not EOF(1)
        Line Input #1, textline
        fileline = Split(Trim(textline), ",")
        If Trim(fileline(1)) <> txtEmail.Tag Then
            filecontent = filecontent & fileline(0) & "," & fileline(1) & Chr(13)
        End If
    Loop
    If Len(filecontent) > 2 Then
        filecontent = Mid(filecontent, 1, Len(filecontent) - 1)
    Else
        filecontent = ""
    End If
    Close #1
    
    '********* Writing changes to the File
    Set a = fs.OpenTextFile(pathstr, ForWriting)
    fileline = Split(filecontent, Chr(13))
    For i = 0 To UBound(fileline)
        a.WriteLine fileline(i)
    Next
    a.Close
    
    '*********** Refreshing the Listview & Grid
    Call GridClear
    lstContacts.ListItems.Clear
    Open pathstr For Input As #1
    i = 1
    Do While Not EOF(1)
        Line Input #1, textline
        fileline = Split(Trim(textline), ",")
        lstContacts.ListItems.Add , , fileline(0)
        lstContacts.ListItems(i).ListSubItems.Add , , fileline(1)
        
        mshContacts.TextMatrix(mshContacts.rows - 1, 0) = fileline(0)
        mshContacts.TextMatrix(mshContacts.rows - 1, 1) = fileline(1)
        mshContacts.rows = mshContacts.rows + 1
        
        i = i + 1
    Loop
    Close #1
    
    Call cmdAddClear_Click
End Sub

Private Sub cmdAddEdit_Click()
    If Trim(txtEmail.Tag) = "" Then
        MsgBox "Select the Contact from the List to Edit", vbInformation, "Smart Easy E-Mail"
        mshContacts.SetFocus
        Exit Sub
    End If
    If Trim(txtName.Text) = "" Or Trim(txtEmail.Text) = "" Then
        MsgBox "Select the Contact from the List to Edit", vbInformation, "Smart Easy E-Mail"
        mshContacts.SetFocus
        Exit Sub
    End If
        
    Dim filecontent As String
    pathstr = Mid(App.Path, 1, 2) & Replace(App.Path & "\" & "Contacts.txt", "\\", "\", 3)
    
    '******** Finding the Corresponding value
    Open pathstr For Input As #1
    Do While Not EOF(1)
        Line Input #1, textline
        fileline = Split(Trim(textline), ",")
        If Trim(fileline(1)) = txtEmail.Tag Then
            filecontent = filecontent & Trim(txtName.Text) & "," & Trim(txtEmail.Text) & Chr(13)
        Else
            filecontent = filecontent & fileline(0) & "," & fileline(1) & Chr(13)
        End If
    Loop
    filecontent = Mid(filecontent, 1, Len(filecontent) - 1)
    Close #1
    
    '********* Writing changes to the File
    Set a = fs.OpenTextFile(pathstr, ForWriting)
    fileline = Split(filecontent, Chr(13))
    For i = 0 To UBound(fileline)
        a.WriteLine fileline(i)
    Next
    a.Close
    
    mshContacts.TextMatrix(mshContacts.RowSel, 0) = Trim(txtName)
    mshContacts.TextMatrix(mshContacts.RowSel, 1) = Trim(txtEmail)
    
    '*********** Adding New contact to the Listview
    lstContacts.ListItems.Clear
    Open pathstr For Input As #1
    i = 1
    Do While Not EOF(1)
        Line Input #1, textline
        fileline = Split(Trim(textline), ",")
        lstContacts.ListItems.Add , , fileline(0)
        lstContacts.ListItems(i).ListSubItems.Add , , fileline(1)
        i = i + 1
    Loop
    Close #1
    
    Call cmdAddClear_Click
End Sub

Private Sub cmdCC_Click()
    If lstContacts.ListItems.Count = 0 Then
        MsgBox "There is Contacts to add", vbInformation, "Smart Easy E-Mail"
        lstContacts.SetFocus
        Exit Sub
    End If
    For i = 1 To lstContacts.ListItems.Count
        If lstContacts.ListItems(i).Checked = True Then
            selFlg = True
            Exit For
        End If
    Next
    
    If selFlg = False Then
        MsgBox "There is no Contact selected to add", vbInformation, "Smart Easy E-Mail"
        lstContacts.SetFocus
        Exit Sub
    End If
    
    For i = 1 To lstContacts.ListItems.Count
        If lstContacts.ListItems(i).Checked = True Then
            lstCC.AddItem lstContacts.ListItems(i).ListSubItems(1)
            lstContacts.ListItems(i).Checked = False
        End If
    Next
End Sub

Private Sub cmdClear_Click()
    lstTo.Clear
    lstCC.Clear
    lstBCC.Clear
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNewContacts_Click()
    fraMain.Visible = False
    fraNewContact.Visible = True
'    mshContacts.Clear
'    mshContacts.rows = 2
'    mshContacts.Cols = 2
'    mshContacts.TextMatrix(0, 0) = "Name"
'    mshContacts.TextMatrix(0, 1) = "Email"
'    mshContacts.Row = 0
'    mshContacts.ColWidth(0) = 3500
'    mshContacts.ColWidth(1) = 3500
'
'    '********* Reading from the File
'    pathstr = Mid(App.Path, 1, 2) & Replace(App.Path & "\" & "Contacts.txt", "\\", "\", 3)
'    If fs.FileExists(pathstr) = True Then
'    Open pathstr For Input As #1
'    Do While Not EOF(1)
'        Line Input #1, textline
'        fileline = Split(Trim(textline), ",")
'        mshContacts.TextMatrix(mshContacts.rows - 1, 0) = fileline(0)
'        mshContacts.TextMatrix(mshContacts.rows - 1, 1) = fileline(1)
'        mshContacts.rows = mshContacts.rows + 1
'    Loop
'    Close #1
'    End If
End Sub
Private Sub cmdOK_Click()
    Dim ToList As String
    Dim CcList As String
    Dim BccList As String
    If Len(Trim(frmMain.txtTo.Text)) > 4 Then
        ToList = frmMain.txtTo.Text & ","
    End If
    If Len(Trim(frmMain.txtCC.Text)) > 4 Then
        CcList = frmMain.txtCC.Text & ","
    End If
    If Len(Trim(frmMain.txtBCC.Text)) > 4 Then
        BccList = frmMain.txtBCC.Text & ","
    End If
    If lstTo.ListCount = 0 And lstCC.ListCount = 0 And lstBCC.ListCount = 0 Then
        MsgBox "There is no Contacts selected", vbInformation, "Smart Easy Email"
        lstContacts.SetFocus
        Exit Sub
    End If
    For i = 0 To lstTo.ListCount - 1
       If InStr(1, frmMain.txtTo.Text, lstTo.List(i)) = 0 Then
         ToList = ToList & lstTo.List(i) & ","
       End If
    Next
    For i = 0 To lstCC.ListCount - 1
        If InStr(1, frmMain.txtCC.Text, lstCC.List(i)) = 0 Then
            CcList = CcList & lstCC.List(i) & ","
        End If
    Next
    For i = 0 To lstBCC.ListCount - 1
        If InStr(1, frmMain.txtBCC.Text, lstBCC.List(i)) = 0 Then
            BccList = BccList & lstBCC.List(i) & ","
        End If
    Next
    
    If Len(ToList) > 2 Then
        frmMain.txtTo.Text = Mid(ToList, 1, Len(ToList) - 1)
    End If
    If Len(CcList) > 2 Then
        frmMain.txtCC.Text = Mid(CcList, 1, Len(CcList) - 1)
    End If
    If Len(BccList) > 2 Then
        frmMain.txtBCC.Text = Mid(BccList, 1, Len(BccList) - 1)
    End If
    Unload Me
End Sub

Private Sub cmdSearchFile_Click()
    Dim filename As String
    Dim delimit As String
    Dim otherlist As String
    frmMain.dlgCommonDialog.Filter = "Text Files (*.txt)|*.txt|Word Document (*.doc)|*.doc"
    frmMain.dlgCommonDialog.ShowOpen
    filename = frmMain.dlgCommonDialog.filename
    If fs.FileExists(filename) = True Then
        delimit = InputBox("Please Enter valid Delimiter to Search from the File.", "Delimiter", ",")
        If (delimit) = vbNullString Or Len((delimit)) > 1 Then
            MsgBox "Please Enter Valid Delimiter", vbInformation, "Smart Easy E-Mail"
            Exit Sub
        End If
        Open filename For Input As #1
        Do While Not EOF(1)
            Line Input #1, textline
            fileline = Split(textline, delimit, -1, vbTextCompare)
            For i = 0 To UBound(fileline)
                'otherlist = otherlist & fileline(i) & ","
                lstCC.AddItem fileline(i)
            Next
        Loop
        Close #1
'        otherlist = Mid(otherlist, 1, Len(otherlist) - 1)
'        If Len(Trim(txtbcc.Text)) > 0 Then
'            txtbcc.Text = txtbcc.Text & otherlist
'        Else
'            txtbcc.Text = otherlist
'        End If
'        MsgBox otherlist
    End If
End Sub

Private Sub cmdTo_Click()
    If lstContacts.ListItems.Count = 0 Then
        MsgBox "There is Contacts to add", vbInformation, "Smart Easy E-Mail"
        lstContacts.SetFocus
        Exit Sub
    End If
    
    For i = 1 To lstContacts.ListItems.Count
        If lstContacts.ListItems(i).Checked = True Then
            selFlg = True
            Exit For
        End If
    Next
    
    If selFlg = False Then
        MsgBox "There is no Contact selected to add", vbInformation, "Smart Easy E-Mail"
        lstContacts.SetFocus
        Exit Sub
    End If
    
    For i = 1 To lstContacts.ListItems.Count
        If lstContacts.ListItems(i).Checked = True Then
            lstTo.AddItem (lstContacts.ListItems(i).ListSubItems(1))
            lstContacts.ListItems(i).Checked = False
        End If
    Next
    selFlg = False
End Sub

Private Sub Form_Load()
    
    '******** Reading from the Text File
    Call GridClear
    pathstr = Mid(App.Path, 1, 2) & Replace(App.Path & "\" & "Contacts.txt", "\\", "\", 3)
    If fs.FileExists(pathstr) = True Then
        Open pathstr For Input As #1
        i = 1
        Do While Not EOF(1)
            Line Input #1, textline
            fileline = Split(Trim(textline), ",")
            lstContacts.ListItems.Add , , fileline(0)
            lstContacts.ListItems(i).ListSubItems.Add , , fileline(1)
            
            mshContacts.TextMatrix(mshContacts.rows - 1, 0) = fileline(0)
            mshContacts.TextMatrix(mshContacts.rows - 1, 1) = fileline(1)
            mshContacts.rows = mshContacts.rows + 1
            i = i + 1
        Loop
        Close #1
    End If
End Sub

Private Sub lstBCC_DblClick()
    lstBCC.RemoveItem (lstBCC.ListIndex)
End Sub

Private Sub lstCC_DblClick()
    lstCC.RemoveItem (lstCC.ListIndex)
End Sub


Private Sub lstContacts_DblClick()
Dim ExFlg As Integer
If lstContacts.ListItems.Count > 0 Then
    ExFlg = 0
    If RecipFlg = 1 Then
        For i = 0 To lstTo.ListCount - 1
            If lstTo.List(i) = lstContacts.SelectedItem.ListSubItems(1).Text Then
                ExFlg = 1
            End If
        Next
        If ExFlg = 0 Then
            lstTo.AddItem (lstContacts.SelectedItem.ListSubItems(1).Text)
        End If
    End If
    If RecipFlg = 2 Then
        For i = 0 To lstCC.ListCount - 1
            If lstCC.List(i) = lstContacts.SelectedItem.ListSubItems(1).Text Then
                ExFlg = 1
            End If
        Next
        If ExFlg = 0 Then
            lstCC.AddItem (lstContacts.SelectedItem.ListSubItems(1).Text)
        End If
    End If
    If RecipFlg = 3 Then
        For i = 0 To lstBCC.ListCount - 1
            If lstBCC.List(i) = lstContacts.SelectedItem.ListSubItems(1).Text Then
                ExFlg = 1
            End If
        Next
        If ExFlg = 0 Then
            lstBCC.AddItem (lstContacts.SelectedItem.ListSubItems(1).Text)
        End If
    End If

End If
End Sub

Private Sub lstContacts_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
End Sub

Private Sub lstTo_DblClick()
    lstTo.RemoveItem (lstTo.ListIndex)
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Or KeyAscii = 39 Or KeyAscii = 96 Then KeyAscii = 0
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Or KeyAscii = 39 Or KeyAscii = 96 Then KeyAscii = 0
End Sub
Private Sub mshContacts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(mshContacts.TextMatrix(1, 1)) <> "" Then
        txtName.Text = mshContacts.TextMatrix(mshContacts.RowSel, 0)
        txtEmail.Text = mshContacts.TextMatrix(mshContacts.RowSel, 1)
        txtEmail.Tag = mshContacts.TextMatrix(mshContacts.RowSel, 1)
        cmdAddEdit.Enabled = True
        cmdAdd.Enabled = False
    End If
End Sub

Private Sub mshContacts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuList
    End If
End Sub
Private Sub lstContacts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuList
    End If
End Sub
Private Sub mnuDelete_Click()
    If fraNewContact.Visible = False And lstContacts.ListItems.Count > 0 Then
        txtName.Text = lstContacts.SelectedItem
        txtEmail.Text = lstContacts.SelectedItem.ListSubItems(1)
        txtEmail.Tag = lstContacts.SelectedItem.ListSubItems(1)
    ElseIf fraNewContact.Visible = True Then
        Call mshContacts_DblClick
    Else
        Exit Sub
    End If
    Call cmdAddDelete_Click
End Sub
Public Sub GridClear()
    mshContacts.Clear
    mshContacts.rows = 2
    mshContacts.Cols = 2
    mshContacts.TextMatrix(0, 0) = "Name"
    mshContacts.TextMatrix(0, 1) = "Email"
    mshContacts.Row = 0
    mshContacts.ColWidth(0) = 3500
    mshContacts.ColWidth(1) = 3500
End Sub
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
Private Sub mnuEdit_Click()
    If fraNewContact.Visible = False And lstContacts.ListItems.Count > 0 Then
        Call cmdNewContacts_Click
        txtName.Text = lstContacts.SelectedItem
        txtEmail.Text = lstContacts.SelectedItem.ListSubItems(1)
        txtEmail.Tag = lstContacts.SelectedItem.ListSubItems(1)
        cmdAddEdit.Enabled = True
        cmdAdd.Enabled = False
        txtName.SetFocus
    ElseIf fraNewContact.Visible = True Then
        Call mshContacts_DblClick
        txtName.SetFocus
    End If
End Sub
Private Sub mshContacts_DblClick()
    If Trim(mshContacts.TextMatrix(1, 1)) <> "" Then
        txtName.Text = mshContacts.TextMatrix(mshContacts.RowSel, 0)
        txtEmail.Text = mshContacts.TextMatrix(mshContacts.RowSel, 1)
        txtEmail.Tag = mshContacts.TextMatrix(mshContacts.RowSel, 1)
        cmdAdd.Enabled = False
        cmdAddEdit.Enabled = True
    End If
End Sub
