VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDialog_Table 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table Editor"
   ClientHeight    =   3420
   ClientLeft      =   6120
   ClientTop       =   4035
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTableWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   1080
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4485
      TabIndex        =   7
      Top             =   2985
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      Height          =   375
      Left            =   3300
      TabIndex        =   6
      Top             =   2970
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   180
      TabIndex        =   1
      Top             =   495
      Width           =   5280
      Begin VB.CommandButton cmdBorderColor 
         Height          =   270
         Left            =   4770
         TabIndex        =   19
         Top             =   945
         Width           =   300
      End
      Begin VB.TextBox combo4 
         Height          =   285
         Left            =   3960
         TabIndex        =   16
         Text            =   "1"
         Top             =   1350
         Width           =   1095
      End
      Begin VB.TextBox combo3 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Text            =   "1"
         Top             =   1350
         Width           =   900
      End
      Begin VB.TextBox combo2 
         Height          =   285
         Left            =   3945
         TabIndex        =   14
         Text            =   "1"
         Top             =   600
         Width           =   1125
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   720
         TabIndex        =   10
         Text            =   "Left"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "3"
         Top             =   1830
         Width           =   900
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3960
         TabIndex        =   4
         Text            =   "3"
         Top             =   1830
         Width           =   1095
      End
      Begin VB.Label lblBorderColor 
         Height          =   270
         Left            =   3960
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border Color:"
         Height          =   195
         Left            =   2625
         TabIndex        =   17
         Top             =   1005
         Width           =   1185
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   5100
         Y1              =   1710
         Y2              =   1710
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cell Padding:"
         Height          =   195
         Left            =   2670
         TabIndex        =   13
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cell Spacing:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Border:"
         Height          =   195
         Left            =   3150
         TabIndex        =   11
         Top             =   600
         Width           =   660
      End
      Begin VB.Label Label4 
         Caption         =   "Align:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5160
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Table Setup:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# Of Rows:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# Of Columns:"
         Height          =   195
         Left            =   2535
         TabIndex        =   2
         Top             =   1830
         Width           =   1275
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   2835
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   5001
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Table Editor"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDialog_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ########################################################################
' #                                                                      #
' #            This is the Table dialog used by EasyASP                  #
' #                                                                      #
' #                                                                      #
' #                  Copyright 1999 Eric Banker                          #
' #                  All Rights Reserved                                 #
' #                                                                      #
' ########################################################################

Option Explicit

Private Sub cmdBorderColor_Click()
    cmdialog.flags = cdlCCFullOpen Or cdlCCRGBInit
    'cmdialog.color = Command1.BackColor
    cmdialog.ShowColor
    
    Dim r As Integer
    Dim G As Integer
    Dim B As Integer
    
    ExtractRGB cmdialog.color, r, G, B
    lblBorderColor.BackColor = RGB(r, G, B) '"#" & Format(Hex(r), "00") & Format(Hex(G), "00") & Format(Hex(B), "00")

End Sub

' Respond to the button clicks with these below
' ----------------------------------------------------------------

Private Sub Command1_Click()
On Error Resume Next
    
    ' Create Variables
    Dim Columns As Integer
    Dim rows As Integer
    Dim td As String
    Dim fulltd As String
    Dim align As String
    Dim cellpad As String
    Dim cellspace As String
    Dim border As String
    
    'Put info into variables
    align = Combo1.Text
    border = combo2.Text
    cellspace = combo3.Text
    cellpad = combo4.Text
    
    ' Place text into a varaible
    td = "      <TD></TD>" & vbCrLf
    
    ' Finish with input variables
    Columns = Text2.Text
    rows = Text1.Text
    
    ' Create the table header
    TableHtml = "<Table"
    
    If align = "" Then
    
    Else
        TableHtml = TableHtml & " align=""" & align & """"
    End If
    
    If cellspace = "" Then
    
    Else
        TableHtml = TableHtml & " cellspacing=""" & cellspace & """"
    End If
    
    If cellpad = "" Then
    
    Else
        TableHtml = TableHtml & " cellpadding=""" & cellpad & """"
    End If
    
    If border = "" Then
    
    Else
        TableHtml = TableHtml & " border=""" & border & """"
    End If
    
    TableHtml = TableHtml & " borderColor=""" & lblBorderColor.BackColor & """"
    
    TableHtml = TableHtml & " width=75%>" & vbCrLf
    
    ' Loop to fill up the TD variable
    Do While rows > 0
        fulltd = fulltd & td
        rows = rows - 1
    Loop
    
    ' Loop to do the same for columns
    Do While Columns > 0
        TableHtml = TableHtml & "   <TR>" & vbCrLf & fulltd & "   </TR>" & vbCrLf
        Columns = Columns - 1
    Loop
    
    ' Print out this mother fucker
    TableHtml = TableHtml & "</Table>" & vbCrLf
    
    ' unload the form
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

' Fill the combo box
' ---------------------------------------------------------------

Private Sub Combo1_DropDown()
        Combo1.Clear
        Combo1.AddItem "Left"
        Combo1.AddItem "Center"
        Combo1.AddItem "Right"
End Sub

Private Sub Form_Load()
lblBorderColor.BackColor = vbBlack
End Sub
