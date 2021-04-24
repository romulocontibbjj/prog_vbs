VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCadFormularioDataCanc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "..."
   ClientHeight    =   3510
   ClientLeft      =   5295
   ClientTop       =   2550
   ClientWidth     =   2655
   ControlBox      =   0   'False
   Icon            =   "frmCadFormularioDataCanc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtUF 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2415
   End
   Begin MSMask.MaskEdBox MskData 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2700
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "UF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1140
      TabIndex        =   4
      Top             =   2400
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmCadFormularioDataCanc.frx":000C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmCadFormularioDataCanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdOk_Click()
    If Len(Trim(MskData.Text)) = 0 And Len(Trim(TxtUF.Text)) = 0 Then
    Exit Sub
    Else
    frmCadFormulario.TxtDataAnt.Text = MskData.Text
    frmCadFormulario.TxtUFAnt.Text = TxtUF.Text
    Unload Me
    End If
End Sub

Private Sub MskData_GotFocus()
Call Date_MskEdBox_GotFocus(MskData)
End Sub

Private Sub MskData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub MskData_LostFocus()
Call Date_MskEdBox_LostFocus(MskData)
End Sub
