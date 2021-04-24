VERSION 5.00
Object = "{76803B87-EB3A-11D1-9BD3-444553540000}#1.0#0"; "FRMAGIC.OCX"
Begin VB.Form FRMFullDrag 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Formulário_Mágico_RomSoft.ArrasteTotal ArrasteTotal1 
      Left            =   2160
      Top             =   360
      _ExtentX        =   2223
      _ExtentY        =   847
   End
   Begin VB.PictureBox PICTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton BTNEnd 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.CheckBox CHKMoveable 
      Caption         =   "Arraste do Formulário"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4455
      Begin VB.CommandButton BTN1 
         Caption         =   "Mova me"
         Height          =   615
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Click and Drag Here"
         Top             =   240
         Width           =   1575
      End
      Begin VB.PictureBox PIC1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         Picture         =   "FRMFullDrag.frx":0000
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         ToolTipText     =   "Click and Drag Here"
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "FRMFullDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub LBL1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub BTN1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then ArrasteTotal1.ArrasteTotal BTN1
End Sub

Private Sub BTNEnd_Click()
End
End Sub

Private Sub CHKMoveable_Click()

If CHKMoveable.Value = 1 Then
    PICTitle.Visible = False
    MsgBox "Agora você pode arrastar o formulário clicando em qualquer área"
Else
    PICTitle.Visible = True
End If
End Sub

Private Sub Form_Load()
PICTitle.Print "Arraste Total RomSoft"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And CHKMoveable.Value = 1 Then ArrasteTotal1.ArrasteTotal Me
End Sub

Private Sub PIC1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then ArrasteTotal1.ArrasteTotal PIC1
End Sub

Private Sub PICTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then ArrasteTotal1.ArrasteTotal FRMFullDrag
End Sub
