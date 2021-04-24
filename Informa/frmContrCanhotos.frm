VERSION 5.00
Begin VB.Form frmContrCanhotos 
   Caption         =   "Controle de Canhotos de NFs"
   ClientHeight    =   5655
   ClientLeft      =   3015
   ClientTop       =   1065
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   5550
   Begin VB.Frame Frame1 
      Caption         =   "Controle de Canhotos das NF dos CTC de POD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar/Voltar"
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Frame fraFaltantes 
         Caption         =   "0 Canhotos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   3600
         TabIndex        =   8
         Top             =   480
         Width           =   1575
         Begin VB.ListBox lstFaltantes 
            Height          =   3375
            ItemData        =   "frmContrCanhotos.frx":0000
            Left            =   120
            List            =   "frmContrCanhotos.frx":0002
            TabIndex        =   4
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Canhotos Faltantes no CTC de POD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraPresentes 
         Caption         =   "0 Canhotos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1575
         Begin VB.ListBox lstPresentes 
            Height          =   3375
            ItemData        =   "frmContrCanhotos.frx":0004
            Left            =   120
            List            =   "frmContrCanhotos.frx":0006
            TabIndex        =   2
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Canhotos Presentes com o CTC de POD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "<<--  Voltar  <<--"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "-->>  Excluir  -->>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblFilialctc 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Filial CTC:"
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
         Left            =   2160
         TabIndex        =   10
         Top             =   960
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmContrCanhotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    frmPod.lblcontroletela = "cancelar"
    Me.Hide
End Sub

Private Sub cmdExcluir_Click()
    lstFaltantes.AddItem lstPresentes.Text
    lstPresentes.RemoveItem lstPresentes.ListIndex
    fraPresentes.Caption = lstPresentes.ListCount & " Canhotos"
    fraFaltantes.Caption = lstFaltantes.ListCount & " Canhotos"
    cmdOk.SetFocus
End Sub

Private Sub cmdExcluir_LostFocus()
    cmdExcluir.Enabled = False
End Sub

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub cmdVoltar_Click()
    lstPresentes.AddItem lstFaltantes.Text
    lstFaltantes.RemoveItem lstFaltantes.ListIndex
    fraPresentes.Caption = lstPresentes.ListCount & " Canhotos"
    fraFaltantes.Caption = lstFaltantes.ListCount & " Canhotos"
    cmdOk.SetFocus
End Sub

Private Sub cmdVoltar_LostFocus()
    cmdVoltar.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmContrCanhotos = Nothing
End Sub

Private Sub lstFaltantes_Click()
    cmdVoltar.Enabled = True
    cmdVoltar.SetFocus
End Sub

Private Sub lstPresentes_Click()
    cmdExcluir.Enabled = True
    cmdExcluir.SetFocus
End Sub
