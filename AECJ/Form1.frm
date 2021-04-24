VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   1530
   ClientTop       =   1545
   ClientWidth     =   10380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   10380
   Begin VB.Frame Frame1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin TabDlg.SSTab SSTab1 
         Height          =   7815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   13785
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "ASSOCIADOS"
         TabPicture(0)   =   "Form1.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lab_posto"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmd_gravar"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame4"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmd_sair"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "COLABORADOES"
         TabPicture(1)   =   "Form1.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         Begin VB.CommandButton cmd_sair 
            Caption         =   "&SAIR"
            Height          =   375
            Left            =   8520
            TabIndex        =   30
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Frame Frame4 
            Caption         =   "ASSOCIADOS"
            Height          =   7335
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   9855
            Begin VB.CommandButton cmd_sair2 
               Caption         =   "SAIR"
               Height          =   375
               Left            =   7080
               TabIndex        =   31
               Top             =   3240
               Width           =   1215
            End
            Begin VB.Frame Frame5 
               Caption         =   "Frame5"
               Height          =   1455
               Left            =   600
               TabIndex        =   29
               Top             =   4560
               Width           =   5775
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Command3"
               Height          =   375
               Left            =   5280
               TabIndex        =   27
               Top             =   3240
               Width           =   1215
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Command2"
               Height          =   375
               Left            =   3360
               TabIndex        =   26
               Top             =   3240
               Width           =   1215
            End
            Begin VB.CommandButton cmd_novo 
               Caption         =   "&NOVO"
               Height          =   375
               Left            =   1320
               TabIndex        =   25
               Top             =   3240
               Width           =   1215
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   "AEJC"
               BeginProperty Font 
                  Name            =   "Monotype Corsiva"
                  Size            =   72
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   1695
               Left            =   1680
               TabIndex        =   28
               Top             =   720
               Width           =   6855
            End
         End
         Begin VB.CommandButton cmd_gravar 
            Caption         =   "&GRAVAR"
            Height          =   375
            Left            =   8520
            TabIndex        =   14
            Top             =   960
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Caption         =   "HISTÓRICO DE PELOTÕES"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   120
            TabIndex        =   17
            Top             =   1920
            Width           =   9855
            Begin VB.CommandButton cmd_insere_pel 
               Caption         =   "&INSERIR"
               Height          =   375
               Left            =   960
               TabIndex        =   23
               Top             =   720
               Width           =   2415
            End
            Begin VB.ComboBox cmb_pel2 
               Height          =   315
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   240
               Width           =   2415
            End
            Begin MSDataGridLib.DataGrid grd_pel 
               Bindings        =   "Form1.frx":0038
               Height          =   975
               Left            =   6480
               TabIndex        =   18
               Top             =   360
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1720
               _Version        =   393216
               HeadLines       =   1
               RowHeight       =   15
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DataMember      =   "sel_pel_juco"
               ColumnCount     =   1
               BeginProperty Column00 
                  DataField       =   "PELOTAO"
                  Caption         =   "PELOTAO"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1046
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                     ColumnWidth     =   1739,906
                  EndProperty
               EndProperty
            End
            Begin VB.Label lab_ordem 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1"
               Height          =   255
               Left            =   4200
               TabIndex        =   22
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lab 
               Caption         =   "ORDEM:"
               Height          =   255
               Left            =   3480
               TabIndex        =   21
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label7 
               Caption         =   "PELOTÃO:"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   855
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00800000&
               BackStyle       =   1  'Opaque
               Height          =   1215
               Left            =   6360
               Top             =   240
               Width           =   3255
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1455
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   8175
            Begin VB.TextBox txt_nome_completo 
               Height          =   285
               Left            =   960
               TabIndex        =   3
               Top             =   240
               Width           =   7095
            End
            Begin VB.TextBox txt_nome_guerra 
               Height          =   285
               Left            =   960
               TabIndex        =   5
               Top             =   600
               Width           =   2415
            End
            Begin VB.TextBox txt_numero 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3720
               TabIndex        =   7
               Top             =   600
               Width           =   975
            End
            Begin VB.ComboBox cmb_graduacao 
               Height          =   315
               ItemData        =   "Form1.frx":004F
               Left            =   5880
               List            =   "Form1.frx":007A
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   600
               Width           =   2175
            End
            Begin VB.ComboBox cmb_pelotao 
               Height          =   315
               ItemData        =   "Form1.frx":00FE
               Left            =   960
               List            =   "Form1.frx":0117
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   960
               Width           =   2415
            End
            Begin VB.ComboBox cmb_cmt 
               Height          =   315
               ItemData        =   "Form1.frx":0172
               Left            =   5880
               List            =   "Form1.frx":017C
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "NOME:"
               Height          =   255
               Left            =   360
               TabIndex        =   15
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label2 
               Caption         =   "GUERRA:"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label3 
               Caption         =   "Nº"
               Height          =   255
               Left            =   3480
               TabIndex        =   10
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Label4 
               Caption         =   "GRADUAÇÃO:"
               Height          =   255
               Left            =   4800
               TabIndex        =   8
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "ÚLT. PEL.:"
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label6 
               Caption         =   "FOI CMT:"
               Height          =   255
               Left            =   5160
               TabIndex        =   4
               Top             =   960
               Width           =   735
            End
         End
         Begin VB.Label lab_posto 
            Alignment       =   2  'Center
            Caption         =   "AEJC"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   615
            Left            =   8400
            TabIndex        =   16
            Top             =   360
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_graduacao_Click()

If cmb_graduacao.ListIndex >= 10 Then
    lab_posto.Caption = "OFICIAL SUPERIOR"
ElseIf cmb_graduacao.ListIndex >= 6 Then
    lab_posto.Caption = "OFICIAL SUBALTERNO"
ElseIf cmb_graduacao.ListIndex >= 2 Then
    lab_posto.Caption = "PRAÇA"
Else
    lab_posto.Caption = cmb_graduacao.Text
End If



End Sub

Private Sub cmd_gravar_Click()

deb_aejc.in_associados txt_nome_completo.Text, txt_nome_guerra.Text, txt_numero.Text, _
    cmb_graduacao.Text, cmb_pelotao.Text, cmb_cmt.Text
    
MsgBox cmb_graduacao.Text & " " & txt_nome_guerra & "Nº" & txt_numero.Text & Chr$(13) & _
    "ADICIONADO", vbInformation, "ADICIONADO"
    
txt_numero.Locked = False

Frame3.Enabled = True

cmb_pel2.SetFocus




    


End Sub

Private Sub cmd_insere_pel_Click()

deb_aejc.in_pel_associados lab_ordem.Caption, cmb_pel2.Text, txt_numero

lab_ordem.Caption = lab_ordem + 1

cmb_pel2.SetFocus

Dim x As Integer
If deb_aejc.rssel_pel_juco.State = 1 Then deb_aejc.rssel_pel_juco.Close
    deb_aejc.sel_pel_juco txt_numero
    x = deb_aejc.rssel_pel_juco.RecordCount
    grd_pel.DataMember = "sel_pel_juco"
    grd_pel.Refresh



End Sub

Private Sub cmd_novo_Click()
Frame4.Visible = False
txt_nome_completo.SetFocus
End Sub

Private Sub cmd_sair_Click()
Frame4.Visible = True


End Sub

Private Sub cmd_sair2_Click()
Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub Form_Load()
Dim xpel As Integer
Dim y As Integer
Dim xcmbpelotao As String
x = cmb_pelotao.ListCount

Do While y <= x

xcmbpelotao = cmb_pelotao.List(y)

If xcmbpelotao = "" Then
    Exit Sub
Else
    cmb_pel2.AddItem xcmbpelotao
    y = y + 1
End If

Loop















End Sub
