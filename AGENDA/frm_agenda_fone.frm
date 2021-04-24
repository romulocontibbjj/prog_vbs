VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_agenda_fone 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AGENDA TELEFÔNICA"
   ClientHeight    =   3765
   ClientLeft      =   2850
   ClientTop       =   4800
   ClientWidth     =   8385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   8385
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.Frame fra_pesquisa 
         Caption         =   "Perquisa"
         Height          =   2775
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   7935
         Begin VB.CommandButton cmd_seleciona 
            Caption         =   "SELECIONA"
            Height          =   495
            Left            =   6960
            TabIndex        =   25
            Top             =   1200
            Width           =   855
         End
         Begin VB.CommandButton cmd_sair_pesq 
            Caption         =   "&SAIR"
            Height          =   255
            Left            =   6960
            TabIndex        =   24
            Top             =   840
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid grd_pesq 
            Bindings        =   "frm_agenda_fone.frx":0000
            Height          =   2415
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   4260
            _Version        =   393216
            BackColor       =   64
            ForeColor       =   16777215
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
            DataMember      =   "sel_pesq_nome"
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "NOME"
               Caption         =   "NOME"
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
            BeginProperty Column01 
               DataField       =   "TEL"
               Caption         =   "TEL"
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
            BeginProperty Column02 
               DataField       =   "CEL"
               Caption         =   "CEL"
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
            BeginProperty Column03 
               DataField       =   "TELCOM"
               Caption         =   "TELCOM"
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
            BeginProperty Column04 
               DataField       =   "EMAIL"
               Caption         =   "EMAIL"
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
            BeginProperty Column05 
               DataField       =   "OBS"
               Caption         =   "OBS"
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
               BeginProperty Column01 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame fra_procura 
         Caption         =   "Procurar Pelo:"
         Height          =   1335
         Left            =   3240
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
         Begin VB.CommandButton cmd_ok 
            Caption         =   "&OK"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton cmd_cancela 
            Caption         =   "&CANCELA"
            Height          =   255
            Left            =   840
            TabIndex        =   20
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton opt_fone 
            Caption         =   "TELEFONE"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton opt_nome 
            Caption         =   "NOME"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&SAIR"
         Height          =   255
         Left            =   7080
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmd_procurar 
         Caption         =   "&Procurar"
         Height          =   255
         Left            =   7080
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmd_inserir 
         Caption         =   "&INSERIR"
         Height          =   255
         Left            =   7080
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txt_obs 
         Height          =   1455
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2040
         Width           =   5655
      End
      Begin VB.TextBox txt_email 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   1680
         Width           =   5655
      End
      Begin MSMask.MaskEdBox mask_tel_com 
         Height          =   330
         Left            =   5640
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(99)9999-9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mask_cel 
         Height          =   330
         Left            =   3360
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(99)9999-9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mask_tel 
         Height          =   330
         Left            =   1200
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(99)9999-9999"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_nome 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AGENDA TELEFONICA"
         BeginProperty Font 
            Name            =   "BatmanForeverAlternate"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   7935
      End
      Begin VB.Label Label6 
         Caption         =   "OBS:"
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "E-MAIL:"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "TEL_COM:"
         Height          =   255
         Left            =   4800
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "CELULAR:"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "TELEFONE:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "NOME:"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   840
         Width           =   615
      End
   End
End
Attribute VB_Name = "frm_agenda_fone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancela_Click()
fra_procura.Visible = False
End Sub

Private Sub cmd_inserir_Click()

If deb_pend.rssel_pesq_nome.State = 1 Then deb_pend.rssel_pesq_nome.Close
    deb_pend.sel_pesq_nome txt_nome
    
    If deb_pend.rssel_pesq_nome.RecordCount = 1 Then
        MsgBox "Nome já cadastrado", vbInformation, "DUPLICIDADE"
        txt_nome.SelStart = 0
        txt_nome.SelLength = Len(txt_nome.Text)
        txt_nome.SetFocus
    Else

        If mask_cel.Text = "(__)____-____" Then
            mask_cel.Mask = ""
            mask_cel.Text = "Não Tem"
        End If
        
        If mask_tel.Text = "(__)____-____" Then
            mask_tel.Mask = ""
            mask_tel.Text = "Não Tem"
        End If
        


deb_pend.in_fones txt_nome, mask_tel, mask_cel, mask_tel_com, LCase(txt_email), txt_obs

MsgBox txt_nome & Chr$(13) & "CADASTRADO", vbInformation, "CADASTRADO"

txt_nome.Text = Empty
txt_email.Text = Empty
txt_obs.Text = Empty
mask_cel.Mask = ""
mask_cel.Text = Empty
mask_cel.Mask = "(99)9999-9999"
mask_tel.Mask = ""
mask_tel.Text = Empty
mask_tel.Mask = "(99)9999-9999)"
mask_tel_com.Mask = ""
mask_tel_com.Text = Empty
mask_tel_com.Mask = "(99)9999-9999"
txt_nome.SetFocus

End If


End Sub

Private Sub cmd_ok_Click()
Dim xpesq As Integer

If opt_nome.Value = True Then
    xpesq = 1
Else
    xpesq = 2
End If

fra_pesquisa.Visible = True
fra_procura.Visible = False

If deb_pend.rssel_pesq_nome.State = 1 Then deb_pend.rssel_pesq_nome.Close
    deb_pend.sel_pesq_nome "%" & txt_nome & "%"
    
    If deb_pend.rssel_pesq_nome.RecordCount < 1 Then
        MsgBox "Não Há registro", vbInformation, "SEM REGISTRO"
    Else
        grd_pesq.DataMember = "sel_pesq_nome"
        grd_pesq.Refresh
    End If
    

End Sub

Private Sub cmd_procurar_Click()
fra_procura.Visible = True
End Sub

Private Sub cmd_sair_Click()
Unload Me

End Sub

Private Sub cmd_sair_pesq_Click()
fra_pesquisa.Visible = False

End Sub

Private Sub cmd_seleciona_Click()

With deb_pend.rssel_pesq_nome

mask_cel.Mask = ""
mask_tel.Mask = ""
mask_tel_com.Mask = ""

txt_nome.Text = .Fields("NOME")
txt_email.Text = .Fields("EMAIL")
txt_obs.Text = .Fields("OBS")
mask_cel.Text = .Fields("CEL")
mask_tel.Text = .Fields("TEL")
mask_tel_com = .Fields("TELCOM")

End With

grd_pesq.DataMember = ""
grd_pesq.Refresh

fra_pesquisa.Visible = False

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub
