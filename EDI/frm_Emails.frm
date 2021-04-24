VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Emails 
   Caption         =   "CADASTRO DE EMAILS"
   ClientHeight    =   4635
   ClientLeft      =   3840
   ClientTop       =   2445
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6540
   Begin VB.Frame Frame1 
      Height          =   4515
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   6390
      Begin VB.CommandButton cmd_Altera 
         Caption         =   "&Alterar"
         Height          =   315
         Left            =   3675
         TabIndex        =   12
         Top             =   3675
         Width           =   1215
      End
      Begin VB.CommandButton cmd_novo 
         Caption         =   "&Novo"
         Height          =   315
         Left            =   3675
         TabIndex        =   11
         Top             =   4050
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Sair 
         Caption         =   "&Sair"
         Height          =   315
         Left            =   5025
         TabIndex        =   10
         Top             =   4050
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Gravar 
         Caption         =   "&Gravar"
         Height          =   315
         Left            =   5025
         TabIndex        =   9
         Top             =   3675
         Width           =   1215
      End
      Begin VB.TextBox txt_Empresa 
         Height          =   285
         Left            =   3750
         TabIndex        =   8
         Top             =   3300
         Width           =   2490
      End
      Begin VB.TextBox txt_Email 
         Height          =   285
         Left            =   75
         TabIndex        =   6
         Top             =   3825
         Width           =   3390
      End
      Begin VB.TextBox txt_Nome 
         Height          =   285
         Left            =   75
         TabIndex        =   4
         Top             =   3300
         Width           =   3390
      End
      Begin VB.Frame Frame2 
         Height          =   2715
         Left            =   75
         TabIndex        =   1
         Top             =   150
         Width           =   6240
         Begin MSDataGridLib.DataGrid Grd_Email 
            Bindings        =   "frm_Emails.frx":0000
            Height          =   2340
            Left            =   75
            TabIndex        =   2
            Top             =   225
            Width           =   6090
            _ExtentX        =   10742
            _ExtentY        =   4128
            _Version        =   393216
            BackColor       =   12648447
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
            DataMember      =   "Sel_Emails"
            ColumnCount     =   4
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
            BeginProperty Column02 
               DataField       =   "EMPRESA"
               Caption         =   "EMPRESA"
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
               DataField       =   "ID"
               Caption         =   "ID"
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
                  ColumnWidth     =   915,024
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label Label3 
         Caption         =   "EMPRESA:"
         Height          =   165
         Left            =   3750
         TabIndex        =   7
         Top             =   3075
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "EMAIL:"
         Height          =   240
         Left            =   75
         TabIndex        =   5
         Top             =   3600
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "NOME:"
         Height          =   165
         Left            =   75
         TabIndex        =   3
         Top             =   3075
         Width           =   540
      End
   End
End
Attribute VB_Name = "frm_Emails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xalr As Integer
Private Sub cmd_Altera_Click()
Dim xid As Integer

xalr = xalr + 1

If xalr = 1 Then

Destrava_tela (1)
cmd_Altera.Caption = "ALTERA"

ElseIf xalr = 2 Then

    xid = Grd_Email.Columns(3)

    deb_edi.Up_Emails txt_Nome.Text, LCase(Trim$(txt_Email.Text)), UCase(txt_Empresa.Text), xid

        If deb_edi.rsSel_Emails.State = 1 Then deb_edi.rsSel_Emails.Close
            deb_edi.Sel_Emails
        
            Grd_Email.DataMember = "Sel_Emails"
            Grd_Email.Refresh

            xalr = 0
    txt_Email.Text = Empty
    txt_Nome.Text = Empty
    txt_Empresa.Text = Empty
        
    MsgBox "Alterações Concluidas", vbInformation, "EMAILS"
    
    Trava_tela (1)

End If


End Sub

Private Sub cmd_Gravar_Click()
Dim xnome As String
Dim xemail As String
Dim xempresa As String

If Len(Trim$(txt_Nome.Text)) = 0 Then

    MsgBox "Digite o Nome do Cliente", vbInformation, "EMAILS"
    txt_Nome.SetFocus
ElseIf Len(Trim$(txt_Empresa.Text)) = 0 Then

    MsgBox "Digite o Nome da Empresa", vbInformation, "EMAILS"
    txt_Empresa.SetFocus
ElseIf Len(Trim$(txt_Email.Text)) = 0 Then

    MsgBox "Digite o Emal", vbInformation, "EMAILS"
    txt_Email.SetFocus
Else
    xnome = txt_Nome.Text
    xemail = LCase(Trim$(txt_Email.Text))
    xempresa = UCase(txt_Empresa.Text)


    deb_edi.In_Emails xnome, xemail, xempresa
    MsgBox "Email Cadastrado com Sucesso"
    
    txt_Email.Text = Empty
    txt_Nome.Text = Empty
    txt_Empresa.Text = Empty
    
    
    
    If deb_edi.rsSel_Emails.State = 1 Then deb_edi.rsSel_Emails.Close
        deb_edi.Sel_Emails
        
        Grd_Email.DataMember = "Sel_Emails"
        Grd_Email.Refresh
    
    Trava_tela (1)
End If


End Sub

Private Sub cmd_novo_Click()

txt_Email.Text = Empty
txt_Nome.Text = Empty
txt_Empresa.Text = Empty

Destrava_tela (1)
txt_Nome.SetFocus
End Sub

Private Sub cmd_Sair_Click()
Unload Me

End Sub

Private Sub Form_Load()

Trava_tela (1)

End Sub

Private Sub Grd_Email_Click()

txt_Nome.Text = Grd_Email.Columns(0)
txt_Email.Text = Grd_Email.Columns(1)
txt_Empresa.Text = Grd_Email.Columns(2)

End Sub
