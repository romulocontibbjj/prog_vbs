VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_limpaOcorr 
   Caption         =   "LIMPA AT_EDI_OCORR"
   ClientHeight    =   2595
   ClientLeft      =   4560
   ClientTop       =   3720
   ClientWidth     =   4005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   4005
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton cmd_sair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmd_limpaEdi 
         Caption         =   "&Limpa EDI"
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mask_dt2 
         Height          =   300
         Left            =   2520
         TabIndex        =   4
         Top             =   1440
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mask_dt1 
         Height          =   300
         Left            =   480
         TabIndex        =   3
         Top             =   1440
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_cgc 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lab_nome 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "NOME:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Data Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Data Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "CGC:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frm_limpaOcorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_limpaEdi_Click()

deb_edi.Up_atEdi txt_cgc, CDate(mask_dt1), CDate(mask_dt2)

MsgBox "Período Atualizado", vbInformation, "EDI"



End Sub

Private Sub cmd_Sair_Click()
Unload Me

End Sub

Private Sub txt_cgc_KeyDown(KeyCode As Integer, Shift As Integer)

If txt_cgc.Text = "?" Then
    frm_cgc.Show 1
    Exit Sub
End If




If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
    KeyCode = 0
    
    txt_cgc.Text = Trim$(txt_cgc.Text)

If Len(Trim$(txt_cgc)) < 8 Then

    MsgBox "Favor Digite mais " & 8 - Len(Trim$(txt_cgc)) & " numeros Para Continuar a Busca"
    
    txt_cgc.SelStart = 0
    txt_cgc.SelLength = Len(txt_cgc)
    txt_cgc.SetFocus
    
    Exit Sub
    
End If



If deb_edi.rsSel_NomeCli.State = 1 Then deb_edi.rsSel_NomeCli.Close
    deb_edi.Sel_NomeCli Trim$(txt_cgc)
    
    If deb_edi.rsSel_NomeCli.RecordCount > 0 Then
    
        lab_nome.Caption = deb_edi.rsSel_NomeCli.Fields("NOME")
    
    Else
        
        MsgBox "CGC Não Cadastrado", vbInformation, "CGC"
        
    End If
    
End If

End Sub

