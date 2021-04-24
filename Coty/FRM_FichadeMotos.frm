VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FRM_FichadeMotos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha de Motos"
   ClientHeight    =   5430
   ClientLeft      =   2145
   ClientTop       =   2595
   ClientWidth     =   9165
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5459.754
   ScaleMode       =   0  'User
   ScaleWidth      =   13071.39
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameDadosCadastrais 
      Caption         =   " Dados Cadastrais de Motos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin MSMask.MaskEdBox mask_ano 
         Height          =   300
         Left            =   6120
         TabIndex        =   7
         Top             =   1200
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "99/99"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmb_Uf 
         Height          =   315
         ItemData        =   "FRM_FichadeMotos.frx":0000
         Left            =   6120
         List            =   "FRM_FichadeMotos.frx":0055
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txt_Cidade 
         Height          =   285
         Left            =   3720
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txt_numero 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txt_letra 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   1
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cmb_motoqueiros 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Frame Frame_cadastroClientes 
         Height          =   735
         Left            =   6480
         TabIndex        =   16
         Top             =   4560
         Width           =   2535
         Begin VB.CommandButton cmd_altera 
            Caption         =   "&Alterar"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmd_ok 
            Caption         =   "OK"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CbtSair 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   1320
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmb_cor 
         Height          =   315
         ItemData        =   "FRM_FichadeMotos.frx":00C5
         Left            =   3720
         List            =   "FRM_FichadeMotos.frx":00D8
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ComboBox cmb_marca 
         Height          =   315
         ItemData        =   "FRM_FichadeMotos.frx":0100
         Left            =   3720
         List            =   "FRM_FichadeMotos.frx":0125
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txt_modelo 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "Ano"
         Height          =   255
         Left            =   6120
         TabIndex        =   21
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "UF"
         Height          =   255
         Left            =   6120
         TabIndex        =   20
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "-"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label5 
         Caption         =   "Motoqueiro"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Cor"
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Modelo"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Marca"
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Placa"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FRM_FichadeMotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CbtSair_Click()
Unload Me
End Sub

Private Sub cmd_altera_Click()
Dim xplaca As String

xplaca = txt_letra.Text & txt_numero.Text

If MsgBox("Deseja Fazer a Alteração na Motocicleta: " & xplaca, vbYesNo, "MOTOS") = vbYes Then

    deb_coty.Up_motod cmb_motoqueiros.Text, cmb_marca.Text, txt_modelo.Text, mask_ano.Text, _
                    xplaca, txt_Cidade.Text, cmb_Uf.Text, cmb_cor.Text, xplaca

    MsgBox "Alteração Conluída com Sucesso", vbInformation, "MOTOS"

End If

If deb_coty.rsSel_Motos.State = 1 Then deb_coty.rsSel_Motos.Close
    deb_coty.Sel_Motos
    
    FRM_CadastrodeMotos.grd_motos.DataMember = "Sel_Motos"
    FRM_CadastrodeMotos.grd_motos.Refresh

Unload Me


End Sub

Private Sub cmd_ok_Click()
Dim xplaca As String

xplaca = txt_letra.Text & txt_numero.Text

If Trim$(Len(xplaca)) < 7 Then
    MsgBox "Corrija a Placa", vbInformation, "MOTOS"
    txt_letra.SelStart = 0
    txt_letra.SelLength = Len(txt_letra.Text)
    txt_letra.SetFocus
    Exit Sub

ElseIf Trim$(Len(txt_Cidade)) = 4 Then
    MsgBox "Corrija a Cidade", vbInformation, "MOTOS"
    txt_Cidade.SelStart = 0
    txt_Cidade.SelLength = Len(txt_letra.Text)
    txt_Cidade.SetFocus
    Exit Sub
    
ElseIf Trim$(Len(txt_modelo.Text)) = 0 Then
    MsgBox "Coloque o Modelo da Moto", vbInformation, "MOTOS"
    txt_modelo.SelStart = 0
    txt_modelo.SelLength = Len(txt_letra.Text)
    txt_modelo.SetFocus
    Exit Sub
    

ElseIf cmb_marca.ListIndex = -1 Then
    MsgBox "Selecione a Marca da Motocicleta", vbInformation, "MOTOS"
    cmb_marca.SetFocus
    Exit Sub
    
ElseIf cmb_motoqueiros.ListIndex = -1 Then
    MsgBox "Selecione o Nome do MotoBoy", vbInformation, "MOTOS"
    cmb_motoqueiros.SetFocus
    Exit Sub

ElseIf cmb_cor.ListIndex = -1 Then
    MsgBox "Selecione a Cor da Motocicleta", vbInformation, "MOTOS"
    cmb_cor.SetFocus
    Exit Sub

End If





deb_coty.In_Moto cmb_motoqueiros.Text, cmb_marca.Text, txt_modelo.Text, mask_ano.Text, _
                    xplaca, txt_Cidade.Text, cmb_Uf.Text, cmb_cor.Text
                    
MsgBox "Motocicleta: " & xplaca & Chr$(13) & "Cadastrada com seu sucesso"

limpa_tela (3)

If deb_coty.rsSel_Motos.State = 1 Then deb_coty.rsSel_Motos.Close
    deb_coty.Sel_Motos
    
    FRM_CadastrodeMotos.grd_motos.DataMember = "Sel_Motos"
    FRM_CadastrodeMotos.grd_motos.Refresh
    
Unload Me



End Sub

Private Sub Form_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))

End Sub

Private Sub txt_letra_Change()

If Len(txt_letra.Text) = 3 Then
    txt_numero.SetFocus
End If


End Sub

Private Sub txt_numero_Change()

If Len(txt_numero.Text) = 4 Then
    txt_Cidade.SetFocus
End If


End Sub
