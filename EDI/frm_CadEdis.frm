VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_CadEdis 
   Caption         =   "CADASTRO DE EDI´S"
   ClientHeight    =   9105
   ClientLeft      =   2025
   ClientTop       =   1395
   ClientWidth     =   10575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   10575
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10515
      Begin VB.Frame fra_periodo 
         Height          =   825
         Left            =   2925
         TabIndex        =   40
         Top             =   600
         Width           =   3540
         Begin VB.CheckBox chk_Canc 
            Caption         =   "Cancelados"
            Height          =   255
            Left            =   2040
            TabIndex        =   50
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chk_entrega 
            BackColor       =   &H8000000A&
            Caption         =   "Somente Entreg"
            Height          =   240
            Left            =   2025
            TabIndex        =   44
            Top             =   150
            Width           =   1440
         End
         Begin VB.TextBox txt_periodo 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   975
            TabIndex        =   42
            Text            =   "0"
            Top             =   150
            Width           =   465
         End
         Begin VB.Label Label13 
            Caption         =   "CANC:"
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
            Left            =   1440
            TabIndex        =   49
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "DIAS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1500
            TabIndex        =   43
            Top             =   150
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "PERÍODO:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   75
            TabIndex        =   41
            Top             =   150
            Width           =   915
         End
      End
      Begin VB.CommandButton cmd_Salvar 
         Caption         =   "&Sair"
         Height          =   390
         Left            =   8640
         TabIndex        =   39
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton cmd_cgc 
         Caption         =   "&?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2925
         TabIndex        =   2
         Top             =   300
         Width           =   315
      End
      Begin VB.Frame Frame4 
         Caption         =   "ARQUIVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   120
         TabIndex        =   32
         Top             =   6240
         Width           =   10215
         Begin VB.CheckBox chk_DDMMAA 
            Caption         =   "DDMMAA"
            Height          =   255
            Left            =   6360
            TabIndex        =   48
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txt_instrucao 
            Height          =   285
            Left            =   2640
            TabIndex        =   46
            Top             =   1680
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CheckBox chk_ddmm 
            Caption         =   "DDMM"
            Height          =   240
            Left            =   6360
            TabIndex        =   38
            Top             =   360
            Width           =   825
         End
         Begin VB.TextBox txt_Arquivo 
            Height          =   285
            Left            =   2625
            TabIndex        =   37
            Top             =   1125
            Width           =   2235
         End
         Begin VB.TextBox txt_Salvar 
            Height          =   285
            Left            =   2625
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   525
            Width           =   3015
         End
         Begin VB.DirListBox Dir1 
            Height          =   2115
            Left            =   150
            TabIndex        =   33
            Top             =   300
            Width           =   2340
         End
         Begin VB.Label lab_somaletras 
            Caption         =   "MAX:100 (0)"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4920
            TabIndex        =   47
            Top             =   1680
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Instrução"
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
            Left            =   2640
            TabIndex        =   45
            Top             =   1440
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Nome Arquivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2625
            TabIndex        =   36
            Top             =   900
            Width           =   1290
         End
         Begin VB.Label Label9 
            Caption         =   "Salvar Em:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2625
            TabIndex        =   34
            Top             =   300
            Width           =   1140
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ENVIOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   7320
         TabIndex        =   19
         Top             =   120
         Width           =   3090
         Begin MSMask.MaskEdBox mask_hora 
            Height          =   315
            Left            =   2175
            TabIndex        =   31
            Top             =   450
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "99:99"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txt_Dia 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   750
            TabIndex        =   29
            Top             =   1800
            Width           =   465
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Dias Úteis"
            Height          =   240
            Left            =   150
            TabIndex        =   28
            Top             =   2160
            Width           =   2340
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Dia:"
            Height          =   240
            Left            =   150
            TabIndex        =   27
            Top             =   1875
            Width           =   615
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Sábado"
            Height          =   240
            Left            =   150
            TabIndex        =   26
            Top             =   1575
            Width           =   1290
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Sexta"
            Height          =   240
            Left            =   150
            TabIndex        =   25
            Top             =   1350
            Width           =   990
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Quinta"
            Height          =   240
            Left            =   150
            TabIndex        =   24
            Top             =   1125
            Width           =   915
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Quarta"
            Height          =   240
            Left            =   150
            TabIndex        =   23
            Top             =   900
            Width           =   1290
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Terça"
            Height          =   195
            Left            =   150
            TabIndex        =   22
            Top             =   675
            Width           =   1515
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Segunda"
            Height          =   240
            Left            =   150
            TabIndex        =   21
            Top             =   450
            Width           =   1965
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Domingo"
            Height          =   240
            Left            =   150
            TabIndex        =   20
            Top             =   225
            Width           =   1365
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Horário"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1950
            TabIndex        =   30
            Top             =   225
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "EMAILS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2340
         Left            =   7320
         TabIndex        =   11
         Top             =   3840
         Width           =   3090
         Begin VB.ListBox list_Email 
            Height          =   2010
            Left            =   75
            TabIndex        =   12
            Top             =   225
            Width           =   2940
         End
      End
      Begin VB.CommandButton cmd_recebe 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6600
         TabIndex        =   5
         Top             =   4800
         Width           =   540
      End
      Begin VB.CommandButton cmb_envia 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6600
         TabIndex        =   4
         Top             =   4320
         Width           =   540
      End
      Begin VB.ListBox txt_email 
         Height          =   1815
         Left            =   1200
         TabIndex        =   10
         Top             =   4200
         Width           =   5295
      End
      Begin VB.CommandButton cmd_gravar 
         Caption         =   "&Gravar"
         Height          =   390
         Left            =   8640
         TabIndex        =   9
         Top             =   2880
         Width           =   1725
      End
      Begin VB.TextBox txt_Mensagem 
         Height          =   1920
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2280
         Width           =   5295
      End
      Begin VB.TextBox txt_assunto 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   1920
         Width           =   5295
      End
      Begin VB.ComboBox cmb_edi 
         Height          =   315
         ItemData        =   "frm_CadEdis.frx":0000
         Left            =   1275
         List            =   "frm_CadEdis.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   675
         Width           =   1515
      End
      Begin VB.TextBox txt_cliente 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   1440
         Width           =   5295
      End
      Begin VB.TextBox txt_cgc 
         Height          =   285
         Left            =   1275
         MaxLength       =   8
         TabIndex        =   1
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "MENSAGEM:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   18
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Label Label6 
         Caption         =   "ASSUNTO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   17
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "E-MAIL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   16
         Top             =   4200
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "EDI:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   15
         Top             =   675
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "CLIENTE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   14
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "C.G.C.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   13
         Top             =   300
         Width           =   540
      End
   End
End
Attribute VB_Name = "frm_CadEdis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xid As Integer

Private Sub Check8_Click()
Dim vobj As Control
txt_Dia.Enabled = True
For Each vobj In frm_CadEdis

    If TypeOf vobj Is CheckBox Then
        
        If Check8.Value = 1 Then
        
            vobj.Enabled = False

        Else
        
            vobj.Enabled = True
            
        End If
        
        
    End If
    

Next
chk_ddmm.Enabled = True
Check8.Enabled = True


End Sub

Private Sub chk_Canc_Click()
If chk_Canc.Value = 1 Then

    chk_entrega.Enabled = False
Else
    chk_entrega.Enabled = True
End If


End Sub

Private Sub chk_entrega_Click()

'If chk_entrega.Value = 1 Then

 '   chk_Canc.Enabled = False
'Else
'    chk_Canc.Enabled = True
'End If



End Sub

Private Sub cmb_edi_Click()

If cmb_edi.Text = "OCOREN" Then
    
    txt_periodo.Text = "0"
    txt_periodo.Enabled = False

ElseIf cmb_edi.Text = "DOCCOB" Then
    
    txt_periodo.Text = "0"
    txt_periodo.Enabled = False

ElseIf cmb_edi.Text = "CONEMB" Or cmb_edi.Text = "CORREIOS" Then
    
    txt_periodo.Enabled = True
    
    If txt_periodo.Text = "" Then
        
        MsgBox ("Informe o período ..."), vbInformation + vbOKOnly
        cmb_edi.SetFocus
    
    End If
    
End If

End Sub

Private Sub cmb_envia_Click()

If list_Email.ListIndex > -1 Then
    
        If deb_edi.rsSel_Email.State = 1 Then deb_edi.rsSel_Email.Close
        deb_edi.Sel_Email Mid(list_Email.List(list_Email.ListIndex), 1, Len(list_Email.List(list_Email.ListIndex)) - 1)
        

    txt_Email.AddItem deb_edi.rsSel_Email.Fields("email") & ";"
    list_Email.RemoveItem list_Email.ListIndex
End If

End Sub

Private Sub cmd_cgc_Click()
frm_cgc.Show 1


End Sub

Private Sub cmd_Gravar_Click()
Dim xcgc As String
Dim xedi As String
Dim xcliente As String
Dim xemail As String
Dim xasssunto As String
Dim xmensagem As String
Dim xSalvar As String
Dim xnomearq As String
Dim xddmm As String
Dim xDia As Integer
Dim xhorario As String
Dim xsemana As String
Dim xentrega As Integer
Dim X As Integer

'If txt_Email.ListCount = 0 Then
    'MsgBox "Favor Selecione No minimo um Email para Envio do EDI", vbCritical, "EMAILS"
'    list_Email.Selected(-1) = True
    'Exit Sub
    
'ElseIf Len(Trim$(txt_assunto.Text)) = 0 Then
 '   MsgBox "Digite o Assunto do Email", vbCritical, "EMAILS"
  '  txt_assunto.SetFocus
   ' Exit Sub
    
If Len(Trim$(txt_Salvar.Text)) = 0 Then
    MsgBox "Favor selecionar o Diretório no Qual será salvo o EDI."
    Dir1.SetFocus
    Exit Sub

End If

If Check8.Value = 1 Then
    xDia = txt_Dia.Text
    
ElseIf Check9.Value = 1 Then
    
    xDia = 0
    xsemana = "0111110"

Else
    xDia = 0
    
    xsemana = Check1.Value & Check2.Value & Check3.Value & Check4.Value & Check5.Value & Check6.Value & Check7.Value
 
    
End If

For X = 0 To txt_Email.ListCount - 1

    xemail = xemail & txt_Email.List(X)

Next

xcgc = txt_cgc.Text
xedi = cmb_edi.Text
xcliente = txt_Cliente.Text
xassunto = txt_assunto.Text
xmensagem = txt_Mensagem.Text
xSalvar = txt_Salvar.Text

If chk_ddmm.Value = 1 Then
    xddmm = 1
Else
    xddmm = 0
End If

xnomearq = txt_Arquivo.Text
xhorario = mask_hora.Text


If cmb_edi.Text = "BONAGURA" Then

    xcgc = 0
    xcliente = "BONAGURA"
    
End If








If xid = 0 Then

    deb_edi.In_CadEnvios xcgc, xedi, xcliente, xemail, xassunto, xmensagem, xSalvar, xnomearq, xddmm, xDia, xhorario, xsemana, Int(txt_periodo.Text), chk_entrega.Value, chk_Canc.Value
    MsgBox "Edi Cadastrado com Sucesso", vbInformation, "EMAILS"
Else

If txt_periodo.Text = "" Then
    txt_periodo.Text = 0
End If


    deb_edi.Up_CadEdis xcgc, xedi, xcliente, xemail, xassunto, xmensagem, xSalvar, xnomearq, xddmm, xDia, xhorario, xsemana, Int(txt_periodo.Text), chk_entrega.Value, chk_Canc.Value, xid
    xid = 0
    
    If deb_edi.rsSel_CadEdis.State = 1 Then deb_edi.rsSel_CadEdis.Close
        deb_edi.Sel_CadEdis
    
        frm_Cadastrados.grd_Edi.DataMember = "Sel_CadEdis"
        frm_Cadastrados.grd_Edi.Refresh
        
    MsgBox "Edi Alterado com Sucesso", vbInformation, "EMAILS"
End If




limpa_tela (1)
mask_hora.Mask = ""
mask_hora.Text = Empty
mask_hora.Mask = "99:99"

txt_Email.Clear

list_Email.Clear

If deb_edi.rsSel_Emails.State = 1 Then deb_edi.rsSel_Emails.Close
    deb_edi.Sel_Emails

    If deb_edi.rsSel_Emails.RecordCount > 0 Then
        deb_edi.rsSel_Emails.MoveFirst
        Do Until deb_edi.rsSel_Emails.EOF
        
        list_Email.AddItem deb_edi.rsSel_Emails.Fields("nome") & ";"
         
        deb_edi.rsSel_Emails.MoveNext
        
        Loop
    End If


End Sub

Private Sub cmd_recebe_Click()
If txt_Email.ListIndex > -1 Then

    list_Email.AddItem txt_Email.List(txt_Email.ListIndex)
    txt_Email.RemoveItem txt_Email.ListIndex
End If
End Sub

Private Sub cmd_Salvar_Click()


If deb_edi.rsSel_CadEdis.State = 1 Then deb_edi.rsSel_CadEdis.Close
        deb_edi.Sel_CadEdis
        
        frm_Cadastrados.grd_Edi.DataMember = "Sel_CadEdis"
        frm_Cadastrados.grd_Edi.Refresh

        frm_Cadastrados.Show

Unload Me

End Sub

Private Sub Dir1_Change()
txt_Salvar.Text = Dir1.Path

End Sub

Private Sub Form_Load()
Dir1.Path = "c:\informa"

If deb_edi.rsSel_Emails.State = 1 Then deb_edi.rsSel_Emails.Close
    deb_edi.Sel_Emails

    If deb_edi.rsSel_Emails.RecordCount > 0 Then
        deb_edi.rsSel_Emails.MoveFirst
        Do Until deb_edi.rsSel_Emails.EOF
        
        list_Email.AddItem deb_edi.rsSel_Emails.Fields("nome") & ";"
         
        deb_edi.rsSel_Emails.MoveNext
        
        Loop
    End If
    

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift _
      As Integer)
  'If KeyCode = vbKeyReturn Then
    'SendKeys "{Tab}"
    'KeyCode = 0
  'End If
End Sub

Private Sub txt_assunto_LostFocus()

txt_assunto.Text = UCase(txt_assunto.Text)

End Sub

Private Sub txt_instrucao_Change()
lab_somaletras.Caption = "MAX:100 (" & Len(txt_instrucao.Text) & ")"
End Sub

Private Sub txt_periodo_lostFocus()

If IsNumeric(txt_periodo.Text) = False Then
    
        MsgBox ("Valor de período inválido ..."), vbInformation + vbOKOnly
        txt_periodo.SetFocus
    
End If

End Sub
