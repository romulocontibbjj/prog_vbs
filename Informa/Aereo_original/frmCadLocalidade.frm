VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCadLocalidade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Localidades"
   ClientHeight    =   6915
   ClientLeft      =   3075
   ClientTop       =   795
   ClientWidth     =   5670
   ControlBox      =   0   'False
   Icon            =   "frmCadLocalidade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   5670
   Begin VB.CommandButton cmd_Alterar 
      Caption         =   "Alterar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1580
      TabIndex        =   4
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmd_Gravar 
      Caption         =   "Gravar"
      Enabled         =   0   'False
      Height          =   435
      Left            =   3040
      TabIndex        =   7
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmd_Canc 
      Caption         =   "Canc/Sair"
      Height          =   435
      Left            =   4500
      TabIndex        =   9
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmd_Nova 
      Caption         =   "Nova"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1035
   End
   Begin VB.Frame fra_Dados 
      Caption         =   "Dados de Localidade"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   660
      Width           =   5415
      Begin VB.TextBox TxtGeo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3420
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox TxtAeroporto 
         Height          =   285
         Left            =   180
         TabIndex        =   5
         Top             =   1200
         Width           =   3195
      End
      Begin VB.TextBox txtLocalidade 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox txtUF 
         Height          =   285
         Left            =   4620
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtSigla 
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Região Geog."
         Height          =   195
         Left            =   3420
         TabIndex        =   16
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Aeroporto"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "UF"
         Height          =   195
         Left            =   4620
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Localidade"
         Height          =   195
         Left            =   840
         TabIndex        =   12
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Sigla"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   435
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGridLocalidade 
      Height          =   4365
      Left            =   120
      TabIndex        =   14
      Top             =   2460
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7699
      _Version        =   393216
      AllowBigSelection=   0   'False
      FocusRect       =   2
      SelectionMode   =   1
   End
   Begin VB.Label lbl_status 
      Caption         =   "INICIO"
      Height          =   195
      Left            =   4020
      TabIndex        =   10
      Top             =   2340
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmCadLocalidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xSigla As String
Private Sub cmd_Alterar_Click()
'If de_informa.rsSel_CONFCidade.State = 1 Then de_informa.rsSel_CONFCidade.Close
'de_informa.Sel_CONFCidade "%" & Trim(txtLocalidade.Text) & "%"
'
'
'    If de_informa.rsSel_CONFCidade.RecordCount = 0 Then
'    MsgBox "Não foi encontrada nenhuma ocorrência similar a esta na Tabela de Cidades. Por favor, tente novamente", vbCritical, ""
'    Exit Sub
'    ElseIf de_informa.rsSel_CONFCidade.RecordCount > 2 Then
'        If MsgBox("Foram encontradas " & de_informa.rsSel_CONFCidade.RecordCount & " ocorrências de Cidades similares ao nome digitado. Pressione Ok para escolher qual seria a Cidade correta ou pressione Cancelar para tentar novamente...", vbOKCancel + vbExclamation, "") = vbOK Then
'        Set xForm = frmCadLocalidade
'        LeaveSub = False
'        FrmCadLocalidadeCONF.Show 1
'            If LeaveSub = True Then Exit Sub
'        Else
'        Exit Sub
'        End If
'    End If


cmd_nova.Enabled = False
cmd_Alterar.Enabled = False
FlexGridLocalidade.Enabled = False
fra_Dados.Enabled = True
cmd_Gravar.Enabled = True
lbl_status.Caption = "ALTERACAO"

xSigla = TxtSigla.Text

TxtSigla.BackColor = &HC0FFFF
txtLocalidade.BackColor = &HC0FFFF
TxtUF.BackColor = &HC0FFFF
TxtAeroporto.BackColor = &HC0FFFF
TxtGeo.BackColor = &HC0FFFF


TxtSigla.SetFocus
End Sub

Private Sub cmd_Canc_Click()

    If lbl_status.Caption = "INCLUSAO" Or lbl_status.Caption = "ALTERACAO" Then
    cmd_nova.Enabled = True
FlexGridLocalidade.Enabled = True
fra_Dados.Enabled = False
cmd_Gravar.Enabled = False


TxtSigla.BackColor = &H80000014
txtLocalidade.BackColor = &H80000014
TxtUF.BackColor = &H80000014
TxtAeroporto.BackColor = &H80000014
TxtGeo.BackColor = &H80000014


TxtSigla.Text = ""
txtLocalidade.Text = ""
TxtUF.Text = ""
TxtAeroporto.Text = ""
TxtGeo.Text = ""

Dim xCont As Integer
FlexGridLocalidade.Clear

    If de_informa.rsSel_CadLocalAir.State = 1 Then de_informa.rsSel_CadLocalAir.Close
    de_informa.Sel_CadLocalAir "%"
    
FlexGridLocalidade.Rows = de_informa.rsSel_CadLocalAir.RecordCount + 1
FlexGridLocalidade.FixedRows = 1
FlexGridLocalidade.FixedCols = 0
FlexGridLocalidade.Cols = 5

FlexGridLocalidade.ColWidth(0) = 500
FlexGridLocalidade.ColWidth(1) = 3700
FlexGridLocalidade.ColWidth(2) = 500
FlexGridLocalidade.ColWidth(3) = 3700
FlexGridLocalidade.ColWidth(4) = 1200
xCont = 0
FlexGridLocalidade.TextMatrix(xCont, 0) = "Sigla"
FlexGridLocalidade.TextMatrix(xCont, 1) = "Localidade"
FlexGridLocalidade.TextMatrix(xCont, 2) = "UF"
FlexGridLocalidade.TextMatrix(xCont, 3) = "Aeroporto"
FlexGridLocalidade.TextMatrix(xCont, 4) = "Região Geog."

xCont = 1
    Do Until de_informa.rsSel_CadLocalAir.EOF
    FlexGridLocalidade.TextMatrix(xCont, 0) = de_informa.rsSel_CadLocalAir.Fields("sigla")
    FlexGridLocalidade.TextMatrix(xCont, 1) = PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("localidade"))
    FlexGridLocalidade.TextMatrix(xCont, 2) = de_informa.rsSel_CadLocalAir.Fields("uf")
    FlexGridLocalidade.TextMatrix(xCont, 3) = PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("aeroporto"))
    FlexGridLocalidade.TextMatrix(xCont, 4) = PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("regiaogeo"))
    xCont = xCont + 1
    de_informa.rsSel_CadLocalAir.MoveNext
    Loop
    

lbl_status.Caption = "INICIO"
    
    Else
    Unload Me
    End If
End Sub



Private Sub cmd_Gravar_Click()
        
    If Len(Trim(txtLocalidade.Text)) = 0 Then
    MsgBox "Não foi informada corretamente a Localidade", vbCritical, ""
    Exit Sub
    ElseIf Len(Trim(TxtUF.Text)) = 0 Then
    MsgBox "Não foi informada corretamente a UF", vbCritical, ""
    Exit Sub
    End If
    
    If Len(Trim(TxtSigla.Text)) = 0 Then
    MsgBox "Não foi informada corretamente a Sigla do Aeroporto", vbCritical, ""
        If MsgBox("Deseja cadastrar esta cidade mesmo sem a sigla?", vbYesNo + vbExclamation, "") = vbNo Then
        Exit Sub
        End If
    End If
        
    
If de_informa.rsSel_CONFCidade.State = 1 Then de_informa.rsSel_CONFCidade.Close
de_informa.Sel_CONFCidade Trim(txtLocalidade.Text) & "%"
    
    
    If de_informa.rsSel_CONFCidade.RecordCount = 0 Then
        If MsgBox("Não foi encontrada nenhuma ocorrência similar a esta na Tabela de Cidades. Deseja cadastrar esta cidade mesmo assim?", vbYesNo + vbExclamation, "") = vbNo Then
        Exit Sub
        End If
    ElseIf de_informa.rsSel_CONFCidade.RecordCount > 2 Then
        If MsgBox("Foram encontradas " & de_informa.rsSel_CONFCidade.RecordCount & " ocorrências de Cidades similares ao nome digitado. Pressione Ok para escolher qual seria a Cidade correta ou pressione Cancelar para tentar novamente...", vbOKCancel + vbExclamation, "") = vbOK Then
        Set xForm = frmCadLocalidade
        LeaveSub = False
        FrmCadLocalidadeCONF.Show 1
            If LeaveSub = True Then Exit Sub
        Else
        Exit Sub
        End If
    End If

cmd_Gravar.Enabled = False
cmd_nova.Enabled = False
cmd_Alterar.Enabled = False

    If Len(Trim(TxtAeroporto)) = 0 Then TxtAeroporto.Text = txtLocalidade.Text
    
If TxtUF.Text = "RS" Or TxtUF.Text = "PR" Or TxtUF.Text = "SC" Then
TxtGeo.Text = "SUL"
ElseIf TxtUF.Text = "SP" Or TxtUF.Text = "MG" Or TxtUF.Text = "ES" Or TxtUF.Text = "RJ" Then
TxtGeo.Text = "SUDESTE"
ElseIf TxtUF.Text = "MT" Or TxtUF.Text = "MS" Or TxtUF.Text = "GO" Or TxtUF.Text = "DF" Then
TxtGeo.Text = "CENTRO-OESTE"
ElseIf TxtUF.Text = "AM" Or TxtUF.Text = "PA" Or TxtUF.Text = "AP" Or TxtUF.Text = "AC" Or TxtUF.Text = "RO" Or TxtUF.Text = "RR" Or TxtUF.Text = "TO" Then
TxtGeo.Text = "NORTE"
Else
TxtGeo.Text = "NORDESTE"
End If



    If lbl_status.Caption = "INCLUSAO" Then
        If de_informa.rsSel_CadLocalAir.State = 1 Then de_informa.rsSel_CadLocalAir.Close
        de_informa.Sel_CadLocalAir UCase(Trim(TxtSigla.Text))
        
        If de_informa.rsSel_CadLocalAir.RecordCount > 0 Then
        MsgBox "Já existe um Aeroporto Cadastrado com esta Sigla. O Cadastramento será cancelado.", vbExclamation, ""
        Else
        de_informa.Ins_CadLocalAir UCase(Trim(TxtSigla.Text)), UCase(Trim(txtLocalidade.Text)), UCase(Trim(TxtUF.Text)), TxtAeroporto.Text, TxtGeo.Text
        End If
    ElseIf lbl_status.Caption = "ALTERACAO" Then
        de_informa.Update_CadLocalAir UCase(Trim(TxtSigla.Text)), UCase(Trim(txtLocalidade.Text)), UCase(Trim(TxtUF.Text)), TxtAeroporto.Text, TxtGeo.Text, xSigla
    End If
    
cmd_nova.Enabled = True
FlexGridLocalidade.Enabled = True
fra_Dados.Enabled = False
cmd_Gravar.Enabled = False


TxtSigla.BackColor = &H80000014
txtLocalidade.BackColor = &H80000014
TxtUF.BackColor = &H80000014
TxtAeroporto.BackColor = &H80000014
TxtGeo.BackColor = &H80000014

TxtSigla.Text = ""
txtLocalidade.Text = ""
TxtUF.Text = ""
TxtAeroporto.Text = ""
TxtGeo.Text = ""


Dim xCont As Integer
FlexGridLocalidade.Clear

    If de_informa.rsSel_CadLocalAir.State = 1 Then de_informa.rsSel_CadLocalAir.Close
    de_informa.Sel_CadLocalAir "%"
    
FlexGridLocalidade.Rows = de_informa.rsSel_CadLocalAir.RecordCount + 1
FlexGridLocalidade.FixedRows = 1
FlexGridLocalidade.FixedCols = 0
FlexGridLocalidade.Cols = 5

FlexGridLocalidade.ColWidth(0) = 500
FlexGridLocalidade.ColWidth(1) = 3700
FlexGridLocalidade.ColWidth(2) = 500
FlexGridLocalidade.ColWidth(3) = 3700
FlexGridLocalidade.ColWidth(4) = 1200
xCont = 0
FlexGridLocalidade.TextMatrix(xCont, 0) = "Sigla"
FlexGridLocalidade.TextMatrix(xCont, 1) = "Localidade"
FlexGridLocalidade.TextMatrix(xCont, 2) = "UF"
FlexGridLocalidade.TextMatrix(xCont, 3) = "Aeroporto"
FlexGridLocalidade.TextMatrix(xCont, 4) = "Região Geog."

xCont = 1
    Do Until de_informa.rsSel_CadLocalAir.EOF
    FlexGridLocalidade.TextMatrix(xCont, 0) = de_informa.rsSel_CadLocalAir.Fields("sigla")
    FlexGridLocalidade.TextMatrix(xCont, 1) = PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("localidade"))
    FlexGridLocalidade.TextMatrix(xCont, 2) = de_informa.rsSel_CadLocalAir.Fields("uf")
    FlexGridLocalidade.TextMatrix(xCont, 3) = PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("aeroporto"))
    FlexGridLocalidade.TextMatrix(xCont, 4) = PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("regiaogeo"))
    xCont = xCont + 1
    de_informa.rsSel_CadLocalAir.MoveNext
    Loop
    

lbl_status.Caption = "INICIO"
End Sub

Private Sub cmd_Nova_Click()
cmd_nova.Enabled = False
cmd_Alterar.Enabled = False
FlexGridLocalidade.Enabled = False
FlexGridLocalidade.Enabled = False
fra_Dados.Enabled = True
cmd_Gravar.Enabled = True
lbl_status.Caption = "INCLUSAO"

TxtSigla.BackColor = &HC0FFFF
txtLocalidade.BackColor = &HC0FFFF
TxtUF.BackColor = &HC0FFFF
TxtAeroporto.BackColor = &HC0FFFF

TxtSigla.Text = ""
txtLocalidade.Text = ""
TxtUF.Text = ""
TxtAeroporto.Text = ""

TxtSigla.SetFocus


End Sub

Private Sub FlexGridLocalidade_Click()
cmd_Alterar.Enabled = True
With FlexGridLocalidade
    TxtSigla.Text = .TextMatrix(.Row, 0)
    txtLocalidade.Text = .TextMatrix(.Row, 1)
    TxtUF.Text = .TextMatrix(.Row, 2)
    TxtAeroporto.Text = .TextMatrix(.Row, 3)
    TxtGeo.Text = .TextMatrix(.Row, 4)
End With
End Sub

Private Sub Form_Load()
Dim xCont As Integer
FlexGridLocalidade.Clear

    If de_informa.rsSel_CadLocalAir.State = 1 Then de_informa.rsSel_CadLocalAir.Close
    de_informa.Sel_CadLocalAir "%"
    
FlexGridLocalidade.Rows = de_informa.rsSel_CadLocalAir.RecordCount + 1
FlexGridLocalidade.FixedRows = 1
FlexGridLocalidade.FixedCols = 0
FlexGridLocalidade.Cols = 5

FlexGridLocalidade.ColWidth(0) = 500
FlexGridLocalidade.ColWidth(1) = 3700
FlexGridLocalidade.ColWidth(2) = 500
FlexGridLocalidade.ColWidth(3) = 3700
FlexGridLocalidade.ColWidth(4) = 1200
xCont = 0
FlexGridLocalidade.TextMatrix(xCont, 0) = "Sigla"
FlexGridLocalidade.TextMatrix(xCont, 1) = "Localidade"
FlexGridLocalidade.TextMatrix(xCont, 2) = "UF"
FlexGridLocalidade.TextMatrix(xCont, 3) = "Aeroporto"
FlexGridLocalidade.TextMatrix(xCont, 4) = "Região Geog."

xCont = 1
    Do Until de_informa.rsSel_CadLocalAir.EOF
    FlexGridLocalidade.TextMatrix(xCont, 0) = de_informa.rsSel_CadLocalAir.Fields("sigla")
    FlexGridLocalidade.TextMatrix(xCont, 1) = PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("localidade"))
    FlexGridLocalidade.TextMatrix(xCont, 2) = de_informa.rsSel_CadLocalAir.Fields("uf")
    FlexGridLocalidade.TextMatrix(xCont, 3) = PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("aeroporto"))
    FlexGridLocalidade.TextMatrix(xCont, 4) = PriMaiuscula(de_informa.rsSel_CadLocalAir.Fields("regiaogeo"))
    xCont = xCont + 1
    de_informa.rsSel_CadLocalAir.MoveNext
    Loop
    
End Sub

Private Sub Teste_Change()

Dim Abc As String
Dim Cont As Integer
Dim Flag As Boolean
Dim TextoAux As String

Abc = "ABCEDFGHIJKLMNOPQRSTUVYXWZ"
TextoAux = ""

    If Len(TxtSigla.Text) > 0 Then
        
        Flag = False
        
        For Cont = 1 To Len(Abc)
            If UCase(Mid(TxtSigla.Text, TxtSigla.SelStart, 1)) = Mid(Abc, Cont, 1) Then
            Flag = True
            End If
        Next
    
        If Flag = False Then
            For Cont = 1 To Len(TxtSigla.Text)
                If Cont <> TxtSigla.SelStart Then
                TextoAux = Mid(TxtSigla, Cont, 1)
                End If
            Next
        Else
        TextoAux = TxtSigla.Text
        End If
        
        TxtSigla.Text = TextoAux
    End If
        
End Sub

Private Sub ListCidade_Click()
ListSIGLA.Selected(ListCidade.ListIndex) = True
ListUF.Selected(ListCidade.ListIndex) = True
cmd_Alterar.Enabled = True

TxtSigla.Text = ListSIGLA.Text
txtLocalidade.Text = ListCidade.Text
TxtUF.Text = ListUF.Text
End Sub

Private Sub ListCidade_Scroll()
ListSIGLA.TopIndex = ListCidade.TopIndex
ListUF.TopIndex = ListUF.TopIndex
End Sub

Private Sub ListSIGLA_Click()
ListCidade.Selected(ListSIGLA.ListIndex) = True
ListUF.Selected(ListSIGLA.ListIndex) = True
cmd_Alterar.Enabled = True

TxtSigla.Text = ListSIGLA.Text
txtLocalidade.Text = ListCidade.Text
TxtUF.Text = ListUF.Text
End Sub

Private Sub ListSIGLA_Scroll()
ListCidade.TopIndex = ListSIGLA.TopIndex
ListUF.TopIndex = ListSIGLA.TopIndex
End Sub

Private Sub ListUF_Click()
ListSIGLA.Selected(ListUF.ListIndex) = True
ListCidade.Selected(ListUF.ListIndex) = True
cmd_Alterar.Enabled = True
TxtSigla.Text = ListSIGLA.Text
txtLocalidade.Text = ListCidade.Text
TxtUF.Text = ListUF.Text
End Sub

Private Sub ListUF_Scroll()
ListCidade.TopIndex = ListUF.TopIndex
ListSIGLA.TopIndex = ListUF.TopIndex
End Sub

Private Sub TxtAeroporto_Change()
TxtAeroporto.Text = UCase(TxtAeroporto.Text)
TxtAeroporto.SelStart = Len(TxtAeroporto.Text)
End Sub

Private Sub TxtAeroporto_GotFocus()
TxtAeroporto.SelStart = 0
TxtAeroporto.SelLength = 10
End Sub

Private Sub TxtAeroporto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub TxtAeroporto_LostFocus()
TxtAeroporto.Text = UCase(TxtAeroporto.Text)
End Sub

Private Sub txtLocalidade_Change()
txtLocalidade.Text = UCase(txtLocalidade.Text)
txtLocalidade.SelStart = Len(txtLocalidade.Text)
End Sub

Private Sub Txtlocalidade_GotFocus()
txtLocalidade.SelStart = 0
txtLocalidade.SelLength = 100
End Sub


Private Sub Txtlocalidade_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub


Private Sub txtLocalidade_LostFocus()
txtLocalidade.Text = SemAcento(UCase(txtLocalidade.Text))
End Sub

Private Sub txtSigla_Change()
TxtSigla.Text = UCase(TxtSigla.Text)
TxtSigla.SelStart = Len(TxtSigla.Text)
End Sub

Private Sub TxtSigla_GotFocus()
TxtSigla.SelStart = 0
TxtSigla.SelLength = 10
End Sub


Private Sub TxtSigla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub


Private Sub TxtSigla_LostFocus()
TxtSigla.Text = UCase(TxtSigla.Text)
End Sub

Private Sub txtUF_Change()
TxtUF.Text = UCase(TxtUF.Text)
TxtUF.SelStart = Len(TxtUF.Text)
End Sub

Private Sub txtUF_GotFocus()
TxtUF.SelStart = 0
TxtUF.SelLength = 10
End Sub


Private Sub txtUF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub


Private Sub txtUF_LostFocus()
If Len(Trim(TxtUF.Text)) > 0 Then
    If TxtUF = "AM" Or TxtUF = "AP" Or TxtUF = "PA" Or TxtUF = "RO" Or TxtUF = "RR" Or TxtUF = "TO" Then
    TxtGeo.Text = "Norte"
    ElseIf TxtUF = "AL" Or TxtUF = "BA" Or TxtUF = "CE" Or TxtUF = "MA" Or TxtUF = "PB" Or TxtUF = "PE" Or TxtUF = "PI" Or TxtUF = "RN" Or TxtUF = "SE" Then
    TxtGeo.Text = "Nordeste"
    ElseIf TxtUF = "DF" Or TxtUF = "GO" Or TxtUF = "MS" Or TxtUF = "MT" Then
    TxtGeo.Text = "Centro-Oeste"
    ElseIf TxtUF = "ES" Or TxtUF = "MG" Or TxtUF = "RJ" Or TxtUF = "SP" Then
    TxtGeo.Text = "Sudeste"
    ElseIf TxtUF = "PR" Or TxtUF = "RS" Or TxtUF = "SC" Then
    TxtGeo.Text = "Sul"
    End If
End If
End Sub



