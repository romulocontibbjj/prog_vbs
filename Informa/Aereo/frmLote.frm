VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Emissão de Etiquetas de Lote"
   ClientHeight    =   4650
   ClientLeft      =   5205
   ClientTop       =   2610
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraGrid 
      Caption         =   "Range dos Lotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   1980
      Width           =   4815
      Begin MSDataGridLib.DataGrid GridFormularios 
         Bindings        =   "frmLote.frx":000C
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3836
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
         DataMember      =   "SelLotes"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Filial"
            Caption         =   "Filial"
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
            DataField       =   "Sigla"
            Caption         =   "Sigla"
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
            DataField       =   "NumeroMin"
            Caption         =   "NumeroMin"
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
            DataField       =   "NumeroMax"
            Caption         =   "NumeroMax"
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
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   540,284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1005,165
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraInserir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.TextBox TxtQte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3660
         MaxLength       =   4
         TabIndex        =   2
         Top             =   840
         Width           =   1035
      End
      Begin VB.TextBox TxtBuscaFilial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   1
         Top             =   540
         Width           =   495
      End
      Begin VB.CommandButton cmdConfirma 
         Caption         =   "Confirmar"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   1260
         Width           =   2295
      End
      Begin VB.CommandButton cmdDesisteIns 
         Caption         =   "Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1260
         Width           =   2295
      End
      Begin VB.TextBox TxtSigla 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade de Lotes que deseja imprimir:"
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   885
         Width           =   2910
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   585
         TabIndex        =   12
         Top             =   585
         Width           =   300
      End
      Begin VB.Label TxtNomeFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Top             =   540
         Width           =   2655
      End
      Begin VB.Label TxtFilial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         TabIndex        =   10
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sigla:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   390
      End
      Begin VB.Label lblFantasiaIns 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1500
         TabIndex        =   7
         Top             =   240
         Width           =   3195
      End
   End
End
Attribute VB_Name = "frmLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBusca_Click()
    If de_informa.rsSel_NumFormularioItem.State = 1 Then de_informa.rsSel_NumFormularioItem.Close
    de_informa.Sel_NumFormularioItem Trim$(txtCodCia), Val(txtNumAWB), TxtFilial2.Caption
    LblStatus = ""
    lblDataStatus = ""
    TxtMotivo.Text = ""
    If de_informa.rsSel_NumFormularioItem.RecordCount > 0 Then
        If de_informa.rsSel_NumFormularioItem.Fields("tem_ocorr") = "C" Then
        LblStatus = "Cancelado"
        TxtMotivo.Enabled = False
        ElseIf de_informa.rsSel_NumFormularioItem.Fields("tem_ocorr") = "D" Then
        LblStatus = "Disponível"
        TxtMotivo.Enabled = True
        ElseIf de_informa.rsSel_NumFormularioItem.Fields("tem_ocorr") = "N" Then
        LblStatus = "Emitido"
        TxtMotivo.Enabled = True
        ElseIf de_informa.rsSel_NumFormularioItem.Fields("tem_ocorr") = "2" Then
        LblStatus = "Em Ocorrência"
        TxtMotivo.Enabled = False
        End If
        
        If Not IsNull(de_informa.rsSel_NumFormularioItem.Fields("canc_obs")) Then
        TxtMotivo.Text = de_informa.rsSel_NumFormularioItem.Fields("canc_obs")
        End If
        
            If de_informa.rsSel_NumFormularioItem.Fields("tem_ocorr") = "C" Then
            lblDataStatus = de_informa.rsSel_NumFormularioItem.Fields("canc_data")
            Else
                If Not IsNull(de_informa.rsSel_NumFormularioItem.Fields("datastatus")) Then
                    lblDataStatus = de_informa.rsSel_NumFormularioItem.Fields("datastatus")
                Else
                    lblDataStatus = ""
                End If
            End If
    Else
        MsgBox "Número de AWB para esta Cia. Não Encontrado!", vbExclamation, ""
    End If
    If Len(Trim$(LblStatus)) > 0 Then
        cmdConfirmaCanc.Enabled = True
    Else
        cmdConfirmaCanc.Enabled = False
    End If
    txtNumAWB.SetFocus
End Sub

Private Sub cmdCancelarForm_Click()
    If Mid(stringdiretos, 35, 1) = "0" Then
    MsgBox "Acesso Negado! Contate o administrador do sitema.", vbCritical, "ACESSO NEGADO!"
    Exit Sub
    End If


    fraConsCanc.Caption = "CANCELAR Formulário AWB"
    fraConsCanc.Enabled = True
    fraInserir.Enabled = False
    cmdCancelarForm.Enabled = False
    cmdInserirForm.Enabled = False
    cmdSair.Enabled = False
    If Len(Trim$(LblStatus)) > 0 Then
        cmdConfirmaCanc.Enabled = True
    Else
        cmdConfirmaCanc.Enabled = False
    End If
    txtCodCia.SetFocus
End Sub

Private Sub cmdConfirmaCanc_Click()
    If Val(txtNumAWB.Text) = 0 Then
    MsgBox "Você não informou o AWB.", vbCritical, ""
    Exit Sub
    End If
    
    If Len(Trim(TxtMotivo.Text)) = 0 Then
    MsgBox "Você não informou o motivo.", vbCritical, ""
    Exit Sub
    End If

    If Mid(LblStatus, 1, 1) = "C" Then
        MsgBox "ERRO! Este Formulário Já Está Cancelado!", vbCritical, ""
        txtNumAWB.SetFocus
        Exit Sub
    Else
        If LblStatus = "Em Ocorrência" Then
        MsgBox "Não é possível Cancelar este AWB pois ele está em processo de Ocorrência!"
        Exit Sub
        End If
    de_informa.Alt_FormularioStatuscanc "C", "TESTE", UCase(Trim(TxtMotivo.Text)), txtCodCia, txtNumAWB, TxtFilial2.Caption
    MsgBox "OK ! Formulário Cancelado.", vbInformation, ""
    cmdBusca_Click
    End If
txtCodCia.Text = ""
lblFantasia.Caption = ""
txtNumAWB.Text = ""
LblStatus.Caption = ""
lblDataStatus.Caption = ""
TxtBuscaFilial2.Text = ""
TxtFilial2.Caption = ""
TxtNomeFilial2.Caption = ""
TxtMotivo.Text = ""
TxtMotivo.Enabled = False
End Sub

Private Sub cmdDesisteCanc_Click()
fraConsCanc.Caption = "Consultar Formulário AWB"
fraConsCanc.Enabled = False
cmdCancelarForm.Enabled = True
cmdInserirForm.Enabled = True
cmdSair.Enabled = True

txtCodCia.Text = ""
lblFantasia.Caption = ""
txtNumAWB.Text = ""
LblStatus.Caption = ""
lblDataStatus.Caption = ""
TxtBuscaFilial2.Text = ""
TxtFilial2.Caption = ""
TxtNomeFilial2.Caption = ""

End Sub

Private Sub cmdConfirma_Click()
'CONFIGURACAO DE IMPRESSORAS - Inicio
Dim SETIMPLinha As String
Dim SETIMPImpressoraAtual As Printer
Dim SETIMPImpressoraPadrao As String
Dim SETIMPAchouIMP As Boolean

    If Dir("c:\printer.cfg") = "" Then
    MsgBox "Você não possui o arquivo de configuração de impressoras. Antes de continuar, é imprescindível que você configure as configure.", vbExclamation, "IMPRESSORAS"
    frmControleImpressoras.Show 1
    End If

AchouIMP = False
    
    If Dir("c:\printer.cfg") <> "" Then
        Open "c:\printer.cfg" For Input As #1
        Do Until EOF(1)
            Line Input #1, SETIMPxLinha
            If Mid(SETIMPxLinha, 1, 3) = "ETL" Then
            SETIMPImpressoraPadrao = Mid(SETIMPxLinha, 5)
            SETIMPAchouIMP = True
            Exit Do
            End If
        Loop
        Close #1
    End If
    
    If SETIMPAchouIMP = False Then
    MsgBox "Não existe impressora configurada para esta operação. Corrija este problema indo ao menu Configurações e depois em Impressoras e configure em qual impressora os AWBs deverão ser impressos.", vbCritical, "ERRO!"
    Exit Sub
    End If
    
    For Each SETIMPImpressoraAtual In Printers
        If SETIMPImpressoraAtual.DeviceName = SETIMPImpressoraPadrao Then
            Set Printer = SETIMPImpressoraAtual
            DoEvents
            Exit For
        End If
    Next
'CONFIGURACAO DE IMPRESSORAS - Fim

Dim xLote As String
Dim xCont As Integer
Dim xNumero As Long
Dim xEtiqueta As Long
Dim xPagina As Long
Dim MargemTOP As Double
Dim MargemLeft As Double
Dim Recuo As Long
Dim xTexto As String


    If ((Val(TxtQte.Text) / 5) - Int(Val(TxtQte.Text) / 5)) > 0 Then
    xPagina = Int(Val(TxtQte.Text) / 5) + 1
    Else
    xPagina = Int(Val(TxtQte.Text) / 5)
    End If


de_informa.cn_informa.BeginTrans

xEtiqueta = 0
Recuo = 13

    For xCont = 1 To Val(TxtQte.Text)
    xEtiqueta = xEtiqueta + 1
        If xEtiqueta > 5 Then
        Printer.NewPage
        xEtiqueta = 1
        End If
        
    If xEtiqueta = 1 Then
    MargemTOP = 15 + Recuo
    ElseIf xEtiqueta = 2 Then
    MargemTOP = 65 + Recuo
    ElseIf xEtiqueta = 3 Then
    MargemTOP = 115 + Recuo
    ElseIf xEtiqueta = 4 Then
    MargemTOP = 165 + Recuo
    ElseIf xEtiqueta = 5 Then
    MargemTOP = 215 + Recuo
    End If
        
    If de_informa.rsCapLote.State = 1 Then de_informa.rsCapLote.Close
    de_informa.CapLote TxtSigla.Text, TxtBuscaFilial.Text
    
        If de_informa.rsCapLote.RecordCount = 0 Then
        xNumero = 1
        Else
        xNumero = de_informa.rsCapLote.Fields("numero") + 1
        End If
    xLote = String(2 - Len(Trim(Str(Val(TxtFilial.Caption)))), "0") & Trim(Str(Val(TxtFilial.Caption))) & UCase(Trim(TxtSigla.Text)) & String(10 - Len(Trim(Str(xNumero))), "0") & Trim(Str(xNumero))
    de_informa.InsereLote xLote, TxtFilial.Caption, TxtSigla.Text, Trim(Str(xNumero))
    
    MargemLeft = 10
    Call EnviaIMP("ARIAL", 25, MargemLeft, MargemTOP, xLote)
    xTexto = "Lotes Internos Intec Cargo - Via Operacional"
    Call EnviaIMP("ARIAL", 10, MargemLeft, MargemTOP + 10, xTexto)
    
    MargemLeft = 110
    Call EnviaIMP("ARIAL", "25", MargemLeft, MargemTOP, xLote)
    xTexto = "Lotes Internos Intec Cargo - Via Emissão"
    Call EnviaIMP("ARIAL", 10, MargemLeft, MargemTOP + 10, xTexto)
    Next
    
    If xEtiqueta < 5 Then
    xEtiqueta = xEtiqueta + 1
        For xCont = xEtiqueta To 5
            If xCont = 1 Then
            MargemTOP = 15 + Recuo
            ElseIf xCont = 2 Then
            MargemTOP = 65 + Recuo
            ElseIf xCont = 3 Then
            MargemTOP = 115 + Recuo
            ElseIf xCont = 4 Then
            MargemTOP = 165 + Recuo
            ElseIf xCont = 5 Then
            MargemTOP = 215 + Recuo
            End If
        MargemLeft = 20
        Call EnviaIMP("ARIAL", "30", MargemLeft, MargemTOP, "EM BRANCO")
        
        MargemLeft = 120
        Call EnviaIMP("ARIAL", "30", MargemLeft, MargemTOP, "EM BRANCO")
        Next
    End If
    
    Printer.EndDoc
de_informa.cn_informa.CommitTrans
Call LimpaTela(Me)
GridFormularios.Refresh
End Sub

Private Sub cmdDesisteIns_Click()
Unload Me
End Sub

Private Sub cmdInserirForm_Click()
    fraConsCanc.Enabled = False
    fraInserir.Enabled = True
    cmdCancelarForm.Enabled = False
    cmdInserirForm.Enabled = False
    cmdSair.Enabled = False
    TxtSigla.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TxtAWB2_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If

End Sub

Private Sub TxtBuscaFilial_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtBuscaFilial_LostFocus()
If Len(Trim(TxtBuscaFilial.Text)) > 0 Then
    TxtBuscaFilial.Text = Trim(String(2 - Len(Trim(Str(Val(TxtBuscaFilial.Text)))), "0")) & Trim(Str(Val(TxtBuscaFilial.Text)))
    If de_informa.rsSelFiliais.State = 1 Then de_informa.rsSelFiliais.Close
    de_informa.SelFiliais TxtBuscaFilial.Text
    If de_informa.rsSelFiliais.RecordCount > 0 Then
    If IsNull(de_informa.rsSelFiliais.Fields("filial")) = False Then TxtFilial.Caption = de_informa.rsSelFiliais.Fields("filial")
    If IsNull(de_informa.rsSelFiliais.Fields("nomefilial")) = False Then TxtNomeFilial.Caption = PriMaiuscula(de_informa.rsSelFiliais.Fields("nomefilial"))
    DoEvents
    End If
End If
End Sub

Private Sub TxtBuscaFilial2_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        Call TxtBuscaFilial2_LostFocus
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtBuscaFilial2_LostFocus()
If Len(Trim(TxtBuscaFilial2.Text)) > 0 Then
    TxtBuscaFilial2.Text = Trim(String(2 - Len(Trim(Str(Val(TxtBuscaFilial2.Text)))), "0")) & Trim(Str(Val(TxtBuscaFilial2.Text)))
    If de_informa.rsSelFiliais.State = 1 Then de_informa.rsSelFiliais.Close
    de_informa.SelFiliais TxtBuscaFilial2.Text
    If de_informa.rsSelFiliais.RecordCount > 0 Then
    If IsNull(de_informa.rsSelFiliais.Fields("filial")) = False Then TxtFilial2.Caption = de_informa.rsSelFiliais.Fields("filial")
    If IsNull(de_informa.rsSelFiliais.Fields("nomefilial")) = False Then TxtNomeFilial2.Caption = PriMaiuscula(de_informa.rsSelFiliais.Fields("nomefilial"))
    DoEvents
    End If
End If
End Sub

Private Sub txtCodCia_GotFocus()
    txtCodCia.SelStart = 0
    txtCodCia.SelLength = 3
End Sub

Private Sub txtCodCia_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
        Call txtCodCia_LostFocus
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
End Sub

Private Sub txtCodCia_LostFocus()
    If Len(Trim$(txtCodCia)) > 0 Then
        txtCodCia = UCase(Trim$(txtCodCia))
        lblFantasia = ""
        If de_informa.rsSel_CiaAereaPorCodigo.State = 1 Then de_informa.rsSel_CiaAereaPorCodigo.Close
        de_informa.Sel_CiaAereaPorCodigo Trim$(txtCodCia.Text)
        If de_informa.rsSel_CiaAereaPorCodigo.RecordCount > 0 Then
            lblFantasia = de_informa.rsSel_CiaAereaPorCodigo.Fields("fantasia")
            If Val(txtNumAWB) > 0 Then
                CmdBusca.Enabled = True
            End If
        Else
            MsgBox "Código de Cia. Aérea Não Encontrado !"
            CmdBusca.Enabled = False
            txtCodCia.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub TxtSigla_Change()
TxtSigla.Text = UCase(TxtSigla.Text)
TxtSigla.SelStart = Len(TxtSigla.Text)
End Sub

Private Sub TxtSigla_GotFocus()
    TxtSigla.SelStart = 0
    TxtSigla.SelLength = 3
End Sub

Private Sub TxtSigla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtSigla_LostFocus()
    If Len(Trim$(TxtSigla)) > 0 Then
        TxtSigla = UCase(Trim$(TxtSigla))
        lblFantasiaIns = ""
        If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
        de_informa.SelAeroportoSigla Trim$(TxtSigla.Text)
        If de_informa.rsSelAeroportoSigla.RecordCount > 0 Then
            lblFantasiaIns = de_informa.rsSelAeroportoSigla.Fields("aeroporto")
        Else
            MsgBox "Aeroporto Não Encontrado !", vbCritical, ""
            TxtSigla.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub TxtDig2_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If

End Sub

Private Sub TxtDigFim_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtDigInic_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtMotivo_Change()
If Len(Trim(TxtMotivo.Text)) > 0 Then
TxtMotivo.Text = UCase(TxtMotivo.Text)
TxtMotivo.SelStart = Len((TxtMotivo.Text))
End If
End Sub

Private Sub txtNumAWB_Change()
    SoNumero (txtNumAWB)
    If Val(txtNumAWB) > 0 And Len(Trim$(lblFantasia)) > 0 Then
        CmdBusca.Enabled = True
        'cmdConfirmaCanc.Enabled = True
    Else
        CmdBusca.Enabled = False
        cmdConfirmaCanc.Enabled = False
        LblStatus = ""
    End If
End Sub

Private Sub txtNumAWB_GotFocus()
    txtNumAWB.SelStart = 0
    txtNumAWB.SelLength = 12
End Sub

Private Sub txtNumAWB_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtNumFim_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtNumInic_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If
End Sub

Sub EnviaIMP(xNomeFont As String, xSizeFont As Double, xCurrentX As Double, xCurrentY As Double, xIMP As String)
Printer.ScaleMode = vbMillimeters
Printer.Font.Name = xNomeFont
Printer.Font.Size = xSizeFont
Printer.FontBold = True
Printer.CurrentX = xCurrentX
Printer.CurrentY = xCurrentY
Printer.Print xIMP
End Sub

