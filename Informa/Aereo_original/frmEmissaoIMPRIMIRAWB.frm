VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEmissaoIMPRIMIRAWB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir AWB"
   ClientHeight    =   3870
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   8970
   ControlBox      =   0   'False
   Icon            =   "frmEmissaoIMPRIMIRAWB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "AWBs Emitidos ainda não impressos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   8715
      Begin MSFlexGridLib.MSFlexGrid FlexAWB 
         Height          =   2415
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   4260
         _Version        =   393216
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   7500
      TabIndex        =   5
      Top             =   3420
      Width           =   1335
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   7500
      TabIndex        =   4
      Top             =   3060
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   7275
      Begin VB.TextBox TxtDig 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   6660
         MaxLength       =   1
         TabIndex        =   3
         Top             =   300
         Width           =   435
      End
      Begin VB.TextBox TxtAWB 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4180
         MaxLength       =   10
         TabIndex        =   2
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox TxtSigla 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2360
         MaxLength       =   2
         TabIndex        =   1
         Top             =   300
         Width           =   435
      End
      Begin VB.TextBox TxtFilial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   540
         MaxLength       =   2
         TabIndex        =   0
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Dígito"
         Height          =   195
         Left            =   6180
         TabIndex        =   10
         Top             =   345
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sigla Cia. Aer."
         Height          =   195
         Left            =   1320
         TabIndex        =   9
         Top             =   345
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   345
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº. AWB"
         Height          =   195
         Left            =   3480
         TabIndex        =   7
         Top             =   345
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmEmissaoIMPRIMIRAWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBuscar_Click()
Dim CodAwb As String

CodAwb = TxtFilial.Text & TxtSigla.Text & String(10 - Len(Trim(Str(Val(TxtAWB.Text)))), "0") & Trim(Str(Val(TxtAWB.Text))) & Trim(Str(Val(TxtDig.Text)))
'CodAwb = TxtFilial.Text & TxtSigla.Text & Trim(Str(Val(TxtAwb.Text)))

If de_informa.rsConsultaAWB.State = 1 Then de_informa.rsConsultaAWB.Close
If de_informa.rsConsultaAWBNF.State = 1 Then de_informa.rsConsultaAWBNF.Close
If de_informa.rsConsultaAWBVolume.State = 1 Then de_informa.rsConsultaAWBVolume.Close

de_informa.ConsultaAWB CodAwb
de_informa.ConsultaAWBNF CodAwb
de_informa.ConsultaAWBVolume CodAwb

    If de_informa.rsConsultaAWB.RecordCount = 0 Then
    MsgBox "Código de AWB não encontrado. Revise os dados digitados e tente novamente.", vbExclamation, ""
    Exit Sub
    End If
    
    If AUXCanc = "IMPRIMIR" Then
        If de_informa.rsConsultaAWB.Fields("cancelado") = "X" Then
        MsgBox "Este AWB está cancelado. Não é possível imprimí-lo.", vbCritical, ""
        Exit Sub
        End If
    End If

Call LimpaTela(frmEmissao)
frmEmissao.FlexGridVolumes.Clear
frmEmissao.FlexGridNFs.Clear

With frmEmissao
    .TxtBuscaFilial.Text = de_informa.rsConsultaAWB.Fields("filial")
    'BUSCA DADOS DE FILIAL
    .TxtBuscaFilial.Text = Trim(String(2 - Len(Trim(Str(Val(.TxtBuscaFilial.Text)))), "0")) & Trim(Str(Val(.TxtBuscaFilial.Text)))
    If de_informa.rsSelFiliais.State = 1 Then de_informa.rsSelFiliais.Close
    de_informa.SelFiliais .TxtBuscaFilial.Text
        If de_informa.rsSelFiliais.RecordCount > 0 Then
        If IsNull(de_informa.rsSelFiliais.Fields("filial")) = False Then .TxtFilial.Caption = de_informa.rsSelFiliais.Fields("filial")
        If IsNull(de_informa.rsSelFiliais.Fields("nomefilial")) = False Then .TxtNomeFilial.Caption = PriMaiuscula(de_informa.rsSelFiliais.Fields("nomefilial"))
        If IsNull(de_informa.rsSelFiliais.Fields("cgc")) = False Then .TxtCGCFilial.Caption = de_informa.rsSelFiliais.Fields("cgc")
        If IsNull(de_informa.rsSelFiliais.Fields("inscrest")) = False Then .TxtInscrEstFilial.Caption = de_informa.rsSelFiliais.Fields("inscrest")
        If IsNull(de_informa.rsSelFiliais.Fields("cidade")) = False Then .TxtCidadeFilial.Caption = PriMaiuscula(de_informa.rsSelFiliais.Fields("cidade"))
        If IsNull(de_informa.rsSelFiliais.Fields("uf")) = False Then .TxtUFFilial.Caption = de_informa.rsSelFiliais.Fields("uf")
        If IsNull(de_informa.rsSelFiliais.Fields("licensaIATA")) = False Then .TxtLicensaFilial.Caption = de_informa.rsSelFiliais.Fields("licensaIATA")
        If IsNull(de_informa.rsSelFiliais.Fields("siglaIATA")) = False Then .TxtSiglaFilial.Caption = de_informa.rsSelFiliais.Fields("siglaIATA")
        DoEvents
        End If
        
    .TxtSiglaCiaAerea.Text = de_informa.rsConsultaAWB.Fields("cia")
    .TxtNomeCiaAerea.Caption = PriMaiuscula(de_informa.rsConsultaAWB.Fields("NOMEcia"))
    .TxtCGCCiaAerea.Caption = de_informa.rsConsultaAWB.Fields("CGCcia")
    .TxtInscrEstCiaAerea.Caption = de_informa.rsConsultaAWB.Fields("INSCRESTcia")
    
    .TxtCodIATA.Text = de_informa.rsConsultaAWB.Fields("codiataprod")
    
    .TxtAWB.Text = de_informa.rsConsultaAWB.Fields("AWB")
    .TxtDig.Text = de_informa.rsConsultaAWB.Fields("DIG")
    
    .TxtDescrIATA.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("descrprodsis"))
    
    
    If de_informa.rsSelEspecie.State = 1 Then de_informa.rsSelEspecie.Close
    de_informa.SelEspecie
    .ComboEspecie.Clear
    Do Until de_informa.rsSelEspecie.EOF
    .ComboEspecie.AddItem UCase(de_informa.rsSelEspecie.Fields("especie"))
    de_informa.rsSelEspecie.MoveNext
    Loop
    
    .ComboEspecie.Text = UCase(de_informa.rsConsultaAWB.Fields("especie"))
    .TxtPesoCubado.Text = de_informa.rsConsultaAWB.Fields("pesocubado")
    .TxtPesoReal.Text = de_informa.rsConsultaAWB.Fields("pesoreal")
    .TxtTotalVM = de_informa.rsConsultaAWB.Fields("valmerc")
    
    .TxtNomeExpedidor.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("NOMEexp"))
    .TxtCGCExpedidor.Text = de_informa.rsConsultaAWB.Fields("cnpjexp")
    .TxtCidadeExpedidor.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("CIDADExp"))
    .TxtUFExpedidor.Text = de_informa.rsConsultaAWB.Fields("UFexp")
    .TxtInscrEstExpedidor.Text = de_informa.rsConsultaAWB.Fields("INSCrESTexp")
    .TxtEndExpedidor.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("ENDexp"))
    .TxtBairroEXP.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("BAIRROexp"))
    .TxtCEPExpedidor.Text = de_informa.rsConsultaAWB.Fields("CEPexp")
    .TxtTelExpedidor.Text = de_informa.rsConsultaAWB.Fields("TELexp")
    .TxtFAXExpedidor.Text = de_informa.rsConsultaAWB.Fields("FAXexp")
    .TxtSeguradoraExpedidor.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("SEGexp"))
    .TxtApoliceExpedidor.Text = de_informa.rsConsultaAWB.Fields("APOLICEEXP")
        
    .TxtSiglaExpedidor.Text = de_informa.rsConsultaAWB.Fields("siglaorigem")
    .TxtAeroportoExpedidor.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("CIDADEorigem")) & " - " & de_informa.rsConsultaAWB.Fields("uforigem") & " (" & PriMaiuscula(de_informa.rsConsultaAWB.Fields("aeroportoorigem")) & ")"
    
    .TxtSiglaDestinatario.Text = de_informa.rsConsultaAWB.Fields("siglaDES")
    .TxtAeroportoDestinatario.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("cidadeDEStino")) & " - " & de_informa.rsConsultaAWB.Fields("ufDEStino") & " (" & PriMaiuscula(de_informa.rsConsultaAWB.Fields("aeroportoDEStino")) & ")"
    
    .TxtSiglaVIA.Text = de_informa.rsConsultaAWB.Fields("siglavia")
    .TxtAeroportoVIA.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("cidadevia")) & " - " & de_informa.rsConsultaAWB.Fields("ufvia") & " (" & PriMaiuscula(de_informa.rsConsultaAWB.Fields("aeroportovia")) & ")"
    
    .TxtNomeDestinatario.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("NOMEdes"))
    .TxtCGCDestinatario.Text = de_informa.rsConsultaAWB.Fields("cnpjdes")
    .TxtCidadeDestinatario.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("CIDADEdes"))
    .TxtUFDestinatario.Text = de_informa.rsConsultaAWB.Fields("UFdes")
    .TxtInscrEstDestinatario.Text = de_informa.rsConsultaAWB.Fields("INSCrESTdes")
    .TxtEndDestinatario.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("ENDdes"))
    .TxtBairroDEST.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("BAIRROdes"))
    .TxtCEPDestinatario.Text = de_informa.rsConsultaAWB.Fields("CEPdes")
    .TxtTelDestinatario.Text = de_informa.rsConsultaAWB.Fields("TELdes")
    .TxtFAXDestinatario.Text = de_informa.rsConsultaAWB.Fields("FAXdes")
    .TxtSeguradoraDestinatario.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("SEGdes"))
    .TxtApoliceDestinatario.Text = de_informa.rsConsultaAWB.Fields("APOLICEdes")
    
    .TxtVolumes.Text = de_informa.rsConsultaAWB.Fields("volumes")
    .TxtOBSEmissao.Text = de_informa.rsConsultaAWB.Fields("obsemissor")
    .TxtTipoTaxa.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("tipotaxa"))
        
        If de_informa.rsConsultaAWB.Fields("KILO") = 0 Then
        .TxtKiloCob.Text = ""
        Else
        .TxtKiloCob.Text = de_informa.rsConsultaAWB.Fields("KILO")
        End If
        
    .TxtFreteNacional.Text = de_informa.rsConsultaAWB.Fields("fretenacional")
    .TxtFreteRegional.Text = de_informa.rsConsultaAWB.Fields("freteregional")
    .TxtADValorem.Text = de_informa.rsConsultaAWB.Fields("advalorem")
    .TxtTipoADVAL.Text = de_informa.rsConsultaAWB.Fields("tipoadval")
    .TxtTXOrigem.Text = de_informa.rsConsultaAWB.Fields("txorigem")
    .TxtTXDestino.Text = de_informa.rsConsultaAWB.Fields("txdestino")
    .TxtTXRedesp.Text = de_informa.rsConsultaAWB.Fields("txredesp")
    .TxtOutros1.Text = de_informa.rsConsultaAWB.Fields("txoutros1")
    .TxtDescrOutros1.Text = de_informa.rsConsultaAWB.Fields("descrtxoutros1")
    .TxtOutros2.Text = de_informa.rsConsultaAWB.Fields("txoutros2")
    .TxtDescrOutros2.Text = de_informa.rsConsultaAWB.Fields("descrtxoutros2")
    .TxtFreteTotal.Text = de_informa.rsConsultaAWB.Fields("fretetotal")
    .TxtAliquota.Text = de_informa.rsConsultaAWB.Fields("aliquota")
    .TxtICMS.Text = de_informa.rsConsultaAWB.Fields("icms")
    .TxtAutorizador.Text = de_informa.rsConsultaAWB.Fields("spotautorizador")
    .TxtKilo.Text = de_informa.rsConsultaAWB.Fields("spotkilo")
    .TxtTotalVM.Text = de_informa.rsConsultaAWB.Fields("valmerc")
    xUsuarioIMP = de_informa.rsConsultaAWB.Fields("emissor")
    xDataIMP = de_informa.rsConsultaAWB.Fields("data")
    xHoraIMP = de_informa.rsConsultaAWB.Fields("hora")
    
    If de_informa.rsConsultaAWB.Fields("perecivel") = "S" Then
    .ChkPerecivel.Value = 1
    Else
    .ChkPerecivel.Value = 0
    End If
    
    If de_informa.rsConsultaAWB.Fields("modal") = "PAGO" Then
    .OptPago.Value = True
    .OptAPagar.Value = False
    Else
    .OptPago.Value = False
    .OptAPagar.Value = True
    End If
    
    If de_informa.rsConsultaAWB.Fields("retira") = "S" Then
    .OptRetiraSim.Value = True
    .OptRetiraNao.Value = False
    Else
    .OptRetiraSim.Value = False
    .OptRetiraNao.Value = True
    End If
    .TxtLocalRetirada.Text = UCase(Trim(de_informa.rsConsultaAWB.Fields("LOCALRETIRADA")))
    
'BUSCA NOTAS

Dim X, Y As Integer
.FlexGridNFs.Rows = de_informa.rsConsultaAWBNF.RecordCount + 1
.FlexGridNFs.Cols = 6
    If .FlexGridNFs.Rows > 1 Then
    .FlexGridNFs.FixedRows = 1
    Else
    .FlexGridNFs.Rows = .FlexGridNFs.Rows + 1
    .FlexGridNFs.FixedRows = 1
    End If
.FlexGridNFs.FixedCols = 0
.FlexGridNFs.TextMatrix(0, 0) = "N/D"
.FlexGridNFs.TextMatrix(0, 1) = "Filial"
.FlexGridNFs.TextMatrix(0, 2) = "CTC"
.FlexGridNFs.TextMatrix(0, 3) = "Nº NF"
.FlexGridNFs.TextMatrix(0, 4) = "Série"
.FlexGridNFs.TextMatrix(0, 5) = "Valor"
.FlexGridNFs.ColWidth(0) = 500
.FlexGridNFs.ColWidth(1) = 900
.FlexGridNFs.ColWidth(2) = 1200
.FlexGridNFs.ColWidth(3) = 1200
.FlexGridNFs.ColWidth(4) = 700
.FlexGridNFs.ColWidth(5) = 1200

Y = 0
    Do Until de_informa.rsConsultaAWBNF.EOF
    Y = Y + 1
    If IsNull(de_informa.rsConsultaAWBNF.Fields("tipo")) = False Then .FlexGridNFs.TextMatrix(Y, 0) = de_informa.rsConsultaAWBNF.Fields("tipo")
    If IsNull(de_informa.rsConsultaAWBNF.Fields("filialctc")) = False Then .FlexGridNFs.TextMatrix(Y, 1) = Mid(de_informa.rsConsultaAWBNF.Fields("filialctc"), 1, 2)
    If IsNull(de_informa.rsConsultaAWBNF.Fields("filialctc")) = False Then .FlexGridNFs.TextMatrix(Y, 2) = Val(Mid(de_informa.rsConsultaAWBNF.Fields("filialctc"), 3))
    If IsNull(de_informa.rsConsultaAWBNF.Fields("nota")) = False Then .FlexGridNFs.TextMatrix(Y, 3) = de_informa.rsConsultaAWBNF.Fields("nota")
    If IsNull(de_informa.rsConsultaAWBNF.Fields("serie")) = False Then .FlexGridNFs.TextMatrix(Y, 4) = de_informa.rsConsultaAWBNF.Fields("serie")
    If IsNull(de_informa.rsConsultaAWBNF.Fields("valor")) = False Then .FlexGridNFs.TextMatrix(Y, 5) = de_informa.rsConsultaAWBNF.Fields("valor")
    de_informa.rsConsultaAWBNF.MoveNext
    Loop
   
'BUSCA VOLUMES
.FlexGridVolumes.Rows = de_informa.rsConsultaAWBVolume.RecordCount + 1
.FlexGridVolumes.Cols = 5
.FlexGridVolumes.FixedCols = 0
    If .FlexGridVolumes.Rows > 1 Then
    .FlexGridVolumes.FixedRows = 1
    Else
    .FlexGridVolumes.Rows = .FlexGridVolumes.Rows + 1
    .FlexGridVolumes.FixedRows = 1
    End If
.FlexGridVolumes.TextMatrix(0, 0) = "Vol."
.FlexGridVolumes.TextMatrix(0, 1) = "Comp. (Cm)"
.FlexGridVolumes.TextMatrix(0, 2) = "Larg. (Cm)"
.FlexGridVolumes.TextMatrix(0, 3) = "Alt. (Cm)"
.FlexGridVolumes.TextMatrix(0, 4) = "Peso (Kg)"
.FlexGridVolumes.ColWidth(0) = 900
.FlexGridVolumes.ColWidth(1) = 900
.FlexGridVolumes.ColWidth(2) = 900
.FlexGridVolumes.ColWidth(3) = 900
.FlexGridVolumes.ColWidth(4) = 900

Y = 0
    Do Until de_informa.rsConsultaAWBVolume.EOF
    Y = Y + 1
    If IsNull(de_informa.rsConsultaAWBVolume.Fields("volumes")) = False Then .FlexGridVolumes.TextMatrix(Y, 0) = de_informa.rsConsultaAWBVolume.Fields("volumes")
    If IsNull(de_informa.rsConsultaAWBVolume.Fields("comprimento")) = False Then .FlexGridVolumes.TextMatrix(Y, 1) = de_informa.rsConsultaAWBVolume.Fields("comprimento")
    If IsNull(de_informa.rsConsultaAWBVolume.Fields("largura")) = False Then .FlexGridVolumes.TextMatrix(Y, 2) = de_informa.rsConsultaAWBVolume.Fields("largura")
    If IsNull(de_informa.rsConsultaAWBVolume.Fields("altura")) = False Then .FlexGridVolumes.TextMatrix(Y, 3) = de_informa.rsConsultaAWBVolume.Fields("altura")
    If IsNull(de_informa.rsConsultaAWBVolume.Fields("pesoreal")) = False Then .FlexGridVolumes.TextMatrix(Y, 4) = de_informa.rsConsultaAWBVolume.Fields("pesoreal")
    de_informa.rsConsultaAWBVolume.MoveNext
    Loop
End With



    If MsgBox("Você deseja fazer alguma alteração neste AWB que você está trazendo?", vbYesNo + vbQuestion, "") = vbNo Then
    Call TravaFrame(frmEmissao, frmEmissao.FraBotoes, 0)
    Acao = "IMPRIMIR"
    frmEmissao.CmdEmitir.Caption = "Imprimir AWB"
    frmEmissao.LblAtualizarFrete.Caption = "Nao"
    Else
    Acao = "ALTERAR"
    frmEmissao.FraAWB.Enabled = False
    frmEmissao.CmdEmitir.Caption = "Gravar AWB"
    frmEmissao.LblAtualizarFrete.Caption = "Sim"
    End If
    
frmEmissao.TxtFreteNacional.Text = Format(frmEmissao.TxtFreteNacional.Text, "###,###,###,##0.00")
frmEmissao.TxtFreteRegional.Text = Format(frmEmissao.TxtFreteRegional.Text, "###,###,###,##0.00")
frmEmissao.TxtKiloCob.Text = Format(frmEmissao.TxtKiloCob.Text, "###,###,###,##0.00")
frmEmissao.TxtADValorem.Text = Format(frmEmissao.TxtADValorem.Text, "###,###,###,##0.00")
frmEmissao.TxtTXOrigem.Text = Format(frmEmissao.TxtTXOrigem.Text, "###,###,###,##0.00")
frmEmissao.TxtTXDestino.Text = Format(frmEmissao.TxtTXDestino.Text, "###,###,###,##0.00")
frmEmissao.TxtTXRedesp.Text = Format(frmEmissao.TxtTXRedesp.Text, "###,###,###,##0.00")
frmEmissao.TxtOutros1.Text = Format(frmEmissao.TxtOutros1.Text, "###,###,###,##0.00")
frmEmissao.TxtOutros2.Text = Format(frmEmissao.TxtOutros2.Text, "###,###,###,##0.00")
frmEmissao.TxtFreteTotal.Text = Format(frmEmissao.TxtFreteTotal.Text, "###,###,###,##0.00")
frmEmissao.TxtICMS.Text = Format(frmEmissao.TxtICMS.Text, "###,###,###,##0.00")
frmEmissao.TxtTotalVM.Text = Format(frmEmissao.TxtTotalVM.Text, "###,###,###,##0.00")
frmEmissao.TxtPesoCubado.Text = Format(frmEmissao.TxtPesoCubado.Text, "###,###,###,##0.0")
frmEmissao.TxtPesoReal.Text = Format(frmEmissao.TxtPesoReal.Text, "###,###,###,##0.0")

Unload Me

End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()

If de_informa.rsAWBsNaoImpressos.State = 1 Then de_informa.rsAWBsNaoImpressos.Close
de_informa.AWBsNaoImpressos "%", "%"

FlexAWB.Clear
FlexAWB.Cols = 10
FlexAWB.FixedCols = 0
FlexAWB.Rows = de_informa.rsAWBsNaoImpressos.RecordCount + 1

FlexAWB.TextMatrix(0, 0) = "Filial"
FlexAWB.TextMatrix(0, 1) = "AWB"
FlexAWB.TextMatrix(0, 2) = "Dig."
FlexAWB.TextMatrix(0, 3) = "Cia."
FlexAWB.TextMatrix(0, 4) = "Origem"
FlexAWB.TextMatrix(0, 5) = "VIA"
FlexAWB.TextMatrix(0, 6) = "Destino"
FlexAWB.TextMatrix(0, 7) = "Data"
FlexAWB.TextMatrix(0, 8) = "Hora"
FlexAWB.TextMatrix(0, 9) = "Emissor"

FlexAWB.ColWidth(0) = 500
FlexAWB.ColWidth(1) = 800
FlexAWB.ColWidth(2) = 400
FlexAWB.ColWidth(3) = 500
FlexAWB.ColWidth(4) = 2500
FlexAWB.ColWidth(5) = 2500
FlexAWB.ColWidth(6) = 2500
FlexAWB.ColWidth(7) = 1000
FlexAWB.ColWidth(8) = 1000
FlexAWB.ColWidth(9) = 1000

X = 0

    
    Do Until de_informa.rsAWBsNaoImpressos.EOF
    X = X + 1
    FlexAWB.TextMatrix(X, 0) = de_informa.rsAWBsNaoImpressos.Fields("filial")
    FlexAWB.TextMatrix(X, 1) = de_informa.rsAWBsNaoImpressos.Fields("awb")
    FlexAWB.TextMatrix(X, 2) = de_informa.rsAWBsNaoImpressos.Fields("dig")
    FlexAWB.TextMatrix(X, 3) = de_informa.rsAWBsNaoImpressos.Fields("cia")
    If IsNull(de_informa.rsAWBsNaoImpressos.Fields("cidadeorigem")) = False Then FlexAWB.TextMatrix(X, 4) = de_informa.rsAWBsNaoImpressos.Fields("cidadeorigem")
    If IsNull(de_informa.rsAWBsNaoImpressos.Fields("cidadevia")) = False Then FlexAWB.TextMatrix(X, 5) = de_informa.rsAWBsNaoImpressos.Fields("cidadevia")
    If IsNull(de_informa.rsAWBsNaoImpressos.Fields("cidadedestino")) = False Then FlexAWB.TextMatrix(X, 6) = de_informa.rsAWBsNaoImpressos.Fields("cidadedestino")
    FlexAWB.TextMatrix(X, 7) = de_informa.rsAWBsNaoImpressos.Fields("data")
    FlexAWB.TextMatrix(X, 8) = de_informa.rsAWBsNaoImpressos.Fields("hora")
    FlexAWB.TextMatrix(X, 9) = de_informa.rsAWBsNaoImpressos.Fields("emissor")
    de_informa.rsAWBsNaoImpressos.MoveNext
    Loop

End Sub

Private Sub TxtAwb_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 80
End Sub

Private Sub TxtAWB_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtAWB_LostFocus()
If Len(Trim(TxtFilial.Text)) > 0 And Len(Trim(TxtSigla.Text)) > 0 And Len(Trim(TxtAWB.Text)) > 0 Then
Me.MousePointer = 11
DoEvents

If de_informa.rsConfereNumeroAWB.State = 1 Then de_informa.rsConfereNumeroAWB.Close
de_informa.ConfereNumeroAWB TxtSigla.Text, TxtFilial.Text, TxtAWB.Text

    If de_informa.rsConfereNumeroAWB.RecordCount = 0 Then
    MsgBox "Este formulário não está cadastrado!.", vbCritical, ""
    TxtAWB.Text = ""
    TxtDig.Text = ""
    TxtAWB.SetFocus
    Me.MousePointer = 0
    DoEvents
    Exit Sub
    ElseIf de_informa.rsConfereNumeroAWB.Fields("tem_ocorr") = "C" Then
    MsgBox "O formulário para este AWB está cancelado. Para utilizá-lo, vá até o cadastro de formulários e descancele-o.", vbCritical, ""
    TxtAWB.Text = ""
    TxtDig.Text = ""
    TxtAWB.SetFocus
    Me.MousePointer = 0
    DoEvents
    Exit Sub
    Else
    TxtDig.Text = de_informa.rsConfereNumeroAWB.Fields("dig")
    End If
Else
TxtAWB.Text = ""
TxtDig.Text = ""
Me.MousePointer = 0
DoEvents
End If
Me.MousePointer = 0
DoEvents
End Sub

Private Sub TxtDig_GotFocus()
TxtDig.SelStart = 0
TxtDig.SelLength = 10
End Sub

Private Sub TxtDig_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtFilial_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 3
End Sub

Private Sub TxtFilial_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        ElseIf KeyAscii <> 8 Then
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtFilial_LostFocus()
TxtFilial.Text = String(2 - Len(TxtFilial.Text), "0") & TxtFilial.Text

Me.MousePointer = 11
DoEvents

If de_informa.rsAWBsNaoImpressos.State = 1 Then de_informa.rsAWBsNaoImpressos.Close
de_informa.AWBsNaoImpressos "%", TxtFilial.Text & "%"

FlexAWB.Clear
FlexAWB.Cols = 10
FlexAWB.FixedCols = 0
FlexAWB.Rows = de_informa.rsAWBsNaoImpressos.RecordCount + 1

FlexAWB.TextMatrix(0, 0) = "Filial"
FlexAWB.TextMatrix(0, 1) = "AWB"
FlexAWB.TextMatrix(0, 2) = "Dig."
FlexAWB.TextMatrix(0, 3) = "Cia."
FlexAWB.TextMatrix(0, 4) = "Origem"
FlexAWB.TextMatrix(0, 5) = "VIA"
FlexAWB.TextMatrix(0, 6) = "Destino"
FlexAWB.TextMatrix(0, 7) = "Data"
FlexAWB.TextMatrix(0, 8) = "Hora"
FlexAWB.TextMatrix(0, 9) = "Emissor"

FlexAWB.ColWidth(0) = 500
FlexAWB.ColWidth(1) = 800
FlexAWB.ColWidth(2) = 400
FlexAWB.ColWidth(3) = 500
FlexAWB.ColWidth(4) = 2500
FlexAWB.ColWidth(5) = 2500
FlexAWB.ColWidth(6) = 2500
FlexAWB.ColWidth(7) = 1000
FlexAWB.ColWidth(8) = 1000
FlexAWB.ColWidth(9) = 1000

X = 0

    
    Do Until de_informa.rsAWBsNaoImpressos.EOF
    X = X + 1
    FlexAWB.TextMatrix(X, 0) = de_informa.rsAWBsNaoImpressos.Fields("filial")
    FlexAWB.TextMatrix(X, 1) = de_informa.rsAWBsNaoImpressos.Fields("awb")
    FlexAWB.TextMatrix(X, 2) = de_informa.rsAWBsNaoImpressos.Fields("dig")
    FlexAWB.TextMatrix(X, 3) = de_informa.rsAWBsNaoImpressos.Fields("cia")
    If IsNull(de_informa.rsAWBsNaoImpressos.Fields("cidadeorigem")) = False Then FlexAWB.TextMatrix(X, 4) = de_informa.rsAWBsNaoImpressos.Fields("cidadeorigem")
    If IsNull(de_informa.rsAWBsNaoImpressos.Fields("cidadevia")) = False Then FlexAWB.TextMatrix(X, 5) = de_informa.rsAWBsNaoImpressos.Fields("cidadevia")
    If IsNull(de_informa.rsAWBsNaoImpressos.Fields("cidadedestino")) = False Then FlexAWB.TextMatrix(X, 6) = de_informa.rsAWBsNaoImpressos.Fields("cidadedestino")
    FlexAWB.TextMatrix(X, 7) = de_informa.rsAWBsNaoImpressos.Fields("data")
    FlexAWB.TextMatrix(X, 8) = de_informa.rsAWBsNaoImpressos.Fields("hora")
    FlexAWB.TextMatrix(X, 9) = de_informa.rsAWBsNaoImpressos.Fields("emissor")
    de_informa.rsAWBsNaoImpressos.MoveNext
    Loop

Me.MousePointer = 0
DoEvents
    
End Sub

Private Sub TxtSigla_GotFocus()
ActiveControl.SelStart = 0
ActiveControl.SelLength = 3
End Sub

Private Sub TxtSigla_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
End Sub

Private Sub TxtSigla_LostFocus()
TxtSigla.Text = UCase(TxtSigla.Text)

Me.MousePointer = 11
DoEvents

If de_informa.rsAWBsNaoImpressos.State = 1 Then de_informa.rsAWBsNaoImpressos.Close
de_informa.AWBsNaoImpressos TxtSigla.Text & "%", TxtFilial.Text & "%"

FlexAWB.Clear
FlexAWB.Cols = 10
FlexAWB.FixedCols = 0
FlexAWB.Rows = de_informa.rsAWBsNaoImpressos.RecordCount + 1

FlexAWB.TextMatrix(0, 0) = "Filial"
FlexAWB.TextMatrix(0, 1) = "AWB"
FlexAWB.TextMatrix(0, 2) = "Dig."
FlexAWB.TextMatrix(0, 3) = "Cia."
FlexAWB.TextMatrix(0, 4) = "Origem"
FlexAWB.TextMatrix(0, 5) = "VIA"
FlexAWB.TextMatrix(0, 6) = "Destino"
FlexAWB.TextMatrix(0, 7) = "Data"
FlexAWB.TextMatrix(0, 8) = "Hora"
FlexAWB.TextMatrix(0, 9) = "Emissor"

FlexAWB.ColWidth(0) = 500
FlexAWB.ColWidth(1) = 800
FlexAWB.ColWidth(2) = 400
FlexAWB.ColWidth(3) = 500
FlexAWB.ColWidth(4) = 2500
FlexAWB.ColWidth(5) = 2500
FlexAWB.ColWidth(6) = 2500
FlexAWB.ColWidth(7) = 1000
FlexAWB.ColWidth(8) = 1000
FlexAWB.ColWidth(9) = 1000

X = 0

    
    Do Until de_informa.rsAWBsNaoImpressos.EOF
    X = X + 1
    FlexAWB.TextMatrix(X, 0) = de_informa.rsAWBsNaoImpressos.Fields("filial")
    FlexAWB.TextMatrix(X, 1) = de_informa.rsAWBsNaoImpressos.Fields("awb")
    FlexAWB.TextMatrix(X, 2) = de_informa.rsAWBsNaoImpressos.Fields("dig")
    FlexAWB.TextMatrix(X, 3) = de_informa.rsAWBsNaoImpressos.Fields("cia")
    If IsNull(de_informa.rsAWBsNaoImpressos.Fields("cidadeorigem")) = False Then FlexAWB.TextMatrix(X, 4) = de_informa.rsAWBsNaoImpressos.Fields("cidadeorigem")
    If IsNull(de_informa.rsAWBsNaoImpressos.Fields("cidadevia")) = False Then FlexAWB.TextMatrix(X, 5) = de_informa.rsAWBsNaoImpressos.Fields("cidadevia")
    If IsNull(de_informa.rsAWBsNaoImpressos.Fields("cidadedestino")) = False Then FlexAWB.TextMatrix(X, 6) = de_informa.rsAWBsNaoImpressos.Fields("cidadedestino")
    FlexAWB.TextMatrix(X, 7) = de_informa.rsAWBsNaoImpressos.Fields("data")
    FlexAWB.TextMatrix(X, 8) = de_informa.rsAWBsNaoImpressos.Fields("hora")
    FlexAWB.TextMatrix(X, 9) = de_informa.rsAWBsNaoImpressos.Fields("emissor")
    de_informa.rsAWBsNaoImpressos.MoveNext
    Loop

Me.MousePointer = 0
DoEvents
    
End Sub

