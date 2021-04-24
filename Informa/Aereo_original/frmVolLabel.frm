VERSION 5.00
Begin VB.Form frmVolLabel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Emissão de Etiquetas de Volume"
   ClientHeight    =   6495
   ClientLeft      =   5085
   ClientTop       =   1425
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmVolLabel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox LogoVARIG 
      Height          =   735
      Left            =   1200
      Picture         =   "frmVolLabel.frx":000C
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   45
      Top             =   7620
      Width           =   855
   End
   Begin VB.PictureBox LogoPANTANAL 
      Height          =   735
      Left            =   2100
      Picture         =   "frmVolLabel.frx":1F4F
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   44
      Top             =   7620
      Width           =   855
   End
   Begin VB.PictureBox LogoTAM 
      Height          =   735
      Left            =   3000
      Picture         =   "frmVolLabel.frx":3CDF
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   43
      Top             =   7620
      Width           =   855
   End
   Begin VB.PictureBox LogoOCEAN 
      Height          =   735
      Left            =   3900
      Picture         =   "frmVolLabel.frx":7E4A
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   42
      Top             =   7620
      Width           =   855
   End
   Begin VB.PictureBox LogoVASP 
      Height          =   735
      Left            =   300
      Picture         =   "frmVolLabel.frx":92AE
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   41
      Top             =   7620
      Width           =   855
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
      Height          =   6315
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   4815
      Begin VB.TextBox TxtLote 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   39
         Top             =   5340
         Width           =   1815
      End
      Begin VB.TextBox TxtBuscaLote 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3540
         MaxLength       =   10
         TabIndex        =   37
         Top             =   5040
         Width           =   1155
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   3480
         TabIndex        =   4
         Top             =   840
         Width           =   1155
      End
      Begin VB.TextBox TxtVolumes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3540
         MaxLength       =   10
         TabIndex        =   5
         Top             =   4740
         Width           =   1155
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados do AWB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   4575
         Begin VB.TextBox TxtCodAWB 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   600
            MaxLength       =   50
            TabIndex        =   36
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox TxtVolumesAWB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3300
            MaxLength       =   10
            TabIndex        =   33
            Top             =   2940
            Width           =   1155
         End
         Begin VB.TextBox TxtPesoTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3300
            MaxLength       =   10
            TabIndex        =   31
            Top             =   2640
            Width           =   1155
         End
         Begin VB.TextBox TxtSiglaDestino 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   29
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox TxtAeroportoDestino 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2100
            MaxLength       =   500
            TabIndex        =   28
            Top             =   2280
            Width           =   2355
         End
         Begin VB.TextBox TxtCidadeDestino 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            MaxLength       =   500
            TabIndex        =   26
            Top             =   1980
            Width           =   2895
         End
         Begin VB.TextBox TxtDestinatario 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            MaxLength       =   500
            TabIndex        =   24
            Top             =   1680
            Width           =   2895
         End
         Begin VB.TextBox TxtSiglaOrigem 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   22
            Top             =   1260
            Width           =   495
         End
         Begin VB.TextBox TxtAeroportoOrigem 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2100
            MaxLength       =   500
            TabIndex        =   21
            Top             =   1260
            Width           =   2355
         End
         Begin VB.TextBox TxtCidadeOrigem 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            MaxLength       =   500
            TabIndex        =   19
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox TxtExpedidor 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            MaxLength       =   500
            TabIndex        =   17
            Top             =   660
            Width           =   2895
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "AWB:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   285
            Width           =   420
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Volumes no AWB:"
            Height          =   195
            Left            =   1980
            TabIndex        =   32
            Top             =   3000
            Width           =   1290
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Peso Total:"
            Height          =   195
            Left            =   2460
            TabIndex        =   30
            Top             =   2685
            Width           =   810
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Aeroporto Destino:"
            Height          =   195
            Left            =   180
            TabIndex        =   27
            Top             =   2325
            Width           =   1320
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Cidade Destino:"
            Height          =   195
            Left            =   375
            TabIndex        =   25
            Top             =   2025
            Width           =   1125
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Destinatário:"
            Height          =   195
            Left            =   615
            TabIndex        =   23
            Top             =   1725
            Width           =   885
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Aeroporto Origem:"
            Height          =   195
            Left            =   225
            TabIndex        =   20
            Top             =   1305
            Width           =   1275
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cidade Origem:"
            Height          =   195
            Left            =   420
            TabIndex        =   18
            Top             =   1005
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Expedidor:"
            Height          =   195
            Left            =   750
            TabIndex        =   16
            Top             =   720
            Width           =   750
         End
      End
      Begin VB.TextBox TxtAwb 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   2
         Top             =   840
         Width           =   1035
      End
      Begin VB.TextBox TxtDig 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   3
         Top             =   840
         Width           =   315
      End
      Begin VB.TextBox TxtBuscaFilial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   540
         MaxLength       =   2
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdConfirma 
         Caption         =   "Confirmar"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   5760
         Width           =   2295
      End
      Begin VB.CommandButton cmdDesisteIns 
         Caption         =   "Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox TxtSigla 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   540
         MaxLength       =   3
         TabIndex        =   1
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lote:"
         Height          =   195
         Left            =   2445
         TabIndex        =   40
         Top             =   5385
         Width           =   360
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Número do Lote:"
         Height          =   195
         Left            =   2280
         TabIndex        =   38
         Top             =   5085
         Width           =   1185
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quantidade de Volumes para Entiquetar:"
         Height          =   195
         Left            =   600
         TabIndex        =   34
         Top             =   4785
         Width           =   2865
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "AWB e Dígito:"
         Height          =   195
         Left            =   900
         TabIndex        =   14
         Top             =   885
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         Height          =   195
         Left            =   165
         TabIndex        =   13
         Top             =   285
         Width           =   345
      End
      Begin VB.Label TxtNomeFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1620
         TabIndex        =   12
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label TxtFilial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sigla:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   585
         Width           =   390
      End
      Begin VB.Label lblFantasiaIns 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   540
         Width           =   3555
      End
   End
End
Attribute VB_Name = "frmVolLabel"
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

Private Sub CmdBuscar_Click()

    If Len(Trim(TxtAWB.Text)) = 0 And Len(Trim(TxtDig.Text)) = 0 Then
    Exit Sub
    ElseIf Len(Trim(TxtFilial.Caption)) = 0 And Len(Trim(TxtSigla.Text)) = 0 Then
    Exit Sub
    End If
    
    
'xCodAwb = String(2 - Len(Trim(Str(Val(TxtFilial.Caption)))), "0") & Trim(Str(Val(TxtFilial.Caption))) & UCase(Trim(TxtSigla.Text)) & String(10 - Len(Trim(Str(Val(TxtAwb.Text)))), "0") & Trim(Str(Val(TxtAwb.Text))) & Trim(Str(Val(TxtDig.Text)))
xCodAwb = String(2 - Len(Trim(Str(Val(TxtFilial.Caption)))), "0") & Trim(Str(Val(TxtFilial.Caption))) & UCase(Trim(TxtSigla.Text)) & Trim(Str(Val(TxtAWB.Text)))

If de_informa.rsConsultaAWB.State = 1 Then de_informa.rsConsultaAWB.Close
de_informa.ConsultaAWB xCodAwb

    If de_informa.rsConsultaAWB.RecordCount = 1 Then
    TxtCodAWB.Text = de_informa.rsConsultaAWB.Fields("awb") & "-" & de_informa.rsConsultaAWB.Fields("dig")
    TxtExpedidor.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("nomeexp"))
    TxtCidadeOrigem.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("cidadexp"))
    TxtSiglaOrigem.Text = de_informa.rsConsultaAWB.Fields("siglaorigem")
    If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
    de_informa.SelAeroportoSigla TxtSiglaOrigem.Text
    TxtAeroportoOrigem.Text = PriMaiuscula(de_informa.rsSelAeroportoSigla.Fields("aeroporto"))
    
    TxtDestinatario.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("nomedes"))
    TxtCidadeDestino.Text = PriMaiuscula(de_informa.rsConsultaAWB.Fields("cidadedes"))
    TxtSiglaDestino.Text = de_informa.rsConsultaAWB.Fields("siglades")
    If de_informa.rsSelAeroportoSigla.State = 1 Then de_informa.rsSelAeroportoSigla.Close
    de_informa.SelAeroportoSigla TxtSiglaDestino.Text
    TxtAeroportoDestino.Text = PriMaiuscula(de_informa.rsSelAeroportoSigla.Fields("aeroporto"))
    
        If de_informa.rsConsultaAWB.Fields("pesoreal") > de_informa.rsConsultaAWB.Fields("pesocubado") Then
        TxtPesoTotal.Text = Format(de_informa.rsConsultaAWB.Fields("pesoreal"), "##0.0")
        Else
        TxtPesoTotal.Text = Format(de_informa.rsConsultaAWB.Fields("pesocubado"), "##0.0")
        End If
        
    TxtVolumesAWB.Text = de_informa.rsConsultaAWB.Fields("volumes")
    
    
    Else
    MsgBox "Sua busca não retornou registro algum.", vbExclamation, ""
    End If
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
            If Mid(SETIMPxLinha, 1, 3) = "ETV" Then
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


frmVolLabel.MousePointer = 11


    If ((Val(TxtVolumes.Text) / 2) - Int(Val(TxtVolumes.Text) / 2)) > 0 Then
    xPagina = Int(Val(TxtVolumes.Text) / 2) + 1
    Else
    xPagina = Int(Val(TxtVolumes.Text) / 2)
    End If

xEtiqueta = 0
Recuo = 13

    For xCont = 1 To Val(TxtVolumes.Text)
    xEtiqueta = xEtiqueta + 1
        If xEtiqueta > 4 Then
        Printer.NewPage
        xEtiqueta = 1
        End If
        
    If xEtiqueta = 1 Then
    MargemTOP = 0
    MargemLeft = -4
    ElseIf xEtiqueta = 2 Then
    MargemTOP = 0
    MargemLeft = 101
    ElseIf xEtiqueta = 3 Then
    MargemTOP = 135
    MargemLeft = -4
    ElseIf xEtiqueta = 4 Then
    MargemTOP = 135
    MargemLeft = 101
    End If
    
    Printer.CurrentY = 9 + MargemTOP
    Printer.CurrentX = 58 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 15
    Printer.Print "    CARGA"
    
    Printer.CurrentY = 16 + MargemTOP
    Printer.CurrentX = 58 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 15
    Printer.Print "DOMÉSTICA"
    
    Printer.CurrentY = 26 + MargemTOP
    Printer.CurrentX = 7 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 7
    Printer.Print "Nº DO CONHECIMENTO"
    
    Printer.CurrentY = 29 + MargemTOP
    Printer.CurrentX = 7 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 12
        If TxtSigla.Text = "RG" Then
        Printer.Print "042 -   " & TxtAWB.Text & " - " & TxtDig.Text
        ElseIf TxtSigla.Text = "OC" Then
        Printer.Print "222 -   " & TxtAWB.Text & " - " & TxtDig.Text
        ElseIf TxtSigla.Text = "P8" Then
        Printer.Print "146 -   " & TxtAWB.Text & " - " & TxtDig.Text
        ElseIf TxtSigla.Text = "KK" Then
        Printer.Print "        " & TxtAWB.Text & " - " & TxtDig.Text
        ElseIf TxtSigla.Text = "VP" Then
        Printer.Print "        " & TxtAWB.Text & " - " & TxtDig.Text
        End If
    
    Printer.CurrentY = 36 + MargemTOP
    Printer.CurrentX = 7 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 7
    Printer.Print "AEROPORTO DE ORIGEM"
    
    Printer.CurrentY = 39 + MargemTOP
    Printer.CurrentX = 7 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 12
    Printer.Print TxtSiglaOrigem.Text & " - " & TxtAeroportoOrigem.Text
    
    
    Printer.CurrentY = 46 + MargemTOP
    Printer.CurrentX = 48 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 8
    Printer.Print "PARA"
    
    Printer.CurrentY = 50 + MargemTOP
    Printer.CurrentX = 7 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 125
    Printer.Print TxtSiglaDestino.Text
    
    Printer.CurrentY = 105 + MargemTOP
    Printer.CurrentX = 7 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 10
    Printer.Print "Lote Interno: " & TxtLote.Text
    
    Printer.CurrentY = 111 + MargemTOP
    Printer.CurrentX = 7 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 7
    Printer.Print "Nº TOTAL DE VOLUMES"
    
    Printer.CurrentY = 117 + MargemTOP
    Printer.CurrentX = 7 + MargemLeft
    Printer.Font = "COURIER NEW"
    Printer.FontBold = True
    Printer.FontSize = 26
    Printer.Print String(3 - Len(Trim(Str(xCont))), " ") & Trim(Str(xCont)) & "/" & String(3 - Len(Trim(Str(Val(TxtVolumes.Text)))), " ") & Trim(Str(Val(TxtVolumes.Text)))
    
    Printer.CurrentY = 111 + MargemTOP
    Printer.CurrentX = 51 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 7
    Printer.Print "PESO TOTAL"
    
    Printer.CurrentY = 116 + MargemTOP
    Printer.CurrentX = 53 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 30
    Printer.Print String(7 - Len(Trim(TxtPesoTotal.Text)), " ") & Trim(TxtPesoTotal.Text)
    
    
    Printer.CurrentY = 125 + MargemTOP
    Printer.CurrentX = 93 + MargemLeft
    Printer.Font = "ARIAL"
    Printer.FontBold = True
    Printer.FontSize = 9
    Printer.Print "KG"
    
    
    Printer.ForeColor = &H80000008     'PRETO
    Printer.ScaleMode = vbMillimeters
    Printer.DrawWidth = 8
    Printer.DrawMode = 9
        
        If TxtSigla.Text = "RG" Then
        Printer.PaintPicture LogoVARIG.Picture, 8 + MargemLeft, 4 + MargemTOP, LogoVARIG.Picture.Width * 0.006, LogoVARIG.Picture.Height * 0.006 'VARIG
        ElseIf TxtSigla.Text = "OC" Then
        Printer.PaintPicture LogoOCEAN.Picture, 7 + MargemLeft, 5 + MargemTOP, LogoOCEAN.Picture.Width * 0.007, LogoOCEAN.Picture.Height * 0.007 'OCEAN AIR
        ElseIf TxtSigla.Text = "P8" Then
        Printer.PaintPicture LogoPANTANAL.Picture, 7 + MargemLeft, 6 + MargemTOP, LogoPANTANAL.Picture.Width * 0.0065, LogoPANTANAL.Picture.Height * 0.0065 'PANTANAL
        ElseIf TxtSigla.Text = "KK" Then
        Printer.PaintPicture LogoTAM.Picture, 7.5 + MargemLeft, 6.6 + MargemTOP, LogoTAM.Picture.Width * 0.0045, LogoTAM.Picture.Height * 0.0045 'TAM
        ElseIf TxtSigla.Text = "VP" Then
        Printer.PaintPicture LogoVASP.Picture, 6 + MargemLeft, 6 + MargemTOP, LogoVASP.Picture.Width * 0.0045, LogoVASP.Picture.Height * 0.0045  'VASP
        End If
    
    Printer.Line (5 + MargemLeft, 5 + MargemTOP)-(50 + MargemLeft, 25 + MargemTOP), , B
    Printer.Line (50 + MargemLeft, 5 + MargemTOP)-(100 + MargemLeft, 25 + MargemTOP), , B
    Printer.Line (5 + MargemLeft, 25 + MargemTOP)-(100 + MargemLeft, 35 + MargemTOP), , B
    Printer.Line (5 + MargemLeft, 35 + MargemTOP)-(100 + MargemLeft, 45 + MargemTOP), , B
    Printer.Line (5 + MargemLeft, 45 + MargemTOP)-(100 + MargemLeft, 110 + MargemTOP), , B
    Printer.Line (5 + MargemLeft, 110 + MargemTOP)-(50 + MargemLeft, 130 + MargemTOP), , B
    Printer.Line (50 + MargemLeft, 110 + MargemTOP)-(100 + MargemLeft, 130 + MargemTOP), , B
    
    
    Next
    
    If xEtiqueta < 4 Then
    xEtiqueta = xEtiqueta + 1
        For xCont = xEtiqueta To 5
            If xCont = 1 Then
            MargemTOP = 0
            MargemLeft = 0
            ElseIf xCont = 2 Then
            MargemTOP = 0
            MargemLeft = 105
            ElseIf xCont = 3 Then
            MargemTOP = 135
            MargemLeft = 0
            ElseIf xCont = 4 Then
            MargemTOP = 135
            MargemLeft = 105
            End If
            
            Printer.CurrentY = 5 + MargemTOP
            Printer.CurrentX = 5 + MargemLeft
            Printer.Font = "ARIAL"
            Printer.FontBold = False
            Printer.FontSize = 330
            Printer.Print "X"
        Next
    End If
    
    Printer.EndDoc
Call LimpaTela(Me)
frmVolLabel.MousePointer = 0
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

Private Sub Text2_Change()

End Sub

Private Sub TxtAwb_GotFocus()
    TxtAWB.SelStart = 0
    TxtAWB.SelLength = 3
End Sub

Private Sub TxtAWB_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
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

Private Sub TxtBuscaLote_GotFocus()
TxtBuscaLote.SelStart = 0
TxtBuscaLote.SelLength = 100
End Sub

Private Sub TxtBuscaLote_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtBuscaLote_LostFocus()
If Val(TxtBuscaLote.Text) > 0 Then

xLote = TxtFilial.Caption & TxtSiglaOrigem.Text & String(10 - Len(Trim(Str(Val(TxtBuscaLote.Text)))), "0") & Trim(Str(Val(TxtBuscaLote.Text)))

    If de_informa.rsVerificaLote.State = 1 Then de_informa.rsVerificaLote.Close
    de_informa.VerificaLote xLote
    
    If de_informa.rsVerificaLote.RecordCount = 1 Then
    TxtLote.Text = de_informa.rsVerificaLote.Fields("lote")
    Else
    TxtLote.Text = ""
    End If
    
End If
End Sub

Private Sub TxtDig_GotFocus()
    TxtAWB.SelStart = 0
    TxtAWB.SelLength = 3
End Sub

Private Sub TxtDig_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
        KeyAscii = 0
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
        If de_informa.rsSel_CiaAereaPorCodigo.State = 1 Then de_informa.rsSel_CiaAereaPorCodigo.Close
        de_informa.Sel_CiaAereaPorCodigo Trim$(TxtSigla.Text)
        If de_informa.rsSel_CiaAereaPorCodigo.RecordCount > 0 Then
            lblFantasiaIns = de_informa.rsSel_CiaAereaPorCodigo.Fields("FANTASIA")
        Else
            MsgBox "Cia. Aérea Não Encontrada!", vbCritical, ""
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

Private Sub TxtVolumes_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Else
        KeyAscii = 0
        End If
    End If
End Sub
