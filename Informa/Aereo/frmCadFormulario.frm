VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCadFormulario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Formulários"
   ClientHeight    =   7485
   ClientLeft      =   3765
   ClientTop       =   630
   ClientWidth     =   5070
   ControlBox      =   0   'False
   Icon            =   "frmCadFormulario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtUFAnt 
      Height          =   285
      Left            =   1980
      TabIndex        =   47
      Top             =   7380
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox TxtDataAnt 
      Height          =   285
      Left            =   180
      TabIndex        =   46
      Top             =   7380
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   3720
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelarForm 
      Caption         =   "Cancelar Formulário"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdInserirForm 
      Caption         =   "Inserir Formulários"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame fraGrid 
      Caption         =   "Formulários AWBs Cadastrados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   120
      TabIndex        =   24
      Top             =   5940
      Width           =   4815
      Begin MSDataGridLib.DataGrid GridFormularios 
         Bindings        =   "frmCadFormulario.frx":000C
         Height          =   1035
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1826
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
         DataMember      =   "Sel_CadFormulario"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "idcadform"
            Caption         =   "idcadform"
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
            DataField       =   "codcia"
            Caption         =   "codcia"
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
            DataField       =   "filial"
            Caption         =   "filial"
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
            DataField       =   "numinicial"
            Caption         =   "numinicial"
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
            DataField       =   "numfinal"
            Caption         =   "numfinal"
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
            DataField       =   "datacadastro"
            Caption         =   "datacadastro"
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
         BeginProperty Column06 
            DataField       =   "usuariocad"
            Caption         =   "usuariocad"
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
               ColumnWidth     =   1289,764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   615,118
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
            BeginProperty Column06 
               ColumnWidth     =   1140,095
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraConsCanc 
      Caption         =   "Consultar Formulário AWB"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   25
      Top             =   3480
      Width           =   4815
      Begin VB.TextBox TxtDig 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         MaxLength       =   12
         TabIndex        =   45
         Top             =   960
         Width           =   315
      End
      Begin VB.CommandButton cmdDesisteCanc 
         Caption         =   "Sair"
         Height          =   315
         Left            =   2760
         TabIndex        =   18
         Top             =   1605
         Width           =   1935
      End
      Begin VB.CommandButton cmdConfirmaCanc 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2760
         TabIndex        =   17
         Top             =   1305
         Width           =   1935
      End
      Begin VB.TextBox TxtMotivo 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1980
         Width           =   3495
      End
      Begin VB.TextBox TxtBuscaFilial2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Buscar AWB"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   15
         Top             =   945
         Width           =   1935
      End
      Begin VB.TextBox txtNumAWB 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1140
         MaxLength       =   12
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtCodCia 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Motivo:"
         Height          =   195
         Left            =   540
         TabIndex        =   43
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label TxtFilial2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         TabIndex        =   42
         Top             =   240
         Width           =   495
      End
      Begin VB.Label TxtNomeFilial2 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   41
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   585
         TabIndex        =   40
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Data Status:"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   1725
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Left            =   570
         TabIndex        =   29
         Top             =   1425
         Width           =   495
      End
      Begin VB.Label lblFantasia 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1500
         TabIndex        =   28
         Top             =   540
         Width           =   3195
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Núm. AWB:"
         Height          =   195
         Left            =   225
         TabIndex        =   27
         Top             =   1005
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cód. Cia:"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   585
         Width           =   645
      End
      Begin VB.Label lblDataStatus 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1140
         TabIndex        =   32
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1140
         TabIndex        =   31
         Top             =   1380
         Width           =   1575
      End
   End
   Begin VB.Frame fraInserir 
      Caption         =   "Inserir Formulários"
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
      Height          =   2835
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   4815
      Begin VB.TextBox TxtBuscaFilial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdConfirmaIns 
         Caption         =   "Confirmar"
         Height          =   315
         Left            =   2400
         TabIndex        =   9
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         Caption         =   "Numeração Sequencial do Formulário"
         Height          =   1455
         Left            =   120
         TabIndex        =   33
         Top             =   900
         Width           =   4575
         Begin VB.TextBox TxtAWB2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   5
            Top             =   540
            Width           =   1215
         End
         Begin VB.TextBox TxtDig2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   2940
            MaxLength       =   1
            TabIndex        =   6
            Top             =   540
            Width           =   315
         End
         Begin VB.TextBox TxtDigFim 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   2940
            MaxLength       =   1
            TabIndex        =   8
            Top             =   840
            Width           =   315
         End
         Begin VB.TextBox TxtDigInic 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   2940
            MaxLength       =   1
            TabIndex        =   4
            Top             =   240
            Width           =   315
         End
         Begin MSComctlLib.ProgressBar progress 
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1140
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.TextBox txtNumFim 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   7
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtNumInic 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "2º AWB:"
            Height          =   195
            Left            =   1020
            TabIndex        =   44
            Top             =   585
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "AWB Final:"
            Height          =   195
            Left            =   840
            TabIndex        =   35
            Top             =   885
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "AWB Inicial:"
            Height          =   195
            Left            =   765
            TabIndex        =   34
            Top             =   285
            Width           =   870
         End
      End
      Begin VB.CommandButton cmdDesisteIns 
         Caption         =   "Sair"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txtCodCiaIns 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   3
         TabIndex        =   2
         Top             =   540
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Filial"
         Height          =   195
         Left            =   585
         TabIndex        =   39
         Top             =   285
         Width           =   300
      End
      Begin VB.Label TxtNomeFilial 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   38
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label TxtFilial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         TabIndex        =   37
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cód. Cia:"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   645
      End
      Begin VB.Label lblFantasiaIns 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1500
         TabIndex        =   22
         Top             =   540
         Width           =   3195
      End
   End
End
Attribute VB_Name = "frmCadFormulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBusca_Click()
    If de_informa.rsSel_NumFormularioItem.State = 1 Then de_informa.rsSel_NumFormularioItem.Close
    de_informa.Sel_NumFormularioItem Trim$(txtCodCia), Val(txtNumAWB), TxtFilial2.Caption
    lblStatus = ""
    lblDataStatus = ""
    TxtMotivo.Text = ""
    If de_informa.rsSel_NumFormularioItem.RecordCount > 0 Then
        If de_informa.rsSel_NumFormularioItem.Fields("tem_ocorr") = "C" Then
        lblStatus = "Cancelado"
        cmdConfirmaCanc.Caption = "Descancelar"
        TxtMotivo.Enabled = False
        ElseIf de_informa.rsSel_NumFormularioItem.Fields("tem_ocorr") = "D" Then
        lblStatus = "Disponível"
        cmdConfirmaCanc.Caption = "Cancelar"
        TxtMotivo.Enabled = True
        TxtMotivo.SetFocus
        ElseIf de_informa.rsSel_NumFormularioItem.Fields("tem_ocorr") = "2" Then
        lblStatus = "Em Ocorrência"
        cmdConfirmaCanc.Caption = "Cancelar"
        TxtMotivo.Enabled = False
        Else
        lblStatus = "Emitido"
        cmdConfirmaCanc.Caption = "Cancelar"
        TxtMotivo.Enabled = True
        TxtMotivo.SetFocus
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
    cmdConfirmaCanc.Enabled = True
    'cmdConfirmaCanc.SetFocus
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
    If Len(Trim$(lblStatus)) > 0 Then
        cmdConfirmaCanc.Enabled = True
    Else
        cmdConfirmaCanc.Enabled = False
    End If
    TxtBuscaFilial2.SetFocus
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

    'If Mid(lblStatus, 1, 1) = "C" Then
        'MsgBox "ERRO! Este Formulário Já Está Cancelado!", vbCritical, ""
        'txtNumAWB.SetFocus
        'Exit Sub
    'Else
        'If lblStatus = "Em Ocorrência" Then
        'MsgBox "Não é possível Cancelar este AWB pois ele está em processo de Ocorrência!"
        'Exit Sub
        'End If
    
    CodAwb = TxtFilial2.Caption & txtCodCia.Text & String(10 - Len(Trim(Str(Val(txtNumAWB.Text)))), "0") & Trim(Str(Val(txtNumAWB.Text))) & Trim(Str(Val(TxtDig.Text)))
    
    If cmdConfirmaCanc.Caption = "Cancelar" Then
    de_informa.cn_informa.BeginTrans
    de_informa.Alt_FormularioStatuscanc "C", xUsuario, UCase(Trim(TxtMotivo.Text)), txtCodCia, txtNumAWB, TxtFilial2.Caption
        If de_informa.rsConsultaAWB.State = 1 Then de_informa.rsConsultaAWB.Close
        de_informa.ConsultaAWB CodAwb
        
            If de_informa.rsConsultaAWB.RecordCount = 0 Then
            Dim xValm As Currency
            Dim xPesoRe As Currency
            Dim xPesoCub As Currency
            xValm = 0
            xPesoRe = 0
            xPesoCub = 0
            DoEvents
            frmCadFormularioDataCanc.Show 1
            DoEvents
            de_informa.InsAWB CodAwb, Trim(Str(Val(txtNumAWB.Text))), Trim(Str(Val(TxtDig.Text))), _
            UCase(Trim(txtCodCia.Text)), "", "", "", _
            Trim(TxtFilial2.Caption), "", "", _
            "", Val(0), xValm, Val(0), Val(0), Val(0), Val(0), xPesoRe, xPesoCub, _
            "", "", "", "", "", "", "", UCase(Trim(TxtUFAnt.Text)), "", "", "", "", _
            "", "", "", "", "", "", "", "", "", "", "", "", _
            "", "", "", "", "", "", "", "", "", "", "", ""
            'Segunda Parte do Insert do AWB
            de_informa.InsAWB2 CodAwb, "", CDbl(0), CDbl(0), CDbl(0), CDbl(0), Val(0), CDbl(0), CDbl(0), CDbl(0), CDbl(0), CDbl(0), CDbl(0), CDbl(0), "", CDbl(0), "", CDbl(0), _
            CDbl(0), CDbl(0), CDbl(0), "", CDbl(0), "", "", "", "", "", "", "", "CANCELADO ANTES DE EMITIR", CDate(TxtDataAnt.Text), CVar(DataHora("Hora")), xUsuario, "SAO"
            TxtDataAnt.Text = ""
            TxtUFAnt.Text = ""
            End If
    de_informa.CancelaAWB xUsuario, UCase(Trim(TxtMotivo.Text)), DataHora("DATA"), DataHora("HORA"), CodAwb
    de_informa.cn_informa.CommitTrans
    MsgBox "OK! Formulário Cancelado.", vbInformation, ""
    cmdBusca_Click
    ElseIf cmdConfirmaCanc.Caption = "Descancelar" Then
    de_informa.cn_informa.BeginTrans
    If de_informa.rsConsultaAWB.State = 1 Then de_informa.rsConsultaAWB.Close
    de_informa.ConsultaAWB CodAwb
        If de_informa.rsConsultaAWB.RecordCount > 0 Then
        de_informa.Alt_FormularioStatuscanc "E", xUsuario, UCase(Trim(TxtMotivo.Text)), txtCodCia, txtNumAWB, TxtFilial2.Caption
        Else
        de_informa.Alt_FormularioStatuscanc "D", xUsuario, UCase(Trim(TxtMotivo.Text)), txtCodCia, txtNumAWB, TxtFilial2.Caption
        End If
    de_informa.DescancelaAWB "", "", "", CDate(DataHora("DATA")), "", CodAwb
    de_informa.cn_informa.CommitTrans
    MsgBox "OK! Formulário Descancelado.", vbInformation, ""
    cmdBusca_Click
    End If
    
'txtCodCia.Text = ""
'lblFantasia.Caption = ""
txtNumAWB.Text = ""
TxtDig.Text = ""
lblStatus.Caption = ""
lblDataStatus.Caption = ""
TxtBuscaFilial2.Text = ""
'TxtFilial2.Caption = ""
'TxtNomeFilial2.Caption = ""
TxtMotivo.Text = ""
TxtMotivo.Enabled = False
txtNumAWB.SetFocus
End Sub

Private Sub cmdConfirmaIns_Click()
    Dim MatrizForm(100000, 2) As Long
    Dim xForm As Long
    Dim xDig As Integer
    Dim Linha As Long
    Dim Linha2 As Long
    Dim xZero As Integer
    
    
    Dim xnumform As Long
    If Val(txtNumInic) > Val(txtNumFim) Then
        MsgBox "ERRO! Número Final MENOR que Número Inicial.", vbCritical, ""
        txtNumInic.SetFocus
        Exit Sub
    ElseIf Val(txtNumInic.Text) = 0 Then
    MsgBox "Não é possível cadastrar um formulário com valor nulo.", vbCritical, ""
    txtNumInic.SetFocus
    Exit Sub
    ElseIf Val(TxtAWB2.Text) = 0 Then
    MsgBox "Não é possível cadastrar um formulário com valor nulo.", vbCritical, ""
    TxtAWB2.SetFocus
    Exit Sub
    ElseIf Val(txtNumFim.Text) = 0 Then
    MsgBox "Não é possível cadastrar um formulário com valor nulo.", vbCritical, ""
    txtNumFim.SetFocus
    Exit Sub
    ElseIf Len(Trim(TxtDigInic.Text)) = 0 Then
    MsgBox "É necessário o dígito para cadastrar o formulário.", vbCritical, ""
    TxtDigInic.SetFocus
    Exit Sub
    ElseIf Len(Trim(TxtDig2.Text)) = 0 Then
    MsgBox "É necessário o dígito para cadastrar o formulário.", vbCritical, ""
    TxtDig2.SetFocus
    Exit Sub
    ElseIf Len(Trim(TxtDigFim.Text)) = 0 Then
    MsgBox "É necessário o dígito para cadastrar o formulário.", vbCritical, ""
    TxtDigFim.SetFocus
    Exit Sub
    ElseIf Val(TxtFilial.Caption) = 0 Then
    MsgBox "É necessário informar a qual filial este Intervalo de Formulários pertencerá!", vbExclamation, ""
    Exit Sub
    End If
    
    'verifica se intervalo já existe
    If de_informa.rsSel_FormularioJaExiste.State = 1 Then de_informa.rsSel_FormularioJaExiste.Close
    de_informa.Sel_FormularioJaExiste txtNumInic, txtNumFim, Trim$(txtCodCiaIns), TxtFilial.Caption
    If de_informa.rsSel_FormularioJaExiste.RecordCount > 0 Then
        MsgBox "Intervalo de Numeração de Formulário Já Existente no Banco de Dados!", vbCritical, ""
        Exit Sub
    End If
    
        If txtCodCiaIns.Text = "OC" Then
            If (Val(TxtDig2.Text) - Val(TxtDigInic.Text)) > 1 And Val(TxtDigInic.Text) <> 9 Then
            MsgBox "O controle dos fomulários não está correto. Favor verificar os números digitados e tente novamente.", vbInformation, ""
            Exit Sub
            Else
                If Val(TxtDig2.Text) = 0 And Val(TxtDigInic.Text) = 0 Then
                Linha = 0
                xDig = Val(TxtDigInic.Text) - 1
                xZero = 1
                    For xForm = Val(txtNumInic.Text) To Val(txtNumFim.Text)
                    Linha = Linha + 1
                    xDig = xDig + 1
                        If xZero = 1 Then
                        xDig = 0
                        xZero = 2
                        ElseIf xZero = 2 Then
                        xDig = 0
                        xZero = 0
                        End If
                        
                        If xDig > 9 Then
                        xDig = 0
                        xZero = 1
                        End If
                    MatrizForm(Linha, 0) = xForm
                    MatrizForm(Linha, 1) = xDig
                    Next
                ElseIf Val(TxtDig2.Text) > 0 And Val(TxtDigInic.Text) = 0 Then
                Linha = 0
                xDig = Val(TxtDigInic.Text) - 1
                xZero = 2
                    For xForm = Val(txtNumInic.Text) To Val(txtNumFim.Text)
                    Linha = Linha + 1
                    xDig = xDig + 1
                        If xZero = 1 Then
                        xDig = 0
                        xZero = 2
                        ElseIf xZero = 2 Then
                        xDig = 0
                        xZero = 0
                        End If
                        
                        If xDig > 9 Then
                        xDig = 0
                        xZero = 1
                        End If
                    MatrizForm(Linha, 0) = xForm
                    MatrizForm(Linha, 1) = xDig
                    Next
                Else
                Linha = 0
                xDig = Val(TxtDigInic.Text) - 1
                xZero = 0
                    For xForm = Val(txtNumInic.Text) To Val(txtNumFim.Text)
                    Linha = Linha + 1
                    xDig = xDig + 1
                        If xZero = 1 Then
                        xDig = 0
                        xZero = 0
                        End If
                        
                        If xDig > 9 Then
                        xDig = 0
                        xZero = 1
                        End If
                    MatrizForm(Linha, 0) = xForm
                    MatrizForm(Linha, 1) = xDig
                    Next
                End If
            End If
        Else
        Linha = 0
        xDig = Val(TxtDigInic.Text) - 1
        
            For xForm = Val(txtNumInic.Text) To Val(txtNumFim.Text)
            Linha = Linha + 1
            xDig = xDig + 1
                If xDig > 6 Then
                xDig = 0
                End If
            MatrizForm(Linha, 0) = xForm
            MatrizForm(Linha, 1) = xDig
            Next
        End If
    
        If xDig <> Val(TxtDigFim.Text) Then
        MsgBox "O controle dos fomulários não está correto. Favor verificar os números digitados e tente novamente.", vbInformation, ""
        Exit Sub
        End If
    
    
    fraInserir.Enabled = False
    cmdConfirmaIns.Enabled = False
    cmdDesisteIns.Enabled = False
    fraGrid.Enabled = False
    frmCadFormulario.MousePointer = 11
    
    de_informa.cn_informa.BeginTrans
    
    'cadastrando formulario (pai)
    de_informa.Ins_CadFormulario Trim$(txtCodCiaIns), Val(txtNumInic), Val(TxtDigInic.Text), Val(txtNumFim), Val(TxtDigFim.Text), xUsuario, TxtFilial.Caption
    
    'busca o ID do cadastramento de formulário que acabou de ser feito acima
    If de_informa.rsSel_BuscaIDFormAir.State = 1 Then de_informa.rsSel_BuscaIDFormAir.Close
    de_informa.Sel_BuscaIDFormAir Trim$(txtCodCiaIns), Val(txtNumInic), Val(TxtDigInic.Text), Val(txtNumFim), Val(TxtDigFim.Text), TxtFilial.Caption
    If de_informa.rsSel_BuscaIDFormAir.RecordCount <= 0 Then
        MsgBox "Impossível Cadastrar ! Procure Suporte o Gestor do Sistema.", vbCritical, ""
        Exit Sub
    End If
    
    'cadastrando formulário ítem
    Linha2 = 0
    progress.Max = Linha
    For Linha2 = 1 To Linha
        xAWB = MatrizForm(Linha2, 0)
        xDig = MatrizForm(Linha2, 1)
        xCodAwb = String(2 - Len(Trim(Str(Val(TxtFilial.Caption)))), "0") & Trim(Str(Val(TxtFilial.Caption))) & UCase(Trim(txtCodCiaIns.Text)) & String(10 - Len(Trim(Str(Val(xAWB)))), "0") & Trim(Str(Val(xAWB))) & Trim(Str(Val(xDig)))

        de_informa.Ins_CadFormItem de_informa.rsSel_BuscaIDFormAir.Fields("idcadform"), xCodAwb, MatrizForm(Linha2, 0), MatrizForm(Linha2, 1), Trim$(txtCodCiaIns), "D", TxtFilial.Caption
        progress.Value = Linha2
        DoEvents
    Next
    de_informa.cn_informa.CommitTrans
    cmdConfirmaIns.Enabled = True
    cmdDesisteIns.Enabled = True
    fraGrid.Enabled = True
    fraInserir.Enabled = True
    'txtCodCiaIns.Text = ""
    'lblFantasiaIns.Caption = ""
    progress.Value = 0
    txtNumInic.Text = ""
    txtNumFim.Text = ""
    'TxtBuscaFilial.Text = ""
    'TxtFilial.Caption = ""
    'TxtNomeFilial.Caption = ""
    TxtDigInic.Text = ""
    TxtDig2.Text = ""
    TxtDigFim.Text = ""
    TxtAWB2.Text = ""
    DoEvents
    
    If de_informa.rsSel_CadFormulario.State = 1 Then de_informa.rsSel_CadFormulario.Close
    de_informa.Sel_CadFormulario
    GridFormularios.DataMember = "sel_cadformulario"
    GridFormularios.Refresh
    Me.MousePointer = 0
    MsgBox "OK! Formulários Cadastrados com sucesso!", vbInformation, "Cadastramento de Formulários Concluído"
    txtNumInic.SetFocus
    
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
TxtDig.Text = ""
lblStatus.Caption = ""
lblDataStatus.Caption = ""
TxtBuscaFilial2.Text = ""
TxtFilial2.Caption = ""
TxtNomeFilial2.Caption = ""

End Sub
Private Sub cmdDesisteIns_Click()
    fraConsCanc.Enabled = False
    fraInserir.Enabled = False
    cmdCancelarForm.Enabled = True
    cmdInserirForm.Enabled = True
    cmdSair.Enabled = True
    
    txtCodCiaIns.Text = ""
    lblFantasiaIns.Caption = ""
    progress.Value = 0
    txtNumInic.Text = ""
    txtNumFim.Text = ""
    TxtBuscaFilial.Text = ""
    TxtFilial.Caption = ""
    TxtNomeFilial.Caption = ""
    DoEvents
    
End Sub

Private Sub cmdInserirForm_Click()
    fraConsCanc.Enabled = False
    fraInserir.Enabled = True
    cmdCancelarForm.Enabled = False
    cmdInserirForm.Enabled = False
    cmdSair.Enabled = False
    TxtBuscaFilial.SetFocus
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
        Call TxtBuscaFilial_LostFocus
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
                cmdBusca.Enabled = True
            End If
        Else
            MsgBox "Código de Cia. Aérea Não Encontrado !"
            cmdBusca.Enabled = False
            txtCodCia.SetFocus
            Exit Sub
        End If
    Else
        cmdBusca.Enabled = False
        cmdConfirmaCanc.Enabled = False
        lblFantasia = ""
        lblStatus = ""
        txtNumAWB = ""
    End If
End Sub

Private Sub txtCodCiaIns_GotFocus()
    txtCodCiaIns.SelStart = 0
    txtCodCiaIns.SelLength = 3
End Sub

Private Sub txtCodCiaIns_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call txtCodCia_LostFocus
    KeyAscii = 0
    SendKeys "{TAB}"
    End If
End Sub

Private Sub txtCodCiaIns_LostFocus()
    If Len(Trim$(txtCodCiaIns)) > 0 Then
        txtCodCiaIns = UCase(Trim$(txtCodCiaIns))
        lblFantasiaIns = ""
        If de_informa.rsSel_CiaAereaPorCodigo.State = 1 Then de_informa.rsSel_CiaAereaPorCodigo.Close
        de_informa.Sel_CiaAereaPorCodigo Trim$(txtCodCiaIns.Text)
        If de_informa.rsSel_CiaAereaPorCodigo.RecordCount > 0 Then
            lblFantasiaIns = de_informa.rsSel_CiaAereaPorCodigo.Fields("fantasia")
            If Val(txtNumAWB) > 0 Then
                cmdBusca.Enabled = True
            End If
        Else
            MsgBox "Código de Cia. Aérea Não Encontrado !", vbCritical, ""
            cmdBusca.Enabled = False
            txtCodCiaIns.SetFocus
            Exit Sub
        End If
    Else
        cmdBusca.Enabled = False
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
        cmdBusca.Enabled = True
        'cmdConfirmaCanc.Enabled = True
    Else
        cmdBusca.Enabled = False
        cmdConfirmaCanc.Enabled = False
        lblStatus = ""
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

Private Sub txtNumAWB_LostFocus()
If Len(Trim(TxtFilial2.Caption)) > 0 And Len(Trim(txtCodCia.Text)) > 0 And Len(Trim(txtNumAWB.Text)) > 0 Then

If de_informa.rsConfereNumeroAWB.State = 1 Then de_informa.rsConfereNumeroAWB.Close
de_informa.ConfereNumeroAWB txtCodCia.Text, TxtFilial2.Caption, txtNumAWB.Text

    If de_informa.rsConfereNumeroAWB.RecordCount = 0 Then
    MsgBox "Este formulário não está cadastrado!.", vbCritical, ""
    txtNumAWB.Text = ""
    TxtDig.Text = ""
    txtNumAWB.SetFocus
    Exit Sub
    Else
    TxtDig.Text = de_informa.rsConfereNumeroAWB.Fields("dig")
    End If
Else
txtNumAWB.Text = ""
TxtDig.Text = ""
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
