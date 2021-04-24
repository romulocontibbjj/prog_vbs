VERSION 5.00
Begin VB.Form frmNovaTabPrz 
   Caption         =   "Inclusão de Nova Tabela de Prazos"
   ClientHeight    =   5415
   ClientLeft      =   2865
   ClientTop       =   1290
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   5760
   Begin VB.Frame Frame4 
      Caption         =   "Processo de Inclusão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   42
      Top             =   120
      Width           =   2415
      Begin VB.Label lblEtapaFim 
         AutoSize        =   -1  'True
         Caption         =   "de  27"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   45
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lblEtapa 
         AutoSize        =   -1  'True
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   44
         Top             =   480
         Width           =   225
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Etapa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   43
         Top             =   480
         Width           =   570
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Sair"
      Height          =   495
      Left            =   2880
      TabIndex        =   41
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar Prazo"
      Height          =   495
      Left            =   4320
      TabIndex        =   40
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dados da Tabela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   5535
      Begin VB.TextBox txtAirInterior 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3480
         TabIndex        =   39
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtRodoInterior 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3480
         TabIndex        =   38
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtAirCapital 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   37
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtRodoCapital 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   36
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Interior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3720
         TabIndex        =   35
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Capital"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   34
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Modal Aéreo........:"
         Height          =   195
         Left            =   600
         TabIndex        =   33
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Modal Rodoviário:"
         Height          =   195
         Left            =   600
         TabIndex        =   32
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label lblRegUF 
         Alignment       =   2  'Center
         Caption         =   "Norte - AC (Acre)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   31
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cod. Tabela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      Begin VB.TextBox txtCodTabela 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5535
      Begin VB.Label lblMT 
         AutoSize        =   -1  'True
         Caption         =   "MT"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   4920
         TabIndex        =   30
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label lblAP 
         AutoSize        =   -1  'True
         Caption         =   "AP"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   1320
         TabIndex        =   29
         Top             =   360
         Width           =   210
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "PA"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   1920
         TabIndex        =   28
         Top             =   360
         Width           =   210
      End
      Begin VB.Label lblRO 
         AutoSize        =   -1  'True
         Caption         =   "RO"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   2520
         TabIndex        =   27
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblRR 
         AutoSize        =   -1  'True
         Caption         =   "RR"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   3120
         TabIndex        =   26
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblTO 
         AutoSize        =   -1  'True
         Caption         =   "TO"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   3720
         TabIndex        =   25
         Top             =   360
         Width           =   225
      End
      Begin VB.Label lblAL 
         AutoSize        =   -1  'True
         Caption         =   "AL"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   4320
         TabIndex        =   24
         Top             =   360
         Width           =   195
      End
      Begin VB.Label lblBA 
         AutoSize        =   -1  'True
         Caption         =   "BA"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   4920
         TabIndex        =   23
         Top             =   360
         Width           =   210
      End
      Begin VB.Label lblSE 
         AutoSize        =   -1  'True
         Caption         =   "SE"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   210
      End
      Begin VB.Label lblPE 
         AutoSize        =   -1  'True
         Caption         =   "PE"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   720
         TabIndex        =   21
         Top             =   720
         Width           =   210
      End
      Begin VB.Label lblPB 
         AutoSize        =   -1  'True
         Caption         =   "PB"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   1320
         TabIndex        =   20
         Top             =   720
         Width           =   210
      End
      Begin VB.Label lblRN 
         AutoSize        =   -1  'True
         Caption         =   "RN"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   1920
         TabIndex        =   19
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lblCE 
         AutoSize        =   -1  'True
         Caption         =   "CE"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   2520
         TabIndex        =   18
         Top             =   720
         Width           =   210
      End
      Begin VB.Label lblPI 
         AutoSize        =   -1  'True
         Caption         =   "PI"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   3120
         TabIndex        =   17
         Top             =   720
         Width           =   150
      End
      Begin VB.Label lblMA 
         AutoSize        =   -1  'True
         Caption         =   "MA"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   3720
         TabIndex        =   16
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lblES 
         AutoSize        =   -1  'True
         Caption         =   "ES"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   4320
         TabIndex        =   15
         Top             =   720
         Width           =   210
      End
      Begin VB.Label lblMG 
         AutoSize        =   -1  'True
         Caption         =   "MG"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   4920
         TabIndex        =   14
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblRJ 
         AutoSize        =   -1  'True
         Caption         =   "RJ"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   195
      End
      Begin VB.Label lblSP 
         AutoSize        =   -1  'True
         Caption         =   "SP"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   720
         TabIndex        =   12
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label lblPR 
         AutoSize        =   -1  'True
         Caption         =   "PR"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         Top             =   1080
         Width           =   225
      End
      Begin VB.Label lblRS 
         AutoSize        =   -1  'True
         Caption         =   "RS"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   1920
         TabIndex        =   10
         Top             =   1080
         Width           =   225
      End
      Begin VB.Label lblSC 
         AutoSize        =   -1  'True
         Caption         =   "SC"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   2520
         TabIndex        =   9
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label lblDF 
         AutoSize        =   -1  'True
         Caption         =   "DF"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   3120
         TabIndex        =   8
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label lblGO 
         AutoSize        =   -1  'True
         Caption         =   "GO"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   3720
         TabIndex        =   7
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label lblMS 
         AutoSize        =   -1  'True
         Caption         =   "MS"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   4320
         TabIndex        =   6
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label lblAM 
         AutoSize        =   -1  'True
         Caption         =   "AM"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   240
      End
      Begin VB.Label lblAC 
         AutoSize        =   -1  'True
         Caption         =   "AC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmNovaTabPrz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    If cmdCancelar.Caption <> "Sair" Then
        If MsgBox("Se você cancelar esta operação, serão desconsiderados todas as inclusões para esta Tabela de Prazos. Você tem certeza que deseja cancelar ?", vbYesNo + vbQuestion, "Confirmação") = vbYes Then
            de_informa.excl_tabprazo txtCodTabela
            MsgBox "OK ! Processo cancelado !"
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdGravar_Click()
    If Val(lblEtapa) = 1 Then
        de_informa.ins_cadprazo txtCodTabela, "AC", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "AC", "A", txtAirCapital, txtAirInterior
        lblEtapa = "02"
        lblAC.ForeColor = &HC00000
        lblAM.Font.Size = 12
        lblAM.Font.Bold = True
        lblAM.ForeColor = &H80000017
        lblRegUF = "Norte - AM (Amazonas)"
        txtCodTabela.BackColor = &H80000014
        txtCodTabela.Enabled = False
        cmdCancelar.Caption = "Cancelar Tudo"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 2 Then
        de_informa.ins_cadprazo txtCodTabela, "AM", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "AM", "A", txtAirCapital, txtAirInterior
        lblEtapa = "03"
        lblAM.ForeColor = &HC00000
        lblAP.Font.Size = 12
        lblAP.Font.Bold = True
        lblAP.ForeColor = &H80000017
        lblRegUF = "Norte - AP (Amapá)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 3 Then
        de_informa.ins_cadprazo txtCodTabela, "AP", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "AP", "A", txtAirCapital, txtAirInterior
        lblEtapa = "04"
        lblAP.ForeColor = &HC00000
        lblPA.Font.Size = 12
        lblPA.Font.Bold = True
        lblPA.ForeColor = &H80000017
        lblRegUF = "Norte - PA (Pará)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 4 Then
        de_informa.ins_cadprazo txtCodTabela, "PA", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "PA", "A", txtAirCapital, txtAirInterior
        lblEtapa = "05"
        lblPA.ForeColor = &HC00000
        lblRO.Font.Size = 12
        lblRO.Font.Bold = True
        lblRO.ForeColor = &H80000017
        lblRegUF = "Norte - RO (Rondônia)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 5 Then
        de_informa.ins_cadprazo txtCodTabela, "RO", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "RO", "A", txtAirCapital, txtAirInterior
        lblEtapa = "06"
        lblRO.ForeColor = &HC00000
        lblRR.Font.Size = 12
        lblRR.Font.Bold = True
        lblRR.ForeColor = &H80000017
        lblRegUF = "Norte - RR (Roraima)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 6 Then
        de_informa.ins_cadprazo txtCodTabela, "RR", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "RR", "A", txtAirCapital, txtAirInterior
        lblEtapa = "07"
        lblRR.ForeColor = &HC00000
        lblTO.Font.Size = 12
        lblTO.Font.Bold = True
        lblTO.ForeColor = &H80000017
        lblRegUF = "Norte - TO (Tocantins)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 7 Then
        de_informa.ins_cadprazo txtCodTabela, "TO", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "TO", "A", txtAirCapital, txtAirInterior
        lblEtapa = "08"
        lblTO.ForeColor = &HC00000
        lblAL.Font.Size = 12
        lblAL.Font.Bold = True
        lblAL.ForeColor = &H80000017
        lblRegUF = "Nordeste - AL (Alagoas)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 8 Then
        de_informa.ins_cadprazo txtCodTabela, "AL", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "AL", "A", txtAirCapital, txtAirInterior
        lblEtapa = "09"
        lblAL.ForeColor = &HC00000
        lblBA.Font.Size = 12
        lblBA.Font.Bold = True
        lblBA.ForeColor = &H80000017
        lblRegUF = "Nordeste - BA (Bahia)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 9 Then
        de_informa.ins_cadprazo txtCodTabela, "BA", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "BA", "A", txtAirCapital, txtAirInterior
        lblEtapa = "10"
        lblBA.ForeColor = &HC00000
        lblSE.Font.Size = 12
        lblSE.Font.Bold = True
        lblSE.ForeColor = &H80000017
        lblRegUF = "Nordeste - SE (Sergipe)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 10 Then
        de_informa.ins_cadprazo txtCodTabela, "SE", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "SE", "A", txtAirCapital, txtAirInterior
        lblEtapa = "11"
        lblSE.ForeColor = &HC00000
        lblPE.Font.Size = 12
        lblPE.Font.Bold = True
        lblPE.ForeColor = &H80000017
        lblRegUF = "Nordeste - PE (Pernambuco)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 11 Then
        de_informa.ins_cadprazo txtCodTabela, "PE", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "PE", "A", txtAirCapital, txtAirInterior
        lblEtapa = "12"
        lblPE.ForeColor = &HC00000
        lblPB.Font.Size = 12
        lblPB.Font.Bold = True
        lblPB.ForeColor = &H80000017
        lblRegUF = "Nordeste - PB (Paraíba)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 12 Then
        de_informa.ins_cadprazo txtCodTabela, "PB", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "PB", "A", txtAirCapital, txtAirInterior
        lblEtapa = "13"
        lblPB.ForeColor = &HC00000
        lblRN.Font.Size = 12
        lblRN.Font.Bold = True
        lblRN.ForeColor = &H80000017
        lblRegUF = "Nordeste - RN (Rio Grande do Norte)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 13 Then
        de_informa.ins_cadprazo txtCodTabela, "RN", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "RN", "A", txtAirCapital, txtAirInterior
        lblEtapa = "14"
        lblRN.ForeColor = &HC00000
        lblCE.Font.Size = 12
        lblCE.Font.Bold = True
        lblCE.ForeColor = &H80000017
        lblRegUF = "Nordeste - CE (Ceará)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 14 Then
        de_informa.ins_cadprazo txtCodTabela, "CE", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "CE", "A", txtAirCapital, txtAirInterior
        lblEtapa = "15"
        lblCE.ForeColor = &HC00000
        lblPI.Font.Size = 12
        lblPI.Font.Bold = True
        lblPI.ForeColor = &H80000017
        lblRegUF = "Nordeste - PI (Piauí)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 15 Then
        de_informa.ins_cadprazo txtCodTabela, "PI", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "PI", "A", txtAirCapital, txtAirInterior
        lblEtapa = "16"
        lblPI.ForeColor = &HC00000
        lblMA.Font.Size = 12
        lblMA.Font.Bold = True
        lblMA.ForeColor = &H80000017
        lblRegUF = "Nordeste - MA (Maranhão)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 16 Then
        de_informa.ins_cadprazo txtCodTabela, "MA", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "MA", "A", txtAirCapital, txtAirInterior
        lblEtapa = "17"
        lblMA.ForeColor = &HC00000
        lblES.Font.Size = 12
        lblES.Font.Bold = True
        lblES.ForeColor = &H80000017
        lblRegUF = "Sudeste - ES (Espírito Santo)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 17 Then
        de_informa.ins_cadprazo txtCodTabela, "ES", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "ES", "A", txtAirCapital, txtAirInterior
        lblEtapa = "18"
        lblES.ForeColor = &HC00000
        lblMG.Font.Size = 12
        lblMG.Font.Bold = True
        lblMG.ForeColor = &H80000017
        lblRegUF = "Sudeste - MG (Minas Gerais)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 18 Then
        de_informa.ins_cadprazo txtCodTabela, "MG", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "MG", "A", txtAirCapital, txtAirInterior
        lblEtapa = "19"
        lblMG.ForeColor = &HC00000
        lblRJ.Font.Size = 12
        lblRJ.Font.Bold = True
        lblRJ.ForeColor = &H80000017
        lblRegUF = "Sudeste - RJ (Rio de Janeiro)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 19 Then
        de_informa.ins_cadprazo txtCodTabela, "RJ", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "RJ", "A", txtAirCapital, txtAirInterior
        lblEtapa = "20"
        lblRJ.ForeColor = &HC00000
        lblSP.Font.Size = 12
        lblSP.Font.Bold = True
        lblSP.ForeColor = &H80000017
        lblRegUF = "Sudeste - SP (São Paulo)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 20 Then
        de_informa.ins_cadprazo txtCodTabela, "SP", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "SP", "A", txtAirCapital, txtAirInterior
        lblEtapa = "21"
        lblSP.ForeColor = &HC00000
        lblPR.Font.Size = 12
        lblPR.Font.Bold = True
        lblPR.ForeColor = &H80000017
        lblRegUF = "Sul - PR (Paraná)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 21 Then
        de_informa.ins_cadprazo txtCodTabela, "PR", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "PR", "A", txtAirCapital, txtAirInterior
        lblEtapa = "22"
        lblPR.ForeColor = &HC00000
        lblRS.Font.Size = 12
        lblRS.Font.Bold = True
        lblRS.ForeColor = &H80000017
        lblRegUF = "Sul - RS (Rio Grande do Sul)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 22 Then
        de_informa.ins_cadprazo txtCodTabela, "RS", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "RS", "A", txtAirCapital, txtAirInterior
        lblEtapa = "23"
        lblRS.ForeColor = &HC00000
        lblSC.Font.Size = 12
        lblSC.Font.Bold = True
        lblSC.ForeColor = &H80000017
        lblRegUF = "Sul - SC (Santa Catarina)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 23 Then
        de_informa.ins_cadprazo txtCodTabela, "SC", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "SC", "A", txtAirCapital, txtAirInterior
        lblEtapa = "24"
        lblSC.ForeColor = &HC00000
        lblDF.Font.Size = 12
        lblDF.Font.Bold = True
        lblDF.ForeColor = &H80000017
        lblRegUF = "Centro Oeste - DF (Brasília)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 24 Then
        de_informa.ins_cadprazo txtCodTabela, "DF", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "DF", "A", txtAirCapital, txtAirInterior
        lblEtapa = "25"
        lblDF.ForeColor = &HC00000
        lblGO.Font.Size = 12
        lblGO.Font.Bold = True
        lblGO.ForeColor = &H80000017
        lblRegUF = "Centro Oeste - GO (Goiás)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 25 Then
        de_informa.ins_cadprazo txtCodTabela, "GO", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "GO", "A", txtAirCapital, txtAirInterior
        lblEtapa = "26"
        lblGO.ForeColor = &HC00000
        lblMS.Font.Size = 12
        lblMS.Font.Bold = True
        lblMS.ForeColor = &H80000017
        lblRegUF = "Centro Oeste - MS (Mato Grosso do Sul)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 26 Then
        de_informa.ins_cadprazo txtCodTabela, "MS", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "MS", "A", txtAirCapital, txtAirInterior
        lblEtapa = "27"
        lblMS.ForeColor = &HC00000
        lblMT.Font.Size = 12
        lblMT.Font.Bold = True
        lblMT.ForeColor = &H80000017
        lblRegUF = "Centro Oeste - MT (Mato Grosso)"
        txtRodoCapital.SetFocus
    ElseIf Val(lblEtapa) = 27 Then
        de_informa.ins_cadprazo txtCodTabela, "MT", "R", txtRodoCapital, txtRodoInterior
        de_informa.ins_cadprazo txtCodTabela, "MT", "A", txtAirCapital, txtAirInterior
        lblMT.ForeColor = &HC00000
        lblEtapaFim.Visible = False
        lblEtapa = "FINALIZADO"
        cmdGravar.Caption = "OK"
        cmdCancelar.Visible = False
        lblRegUF = "Processo Finalizado"
        If de_informa.rsSel_TabPrazoGro.State = 1 Then de_informa.rsSel_TabPrazoGro.Close
        de_informa.Sel_TabPrazoGro
        frmCadPrazos.gridCadPrazo.DataMember = "sel_tabprazogro"
        frmCadPrazos.gridCadPrazo.Refresh
        cmdGravar.SetFocus
        txtRodoCapital = ""
        txtRodoInterior = ""
        txtAirCapital = ""
        txtAirInterior = ""
    ElseIf cmdGravar.Caption = "OK" Then
        
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "INCLUSÃO", xusuario, "CAD. DE PRAZOS DE ENTREGA: " & txtCodTabela
    
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNovaTabPrz = Nothing
End Sub

Private Sub txtAirCapital_GotFocus()
    txtAirCapital.SelStart = 0
    txtAirCapital.SelLength = 3
End Sub

Private Sub txtAirCapital_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtAirInterior_GotFocus()
    txtAirInterior.SelStart = 0
    txtAirInterior.SelLength = 3
End Sub

Private Sub txtAirInterior_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtCodTabela_GotFocus()
    txtCodTabela.SelStart = 0
    txtCodTabela.SelLength = 6
End Sub

Private Sub txtCodTabela_LostFocus()
    txtCodTabela = UCase(txtCodTabela)
End Sub

Private Sub txtRodoCapital_GotFocus()
    txtRodoCapital.SelStart = 0
    txtRodoCapital.SelLength = 3
End Sub

Private Sub txtRodoCapital_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub txtRodoInterior_GotFocus()
    txtRodoInterior.SelStart = 0
    txtRodoInterior.SelLength = 3
End Sub

Private Sub txtRodoInterior_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
