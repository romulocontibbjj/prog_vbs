VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFaturaEletronica 
   Caption         =   "Gera Arquivo de Fatura Eletrônica"
   ClientHeight    =   7215
   ClientLeft      =   2010
   ClientTop       =   1845
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   10560
   Begin VB.Frame Frame1 
      Caption         =   "Fatura Eletrônica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   6135
      Begin VB.OptionButton optModelo3 
         Caption         =   "Modelo3"
         Height          =   255
         Left            =   4920
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optModelo2 
         Caption         =   "Modelo2"
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optModelo1 
         Caption         =   "Modelo 1"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtFilial 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   0
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtFatura 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscaPreFat 
         Caption         =   "Busca "
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   4800
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número da Filial-Fatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1665
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da Fatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   10335
      Begin VB.CommandButton cmdGeraArq 
         Caption         =   "Gera Arquivo de Fatura Eletrônica ..."
         Height          =   375
         Left            =   6960
         TabIndex        =   3
         Top             =   840
         Width           =   3255
      End
      Begin VB.CommandButton cmdImprime 
         Caption         =   "Imprimir ..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexFatEletr 
         Height          =   4215
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7435
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblVencto 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5760
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblClienteCNPJ 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Vencto:"
         Height          =   195
         Left            =   5160
         TabIndex        =   12
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Valor Bruto da Fatura:"
         Height          =   195
         Left            =   6960
         TabIndex        =   9
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label lblValorFatura 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8640
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Fatura Eletrônica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   6720
      TabIndex        =   17
      Top             =   360
      Width           =   3360
   End
End
Attribute VB_Name = "frmFaturaEletronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscaPreFat_Click()
    Dim xValorFatura As Currency
    
    flexFatEletr.Rows = 1
    flexFatEletr.Rows = 2
    flexFatEletr.FixedRows = 1
    
    If optModelo1 = True Then
    
        flexFatEletr.Rows = 2
        flexFatEletr.Cols = 31
        flexFatEletr.Row = 0
        flexFatEletr.Col = 1
        flexFatEletr.Text = "Filial"
        flexFatEletr.Col = 2
        flexFatEletr.Text = "CTC"
        flexFatEletr.Col = 3
        flexFatEletr.Text = "Remetente"
        flexFatEletr.Col = 4
        flexFatEletr.Text = "Origem"
        flexFatEletr.Col = 5
        flexFatEletr.Text = "Destinatario"
        flexFatEletr.Col = 6
        flexFatEletr.Text = "Cidade Destino"
        flexFatEletr.Col = 7
        flexFatEletr.Text = "UF"
        flexFatEletr.Col = 8
        flexFatEletr.Text = "NFS"
        flexFatEletr.Col = 9
        flexFatEletr.Text = "Valor Merc."
        flexFatEletr.Col = 10
        flexFatEletr.Text = "Volumes"
        flexFatEletr.Col = 11
        flexFatEletr.Text = "Peso"
        flexFatEletr.Col = 12
        flexFatEletr.Text = "Peso Tax."
        flexFatEletr.Col = 13
        flexFatEletr.Text = "Frete Total Liq."
        flexFatEletr.Col = 14
        flexFatEletr.Text = "Frete Total Bruto"
        flexFatEletr.Col = 15
        flexFatEletr.Text = "Frete Peso"
        flexFatEletr.Col = 16
        flexFatEletr.Text = "Frete Valor"
        flexFatEletr.Col = 17
        flexFatEletr.Text = "Gris"
        flexFatEletr.Col = 18
        flexFatEletr.Text = "Tx Urgencia"
        flexFatEletr.Col = 19
        flexFatEletr.Text = "Tx Coleta"
        flexFatEletr.Col = 20
        flexFatEletr.Text = "Tx Entrega/Red"
        flexFatEletr.Col = 21
        flexFatEletr.Text = "Pedagio"
        flexFatEletr.Col = 22
        flexFatEletr.Text = "Tx Outros"
        flexFatEletr.Col = 23
        flexFatEletr.Text = "Frete/Val.Merc (%)"
        flexFatEletr.Col = 24
        flexFatEletr.Text = "Data CTC"
        flexFatEletr.Col = 25
        flexFatEletr.Text = "Filial-Fatura"
        flexFatEletr.Col = 26
        flexFatEletr.Text = "Modal"
        flexFatEletr.Col = 27
        flexFatEletr.Text = "Obs Emissao"
        flexFatEletr.Col = 28
        flexFatEletr.Text = "URG/PRI/NOR"
        flexFatEletr.Col = 29
        flexFatEletr.Text = "NAT.PROD."
        flexFatEletr.Col = 30
        flexFatEletr.Text = "COD.TAB"
        
            
        flexFatEletr.ColWidth(0) = 200
        flexFatEletr.ColWidth(1) = 400
        flexFatEletr.ColWidth(2) = 800
        flexFatEletr.ColWidth(3) = 1900
        flexFatEletr.ColWidth(4) = 1900
        flexFatEletr.ColWidth(5) = 1900
        flexFatEletr.ColWidth(6) = 1900
        flexFatEletr.ColWidth(7) = 400
        flexFatEletr.ColWidth(8) = 2500
        flexFatEletr.ColWidth(9) = 1200
        flexFatEletr.ColWidth(10) = 800
        flexFatEletr.ColWidth(11) = 600
        flexFatEletr.ColWidth(12) = 600
        flexFatEletr.ColWidth(13) = 800
        flexFatEletr.ColWidth(14) = 1200
        flexFatEletr.ColWidth(15) = 1200
        flexFatEletr.ColWidth(16) = 1200
        flexFatEletr.ColWidth(17) = 1200
        flexFatEletr.ColWidth(18) = 1200
        flexFatEletr.ColWidth(19) = 1200
        flexFatEletr.ColWidth(20) = 1200
        flexFatEletr.ColWidth(21) = 1200
        flexFatEletr.ColWidth(22) = 1200
        flexFatEletr.ColWidth(23) = 1200
        flexFatEletr.ColWidth(24) = 1200
        flexFatEletr.ColWidth(25) = 1200
        flexFatEletr.ColWidth(26) = 1200
        flexFatEletr.ColWidth(27) = 4000
        flexFatEletr.ColWidth(28) = 1400
        flexFatEletr.ColWidth(29) = 1400
        flexFatEletr.ColWidth(30) = 1400
    
        If de_informa.rsSel_GeraArqFatEletr.State = 1 Then de_informa.rsSel_GeraArqFatEletr.Close
        de_informa.Sel_GeraArqFatEletr TransFatur(txtFilial, txtFatura)
        
        If de_informa.rsSel_GeraArqFatEletr.RecordCount > 0 Then
            flexFatEletr.Rows = de_informa.rsSel_GeraArqFatEletr.RecordCount + 2
            flexFatEletr.FixedRows = 1
            lblClienteCNPJ = de_informa.rsSel_GeraArqFatEletr.Fields("cliente_cgc")
            lblCliente = de_informa.rsSel_GeraArqFatEletr.Fields("cliente_nome")
            lblVencto = de_informa.rsSel_GeraArqFatEletr.Fields("vencimento")
            lblValorFatura = Format(de_informa.rsSel_GeraArqFatEletr.Fields("valorfatura"), "#,###,##0.00")
            For xcont = 1 To de_informa.rsSel_GeraArqFatEletr.RecordCount
                flexFatEletr.TextMatrix(xcont, 1) = de_informa.rsSel_GeraArqFatEletr.Fields("filial")
                flexFatEletr.TextMatrix(xcont, 2) = de_informa.rsSel_GeraArqFatEletr.Fields("ctc")
                flexFatEletr.TextMatrix(xcont, 3) = de_informa.rsSel_GeraArqFatEletr.Fields("remet_nome")
                flexFatEletr.TextMatrix(xcont, 4) = Trim$(de_informa.rsSel_GeraArqFatEletr.Fields("cidade_orig")) & "-" & de_informa.rsSel_GeraArqFatEletr.Fields("uf_orig")
                flexFatEletr.TextMatrix(xcont, 5) = de_informa.rsSel_GeraArqFatEletr.Fields("dest_nome")
                flexFatEletr.TextMatrix(xcont, 6) = de_informa.rsSel_GeraArqFatEletr.Fields("dest_cidade")
                flexFatEletr.TextMatrix(xcont, 7) = de_informa.rsSel_GeraArqFatEletr.Fields("dest_uf")
                flexFatEletr.TextMatrix(xcont, 8) = de_informa.rsSel_GeraArqFatEletr.Fields("nfs")
                flexFatEletr.TextMatrix(xcont, 9) = de_informa.rsSel_GeraArqFatEletr.Fields("valmerc")
                flexFatEletr.TextMatrix(xcont, 10) = de_informa.rsSel_GeraArqFatEletr.Fields("volumes")
                flexFatEletr.TextMatrix(xcont, 11) = de_informa.rsSel_GeraArqFatEletr.Fields("peso")
                flexFatEletr.TextMatrix(xcont, 12) = de_informa.rsSel_GeraArqFatEletr.Fields("pesotax")
                flexFatEletr.TextMatrix(xcont, 13) = de_informa.rsSel_GeraArqFatEletr.Fields("frete")
                flexFatEletr.TextMatrix(xcont, 14) = de_informa.rsSel_GeraArqFatEletr.Fields("fretetotalbruto")
                If IsNull(de_informa.rsSel_GeraArqFatEletr.Fields("fretepesobr")) Then
                    flexFatEletr.TextMatrix(xcont, 15) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 15) = de_informa.rsSel_GeraArqFatEletr.Fields("fretepesobr")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletr.Fields("fretevalorbr")) Then
                    flexFatEletr.TextMatrix(xcont, 16) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 16) = de_informa.rsSel_GeraArqFatEletr.Fields("fretevalorbr")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletr.Fields("grisbr")) Then
                    flexFatEletr.TextMatrix(xcont, 17) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 17) = de_informa.rsSel_GeraArqFatEletr.Fields("grisbr")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletr.Fields("txurgenciabr")) Then
                    flexFatEletr.TextMatrix(xcont, 18) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 18) = de_informa.rsSel_GeraArqFatEletr.Fields("txurgenciabr")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletr.Fields("txcoletabr")) Then
                    flexFatEletr.TextMatrix(xcont, 19) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 19) = de_informa.rsSel_GeraArqFatEletr.Fields("txcoletabr")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletr.Fields("txentregared")) Then
                    flexFatEletr.TextMatrix(xcont, 20) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 20) = de_informa.rsSel_GeraArqFatEletr.Fields("txentregared")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletr.Fields("pedagiobr")) Then
                    flexFatEletr.TextMatrix(xcont, 21) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 21) = de_informa.rsSel_GeraArqFatEletr.Fields("pedagiobr")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletr.Fields("txoutrosbr")) Then
                    flexFatEletr.TextMatrix(xcont, 22) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 22) = de_informa.rsSel_GeraArqFatEletr.Fields("txoutrosbr")
                End If
                If de_informa.rsSel_GeraArqFatEletr.Fields("valmerc") > 1 And de_informa.rsSel_GeraArqFatEletr.Fields("frete") > 1 Then
                    flexFatEletr.TextMatrix(xcont, 23) = Mid$(de_informa.rsSel_GeraArqFatEletr.Fields("perc"), 1, _
                                                         InStr(1, de_informa.rsSel_GeraArqFatEletr.Fields("perc"), ".", vbTextCompare) - 1) & "," & _
                                                         Mid$(de_informa.rsSel_GeraArqFatEletr.Fields("perc"), _
                                                         InStr(1, de_informa.rsSel_GeraArqFatEletr.Fields("perc"), ".", vbTextCompare) + 1)
                Else
                    flexFatEletr.TextMatrix(xcont, 23) = "0%"
                End If
                flexFatEletr.TextMatrix(xcont, 24) = de_informa.rsSel_GeraArqFatEletr.Fields("data")
                flexFatEletr.TextMatrix(xcont, 25) = de_informa.rsSel_GeraArqFatEletr.Fields("Fatura")
                flexFatEletr.TextMatrix(xcont, 26) = de_informa.rsSel_GeraArqFatEletr.Fields("modal")
                flexFatEletr.TextMatrix(xcont, 27) = de_informa.rsSel_GeraArqFatEletr.Fields("obs_emissao")
                flexFatEletr.TextMatrix(xcont, 28) = de_informa.rsSel_GeraArqFatEletr.Fields("prioridade")
                flexFatEletr.TextMatrix(xcont, 29) = de_informa.rsSel_GeraArqFatEletr.Fields("natureza")
                flexFatEletr.TextMatrix(xcont, 30) = de_informa.rsSel_GeraArqFatEletr.Fields("tabfrete")
                de_informa.rsSel_GeraArqFatEletr.MoveNext
            Next
        Else
            MsgBox "Arquivo de Fatura Não Encontrado !"
            Exit Sub
        End If
        
    ElseIf optModelo2.Value = True Then
    
        flexFatEletr.Rows = 2
        flexFatEletr.Cols = 21
        flexFatEletr.Row = 0
        flexFatEletr.Col = 1
        flexFatEletr.Text = "Filial-Fatura"
        flexFatEletr.Col = 2
        flexFatEletr.Text = "Cliente_Nome"
        flexFatEletr.Col = 3
        flexFatEletr.Text = "Vencimento"
        flexFatEletr.Col = 4
        flexFatEletr.Text = "FilialCTC"
        flexFatEletr.Col = 5
        flexFatEletr.Text = "NFS"
        flexFatEletr.Col = 6
        flexFatEletr.Text = "Data"
        flexFatEletr.Col = 7
        flexFatEletr.Text = "Dest_Nome"
        flexFatEletr.Col = 8
        flexFatEletr.Text = "Endereco"
        flexFatEletr.Col = 9
        flexFatEletr.Text = "Cidade"
        flexFatEletr.Col = 10
        flexFatEletr.Text = "UF"
        flexFatEletr.Col = 11
        flexFatEletr.Text = "Natureza"
        flexFatEletr.Col = 12
        flexFatEletr.Text = "Prioridade"
        flexFatEletr.Col = 13
        flexFatEletr.Text = "Peso"
        flexFatEletr.Col = 14
        flexFatEletr.Text = "Peso Tax."
        flexFatEletr.Col = 15
        flexFatEletr.Text = "Valmerc."
        flexFatEletr.Col = 16
        flexFatEletr.Text = "Regiao"
        flexFatEletr.Col = 17
        flexFatEletr.Text = "Modal"
        flexFatEletr.Col = 18
        flexFatEletr.Text = "Frete"
        flexFatEletr.Col = 19
        flexFatEletr.Text = "Remetente"
            
        flexFatEletr.ColWidth(0) = 200
        flexFatEletr.ColWidth(1) = 1600
        flexFatEletr.ColWidth(2) = 1600
        flexFatEletr.ColWidth(3) = 1600
        flexFatEletr.ColWidth(4) = 1600
        flexFatEletr.ColWidth(5) = 1600
        flexFatEletr.ColWidth(6) = 1600
        flexFatEletr.ColWidth(7) = 1600
        flexFatEletr.ColWidth(8) = 1600
        flexFatEletr.ColWidth(9) = 1600
        flexFatEletr.ColWidth(10) = 1600
        flexFatEletr.ColWidth(11) = 1600
        flexFatEletr.ColWidth(12) = 1600
        flexFatEletr.ColWidth(13) = 1600
        flexFatEletr.ColWidth(14) = 1600
        flexFatEletr.ColWidth(15) = 1600
        flexFatEletr.ColWidth(16) = 1600
        flexFatEletr.ColWidth(17) = 1600
        flexFatEletr.ColWidth(18) = 1600
        flexFatEletr.ColWidth(19) = 1600
    
        If de_informa.rsSel_GeraArqFatEletr.State = 1 Then de_informa.rsSel_GeraArqFatEletr.Close
        de_informa.Sel_GeraArqFatEletr TransFatur(txtFilial, txtFatura)
        
        If de_informa.rsSel_GeraArqFatEletr.RecordCount > 0 Then
            flexFatEletr.Rows = de_informa.rsSel_GeraArqFatEletr.RecordCount + 2
            flexFatEletr.FixedRows = 1
            lblClienteCNPJ = de_informa.rsSel_GeraArqFatEletr.Fields("cliente_cgc")
            lblCliente = de_informa.rsSel_GeraArqFatEletr.Fields("cliente_nome")
            lblVencto = de_informa.rsSel_GeraArqFatEletr.Fields("vencimento")
            lblValorFatura = Format(de_informa.rsSel_GeraArqFatEletr.Fields("valorfatura"), "#,###,##0.00")
            For xcont = 1 To de_informa.rsSel_GeraArqFatEletr.RecordCount
                flexFatEletr.TextMatrix(xcont, 1) = TransFatur(txtFilial, txtFatura)
                flexFatEletr.TextMatrix(xcont, 2) = de_informa.rsSel_GeraArqFatEletr.Fields("cliente_nome")
                flexFatEletr.TextMatrix(xcont, 3) = de_informa.rsSel_GeraArqFatEletr.Fields("vencimento")
                flexFatEletr.TextMatrix(xcont, 4) = de_informa.rsSel_GeraArqFatEletr.Fields("filialctc")
                flexFatEletr.TextMatrix(xcont, 5) = de_informa.rsSel_GeraArqFatEletr.Fields("nfs")
                flexFatEletr.TextMatrix(xcont, 6) = de_informa.rsSel_GeraArqFatEletr.Fields("data")
                flexFatEletr.TextMatrix(xcont, 7) = de_informa.rsSel_GeraArqFatEletr.Fields("dest_nome")
                flexFatEletr.TextMatrix(xcont, 8) = de_informa.rsSel_GeraArqFatEletr.Fields("dest_end")
                flexFatEletr.TextMatrix(xcont, 9) = de_informa.rsSel_GeraArqFatEletr.Fields("dest_cidade")
                flexFatEletr.TextMatrix(xcont, 10) = de_informa.rsSel_GeraArqFatEletr.Fields("dest_uf")
                flexFatEletr.TextMatrix(xcont, 11) = de_informa.rsSel_GeraArqFatEletr.Fields("naturezaobs")
                flexFatEletr.TextMatrix(xcont, 12) = de_informa.rsSel_GeraArqFatEletr.Fields("prioridade")
                flexFatEletr.TextMatrix(xcont, 13) = de_informa.rsSel_GeraArqFatEletr.Fields("peso")
                flexFatEletr.TextMatrix(xcont, 14) = de_informa.rsSel_GeraArqFatEletr.Fields("pesotax")
                flexFatEletr.TextMatrix(xcont, 15) = de_informa.rsSel_GeraArqFatEletr.Fields("valmerc")
                flexFatEletr.TextMatrix(xcont, 16) = de_informa.rsSel_GeraArqFatEletr.Fields("regiao")
                flexFatEletr.TextMatrix(xcont, 17) = de_informa.rsSel_GeraArqFatEletr.Fields("modal")
                flexFatEletr.TextMatrix(xcont, 18) = de_informa.rsSel_GeraArqFatEletr.Fields("frete")
                flexFatEletr.TextMatrix(xcont, 19) = de_informa.rsSel_GeraArqFatEletr.Fields("remet_nome")
                
                de_informa.rsSel_GeraArqFatEletr.MoveNext
            Next
    
        Else
            MsgBox "Arquivo de Fatura Não Encontrado !"
            Exit Sub
        End If
        
     ElseIf optModelo3.Value = True Then
        flexFatEletr.Rows = 2
        flexFatEletr.Cols = 29
        flexFatEletr.Row = 0
        flexFatEletr.Col = 1
        flexFatEletr.Text = "Filial"
        flexFatEletr.Col = 2
        flexFatEletr.Text = "CTC"
        flexFatEletr.Col = 3
        flexFatEletr.Text = "Remetente"
        flexFatEletr.Col = 4
        flexFatEletr.Text = "Origem"
        flexFatEletr.Col = 5
        flexFatEletr.Text = "Destinatario"
        flexFatEletr.Col = 6
        flexFatEletr.Text = "Cidade Destino"
        flexFatEletr.Col = 7
        flexFatEletr.Text = "UF"
        flexFatEletr.Col = 8
        flexFatEletr.Text = "NFS"
        flexFatEletr.Col = 9
        flexFatEletr.Text = "Valor Merc."
        flexFatEletr.Col = 10
        flexFatEletr.Text = "Volumes"
        flexFatEletr.Col = 11
        flexFatEletr.Text = "Peso"
        flexFatEletr.Col = 12
        flexFatEletr.Text = "Peso Tax."
        flexFatEletr.Col = 13
        flexFatEletr.Text = "Frete Total Liq."
        flexFatEletr.Col = 14
        flexFatEletr.Text = "Frete Total Bruto"
        flexFatEletr.Col = 15
        flexFatEletr.Text = "Frete Peso"
        flexFatEletr.Col = 16
        flexFatEletr.Text = "Frete Valor"
        flexFatEletr.Col = 17
        flexFatEletr.Text = "Gris"
        flexFatEletr.Col = 18
        flexFatEletr.Text = "Tx Urgencia"
        flexFatEletr.Col = 19
        flexFatEletr.Text = "Tx Coleta"
        flexFatEletr.Col = 20
        flexFatEletr.Text = "Tx Entrega/Red"
        flexFatEletr.Col = 21
        flexFatEletr.Text = "Pedagio"
        flexFatEletr.Col = 22
        flexFatEletr.Text = "Tx Outros"
        flexFatEletr.Col = 23
        flexFatEletr.Text = "Frete/Val.Merc (%)"
        flexFatEletr.Col = 24
        flexFatEletr.Text = "Data CTC"
        flexFatEletr.Col = 25
        flexFatEletr.Text = "Filial-Fatura"
        flexFatEletr.Col = 26
        flexFatEletr.Text = "Modal"
        flexFatEletr.Col = 27
        flexFatEletr.Text = "Obs Emissao"
        flexFatEletr.Col = 28
        flexFatEletr.Text = "URG/PRI/NOR"
            
        flexFatEletr.ColWidth(0) = 200
        flexFatEletr.ColWidth(1) = 400
        flexFatEletr.ColWidth(2) = 800
        flexFatEletr.ColWidth(3) = 1900
        flexFatEletr.ColWidth(4) = 1900
        flexFatEletr.ColWidth(5) = 1900
        flexFatEletr.ColWidth(6) = 1900
        flexFatEletr.ColWidth(7) = 400
        flexFatEletr.ColWidth(8) = 2500
        flexFatEletr.ColWidth(9) = 1200
        flexFatEletr.ColWidth(10) = 800
        flexFatEletr.ColWidth(11) = 600
        flexFatEletr.ColWidth(12) = 600
        flexFatEletr.ColWidth(13) = 800
        flexFatEletr.ColWidth(14) = 1200
        flexFatEletr.ColWidth(15) = 1200
        flexFatEletr.ColWidth(16) = 1200
        flexFatEletr.ColWidth(17) = 1200
        flexFatEletr.ColWidth(18) = 1200
        flexFatEletr.ColWidth(19) = 1200
        flexFatEletr.ColWidth(20) = 1200
        flexFatEletr.ColWidth(21) = 1200
        flexFatEletr.ColWidth(22) = 1200
        flexFatEletr.ColWidth(23) = 1200
        flexFatEletr.ColWidth(24) = 1200
        flexFatEletr.ColWidth(25) = 1200
        flexFatEletr.ColWidth(26) = 1200
        flexFatEletr.ColWidth(27) = 4000
        flexFatEletr.ColWidth(28) = 1400
        
        If de_informa.rsSel_GeraArqFatEletrCTR.State = 1 Then de_informa.rsSel_GeraArqFatEletrCTR.Close
        de_informa.Sel_GeraArqFatEletrCTR TransFatur(txtFilial, txtFatura)
        
        If de_informa.rsSel_GeraArqFatEletrCTR.RecordCount > 0 Then
            flexFatEletr.Rows = de_informa.rsSel_GeraArqFatEletrCTR.RecordCount + 2
            flexFatEletr.FixedRows = 1
            lblClienteCNPJ = de_informa.rsSel_GeraArqFatEletrCTR.Fields("cliente_cgc")
            lblCliente = de_informa.rsSel_GeraArqFatEletrCTR.Fields("cliente_nome")
            lblVencto = de_informa.rsSel_GeraArqFatEletrCTR.Fields("vencimento")
            lblValorFatura = Format(de_informa.rsSel_GeraArqFatEletrCTR.Fields("valorfatura"), "#,###,##0.00")
            For xcont = 1 To de_informa.rsSel_GeraArqFatEletrCTR.RecordCount
                flexFatEletr.TextMatrix(xcont, 1) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("filial")
                flexFatEletr.TextMatrix(xcont, 2) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("ctc")
                flexFatEletr.TextMatrix(xcont, 3) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("remet_nome")
                flexFatEletr.TextMatrix(xcont, 4) = Trim$(de_informa.rsSel_GeraArqFatEletrCTR.Fields("cidade_orig")) & "-" & de_informa.rsSel_GeraArqFatEletrCTR.Fields("uf_orig")
                flexFatEletr.TextMatrix(xcont, 5) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("dest_nome")
                flexFatEletr.TextMatrix(xcont, 6) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("dest_cidade")
                flexFatEletr.TextMatrix(xcont, 7) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("dest_uf")
                flexFatEletr.TextMatrix(xcont, 8) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("nfs")
                flexFatEletr.TextMatrix(xcont, 9) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("valmerc")
                flexFatEletr.TextMatrix(xcont, 10) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("volumes")
                flexFatEletr.TextMatrix(xcont, 11) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("peso")
                flexFatEletr.TextMatrix(xcont, 12) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("pesotax")
                flexFatEletr.TextMatrix(xcont, 13) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("frete")
                flexFatEletr.TextMatrix(xcont, 14) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("fretetotalbruto")
                If IsNull(de_informa.rsSel_GeraArqFatEletrCTR.Fields("fretepesobr")) Then
                    flexFatEletr.TextMatrix(xcont, 15) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 15) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("fretepesobr")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletrCTR.Fields("fretevalorbr")) Then
                    flexFatEletr.TextMatrix(xcont, 16) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 16) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("fretevalorbr")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletrCTR.Fields("grisbr")) Then
                    flexFatEletr.TextMatrix(xcont, 17) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 17) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("grisbr")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletrCTR.Fields("txurgenciabr")) Then
                    flexFatEletr.TextMatrix(xcont, 18) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 18) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("txurgenciabr")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletrCTR.Fields("txcoletabr")) Then
                    flexFatEletr.TextMatrix(xcont, 19) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 19) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("txcoletabr")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletrCTR.Fields("txentregared")) Then
                    flexFatEletr.TextMatrix(xcont, 20) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 20) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("txentregared")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletrCTR.Fields("pedagiobr")) Then
                    flexFatEletr.TextMatrix(xcont, 21) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 21) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("pedagiobr")
                End If
                If IsNull(de_informa.rsSel_GeraArqFatEletrCTR.Fields("txoutrosbr")) Then
                    flexFatEletr.TextMatrix(xcont, 22) = "0"
                Else
                    flexFatEletr.TextMatrix(xcont, 22) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("txoutrosbr")
                End If
                If de_informa.rsSel_GeraArqFatEletrCTR.Fields("valmerc") > 1 And de_informa.rsSel_GeraArqFatEletrCTR.Fields("frete") > 1 Then
                    flexFatEletr.TextMatrix(xcont, 23) = Mid$(de_informa.rsSel_GeraArqFatEletrCTR.Fields("perc"), 1, _
                                                         InStr(1, de_informa.rsSel_GeraArqFatEletrCTR.Fields("perc"), ".", vbTextCompare) - 1) & "," & _
                                                         Mid$(de_informa.rsSel_GeraArqFatEletrCTR.Fields("perc"), _
                                                         InStr(1, de_informa.rsSel_GeraArqFatEletrCTR.Fields("perc"), ".", vbTextCompare) + 1)
                                                         
                Else
                    flexFatEletr.TextMatrix(xcont, 23) = "0%"
                End If
                flexFatEletr.TextMatrix(xcont, 24) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("data")
                flexFatEletr.TextMatrix(xcont, 25) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("Fatura")
                flexFatEletr.TextMatrix(xcont, 26) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("modal")
                flexFatEletr.TextMatrix(xcont, 27) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("obs_emissao")
                flexFatEletr.TextMatrix(xcont, 28) = de_informa.rsSel_GeraArqFatEletrCTR.Fields("prioridade")
                de_informa.rsSel_GeraArqFatEletrCTR.MoveNext
            Next
    
        Else
            MsgBox "Arquivo de Fatura Não Encontrado !"
            Exit Sub
        End If
        
    End If
    
    DoEvents
    
    flexFatEletr_Click
    
End Sub
Private Sub cmdGeraArq_Click()
    Dim xcont As Long, xFiles As String, xlinha As String
    If flexFatEletr.Rows < 3 Then
        MsgBox "Não há Dados Para Geração de Arquivos !", vbExclamation, "Ops"
    Else
            
        xFiles = "C:\FATURA\" & Trim$(Mid$(lblCliente, 1, 5)) & "_" & TransFatur(txtFilial, txtFatura) & "_" & zeros(Day(Date), 2) & zeros(Month(Date), 2) & "_" & Mid$(Trim$(CVar(Time())), 1, 2) & Mid$(Trim$(CVar(Time())), 4, 2) & ".txt"
        
        Open xFiles For Output As #1
    
        For xcont = 0 To flexFatEletr.Rows - 1
        
            If optModelo1.Value = True Then
            
                xlinha = flexFatEletr.TextMatrix(xcont, 1) & "#" & flexFatEletr.TextMatrix(xcont, 2) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 3) & "#" & flexFatEletr.TextMatrix(xcont, 4) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 5) & "#" & flexFatEletr.TextMatrix(xcont, 6) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 7) & "#" & flexFatEletr.TextMatrix(xcont, 8) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 9) & "#" & flexFatEletr.TextMatrix(xcont, 10) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 11) & "#" & flexFatEletr.TextMatrix(xcont, 12) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 13) & "#" & flexFatEletr.TextMatrix(xcont, 14) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 15) & "#" & flexFatEletr.TextMatrix(xcont, 16) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 17) & "#" & flexFatEletr.TextMatrix(xcont, 18) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 19) & "#" & flexFatEletr.TextMatrix(xcont, 20) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 21) & "#" & flexFatEletr.TextMatrix(xcont, 22) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 23) & "#" & flexFatEletr.TextMatrix(xcont, 24) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 25) & "#" & flexFatEletr.TextMatrix(xcont, 26) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 27) & "#" & flexFatEletr.TextMatrix(xcont, 28) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 29) & "#" & flexFatEletr.TextMatrix(xcont, 30) & "#"
                    
            ElseIf optModelo2.Value = True Then
            
                xlinha = flexFatEletr.TextMatrix(xcont, 1) & "#" & flexFatEletr.TextMatrix(xcont, 2) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 3) & "#" & flexFatEletr.TextMatrix(xcont, 4) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 5) & "#" & flexFatEletr.TextMatrix(xcont, 6) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 7) & "#" & flexFatEletr.TextMatrix(xcont, 8) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 9) & "#" & flexFatEletr.TextMatrix(xcont, 10) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 11) & "#" & flexFatEletr.TextMatrix(xcont, 12) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 13) & "#" & flexFatEletr.TextMatrix(xcont, 14) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 15) & "#" & flexFatEletr.TextMatrix(xcont, 16) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 17) & "#" & flexFatEletr.TextMatrix(xcont, 18) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 19) & "#"
            
            ElseIf optModelo3.Value = True Then
                xlinha = flexFatEletr.TextMatrix(xcont, 1) & "#" & flexFatEletr.TextMatrix(xcont, 2) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 3) & "#" & flexFatEletr.TextMatrix(xcont, 4) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 5) & "#" & flexFatEletr.TextMatrix(xcont, 6) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 7) & "#" & flexFatEletr.TextMatrix(xcont, 8) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 9) & "#" & flexFatEletr.TextMatrix(xcont, 10) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 11) & "#" & flexFatEletr.TextMatrix(xcont, 12) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 13) & "#" & flexFatEletr.TextMatrix(xcont, 14) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 15) & "#" & flexFatEletr.TextMatrix(xcont, 16) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 17) & "#" & flexFatEletr.TextMatrix(xcont, 18) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 19) & "#" & flexFatEletr.TextMatrix(xcont, 20) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 21) & "#" & flexFatEletr.TextMatrix(xcont, 22) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 23) & "#" & flexFatEletr.TextMatrix(xcont, 24) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 25) & "#" & flexFatEletr.TextMatrix(xcont, 26) & "#" & _
                        flexFatEletr.TextMatrix(xcont, 27) & "#" & flexFatEletr.TextMatrix(xcont, 28) & "#"
            End If
                    
            Print #1, xlinha
            
            DoEvents
        Next
        
        Close #1
        
        MsgBox "OK ! Arquivo Gerado em " & xFiles & "." & Chr(13) + Chr(10) + Chr(13) + Chr(10) + _
        "O Arquivo Gerado é do Tipo Texto ( TXT com Delimitador # ) e você pode abrí-lo em diversos aplicativos. Para Abrí-lo no MS-Excel, em ABRIR escolha ARQUIVOS DO TIPO = Arquivos de Texto e selecione o arquivo no local indicado acima. Na Caixa ASSISTENTE DE IMPORTAÇÃO escolha DELIMITADO e o caracter delimitador escolha OUTROS e digite # . Clique em Concluir e o arquivo será importado para o MS-Excel.", vbInformation, "Geração de Arquivo TXT"

    End If
        


End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub flexFatEletr_Click()
    DoEvents
End Sub

Private Sub Form_Load()
    flexFatEletr.Rows = 2
    flexFatEletr.Cols = 29
    flexFatEletr.Row = 0
    flexFatEletr.Col = 1
    flexFatEletr.Text = "Filial"
    flexFatEletr.Col = 2
    flexFatEletr.Text = "CTC"
    flexFatEletr.Col = 3
    flexFatEletr.Text = "Remetente"
    flexFatEletr.Col = 4
    flexFatEletr.Text = "Origem"
    flexFatEletr.Col = 5
    flexFatEletr.Text = "Destinatario"
    flexFatEletr.Col = 6
    flexFatEletr.Text = "Cidade Destino"
    flexFatEletr.Col = 7
    flexFatEletr.Text = "UF"
    flexFatEletr.Col = 8
    flexFatEletr.Text = "NFS"
    flexFatEletr.Col = 9
    flexFatEletr.Text = "Valor Merc."
    flexFatEletr.Col = 10
    flexFatEletr.Text = "Volumes"
    flexFatEletr.Col = 11
    flexFatEletr.Text = "Peso"
    flexFatEletr.Col = 12
    flexFatEletr.Text = "Peso Tax."
    flexFatEletr.Col = 13
    flexFatEletr.Text = "Frete Total Liq."
    flexFatEletr.Col = 14
    flexFatEletr.Text = "Frete Total Bruto"
    flexFatEletr.Col = 15
    flexFatEletr.Text = "Frete Peso"
    flexFatEletr.Col = 16
    flexFatEletr.Text = "Frete Valor"
    flexFatEletr.Col = 17
    flexFatEletr.Text = "Gris"
    flexFatEletr.Col = 18
    flexFatEletr.Text = "Tx Urgencia"
    flexFatEletr.Col = 19
    flexFatEletr.Text = "Tx Coleta"
    flexFatEletr.Col = 20
    flexFatEletr.Text = "Tx Entrega/Red"
    flexFatEletr.Col = 21
    flexFatEletr.Text = "Pedagio"
    flexFatEletr.Col = 22
    flexFatEletr.Text = "Tx Outros"
    flexFatEletr.Col = 23
    flexFatEletr.Text = "Frete/Val.Merc (%)"
    flexFatEletr.Col = 24
    flexFatEletr.Text = "Data CTC"
    flexFatEletr.Col = 25
    flexFatEletr.Text = "Filial-Fatura"
    flexFatEletr.Col = 26
    flexFatEletr.Text = "Modal"
    flexFatEletr.Col = 27
    flexFatEletr.Text = "Obs Emissao"
    flexFatEletr.Col = 28
    flexFatEletr.Text = "URG/PRI/NOR"
        
    flexFatEletr.ColWidth(0) = 200
    flexFatEletr.ColWidth(1) = 400
    flexFatEletr.ColWidth(2) = 800
    flexFatEletr.ColWidth(3) = 1900
    flexFatEletr.ColWidth(4) = 1900
    flexFatEletr.ColWidth(5) = 1900
    flexFatEletr.ColWidth(6) = 1900
    flexFatEletr.ColWidth(7) = 400
    flexFatEletr.ColWidth(8) = 2500
    flexFatEletr.ColWidth(9) = 1200
    flexFatEletr.ColWidth(10) = 800
    flexFatEletr.ColWidth(11) = 600
    flexFatEletr.ColWidth(12) = 600
    flexFatEletr.ColWidth(13) = 800
    flexFatEletr.ColWidth(14) = 1200
    flexFatEletr.ColWidth(15) = 1200
    flexFatEletr.ColWidth(16) = 1200
    flexFatEletr.ColWidth(17) = 1200
    flexFatEletr.ColWidth(18) = 1200
    flexFatEletr.ColWidth(19) = 1200
    flexFatEletr.ColWidth(20) = 1200
    flexFatEletr.ColWidth(21) = 1200
    flexFatEletr.ColWidth(22) = 1200
    flexFatEletr.ColWidth(23) = 1200
    flexFatEletr.ColWidth(24) = 1200
    flexFatEletr.ColWidth(25) = 1200
    flexFatEletr.ColWidth(26) = 1200
    flexFatEletr.ColWidth(27) = 4000
    flexFatEletr.ColWidth(28) = 1400
    
End Sub

Private Sub txtPreFat_Change()
    If Not IsNumeric(txtPreFat) Then
        SendKeys "{BACKSPACE}"
    End If
End Sub

Private Sub txtPreFat_GotFocus()
    txtPreFat.SelStart = 0
    txtPreFat.SelLength = 2
End Sub

Private Sub txtPreFat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtPreFat)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub

Private Sub optModelo1_Click()
    flexFatEletr.Rows = 1
    flexFatEletr.Rows = 2
    flexFatEletr.FixedRows = 1

End Sub

Private Sub optModelo2_Click()
    flexFatEletr.Rows = 1
    flexFatEletr.Rows = 2
    flexFatEletr.FixedRows = 1

End Sub

Private Sub txtFatura_Change()
    If Len(Trim$(txtFatura)) > 0 Then
        If Not IsNumeric(txtFatura) Or Mid$(txtFatura, Len(txtFatura), 1) = "," Or Mid$(txtFatura, Len(txtFatura), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
    DoEvents
End Sub
Private Sub txtFatura_GotFocus()
    txtFatura.SelStart = 0
    txtFatura.SelLength = 8
End Sub
Private Sub txtFatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtFatura)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub


Private Sub txtFilial_Change()
    If Len(Trim$(txtFilial)) > 0 Then
        If Not IsNumeric(txtFilial) Or Mid$(txtFilial, Len(txtFilial), 1) = "," Or Mid$(txtFilial, Len(txtFilial), 1) = "." Then
            SendKeys "{BACKSPACE}"
        End If
    End If
    DoEvents
    If Len(Trim$(txtFilial)) = 2 Then
        txtFatura.SetFocus
    End If
End Sub
Private Sub txtFilial_GotFocus()
    txtFilial.SelStart = 0
    txtFilial.SelLength = 2
End Sub
Private Sub txtFilial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtFilial)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub txtFilial_LostFocus()
    If Len(Trim$(txtFilial)) = 1 Then
        txtFilial = "0" & Trim$(txtFilial)
    End If
End Sub

