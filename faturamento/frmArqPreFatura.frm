VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmArqPreFatura 
   Caption         =   "Arquivo de Pré-Fatura"
   ClientHeight    =   6960
   ClientLeft      =   645
   ClientTop       =   1530
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10560
   Begin VB.Frame Frame2 
      Caption         =   "Dados da Pré-Fatura"
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
      TabIndex        =   4
      Top             =   1200
      Width           =   10335
      Begin VB.CommandButton cmdImprime 
         Caption         =   "Imprimir ..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         TabIndex        =   15
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdGeraArq 
         Caption         =   "Gera Arquivo de Pré-Fatura ..."
         Height          =   375
         Left            =   6960
         TabIndex        =   13
         Top             =   840
         Width           =   3255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexPrefat 
         Height          =   4215
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   7435
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Valor Bruto da Fatura:"
         Height          =   195
         Left            =   6960
         TabIndex        =   11
         Top             =   360
         Width           =   1545
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
      Begin VB.Label lblCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Vencto:"
         Height          =   195
         Left            =   5160
         TabIndex        =   8
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblClienteCNPJ 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblVencto 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5760
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pré-Fatura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.OptionButton optModelo3 
         Caption         =   "Modelo3"
         Height          =   255
         Left            =   4920
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optModelo1 
         Caption         =   "Modelo 1"
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optModelo2 
         Caption         =   "Modelo2"
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   4800
         TabIndex        =   14
         Top             =   550
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscaPreFat 
         Caption         =   "Busca "
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   550
         Width           =   1095
      End
      Begin VB.TextBox txtPreFat 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   2
         Top             =   550
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número da Filial+Pré-Fatura:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   550
         Width           =   1995
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "PRÉ-FATURA"
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
      Left            =   7185
      TabIndex        =   16
      Top             =   360
      Width           =   2475
   End
End
Attribute VB_Name = "frmArqPreFatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscaPreFat_Click()
    Dim xValorFatura As Currency
    
    flexPrefat.Rows = 1
    flexPrefat.Rows = 2
    flexPrefat.FixedRows = 1
    
    If de_informa.rsSel_GerarArqPrefat.State = 1 Then de_informa.rsSel_GerarArqPrefat.Close
    de_informa.Sel_GerarArqPrefat Trim$(txtPreFat)
    
    If optModelo1 = True Then
        
        flexPrefat.Rows = 2
        flexPrefat.Cols = 30
        flexPrefat.Row = 0
        flexPrefat.Col = 1
        flexPrefat.Text = "Filial"
        flexPrefat.Col = 2
        flexPrefat.Text = "CTC"
        flexPrefat.Col = 3
        flexPrefat.Text = "Remetente"
        flexPrefat.Col = 4
        flexPrefat.Text = "Origem"
        flexPrefat.Col = 5
        flexPrefat.Text = "Destinatario"
        flexPrefat.Col = 6
        flexPrefat.Text = "Cidade Destino"
        flexPrefat.Col = 7
        flexPrefat.Text = "UF"
        flexPrefat.Col = 8
        flexPrefat.Text = "NFS"
        flexPrefat.Col = 9
        flexPrefat.Text = "Valor Merc."
        flexPrefat.Col = 10
        flexPrefat.Text = "Volumes"
        flexPrefat.Col = 11
        flexPrefat.Text = "Peso"
        flexPrefat.Col = 12
        flexPrefat.Text = "Peso Tax."
        flexPrefat.Col = 13
        flexPrefat.Text = "Frete Total Liq."
        flexPrefat.Col = 14
        flexPrefat.Text = "Frete Total Bruto"
        flexPrefat.Col = 15
        flexPrefat.Text = "Frete Peso"
        flexPrefat.Col = 16
        flexPrefat.Text = "Frete Valor"
        flexPrefat.Col = 17
        flexPrefat.Text = "Gris"
        flexPrefat.Col = 18
        flexPrefat.Text = "Tx Urgencia"
        flexPrefat.Col = 19
        flexPrefat.Text = "Tx Coleta"
        flexPrefat.Col = 20
        flexPrefat.Text = "Tx Entrega/Red"
        flexPrefat.Col = 21
        flexPrefat.Text = "Pedagio"
        flexPrefat.Col = 22
        flexPrefat.Text = "Tx Outros"
        flexPrefat.Col = 23
        flexPrefat.Text = "Data CTC"
        flexPrefat.Col = 24
        flexPrefat.Text = "Modal"
        flexPrefat.Col = 25
        flexPrefat.Text = "Obs Emissao"
        flexPrefat.Col = 26
        flexPrefat.Text = "URG/PRI/NOR"
        flexPrefat.Col = 27
        flexPrefat.Text = "Aliq_ICMS"
        flexPrefat.Col = 28
        flexPrefat.Text = "Tabela"
        flexPrefat.Col = 29
        flexPrefat.Text = "Natureza"
        
        flexPrefat.ColWidth(0) = 200
        flexPrefat.ColWidth(1) = 400
        flexPrefat.ColWidth(2) = 800
        flexPrefat.ColWidth(3) = 1900
        flexPrefat.ColWidth(4) = 1900
        flexPrefat.ColWidth(5) = 1900
        flexPrefat.ColWidth(6) = 1900
        flexPrefat.ColWidth(7) = 400
        flexPrefat.ColWidth(8) = 2500
        flexPrefat.ColWidth(9) = 1200
        flexPrefat.ColWidth(10) = 800
        flexPrefat.ColWidth(11) = 600
        flexPrefat.ColWidth(12) = 600
        flexPrefat.ColWidth(13) = 800
        flexPrefat.ColWidth(14) = 1200
        flexPrefat.ColWidth(15) = 1200
        flexPrefat.ColWidth(16) = 1200
        flexPrefat.ColWidth(17) = 1200
        flexPrefat.ColWidth(18) = 1200
        flexPrefat.ColWidth(19) = 1200
        flexPrefat.ColWidth(20) = 1200
        flexPrefat.ColWidth(21) = 1200
        flexPrefat.ColWidth(22) = 1200
        flexPrefat.ColWidth(23) = 1200
        flexPrefat.ColWidth(24) = 1200
        flexPrefat.ColWidth(25) = 4000
        flexPrefat.ColWidth(26) = 1400
        flexPrefat.ColWidth(27) = 1200
        flexPrefat.ColWidth(28) = 1200
        flexPrefat.ColWidth(29) = 1200
        
        If de_informa.rsSel_GerarArqPrefat.RecordCount > 0 Then
            
            flexPrefat.Rows = de_informa.rsSel_GerarArqPrefat.RecordCount + 2
            flexPrefat.FixedRows = 1
            lblClienteCNPJ = de_informa.rsSel_GerarArqPrefat.Fields("cliente_cgc")
            lblCliente = de_informa.rsSel_GerarArqPrefat.Fields("cliente_nome")
            lblVencto = de_informa.rsSel_GerarArqPrefat.Fields("vencimento")
            xValorFatura = 0
            
            For xcont = 1 To de_informa.rsSel_GerarArqPrefat.RecordCount
                
                xValorFatura = xValorFatura + de_informa.rsSel_GerarArqPrefat.Fields("frete")
                flexPrefat.TextMatrix(xcont, 1) = Mid$(de_informa.rsSel_GerarArqPrefat.Fields("filialctc"), 1, 2)
                flexPrefat.TextMatrix(xcont, 2) = Mid$(de_informa.rsSel_GerarArqPrefat.Fields("filialctc"), 3, 8)
                flexPrefat.TextMatrix(xcont, 3) = de_informa.rsSel_GerarArqPrefat.Fields("remet_nome")
                flexPrefat.TextMatrix(xcont, 4) = Trim$(de_informa.rsSel_GerarArqPrefat.Fields("remet_cidade")) & "-" & de_informa.rsSel_GerarArqPrefat.Fields("remet_uf")
                flexPrefat.TextMatrix(xcont, 5) = de_informa.rsSel_GerarArqPrefat.Fields("dest_nome")
                flexPrefat.TextMatrix(xcont, 6) = Trim$(de_informa.rsSel_GerarArqPrefat.Fields("dest_cidade"))
                flexPrefat.TextMatrix(xcont, 7) = Trim$(de_informa.rsSel_GerarArqPrefat.Fields("dest_uf"))
                flexPrefat.TextMatrix(xcont, 8) = de_informa.rsSel_GerarArqPrefat.Fields("nfs")
                flexPrefat.TextMatrix(xcont, 9) = de_informa.rsSel_GerarArqPrefat.Fields("valmerc")
                flexPrefat.TextMatrix(xcont, 10) = de_informa.rsSel_GerarArqPrefat.Fields("volumes")
                flexPrefat.TextMatrix(xcont, 11) = de_informa.rsSel_GerarArqPrefat.Fields("peso")
                flexPrefat.TextMatrix(xcont, 12) = de_informa.rsSel_GerarArqPrefat.Fields("pesotax")
                flexPrefat.TextMatrix(xcont, 13) = de_informa.rsSel_GerarArqPrefat.Fields("frete")
                flexPrefat.TextMatrix(xcont, 14) = de_informa.rsSel_GerarArqPrefat.Fields("fretetotalbruto")
                
                If de_informa.rsSel_GerarArqPrefat.Fields("frete") = de_informa.rsSel_GerarArqPrefat.Fields("fretetotal") Then
                    
                    flexPrefat.TextMatrix(xcont, 15) = de_informa.rsSel_GerarArqPrefat.Fields("fretepeso")
                    flexPrefat.TextMatrix(xcont, 16) = de_informa.rsSel_GerarArqPrefat.Fields("fretevalor")
                    flexPrefat.TextMatrix(xcont, 17) = de_informa.rsSel_GerarArqPrefat.Fields("gris")
                    flexPrefat.TextMatrix(xcont, 18) = de_informa.rsSel_GerarArqPrefat.Fields("txurgencia")
                    flexPrefat.TextMatrix(xcont, 19) = de_informa.rsSel_GerarArqPrefat.Fields("txcoleta")
                    flexPrefat.TextMatrix(xcont, 20) = de_informa.rsSel_GerarArqPrefat.Fields("txentregared")
                    flexPrefat.TextMatrix(xcont, 21) = de_informa.rsSel_GerarArqPrefat.Fields("pedagio")
                    flexPrefat.TextMatrix(xcont, 22) = de_informa.rsSel_GerarArqPrefat.Fields("txoutros")
                    flexPrefat.TextMatrix(xcont, 27) = de_informa.rsSel_GerarArqPrefat.Fields("fretetotalbruto") - de_informa.rsSel_GerarArqPrefat.Fields("fretetotal")
                    flexPrefat.TextMatrix(xcont, 28) = de_informa.rsSel_GerarArqPrefat.Fields("tabfrete")
                    flexPrefat.TextMatrix(xcont, 29) = de_informa.rsSel_GerarArqPrefat.Fields("naturezaobs")
                
                Else
                    
                    flexPrefat.TextMatrix(xcont, 15) = de_informa.rsSel_GerarArqPrefat.Fields("fretepesobr")
                    flexPrefat.TextMatrix(xcont, 16) = de_informa.rsSel_GerarArqPrefat.Fields("fretevalorbr")
                    flexPrefat.TextMatrix(xcont, 17) = de_informa.rsSel_GerarArqPrefat.Fields("grisbr")
                    flexPrefat.TextMatrix(xcont, 18) = de_informa.rsSel_GerarArqPrefat.Fields("txurgenciabr")
                    flexPrefat.TextMatrix(xcont, 19) = de_informa.rsSel_GerarArqPrefat.Fields("txcoletabr")
                    flexPrefat.TextMatrix(xcont, 20) = de_informa.rsSel_GerarArqPrefat.Fields("txentregaredbr")
                    flexPrefat.TextMatrix(xcont, 21) = de_informa.rsSel_GerarArqPrefat.Fields("pedagiobr")
                    flexPrefat.TextMatrix(xcont, 22) = de_informa.rsSel_GerarArqPrefat.Fields("txoutrosbr")
                    flexPrefat.TextMatrix(xcont, 27) = de_informa.rsSel_GerarArqPrefat.Fields("fretetotalbruto") - de_informa.rsSel_GerarArqPrefat.Fields("fretetotal")
                    flexPrefat.TextMatrix(xcont, 28) = de_informa.rsSel_GerarArqPrefat.Fields("tabfrete")
                    flexPrefat.TextMatrix(xcont, 29) = de_informa.rsSel_GerarArqPrefat.Fields("naturezaobs")
                
                End If
                
                flexPrefat.TextMatrix(xcont, 23) = de_informa.rsSel_GerarArqPrefat.Fields("data")
                flexPrefat.TextMatrix(xcont, 24) = de_informa.rsSel_GerarArqPrefat.Fields("modal")
                flexPrefat.TextMatrix(xcont, 25) = de_informa.rsSel_GerarArqPrefat.Fields("obs_emissao")
                flexPrefat.TextMatrix(xcont, 26) = de_informa.rsSel_GerarArqPrefat.Fields("prioridade")
                
                de_informa.rsSel_GerarArqPrefat.MoveNext
                
            Next
            
            lblValorFatura = Format(xValorFatura, "#,###,##0.00")
        
        Else
            
            MsgBox "Arquivo de Pré-Fatura Não Encontrado !"
            Exit Sub
        
        End If
        
    ElseIf optModelo2 = True Then
            
        flexPrefat.Rows = 2
        flexPrefat.Cols = 20
        flexPrefat.Row = 0
        flexPrefat.Col = 1
        flexPrefat.Text = "Pre-Fatura"
        flexPrefat.Col = 2
        flexPrefat.Text = "Cliente_Nome"
        flexPrefat.Col = 3
        flexPrefat.Text = "Vencimento"
        flexPrefat.Col = 4
        flexPrefat.Text = "FilialCTC"
        flexPrefat.Col = 5
        flexPrefat.Text = "NFS"
        flexPrefat.Col = 6
        flexPrefat.Text = "Data"
        flexPrefat.Col = 7
        flexPrefat.Text = "Dest_Nome"
        flexPrefat.Col = 8
        flexPrefat.Text = "Endereco"
        flexPrefat.Col = 9
        flexPrefat.Text = "Cidade"
        flexPrefat.Col = 10
        flexPrefat.Text = "UF"
        flexPrefat.Col = 11
        flexPrefat.Text = "Natureza"
        flexPrefat.Col = 12
        flexPrefat.Text = "Prioridade"
        flexPrefat.Col = 13
        flexPrefat.Text = "Peso"
        flexPrefat.Col = 14
        flexPrefat.Text = "Peso Tax."
        flexPrefat.Col = 15
        flexPrefat.Text = "Valmerc."
        flexPrefat.Col = 16
        flexPrefat.Text = "Regiao"
        flexPrefat.Col = 17
        flexPrefat.Text = "Modal"
        flexPrefat.Col = 18
        flexPrefat.Text = "Frete"
        flexPrefat.Col = 19
        flexPrefat.Text = "Remetente"
            
        flexPrefat.ColWidth(0) = 200
        flexPrefat.ColWidth(1) = 1600
        flexPrefat.ColWidth(2) = 1600
        flexPrefat.ColWidth(3) = 1600
        flexPrefat.ColWidth(4) = 1600
        flexPrefat.ColWidth(5) = 1600
        flexPrefat.ColWidth(6) = 1600
        flexPrefat.ColWidth(7) = 1600
        flexPrefat.ColWidth(8) = 1600
        flexPrefat.ColWidth(9) = 1600
        flexPrefat.ColWidth(10) = 1600
        flexPrefat.ColWidth(11) = 1600
        flexPrefat.ColWidth(12) = 1600
        flexPrefat.ColWidth(13) = 1600
        flexPrefat.ColWidth(14) = 1600
        flexPrefat.ColWidth(15) = 1600
        flexPrefat.ColWidth(16) = 1600
        flexPrefat.ColWidth(17) = 1600
        flexPrefat.ColWidth(18) = 1600
        flexPrefat.ColWidth(19) = 1600
            
        If de_informa.rsSel_GerarArqPrefat.RecordCount > 0 Then
            flexPrefat.Rows = de_informa.rsSel_GerarArqPrefat.RecordCount + 2
            flexPrefat.FixedRows = 1
            lblClienteCNPJ = de_informa.rsSel_GerarArqPrefat.Fields("cliente_cgc")
            lblCliente = de_informa.rsSel_GerarArqPrefat.Fields("cliente_nome")
            lblVencto = de_informa.rsSel_GerarArqPrefat.Fields("vencimento")
            xValorFatura = 0
            For xcont = 1 To de_informa.rsSel_GerarArqPrefat.RecordCount
                
                xValorFatura = xValorFatura + de_informa.rsSel_GerarArqPrefat.Fields("frete")
                flexPrefat.TextMatrix(xcont, 1) = de_informa.rsSel_GerarArqPrefat.Fields("filialprefatura")
                flexPrefat.TextMatrix(xcont, 2) = de_informa.rsSel_GerarArqPrefat.Fields("remet_nome")
                flexPrefat.TextMatrix(xcont, 3) = de_informa.rsSel_GerarArqPrefat.Fields("vencimento")
                flexPrefat.TextMatrix(xcont, 4) = Trim$(de_informa.rsSel_GerarArqPrefat.Fields("filialctc"))
                flexPrefat.TextMatrix(xcont, 5) = de_informa.rsSel_GerarArqPrefat.Fields("nfs")
                flexPrefat.TextMatrix(xcont, 6) = Trim$(de_informa.rsSel_GerarArqPrefat.Fields("data"))
                flexPrefat.TextMatrix(xcont, 7) = Trim$(de_informa.rsSel_GerarArqPrefat.Fields("dest_nome"))
                flexPrefat.TextMatrix(xcont, 8) = de_informa.rsSel_GerarArqPrefat.Fields("dest_end")
                flexPrefat.TextMatrix(xcont, 9) = de_informa.rsSel_GerarArqPrefat.Fields("dest_cidade")
                flexPrefat.TextMatrix(xcont, 10) = de_informa.rsSel_GerarArqPrefat.Fields("dest_uf")
                flexPrefat.TextMatrix(xcont, 11) = de_informa.rsSel_GerarArqPrefat.Fields("natureza")
                flexPrefat.TextMatrix(xcont, 12) = de_informa.rsSel_GerarArqPrefat.Fields("prioridade")
                flexPrefat.TextMatrix(xcont, 13) = de_informa.rsSel_GerarArqPrefat.Fields("peso")
                flexPrefat.TextMatrix(xcont, 14) = de_informa.rsSel_GerarArqPrefat.Fields("pesotax")
                flexPrefat.TextMatrix(xcont, 15) = de_informa.rsSel_GerarArqPrefat.Fields("valmerc")
                flexPrefat.TextMatrix(xcont, 16) = de_informa.rsSel_GerarArqPrefat.Fields("regiao")
                flexPrefat.TextMatrix(xcont, 17) = de_informa.rsSel_GerarArqPrefat.Fields("modal")
                flexPrefat.TextMatrix(xcont, 18) = de_informa.rsSel_GerarArqPrefat.Fields("frete")
                flexPrefat.TextMatrix(xcont, 19) = de_informa.rsSel_GerarArqPrefat.Fields("remet_nome")
                
                de_informa.rsSel_GerarArqPrefat.MoveNext
                
            Next
            
            lblValorFatura = Format(xValorFatura, "#,###,##0.00")
        
        Else
            MsgBox "Arquivo de Pré-Fatura Não Encontrado !"
            Exit Sub
        End If
    
    ElseIf optModelo3 = True Then
        flexPrefat.Rows = 2
        flexPrefat.Cols = 29
        
        flexPrefat.Row = 0
        flexPrefat.Col = 1
        flexPrefat.Text = "Filial"
        flexPrefat.Col = 2
        flexPrefat.Text = "CTC"
        flexPrefat.Col = 3
        flexPrefat.Text = "Remetente"
        flexPrefat.Col = 4
        flexPrefat.Text = "Origem"
        flexPrefat.Col = 5
        flexPrefat.Text = "Destinatario"
        flexPrefat.Col = 6
        flexPrefat.Text = "Cidade Destino"
        flexPrefat.Col = 7
        flexPrefat.Text = "UF"
        flexPrefat.Col = 8
        flexPrefat.Text = "NFS"
        flexPrefat.Col = 9
        flexPrefat.Text = "Valor Merc."
        flexPrefat.Col = 10
        flexPrefat.Text = "Volumes"
        flexPrefat.Col = 11
        flexPrefat.Text = "Peso"
        flexPrefat.Col = 12
        flexPrefat.Text = "Peso Tax."
        flexPrefat.Col = 13
        flexPrefat.Text = "Frete Total Liq."
        flexPrefat.Col = 14
        flexPrefat.Text = "Frete Total Bruto"
        flexPrefat.Col = 15
        flexPrefat.Text = "Frete Peso"
        flexPrefat.Col = 16
        flexPrefat.Text = "Frete Valor"
        flexPrefat.Col = 17
        flexPrefat.Text = "Gris"
        flexPrefat.Col = 18
        flexPrefat.Text = "Tx Urgencia"
        flexPrefat.Col = 19
        flexPrefat.Text = "Tx Coleta"
        flexPrefat.Col = 20
        flexPrefat.Text = "Tx Entrega/Red"
        flexPrefat.Col = 21
        flexPrefat.Text = "Pedagio"
        flexPrefat.Col = 22
        flexPrefat.Text = "Tx Outros"
        
        flexPrefat.Col = 23
        flexPrefat.Text = "Aliquota"
        flexPrefat.Col = 24
        flexPrefat.Text = "ICMS"
        
        flexPrefat.Col = 25
        flexPrefat.Text = "Data CTC"
        flexPrefat.Col = 26
        flexPrefat.Text = "Modal"
        flexPrefat.Col = 27
        flexPrefat.Text = "Obs Emissao"
        flexPrefat.Col = 28
        flexPrefat.Text = "Natureza"
            
        flexPrefat.ColWidth(0) = 200
        flexPrefat.ColWidth(1) = 400
        flexPrefat.ColWidth(2) = 800
        flexPrefat.ColWidth(3) = 1900
        flexPrefat.ColWidth(4) = 1900
        flexPrefat.ColWidth(5) = 1900
        flexPrefat.ColWidth(6) = 1900
        flexPrefat.ColWidth(7) = 400
        flexPrefat.ColWidth(8) = 2500
        flexPrefat.ColWidth(9) = 1200
        flexPrefat.ColWidth(10) = 800
        flexPrefat.ColWidth(11) = 600
        flexPrefat.ColWidth(12) = 600
        flexPrefat.ColWidth(13) = 800
        flexPrefat.ColWidth(14) = 1200
        flexPrefat.ColWidth(15) = 1200
        flexPrefat.ColWidth(16) = 1200
        flexPrefat.ColWidth(17) = 1200
        flexPrefat.ColWidth(18) = 1200
        flexPrefat.ColWidth(19) = 1200
        flexPrefat.ColWidth(20) = 1200
        flexPrefat.ColWidth(21) = 1200
        flexPrefat.ColWidth(22) = 1200
        flexPrefat.ColWidth(23) = 1200
        flexPrefat.ColWidth(24) = 1200
        flexPrefat.ColWidth(25) = 1200
        flexPrefat.ColWidth(26) = 1400
        flexPrefat.ColWidth(27) = 4000
        flexPrefat.ColWidth(28) = 800
        
        If de_informa.rsSel_GerarArqPrefat.RecordCount > 0 Then
            flexPrefat.Rows = de_informa.rsSel_GerarArqPrefat.RecordCount + 2
            flexPrefat.FixedRows = 1
            lblClienteCNPJ = de_informa.rsSel_GerarArqPrefat.Fields("cliente_cgc")
            lblCliente = de_informa.rsSel_GerarArqPrefat.Fields("cliente_nome")
            lblVencto = de_informa.rsSel_GerarArqPrefat.Fields("vencimento")
            xValorFatura = 0
            For xcont = 1 To de_informa.rsSel_GerarArqPrefat.RecordCount
                xValorFatura = xValorFatura + de_informa.rsSel_GerarArqPrefat.Fields("frete")
                flexPrefat.TextMatrix(xcont, 1) = Mid$(de_informa.rsSel_GerarArqPrefat.Fields("filialctc"), 1, 2)
                flexPrefat.TextMatrix(xcont, 2) = Mid$(de_informa.rsSel_GerarArqPrefat.Fields("filialctc"), 3, 8)
                flexPrefat.TextMatrix(xcont, 3) = de_informa.rsSel_GerarArqPrefat.Fields("remet_nome")
                flexPrefat.TextMatrix(xcont, 4) = Trim$(de_informa.rsSel_GerarArqPrefat.Fields("remet_cidade")) & "-" & de_informa.rsSel_GerarArqPrefat.Fields("remet_uf")
                flexPrefat.TextMatrix(xcont, 5) = de_informa.rsSel_GerarArqPrefat.Fields("dest_nome")
                flexPrefat.TextMatrix(xcont, 6) = Trim$(de_informa.rsSel_GerarArqPrefat.Fields("dest_cidade"))
                flexPrefat.TextMatrix(xcont, 7) = Trim$(de_informa.rsSel_GerarArqPrefat.Fields("dest_uf"))
                flexPrefat.TextMatrix(xcont, 8) = de_informa.rsSel_GerarArqPrefat.Fields("nfs")
                flexPrefat.TextMatrix(xcont, 9) = de_informa.rsSel_GerarArqPrefat.Fields("valmerc")
                flexPrefat.TextMatrix(xcont, 10) = de_informa.rsSel_GerarArqPrefat.Fields("volumes")
                flexPrefat.TextMatrix(xcont, 11) = de_informa.rsSel_GerarArqPrefat.Fields("peso")
                flexPrefat.TextMatrix(xcont, 12) = de_informa.rsSel_GerarArqPrefat.Fields("pesotax")
                flexPrefat.TextMatrix(xcont, 13) = de_informa.rsSel_GerarArqPrefat.Fields("frete")
                flexPrefat.TextMatrix(xcont, 14) = de_informa.rsSel_GerarArqPrefat.Fields("fretetotalbruto")
                If de_informa.rsSel_GerarArqPrefat.Fields("frete") = de_informa.rsSel_GerarArqPrefat.Fields("fretetotal") Then
                    flexPrefat.TextMatrix(xcont, 15) = de_informa.rsSel_GerarArqPrefat.Fields("fretepeso")
                    flexPrefat.TextMatrix(xcont, 16) = de_informa.rsSel_GerarArqPrefat.Fields("fretevalor")
                    flexPrefat.TextMatrix(xcont, 17) = de_informa.rsSel_GerarArqPrefat.Fields("gris")
                    flexPrefat.TextMatrix(xcont, 18) = de_informa.rsSel_GerarArqPrefat.Fields("txurgencia")
                    flexPrefat.TextMatrix(xcont, 19) = de_informa.rsSel_GerarArqPrefat.Fields("txcoleta")
                    flexPrefat.TextMatrix(xcont, 20) = de_informa.rsSel_GerarArqPrefat.Fields("txentregared")
                    flexPrefat.TextMatrix(xcont, 21) = de_informa.rsSel_GerarArqPrefat.Fields("pedagio")
                    flexPrefat.TextMatrix(xcont, 22) = de_informa.rsSel_GerarArqPrefat.Fields("txoutros")
                Else
                    flexPrefat.TextMatrix(xcont, 15) = de_informa.rsSel_GerarArqPrefat.Fields("fretepesobr")
                    flexPrefat.TextMatrix(xcont, 16) = de_informa.rsSel_GerarArqPrefat.Fields("fretevalorbr")
                    flexPrefat.TextMatrix(xcont, 17) = de_informa.rsSel_GerarArqPrefat.Fields("grisbr")
                    flexPrefat.TextMatrix(xcont, 18) = de_informa.rsSel_GerarArqPrefat.Fields("txurgenciabr")
                    flexPrefat.TextMatrix(xcont, 19) = de_informa.rsSel_GerarArqPrefat.Fields("txcoletabr")
                    flexPrefat.TextMatrix(xcont, 20) = de_informa.rsSel_GerarArqPrefat.Fields("txentregaredbr")
                    flexPrefat.TextMatrix(xcont, 21) = de_informa.rsSel_GerarArqPrefat.Fields("pedagiobr")
                    flexPrefat.TextMatrix(xcont, 22) = de_informa.rsSel_GerarArqPrefat.Fields("txoutrosbr")
                End If
                flexPrefat.TextMatrix(xcont, 23) = de_informa.rsSel_GerarArqPrefat.Fields("Aliquota")
                flexPrefat.TextMatrix(xcont, 24) = de_informa.rsSel_GerarArqPrefat.Fields("ICMS")
                flexPrefat.TextMatrix(xcont, 25) = de_informa.rsSel_GerarArqPrefat.Fields("data")
                flexPrefat.TextMatrix(xcont, 26) = de_informa.rsSel_GerarArqPrefat.Fields("modal")
                flexPrefat.TextMatrix(xcont, 27) = de_informa.rsSel_GerarArqPrefat.Fields("obs_emissao")
                flexPrefat.TextMatrix(xcont, 28) = de_informa.rsSel_GerarArqPrefat.Fields("naturezaobs")
                
                de_informa.rsSel_GerarArqPrefat.MoveNext
                
            Next
            lblValorFatura = Format(xValorFatura, "#,###,##0.00")
        Else
            MsgBox "Arquivo de Pré-Fatura Não Encontrado !"
            Exit Sub
        End If
    
    End If
End Sub

Private Sub cmdGeraArq_Click()
    Dim xcont As Long, xFiles As String, xlinha As String
    If flexPrefat.Rows < 3 Then
        MsgBox "Não há Dados Para Geração de Arquivos !", vbExclamation, "Ops"
    Else
        
        xFiles = "C:\PREFATURA\" & Trim$(Mid$(lblCliente, 1, 5)) & "_" & Trim$(txtPreFat) & "_" & zeros(Day(Date), 2) & zeros(Month(Date), 2) & "_" & Mid$(Trim$(CVar(Time())), 1, 2) & Mid$(Trim$(CVar(Time())), 4, 2) & ".txt"
        
        Open xFiles For Output As #1
    
        xlinha = "PRE-FATURA: " & "#" & txtPreFat & "###" & "VENCTO: " & lblVencto & "#"
        Print #1, xlinha
        xlinha = "VALOR R$: " & "#" & CDbl(SoNumeros(lblValorFatura)) / 100 & "#"
        Print #1, xlinha
        Print #1, "#"
        
        If optModelo1 = True Then
        
            For xcont = 0 To flexPrefat.Rows - 1
            
                xlinha = flexPrefat.TextMatrix(xcont, 1) & "#" & flexPrefat.TextMatrix(xcont, 2) & "#" & _
                        flexPrefat.TextMatrix(xcont, 3) & "#" & flexPrefat.TextMatrix(xcont, 4) & "#" & _
                        flexPrefat.TextMatrix(xcont, 5) & "#" & flexPrefat.TextMatrix(xcont, 6) & "#" & _
                        flexPrefat.TextMatrix(xcont, 7) & "#" & flexPrefat.TextMatrix(xcont, 8) & "#" & _
                        flexPrefat.TextMatrix(xcont, 9) & "#" & flexPrefat.TextMatrix(xcont, 10) & "#" & _
                        flexPrefat.TextMatrix(xcont, 11) & "#" & flexPrefat.TextMatrix(xcont, 12) & "#" & _
                        flexPrefat.TextMatrix(xcont, 13) & "#" & flexPrefat.TextMatrix(xcont, 14) & "#" & _
                        flexPrefat.TextMatrix(xcont, 15) & "#" & flexPrefat.TextMatrix(xcont, 16) & "#" & _
                        flexPrefat.TextMatrix(xcont, 17) & "#" & flexPrefat.TextMatrix(xcont, 18) & "#" & _
                        flexPrefat.TextMatrix(xcont, 19) & "#" & flexPrefat.TextMatrix(xcont, 20) & "#" & _
                        flexPrefat.TextMatrix(xcont, 21) & "#" & flexPrefat.TextMatrix(xcont, 22) & "#" & _
                        flexPrefat.TextMatrix(xcont, 23) & "#" & flexPrefat.TextMatrix(xcont, 24) & "#" & _
                        flexPrefat.TextMatrix(xcont, 25) & "#" & flexPrefat.TextMatrix(xcont, 26) & "#" & _
                        flexPrefat.TextMatrix(xcont, 27) & "#" & flexPrefat.TextMatrix(xcont, 28) & "#" & _
                        flexPrefat.TextMatrix(xcont, 27) & "#" & flexPrefat.TextMatrix(xcont, 29)
                        
                Print #1, xlinha
                DoEvents
            Next
        
        ElseIf optModelo2 = True Then
        
            For xcont = 0 To flexPrefat.Rows - 1
            
                xlinha = flexPrefat.TextMatrix(xcont, 1) & "#" & flexPrefat.TextMatrix(xcont, 2) & "#" & _
                        flexPrefat.TextMatrix(xcont, 3) & "#" & flexPrefat.TextMatrix(xcont, 4) & "#" & _
                        flexPrefat.TextMatrix(xcont, 5) & "#" & flexPrefat.TextMatrix(xcont, 6) & "#" & _
                        flexPrefat.TextMatrix(xcont, 7) & "#" & flexPrefat.TextMatrix(xcont, 8) & "#" & _
                        flexPrefat.TextMatrix(xcont, 9) & "#" & flexPrefat.TextMatrix(xcont, 10) & "#" & _
                        flexPrefat.TextMatrix(xcont, 11) & "#" & flexPrefat.TextMatrix(xcont, 12) & "#" & _
                        flexPrefat.TextMatrix(xcont, 13) & "#" & flexPrefat.TextMatrix(xcont, 14) & "#" & _
                        flexPrefat.TextMatrix(xcont, 15) & "#" & flexPrefat.TextMatrix(xcont, 16) & "#" & _
                        flexPrefat.TextMatrix(xcont, 17) & "#" & flexPrefat.TextMatrix(xcont, 18) & "#" & _
                        flexPrefat.TextMatrix(xcont, 19) & "#"
                Print #1, xlinha
                DoEvents
            Next
        
        ElseIf optModelo3 = True Then
            For xcont = 0 To flexPrefat.Rows - 1
                xlinha = flexPrefat.TextMatrix(xcont, 1) & "#" & flexPrefat.TextMatrix(xcont, 2) & "#" & _
                        flexPrefat.TextMatrix(xcont, 3) & "#" & flexPrefat.TextMatrix(xcont, 4) & "#" & _
                        flexPrefat.TextMatrix(xcont, 5) & "#" & flexPrefat.TextMatrix(xcont, 6) & "#" & _
                        flexPrefat.TextMatrix(xcont, 7) & "#" & flexPrefat.TextMatrix(xcont, 8) & "#" & _
                        flexPrefat.TextMatrix(xcont, 9) & "#" & flexPrefat.TextMatrix(xcont, 10) & "#" & _
                        flexPrefat.TextMatrix(xcont, 11) & "#" & flexPrefat.TextMatrix(xcont, 12) & "#" & _
                        flexPrefat.TextMatrix(xcont, 13) & "#" & flexPrefat.TextMatrix(xcont, 14) & "#" & _
                        flexPrefat.TextMatrix(xcont, 15) & "#" & flexPrefat.TextMatrix(xcont, 16) & "#" & _
                        flexPrefat.TextMatrix(xcont, 17) & "#" & flexPrefat.TextMatrix(xcont, 18) & "#" & _
                        flexPrefat.TextMatrix(xcont, 19) & "#" & flexPrefat.TextMatrix(xcont, 20) & "#" & _
                        flexPrefat.TextMatrix(xcont, 21) & "#" & flexPrefat.TextMatrix(xcont, 22) & "#" & _
                        flexPrefat.TextMatrix(xcont, 23) & "#" & flexPrefat.TextMatrix(xcont, 24) & "#" & _
                        flexPrefat.TextMatrix(xcont, 25) & "#" & flexPrefat.TextMatrix(xcont, 26) & "#" & _
                        flexPrefat.TextMatrix(xcont, 27) & "#"
                Print #1, xlinha
                DoEvents
            Next
        
        End If
        
        Close #1
        
        MsgBox "OK ! Arquivo Gerado em " & xFiles & "." & Chr(13) + Chr(10) + Chr(13) + Chr(10) + _
        "O Arquivo Gerado é do Tipo Texto ( TXT com Delimitador # ) e você pode abrí-lo em diversos aplicativos. Para Abrí-lo no MS-Excel, em ABRIR escolha ARQUIVOS DO TIPO = Arquivos de Texto e selecione o arquivo no local indicado acima. Na Caixa ASSISTENTE DE IMPORTAÇÃO escolha DELIMITADO e o caracter delimitador escolha OUTROS e digite # . Clique em Concluir e o arquivo será importado para o MS-Excel.", vbInformation, "Geração de Arquivo TXT"

    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    flexPrefat.Rows = 2
    flexPrefat.Cols = 27
    flexPrefat.Row = 0
    flexPrefat.Col = 1
    flexPrefat.Text = "Filial"
    flexPrefat.Col = 2
    flexPrefat.Text = "CTC"
    flexPrefat.Col = 3
    flexPrefat.Text = "Remetente"
    flexPrefat.Col = 4
    flexPrefat.Text = "Origem"
    flexPrefat.Col = 5
    flexPrefat.Text = "Destinatario"
    flexPrefat.Col = 6
    flexPrefat.Text = "Cidade Destino"
    flexPrefat.Col = 7
    flexPrefat.Text = "UF"
    flexPrefat.Col = 8
    flexPrefat.Text = "NFS"
    flexPrefat.Col = 9
    flexPrefat.Text = "Valor Merc."
    flexPrefat.Col = 10
    flexPrefat.Text = "Volumes"
    flexPrefat.Col = 11
    flexPrefat.Text = "Peso"
    flexPrefat.Col = 12
    flexPrefat.Text = "Peso Tax."
    flexPrefat.Col = 13
    flexPrefat.Text = "Frete Total Liq."
    flexPrefat.Col = 14
    flexPrefat.Text = "Frete Total Bruto"
    flexPrefat.Col = 15
    flexPrefat.Text = "Frete Peso"
    flexPrefat.Col = 16
    flexPrefat.Text = "Frete Valor"
    flexPrefat.Col = 17
    flexPrefat.Text = "Gris"
    flexPrefat.Col = 18
    flexPrefat.Text = "Tx Urgencia"
    flexPrefat.Col = 19
    flexPrefat.Text = "Tx Coleta"
    flexPrefat.Col = 20
    flexPrefat.Text = "Tx Entrega/Red"
    flexPrefat.Col = 21
    flexPrefat.Text = "Pedagio"
    flexPrefat.Col = 22
    flexPrefat.Text = "Tx Outros"
    flexPrefat.Col = 23
    flexPrefat.Text = "Data CTC"
    flexPrefat.Col = 24
    flexPrefat.Text = "Modal"
    flexPrefat.Col = 25
    flexPrefat.Text = "Obs Emissao"
    flexPrefat.Col = 26
    flexPrefat.Text = "URG/PRI/NOR"
        
    flexPrefat.ColWidth(0) = 200
    flexPrefat.ColWidth(1) = 400
    flexPrefat.ColWidth(2) = 800
    flexPrefat.ColWidth(3) = 1900
    flexPrefat.ColWidth(4) = 1900
    flexPrefat.ColWidth(5) = 1900
    flexPrefat.ColWidth(6) = 1900
    flexPrefat.ColWidth(7) = 400
    flexPrefat.ColWidth(8) = 2500
    flexPrefat.ColWidth(9) = 1200
    flexPrefat.ColWidth(10) = 800
    flexPrefat.ColWidth(11) = 600
    flexPrefat.ColWidth(12) = 600
    flexPrefat.ColWidth(13) = 800
    flexPrefat.ColWidth(14) = 1200
    flexPrefat.ColWidth(15) = 1200
    flexPrefat.ColWidth(16) = 1200
    flexPrefat.ColWidth(17) = 1200
    flexPrefat.ColWidth(18) = 1200
    flexPrefat.ColWidth(19) = 1200
    flexPrefat.ColWidth(20) = 1200
    flexPrefat.ColWidth(21) = 1200
    flexPrefat.ColWidth(22) = 1200
    flexPrefat.ColWidth(23) = 1200
    flexPrefat.ColWidth(24) = 1200
    flexPrefat.ColWidth(25) = 4000
    flexPrefat.ColWidth(26) = 1400
    
End Sub

Private Sub optModelo1_Click()
    flexPrefat.Rows = 1
    flexPrefat.Rows = 2
    flexPrefat.FixedRows = 1

End Sub

Private Sub optModelo2_Click()
    flexPrefat.Rows = 1
    flexPrefat.Rows = 2
    flexPrefat.FixedRows = 1

End Sub

Private Sub txtPreFat_Change()
    If Not IsNumeric(txtPreFat) Then
        SendKeys "{BACKSPACE}"
    End If
End Sub

Private Sub txtPreFat_GotFocus()
    txtPreFat.SelStart = 0
    txtPreFat.SelLength = 8
End Sub

Private Sub txtPreFat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim$(txtPreFat)) > 0 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
Private Sub lixo()


    flexPrefat.Rows = 2
    flexPrefat.Cols = 19
    flexPrefat.Row = 0
    flexPrefat.Col = 1
    flexPrefat.Text = "Filial-CTC"
    flexPrefat.Col = 2
    flexPrefat.Text = "Data"
    flexPrefat.Col = 3
    flexPrefat.Text = "Remetente"
    flexPrefat.Col = 4
    flexPrefat.Text = "Destinatário"
    flexPrefat.Col = 5
    flexPrefat.Text = "Frete Cobr."
    flexPrefat.Col = 6
    flexPrefat.Text = "NFs"
    flexPrefat.Col = 7
    flexPrefat.Text = "Valor Merc."
    flexPrefat.Col = 8
    flexPrefat.Text = "Peso"
    flexPrefat.Col = 9
    flexPrefat.Text = "Frete Peso"
    flexPrefat.Col = 10
    flexPrefat.Text = "Frete Valor"
    flexPrefat.Col = 11
    flexPrefat.Text = "Gris"
    flexPrefat.Col = 12
    flexPrefat.Text = "Tx.Coleta"
    flexPrefat.Col = 13
    flexPrefat.Text = "Tx.Entrega"
    flexPrefat.Col = 14
    flexPrefat.Text = "Tx.Urgencia"
    flexPrefat.Col = 15
    flexPrefat.Text = "Pedágio"
    flexPrefat.Col = 16
    flexPrefat.Text = "Tx.Outros"
    flexPrefat.Col = 17
    flexPrefat.Text = "Descr.Outros"
    flexPrefat.Col = 18
    flexPrefat.Text = "Emissor"
        
    flexPrefat.ColWidth(0) = 150
    flexPrefat.ColWidth(1) = 1000
    flexPrefat.ColWidth(2) = 800
    flexPrefat.ColWidth(3) = 1000
    flexPrefat.ColWidth(4) = 1000
    flexPrefat.ColWidth(5) = 1000
    flexPrefat.ColWidth(6) = 1000
    flexPrefat.ColWidth(7) = 1000
    flexPrefat.ColWidth(8) = 1000
    flexPrefat.ColWidth(9) = 1000
    flexPrefat.ColWidth(10) = 1000
    flexPrefat.ColWidth(11) = 1000
    flexPrefat.ColWidth(12) = 1000
    flexPrefat.ColWidth(13) = 1000
    flexPrefat.ColWidth(14) = 1000
    flexPrefat.ColWidth(15) = 1000
    flexPrefat.ColWidth(16) = 1000
    flexPrefat.ColWidth(17) = 1000
    flexPrefat.ColWidth(18) = 1000
End Sub
