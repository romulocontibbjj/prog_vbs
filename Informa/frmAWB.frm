VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAWB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informações de AWB"
   ClientHeight    =   6375
   ClientLeft      =   885
   ClientTop       =   1590
   ClientWidth     =   10695
   Icon            =   "frmAWB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   8760
      TabIndex        =   47
      Top             =   7200
      Width           =   1755
   End
   Begin VB.Frame Frame5 
      Caption         =   "NFs Constantes neste AWB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3720
      Left            =   5640
      TabIndex        =   45
      Top             =   60
      Width           =   4935
      Begin MSFlexGridLib.MSFlexGrid FlexGridNFs 
         Height          =   3375
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   5953
         _Version        =   393216
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Vôo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   120
      TabIndex        =   30
      Top             =   3780
      Width           =   10455
      Begin VB.TextBox TxtDataPartidaCon 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1500
         Width           =   870
      End
      Begin VB.TextBox TxtHoraPartidaCon 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1500
         Width           =   870
      End
      Begin VB.TextBox TxtDataChegadaCon 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1200
         Width           =   870
      End
      Begin VB.TextBox TxtHoraChegadaCon 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1200
         Width           =   870
      End
      Begin VB.TextBox TxtConexao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   900
         Width           =   4065
      End
      Begin VB.TextBox TxtOBS 
         BackColor       =   &H00FFFFFF&
         Height          =   2145
         Left            =   5520
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   240
         Width           =   4785
      End
      Begin VB.TextBox TxtVolumesRetira 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   2100
         Width           =   870
      End
      Begin VB.TextBox TxtHoraChegada 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1800
         Width           =   870
      End
      Begin VB.TextBox TxtDataChegada 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1800
         Width           =   870
      End
      Begin VB.TextBox TxtHoraPartida 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   600
         Width           =   870
      End
      Begin VB.TextBox TxtDataPartida 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   600
         Width           =   870
      End
      Begin VB.TextBox TxtRetirou 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2100
         Width           =   555
      End
      Begin VB.TextBox TxtVoo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   705
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   300
         Width           =   4710
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dt Partida Conexão"
         Height          =   195
         Left            =   330
         TabIndex        =   58
         Top             =   1560
         Width           =   1380
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hora Partida Conexão"
         Height          =   195
         Left            =   2910
         TabIndex        =   57
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dt. Chegada Conexão"
         Height          =   195
         Left            =   135
         TabIndex        =   54
         Top             =   1260
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hora Chegada Conexão"
         Height          =   195
         Left            =   2760
         TabIndex        =   53
         Top             =   1260
         Width           =   1710
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Com Conexão em "
         Height          =   195
         Left            =   60
         TabIndex        =   50
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Volumes Retirados"
         Height          =   195
         Left            =   3150
         TabIndex        =   44
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hora Chegada"
         Height          =   195
         Left            =   3435
         TabIndex        =   42
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Chegada"
         Height          =   195
         Left            =   675
         TabIndex        =   40
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hora Partida"
         Height          =   195
         Left            =   3585
         TabIndex        =   38
         Top             =   660
         Width           =   885
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cliente Retirou Mercadoria?"
         Height          =   195
         Left            =   60
         TabIndex        =   36
         Top             =   2160
         Width           =   1965
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No. Vôo"
         Height          =   195
         Left            =   60
         TabIndex        =   35
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Partida"
         Height          =   195
         Left            =   825
         TabIndex        =   34
         Top             =   660
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Destinatário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   5475
      Begin VB.TextBox TxtUfDES 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4980
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   540
         Width           =   390
      End
      Begin VB.TextBox TxtCidadeDES 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   540
         Width           =   4845
      End
      Begin VB.TextBox TxtNomeDES 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   3645
      End
      Begin VB.TextBox TxtCGCDES 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Informações do AWB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   10
      Top             =   2460
      Width           =   5475
      Begin VB.TextBox TxtFretetotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox TxtHoraEmissao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   1050
      End
      Begin VB.TextBox TxtEmissor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   900
         Width           =   1050
      End
      Begin VB.TextBox TxtDescr 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   300
         Width           =   2325
      End
      Begin VB.TextBox TxtPesoReal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxtVolumes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox TxtDataEmissao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   300
         Width           =   1050
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Data Emissão"
         Height          =   195
         Left            =   3240
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Hora Emissão"
         Height          =   195
         Left            =   3240
         TabIndex        =   23
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emissor"
         Height          =   195
         Left            =   3720
         TabIndex        =   22
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Frete Aéreo Total"
         Height          =   195
         Left            =   1080
         TabIndex        =   21
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Peso Real"
         Height          =   195
         Left            =   1560
         TabIndex        =   20
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Volumes"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   660
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Expedidor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   5475
      Begin VB.TextBox TxtCGCEXP 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1590
      End
      Begin VB.TextBox TxtNomeEXP 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   3645
      End
      Begin VB.TextBox TxtCidadeEXP 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   540
         Width           =   4845
      End
      Begin VB.TextBox TxtUFEXP 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4980
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   390
      End
   End
   Begin VB.Frame fraAWB 
      Caption         =   "AWB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5475
      Begin VB.TextBox TxtDig 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4980
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   390
      End
      Begin VB.TextBox txtAWB 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1845
      End
      Begin VB.TextBox txtCia 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   540
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2565
      End
      Begin VB.TextBox txtFilial 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmAWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
If De_Aereo.rsSelAWB.State = 1 Then De_Aereo.rsSelAWB.Close
If De_Aereo.rsSelAWB_CTC.State = 1 Then De_Aereo.rsSelAWB_CTC.Close
        
        xfilialctc = String(2 - Len(Trim(Str(Val(frmSac.txtFilial.Text)))), "0") & Trim(Str(Val(frmSac.txtFilial.Text))) & String(8 - Len(Trim(Str(Val(frmSac.txtCtc.Text)))), "0") & Trim(Str(Val(frmSac.txtCtc.Text)))
        
        If De_Aereo.rsSelAWB_CTC.State = 1 Then De_Aereo.rsSelAWB_CTC.Close
        De_Aereo.SelAWB_CTC xfilialctc
        
        If De_Aereo.rsSelAWB_CTC.RecordCount > 0 Then
        FlexGridNFs.Rows = 0
        Call limpatela(Me)
        txtFilial.Text = De_Aereo.rsSelAWB_CTC.Fields("filial")
        
                If IsNull(De_Aereo.rsSelAWB_CTC.Fields("cia")) = True Then
                    If De_Aereo.rsSelAWB_CTC.Fields("cia") = "KK" Then
                    txtCia.Text = "Tam"
                    ElseIf De_Aereo.rsSelAWB_CTC.Fields("cia") = "RG" Then
                    txtCia.Text = "Varig"
                    ElseIf De_Aereo.rsSelAWB_CTC.Fields("cia") = "P8" Then
                    txtCia.Text = "Pantanal"
                    ElseIf De_Aereo.rsSelAWB_CTC.Fields("cia") = "VP" Then
                    txtCia.Text = "Vasp"
                    ElseIf De_Aereo.rsSelAWB_CTC.Fields("cia") = "OC" Then
                    txtCia.Text = "Ocean Air"
                    End If
                Else
                txtCia.Text = PriMaiuscula(De_Aereo.rsSelAWB_CTC.Fields("nomecia"))
                End If
                
        If IsNull(De_Aereo.rsSelAWB_CTC.Fields("awb")) = False Then txtAWB.Text = De_Aereo.rsSelAWB_CTC.Fields("awb")
        If IsNull(De_Aereo.rsSelAWB_CTC.Fields("dig")) = False Then TxtDig.Text = De_Aereo.rsSelAWB_CTC.Fields("dig")
        
        If IsNull(De_Aereo.rsSelAWB_CTC.Fields("descrprodsis")) = False Then TxtDescr.Text = PriMaiuscula(De_Aereo.rsSelAWB_CTC.Fields("descrprodsis"))
        
        If IsNull(De_Aereo.rsSelAWB_CTC.Fields("emissor")) = False Then TxtEmissor.Text = De_Aereo.rsSelAWB_CTC.Fields("emissor")
        If IsNull(De_Aereo.rsSelAWB_CTC.Fields("data")) = False Then TxtDataEmissao.Text = De_Aereo.rsSelAWB_CTC.Fields("data")
        If IsNull(De_Aereo.rsSelAWB_CTC.Fields("hora")) = False Then TxtHoraEmissao.Text = De_Aereo.rsSelAWB_CTC.Fields("hora")
        
        If IsNull(De_Aereo.rsSelAWB_CTC.Fields("pesoreal")) = False Then TxtPesoReal.Text = Format(De_Aereo.rsSelAWB_CTC.Fields("pesoreal"), "###0.0")
        If IsNull(De_Aereo.rsSelAWB_CTC.Fields("fretetotal")) = False Then TxtFretetotal.Text = Format(De_Aereo.rsSelAWB_CTC.Fields("fretetotal"), "##,##0.00")
        If IsNull(De_Aereo.rsSelAWB_CTC.Fields("volumes")) = False Then TxtVolumes.Text = De_Aereo.rsSelAWB_CTC.Fields("volumes")
        
        If De_Aereo.rsSelAWB.State = 1 Then De_Aereo.rsSelAWB.Close
        De_Aereo.SelAWB De_Aereo.rsSelAWB_CTC.Fields("codawb")
        
        FlexGridNFs.Clear
        FlexGridNFs.Rows = De_Aereo.rsSelAWB.RecordCount + 1
        FlexGridNFs.Cols = 6
        FlexGridNFs.FixedCols = 0
        FlexGridNFs.FixedRows = 1
        
        FlexGridNFs.TextMatrix(0, 0) = "NF"
        FlexGridNFs.TextMatrix(0, 1) = "Série"
        FlexGridNFs.TextMatrix(0, 2) = "Valor"
        FlexGridNFs.TextMatrix(0, 3) = "FilialCTC"
        FlexGridNFs.TextMatrix(0, 4) = "Remetente"
        FlexGridNFs.TextMatrix(0, 5) = "Destinatário"
        
        FlexGridNFs.ColWidth(0) = 700
        FlexGridNFs.ColWidth(1) = 500
        FlexGridNFs.ColWidth(2) = 1300
        FlexGridNFs.ColWidth(3) = 1200
        FlexGridNFs.ColWidth(4) = 3500
        FlexGridNFs.ColWidth(5) = 3500
        
        xCodAwb = De_Aereo.rsSelAWB.Fields("codawb")
        
        
        X = 0
        
            Do Until De_Aereo.rsSelAWB.EOF
            X = X + 1
            
            If Not IsNull(De_Aereo.rsSelAWB.Fields("nota")) Then FlexGridNFs.TextMatrix(X, 0) = De_Aereo.rsSelAWB.Fields("nota")
            If Not IsNull(De_Aereo.rsSelAWB.Fields("SERIE")) Then FlexGridNFs.TextMatrix(X, 1) = De_Aereo.rsSelAWB.Fields("serie")
            If Not IsNull(De_Aereo.rsSelAWB.Fields("VALOR")) Then FlexGridNFs.TextMatrix(X, 2) = Format(De_Aereo.rsSelAWB.Fields("VALOR"), "##,##0.00")
            If Not IsNull(De_Aereo.rsSelAWB.Fields("FILIALCTC")) Then FlexGridNFs.TextMatrix(X, 3) = De_Aereo.rsSelAWB.Fields("FILIALCTC")
            If Not IsNull(De_Aereo.rsSelAWB.Fields("REMET_NOME")) Then FlexGridNFs.TextMatrix(X, 4) = PriMaiuscula(De_Aereo.rsSelAWB.Fields("REMET_NOME"))
            If Not IsNull(De_Aereo.rsSelAWB.Fields("DEST_NOME")) Then FlexGridNFs.TextMatrix(X, 5) = PriMaiuscula(De_Aereo.rsSelAWB.Fields("DEST_NOME"))
            
            De_Aereo.rsSelAWB.MoveNext
            Loop
            
            
        
        
        If Not IsNull(De_Aereo.rsSelAWB_CTC.Fields("nomeexp")) Then TxtNomeEXP.Text = PriMaiuscula(De_Aereo.rsSelAWB_CTC.Fields("nomeexp"))
        If Not IsNull(De_Aereo.rsSelAWB_CTC.Fields("cidadexp")) Then TxtCidadeEXP.Text = PriMaiuscula(De_Aereo.rsSelAWB_CTC.Fields("cidadexp"))
        If Not IsNull(De_Aereo.rsSelAWB_CTC.Fields("ufexp")) Then TxtUFEXP.Text = De_Aereo.rsSelAWB_CTC.Fields("ufexp")
        If Not IsNull(De_Aereo.rsSelAWB_CTC.Fields("cnpjexp")) Then TxtCGCEXP.Text = De_Aereo.rsSelAWB_CTC.Fields("cnpjexp")
        
        If Not IsNull(De_Aereo.rsSelAWB_CTC.Fields("nomedes")) Then TxtNomeDES.Text = PriMaiuscula(De_Aereo.rsSelAWB_CTC.Fields("nomedes"))
        If Not IsNull(De_Aereo.rsSelAWB_CTC.Fields("cidadedes")) Then TxtCidadeDES.Text = PriMaiuscula(De_Aereo.rsSelAWB_CTC.Fields("cidadedes"))
        If Not IsNull(De_Aereo.rsSelAWB_CTC.Fields("ufdes")) Then TxtUfDES.Text = De_Aereo.rsSelAWB_CTC.Fields("ufdes")
        If Not IsNull(De_Aereo.rsSelAWB_CTC.Fields("cnpjdes")) Then txtCGCDes.Text = De_Aereo.rsSelAWB_CTC.Fields("cnpjdes")
        
        De_Aereo.rsSelAWB.MoveFirst
        
        If De_Aereo.rsSelAwbVoo.State = 1 Then De_Aereo.rsSelAwbVoo.Close
        De_Aereo.SelAwbVoo De_Aereo.rsSelAWB.Fields("codawb")
        
        If De_Aereo.rsSelAwbVoo.RecordCount > 0 Then
        If IsNull(De_Aereo.rsSelAwbVoo.Fields("voo")) = False Then TxtVoo.Text = De_Aereo.rsSelAwbVoo.Fields("voo")
        If IsNull(De_Aereo.rsSelAwbVoo.Fields("data_partida")) = False Then TxtDataPartida = De_Aereo.rsSelAwbVoo.Fields("data_partida")
        If IsNull(De_Aereo.rsSelAwbVoo.Fields("hora_partida")) = False Then TxtHoraPartida.Text = De_Aereo.rsSelAwbVoo.Fields("hora_partida")
            If IsNull(De_Aereo.rsSelAwbVoo.Fields("CONCIDADE")) = False And IsNull(De_Aereo.rsSelAwbVoo.Fields("CONUF")) = False And IsNull(De_Aereo.rsSelAwbVoo.Fields("CONAEROPORTO")) = False Then
                If Len(Trim(De_Aereo.rsSelAwbVoo.Fields("CONAEROPORTO"))) > 0 And Len(Trim(De_Aereo.rsSelAwbVoo.Fields("CONCIDADE"))) And Len(Trim(De_Aereo.rsSelAwbVoo.Fields("CONUF"))) > 0 Then
                TxtConexao.Text = PriMaiuscula(Trim(De_Aereo.rsSelAwbVoo.Fields("CONCIDADE"))) & " - " & PriMaiuscula(Trim(De_Aereo.rsSelAwbVoo.Fields("CONUF"))) & " (" & PriMaiuscula(Trim(De_Aereo.rsSelAwbVoo.Fields("CONAEROPORTO"))) & ")"
                ElseIf Len(Trim(De_Aereo.rsSelAwbVoo.Fields("CONCIDADE"))) And Len(Trim(De_Aereo.rsSelAwbVoo.Fields("CONUF"))) > 0 Then
                TxtConexao.Text = PriMaiuscula(Trim(De_Aereo.rsSelAwbVoo.Fields("CONCIDADE"))) & " - " & PriMaiuscula(Trim(De_Aereo.rsSelAwbVoo.Fields("CONUF")))
                End If
            End If
        If IsNull(De_Aereo.rsSelAwbVoo.Fields("condtcheg")) = False Then TxtDataChegadaCon.Text = De_Aereo.rsSelAwbVoo.Fields("condtcheg")
        If IsNull(De_Aereo.rsSelAwbVoo.Fields("conhoracheg")) = False Then TxtHoraChegadaCon.Text = De_Aereo.rsSelAwbVoo.Fields("conhoracheg")
        If IsNull(De_Aereo.rsSelAwbVoo.Fields("condtpart")) = False Then TxtDataPartidaCon.Text = De_Aereo.rsSelAwbVoo.Fields("condtpart")
        If IsNull(De_Aereo.rsSelAwbVoo.Fields("conhoracheg")) = False Then TxtHoraPartidaCon.Text = De_Aereo.rsSelAwbVoo.Fields("conhoracheg")
        If IsNull(De_Aereo.rsSelAwbVoo.Fields("data_chegada")) = False Then TxtDataChegada.Text = De_Aereo.rsSelAwbVoo.Fields("data_chegada")
        If IsNull(De_Aereo.rsSelAwbVoo.Fields("hora_chegada")) = False Then TxtHoraChegada.Text = De_Aereo.rsSelAwbVoo.Fields("hora_chegada")
        If IsNull(De_Aereo.rsSelAwbVoo.Fields("volumesretirados")) = False Then TxtVolumesRetira.Text = De_Aereo.rsSelAwbVoo.Fields("volumesretirados")
        If IsNull(De_Aereo.rsSelAwbVoo.Fields("obs")) = False Then TxtOBS.Text = De_Aereo.rsSelAwbVoo.Fields("obs")
            If De_Aereo.rsSelAwbVoo.Fields("clienteretirou") = "S" Then
            TxtRetirou.Text = "Sim"
            Else
            TxtRetirou.Text = "Não"
            End If
        End If
    End If
    
End Sub

Private Sub Text22_Change()

End Sub


Public Function PriMaiuscula(Texto) As String
Texto = LCase(Texto)
xmaiuscula = "SIM"
xtexto2 = ""


        For X = 1 To Len(Trim(Texto)) Step 1
           If xmaiuscula = "SIM" Then
            xtexto2 = xtexto2 & UCase(Mid(Trim(Texto), X, 1))
            Else
            xtexto2 = xtexto2 & Mid(Trim(Texto), X, 1)
            End If

            If Mid(Trim(Texto), X, 1) = " " Or Mid(Trim(Texto), X, 1) = "." Or Mid(Trim(Texto), X, 1) = "/" Or Mid(Trim(Texto), X, 1) = "\" Or Mid(Trim(Texto), X, 1) = ";" Or Mid(Trim(Texto), X, 1) = ":" Or Mid(Trim(Texto), X, 1) = "_" Or Mid(Trim(Texto), X, 1) = "&" Or Mid(Trim(Texto), X, 1) = "-" Then
                If Mid(Trim(Texto), X, 1) = " " Then
                    If Mid(Trim(Texto), X, 4) = " do " Then
                    xmaiuscula = "NAO"
                    ElseIf Mid(Trim(Texto), X, 4) = " da " Then
                    xmaiuscula = "NAO"
                    ElseIf Mid(Trim(Texto), X, 4) = " de " Then
                    xmaiuscula = "NAO"
                    ElseIf Mid(Trim(Texto), X, 5) = " das " Then
                    xmaiuscula = "NAO"
                    ElseIf Mid(Trim(Texto), X, 5) = " dos " Then
                    xmaiuscula = "NAO"
                    Else
                    xmaiuscula = "SIM"
                    End If
                Else
                xmaiuscula = "SIM"
                End If
            Else
            xmaiuscula = "NAO"
            End If
        Next
        
PriMaiuscula = xtexto2
End Function

