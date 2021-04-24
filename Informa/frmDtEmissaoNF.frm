VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDtEmissaoNF 
   Caption         =   "Dados de Data de Emissao de NF (Exclusivo HEXAL)"
   ClientHeight    =   6330
   ClientLeft      =   2595
   ClientTop       =   945
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   5925
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "HEXAL"
      TabPicture(0)   =   "frmDtEmissaoNF.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "ALLERGAN"
      TabPicture(1)   =   "frmDtEmissaoNF.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "MEDLEY"
      TabPicture(2)   =   "frmDtEmissaoNF.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "FARMASA"
      TabPicture(3)   =   "frmDtEmissaoNF.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).Control(1)=   "Frame7"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame8 
         Caption         =   "Gerar Arquivo FARMASA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   53
         Top             =   2880
         Width           =   5175
         Begin VB.CommandButton cmdGerarFarmasa 
            Caption         =   "Gerar Arquivo FARMASA"
            Enabled         =   0   'False
            Height          =   735
            Left            =   3360
            TabIndex        =   54
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblContArq4 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2160
            TabIndex        =   57
            Top             =   1440
            Width           =   90
         End
         Begin VB.Label Label30 
            Caption         =   "NF Processadas:"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            Caption         =   "Gera Arquivo para cliente no formato solicitado. Diretório C:\INFORMA\FARMASA\FARMASA.TXT"
            Height          =   675
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   3165
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ler Arquivo FARMASA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   44
         Top             =   600
         Width           =   5175
         Begin VB.CommandButton cmdLerFarmasa 
            Caption         =   "Ler Arquivo BomiFarma / Farmasa"
            Enabled         =   0   'False
            Height          =   855
            Left            =   3360
            TabIndex        =   45
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            Caption         =   "Falta Interface da Bomi. O layout programado foi feito a partir de um XLS transformado em DBF e depois em TXT(aqui na INTEC)"
            Height          =   855
            Left            =   2520
            TabIndex        =   58
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            Caption         =   "Arquivo NOTAS.TXT no diretório C:\INFORMA\FARMASA\NOTAS.TXT"
            Height          =   495
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label lblLida4 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2040
            TabIndex        =   51
            Top             =   1200
            Width           =   90
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "NF Lidas"
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label lblNaoOk4 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2040
            TabIndex        =   49
            Top             =   1680
            Width           =   90
         End
         Begin VB.Label lblOk4 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2040
            TabIndex        =   48
            Top             =   1440
            Width           =   90
         End
         Begin VB.Label Label23 
            Caption         =   "NF Processadas:"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "NFs Não Encontradas:"
            Height          =   195
            Left            =   240
            TabIndex        =   46
            Top             =   1680
            Width           =   1620
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Ler Arquivo MEDLEY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   35
         Top             =   600
         Width           =   5175
         Begin VB.CommandButton cmdLerMedley 
            Caption         =   "Ler Arquivo BomiFarma / Medley"
            Enabled         =   0   'False
            Height          =   615
            Left            =   3240
            TabIndex        =   36
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "Arquivo Com Rotina Desatualizada. Necessário Rever."
            Height          =   615
            Left            =   3360
            TabIndex        =   59
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "NFs Não Encontradas:"
            Height          =   195
            Left            =   360
            TabIndex        =   43
            Top             =   1680
            Width           =   1620
         End
         Begin VB.Label Label21 
            Caption         =   "NF Processadas:"
            Height          =   255
            Left            =   360
            TabIndex        =   42
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblOk3 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   41
            Top             =   1440
            Width           =   90
         End
         Begin VB.Label lblNaoOk3 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   40
            Top             =   1680
            Width           =   90
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "NF Lidas"
            Height          =   195
            Left            =   360
            TabIndex        =   39
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label lblLida3 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   38
            Top             =   1200
            Width           =   90
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "Arquivo NOTAS.TXT no diretório C:\INFORMA\MEDLEY"
            Height          =   495
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Gerar Arquivo MEDLEY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   30
         Top             =   2880
         Width           =   5175
         Begin VB.CommandButton cmdGeraMedley 
            Caption         =   "Gerar Arquivo MEDLEY"
            Enabled         =   0   'False
            Height          =   735
            Left            =   3360
            TabIndex        =   31
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label19 
            Caption         =   "Arquivo Com Rotina Desatualizada. Necessário Rever."
            Height          =   615
            Left            =   3480
            TabIndex        =   60
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "Gera Arquivo para cliente no formato solicitado. Diretório C:\INFORMA\MEDLEY\MEDLEY.TXT"
            Height          =   675
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   3165
         End
         Begin VB.Label Label10 
            Caption         =   "NF Processadas:"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblContArq3 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2160
            TabIndex        =   32
            Top             =   1440
            Width           =   90
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ler Arquivo ALLERGAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   21
         Top             =   600
         Width           =   5175
         Begin VB.CommandButton cmdLerAllergan 
            Caption         =   "Ler Arquivo BomiFarma / Allergan"
            Height          =   1095
            Left            =   3360
            TabIndex        =   22
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Arquivo NOTAS.TXT no diretório C:\INFORMA\ALLERGAN"
            Height          =   495
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label lblLida2 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   28
            Top             =   1200
            Width           =   90
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "NF Lidas"
            Height          =   195
            Left            =   360
            TabIndex        =   27
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label lblNaoOK2 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   26
            Top             =   1680
            Width           =   90
         End
         Begin VB.Label lblOk2 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   25
            Top             =   1440
            Width           =   90
         End
         Begin VB.Label Label12 
            Caption         =   "NF Processadas:"
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "NFs Não Encontradas:"
            Height          =   195
            Left            =   360
            TabIndex        =   23
            Top             =   1680
            Width           =   1620
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Gerar Arquivo ALLERGAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         TabIndex        =   16
         Top             =   2880
         Width           =   5175
         Begin VB.CommandButton cmdGerarAllergan 
            Caption         =   "Gerar Arquivo Allergan"
            Height          =   735
            Left            =   3360
            TabIndex        =   17
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblContArq2 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2160
            TabIndex        =   20
            Top             =   1440
            Width           =   90
         End
         Begin VB.Label Label15 
            Caption         =   "NF Processadas:"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Caption         =   "Gera Arquivo para cliente no formato solicitado. Diretório C:\INFORMA\ALLERGAN\ALLERGAN.TXT"
            Height          =   675
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   3165
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ler Arquivo HEXAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   4935
         Begin VB.CommandButton cmdLer 
            Caption         =   "Ler Arquivo BomiFarma / Cliente Hexal"
            Height          =   1095
            Left            =   3120
            TabIndex        =   8
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "NFs Não Encontradas:"
            Height          =   195
            Left            =   360
            TabIndex        =   15
            Top             =   1680
            Width           =   1620
         End
         Begin VB.Label Label2 
            Caption         =   "NF Processadas:"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblOk 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   13
            Top             =   1440
            Width           =   90
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Arquivo NOTAS.TXT no diretório C:\INFORMA\HEXAL"
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label lblNaoOK 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   11
            Top             =   1680
            Width           =   90
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "NF Lidas"
            Height          =   195
            Left            =   360
            TabIndex        =   10
            Top             =   1200
            Width           =   630
         End
         Begin VB.Label lblLida 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   9
            Top             =   1200
            Width           =   90
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Gerar Arquivo HEXAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   240
         TabIndex        =   2
         Top             =   2880
         Width           =   4935
         Begin VB.CommandButton cmdGerarArq 
            Caption         =   "Gerar Arquivo Hexal"
            Height          =   1095
            Left            =   3120
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Gera Arquivo para cliente no formato solicitado. Diretório C:\INFORMA\HEXAL\HEXAL.TXT"
            Height          =   675
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   2685
         End
         Begin VB.Label Label5 
            Caption         =   "NF Processadas:"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblContArq 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2280
            TabIndex        =   4
            Top             =   1440
            Width           =   90
         End
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "SAIR"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   5760
      Width           =   1575
   End
End
Attribute VB_Name = "frmDtEmissaoNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGeraMedley_Click()

    Dim xuf_dest As String, xcidade_dest As String, xdata_interface As Date, xhora_interface As String
    Dim xnumnf As String, xemissao_nf As Date, xDest_nome As String, xvalornf As Variant
    Dim xdatactc As Date, xfilialctc As String, xmodal As String, xprev_entrega As Date
    Dim xlinha As String, xcontador As Long, xcliente As String, xdataentrega As Variant
    Dim xdias As Variant, xocorr As String, xnoprazo As String, xdataocorr As Variant
    Dim xdiasatraso As Variant, xobs_ocorr As String, xtransp_sub As String
    
    If de_informa.rssel_GeraArqMedley.State = 1 Then de_informa.rssel_GeraArqMedley.Close
    de_informa.sel_GeraArqMedley "50929710%"
    If de_informa.rssel_GeraArqMedley.RecordCount > 0 Then
        Open "C:\INFORMA\MEDLEY\MEDLEY.TXT" For Output As #1
        'cria cabeçário do arquivo (campos)
        xlinha = "Num.NF#Emissao NF#Cliente Dest#Cidade#UF#Data Coleta#Modal#Prazo Contr.#Prev.Entrega#Data Entrega#Prazo Entrega#Dias Antec/Atraso#Antec/Atraso#Ocorrência#Observação#SubContratado#"
        Print #1, xlinha
        xcontador = 0
        Do Until de_informa.rssel_GeraArqMedley.EOF
            xcontador = xcontador + 1
            lblContArq3 = xcontador
            xcliente = "50929710%"
            
            'DEMAIS DADOS DA RS
            xnumnf = de_informa.rssel_GeraArqMedley.Fields("numnf")
            If IsNull(de_informa.rssel_GeraArqMedley.Fields("emissao_nf")) Then
                xEmissao = ""
            Else
                xemissao_nf = de_informa.rssel_GeraArqMedley.Fields("emissao_nf")
            End If
            xDest_nome = de_informa.rssel_GeraArqMedley.Fields("dest_nome")
            xcidade_dest = de_informa.rssel_GeraArqMedley.Fields("cidade_dest")
            xuf_dest = de_informa.rssel_GeraArqMedley.Fields("uf_dest")
            xdatactc = de_informa.rssel_GeraArqMedley.Fields("data")
            xfilialctc = de_informa.rssel_GeraArqMedley.Fields("filialctc")
            xmodal = de_informa.rssel_GeraArqMedley.Fields("modal")
            xprev_entr = de_informa.rssel_GeraArqMedley.Fields("prev_entrega")
            xtransp_sub = de_informa.rssel_GeraArqMedley.Fields("transp_sub")
            xobs_ocorr = ""
            
            'busca prazo
            
            xbuscaprazo = buscaprazo2(xuf_dest, xcidade_dest, "TAB010", xmodal)
            xprazo_TT = Val(Mid$(xbuscaprazo, 1, 2))
            
            'verifica horário de corte - HORA
            If Val(Mid$(de_informa.rssel_GeraArqMedley.Fields("hora"), 1, 2)) > Val(Mid$(xbuscaprazo, 4, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            ElseIf Val(Mid$(de_informa.rssel_GeraArqMedley.Fields("hora"), 1, 2)) = Val(Mid$(xbuscaprazo, 4, 2)) And _
                   Val(Mid$(de_informa.rssel_GeraArqMedley.Fields("hora"), 4, 2)) > Val(Mid$(xbuscaprazo, 7, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            Else
                If diautil(de_informa.rssel_GeraArqMedley.Fields("data"), de_informa.rssel_GeraArqMedley.Fields("uf_dest"), _
                   de_informa.rssel_GeraArqMedley.Fields("cidade_dest")) = False And xprazo_TT = 0 Then
                    xprazo_TT = xprazo_TT + 1
                End If
            End If
            
            xprazo = xprazo_TT
            
            'busca entrega
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr xfilialctc, "01"
            If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
                xdataentrega = de_informa.rsSel_ConsOcorr.Fields("data")
                xdias = de_informa.rsSel_ConsOcorr.Fields("diasuteis")
                xdiasatraso = xprazo - xdias
                If Not IsNull(de_informa.rsSel_ConsOcorr.Fields("obs_ocorr")) Then
                    xobs_ocorr = Trim$(de_informa.rsSel_ConsOcorr.Fields("obs_ocorr"))
                Else
                    xobs_ocorr = ""
                End If
                If xdiasatraso = 0 Then
                    xnoprazo = "NO PRAZO"
                ElseIf xdiasatraso > 0 Then
                    xnoprazo = "ANTECIPADO"
                ElseIf xdiasatraso < 0 Then
                    xnoprazo = "COM ATRASO"
                End If
            Else
                xdataentrega = ""
                xdias = ""
                xnoprazo = ""
                xdiasatraso = ""
            End If
            
            'busca ocorrências
            
            xocorr = ""
            If de_informa.rsSel_GeraArqMedleyOcorr.State = 1 Then de_informa.rsSel_GeraArqMedleyOcorr.Close
            de_informa.Sel_GeraArqMedleyOcorr xfilialctc
            If de_informa.rsSel_GeraArqMedleyOcorr.RecordCount > 0 Then
                Do Until de_informa.rsSel_GeraArqMedleyOcorr.EOF
                    xdataocorr = de_informa.rsSel_GeraArqMedleyOcorr.Fields("data")
                    xdataocorr = Trim$(Str(Day(xdataocorr))) & "/" & Trim$(Str(Month(xdataocorr))) & "/" & Trim$(Str(Year(xdataocorr)))
                    If Not IsNull(de_informa.rsSel_GeraArqMedleyOcorr.Fields("obs_ocorr")) Then
                        xobs_ocorr = xobs_ocorr & "; " & Trim$(de_informa.rsSel_GeraArqMedleyOcorr.Fields("obs_ocorr"))
                    End If
                    xocorr = xocorr & xdataocorr & "-" & Trim$(de_informa.rsSel_GeraArqMedleyOcorr.Fields("descr_ocorr")) & " ; "
                    de_informa.rsSel_GeraArqMedleyOcorr.MoveNext
                Loop
            End If
        
            'monta e grava xlinha
                    
            xlinha = xnumnf & "#" & xemissao_nf & "#" & xDest_nome & "#" & xcidade_dest & "#" & xuf_dest & "#" & _
                     xdatactc & "#" & xmodal & "#" & xprazo & "#" & xprev_entr & "#" & xdataentrega & "#" & xdias & "#" & _
                     xdiasatraso & "#" & xnoprazo & "#" & xocorr & "#" & xobs_ocorr & "#" & xtransp_sub & "#"
            
            Print #1, xlinha
        
            'de_informa.alt_EnvAllerganSim Val(xnumnf), xfilialctc
            de_informa.rssel_GeraArqMedley.MoveNext
            DoEvents


        Loop
        Close #1
        MsgBox "Processo Finalizado !"
    Else
        MsgBox "Não há Dados a serem Gerados !"
    End If



End Sub

Private Sub cmdGerarAllergan_Click()
    Dim xuf_dest As String, xcidade_dest As String, xdata_interface As Date, xhora_interface As String
    Dim xnumnf As String, xemissao_nf As Date, xDest_nome As String, xvalornf As Variant
    Dim xdatactc As Date, xfilialctc As String, xmodal As String, xprev_entrega As Date
    Dim xlinha As String, xcontador As Long, xcliente As String, xdataentrega As Variant
    Dim xdias As Variant, xocorr As String, xnoprazo As String, xdataocorr As Variant
    Dim xobs As String
    
    If de_informa.rsSel_GeraArqAllergan.State = 1 Then de_informa.rsSel_GeraArqAllergan.Close
    de_informa.Sel_GeraArqAllergan "43426626000924"
    If de_informa.rsSel_GeraArqAllergan.RecordCount > 0 Then
        Open "C:\INFORMA\ALLERGAN\ALLERGAN.TXT" For Output As #1
        'cria cabeçário do arquivo (campos)
        xlinha = "Num.NF#Data Arquivo#Hora Arquivo#Emissao NF#Cliente Dest#Cidade#UF#Valor#Emissao CTC#Dt Entrega#Hora Entrega#Recebedor#Prev.Entrega#Prazo Contr.#Antecep/Antec#Dias Antec/Atraso#Ocorrência/Observações"
        Print #1, xlinha
        xcontador = 0
        Do Until de_informa.rsSel_GeraArqAllergan.EOF
            xcontador = xcontador + 1
            lblContArq2 = xcontador
            xcliente = "43426626%"
            
            'DEMAIS DADOS DA RS
            xdata_interface = de_informa.rsSel_GeraArqAllergan.Fields("data_interface")
            xhora_interface = de_informa.rsSel_GeraArqAllergan.Fields("hora_interface")
            xnumnf = de_informa.rsSel_GeraArqAllergan.Fields("numnf")
            xemissao_nf = de_informa.rsSel_GeraArqAllergan.Fields("emissao_nf")
            xDest_nome = de_informa.rsSel_GeraArqAllergan.Fields("dest_nome")
            xcidade_dest = de_informa.rsSel_GeraArqAllergan.Fields("cidade_dest")
            xuf_dest = de_informa.rsSel_GeraArqAllergan.Fields("uf_dest")
            xvalornf = de_informa.rsSel_GeraArqAllergan.Fields("valornf") 'valor do CTC
            If Len(de_informa.rsSel_GeraArqAllergan.Fields("nfs")) > 7 Then 'indica que há mais de uma NF no CTC
                xvalornf = ""
            End If
            xdatactc = de_informa.rsSel_GeraArqAllergan.Fields("data")
            xfilialctc = de_informa.rsSel_GeraArqAllergan.Fields("filialctc")
            xmodal = Mid$(de_informa.rsSel_GeraArqAllergan.Fields("modal"), 1, 1)
            xprev_entrega = de_informa.rsSel_GeraArqAllergan.Fields("prev_entrega")
            
            'busca entrega
            
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr xfilialctc, "01"
            If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
                xdataentrega = de_informa.rsSel_ConsOcorr.Fields("data")
                xhoraentrega = de_informa.rsSel_ConsOcorr.Fields("hora")
                If Not IsNull(de_informa.rsSel_ConsOcorr.Fields("receb")) Then
                    If Len(Trim$(de_informa.rsSel_ConsOcorr.Fields("recebpre"))) > Len(Trim$(de_informa.rsSel_ConsOcorr.Fields("receb"))) Then
                        xrecebedor = de_informa.rsSel_ConsOcorr.Fields("recebpre")
                    Else
                        xrecebedor = de_informa.rsSel_ConsOcorr.Fields("receb")
                    End If
                Else
                    xrecebedor = de_informa.rsSel_ConsOcorr.Fields("recebpre")
                End If
            Else
                xdataentrega = ""
                xhoraentrega = ""
                xrecebedor = ""
            End If
            
            'busca prazo
            
            xbuscaprazo = buscaprazo2(xuf_dest, xcidade_dest, "TAB009", xmodal)
            xprazo_TT = Val(Mid$(xbuscaprazo, 1, 2))
            
            'verifica horário de corte - HORA
            If Val(Mid$(de_informa.rsSel_GeraArqAllergan.Fields("hora"), 1, 2)) > Val(Mid$(xbuscaprazo, 4, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            ElseIf Val(Mid$(de_informa.rsSel_GeraArqAllergan.Fields("hora"), 1, 2)) = Val(Mid$(xbuscaprazo, 4, 2)) And _
                   Val(Mid$(de_informa.rsSel_GeraArqAllergan.Fields("hora"), 4, 2)) > Val(Mid$(xbuscaprazo, 7, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            Else
                If diautil(de_informa.rsSel_GeraArqAllergan.Fields("data"), de_informa.rsSel_GeraArqAllergan.Fields("uf_dest"), _
                   de_informa.rsSel_GeraArqAllergan.Fields("cidade_dest")) = False And xprazo_TT = 0 Then
                    xprazo_TT = xprazo_TT + 1
                End If
            End If
            
            xprazo = xprazo_TT
            
            If xdataentrega = "" Then
                xdias = ""
                xnoprazo = ""
            Else
                xdias = xdataentrega - xprev_entrega
                If xdataentrega = xprev_entrega Then
                    xnoprazo = "NO PRAZO"
                ElseIf xdataentrega < xprev_entrega Then
                    xnoprazo = "ANTECIPADO"
                ElseIf xdataentrega > xprev_entrega Then
                    xnoprazo = "ATRASO"
                End If
            End If
            
            'busca ocorrências
            
            xocorr = ""
            If de_informa.rsSel_GeraArqAllerganOcorr.State = 1 Then de_informa.rsSel_GeraArqAllerganOcorr.Close
            de_informa.Sel_GeraArqAllerganOcorr xfilialctc
            If de_informa.rsSel_GeraArqAllerganOcorr.RecordCount > 0 Then
                Do Until de_informa.rsSel_GeraArqAllerganOcorr.EOF
                    xdataocorr = de_informa.rsSel_GeraArqAllerganOcorr.Fields("data")
                    If Not IsNull(de_informa.rsSel_GeraArqAllerganOcorr.Fields("obs_ocorr")) Then
                        xobs = " (" & Trim$(de_informa.rsSel_GeraArqAllerganOcorr.Fields("obs_ocorr")) & ") "
                    Else
                        xobs = ""
                    End If
                    xdataocorr = Trim$(Str(Day(xdataocorr))) & "/" & Trim$(Str(Month(xdataocorr))) & "/" & Trim$(Str(Year(xdataocorr)))
                    xocorr = xocorr & xdataocorr & "-" & Trim$(de_informa.rsSel_GeraArqAllerganOcorr.Fields("descr_ocorr")) & xobs & " ; "
                    de_informa.rsSel_GeraArqAllerganOcorr.MoveNext
                Loop
            End If
        
            'monta e grava xlinha
                    
            xlinha = xnumnf & "#" & xdata_interface & "#" & xhora_interface & "#" & xemissao_nf & "#" & xDest_nome & "#" & xcidade_dest & "#" & xuf_dest & "#" & xvalornf & "#" & xdatactc & "#" & xdataentrega & "#" & xhoraentrega & "#" & xrecebedor & "#" & xprev_entrega & "#" & xprazo & "#" & xnoprazo & "#" & xdias & "#" & xocorr & "#"
            
            Print #1, xlinha
        
            de_informa.alt_EnvAllerganSim Val(xnumnf), xfilialctc
            de_informa.rsSel_GeraArqAllergan.MoveNext
            DoEvents


        Loop
        Close #1
        MsgBox "Processo Finalizado !"
    Else
        MsgBox "Não há Dados a serem Gerados !"
    End If


End Sub

Private Sub cmdGerarArq_Click()
    Dim xnumpedido As String, xdtpedido As Date, xnumnf As String, xemissao_nf As Date, xDest_nome As String
    Dim xcidade_dest As String, xuf_dest As String, xvalornf As Currency, xdatactc As Date
    Dim xdataentrega As Variant, xtempofat As Long, xprazo As Long, xnoprazo As String, xdias As Variant, xcliente As String
    Dim xprev_entrega As Date, xfilialctc As String, xmodal As String, xocorr As String, xdataocorr As Variant
    Dim xcontador As Long
    
    'CGC CLIENTE HEXAL
    xcliente = "61286647000"
    
    If de_informa.rsSel_GerarArqHexal1.State Then de_informa.rsSel_GerarArqHexal1.Close
    de_informa.Sel_GerarArqHexal1 Mid$(xcliente, 1, 8) & "%"
    If de_informa.rsSel_GerarArqHexal1.RecordCount > 0 Then
        Open "C:\INFORMA\HEXAL\HEXAL.TXT" For Output As #1
        'cria cabeçário do arquivo (campos)
        xlinha = "Num.Pedido#Emissao Ped.#Num.NF#Emissao NF#Tempo Fat#Cliente#Cidade#UF#Valor#Emissao CTC#Dt Entrega#Prev.Entrega#Prazo Contr.#Antecep/Antec#Dias Antec/Atraso#Obs/Ocorr#"
        Print #1, xlinha
        xcontador = 0
        Do Until de_informa.rsSel_GerarArqHexal1.EOF
            xcontador = xcontador + 1
            lblContArq = xcontador
            xcliente = "61286647%"
            'DEMAIS DADOS DA RS
            xnumpedido = de_informa.rsSel_GerarArqHexal1.Fields("numpedido")
            xdtpedido = de_informa.rsSel_GerarArqHexal1.Fields("dtpedido")
            xnumnf = de_informa.rsSel_GerarArqHexal1.Fields("numnf")
            xemissao_nf = de_informa.rsSel_GerarArqHexal1.Fields("emissao_nf")
            xtempofat = xemissao_nf - xdtpedido
            xDest_nome = de_informa.rsSel_GerarArqHexal1.Fields("dest_nome")
            xcidade_dest = de_informa.rsSel_GerarArqHexal1.Fields("cidade_dest")
            xuf_dest = de_informa.rsSel_GerarArqHexal1.Fields("uf_dest")
            xvalornf = de_informa.rsSel_GerarArqHexal1.Fields("valornf")
            xdatactc = de_informa.rsSel_GerarArqHexal1.Fields("data")
            xfilialctc = de_informa.rsSel_GerarArqHexal1.Fields("filialctc")
            xmodal = Mid$(de_informa.rsSel_GerarArqHexal1.Fields("modal"), 1, 1)
            xprev_entrega = de_informa.rsSel_GerarArqHexal1.Fields("prev_entrega")
            xobs_ocorr = ""
            
            'busca entrega
            
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr xfilialctc, "01"
            If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
                xdataentrega = de_informa.rsSel_ConsOcorr.Fields("data")
            Else
                xdataentrega = ""
            End If
            
            'busca prazo
            
            xbuscaprazo = buscaprazo2(xuf_dest, xcidade_dest, "TAB005", xmodal)
            xprazo_TT = Val(Mid$(xbuscaprazo, 1, 2))
            
            'verifica horário de corte - HORA
            If Val(Mid$(de_informa.rsSel_GerarArqHexal1.Fields("hora"), 1, 2)) > Val(Mid$(xbuscaprazo, 4, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            ElseIf Val(Mid$(de_informa.rsSel_GerarArqHexal1.Fields("hora"), 1, 2)) = Val(Mid$(xbuscaprazo, 4, 2)) And _
                   Val(Mid$(de_informa.rsSel_GerarArqHexal1.Fields("hora"), 4, 2)) > Val(Mid$(xbuscaprazo, 7, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            Else
                If diautil(de_informa.rsSel_GerarArqHexal1.Fields("data"), de_informa.rsSel_GerarArqHexal1.Fields("uf_dest"), _
                   de_informa.rsSel_GerarArqHexal1.Fields("cidade_dest")) = False And xprazo_TT = 0 Then
                    xprazo_TT = xprazo_TT + 1
                End If
            End If
            
            xprazo = xprazo_TT
            
            If xdataentrega = "" Then
                xdias = ""
                xnoprazo = ""
            Else
                xdias = xdataentrega - xprev_entrega
                If xdataentrega = xprev_entrega Then
                    xnoprazo = "NO PRAZO"
                ElseIf xdataentrega < xprev_entrega Then
                    xnoprazo = "ANTECIPADO"
                ElseIf xdataentrega > xprev_entrega Then
                    xnoprazo = "ATRASO"
                End If
            End If
            
            'busca ocorrências
            
            xocorr = ""
            If de_informa.rsSel_GerarArqHexal2Ocorr.State = 1 Then de_informa.rsSel_GerarArqHexal2Ocorr.Close
            de_informa.Sel_GerarArqHexal2Ocorr xfilialctc
            If de_informa.rsSel_GerarArqHexal2Ocorr.RecordCount > 0 Then
                Do Until de_informa.rsSel_GerarArqHexal2Ocorr.EOF
                    xdataocorr = de_informa.rsSel_GerarArqHexal2Ocorr.Fields("data")
                    xdataocorr = Trim$(Str(Day(xdataocorr))) & "/" & Trim$(Str(Month(xdataocorr))) & "/" & Trim$(Str(Year(xdataocorr)))
                    xocorr = xocorr & xdataocorr & "-" & Trim$(de_informa.rsSel_GerarArqHexal2Ocorr.Fields("descr_ocorr")) & " ; "
                    de_informa.rsSel_GerarArqHexal2Ocorr.MoveNext
                Loop
            End If
        
            'monta e grava xlinha
                    
            xlinha = xnumpedido & "#" & xdtpedido & "#" & xnumnf & "#" & xemissao_nf & "#" & xtempofat & "#" & xDest_nome & "#" & xcidade_dest & "#" & xuf_dest & "#" & xvalornf & "#" & xdatactc & "#" & xdataentrega & "#" & xprev_entrega & "#" & xprazo & "#" & xnoprazo & "#" & xdias & "#" & xocorr & "#"
            
            Print #1, xlinha
        
            de_informa.Alt_EnvHexalSim Val(xnumnf), xfilialctc
            de_informa.rsSel_GerarArqHexal1.MoveNext
            DoEvents
        
        Loop
        Close #1
        MsgBox "Processo Finalizado !"
    Else
        MsgBox "Não há Dados a serem Gerados !"
    End If

    
        
        

End Sub

Private Sub cmdGerarFarmasa_Click()
    Dim xnumpedido As String, xdtpedido As Date, xnumnf As String, xemissao_nf As Date, xDest_nome As String
    Dim xcidade_dest As String, xuf_dest As String, xvalornf As Currency, xdatactc As Date
    Dim xdataentrega As Variant, xtempofat As Long, xprazo As Long, xnoprazo As String, xdias As Variant, xcliente As String
    Dim xprev_entrega As Date, xfilialctc As String, xmodal As String, xocorr As String, xdataocorr As Variant
    Dim xcontador As Long
    
    'CGC CLIENTE FARMASA
    xcliente = "61150819"
    
    If de_informa.rsSel_GeraArqFarmasa.State Then de_informa.rsSel_GeraArqFarmasa.Close
    de_informa.Sel_GeraArqFarmasa Mid$(xcliente, 1, 8) & "%"
    If de_informa.rsSel_GeraArqFarmasa.RecordCount > 0 Then
        Open "C:\INFORMA\FARMASA\FARMASA.TXT" For Output As #1
        'cria cabeçário do arquivo (campos)
        xlinha = "Num.Pedido#Data Marc.#Num.NF#Emissao NF#Tempo Fat.#Cliente#Cidade#UF#Valor#Emissao CTC#Dt Entrega#Prev.Entrega#Prazo Contr.#Antecep/Antec#Dias Antec/Atraso#Obs/Ocorr#"
        Print #1, xlinha
        xcontador = 0
        Do Until de_informa.rsSel_GeraArqFarmasa.EOF
            xcontador = xcontador + 1
            lblContArq4 = xcontador
            xcliente = "61150819%"
            'DEMAIS DADOS DA RS
            xnumpedido = "" 'de_informa.rsSel_GeraArqFarmasa.Fields("numpedido")
            xdtpedido = 0  'de_informa.rsSel_GeraArqFarmasa.Fields("dtpedido")
            xnumnf = de_informa.rsSel_GeraArqFarmasa.Fields("numnf")
            xemissao_nf = "1970/03/11"  'de_informa.rsSel_GeraArqFarmasa.Fields("emissao_nf")
            xtempofat = 0 'xemissao_nf - xdtpedido
            xDest_nome = de_informa.rsSel_GeraArqFarmasa.Fields("dest_nome")
            xcidade_dest = de_informa.rsSel_GeraArqFarmasa.Fields("cidade_dest")
            xuf_dest = de_informa.rsSel_GeraArqFarmasa.Fields("uf_dest")
            xvalornf = 0  'de_informa.rsSel_GeraArqFarmasa.Fields("valornf")
            xdatactc = de_informa.rsSel_GeraArqFarmasa.Fields("data")
            xfilialctc = de_informa.rsSel_GeraArqFarmasa.Fields("filialctc")
            xmodal = Mid$(de_informa.rsSel_GeraArqFarmasa.Fields("modal"), 1, 1)
            xprev_entrega = de_informa.rsSel_GeraArqFarmasa.Fields("prev_entrega")
            xobs_ocorr = ""
            
            'busca entrega
            
            If de_informa.rsSel_ConsOcorr.State = 1 Then de_informa.rsSel_ConsOcorr.Close
            de_informa.Sel_ConsOcorr xfilialctc, "01"
            If de_informa.rsSel_ConsOcorr.RecordCount > 0 Then
                xdataentrega = de_informa.rsSel_ConsOcorr.Fields("data")
            Else
                xdataentrega = ""
            End If
            
            'busca prazo
            
            xbuscaprazo = buscaprazo2(xuf_dest, xcidade_dest, "TAB002", xmodal)
            xprazo_TT = Val(Mid$(xbuscaprazo, 1, 2))
            
            'verifica horário de corte - HORA
            If Val(Mid$(de_informa.rssel_GeraArqMedley.Fields("hora"), 1, 2)) > Val(Mid$(xbuscaprazo, 4, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            ElseIf Val(Mid$(de_informa.rssel_GeraArqMedley.Fields("hora"), 1, 2)) = Val(Mid$(xbuscaprazo, 4, 2)) And _
                   Val(Mid$(de_informa.rssel_GeraArqMedley.Fields("hora"), 4, 2)) > Val(Mid$(xbuscaprazo, 7, 2)) Then
                'emissao posterior ao horário de corte
                xprazo_TT = xprazo_TT + 1
            Else
                If diautil(de_informa.rssel_GeraArqMedley.Fields("data"), de_informa.rssel_GeraArqMedley.Fields("uf_dest"), _
                   de_informa.rssel_GeraArqMedley.Fields("cidade_dest")) = False And xprazo_TT = 0 Then
                    xprazo_TT = xprazo_TT + 1
                End If
            End If
            
            xprazo = xprazo_TT
            
            If xdataentrega = "" Then
                xdias = ""
                xnoprazo = ""
            Else
                xdias = xdataentrega - xprev_entrega
                If xdataentrega = xprev_entrega Then
                    xnoprazo = "NO PRAZO"
                ElseIf xdataentrega < xprev_entrega Then
                    xnoprazo = "ANTECIPADO"
                ElseIf xdataentrega > xprev_entrega Then
                    xnoprazo = "ATRASO"
                End If
            End If
            
            'busca ocorrências
            
            xocorr = ""
            If de_informa.rsSel_GeraFarmasaOcorr.State = 1 Then de_informa.rsSel_GeraFarmasaOcorr.Close
            de_informa.Sel_GeraFarmasaOcorr xfilialctc
            If de_informa.rsSel_GeraFarmasaOcorr.RecordCount > 0 Then
                Do Until de_informa.rsSel_GeraFarmasaOcorr.EOF
                    xdataocorr = de_informa.rsSel_GeraFarmasaOcorr.Fields("data")
                    xdataocorr = Trim$(Str(Day(xdataocorr))) & "/" & Trim$(Str(Month(xdataocorr))) & "/" & Trim$(Str(Year(xdataocorr)))
                    xocorr = xocorr & xdataocorr & "-" & Trim$(de_informa.rsSel_GeraFarmasaOcorr.Fields("descr_ocorr")) & " ; "
                    de_informa.rsSel_GeraFarmasaOcorr.MoveNext
                Loop
            End If
        
            'monta e grava xlinha
                    
            xlinha = xnumpedido & "#" & xdtpedido & "#" & xnumnf & "#" & xemissao_nf & "#" & xtempofat & "#" & xDest_nome & "#" & xcidade_dest & "#" & xuf_dest & "#" & xvalornf & "#" & xdatactc & "#" & xdataentrega & "#" & xprev_entrega & "#" & xprazo & "#" & xnoprazo & "#" & xdias & "#" & xocorr & "#"
            
            Print #1, xlinha
        
            'de_informa.Alt_EnvHexalSim Val(xnumnf), xfilialctc
            de_informa.rsSel_GeraArqFarmasa.MoveNext
            DoEvents
        
        Loop
        Close #1
        MsgBox "Processo Finalizado !"
    Else
        MsgBox "Não há Dados a serem Gerados !"
    End If

    
        
        

End Sub

Private Sub cmdLer_Click()
    Dim xlinha As String, xnf As Long, xdatanf As Date, xnumpedido As String, xdtpedido As Date, xvalornf As Double

     Open "C:\INFORMA\HEXAL\NOTAS.TXT" For Input As #1
     lblOk = 0
     lblNaoOK = 0
     lblLida = 0
     Do Until EOF(1)
        lblLida = Val(lblLida) + 1
        Line Input #1, xlinha
        xnf = Val(Mid$(xlinha, 58, 6))
        xdatanf = CDate(Mid$(xlinha, 235, 2) & "/" & Mid$(xlinha, 233, 2) & "/" & Mid$(xlinha, 229, 4))
        xnumpedido = Trim$(Mid$(xlinha, 5, 6))
        xdtpedido = CDate(Mid$(xlinha, 22, 2) & "/" & Mid$(xlinha, 20, 2) & "/" & Mid$(xlinha, 16, 4))
        xvalornf = CDbl(Val(Mid$(xlinha, 297, 10)))
        'MsgBox Month(xdatanf)
        If de_informa.rsSel_CgcNFEmissao.State = 1 Then de_informa.rsSel_CgcNFEmissao.Close
        de_informa.Sel_CgcNFEmissao "61286647%", xnf
        If de_informa.rsSel_CgcNFEmissao.RecordCount > 0 Then
            de_informa.Alt_DtEmissaoNF xdatanf, xnumpedido, xdtpedido, xvalornf, "61286647%", xnf
            lblOk = Val(lblOk) + 1
        Else
            lblNaoOK = Val(lblNaoOK) + 1
        End If
        DoEvents
     Loop
     MsgBox "Processo Finalizado !"
     Close #1
End Sub

Private Sub lbl_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdLerAllergan_Click()
    Dim xlinha As String, xnf As Long, xdatanf As Date, xdataarquivo As Date, xhoraarquivo As String

     Open "C:\INFORMA\ALLERGAN\NOTAS.TXT" For Input As #1
     lblOk2 = 0
     lblNaoOK2 = 0
     lblLida2 = 0
     Do Until EOF(1)
        Line Input #1, xlinha
        If Val(Mid$(xlinha, 1, 7)) > 0 Then
            lblLida2 = Val(lblLida2) + 1
            xnf = Val(Mid$(xlinha, 24, 6))
            xdataarquivo = CDate("20" & Mid$(xlinha, 17, 2) & "/" & Mid$(xlinha, 14, 2) & "/" & Mid$(xlinha, 11, 2))
            xdatanf = CDate("20" & Mid$(xlinha, 41, 2) & "/" & Mid$(xlinha, 38, 2) & "/" & Mid$(xlinha, 35, 2))
            xhoraarquivo = Mid$(xlinha, 48, 5)
            'MsgBox xnf
            'MsgBox xdatanf
            'MsgBox xdataarquivo
            'MsgBox xhoraarquivo
            If de_informa.rsSel_CgcNFEmissao.State = 1 Then de_informa.rsSel_CgcNFEmissao.Close
            de_informa.Sel_CgcNFEmissao "43426626%", xnf
            If de_informa.rsSel_CgcNFEmissao.RecordCount > 0 Then
                de_informa.alt_dadosallergan xdatanf, xdataarquivo, xhoraarquivo, "43426626%", xnf
                lblOk2 = Val(lblOk2) + 1
            Else
                lblNaoOK2 = Val(lblNaoOK2) + 1
            End If
            DoEvents
        End If
     Loop
     MsgBox "Processo Finalizado !"
     Close #1
End Sub

Private Sub cmdLerFarmasa_Click()
    Dim xlinha As String, xnf As Long, xdatanf As Date, xnumpedido As String, xdtpedido As Date, xvalornf As Double

     Open "C:\INFORMA\FARMASA\NOTAS.TXT" For Input As #1
     lblOk4 = 0
     lblNaoOk4 = 0
     lblLida4 = 0
     Do Until EOF(1)
        lblLida4 = Val(lblLida4) + 1
        Line Input #1, xlinha
        xnf = Val(Mid$(xlinha, 29, 5))
        xdatanf = CDate(Mid$(xlinha, 40, 4) & "/" & Mid$(xlinha, 37, 2) & "/" & Mid$(xlinha, 34, 2))
        xnumpedido = Trim$(Mid$(xlinha, 1, 5))
        xdtpedido = CDate(Mid$(xlinha, 12, 4) & "/" & Mid$(xlinha, 9, 2) & "/" & Mid$(xlinha, 6, 2))
        xvalornf = CDbl(Val(Mid$(xlinha, 131, 15)))
        'MsgBox Month(xdatanf)
        If de_informa.rsSel_CgcNFEmissao.State = 1 Then de_informa.rsSel_CgcNFEmissao.Close
        de_informa.Sel_CgcNFEmissao "61150819%", xnf
        If de_informa.rsSel_CgcNFEmissao.RecordCount > 0 Then
            de_informa.Alt_DtEmissaoNF xdatanf, xnumpedido, xdtpedido, xvalornf, "61150819%", xnf
            lblOk4 = Val(lblOk4) + 1
        Else
            lblNaoOk4 = Val(lblNaoOk4) + 1
        End If
        DoEvents
     Loop
     MsgBox "Processo Finalizado !"
     Close #1

End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDtEmissaoNF = Nothing
End Sub
