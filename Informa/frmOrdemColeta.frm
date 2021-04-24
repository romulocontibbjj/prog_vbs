VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmOrdemColeta 
   Caption         =   "Ordem de Coleta"
   ClientHeight    =   6915
   ClientLeft      =   495
   ClientTop       =   1050
   ClientWidth     =   11175
   ControlBox      =   0   'False
   Icon            =   "frmOrdemColeta.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   11175
   WindowState     =   2  'Maximized
   Begin VB.Frame FraAux 
      Height          =   4515
      Index           =   0
      Left            =   1140
      TabIndex        =   58
      Top             =   960
      Visible         =   0   'False
      Width           =   9795
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexAUX 
         Height          =   4215
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   180
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   7435
         _Version        =   393216
         BackColor       =   12648447
         ForeColor       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Solicitante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   180
      TabIndex        =   36
      Top             =   60
      Width           =   11595
      Begin VB.TextBox TxtEnderecoSol 
         Height          =   285
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   1080
         Width           =   4515
      End
      Begin VB.TextBox TxtUFSol 
         Height          =   285
         Index           =   0
         Left            =   5460
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   1380
         Width           =   375
      End
      Begin VB.TextBox TxtCidadeSol 
         Height          =   285
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   1380
         Width           =   4095
      End
      Begin VB.TextBox TxtFilial 
         Height          =   285
         Left            =   540
         MaxLength       =   2
         TabIndex        =   0
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtCGCSol 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1395
      End
      Begin VB.TextBox TxtTelSol 
         Height          =   285
         Index           =   0
         Left            =   8160
         TabIndex        =   5
         Top             =   540
         Width           =   3315
      End
      Begin VB.TextBox TxtNomeSol 
         Height          =   285
         Index           =   0
         Left            =   8160
         TabIndex        =   4
         Top             =   240
         Width           =   3315
      End
      Begin VB.TextBox TxtNomeEmpresaSol 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   540
         Width           =   4935
      End
      Begin VB.TextBox TxtBuscaSol 
         Height          =   285
         Index           =   0
         Left            =   2940
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   76
         Top             =   1125
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Cidade/UF:"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   75
         Top             =   1425
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   285
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   7320
         TabIndex        =   40
         Top             =   585
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nome Solicitante:"
         Height          =   195
         Left            =   6840
         TabIndex        =   39
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ, CPF, Apelido ou ?"
         Height          =   195
         Left            =   4740
         TabIndex        =   38
         Top             =   285
         Width           =   1770
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Solicitante:"
         Height          =   195
         Left            =   1500
         TabIndex        =   37
         Top             =   285
         Width           =   1305
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Local de Coleta (Origem)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   180
      TabIndex        =   41
      Top             =   1020
      Width           =   11595
      Begin VB.TextBox TxtNomeSol 
         Height          =   285
         Index           =   1
         Left            =   8220
         TabIndex        =   80
         Top             =   1800
         Width           =   3315
      End
      Begin VB.TextBox TxtTelSol 
         Height          =   285
         Index           =   1
         Left            =   8220
         TabIndex        =   79
         Top             =   2100
         Width           =   3315
      End
      Begin VB.TextBox TxtBuscaSol 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   6
         Top             =   300
         Width           =   1695
      End
      Begin VB.TextBox TxtNomeEmpresaSol 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   4395
      End
      Begin VB.TextBox TxtCGCSol 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox TxtCidadeSol 
         Height          =   285
         Index           =   1
         Left            =   6960
         TabIndex        =   10
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox TxtUFSol 
         Height          =   285
         Index           =   1
         Left            =   11100
         TabIndex        =   11
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox TxtEnderecoSol 
         Height          =   285
         Index           =   1
         Left            =   6960
         TabIndex        =   9
         Top             =   300
         Width           =   4515
      End
      Begin VB.TextBox TxtEnderecoColeta 
         Height          =   285
         Left            =   900
         TabIndex        =   12
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox TxtCidadeColeta 
         Height          =   285
         Left            =   6960
         TabIndex        =   13
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox TxtUFColeta 
         Height          =   285
         Left            =   11100
         TabIndex        =   14
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Index           =   1
         Left            =   6150
         TabIndex        =   78
         Top             =   345
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Cidade/UF:"
         Height          =   195
         Index           =   1
         Left            =   6060
         TabIndex        =   77
         Top             =   645
         Width           =   825
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Endereço de Coleta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   960
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00400000&
         BorderWidth     =   2
         X1              =   11460
         X2              =   1860
         Y1              =   1050
         Y2              =   1065
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   1245
         Width           =   735
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Cidade/UF:"
         Height          =   195
         Left            =   6060
         TabIndex        =   64
         Top             =   1245
         Width           =   825
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ, CPF, Apelido ou ?"
         Height          =   195
         Left            =   3360
         TabIndex        =   45
         Top             =   300
         Width           =   1770
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Remetente:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Local de Entrega (Destino)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   180
      TabIndex        =   43
      Top             =   2700
      Width           =   11595
      Begin VB.TextBox TxtNomeSol 
         Height          =   285
         Index           =   2
         Left            =   8160
         TabIndex        =   82
         Top             =   1200
         Width           =   3315
      End
      Begin VB.TextBox TxtTelSol 
         Height          =   285
         Index           =   2
         Left            =   8160
         TabIndex        =   81
         Top             =   1500
         Width           =   3315
      End
      Begin VB.TextBox TxtBuscaSol 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   15
         Top             =   300
         Width           =   1695
      End
      Begin VB.TextBox TxtNomeEmpresaSol 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   17
         Top             =   600
         Width           =   4395
      End
      Begin VB.TextBox TxtCGCSol 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox TxtCidadeSol 
         Height          =   285
         Index           =   2
         Left            =   6960
         TabIndex        =   19
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox TxtUFSol 
         Height          =   285
         Index           =   2
         Left            =   11100
         TabIndex        =   20
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox TxtEnderecoSol 
         Height          =   285
         Index           =   2
         Left            =   6960
         TabIndex        =   18
         Top             =   300
         Width           =   4515
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Destinatário:"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ, CPF, Apelido ou ?"
         Height          =   195
         Left            =   3360
         TabIndex        =   62
         Top             =   300
         Width           =   1770
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   6150
         TabIndex        =   61
         Top             =   345
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cidade/UF:"
         Height          =   195
         Left            =   6060
         TabIndex        =   60
         Top             =   645
         Width           =   825
      End
   End
   Begin VB.CommandButton CmdCadCli 
      Caption         =   "Cadastro de Clientes"
      Height          =   435
      Left            =   9720
      TabIndex        =   35
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Frame Frame5 
      Caption         =   "Dados da Carga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   180
      TabIndex        =   44
      Top             =   3780
      Width           =   9435
      Begin VB.TextBox TxtPrioridade 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   28
         Top             =   2160
         Width           =   375
      End
      Begin MSMask.MaskEdBox TxtDataColeta 
         Height          =   285
         Left            =   5880
         TabIndex        =   30
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.ComboBox ComboEspecie 
         Height          =   315
         Left            =   1560
         TabIndex        =   23
         Top             =   900
         Width           =   2295
      End
      Begin VB.CommandButton CmdEspecie 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         TabIndex        =   24
         Top             =   900
         Width           =   315
      End
      Begin VB.TextBox TxtVolumes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   27
         Top             =   1860
         Width           =   2655
      End
      Begin VB.TextBox TxtHorario 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   7980
         TabIndex        =   31
         Top             =   300
         Width           =   1275
      End
      Begin VB.TextBox TxtPeso 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   26
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox TxtOBS 
         Height          =   1575
         Left            =   4320
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   870
         Width           =   4935
      End
      Begin VB.TextBox TxtNFs 
         Height          =   285
         Left            =   1560
         MaxLength       =   255
         TabIndex        =   29
         Top             =   2460
         Width           =   7695
      End
      Begin VB.TextBox TxtNatureza 
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox TxtValMerc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Top             =   1260
         Width           =   2655
      End
      Begin VB.TextBox TxtTipoFrete 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LblPrioridade 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "3 - Urg."
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   70
         Top             =   2205
         Width           =   525
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Prioridade:"
         Height          =   195
         Left            =   735
         TabIndex        =   69
         Top             =   2205
         Width           =   750
      End
      Begin VB.Label LblPrioridade 
         AutoSize        =   -1  'True
         Caption         =   "1 - Norm."
         Height          =   195
         Index           =   0
         Left            =   1980
         TabIndex        =   68
         Top             =   2205
         Width           =   645
      End
      Begin VB.Label LblPrioridade 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "2 - Prior."
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   67
         Top             =   2205
         Width           =   585
      End
      Begin VB.Label LblTipoFrete 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "2 - A Pagar"
         Height          =   195
         Index           =   1
         Left            =   3420
         TabIndex        =   57
         Top             =   285
         Width           =   795
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Volumes:"
         Height          =   195
         Left            =   825
         TabIndex        =   56
         Top             =   1905
         Width           =   645
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Horário:"
         Height          =   195
         Left            =   7335
         TabIndex        =   55
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Espécie:"
         Height          =   195
         Left            =   855
         TabIndex        =   54
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Peso:"
         Height          =   195
         Left            =   1065
         TabIndex        =   53
         Top             =   1605
         Width           =   405
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Obs. de Coleta:"
         Height          =   195
         Left            =   4380
         TabIndex        =   52
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "NFs:"
         Height          =   195
         Left            =   1140
         TabIndex        =   51
         Top             =   2505
         Width           =   330
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Data de Coleta:"
         Height          =   195
         Left            =   4680
         TabIndex        =   50
         Top             =   345
         Width           =   1110
      End
      Begin VB.Label LblTipoFrete 
         AutoSize        =   -1  'True
         Caption         =   "1 - Pago"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   49
         Top             =   285
         Width           =   600
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Natureza:"
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   48
         Top             =   645
         Width           =   690
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vlr.Merc:"
         Height          =   195
         Left            =   840
         TabIndex        =   47
         Top             =   1305
         Width           =   630
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo Frete:"
         Height          =   195
         Left            =   705
         TabIndex        =   46
         Top             =   285
         Width           =   765
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Sair"
      Height          =   615
      Left            =   9720
      TabIndex        =   34
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton CmdGravar 
      Caption         =   "Emitir"
      Height          =   615
      Left            =   9720
      TabIndex        =   33
      Top             =   5280
      Width           =   2055
   End
End
Attribute VB_Name = "frmOrdemColeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Y As Integer
Public xIndexBusca As Integer


Public Sub TextMoneyBox_KeyPress(KeyAsciiRequired As Integer, TextBoxRequired As TextBox)
    If KeyAsciiRequired < 48 Or KeyAsciiRequired > 57 Then
        If KeyAsciiRequired = 13 Then
        SendKeys "{TAB}"
        KeyAsciiRequired = 0
        Else
            If KeyAsciiRequired = 8 Then
                If Len(TextBoxRequired.Text) > 0 Then
                TextBoxRequired.Text = SoNumeros(TextBoxRequired.Text)
                TextBoxRequired.Text = Mid(TextBoxRequired.Text, 1, Len(TextBoxRequired.Text) - 1)
                TextBoxRequired.Text = Val(TextBoxRequired.Text) / 100
                End If
            KeyAsciiRequired = 0
            Else
            KeyAsciiRequired = 0
            End If
        End If
    Else
    TextBoxRequired.Text = SoNumeros(TextBoxRequired.Text) & Chr(KeyAsciiRequired)
    TextBoxRequired.Text = Val(TextBoxRequired.Text) / 100
    KeyAsciiRequired = 0
    End If

TextBoxRequired.Text = Format(TextBoxRequired.Text, "###,###,###,##0.00")
End Sub

Public Sub TextMoneyBox_GotFocus(TxtBoxRequired As TextBox)
If Trim(TxtBoxRequired.Text) = "" Then TxtBoxRequired.Text = "0,00"
TxtBoxRequired.SelStart = 0
TxtBoxRequired.SelLength = Len(TxtBoxRequired.Text) + 1
End Sub

Public Sub TextMoneyBox_Change(TxtBoxRequired As TextBox)
    If Len(Trim(TxtBoxRequired.Text)) = 0 Then
    TxtBoxRequired.Text = "0.00"
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    ElseIf CDbl(TxtBoxRequired.Text) = 0 Then
    TxtBoxRequired.Text = "0.00"
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    Else
    TxtBoxRequired.Text = Format((CDbl(TxtBoxRequired.Text)), "###,##0.00")
    TxtBoxRequired.SelStart = Len(TxtBoxRequired.Text)
    End If
End Sub



Public Sub OrdenaListBox(xListBox As ListBox)
xListBox.Visible = False
DoEvents
    For I = 0 To xListBox.ListCount - 1
        For J = 0 To xListBox.ListCount - 2
            If xListBox.List(I) < xListBox.List(J) Then
            xAux = xListBox.List(J)
            xListBox.List(J) = xListBox.List(I)
            xListBox.List(I) = xAux
            End If
        Next
    Next
xListBox.Visible = True
DoEvents
End Sub

Public Sub TransfereItemDeListBox(List_Origem As ListBox, List_Destino As ListBox)
List_Origem.Visible = False
List_Destino.Visible = False
DoEvents
    For X = 0 To List_Origem.ListCount - 1
        If List_Origem.Selected(X) = True Then
        List_Destino.AddItem List_Origem.List(X)
        End If
    Next
    
X = 0
    Do While True
        
        If X > List_Origem.ListCount - 1 Then
        Exit Do
        End If
        
        If List_Origem.Selected(X) = True Then
        List_Origem.RemoveItem (X)
        X = X - 1
        End If
    X = X + 1
    Loop
List_Origem.Visible = True
List_Destino.Visible = True
DoEvents
End Sub

Public Sub LimpaFrame(xtela As Form, xFrameCaption As String)
    Dim xmask As String
    Dim xcontrol As Control
    For Each xcontrol In xtela
        If TypeOf xcontrol Is TextBox Then
            If xcontrol.Container = xFrameCaption Then
            xcontrol.Text = ""
            End If
        ElseIf TypeOf xcontrol Is Label Then
            If xcontrol.Container = xFrameCaption Then
                If xcontrol.BorderStyle = 1 Then
                    xcontrol.Caption = ""
                End If
            End If
        ElseIf TypeOf xcontrol Is MaskEdBox Then
            If xcontrol.Container = xFrameCaption Then
            xcontrol.Mask = ""
            xcontrol.Text = ""
            End If
        End If
    Next
End Sub

Public Sub TravaFrame(xtela As Form, xFrame As frame, Tipo As Integer)
    Dim xcontrol As Control
    For Each xcontrol In xtela
        If TypeOf xcontrol Is frame And xcontrol <> xFrame And Tipo = 0 Then
        xcontrol.Enabled = False
        ElseIf TypeOf xcontrol Is frame And xcontrol <> xFrame And Tipo = 1 Then
        xcontrol.Enabled = True
        End If
    Next
End Sub
Public Function NomeMes(NumeroMes As Integer) As String
NomeMes = ""
If NumeroMes = 1 Then NomeMes = "Janeiro"
If NumeroMes = 2 Then NomeMes = "Fevereiro"
If NumeroMes = 3 Then NomeMes = "Março"
If NumeroMes = 4 Then NomeMes = "Abril"
If NumeroMes = 5 Then NomeMes = "Maio"
If NumeroMes = 6 Then NomeMes = "Junho"
If NumeroMes = 7 Then NomeMes = "Julho"
If NumeroMes = 8 Then NomeMes = "Agosto"
If NumeroMes = 9 Then NomeMes = "Setembro"
If NumeroMes = 10 Then NomeMes = "Outubro"
If NumeroMes = 11 Then NomeMes = "Novembro"
If NumeroMes = 12 Then NomeMes = "Dezembro"
End Function

Public Function NumMes(NomedoMes As String) As Integer
NomedoMes = LCase(Trim(NomedoMes))
NumMes = 0
If NomedoMes = "janeiro" Then NumMes = 1
If NomedoMes = "fevereiro" Then NumMes = 2
If NomedoMes = "março" Then NumMes = 3
If NomedoMes = "abril" Then NumMes = 4
If NomedoMes = "maio" Then NumMes = 5
If NomedoMes = "junho" Then NumMes = 6
If NomedoMes = "julho" Then NumMes = 7
If NomedoMes = "agosto" Then NumMes = 8
If NomedoMes = "setembro" Then NumMes = 9
If NomedoMes = "outubro" Then NumMes = 10
If NomedoMes = "novembro" Then NumMes = 11
If NomedoMes = "dezembro" Then NumMes = 12
End Function

Public Function SemPonto(Texto As String) As Long
Dim Cont As Integer
Dim TextoAux As String
TextoAux = ""
    For Cont = 1 To Len(Texto)
        If Mid(Texto, Cont, 1) <> "." And Mid(Texto, Cont, 1) <> "," And Mid(Texto, Cont, 1) <> "%" Then
        TextoAux = TextoAux & Mid(Texto, Cont, 1)
        End If
    Next
    
    If Len(TextoAux) <= 9 Then
    SemPonto = Val(TextoAux)
    Else
    SemPonto = 0
    End If
End Function

Private Sub CmdCadCli_Click()
frmCadClientes.Show 1
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdGravar_Click()
If Len(Trim(TxtFilial.Text)) = 0 Then
MsgBox "Você não informou a Filial da Coleta.", vbCritical
Exit Sub
ElseIf Len(Trim(TxtEnderecoColeta.Text)) = 0 Then
MsgBox "Você não informou o Endereço de Coleta.", vbCritical
Exit Sub
ElseIf Len(Trim(TxtCidadeColeta.Text)) = 0 Then
MsgBox "Você não informou a Cidade de Coleta.", vbCritical
Exit Sub
ElseIf Len(Trim(TxtUFColeta.Text)) = 0 Then
MsgBox "Você não informou a UF de Coleta.", vbCritical
Exit Sub
ElseIf Len(Trim(TxtTipoFrete.Text)) = 0 Then
MsgBox "Você não informou o Tipo do Frete.", vbCritical
Exit Sub
ElseIf Len(Trim(ComboEspecie.Text)) = 0 Then
MsgBox "Você não informou a Espécie.", vbCritical
Exit Sub
ElseIf Len(Trim(TxtPeso.Text)) = 0 Then
MsgBox "Você não informou o Peso.", vbCritical
Exit Sub
ElseIf Len(Trim(TxtVolumes.Text)) = 0 Then
MsgBox "Você não informou os Volumes.", vbCritical
Exit Sub
ElseIf Len(Trim(TxtPrioridade.Text)) = 0 Then
MsgBox "Você não informou a Prioridade desta Coleta.", vbCritical
Exit Sub
ElseIf Len(Trim(TxtDataColeta.Text)) = 0 Then
MsgBox "Você não informou a Data de Coleta.", vbCritical
Exit Sub
End If

Dim xFilialColeta As String
Dim xfilial As String
Dim XColeta As String
Dim xprioridade As String
Dim xSolCGC As String
Dim xSolEmpNome As String
Dim xSolNome As String
Dim xSolTel As String
Dim xRemetCgc As String
Dim xRemetNome As String
Dim xRemetEnd As String
Dim xRemetCidade As String
Dim xRemetUF As String
Dim xColetaEnd As String
Dim xColetaCidade As String
Dim xColetaUF As String
Dim xDestCGC As String
Dim xDestNome As String
Dim xDestEnd As String
Dim xDestCidade As String
Dim xDestUF As String
Dim xDataColeta As String
Dim xHoraColeta As String
Dim xTipoFrete As String
Dim xnatureza As String
Dim xEspecie As String
Dim xvalmerc As Currency
Dim xpeso As Currency
Dim xvolumes As Currency
Dim xNfs As String
Dim xobs As String
Dim xEmissor As String
Dim xTem_Ocorr As String

Dim xEmissaoColeta As String
Dim xCod_Ocorr As String
Dim xDescricao As String
Dim xObsOcorr As String
Dim xdata As String
Dim xhora As String




de_informa.cn_informa.BeginTrans

If de_informa.rsSel_ProxColeta.State = 1 Then de_informa.rsSel_ProxColeta.Close
de_informa.Sel_ProxColeta

    If IsNull(de_informa.rsSel_ProxColeta.Fields("coleta")) = True Then
    XColeta = "300000"
    Else
    XColeta = Trim(Str(CDbl(de_informa.rsSel_ProxColeta.Fields("coleta")) + 1))
    End If

xfilial = String(2 - Len(Trim(TxtFilial.Text)), "0") & Trim(TxtFilial.Text)
xFilialColeta = transctc(xfilial, XColeta)

    If Val(TxtPrioridade.Text) = 1 Then
    xprioridade = "NORMAL"
    ElseIf Val(TxtPrioridade.Text) = 2 Then
    xprioridade = "PRIORIDADE"
    ElseIf Val(TxtPrioridade.Text) = 3 Then
    xprioridade = "URGENTE"
    End If

xSolCGC = Trim(TxtCGCSol(0).Text)
xSolEmpNome = UCase(Trim(TxtNomeEmpresaSol(0).Text))
xSolNome = UCase(Trim(TxtNomeSol(0).Text))
xSolTel = UCase(Trim(TxtTelSol(0).Text))
xRemetCgc = Trim(TxtCGCSol(1).Text)
xRemetNome = UCase(Trim(TxtNomeEmpresaSol(1).Text))
xRemetEnd = UCase(Trim(TxtEnderecoSol(1).Text))
xRemetCidade = UCase(Trim(TxtCidadeSol(1).Text))
xRemetUF = UCase(Trim(TxtUFSol(1).Text))
xColetaEnd = UCase(Trim(TxtEnderecoColeta.Text))
xColetaCidade = UCase(Trim(TxtCidadeColeta.Text))
xColetaUF = UCase(Trim(TxtUFColeta.Text))
xDestCGC = Trim(UCase(TxtCGCSol(2).Text))
xDestNome = UCase(Trim(TxtNomeEmpresaSol(2).Text))
xDestEnd = UCase(Trim(TxtEnderecoSol(2).Text))
xDestCidade = UCase(Trim(TxtCidadeSol(2).Text))
xDestUF = UCase(Trim(TxtUFSol(2).Text))
xDataColeta = TxtDataColeta.Text
xHoraColeta = TxtHorario.Text
xTipoFrete = UCase(Trim(TxtTipoFrete.Text))
xnatureza = UCase(Trim(TxtNatureza.Text))
xEspecie = UCase(Trim(ComboEspecie.Text))
    If IsNumeric(TxtValMerc.Text) = False Then
    TxtValMerc.Text = 0
    End If
xvalmerc = CDbl(TxtValMerc.Text)
xpeso = CDbl(TxtPeso.Text)
xvolumes = CDbl(TxtVolumes.Text)
xNfs = UCase(Trim(TxtNFs.Text))
xobs = UCase(Trim(TxtOBS.Text))
xEmissor = xusuario
xTem_Ocorr = "N"

'xEmissaoColeta = xEmissao
'xCod_Ocorr = "91"
'xDescricao = "COLETA PROGRAMADA"
'xObsOcorr = "DATA AGENDADA PRIMORDIALMENTE"
'xdata = CDate(TxtDataColeta.Text)
'xhora = UCase(Trim(TxtHorario.Text))

de_informa.ColetaIns xFilialColeta, xfilial, XColeta, xprioridade, xSolCGC, xSolEmpNome, xSolNome, xSolTel, _
xRemetCgc, xRemetNome, xRemetEnd, xRemetCidade, xRemetUF, xColetaEnd, _
xColetaCidade, xColetaUF, xDestCGC, xDestNome, xDestEnd, xDestCidade, _
xDestUF, CDate(xDataColeta), xHoraColeta, xTipoFrete, xnatureza, xEspecie, xvalmerc, xpeso, xvolumes, xNfs, _
xobs, CDate(datahora("DATA")), datahora("HORA"), xEmissor, xTem_Ocorr

'LOG DE USUÁRIO
de_informa.ins_LogUsuario "EMISSAO", xusuario, "COLETA:" & transctc(xfilial, XColeta)


Call ImprimeColeta(xFilialColeta)

MsgBox "Ordem de Coleta emitida com sucesso! Anote o número de sua Coleta: " & Mid(xFilialColeta, 1, 2) & "-" & Trim(Str(Val(Mid(xFilialColeta, 3)))), vbInformation

de_informa.cn_informa.CommitTrans

Call limpatela(Me)
TxtFilial.SetFocus
ComboEspecie.Text = ""
LblTipoFrete(0).FontBold = False
LblTipoFrete(1).FontBold = False
LblPrioridade(0).FontBold = False
LblPrioridade(1).FontBold = False
LblPrioridade(2).FontBold = False

End Sub

Private Sub ComboEspecie_KeyPress(KeyAscii As Integer)
Dim xTextoVelho As String, xTextoNovo As String, Y As Integer

'Asc (UCase(Chr(KeyAscii)))

    If KeyAscii <> 13 And KeyAscii <> 8 Then
    xTextoVelho = Left(ComboEspecie.Text, ComboEspecie.SelStart) & Chr(KeyAscii)
    xTextoNovo = ""
        For Y = 0 To ComboEspecie.ListCount - 1
            If Len(xTextoVelho) <= Len(ComboEspecie.List(Y)) Then
                If UCase(Mid(ComboEspecie.List(Y), 1, Len(xTextoVelho))) = UCase(xTextoVelho) Then
                xTextoNovo = Mid(ComboEspecie.List(Y), Len(xTextoVelho) + 1)
                Y = ComboEspecie.ListCount
                End If
            End If
        Next
    ComboEspecie.Text = UCase(xTextoVelho) & xTextoNovo
    ComboEspecie.SelStart = Len(xTextoVelho)
    ComboEspecie.SelLength = 1000
    ElseIf KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    ElseIf KeyAscii = 8 Then
        If Len(ComboEspecie.Text) > 0 Then
            If ComboEspecie.SelStart > 0 Then
            xTextoVelho = Mid(ComboEspecie.Text, 1, ComboEspecie.SelStart - 1)
            Else
            xTextoVelho = Mid(ComboEspecie.Text, 1, ComboEspecie.SelStart)
            End If
        ComboEspecie.Text = UCase(xTextoVelho)
        ComboEspecie.SelStart = Len(xTextoVelho)
        ComboEspecie.SelLength = 1000
        End If
    End If
KeyAscii = 0
End Sub

Private Sub ComboEspecie_LostFocus()
Dim Y As Integer, xTexto As String

xTexto = ""

        For Y = 0 To ComboEspecie.ListCount - 1
            If UCase(Trim((ComboEspecie.Text))) = UCase(Trim(ComboEspecie.List(Y))) Then
            xTexto = ComboEspecie.List(Y)
            Y = ComboEspecie.ListCount
            End If
        Next
ComboEspecie.Text = xTexto
End Sub

Private Sub FlexAUX_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
    TxtCGCSol(xIndexBusca).Text = FlexAUX(Index).TextMatrix(FlexAUX(Index).Row, 0)
    TxtNomeEmpresaSol(xIndexBusca).Text = FlexAUX(Index).TextMatrix(FlexAUX(Index).Row, 1)
    TxtEnderecoSol(xIndexBusca).Text = FlexAUX(Index).TextMatrix(FlexAUX(Index).Row, 2)
    TxtCidadeSol(xIndexBusca).Text = FlexAUX(Index).TextMatrix(FlexAUX(Index).Row, 3)
    TxtUFSol(xIndexBusca).Text = FlexAUX(Index).TextMatrix(FlexAUX(Index).Row, 4)
    TxtNomeSol(xIndexBusca).Text = FlexAUX(Index).TextMatrix(FlexAUX(Index).Row, 5)
    TxtTelSol(xIndexBusca).Text = FlexAUX(Index).TextMatrix(FlexAUX(Index).Row, 6)
    FlexAUX_LostFocus (Index)
    ElseIf KeyAscii = 27 Then
    FlexAUX_LostFocus (Index)
    End If
End Sub

Private Sub FlexAUX_LostFocus(Index As Integer)
    TxtNomeSol(xIndexBusca).SetFocus
    FlexAUX(Index).Clear
    FlexAUX(Index).Rows = 0
    FraAux(Index).Visible = False
    DoEvents
End Sub

Private Sub Form_Load()
mdiInforma.Toolbar1.Enabled = False

If de_informa.rsSel_Especie.State = 1 Then de_informa.rsSel_Especie.Close
de_informa.Sel_Especie
    Do Until de_informa.rsSel_Especie.EOF
    ComboEspecie.AddItem PriMaiuscula(de_informa.rsSel_Especie.Fields("especie"))
    de_informa.rsSel_Especie.MoveNext
    Loop
End Sub


Private Sub Label2_Click()

End Sub

Private Sub TxtBuscaDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtBuscaRem_GotFocus()
TxtBuscaRem.SelStart = 0
TxtBuscaRem.SelLength = 300
End Sub

Private Sub TxtBuscaRem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInforma.Toolbar1.Enabled = True
End Sub

Private Sub TxtBuscaSol_Change(Index As Integer)
Dim X As Integer
X = TxtBuscaSol(Index).SelStart
TxtBuscaSol(Index).Text = UCase(TxtBuscaSol(Index).Text)
TxtBuscaSol(Index).SelStart = X
End Sub

Private Sub TxtBuscaSol_GotFocus(Index As Integer)
TxtBuscaSol(Index).SelStart = 0
TxtBuscaSol(Index).SelLength = 300
End Sub

Private Sub TxtBuscaSol_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtBuscaSol_LostFocus(Index As Integer)
Dim xRs As Recordset
xIndexBusca = Index

Me.MousePointer = 11
DoEvents
If Len(Trim(TxtBuscaSol(Index).Text)) > 0 Then
If de_informa.rsSel_CadCliApelido.State = 1 Then de_informa.rsSel_CadCliApelido.Close
de_informa.Sel_cadcliapelido Trim(TxtBuscaSol(Index).Text) & "%"
    If de_informa.rsSel_CadCliApelido.RecordCount = 0 Then
        If de_informa.rsSel_CadCliNome.State = 1 Then de_informa.rsSel_CadCliNome.Close
        de_informa.Sel_CadClinome Trim(TxtBuscaSol(Index).Text) & "%"
            If de_informa.rsSel_CadCliNome.RecordCount = 0 Then
                If de_informa.rsSel_CadCliCGC2.State = 1 Then de_informa.rsSel_CadCliCGC2.Close
                de_informa.Sel_cadclicgc2 Trim(TxtBuscaSol(Index).Text) & "%"
                Set xRs = de_informa.rsSel_CadCliCGC2
            Else
            Set xRs = de_informa.rsSel_CadCliNome
            End If
    Else
    Set xRs = de_informa.rsSel_CadCliApelido
    End If
 
    If xRs.RecordCount = 1 Then
    TxtCGCSol(Index).Text = xRs.Fields("cgc")
    TxtNomeEmpresaSol(Index).Text = PriMaiuscula(xRs.Fields("nome"))
    TxtCidadeSol(Index).Text = PriMaiuscula(xRs.Fields("CIDADE"))
    TxtEnderecoSol(Index).Text = PriMaiuscula(xRs.Fields("ENDERECO"))
    TxtUFSol(Index).Text = UCase(xRs.Fields("UF"))
        If IsNull(xRs.Fields("CONTATO1")) = False And xRs.Fields("CONTATO1") <> "" Then
        TxtNomeSol(Index).Text = PriMaiuscula(xRs.Fields("CONTATO1"))
        End If
        
        If IsNull(xRs.Fields("FONECONTATO1")) = False And xRs.Fields("FONECONTATO1") <> "" Then
        TxtTelSol(Index).Text = PriMaiuscula(xRs.Fields("FONECONTATO1"))
        End If
    
    ElseIf xRs.RecordCount > 1 Then
    FlexAUX(0).Clear
    FlexAUX(0).Rows = xRs.RecordCount + 1
    FlexAUX(0).Cols = 7
    FlexAUX(0).FixedRows = 1
    FlexAUX(0).FixedCols = 0
    FlexAUX(0).TextMatrix(0, 0) = "CGC"
    FlexAUX(0).TextMatrix(0, 1) = "Nome do Cliente"
    FlexAUX(0).TextMatrix(0, 2) = "Endereço do Cliente"
    FlexAUX(0).TextMatrix(0, 3) = "Cidade"
    FlexAUX(0).TextMatrix(0, 4) = "UF"
    FlexAUX(0).TextMatrix(0, 5) = "Contato"
    FlexAUX(0).TextMatrix(0, 6) = "Telefone"
    FlexAUX(0).ColWidth(0) = 1500
    FlexAUX(0).ColWidth(1) = 4000
    FlexAUX(0).ColWidth(2) = 5000
    FlexAUX(0).ColWidth(3) = 4000
    FlexAUX(0).ColWidth(4) = 800
    FlexAUX(0).ColWidth(5) = 1500
    FlexAUX(0).ColWidth(6) = 1500
    Y = 0
        Do Until xRs.EOF
        Y = Y + 1
        If IsNull(xRs.Fields("cgc")) = False Then FlexAUX(0).TextMatrix(Y, 0) = xRs.Fields("cgc")
        If IsNull(xRs.Fields("nome")) = False Then FlexAUX(0).TextMatrix(Y, 1) = PriMaiuscula(xRs.Fields("nome"))
        If IsNull(xRs.Fields("endereco")) = False Then FlexAUX(0).TextMatrix(Y, 2) = PriMaiuscula(xRs.Fields("endereco"))
        If IsNull(xRs.Fields("cidade")) = False Then FlexAUX(0).TextMatrix(Y, 3) = PriMaiuscula(xRs.Fields("cidade"))
        If IsNull(xRs.Fields("uf")) = False Then FlexAUX(0).TextMatrix(Y, 4) = xRs.Fields("uf")
        If IsNull(xRs.Fields("contato1")) = False Then FlexAUX(0).TextMatrix(Y, 5) = PriMaiuscula(xRs.Fields("contato1"))
        If IsNull(xRs.Fields("fonecontato1")) = False Then FlexAUX(0).TextMatrix(Y, 6) = xRs.Fields("fonecontato1")
        xRs.MoveNext
        Loop
    FraAux(0).Visible = True
    FlexAUX(0).SetFocus
    ElseIf xRs.RecordCount = 0 Then
        If MsgBox("Cliente não encontrado! Você deseja cadastrar este Cliente?", vbYesNo + vbCritical) = vbYes Then
        frmCadClientes.Show 1
        End If
    End If
TxtBuscaSol(Index).Text = ""
'TxtBuscaSol (xIndexBusca).SetFocus
End If

Me.MousePointer = 0
DoEvents
End Sub


Private Sub TxtBuscarem_Change()
Dim X As Integer
X = TxtBuscaRem.SelStart
TxtBuscaRem.Text = UCase(TxtBuscaRem.Text)
TxtBuscaRem.SelStart = X
End Sub

Private Sub TxtBuscarem_LostFocus()
Me.MousePointer = 11
DoEvents
If Len(Trim(TxtBuscaRem.Text)) > 0 Then
If de_informa.rsSel_CadCliApelido.State = 1 Then de_informa.rsSel_CadCliApelido.Close
de_informa.Sel_cadcliapelido Trim(TxtBuscaRem.Text) & "%"
    If de_informa.rsSel_CadCliApelido.RecordCount = 1 Then
    TxtCGCRem.Text = de_informa.rsSel_CadCliApelido.Fields("cgc")
    TxtNomeRem.Text = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("nome"))
    TxtEnderecoRem.Text = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("endereco"))
    TxtCidadeRem.Text = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("cidade"))
    TxtUFRem.Text = de_informa.rsSel_CadCliApelido.Fields("uf")
    ElseIf de_informa.rsSel_CadCliApelido.RecordCount > 1 Then
    FlexAUX(1).Clear
    FlexAUX(1).Rows = de_informa.rsSel_CadCliApelido.RecordCount + 1
    FlexAUX(1).Cols = 5
    FlexAUX(1).FixedRows = 1
    FlexAUX(1).FixedCols = 0
    FlexAUX(1).TextMatrix(0, 0) = "CGC"
    FlexAUX(1).TextMatrix(0, 1) = "Nome do Cliente"
    FlexAUX(1).TextMatrix(0, 2) = "Endereço do Cliente"
    FlexAUX(1).TextMatrix(0, 3) = "Cidade"
    FlexAUX(1).TextMatrix(0, 4) = "UF"
    FlexAUX(1).ColWidth(0) = 1500
    FlexAUX(1).ColWidth(1) = 4000
    FlexAUX(1).ColWidth(2) = 5000
    FlexAUX(1).ColWidth(3) = 4000
    FlexAUX(1).ColWidth(4) = 800
    Y = 0
        Do Until de_informa.rsSel_CadCliApelido.EOF
        Y = Y + 1
        If IsNull(de_informa.rsSel_CadCliApelido.Fields("cgc")) = False Then FlexAUX(1).TextMatrix(Y, 0) = de_informa.rsSel_CadCliApelido.Fields("cgc")
        If IsNull(de_informa.rsSel_CadCliApelido.Fields("nome")) = False Then FlexAUX(1).TextMatrix(Y, 1) = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("nome"))
        If IsNull(de_informa.rsSel_CadCliApelido.Fields("endereco")) = False Then FlexAUX(1).TextMatrix(Y, 2) = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("endereco"))
        If IsNull(de_informa.rsSel_CadCliApelido.Fields("cidade")) = False Then FlexAUX(1).TextMatrix(Y, 3) = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("cidade"))
        If IsNull(de_informa.rsSel_CadCliApelido.Fields("uf")) = False Then FlexAUX(1).TextMatrix(Y, 4) = de_informa.rsSel_CadCliApelido.Fields("uf")
        de_informa.rsSel_CadCliApelido.MoveNext
        Loop
    FraAux(1).Visible = True
    FlexAUX(1).SetFocus
    ElseIf de_informa.rsSel_CadCliApelido.RecordCount = 0 Then
    If de_informa.rsSel_CadCliNome.State = 1 Then de_informa.rsSel_CadCliNome.Close
    de_informa.Sel_CadClinome Trim(TxtBuscaRem.Text) & "%"
        If de_informa.rsSel_CadCliNome.RecordCount = 1 Then
        TxtCGCRem.Text = de_informa.rsSel_CadCliNome.Fields("cgc")
        TxtNomeRem.Text = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("nome"))
        TxtEnderecoRem.Text = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("endereco"))
        TxtCidadeRem.Text = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("cidade"))
        TxtUFRem.Text = de_informa.rsSel_CadCliNome.Fields("uf")
        ElseIf de_informa.rsSel_CadCliNome.RecordCount > 1 Then
        FlexAUX(1).Clear
        FlexAUX(1).Rows = de_informa.rsSel_CadCliNome.RecordCount + 1
        FlexAUX(1).Cols = 5
        FlexAUX(1).FixedRows = 1
        FlexAUX(1).FixedCols = 0
        FlexAUX(1).TextMatrix(0, 0) = "CGC"
        FlexAUX(1).TextMatrix(0, 1) = "Nome do Cliente"
        FlexAUX(1).TextMatrix(0, 2) = "Endereço do Cliente"
        FlexAUX(1).TextMatrix(0, 3) = "Cidade"
        FlexAUX(1).TextMatrix(0, 4) = "UF"
        FlexAUX(1).ColWidth(0) = 1500
        FlexAUX(1).ColWidth(1) = 4000
        FlexAUX(1).ColWidth(2) = 5000
        FlexAUX(1).ColWidth(3) = 4000
        FlexAUX(1).ColWidth(4) = 800
        Y = 0
            Do Until de_informa.rsSel_CadCliNome.EOF
            Y = Y + 1
            If IsNull(de_informa.rsSel_CadCliNome.Fields("cgc")) = False Then FlexAUX(1).TextMatrix(Y, 0) = de_informa.rsSel_CadCliNome.Fields("cgc")
            If IsNull(de_informa.rsSel_CadCliNome.Fields("nome")) = False Then FlexAUX(1).TextMatrix(Y, 1) = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("nome"))
            If IsNull(de_informa.rsSel_CadCliNome.Fields("endereco")) = False Then FlexAUX(1).TextMatrix(Y, 2) = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("endereco"))
            If IsNull(de_informa.rsSel_CadCliNome.Fields("cidade")) = False Then FlexAUX(1).TextMatrix(Y, 3) = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("cidade"))
            If IsNull(de_informa.rsSel_CadCliNome.Fields("uf")) = False Then FlexAUX(1).TextMatrix(Y, 4) = de_informa.rsSel_CadCliNome.Fields("uf")
            de_informa.rsSel_CadCliNome.MoveNext
            Loop
        FraAux(1).Visible = True
        FlexAUX(1).SetFocus
        ElseIf de_informa.rsSel_CadCliNome.RecordCount = 0 Then
        If de_informa.rsSel_CadCliCGC2.State = 1 Then de_informa.rsSel_CadCliCGC2.Close
        de_informa.Sel_cadclicgc2 Trim(TxtBuscaRem.Text) & "%"
            If de_informa.rsSel_CadCliCGC2.RecordCount = 1 Then
            TxtCGCRem.Text = de_informa.rsSel_CadCliCGC2.Fields("cgc")
            TxtNomeRem.Text = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("nome"))
            TxtEnderecoRem.Text = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("endereco"))
            TxtCidadeRem.Text = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("cidade"))
            TxtUFRem.Text = de_informa.rsSel_CadCliCGC2.Fields("uf")
            ElseIf de_informa.rsSel_CadCliCGC2.RecordCount > 1 Then
            FlexAUX(1).Clear
            FlexAUX(1).Rows = de_informa.rsSel_CadCliCGC2.RecordCount + 1
            FlexAUX(1).Cols = 5
            FlexAUX(1).FixedRows = 1
            FlexAUX(1).FixedCols = 0
            FlexAUX(1).TextMatrix(0, 0) = "CGC"
            FlexAUX(1).TextMatrix(0, 1) = "Nome do Cliente"
            FlexAUX(1).TextMatrix(0, 2) = "Endereço do Cliente"
            FlexAUX(1).TextMatrix(0, 3) = "Cidade"
            FlexAUX(1).TextMatrix(0, 4) = "UF"
            FlexAUX(1).ColWidth(0) = 1500
            FlexAUX(1).ColWidth(1) = 4000
            FlexAUX(1).ColWidth(2) = 5000
            FlexAUX(1).ColWidth(3) = 4000
            FlexAUX(1).ColWidth(4) = 800
            Y = 0
                Do Until de_informa.rsSel_CadCliCGC2.EOF
                Y = Y + 1
                If IsNull(de_informa.rsSel_CadCliCGC2.Fields("cgc")) = False Then FlexAUX(1).TextMatrix(Y, 0) = de_informa.rsSel_CadCliCGC2.Fields("cgc")
                If IsNull(de_informa.rsSel_CadCliCGC2.Fields("nome")) = False Then FlexAUX(1).TextMatrix(Y, 1) = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("nome"))
                If IsNull(de_informa.rsSel_CadCliCGC2.Fields("endereco")) = False Then FlexAUX(1).TextMatrix(Y, 2) = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("endereco"))
                If IsNull(de_informa.rsSel_CadCliCGC2.Fields("cidade")) = False Then FlexAUX(1).TextMatrix(Y, 3) = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("cidade"))
                If IsNull(de_informa.rsSel_CadCliCGC2.Fields("uf")) = False Then FlexAUX(1).TextMatrix(Y, 4) = de_informa.rsSel_CadCliCGC2.Fields("uf")
                de_informa.rsSel_CadCliCGC2.MoveNext
                Loop
            FraAux(1).Visible = True
            FlexAUX(1).SetFocus
            Else
                If MsgBox("Cliente não encontrado! Você deseja cadastrar este Cliente?", vbYesNo + vbCritical) = vbYes Then
                frmCadClientes.Show 1
                End If
            End If
        End If
    End If
TxtBuscaRem.Text = ""
'TxtBuscaRem.SetFocus
End If
Me.MousePointer = 0
DoEvents
End Sub

Private Sub TxtBuscaDES_Change()
Dim X As Integer
X = TxtBuscaDes.SelStart
TxtBuscaDes.Text = UCase(TxtBuscaDes.Text)
TxtBuscaDes.SelStart = X
End Sub

Private Sub TxtBuscaDES_LostFocus()
Me.MousePointer = 11
DoEvents
If Len(Trim(TxtBuscaDes.Text)) > 0 Then
If de_informa.rsSel_CadCliApelido.State = 1 Then de_informa.rsSel_CadCliApelido.Close
de_informa.Sel_cadcliapelido Trim(TxtBuscaDes.Text) & "%"
    If de_informa.rsSel_CadCliApelido.RecordCount = 1 Then
    TxtCGCDES.Text = de_informa.rsSel_CadCliApelido.Fields("cgc")
    TxtNomeDES.Text = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("nome"))
    TxtEnderecoDes.Text = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("endereco"))
    TxtCidadeDes.Text = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("cidade"))
    TxtUfDES.Text = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("uf"))
    ElseIf de_informa.rsSel_CadCliApelido.RecordCount > 1 Then
    FlexAUX(2).Clear
    FlexAUX(2).Rows = de_informa.rsSel_CadCliApelido.RecordCount + 1
    FlexAUX(2).Cols = 5
    FlexAUX(2).FixedRows = 1
    FlexAUX(2).FixedCols = 0
    FlexAUX(2).TextMatrix(0, 0) = "CGC"
    FlexAUX(2).TextMatrix(0, 1) = "Nome do Cliente"
    FlexAUX(2).TextMatrix(0, 2) = "Endereço do Cliente"
    FlexAUX(2).TextMatrix(0, 3) = "Cidade"
    FlexAUX(2).TextMatrix(0, 4) = "UF"
    FlexAUX(2).ColWidth(0) = 1500
    FlexAUX(2).ColWidth(1) = 4000
    FlexAUX(2).ColWidth(2) = 5000
    FlexAUX(2).ColWidth(3) = 4000
    FlexAUX(2).ColWidth(4) = 800
    Y = 0
        Do Until de_informa.rsSel_CadCliApelido.EOF
        Y = Y + 1
        If IsNull(de_informa.rsSel_CadCliApelido.Fields("cgc")) = False Then FlexAUX(2).TextMatrix(Y, 0) = de_informa.rsSel_CadCliApelido.Fields("cgc")
        If IsNull(de_informa.rsSel_CadCliApelido.Fields("nome")) = False Then FlexAUX(2).TextMatrix(Y, 1) = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("nome"))
        If IsNull(de_informa.rsSel_CadCliApelido.Fields("endereco")) = False Then FlexAUX(2).TextMatrix(Y, 2) = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("endereco"))
        If IsNull(de_informa.rsSel_CadCliApelido.Fields("cidade")) = False Then FlexAUX(2).TextMatrix(Y, 3) = PriMaiuscula(de_informa.rsSel_CadCliApelido.Fields("cidade"))
        If IsNull(de_informa.rsSel_CadCliApelido.Fields("uf")) = False Then FlexAUX(2).TextMatrix(Y, 4) = de_informa.rsSel_CadCliApelido.Fields("uf")
        de_informa.rsSel_CadCliApelido.MoveNext
        Loop
    FraAux(2).Visible = True
    FlexAUX(2).SetFocus
    ElseIf de_informa.rsSel_CadCliApelido.RecordCount = 0 Then
    If de_informa.rsSel_CadCliNome.State = 1 Then de_informa.rsSel_CadCliNome.Close
    de_informa.Sel_CadClinome Trim(TxtBuscaDes.Text) & "%"
        If de_informa.rsSel_CadCliNome.RecordCount = 1 Then
        TxtCGCDES.Text = de_informa.rsSel_CadCliNome.Fields("cgc")
        TxtNomeDES.Text = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("nome"))
        TxtEnderecoDes.Text = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("endereco"))
        TxtCidadeDes.Text = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("cidade"))
        TxtUfDES.Text = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("uf"))
        ElseIf de_informa.rsSel_CadCliNome.RecordCount > 1 Then
        FlexAUX(2).Clear
        FlexAUX(2).Rows = de_informa.rsSel_CadCliNome.RecordCount + 1
        FlexAUX(2).Cols = 5
        FlexAUX(2).FixedRows = 1
        FlexAUX(2).FixedCols = 0
        FlexAUX(2).TextMatrix(0, 0) = "CGC"
        FlexAUX(2).TextMatrix(0, 1) = "Nome do Cliente"
        FlexAUX(2).TextMatrix(0, 2) = "Endereço do Cliente"
        FlexAUX(2).TextMatrix(0, 3) = "Cidade"
        FlexAUX(2).TextMatrix(0, 4) = "UF"
        FlexAUX(2).ColWidth(0) = 1500
        FlexAUX(2).ColWidth(1) = 4000
        FlexAUX(2).ColWidth(2) = 5000
        FlexAUX(2).ColWidth(3) = 4000
        FlexAUX(2).ColWidth(4) = 800
        Y = 0
            Do Until de_informa.rsSel_CadCliNome.EOF
            Y = Y + 1
            If IsNull(de_informa.rsSel_CadCliNome.Fields("cgc")) = False Then FlexAUX(2).TextMatrix(Y, 0) = de_informa.rsSel_CadCliNome.Fields("cgc")
            If IsNull(de_informa.rsSel_CadCliNome.Fields("nome")) = False Then FlexAUX(2).TextMatrix(Y, 1) = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("nome"))
            If IsNull(de_informa.rsSel_CadCliNome.Fields("endereco")) = False Then FlexAUX(2).TextMatrix(Y, 2) = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("endereco"))
            If IsNull(de_informa.rsSel_CadCliNome.Fields("cidade")) = False Then FlexAUX(2).TextMatrix(Y, 3) = PriMaiuscula(de_informa.rsSel_CadCliNome.Fields("cidade"))
            If IsNull(de_informa.rsSel_CadCliNome.Fields("uf")) = False Then FlexAUX(2).TextMatrix(Y, 4) = de_informa.rsSel_CadCliNome.Fields("uf")
            de_informa.rsSel_CadCliNome.MoveNext
            Loop
        FraAux(2).Visible = True
        FlexAUX(2).SetFocus
        ElseIf de_informa.rsSel_CadCliNome.RecordCount = 0 Then
        If de_informa.rsSel_CadCliCGC2.State = 1 Then de_informa.rsSel_CadCliCGC2.Close
        de_informa.Sel_cadclicgc2 Trim(TxtBuscaDes.Text) & "%"
            If de_informa.rsSel_CadCliCGC2.RecordCount = 1 Then
            TxtCGCDES.Text = de_informa.rsSel_CadCliCGC2.Fields("cgc")
            TxtNomeDES.Text = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("nome"))
            TxtEnderecoDes.Text = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("endereco"))
            TxtCidadeDes.Text = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("cidade"))
            TxtUfDES.Text = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("uf"))
            ElseIf de_informa.rsSel_CadCliCGC2.RecordCount > 1 Then
            FlexAUX(2).Clear
            FlexAUX(2).Rows = de_informa.rsSel_CadCliCGC2.RecordCount + 1
            FlexAUX(2).Cols = 5
            FlexAUX(2).FixedRows = 1
            FlexAUX(2).FixedCols = 0
            FlexAUX(2).TextMatrix(0, 0) = "CGC"
            FlexAUX(2).TextMatrix(0, 1) = "Nome do Cliente"
            FlexAUX(2).TextMatrix(0, 2) = "Endereço do Cliente"
            FlexAUX(2).TextMatrix(0, 3) = "Cidade"
            FlexAUX(2).TextMatrix(0, 4) = "UF"
            FlexAUX(2).ColWidth(0) = 1500
            FlexAUX(2).ColWidth(1) = 4000
            FlexAUX(2).ColWidth(2) = 5000
            FlexAUX(2).ColWidth(3) = 4000
            FlexAUX(2).ColWidth(4) = 800
            Y = 0
                Do Until de_informa.rsSel_CadCliCGC2.EOF
                Y = Y + 1
                If IsNull(de_informa.rsSel_CadCliCGC2.Fields("cgc")) = False Then FlexAUX(2).TextMatrix(Y, 0) = de_informa.rsSel_CadCliCGC2.Fields("cgc")
                If IsNull(de_informa.rsSel_CadCliCGC2.Fields("nome")) = False Then FlexAUX(2).TextMatrix(Y, 1) = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("nome"))
                If IsNull(de_informa.rsSel_CadCliCGC2.Fields("endereco")) = False Then FlexAUX(2).TextMatrix(Y, 2) = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("endereco"))
                If IsNull(de_informa.rsSel_CadCliCGC2.Fields("cidade")) = False Then FlexAUX(2).TextMatrix(Y, 3) = PriMaiuscula(de_informa.rsSel_CadCliCGC2.Fields("cidade"))
                If IsNull(de_informa.rsSel_CadCliCGC2.Fields("uf")) = False Then FlexAUX(2).TextMatrix(Y, 4) = de_informa.rsSel_CadCliCGC2.Fields("uf")
                de_informa.rsSel_CadCliCGC2.MoveNext
                Loop
            FraAux(2).Visible = True
            FlexAUX(2).SetFocus
            Else
                If MsgBox("Cliente não encontrado! Você deseja cadastrar este Cliente?", vbYesNo + vbCritical) = vbYes Then
                frmCadClientes.Show 1
                End If
            End If
        End If
    End If
TxtBuscaDes.Text = ""
'TxtBuscaDes.SetFocus
End If
Me.MousePointer = 0
DoEvents
End Sub

Private Sub TxtCGCSol_GotFocus(Index As Integer)
TxtCGCSol(Index).SelStart = 0
TxtCGCSol(Index).SelLength = 300
End Sub

Private Sub TxtCGCSol_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtCidadeColeta_Change()
Dim X As Integer
X = TxtCidadeColeta.SelStart
TxtCidadeColeta.Text = PriMaiuscula(TxtCidadeColeta.Text)
TxtCidadeColeta.SelStart = X
End Sub

Private Sub TxtCidadeColeta_GotFocus()
    If Len(Trim(TxtCidadeColeta.Text)) = 0 Then
    TxtCidadeColeta.Text = TxtCidadeSol(1).Text
    TxtCidadeColeta.SelStart = 0
    TxtCidadeColeta.SelLength = 300
    Else
    TxtCidadeColeta.SelStart = 0
    TxtCidadeColeta.SelLength = 300
    End If
End Sub

Private Sub TxtCidadeColeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtCidadeDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub


Private Sub TxtCidadeRem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtColeta_Change()
    If Len(TxtColeta.Text) = 10 Then
    SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtColeta_GotFocus()
TxtColeta.SelStart = 0
TxtColeta.SelLength = 100
End Sub

Private Sub txtcoleta_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False Then
        If KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        ElseIf KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtCidadeSol_GotFocus(Index As Integer)
TxtCidadeSol(Index).SelStart = 0
TxtCidadeSol(Index).SelLength = 300
End Sub

Private Sub TxtCidadeSol_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtDataColeta_GotFocus()
    If TxtDataColeta.Text = "" Then
    TxtDataColeta.Text = Date
    Else
    Call Date_MskEdBox_GotFocus(TxtDataColeta)
    End If
End Sub

Private Sub TxtDataColeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtDataColeta_LostFocus()
Call Date_MskEdBox_LostFocus(TxtDataColeta)
End Sub

Private Sub TxtEnderecoColeta_Change()
Dim X As Integer
X = TxtEnderecoColeta.SelStart
TxtEnderecoColeta.Text = PriMaiuscula(TxtEnderecoColeta.Text)
TxtEnderecoColeta.SelStart = X
End Sub

Private Sub TxtEnderecoColeta_GotFocus()
    If Len(Trim(TxtEnderecoColeta.Text)) = 0 Then
    TxtEnderecoColeta.Text = TxtEnderecoSol(1).Text
    TxtEnderecoColeta.SelStart = 0
    TxtEnderecoColeta.SelLength = 300
    Else
    TxtEnderecoColeta.SelStart = 0
    TxtEnderecoColeta.SelLength = 300
    End If
End Sub

Private Sub TxtEnderecoColeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtEnderecoDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtEnderecoRem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtEnderecoSol_GotFocus(Index As Integer)
TxtEnderecoSol(Index).SelStart = 0
TxtEnderecoSol(Index).SelLength = 300
End Sub

Private Sub TxtEnderecoSol_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtFilial_Change()
    If Len(TxtFilial.Text) = 2 Then
    SendKeys "{TAB}"
    End If
End Sub

Private Sub TxtFilial_GotFocus()
TxtFilial.SelStart = 0
TxtFilial.SelLength = 100
End Sub

Private Sub txtfilial_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False Then
        If KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        ElseIf KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
    End If
End Sub




Private Sub TxtHorario_Change()
Dim X As Integer
X = TxtHorario.SelStart
TxtHorario.Text = UCase(TxtHorario.Text)
TxtHorario.SelStart = X
End Sub

Private Sub TxtHorario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtNatureza_Change()
Dim X As Integer
X = TxtNatureza.SelStart
TxtNatureza.Text = PriMaiuscula(TxtNatureza.Text)
TxtNatureza.SelStart = X
End Sub

Private Sub TxtNatureza_GotFocus()
TxtNatureza.SelStart = 0
TxtNatureza.SelLength = 300
End Sub

Private Sub TxtNatureza_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtNFs_Change()
Dim X As Integer
X = TxtNFs.SelStart
TxtNFs.Text = PriMaiuscula(TxtNFs.Text)
TxtNFs.SelStart = X
End Sub

Private Sub TxtNFs_GotFocus()
TxtNFs.SelStart = 0
TxtNFs.SelLength = 300
End Sub

Private Sub TxtNFs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtNomeDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtNomeEmpresaSol_GotFocus(Index As Integer)
TxtNomeEmpresaSol(Index).SelStart = 0
TxtNomeEmpresaSol(Index).SelLength = 300
End Sub

Private Sub TxtNomeEmpresaSol_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtNomeSol_Change(Index As Integer)
Dim X As Integer
X = TxtNomeSol(Index).SelStart
TxtNomeSol(Index).Text = PriMaiuscula(TxtNomeSol(Index).Text)
TxtNomeSol(Index).SelStart = X
End Sub

Private Sub TxtNomeSol_GotFocus(Index As Integer)
TxtNomeSol(Index).SelStart = 0
TxtNomeSol(Index).SelLength = 300
End Sub

Private Sub TxtNomeSol_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtOBS_Change()
Dim X As Integer
X = TxtOBS.SelStart
TxtOBS.Text = UCase(TxtOBS.Text)
TxtOBS.SelStart = X
End Sub

Private Sub TxtOBS_GotFocus()
TxtVolumes.SelStart = 0
TxtVolumes.SelLength = 300
End Sub

Private Sub TxtOBS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtPeso_GotFocus()
TxtPeso.SelStart = 0
TxtPeso.SelLength = 300
End Sub

Private Sub TxtPeso_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False Then
        If KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        ElseIf KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtTelSol_Change(Index As Integer)
Dim X As Integer
X = TxtTelSol(Index).SelStart
TxtTelSol(Index).Text = PriMaiuscula(TxtTelSol(Index).Text)
TxtTelSol(Index).SelStart = X
End Sub

Private Sub TxtTelSol_GotFocus(Index As Integer)
TxtTelSol(Index).SelStart = 0
TxtTelSol(Index).SelLength = 300
End Sub

Private Sub TxtTelSol_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtTipoFrete_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) <> "1" And Chr(KeyAscii) <> "2" Then
        If KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        ElseIf KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Exit Sub
        End If
    End If
    
    If KeyAscii > 0 And KeyAscii <> 8 Then
        If Chr(KeyAscii) = "1" Then
        LblTipoFrete(0).FontBold = True
        LblTipoFrete(1).FontBold = False
        TxtTipoFrete.Text = "1"
        ElseIf Chr(KeyAscii) = "2" Then
        LblTipoFrete(1).FontBold = True
        LblTipoFrete(0).FontBold = False
        TxtTipoFrete.Text = "2"
        End If
    ElseIf KeyAscii = 0 Or KeyAscii = 8 Then
    LblTipoFrete(1).FontBold = False
    LblTipoFrete(0).FontBold = False
        If KeyAscii = 8 Then
        TxtTipoFrete.Text = ""
        End If
    End If
KeyAscii = 0
DoEvents
End Sub

Private Sub TxtUFColeta_Change()
Dim X As Integer
X = TxtUFColeta.SelStart
TxtUFColeta.Text = UCase(TxtUFColeta.Text)
TxtUFColeta.SelStart = X
End Sub

Private Sub TxtUFColeta_GotFocus()
    If Len(Trim(TxtUFColeta.Text)) = 0 Then
    TxtUFColeta.Text = TxtUFSol(1).Text
    TxtUFColeta.SelStart = 0
    TxtUFColeta.SelLength = 300
    Else
    TxtUFColeta.SelStart = 0
    TxtUFColeta.SelLength = 300
    End If
End Sub

Private Sub TxtUFColeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtUFDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtUFRem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtUFSol_GotFocus(Index As Integer)
TxtUFSol(Index).SelStart = 0
TxtUFSol(Index).SelLength = 300
End Sub

Private Sub TxtUFSol_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
    End If
End Sub

Private Sub TxtValMerc_GotFocus()
TxtValMerc.SelStart = 0
TxtValMerc.SelLength = 300
End Sub

Private Sub TxtValMerc_KeyPress(KeyAscii As Integer)
Call TextMoneyBox_KeyPress(KeyAscii, TxtValMerc)
End Sub

Private Sub TxtVolumes_GotFocus()
TxtVolumes.SelStart = 0
TxtVolumes.SelLength = 300
End Sub

Private Sub TxtVolumes_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = False Then
        If KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        ElseIf KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtPrioridade_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) <> "1" And Chr(KeyAscii) <> "2" And Chr(KeyAscii) <> "3" Then
        If KeyAscii <> 13 And KeyAscii <> 8 Then
        KeyAscii = 0
        ElseIf KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        Exit Sub
        End If
    End If
    
    If KeyAscii > 0 And KeyAscii <> 8 Then
        If Chr(KeyAscii) = "1" Then
        LblPrioridade(0).FontBold = True
        LblPrioridade(1).FontBold = False
        LblPrioridade(2).FontBold = False
        TxtPrioridade.Text = "1"
        ElseIf Chr(KeyAscii) = "2" Then
        LblPrioridade(0).FontBold = False
        LblPrioridade(1).FontBold = True
        LblPrioridade(2).FontBold = False
        TxtPrioridade.Text = "2"
        ElseIf Chr(KeyAscii) = "3" Then
        LblPrioridade(0).FontBold = False
        LblPrioridade(1).FontBold = False
        LblPrioridade(2).FontBold = True
        TxtPrioridade.Text = "3"
        End If
    ElseIf KeyAscii = 0 Or KeyAscii = 8 Then
    LblPrioridade(1).FontBold = False
    LblPrioridade(0).FontBold = False
        If KeyAscii = 8 Then
        TxtPrioridade.Text = ""
        End If
    End If
KeyAscii = 0
DoEvents
End Sub


Public Sub ImprimeColeta(XColeta As String)
    Dim xVIAs As Integer
    Dim xLin As Integer
    Dim ximpr_cfg As String, ximpr_inst As Printer
    Dim xlinha As String
    Dim X As Double
    Dim Y As Double
    
    
    'busca impressora para este documento
    If Dir(App.Path & "\coletaimp.cfg") <> "" Then
        
        Open App.Path & "\coletaimp.cfg" For Input As #1
        
        Do Until EOF(1)
            Line Input #1, xlinha
            If Mid$(xlinha, 1, 3) = "COL" Then
                ximpr_cfg = Trim$(Mid$(xlinha, 5))
                Exit Do
            End If
        Loop
        
        Close #1
        
    Else
    
        MsgBox "Não está Configurado a Impressora para Este Documento: Coleta " & transctc(TxtFilial.Text, TxtColeta.Text)
        Exit Sub
        
    End If
    
    'seta impressora
    
    For Each ximpr_inst In Printers
        If ximpr_inst.DeviceName = ximpr_cfg Then
            Set Printer = ximpr_inst
            DoEvents
            Exit For
        End If
    Next
    
    'BUSCA A MINUTA A SER IMPRESSAO
    
    If de_informa.rsColetaSel.State = 1 Then de_informa.rsColetaSel.Close
    de_informa.ColetaSel XColeta
    
    If de_informa.rsColetaSel.RecordCount < 1 Then
        MsgBox "Coleta Para Impressão Inexistente!"
        Exit Sub
    End If
    
Dim xIMPColeta As String
Dim xIMPDataEmissao As String
Dim xIMPHoraEmissao As String
Dim xIMPEmissor As String
Dim xIMPSolEmpNome As String
Dim xIMPSolContato As String
Dim xIMPSolTel As String
Dim xIMPNomeRem As String
Dim xIMPEndRem As String
Dim xIMPCidadeUfRem As String
Dim xIMPNomeDest As String
Dim xIMPEndDest As String
Dim xIMPCidadeUfDest As String
Dim xIMPLocalColeta As String
Dim xIMPDataColeta As String
Dim xIMPHoraColeta As String
Dim xIMPVolumes As String
Dim xIMPPeso As String
Dim xIMPEspecie As String
Dim xIMPNatureza As String
Dim xIMPTipoFrete As String

Dim xIMPVolumesT As String
Dim xIMPPesoT As String
Dim xIMPEspecieT As String
Dim xIMPNaturezaT As String
Dim xIMPTipoFreteT As String


Dim xIMPNfs As String
Dim xIMPPrioridade As String
Dim xIMPOBServacoes As String
Dim xAux As String

Dim xIMPNfs1 As String
Dim xIMPNfs2 As String
Dim xIMPNfs3 As String
Dim xIMPNfs4 As String
Dim xIMPNfs5 As String

Dim xIMPOBServacoes1 As String
Dim xIMPOBServacoes2 As String
Dim xIMPOBServacoes3 As String
Dim xIMPOBServacoes4 As String
Dim xIMPOBServacoes5 As String

xAux = " "

xIMPColeta = Mid(XColeta, 1, 2) & "-" & Mid(XColeta, 3)
xIMPDataEmissao = de_informa.rsColetaSel.Fields("dataemissao")
xIMPHoraEmissao = de_informa.rsColetaSel.Fields("horaemissao")
xIMPEmissor = de_informa.rsColetaSel.Fields("emissor")
xIMPSolEmpNome = de_informa.rsColetaSel.Fields("SOLempnome")
xIMPSolContato = de_informa.rsColetaSel.Fields("SOLnome")
xIMPSolTel = de_informa.rsColetaSel.Fields("SOLtel")
xIMPNomeRem = de_informa.rsColetaSel.Fields("remetnome")
xIMPEndRem = de_informa.rsColetaSel.Fields("remetend")
xIMPCidadeUfRem = de_informa.rsColetaSel.Fields("remetcidade") & " - " & de_informa.rsColetaSel.Fields("remetuf")
xIMPNomeDest = de_informa.rsColetaSel.Fields("destnome")
xIMPEndDest = de_informa.rsColetaSel.Fields("destend")
xIMPCidadeUfDest = de_informa.rsColetaSel.Fields("destcidade") & " - " & de_informa.rsColetaSel.Fields("destuf")
xIMPLocalColeta = "LOCAL DE COLETA: " & de_informa.rsColetaSel.Fields("coletaend") & " - " & de_informa.rsColetaSel.Fields("coletacidade") & " - " & de_informa.rsColetaSel.Fields("coletauf")
xIMPDataColeta = "DATA A COLETAR: " & de_informa.rsColetaSel.Fields("datacoleta")
xIMPHoraColeta = "HORARIO: " & de_informa.rsColetaSel.Fields("horacoleta")

xIMPVolumesT = "VOLs."
xIMPPesoT = "PESO"
xIMPEspecieT = "ESPECIE"
xIMPNaturezaT = "NATUREZA"
xIMPTipoFreteT = "FRETE"


xIMPVolumes = de_informa.rsColetaSel.Fields("volumes")
xIMPPeso = de_informa.rsColetaSel.Fields("peso")
xIMPEspecie = de_informa.rsColetaSel.Fields("especie")
xIMPNatureza = de_informa.rsColetaSel.Fields("natureza")
xIMPTipoFrete = IIf(de_informa.rsColetaSel.Fields("tipofrete") = "1", "PAGO", "A PAGAR")
xIMPNfs = de_informa.rsColetaSel.Fields("NFS")
xIMPPrioridade = de_informa.rsColetaSel.Fields("prioridade")
xIMPOBServacoes = de_informa.rsColetaSel.Fields("OBS")

xIMPColeta = Mid(xIMPColeta, 1, 12) & String(12 - Len(Mid(xIMPColeta, 1, 12)), xAux)
xIMPDataEmissao = Mid(xIMPDataEmissao, 1, 11) & String(11 - Len(Mid(xIMPDataEmissao, 1, 11)), xAux)
xIMPHoraEmissao = Mid(xIMPHoraEmissao, 1, 11) & String(11 - Len(Mid(xIMPHoraEmissao, 1, 11)), xAux)
xIMPEmissor = Mid(xIMPEmissor, 1, 11) & String(11 - Len(Mid(xIMPEmissor, 1, 11)), xAux)

xIMPSolEmpNome = Mid(xIMPSolEmpNome, 1, 78) & String(78 - Len(Mid(xIMPSolEmpNome, 1, 78)), xAux)
xIMPSolContato = Mid(xIMPSolContato, 1, 78) & String(78 - Len(Mid(xIMPSolContato, 1, 78)), xAux)
xIMPSolTel = Mid(xIMPSolTel, 1, 78) & String(78 - Len(Mid(xIMPSolTel, 1, 78)), xAux)
xIMPNomeRem = Mid(xIMPNomeRem, 1, 78) & String(78 - Len(Mid(xIMPNomeRem, 1, 78)), xAux)
xIMPEndRem = Mid(xIMPEndRem, 1, 78) & String(78 - Len(Mid(xIMPEndRem, 1, 78)), xAux)
xIMPCidadeUfRem = Mid(xIMPCidadeUfRem, 1, 78) & String(78 - Len(Mid(xIMPCidadeUfRem, 1, 78)), xAux)
xIMPNomeDest = Mid(xIMPNomeDest, 1, 78) & String(78 - Len(Mid(xIMPNomeDest, 1, 78)), xAux)
xIMPEndDest = Mid(xIMPEndDest, 1, 78) & String(78 - Len(Mid(xIMPEndDest, 1, 78)), xAux)
xIMPCidadeUfDest = Mid(xIMPCidadeUfDest, 1, 78) & String(78 - Len(Mid(xIMPCidadeUfDest, 1, 78)), xAux)
'xIMPLocalColeta = Mid(xIMPLocalColeta, 1, 92) & String(92 - Len(Mid(xIMPLocalColeta, 1, 92)), xAux)

xIMPLocalColeta = String((92 - Len(Mid(xIMPLocalColeta, 1, 92))) / 2, xAux) & Mid(xIMPLocalColeta, 1, 92) & String((92 - Len(Mid(xIMPLocalColeta, 1, 92))) / 2, xAux)
xIMPDataColeta = String((45 - Len(Mid(xIMPDataColeta, 1, 45))) / 2, xAux) & Mid(xIMPDataColeta, 1, 45) & String((45 - Len(Mid(xIMPDataColeta, 1, 45))) / 2, xAux)
xIMPHoraColeta = String((44 - Len(Mid(xIMPHoraColeta, 1, 44))) / 2, xAux) & Mid(xIMPHoraColeta, 1, 44) & String((44 - Len(Mid(xIMPHoraColeta, 1, 44))) / 2, xAux)
xIMPVolumes = String(15 - Len(Mid(xIMPVolumes, 1, 15)), xAux) & Mid(xIMPVolumes, 1, 15)
xIMPPeso = String(12 - Len(Mid(xIMPPeso, 1, 12)), xAux) & Mid(xIMPPeso, 1, 12)
xIMPEspecie = Mid(xIMPEspecie, 1, 22) & String(22 - Len(Mid(xIMPEspecie, 1, 22)), xAux)
xIMPNatureza = Mid(xIMPNatureza, 1, 22) & String(22 - Len(Mid(xIMPNatureza, 1, 22)), xAux)
xIMPTipoFrete = Mid(xIMPTipoFrete, 1, 9) & String(9 - Len(Mid(xIMPTipoFrete, 1, 9)), xAux)

xIMPVolumesT = String((15 - Len(Mid(xIMPVolumesT, 1, 15))) / 2, xAux) & Mid(xIMPVolumesT, 1, 15) & String((15 - Len(Mid(xIMPVolumesT, 1, 15))) / 2, xAux)
xIMPPesoT = String((12 - Len(Mid(xIMPPesoT, 1, 12))) / 2, xAux) & Mid(xIMPPesoT, 1, 12) & String((12 - Len(Mid(xIMPPesoT, 1, 12))) / 2, xAux)
xIMPEspecieT = String((22 - Len(Mid(xIMPEspecieT, 1, 22))) / 2, xAux) & Mid(xIMPEspecieT, 1, 22) & String((22 - Len(Mid(xIMPEspecieT, 1, 22))) / 2, xAux)
xIMPNaturezaT = String((22 - Len(Mid(xIMPNaturezaT, 1, 22))) / 2, xAux) & Mid(xIMPNaturezaT, 1, 22) & String((22 - Len(Mid(xIMPNaturezaT, 1, 22))) / 2, xAux)
xIMPTipoFreteT = String((9 - Len(Mid(xIMPTipoFreteT, 1, 9))) / 2, xAux) & Mid(xIMPTipoFreteT, 1, 9) & String((9 - Len(Mid(xIMPTipoFreteT, 1, 9))) / 2, xAux)


xIMPNfs = Mid(xIMPNfs, 1, 310) & String(310 - Len(Mid(xIMPNfs, 1, 310)), xAux)
xIMPPrioridade = Mid(xIMPPrioridade, 1, 10) & String(10 - Len(Mid(xIMPPrioridade, 1, 10)), xAux)
xIMPOBServacoes = Mid(xIMPOBServacoes, 1, 310) & String(310 - Len(Mid(xIMPOBServacoes, 1, 310)), xAux)
xIMPNfs1 = Mid(xIMPNfs, 1, 58)
xIMPNfs2 = Mid(xIMPNfs, 59, 63)
xIMPNfs3 = Mid(xIMPNfs, 122, 63)
xIMPNfs4 = Mid(xIMPNfs, 185, 63)
xIMPNfs5 = Mid(xIMPNfs, 248, 63)
xIMPOBServacoes1 = Mid(xIMPOBServacoes, 1, 58)
xIMPOBServacoes2 = Mid(xIMPOBServacoes, 59, 63)
xIMPOBServacoes3 = Mid(xIMPOBServacoes, 122, 63)
xIMPOBServacoes4 = Mid(xIMPOBServacoes, 185, 63)
xIMPOBServacoes5 = Mid(xIMPOBServacoes, 248, 63)

frmLogo.Text1.Text = XColeta
    
    For xVIAs = 1 To 1  'DUAS VIAS
    
        If xVIAs = 1 Then
        xLin = 0
        Printer.DrawStyle = 0
        Printer.ForeColor = &H80000008  'PRETO
        'Printer.DrawWidth = 8
        Printer.DrawMode = 9
        ElseIf xVIAs = 2 Then
        xLin = 149
        Printer.DrawStyle = 0
        Printer.ForeColor = &H80000008  'PRETO
        'Printer.DrawWidth = 8
        Printer.DrawMode = 9
        End If
    Printer.CurrentX = 0
    Printer.CurrentY = 0 + xLin
    
    Printer.FontName = "Courier New"
    Printer.FontSize = 3
    Printer.Print
    Printer.FontSize = 6
    Printer.Print Spc(24); "INTEC-Integração Nacional de Transportes de Encom. e Cargas Ltda"
    Printer.FontSize = 1
    Printer.Print ""
    Printer.FontSize = 6
    Printer.Print Spc(24); "AV. MARG. DIREITA DO RIO TIETÊ, 504 - BARUERI/SP - CEP 06455-050"
    Printer.FontSize = 1
    Printer.Print ""
    Printer.FontSize = 6
    Printer.Print Spc(24); "CNPJ: 52.134.798-0001-68         INSCR.ESTADUAL: 206.182.910.118"
    Printer.FontSize = 1
    Printer.Print ""
    Printer.FontSize = 6
    Printer.Print Spc(24); "TELEFONES: (11) 4689-7575 / 4193-5921           www.intec.com.br"
            
    'GRAFICOS
    Printer.ForeColor = &H80000008   'PRETO
    Printer.ScaleMode = vbMillimeters
    Printer.Line (0, 0 + xLin)-(198, 15 + xLin), , B
    Printer.ForeColor = &H80000008   'PRETO
    Printer.Line (113, 0 + xLin)-(198, 15 + xLin), , B 'QUADRO MINUTA DE TRANSP, NUMERO DA MINUTA DATA, EMISSOR, ETC
    Printer.FontName = "ARIAL"
    Printer.FontSize = 30
    Printer.FontBold = True
    Printer.CurrentX = 130
    Printer.CurrentY = 2 + xLin
    Printer.Print "COLETA"
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.ForeColor = &HC0C0C0     'CINZA
    Printer.Line (0, 15 + xLin)-(198, 20 + xLin), , BF 'QUADRO MINUTA DE TRANSP, NUMERO DA MINUTA DATA, EMISSOR, ETC
    Printer.ForeColor = &H80000008   'PRETO
    Printer.Line (0, 15 + xLin)-(198, 20 + xLin), , B 'QUADRO MINUTA DE TRANSP, NUMERO DA MINUTA DATA, EMISSOR, ETC
    Printer.CurrentX = 0
    Printer.CurrentY = 15.5 + xLin
    Printer.FontBold = True
    Printer.Print " Numero: " & xIMPColeta & "   DATA: " & xIMPDataEmissao & "   HORA: " & xIMPHoraEmissao & "   EMISSOR: " & xIMPEmissor & "   " & Trim(Str(xVIAs)) & "ª VIA"
    Printer.FontBold = False
    
    
    Printer.PaintPicture frmLogo.piclogo.Picture, 1, 1 + xLin, frmLogo.piclogo.Picture.Width * 0.0013, frmLogo.piclogo.Picture.Height * 0.0013
    Printer.PaintPicture frmLogo.Picture1, 140, 113 + xLin, frmLogo.Picture1.Picture.Width * 0.0068, frmLogo.Picture1.Picture.Height * 0.008
    
    
    Printer.Line (0, 20 + xLin)-(198, 33 + xLin), , B
    Printer.Line (0, 33 + xLin)-(198, 49 + xLin), , B
    Printer.Line (0, 49 + xLin)-(198, 63 + xLin), , B
    Printer.ForeColor = &HC0C0C0     'CINZA
    Printer.Line (0, 63 + xLin)-(198, 66.5 + xLin), , BF
    Printer.ForeColor = &H80000008   'PRETO
    Printer.Line (0, 63 + xLin)-(198, 66.5 + xLin), , B
    Printer.ForeColor = &HC0C0C0     'CINZA
    Printer.Line (0, 66.5 + xLin)-(198, 70.5 + xLin), , BF
    Printer.ForeColor = &H80000008   'PRETO
    Printer.Line (0, 66.5 + xLin)-(198, 70.5 + xLin), , B
    Printer.Line (0, 70.5 + xLin)-(37, 74.5 + xLin), , B
    Printer.Line (37, 70.5 + xLin)-(69, 74.5 + xLin), , B
    Printer.Line (69, 70.5 + xLin)-(121, 74.5 + xLin), , B
    Printer.Line (121, 70.5 + xLin)-(175, 74.5 + xLin), , B
    Printer.Line (175, 70.5 + xLin)-(198, 74.5 + xLin), , B
    
    Printer.Line (0, 74.5 + xLin)-(37, 78.5 + xLin), , B
    Printer.Line (37, 74.5 + xLin)-(69, 78.5 + xLin), , B
    Printer.Line (69, 74.5 + xLin)-(121, 78.5 + xLin), , B
    Printer.Line (121, 74.5 + xLin)-(175, 78.5 + xLin), , B
    Printer.Line (175, 74.5 + xLin)-(198, 78.5 + xLin), , B
    
    Printer.Line (0, 78.5 + xLin)-(138, 97 + xLin), , B
        If UCase(Trim(xIMPPrioridade)) = "URGENTE" Then
        Printer.ForeColor = &HC0C0C0     'CINZA
        Printer.Line (138, 78.5 + xLin)-(198, 97 + xLin), , BF
        Printer.ForeColor = &H80000008   'PRETO
        End If
    Printer.Line (138, 78.5 + xLin)-(198, 97 + xLin), , B
    
    Printer.Line (0, 97 + xLin)-(138, 116 + xLin), , B
    Printer.Line (138, 97 + xLin)-(198, 138 + xLin), , B
    
    Printer.ForeColor = &HC0C0C0     'CINZA
    Printer.Line (0, 116 + xLin)-(138, 120 + xLin), , BF
    Printer.ForeColor = &H80000008   'PRETO
    Printer.Line (0, 116 + xLin)-(138, 120 + xLin), , B
    
    Printer.Line (0, 120 + xLin)-(138, 138 + xLin), , B
    
    Printer.CurrentX = 0
    Printer.CurrentY = 20.5 + xLin
    
    Printer.Print " SOLICITANTE : "; xIMPSolEmpNome
    Printer.Print " CONTATO     : "; xIMPSolContato
    Printer.Print " TELEFONE    : "; xIMPSolTel
    Printer.Print ""
    Printer.Print " REMETENTE   : "; xIMPNomeRem
    Printer.Print " ENDERECO    : "; xIMPEndRem
    Printer.Print " CIDADE-UF   : "; xIMPCidadeUfRem
    Printer.Print ""
    Printer.Print " DESTINATARIO: "; xIMPNomeDest
    Printer.Print " ENDERECO    : "; xIMPEndDest
    Printer.Print " CIDADE-UF   : "; xIMPCidadeUfDest
    Printer.FontSize = 2
    Printer.Print ""
    Printer.FontSize = 10
    Printer.FontBold = True
    'Printer.Print " LOCAL DE COLETA: "
    Printer.Print " "; xIMPLocalColeta
    Printer.Print " "; xIMPDataColeta & "   " & xIMPHoraColeta
    Printer.Print " "; xIMPVolumesT; "   "; xIMPPesoT; "   "; xIMPEspecieT; "  "; xIMPNaturezaT; "   "; xIMPTipoFreteT
    Printer.FontBold = False
    Printer.Print " "; xIMPVolumes; "   "; xIMPPeso; "   "; xIMPEspecie; "   "; xIMPNatureza; "   "; xIMPTipoFrete
    Printer.Print " "; "NFs: " & xIMPNfs1 '& "  Prioridade:"
    Printer.Print " "; xIMPNfs2
    Printer.Print " "; xIMPNfs3
    Printer.Print " "; xIMPNfs4
    Printer.Print " "; xIMPNfs5
    Printer.Print " "; "OBS: " & xIMPOBServacoes1
    Printer.Print " "; xIMPOBServacoes2
    Printer.Print " "; xIMPOBServacoes3
    Printer.Print " "; xIMPOBServacoes4
    Printer.Print " "; xIMPOBServacoes5
    Printer.FontBold = True
    Printer.Print "                          R E C E B I M E N T O "
    Printer.FontSize = 8
    Printer.Print " NOME:"
    Printer.Print ""
    Printer.Print " Nº RG:"
    Printer.Print "                                                          ____________________"
    Printer.Print " DATA/HORA:                                                    ASSINATURA"
    
    X = Printer.CurrentX
    Y = Printer.CurrentY
    Printer.FontName = "ARIAL"
    Printer.FontSize = 25
    Printer.FontBold = True
    Printer.CurrentX = 140
    Printer.CurrentY = 83 + xLin
    Printer.Print xIMPPrioridade
    Printer.FontName = "COURIER NEW"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentX = X
    Printer.CurrentY = Y
    
    Next
    
    Printer.EndDoc
    
End Sub


