VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIncluiTabAerea 
   Caption         =   "Cadastro de Nova Tabela"
   ClientHeight    =   6600
   ClientLeft      =   3435
   ClientTop       =   1185
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   5280
   Begin VB.TextBox txtDescrTabela 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1680
      TabIndex        =   39
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtTaxaMinima 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "?"
      Height          =   285
      Left            =   2400
      TabIndex        =   36
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdBuscaLocal 
      Caption         =   "?"
      Height          =   285
      Left            =   2400
      TabIndex        =   22
      Top             =   1320
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cadastros das Tabelas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   5055
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   3720
         TabIndex        =   23
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "Gravar"
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tabela TE TC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   4815
         Begin VB.CheckBox chkUsarGeral 
            Caption         =   "Usar Tab. Geral"
            Height          =   195
            Left            =   3120
            TabIndex        =   34
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtValorKG 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   12
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdBuscaTETC 
            Caption         =   "?"
            Height          =   285
            Left            =   2280
            TabIndex        =   31
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtCodTETC 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblDescTETC 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   525
            Left            =   1200
            TabIndex        =   14
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label lblPesominimo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3600
            TabIndex        =   15
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Peso Mínimo:"
            Height          =   195
            Left            =   2520
            TabIndex        =   35
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Valor Por Kg:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   1320
            Width           =   930
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cód. TE TC:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Caption         =   "Tabela Geral"
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
         TabIndex        =   20
         Top             =   360
         Width           =   4815
         Begin VB.TextBox txtAcima1000 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3600
            TabIndex        =   10
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtAte1000 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3600
            TabIndex        =   9
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtAte500 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3600
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtAte300 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   7
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtAte50 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   6
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtAte25 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Até 1.000Kgs:"
            Height          =   195
            Left            =   2400
            TabIndex        =   29
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Até 300Kgs:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   870
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Até 50Kgs:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   780
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Até 500Kgs:"
            Height          =   195
            Left            =   2400
            TabIndex        =   26
            Top             =   360
            Width           =   870
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Acima 1.000Kgs:"
            Height          =   195
            Left            =   2400
            TabIndex        =   25
            Top             =   1080
            Width           =   1200
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Até 25Kgs:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   780
         End
      End
   End
   Begin VB.TextBox txtSiglaLocal 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin MSMask.MaskEdBox mskVigenciaInicio 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd.M.yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCodCia 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblLocalidade 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   38
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblFantasia 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   37
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Taxa Mínima:"
      Height          =   195
      Left            =   240
      TabIndex        =   32
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Localidade:"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Vigência Em:"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   960
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descr. da Tabela:"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cod.Cia Aérea:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1065
   End
End
Attribute VB_Name = "frmIncluiTabAerea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label16_Click()

End Sub

Private Sub txtCia_Change()

End Sub

Private Sub txtCia_LinkNotify()

End Sub

Private Sub txtCia_LostFocus()


End Sub

Private Sub cmdGravar_Click()
    
    'Gravar a tabela (pai) tb_airtabpreco
    xdata = Date & " " & Time
    de_informa.Ins_TabPreco txtCodCia, txtDescrTabela, "INATIVA", mskVigenciaInicio.Text, CDate(xdata)
    If de_informa.rsSel_IDTabPrecoData.State = 1 Then de_informa.rsSel_IDTabPrecoData.Close
    de_informa.Sel_IDTabPrecoData xdata
    xidtabela = de_informa.rsSel_IDTabPrecoData.Fields("idtabela")
    
    'Gravar a tabela geral (filho) tb_airtabprecogeral
    de_informa.Ins_TabPrecogeral xidtabela, txtSiglaLocal, CDbl(txtTaxaMinima), CDbl(txtAte25), CDbl(txtAte50), CDbl(txtAte300), CDbl(txtAte500), CDbl(txtAte1000), CDbl(txtAcima1000)
    
    'Gravar a tabela TETC (filho) tb_airtabprecotetc
    If chkUsarGeral = 1 Then
        xusargeral = "S"
    Else
        xusargeral = "N"
    End If
    de_informa.Ins_TabPrecotetc xidtabela, txtSiglaLocal, txtCodTETC, CDbl(txtTaxaMinima), CDbl(txtValorKG), xusargeral
    
    MsgBox "OK ! Registro Gravado."
    
    cmdGravar.Enabled = False
    cmdSair.SetFocus
    
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub txtAcima1000_GotFocus()
    If Val(txtAcima1000) > 0 Then txtAcima1000 = CDbl(txtAcima1000)
    txtAcima1000.SelStart = 0
    txtAcima1000.SelLength = Len(Trim$(txtAcima1000))
End Sub

Private Sub txtAcima1000_LostFocus()
    If Val(txtAcima1000) > 0 Then txtAcima1000 = Format(txtAcima1000, "###,##0.00")
End Sub

Private Sub txtAte1000_GotFocus()
    If Val(txtAte1000) > 0 Then txtAte1000 = CDbl(txtAte1000)
    txtAte1000.SelStart = 0
    txtAte1000.SelLength = Len(Trim$(txtAte1000))
End Sub

Private Sub txtAte1000_LostFocus()
    If Val(txtAte1000) > 0 Then txtAte1000 = Format(txtAte1000, "###,##0.00")
End Sub

Private Sub txtAte25_GotFocus()
    If Val(txtAte25) > 0 Then txtAte25 = CDbl(txtAte25)
    txtAte25.SelStart = 0
    txtAte25.SelLength = Len(Trim$(txtAte25))
End Sub

Private Sub txtAte25_LostFocus()
    If Val(txtAte25) > 0 Then txtAte25 = Format(txtAte25, "###,##0.00")
End Sub

Private Sub txtAte300_GotFocus()
    If Val(txtAte300) > 0 Then txtAte300 = CDbl(txtAte300)
    txtAte300.SelStart = 0
    txtAte300.SelLength = Len(Trim$(txtAte300))
End Sub

Private Sub txtAte300_LostFocus()
    If Val(txtAte300) > 0 Then txtAte300 = Format(txtAte300, "###,##0.00")
End Sub

Private Sub txtAte50_GotFocus()
    If Val(txtAte50) > 0 Then txtAte50 = CDbl(txtAte50)
    txtAte50.SelStart = 0
    txtAte50.SelLength = Len(Trim$(txtAte50))
End Sub

Private Sub txtAte50_LostFocus()
    If Val(txtAte50) > 0 Then txtAte50 = Format(txtAte50, "###,##0.00")
End Sub

Private Sub txtAte500_GotFocus()
    If Val(txtAte500) > 0 Then txtAte500 = CDbl(txtAte500)
    txtAte500.SelStart = 0
    txtAte500.SelLength = Len(Trim$(txtAte500))
End Sub

Private Sub txtAte500_LostFocus()
    If Val(txtAte500) > 0 Then txtAte500 = Format(txtAte500, "###,##0.00")
End Sub

Private Sub txtCodCia_LostFocus()
    If Len(Trim$(txtCodCia)) > 0 Then
        txtCodCia = UCase(Trim$(txtCodCia))
        lblFantasia = ""
        If de_informa.rsSel_CiaAereaPorCodigo.State = 1 Then de_informa.rsSel_CiaAereaPorCodigo.Close
        de_informa.Sel_CiaAereaPorCodigo Trim$(txtCodCia.Text)
        If de_informa.rsSel_CiaAereaPorCodigo.RecordCount > 0 Then
            lblFantasia = de_informa.rsSel_CiaAereaPorCodigo.Fields("fantasia")
        Else
            MsgBox "Código de Cia. Aérea Não Encontrado !"
            txtCodCia.SetFocus
            Exit Sub
        End If
    Else
        lblFantasia = ""
    End If
End Sub

Private Sub txtCodTETC_LostFocus()
    If Len(Trim$(txtCodTETC)) > 0 Then
        txtCodTETC = UCase(Trim$(txtCodTETC))
        lblDescTETC = ""
        If de_informa.rsSel_CadTETC.State = 1 Then de_informa.rsSel_CadTETC.Close
        de_informa.Sel_CadTETC Trim$(txtCodTETC.Text)
        If de_informa.rsSel_CadTETC.RecordCount > 0 Then
            lblDescTETC = de_informa.rsSel_CadTETC.Fields("descricao")
        Else
            MsgBox "Código de Categoria de Tarifa TE TC Não Encontrado !"
            txtCodTETC.SetFocus
            Exit Sub
        End If
    Else
        lblDescTETC = ""
    End If
End Sub

Private Sub txtDescrTabela_Change()
Call TextMoneyBox_Change(txtDescrTabela)
End Sub

Private Sub txtDescrTabela_GotFocus()
Call TextMoneyBox_GotFocus(txtDescrTabela)
End Sub

Private Sub txtDescrTabela_KeyPress(KeyAscii As Integer)
Call TextMoneyBox_KeyPress(KeyAscii)
End Sub

Private Sub txtSiglaLocal_LostFocus()
    If Len(Trim$(txtSiglaLocal)) > 0 Then
        txtSiglaLocal = UCase(Trim$(txtSiglaLocal))
        lblLocalidade = ""
        If de_informa.rsSel_CadLocalAir.State = 1 Then de_informa.rsSel_CadLocalAir.Close
        de_informa.Sel_CadLocalAir Trim$(txtSiglaLocal.Text)
        If de_informa.rsSel_CadLocalAir.RecordCount > 0 Then
            lblLocalidade = de_informa.rsSel_CadLocalAir.Fields("localidade")
        Else
            MsgBox "Código de Localidade Destino Não Encontrado !"
            txtSiglaLocal.SetFocus
            Exit Sub
        End If
    Else
        lblLocalidade = ""
    End If

End Sub

Private Sub txtTaxaMinima_GotFocus()
    If Val(txtTaxaMinima) > 0 Then txtTaxaMinima = CDbl(txtTaxaMinima)
    txtTaxaMinima.SelStart = 0
    txtTaxaMinima.SelLength = Len(Trim$(txtTaxaMinima))
End Sub

Private Sub txtTaxaMinima_LostFocus()
    If Val(txtTaxaMinima) > 0 Then txtTaxaMinima = Format(txtTaxaMinima, "###,##0.00")
End Sub

Private Sub txtValorKG_GotFocus()
    If Val(txtValorKG) > 0 Then txtValorKG = CDbl(txtValorKG)
    txtValorKG.SelStart = 0
    txtValorKG.SelLength = Len(Trim$(txtValorKG))
End Sub

Private Sub txtValorKG_LostFocus()
    If Val(txtValorKG) > 0 Then
        txtValorKG = Format(txtValorKG, "###,##0.00")
        If txtTaxaMinima > 0 Then
            lblPesominimo = Format(CDbl(txtTaxaMinima) / (txtValorKG), "##,##0.0")
        Else
            lblPesominimo = ""
        End If
    Else
        lblPesominimo = ""
    End If
End Sub
