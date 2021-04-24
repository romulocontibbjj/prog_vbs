VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRastrPrazo 
   Caption         =   "Rastrear Cálculo de Prazo de Entrega"
   ClientHeight    =   6480
   ClientLeft      =   1185
   ClientTop       =   1245
   ClientWidth     =   9645
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9645
   Begin VB.Frame Frame3 
      Caption         =   "Comandos"
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
      Left            =   8280
      TabIndex        =   23
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton cmdImprTela 
         Height          =   615
         Left            =   240
         Picture         =   "frmRastrPrazo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do CTC / Entrega"
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
      TabIndex        =   1
      Top             =   120
      Width           =   5535
      Begin VB.Frame Frame2 
         Caption         =   "Prazos"
         Height          =   1095
         Left            =   3120
         TabIndex        =   18
         Top             =   1200
         Width           =   2295
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Dias Úteis:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   720
            Width           =   765
         End
         Begin VB.Label lblPrazo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   21
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Meta:"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   405
         End
         Begin VB.Label lblMeta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1080
            TabIndex        =   19
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label lblHsEntr 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblHsEmiss 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Hs:"
         Height          =   195
         Left            =   1920
         TabIndex        =   15
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Hs:"
         Height          =   195
         Left            =   1920
         TabIndex        =   14
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label lblEntrega 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblEmissao 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblModal 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3960
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Modalidade:"
         Height          =   195
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   870
      End
      Begin VB.Label lblUfDest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5040
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Entrega:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label lblCidadeDest 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Emissão:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cidade Dest:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "UF Dest:"
         Height          =   195
         Left            =   4320
         TabIndex        =   4
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblFilialCTC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filial CTC:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   705
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FlexRastrEntrega 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      Appearance      =   0
   End
   Begin MSComCtl2.MonthView Calend 
      Height          =   2310
      Left            =   5760
      TabIndex        =   25
      Top             =   240
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      ShowToday       =   0   'False
      StartOfWeek     =   90832897
      CurrentDate     =   37305
   End
End
Attribute VB_Name = "frmRastrPrazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprTela_Click()
    If Printer.Orientation = vbPRORPortrait Then Printer.Orientation = vbPRORLandscape
    Me.PrintForm
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    FlexRastrEntrega.ColWidth(0) = 1000
    FlexRastrEntrega.ColWidth(1) = 3650
    FlexRastrEntrega.ColWidth(2) = 1200
    FlexRastrEntrega.ColWidth(3) = 2600
    FlexRastrEntrega.ColWidth(4) = 650
    Calend.Value = CDate(lblEmissao)
    FlexRastrEntrega.Row = 0
    FlexRastrEntrega.Col = 0
    FlexRastrEntrega.Text = "Data"
    FlexRastrEntrega.Col = 1
    FlexRastrEntrega.Text = "Ocorrência"
    FlexRastrEntrega.Col = 2
    FlexRastrEntrega.Text = "Dia"
    FlexRastrEntrega.Col = 3
    FlexRastrEntrega.Text = "Obs."
    FlexRastrEntrega.Col = 4
    FlexRastrEntrega.Text = "Dia Útil"
    xuf = lblUFdest
    DoEvents
    
    Dim xy As Integer, xdias As Integer, xlinflex As Integer, xdiasuteis As Integer, xemissNaoUtil As String
    Dim xdutilSN As String
    
    xy = CDate(lblEntrega) - CDate(lblEmissao)
'    If Weekday(CDate(lblEmissao)) = 1 Or Weekday(CDate(lblEmissao)) = 7 Then
'        xemissNaoUtil = "S"
'    End If
    xlinflex = 0
    xdiasuteis = 0
    FlexRastrEntrega.Rows = xy + 2
    For xdias = 0 To xy
        xdutilSN = "S"
        xlinflex = xlinflex + 1
        FlexRastrEntrega.Row = xlinflex
        FlexRastrEntrega.Col = 0
        FlexRastrEntrega.Text = CDate(lblEmissao.Caption) + xdias  'DATA
        FlexRastrEntrega.Col = 2
        FlexRastrEntrega.Text = diasemana(CDate(lblEmissao) + xdias)  'DIA SEMANA
        FlexRastrEntrega.Col = 3
        'VERIFICA SE É FERIADO

        If de_informa.rsSel_Feriado.State = 1 Then de_informa.rsSel_Feriado.Close
        de_informa.Sel_Feriado Month(CDate(lblEmissao) + xdias), Day(CDate(lblEmissao) + xdias)
        If de_informa.rsSel_Feriado.RecordCount > 0 Then
            de_informa.rsSel_Feriado.MoveFirst
            Do Until de_informa.rsSel_Feriado.EOF
                If de_informa.rsSel_Feriado.Fields("uf") = "BR" Then 'feriado nacional
                    If de_informa.rsSel_Feriado.Fields("tipo") = "V" Then  'feriado variável
                        If Year(CDate(lblEmissao) + xdias) = de_informa.rsSel_Feriado.Fields("ano") Then 'verif. se bate o ano, pois é feriado variável
                            FlexRastrEntrega.Col = 1
                            FlexRastrEntrega.Text = "*  F E R I A D O"
                            FlexRastrEntrega.Col = 3
                            xdutilSN = "N"
                            FlexRastrEntrega.Text = de_informa.rsSel_Feriado.Fields("descricao")
                        End If
                    Else 'feriado fixo, nao verif. o ano pois todo ano é a mesma data
                        FlexRastrEntrega.Col = 1
                        FlexRastrEntrega.Text = "*  F E R I A D O"
                        FlexRastrEntrega.Col = 3
                        xdutilSN = "N"
                        FlexRastrEntrega.Text = de_informa.rsSel_Feriado.Fields("descricao")
                    End If
                ElseIf de_informa.rsSel_Feriado.Fields("uf") <> "BR" _
                And de_informa.rsSel_Feriado.Fields("cidade") = "" Then 'feriado estadual
                    If xuf = de_informa.rsSel_Feriado.Fields("uf") Then
                        FlexRastrEntrega.Col = 1
                        FlexRastrEntrega.Text = "*  F E R I A D O"
                        FlexRastrEntrega.Col = 3
                        xdutilSN = "N"
                        FlexRastrEntrega.Text = de_informa.rsSel_Feriado.Fields("descricao")
                    End If
                ElseIf de_informa.rsSel_Feriado.Fields("uf") <> "BR" _
                And de_informa.rsSel_Feriado.Fields("cidade") <> "" Then 'feriado local/municipal
                    If xuf = de_informa.rsSel_Feriado.Fields("uf") _
                    And xcidade = de_informa.rsSel_Feriado.Fields("cidade") Then
                        FlexRastrEntrega.Col = 1
                        FlexRastrEntrega.Text = "*  F E R I A D O"
                        FlexRastrEntrega.Col = 3
                        xdutilSN = "N"
                        FlexRastrEntrega.Text = de_informa.rsSel_Feriado.Fields("descricao")
                    End If
                End If
                de_informa.rsSel_Feriado.MoveNext
            Loop
        End If
        
        'DIA ÚTIL / FIM DE SEMANA / EMISSÃO / OCORRÊNCIA
        
        If Weekday(CDate(lblEmissao) + xdias) = 7 Or Weekday(CDate(lblEmissao) + xdias) = 1 Then
            FlexRastrEntrega.Text = "FINAL DE SEMANA"
            xdutilSN = "N"
        End If
        
        If CDate(lblEmissao) + xdias = CDate(lblEmissao) Then
            FlexRastrEntrega.Col = 1
            FlexRastrEntrega.Text = "EMISSÃO DO CTC  " & FlexRastrEntrega.Text
            FlexRastrEntrega.Col = 3
            If Len(Trim$(FlexRastrEntrega.Text)) < 3 Then
                FlexRastrEntrega.Text = "DIA DA EMISSÃO DO CTC"
            End If
        End If
        
        If xdutilSN = "S" Then
            FlexRastrEntrega.Col = 3
            If xdiasuteis = 0 Then
                FlexRastrEntrega.Text = "DIA ÚTIL (D-Zero) - " & FlexRastrEntrega.Text
            Else
                FlexRastrEntrega.Text = "DIA ÚTIL - " & FlexRastrEntrega.Text
            End If
            FlexRastrEntrega.Col = 4
            FlexRastrEntrega.Text = xdiasuteis
            xdiasuteis = xdiasuteis + 1
        Else
            FlexRastrEntrega.Col = 4
            FlexRastrEntrega.Text = "       -"
        End If
            
        'VERIFICA SE É ENTREGA
        
        If CDate(lblEmissao) + xdias = CDate(lblEntrega) Then
            FlexRastrEntrega.Col = 1
            If Me.Caption = "Rastrear Cálculo de Prazo de Entrega" Then
                FlexRastrEntrega.Text = FlexRastrEntrega.Text & " - ENTREGA REALIZADA"
                
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "CONSULTA", xusuario, "RASTREAR DIAS ÚTEIS DA ENTREGA REALIZADA. CTC: " & lblFilialctc
            
            Else
                FlexRastrEntrega.Text = FlexRastrEntrega.Text & " - PREVISÃO DE ENTREGA"
                
                'LOG DE USUÁRIO
                de_informa.ins_LogUsuario "CONSULTA", xusuario, "RASTREAR DIAS ÚTEIS DA PREVISÃO DE ENTREGA. CTC: " & lblFilialctc
                
            End If
            
            FlexRastrEntrega.Col = 4
            If xdutilSN = "N" And FlexRastrEntrega.Text = "       -" Then
                FlexRastrEntrega.Text = xdiasuteis
            End If
            
            
        End If
        
    Next xdias
    
        

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRastrPrazo = Nothing
End Sub
