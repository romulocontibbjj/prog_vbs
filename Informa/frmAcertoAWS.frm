VERSION 5.00
Begin VB.Form frmAcertoAWS 
   Caption         =   "Acerto Datas do AWS (Minutas)"
   ClientHeight    =   2880
   ClientLeft      =   2535
   ClientTop       =   2445
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   6585
   Begin VB.CommandButton Command2 
      Caption         =   "Acertar Data de Entrega"
      Height          =   735
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acertar Data de Emissao"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label lblCtcNaoOK 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   5760
      TabIndex        =   8
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CTCs NÃO OK:"
      Height          =   195
      Left            =   4200
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblCtcOk 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   5760
      TabIndex        =   6
      Top             =   1440
      Width           =   90
   End
   Begin VB.Label Label5 
      Caption         =   "CTCs OK:"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblMinLida 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   5760
      TabIndex        =   4
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Minutas Lidas:"
      Height          =   195
      Left            =   4200
      TabIndex        =   3
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label lbltotminutas 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   5760
      TabIndex        =   2
      Top             =   720
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total de Minutas:"
      Height          =   195
      Left            =   4200
      TabIndex        =   1
      Top             =   720
      Width           =   1230
   End
End
Attribute VB_Name = "frmAcertoAWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If de_informa.rsAcerto_SelBase.State = 1 Then de_informa.rsAcerto_SelBase.Close
    de_informa.Acerto_SelBase
    lbltotminutas = de_informa.rsAcerto_SelBase.RecordCount
    xbase = 0
    xok = 0
    xnao = 0
    DoEvents
    Do Until de_informa.rsAcerto_SelBase.EOF
        xbase = xbase + 1
        lblMinLida = xbase
        If de_informa.rsAcerto_SelCTC.State = 1 Then de_informa.rsAcerto_SelCTC.Close
        de_informa.Acerto_SelCTC Mid$(de_informa.rsAcerto_SelBase.Fields("REMET_CGC"), 2, 14), Val(de_informa.rsAcerto_SelBase.Fields("NUMNF"))  ', CDate("2002/11/29")
        If de_informa.rsAcerto_SelCTC.RecordCount > 0 Then
            xok = xok + 1
            lblCtcOk = xok
            xdataemi = de_informa.rsAcerto_SelBase.Fields("data")
            xdataemi = "2002/" & Mid$(xdataemi, 4, 3) & Mid$(xdataemi, 1, 2)
            de_informa.Acerto_AltDataCTC xdataemi, de_informa.rsAcerto_SelCTC.Fields("filialctc")
            de_informa.Acerto_AltDataOcorr xdataemi, de_informa.rsAcerto_SelCTC.Fields("filialctc")
        Else
            xnao = xnao + 1
            lblCtcNaoOK = xnao
        End If
        de_informa.rsAcerto_SelBase.MoveNext
        DoEvents
    Loop
End Sub

Private Sub Command2_Click()
    If de_informa.rsAcerto_SelBase.State = 1 Then de_informa.rsAcerto_SelBase.Close
    de_informa.Acerto_SelBase
    lbltotminutas = de_informa.rsAcerto_SelBase.RecordCount
    xbase = 0
    xok = 0
    xnao = 0
    DoEvents
    Do Until de_informa.rsAcerto_SelBase.EOF
        xbase = xbase + 1
        lblMinLida = xbase
        If de_informa.rsAcerto_SelCTC.State = 1 Then de_informa.rsAcerto_SelCTC.Close
        de_informa.Acerto_SelCTC Mid$(de_informa.rsAcerto_SelBase.Fields("REMET_CGC"), 2, 14), Val(de_informa.rsAcerto_SelBase.Fields("NUMNF"))  ', CDate("2002/11/29")
        If de_informa.rsAcerto_SelCTC.RecordCount > 0 Then
            xdataent = Trim$(de_informa.rsAcerto_SelBase.Fields("entrega"))
            If xdataent <> "/  /" Then
                xdataent = "2002/" & Mid$(xdataent, 4, 3) & Mid$(xdataent, 1, 2)
                If de_informa.rsAcerto_SelCTC.Fields("tem_ocorr") <> "1" And de_informa.rsAcerto_SelCTC.Fields("tem_ocorr") <> "0" And IsDate(xdataent) Then
                    xok = xok + 1
                    lblCtcOk = xok
                    If IsNull(de_informa.rsAcerto_SelBase.Fields("recebedor")) Then
                        xrecebedor = ""
                    Else
                        xrecebedor = de_informa.rsAcerto_SelBase.Fields("recebedor")
                    End If
                    de_informa.ins_ocorr1 de_informa.rsAcerto_SelCTC.Fields("filialctc"), _
                                            de_informa.rsAcerto_SelCTC.Fields("data"), _
                                            de_informa.rsAcerto_SelBase.Fields("REMET_CGC"), _
                                            "01", "ENTREGA REALIZADA", CDate(xdataent), "00:00", CDate(xdataent), "00:00", _
                                            xrecebedor, "AUTOMATICO", CVar(Date) & " " & CVar(Time()), "S", Date
                    de_informa.alt_temocorr_sn "1", de_informa.rsAcerto_SelCTC.Fields("filialctc")  'atualiza arquivo de CTC com tem_ocorr = 1
                    frmAtualPrazos.Show 1
                End If
            End If
        Else
            xnao = xnao + 1
            lblCtcNaoOK = xnao
        End If
        de_informa.rsAcerto_SelBase.MoveNext
        DoEvents
    Loop
End Sub
