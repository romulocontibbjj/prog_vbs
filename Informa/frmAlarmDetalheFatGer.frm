VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAlarmDetalheFatGer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalhe do Movimento do Mês"
   ClientHeight    =   8385
   ClientLeft      =   1200
   ClientTop       =   1155
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11415
   Begin VB.Timer TimerDetalhe 
      Interval        =   100
      Left            =   8760
      Top             =   7920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAIR"
      Height          =   375
      Left            =   9720
      TabIndex        =   11
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   0
      TabIndex        =   4
      Top             =   5880
      Width           =   11415
      Begin VB.Label Label4 
         Caption         =   $"frmAlarmDetalheFatGer.frx":0000
         Height          =   1215
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mês Anterior X% representava o período de 01 a 1"
         Height          =   195
         Left            =   6120
         TabIndex        =   7
         Top             =   960
         Width           =   3585
      End
      Begin VB.Label Label2 
         Caption         =   "Representa X% do total de dias deste Mês"
         Height          =   255
         Left            =   6000
         TabIndex        =   6
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mês Atual: dia 1 ao dia 11 "
         Height          =   195
         Left            =   6120
         TabIndex        =   5
         Top             =   240
         Width           =   1890
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mes Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11415
      Begin MSChart20Lib.MSChart graf0 
         Height          =   2535
         Left            =   120
         OleObjectBlob   =   "frmAlarmDetalheFatGer.frx":00E3
         TabIndex        =   3
         Top             =   240
         Width           =   9135
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexAlarmDetalheDia0 
         Height          =   2535
         Left            =   9240
         TabIndex        =   10
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mes Atual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   11415
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexAlarmDetalheDia1 
         Height          =   2535
         Left            =   9240
         TabIndex        =   9
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin MSChart20Lib.MSChart graf1 
         Height          =   2535
         Left            =   120
         OleObjectBlob   =   "frmAlarmDetalheFatGer.frx":1F4B
         TabIndex        =   1
         Top             =   240
         Width           =   9135
      End
   End
End
Attribute VB_Name = "frmAlarmDetalheFatGer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAtualizar_Click()

End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub TimerDetalhe_Timer()
TimerDetalhe.Interval = 0
flexAlarmDetalheDia0.Cols = 3
flexAlarmDetalheDia0.ColWidth(0) = 200
flexAlarmDetalheDia0.ColWidth(1) = 350
flexAlarmDetalheDia0.ColWidth(2) = 1200
flexAlarmDetalheDia0.TextMatrix(0, 1) = "Dia"
flexAlarmDetalheDia0.TextMatrix(0, 2) = "Frete"

flexAlarmDetalheDia1.Cols = 3
flexAlarmDetalheDia1.ColWidth(0) = 200
flexAlarmDetalheDia1.ColWidth(1) = 350
flexAlarmDetalheDia1.ColWidth(2) = 1200
flexAlarmDetalheDia1.TextMatrix(0, 1) = "Dia"
flexAlarmDetalheDia1.TextMatrix(0, 2) = "Frete"

'dados do mês anterior
If de_informa.rsSel_GrafFatDiaMes0.State = 1 Then de_informa.rsSel_GrafFatDiaMes0.Close
de_informa.Sel_GrafFatDiaMes0 Mid$(frmAlarmeUrg.comboMesAnoGeral.ItemData(frmAlarmeUrg.comboMesAnoGeral.ListIndex + 1), 1, 4), _
                          Mid$(frmAlarmeUrg.comboMesAnoGeral.ItemData(frmAlarmeUrg.comboMesAnoGeral.ListIndex + 1), 5, 2)
                          
If de_informa.rsSel_GrafFatDiaMes0.RecordCount < 1 Then
    MsgBox "Não será possível Montar esta análise, pois o mes anterior não houve movimento."
    Unload Me
    Exit Sub
End If

graf0.RowCount = UltDiaMes(Mid$(frmAlarmeUrg.comboMesAnoGeral.ItemData(frmAlarmeUrg.comboMesAnoGeral.ListIndex + 1), 5, 2), _
                                                       Mid$(frmAlarmeUrg.comboMesAnoGeral.ItemData(frmAlarmeUrg.comboMesAnoGeral.ListIndex + 1), 1, 4))

flexAlarmDetalheDia0.Rows = de_informa.rsSel_GrafFatDiaMes0.RecordCount + 1

For xcont = 1 To de_informa.rsSel_GrafFatDiaMes0.RecordCount
    graf0.Row = xcont
    graf0.Data = de_informa.rsSel_GrafFatDiaMes0.Fields("frete")
    graf0.RowLabel = de_informa.rsSel_GrafFatDiaMes0.Fields("dia")
    flexAlarmDetalheDia0.TextMatrix(xcont, 1) = de_informa.rsSel_GrafFatDiaMes0.Fields("dia")
    flexAlarmDetalheDia0.TextMatrix(xcont, 2) = Format(de_informa.rsSel_GrafFatDiaMes0.Fields("frete"), "###,###,##0.00")
    de_informa.rsSel_GrafFatDiaMes0.MoveNext
Next
de_informa.rsSel_GrafFatDiaMes0.MoveLast
For xcont = de_informa.rsSel_GrafFatDiaMes0.Fields("dia") + 1 To graf1.RowCount
    graf0.Row = xcont
    'tentar não mostar o line/dados para esta sequencia pois está zerada
    graf0.Data = ""
    graf0.RowLabel = xcont
Next
de_informa.rsSel_GrafFatDiaMes0.MoveFirst


'dados do mês atual
If de_informa.rsSel_GrafFatDiaMes1.State = 1 Then de_informa.rsSel_GrafFatDiaMes1.Close
de_informa.Sel_GrafFatDiaMes1 Mid$(frmAlarmeUrg.comboMesAnoGeral.ItemData(frmAlarmeUrg.comboMesAnoGeral.ListIndex), 1, 4), _
                          Mid$(frmAlarmeUrg.comboMesAnoGeral.ItemData(frmAlarmeUrg.comboMesAnoGeral.ListIndex), 5, 2)
                          
If de_informa.rsSel_GrafFatDiaMes1.RecordCount < 1 Then
    MsgBox "Não será possível Montar esta análise, pois o mes atual não houve movimento."
    Unload Me
    Exit Sub
End If

graf1.RowCount = UltDiaMes(Mid$(frmAlarmeUrg.comboMesAnoGeral.ItemData(frmAlarmeUrg.comboMesAnoGeral.ListIndex), 5, 2), _
                                                       Mid$(frmAlarmeUrg.comboMesAnoGeral.ItemData(frmAlarmeUrg.comboMesAnoGeral.ListIndex), 1, 4))
                                                       
flexAlarmDetalheDia1.Rows = de_informa.rsSel_GrafFatDiaMes1.RecordCount + 1
                                                       
For xcont = 1 To de_informa.rsSel_GrafFatDiaMes1.RecordCount
    graf1.Row = xcont
    graf1.Data = de_informa.rsSel_GrafFatDiaMes1.Fields("frete")
    graf1.RowLabel = de_informa.rsSel_GrafFatDiaMes1.Fields("dia")
    flexAlarmDetalheDia1.TextMatrix(xcont, 1) = de_informa.rsSel_GrafFatDiaMes1.Fields("dia")
    flexAlarmDetalheDia1.TextMatrix(xcont, 2) = Format(de_informa.rsSel_GrafFatDiaMes1.Fields("frete"), "###,###,##0.00")
    de_informa.rsSel_GrafFatDiaMes1.MoveNext
Next
de_informa.rsSel_GrafFatDiaMes1.MoveLast
For xcont = de_informa.rsSel_GrafFatDiaMes1.Fields("dia") + 1 To graf1.RowCount
    graf1.Row = xcont
    'tentar não mostar o line/dados para esta sequencia pois está zerada
    graf1.Data = ""
    graf1.RowLabel = xcont
Next
de_informa.rsSel_GrafFatDiaMes1.MoveFirst



End Sub
