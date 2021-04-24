VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVLBuscaCli 
   Caption         =   "Busca Clientes Destinatários - Fox Film"
   ClientHeight    =   4845
   ClientLeft      =   1935
   ClientTop       =   1500
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7800
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7575
      Begin MSDataGridLib.DataGrid gridClientes 
         Bindings        =   "frmVLBuscaCli.frx":0000
         Height          =   2895
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5106
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "Sel_VLBuscaCli"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "serie"
            Caption         =   "serie"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "cliente_cgc"
            Caption         =   "cliente_cgc"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cliente_sap"
            Caption         =   "cliente_sap"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "cliente_nome"
            Caption         =   "cliente_nome"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cliente_fantasia"
            Caption         =   "cliente_fantasia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "grupo"
            Caption         =   "grupo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "tipo"
            Caption         =   "tipo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            Size            =   243
            BeginProperty Column00 
               ColumnWidth     =   464,882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1409,953
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2475,213
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1530,142
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1695,118
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busca por Nome / Parte do Nome"
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
      TabIndex        =   3
      Top             =   3480
      Width           =   7575
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6000
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancela"
         Height          =   375
         Left            =   6000
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Busca..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3960
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNome 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   0
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblAguarde 
         AutoSize        =   -1  'True
         Caption         =   "Aguarde ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Cliente:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmVLBuscaCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdBusca_Click()
    Me.MousePointer = 11
    DoEvents
    cmdBusca.Enabled = False
    DoEvents
    cmdConfirmar.Enabled = False
    DoEvents
    cmdCancelar.Enabled = False
    DoEvents
    lblAguarde.Visible = True
    DoEvents
    
    If de_informa.rsSel_VLBuscaCli.State = 1 Then de_informa.rsSel_VLBuscaCli.Close
    de_informa.Sel_VLBuscaCli "%" & Trim$(txtNome) & "%"
    
    gridClientes.DataMember = "sel_vlbuscacli"
    gridClientes.Refresh
    
    Me.MousePointer = 0
    cmdBusca.Enabled = True
    cmdConfirmar.Enabled = True
    cmdCancelar.Enabled = True
    lblAguarde.Visible = False
    DoEvents
    
    If de_informa.rsSel_VLBuscaCli.RecordCount = 0 Then
        MsgBox "Não Encontrado Cliente com Este Nome !"
        cmdConfirmar.Enabled = False
        txtNome.SetFocus
    Else
        cmdConfirmar.Enabled = True
    End If
    
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdConfirmar_Click()
    If frmVideoLarCtr.optCliCnpj.Value = True Then
        frmVideoLarCtr.txtFiltroCliente = gridClientes.Columns(1)
    ElseIf frmVideoLarCtr.optCliCodSap.Value = True Then
        frmVideoLarCtr.txtFiltroCliente = gridClientes.Columns(2)
    ElseIf frmVideoLarCtr.optCliFantasia.Value = True Then
        frmVideoLarCtr.txtFiltroCliente = gridClientes.Columns(4)
    End If

    Unload Me
End Sub
Private Sub Form_Load()
    If de_informa.rsSel_VLBuscaCli.State = 1 Then de_informa.rsSel_VLBuscaCli.Close
    gridClientes.DataMember = "sel_vlbuscacli"
    gridClientes.Refresh
End Sub
Private Sub txtNome_Change()
    If Len(txtNome.Text) >= 3 Then
        cmdBusca.Enabled = True
    Else
        cmdBusca.Enabled = False
    End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'TECLA ENTER
        KeyAscii = 0
        SendKeys "{TAB}"  'ENVIA UM TAB
    End If
End Sub
