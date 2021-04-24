VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRM_CadastrodeMotos 
   Caption         =   "Cadastro de Motos"
   ClientHeight    =   6150
   ClientLeft      =   1560
   ClientTop       =   3375
   ClientWidth     =   10110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6493.401
   ScaleMode       =   0  'User
   ScaleWidth      =   9830.074
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame_cadastroClientes 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton CbtSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CbtIncluir 
         Caption         =   "&Incluir"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CbtExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CbtAlterar 
         Caption         =   "&Alterar"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid grd_motos 
      Bindings        =   "FRM_CadastrodeMotos.frx":0000
      Height          =   5295
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9340
      _Version        =   393216
      BackColor       =   8388608
      ForeColor       =   65535
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
      DataMember      =   "Sel_Motos"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "MOTOQUEIRO"
         Caption         =   "MOTOQUEIRO"
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
         DataField       =   "PLACA"
         Caption         =   "PLACA"
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
         DataField       =   "MARCA"
         Caption         =   "MARCA"
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
         DataField       =   "MODELO"
         Caption         =   "MODELO"
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
         DataField       =   "ANO"
         Caption         =   "ANO"
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
         DataField       =   "CIDADE"
         Caption         =   "CIDADE"
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
         DataField       =   "UF"
         Caption         =   "UF"
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
      BeginProperty Column07 
         DataField       =   "COR"
         Caption         =   "COR"
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
         BeginProperty Column00 
            ColumnWidth     =   1691,731
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   889,688
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1691,731
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1691,731
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   743,612
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1691,731
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   598,087
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1691,731
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRM_CadastrodeMotos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CbtAlterar_Click()
FRM_FichadeMotos.Show

With deb_coty.rsSel_Motos

FRM_FichadeMotos.txt_letra.Text = Mid(.Fields("PLACA"), 1, 3)
FRM_FichadeMotos.txt_numero.Text = Mid(.Fields("PLACA"), 4, 4)
FRM_FichadeMotos.cmb_cor.Text = .Fields("COR")
FRM_FichadeMotos.cmb_marca.Text = .Fields("MARCA")
FRM_FichadeMotos.cmb_motoqueiros.Text = .Fields("MOTOQUEIRO")
FRM_FichadeMotos.txt_Cidade.Text = .Fields("CIDADE")
FRM_FichadeMotos.txt_modelo.Text = .Fields("MODELO")
FRM_FichadeMotos.cmb_Uf.Text = .Fields("UF")
FRM_FichadeMotos.mask_ano.Text = .Fields("ANO")


End With


If deb_coty.rsSel_Motoqueiros.State = 1 Then deb_coty.rsSel_Motoqueiros.Close
        deb_coty.Sel_Motoqueiros
        
        If deb_coty.rsSel_Motoqueiros.RecordCount > 0 Then
            deb_coty.rsSel_Motoqueiros.MoveFirst
            Do Until deb_coty.rsSel_Motoqueiros.EOF
                
                xnome = branco_separa(deb_coty.rsSel_Motoqueiros.Fields("nome"))
                FRM_FichadeMotos.cmb_motoqueiros.AddItem xnome
                deb_coty.rsSel_Motoqueiros.MoveNext
            Loop
        
        End If

FRM_FichadeMotos.txt_letra.SetFocus
FRM_FichadeMotos.txt_letra.Locked = True
FRM_FichadeMotos.txt_numero.Locked = True
FRM_FichadeMotos.cmd_altera.Visible = True

End Sub

Private Sub CbtExcluir_Click()

If deb_coty.rsSel_Motos.State = 1 Then deb_coty.rsSel_Motos.Close
        deb_coty.Sel_Motos
        
    If deb_coty.rsSel_Motos.RecordCount = 0 Then
        MsgBox "Não Há Registros para Exclusão", vbExclamation, "MOTOCICLETAS"
        Exit Sub
    Else
        
        
        If MsgBox("Excluir Moto: " & deb_coty.rsSel_Motos.Fields("PLACA") & " ?", vbYesNo, "MOTOCICLETAS") = vbYes Then
            deb_coty.Del_moto deb_coty.rsSel_Motos.Fields("PLACA")

            If deb_coty.rsSel_Motos.State = 1 Then deb_coty.rsSel_Motos.Close
            deb_coty.Sel_Motos
    
            FRM_CadastrodeMotos.grd_motos.DataMember = "Sel_Motos"
            FRM_CadastrodeMotos.grd_motos.Refresh
        
        End If
    
    End If


End Sub

Private Sub CbtIncluir_Click()
Dim xnome As String
FRM_FichadeMotos.Show

       If deb_coty.rsSel_Motoqueiros.State = 1 Then deb_coty.rsSel_Motoqueiros.Close
        deb_coty.Sel_Motoqueiros
        
        If deb_coty.rsSel_Motoqueiros.RecordCount > 0 Then
            deb_coty.rsSel_Motoqueiros.MoveFirst
            Do Until deb_coty.rsSel_Motoqueiros.EOF
                
                xnome = deb_coty.rsSel_Motoqueiros.Fields("nome")
                FRM_FichadeMotos.cmb_motoqueiros.AddItem xnome
                deb_coty.rsSel_Motoqueiros.MoveNext
            Loop
        
        End If
        
End Sub

Private Sub CbtSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

If deb_coty.rsSel_Motos.State = 1 Then deb_coty.rsSel_Motos.Close
    deb_coty.Sel_Motos
    
    grd_motos.DataMember = "Sel_Motos"
    grd_motos.Refresh


End Sub
