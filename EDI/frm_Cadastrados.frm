VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Cadastrados 
   Caption         =   "EDIS CADASTRADOS"
   ClientHeight    =   6420
   ClientLeft      =   630
   ClientTop       =   1635
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   9180
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin VB.CommandButton Cmd_Excluir 
         Caption         =   "&EXCLUIR"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_Sair 
         Caption         =   "&SAIR"
         Height          =   375
         Left            =   4800
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Alterar 
         Caption         =   "&ALTERAR"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_inserir 
         Caption         =   "&INSERIR"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid grd_Edi 
         Bindings        =   "frm_Cadastrados.frx":0000
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9340
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
         DataMember      =   "Sel_CadEdis"
         ColumnCount     =   16
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "id"
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
            DataField       =   "cgc"
            Caption         =   "cgc"
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
            DataField       =   "edi"
            Caption         =   "edi"
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
            DataField       =   "cliente"
            Caption         =   "cliente"
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
            DataField       =   "email"
            Caption         =   "email"
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
            DataField       =   "assunto"
            Caption         =   "assunto"
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
            DataField       =   "mensagem"
            Caption         =   "mensagem"
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
            DataField       =   "Salvar"
            Caption         =   "Salvar"
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
         BeginProperty Column08 
            DataField       =   "nomearq"
            Caption         =   "nomearq"
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
         BeginProperty Column09 
            DataField       =   "ddmm"
            Caption         =   "ddmm"
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
         BeginProperty Column10 
            DataField       =   "dia"
            Caption         =   "dia"
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
         BeginProperty Column11 
            DataField       =   "horario"
            Caption         =   "horario"
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
         BeginProperty Column12 
            DataField       =   "semana"
            Caption         =   "semana"
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
         BeginProperty Column13 
            DataField       =   "periodo"
            Caption         =   "periodo"
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
         BeginProperty Column14 
            DataField       =   "entrega"
            Caption         =   "entrega"
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
         BeginProperty Column15 
            DataField       =   "cancelados"
            Caption         =   "cancelados"
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
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   915,024
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_Cadastrados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Alterar_Click()

With grd_Edi
    frm_CadEdis.xid = .Columns(0)
    frm_CadEdis.txt_cgc.Text = .Columns(1)
    frm_CadEdis.txt_Cliente.Text = .Columns(3)
    frm_CadEdis.cmb_edi.Text = .Columns(2)
    
    If .Columns(2) = "CONEMB" Or .Columns(2) = "CORREIOS" Then
        frm_CadEdis.txt_periodo.Text = .Columns(13)
        frm_CadEdis.chk_entrega.Value = .Columns(14)
    End If
    
    
    frm_CadEdis.txt_Email.AddItem .Columns(4)
    frm_CadEdis.txt_assunto.Text = .Columns(5)
    frm_CadEdis.txt_Mensagem.Text = .Columns(6)
    frm_CadEdis.txt_Salvar.Text = .Columns(7)
    frm_CadEdis.txt_Arquivo.Text = .Columns(8)
    If .Columns(9) <> "" Then
        frm_CadEdis.chk_ddmm.Value = .Columns(9)
    Else
        frm_CadEdis.chk_ddmm.Value = 0
    End If
    frm_CadEdis.txt_Dia.Text = .Columns(10)
    frm_CadEdis.mask_hora.Text = .Columns(11)
    
    
    If .Columns(12) = "0111110" Then
        frm_CadEdis.Check9.Value = 1
    ElseIf .Columns(10) > 0 Then
        frm_CadEdis.Check8.Value = 1
    Else
        frm_CadEdis.Check1.Value = Mid(.Columns(12), 1, 1)
        frm_CadEdis.Check2.Value = Mid(.Columns(12), 2, 1)
        frm_CadEdis.Check3.Value = Mid(.Columns(12), 3, 1)
        frm_CadEdis.Check4.Value = Mid(.Columns(12), 4, 1)
        frm_CadEdis.Check5.Value = Mid(.Columns(12), 5, 1)
        frm_CadEdis.Check6.Value = Mid(.Columns(12), 6, 1)
        frm_CadEdis.Check7.Value = Mid(.Columns(12), 7, 1)
            
    End If
    
      
    
    frm_CadEdis.Show
    Unload Me
    
    
    
    
        
    

End With

End Sub

Private Sub Cmd_Excluir_Click()

If MsgBox("Deseja Realmente Excluir o EDI???", vbInformation + vbYesNo, "EDI - EXCLUSÃO") = vbYes Then

    deb_edi.Del_CadEdis grd_Edi.Columns(0)
    
    If deb_edi.rsSel_CadEdis.State = 1 Then deb_edi.rsSel_CadEdis.Close
        deb_edi.Sel_CadEdis
        
        grd_Edi.DataMember = "Sel_CadEdis"
        grd_Edi.Refresh
         
    
    
End If



End Sub

Private Sub cmd_inserir_Click()

frm_CadEdis.Show

frm_CadEdis.xid = 0


Unload Me


End Sub

Private Sub cmd_Sair_Click()

Unload Me


End Sub

