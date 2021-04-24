VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiInforma 
   AutoShowChildren=   0   'False
   BackColor       =   &H0000FFFF&
   Caption         =   "Sistema Informa (Lite) - Módulo de Informação - Intec Cargo - V1.6"
   ClientHeight    =   5595
   ClientLeft      =   2235
   ClientTop       =   1905
   ClientWidth     =   7965
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiInformalite.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Timer tmAlarmeUrgencia 
      Left            =   240
      Top             =   2160
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   1535
      ButtonWidth     =   2302
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Consulta NF"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Acomp. VideoLar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5325
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.ToolTipText     =   "usuário ativo no momento"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13758
            MinWidth        =   13758
            Object.ToolTipText     =   "Observações"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "05/11/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":6936
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":6E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":6F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":70A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":722A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":7BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":7D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":82CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":86DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":8886
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":8C46
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":94DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiInformalite.frx":9A12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSac 
      Caption         =   "Consulta SAC"
   End
   Begin VB.Menu mnuVideolar 
      Caption         =   "Controle VIDEOLAR"
   End
   Begin VB.Menu mnuManut 
      Caption         =   "Manutenção da Base"
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "mdiInforma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Activate()
    Dim dia As String
    StatusBar1.Panels.Item(1).Text = "USR: " & xusuario
    StatusBar1.Panels.Item(3).Text = diasemana(datahora("data"))
End Sub

Private Sub MDIForm_Load()
    xusuario = ""
    frmAcesso.Show 1
    xamarelo1 = &HC0FFFF
    xamarelo2 = &HFFFF&
    xbranco = &H80000014
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If xusuario <> "" Then
        'LOG DE USUÁRIO
        de_informa.ins_LogUsuario "LOGOFF", xusuario, "OK"
    End If

    End
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Set mdiInforma = Nothing
End Sub

Private Sub mnuManut_Click()
    frmVLManutBase.Show 1
End Sub

Private Sub mnuSac_Click()
    If Mid$(xdireitos, 10, 1) = "0" Then
        MsgBox "Acesso Não Permitido !"
    Else
        frmVLSac.Show
    End If
End Sub
Private Sub mnuSair_Click()
'    Dim xfilialctc As String, xdata As Date, xcod As String, xdatausu As Date
    Unload Me
    
'    If de_informa.rssel_ComNull.State = 1 Then de_informa.rssel_ComNull.Close
'    de_informa.sel_ComNull
 
'    Do Until de_informa.rssel_ComNull.EOF
'        xfilialctc = de_informa.rssel_ComNull.Fields("filialctc")
'        xcod = de_informa.rssel_ComNull.Fields("cod_ocorr")
'        xdatausu = de_informa.rssel_ComNull.Fields("usu_dataocorr")
'        xdata = de_informa.rssel_ComNull.Fields("data")
        
'        If IsDate(xdata) Then
'            de_informa.Alt_DataCorreta xdata, xfilialctc, xcod, xdatausu
'        Else
'            MsgBox "nao"
'        End If
        
'        de_informa.rssel_ComNull.MoveNext
   
   
'    Loop
    
    
'    MsgBox "ababour"
    
    


    
'    If de_informa.rsSel_Acerto1.State = 1 Then de_informa.rsSel_Acerto1.Close
'    de_informa.Sel_Acerto1
'    Do Until de_informa.rsSel_Acerto1.EOF
'        xfilialctc = de_informa.rsSel_Acerto1.Fields("filialctc")
'        xdata = CDate(de_informa.rsSel_Acerto1.Fields("data"))
'        xcod = de_informa.rsSel_Acerto1.Fields("cod")
'        de_informa.Alt_AcertaOcorr xdata, xfilialctc, xcod
'        de_informa.rsSel_Acerto1.MoveNext
    
    
    
'    Loop
    
    
End Sub
Private Sub mnuVideolar_Click()
    frmVideoLarCtr.Show
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        mnuSac_Click
    ElseIf Button.Index = 2 Then
        mnuVideolar_Click
    ElseIf Button.Index = 3 Then
        mnuSair_Click
    End If
End Sub
