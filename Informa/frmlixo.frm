VERSION 5.00
Begin VB.Form frmlixo 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdprocessar 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "utilizado para colocar pré-baixas nos casos de ocorr 39 e 84 que são referente a CTC/NF retidos para COnferencia."
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   2400
      Width           =   4335
   End
End
Attribute VB_Name = "frmlixo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprocessar_Click()
    If de_informa.rssel_teste1.State = 1 Then de_informa.rssel_teste1.Close
    de_informa.sel_teste1
    
    Do Until de_informa.rssel_teste1.EOF
        xfilialctc = de_informa.rssel_teste1.Fields("filialctc")
        de_informa.ins_ocorr1 xfilialctc, _
                              de_informa.rssel_teste1.Fields("emissao"), _
                              de_informa.rssel_teste1.Fields("remet_cgc"), _
                              "01", "ENTREGA REALIZADA", _
                              de_informa.rssel_teste1.Fields("ocorr"), _
                              de_informa.rssel_teste1.Fields("hsocorr"), _
                              de_informa.rssel_teste1.Fields("ocorr"), _
                              de_informa.rssel_teste1.Fields("hsocorr"), ".", _
                              "AUTO-PREBX", de_informa.rssel_teste1.Fields("dthsusu"), "S", datahora("data")
                              
        de_informa.alt_temocorr_sn "1", xfilialctc
                              
        de_informa.rssel_teste1.MoveNext
        
    Loop
    MsgBox "fim"
End Sub
