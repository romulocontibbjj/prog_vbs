VERSION 5.00
Begin VB.Form FRM_principal 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Tela Principal"
   ClientHeight    =   8340
   ClientLeft      =   825
   ClientTop       =   1665
   ClientWidth     =   11910
   Icon            =   "FRM_principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Menu Mnu_Arquivo 
      Caption         =   "Arquivo"
      Index           =   1
      Begin VB.Menu Mnu_cadastros 
         Caption         =   "Efetuar Cadastro"
         Index           =   2
         Begin VB.Menu Mnu_CadastroClientes 
            Caption         =   "Cliestes"
            Index           =   3
         End
         Begin VB.Menu Mnu_CadastroFuncionarios 
            Caption         =   "Motoqueiros"
            Index           =   4
         End
         Begin VB.Menu Mnu_CadastroMoto 
            Caption         =   "Motocicletas"
            Index           =   5
         End
      End
      Begin VB.Menu Mnu_Sair 
         Caption         =   "Sair"
         Index           =   13
      End
   End
   Begin VB.Menu MenuEfetuarPagamento 
      Caption         =   "Pagamentos"
      Index           =   6
      Begin VB.Menu MnuPagamentosVales 
         Caption         =   "Vale de Funcionários"
         Index           =   8
      End
      Begin VB.Menu Mnu_pagamentosDespesas 
         Caption         =   "Despesas Diverças"
         Index           =   9
      End
   End
   Begin VB.Menu MnuRelatorisos 
      Caption         =   "Relatórios"
      Index           =   10
      Begin VB.Menu MnuRelatoriosDespesas 
         Caption         =   "Despesa mensal"
         Index           =   11
      End
      Begin VB.Menu MnuRelatoriosSalarios 
         Caption         =   "Débito Salários"
         Index           =   12
      End
   End
End
Attribute VB_Name = "FRM_principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Mnu_CadastroClientes_Click(Index As Integer)
FRM_CadastrodeClientes.Show modal, Me
End Sub

Private Sub Mnu_CadastroFuncionarios_Click(Index As Integer)
FRM_CadastrodeMotoqueiros.Show modal, Me
End Sub

Private Sub Mnu_CadastroMoto_Click(Index As Integer)
FRM_CadastrodeMotos.Show modal, Me
End Sub

Private Sub Mnu_Sair_Click(Index As Integer)
End
End Sub
