VERSION 5.00
Begin VB.Form Frm_Produtos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtos"
   ClientHeight    =   4545
   ClientLeft      =   2790
   ClientTop       =   3225
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5430
   Begin VB.TextBox txtDiscontinued 
      DataField       =   "Discontinued"
      DataMember      =   "Products"
      DataSource      =   "deb_nwind"
      Height          =   285
      Left            =   1470
      TabIndex        =   19
      Top             =   3960
      Width           =   330
   End
   Begin VB.TextBox txtReorderLevel 
      DataField       =   "ReorderLevel"
      DataMember      =   "Products"
      DataSource      =   "deb_nwind"
      Height          =   285
      Left            =   1470
      TabIndex        =   17
      Top             =   3585
      Width           =   330
   End
   Begin VB.TextBox txtUnitsOnOrder 
      DataField       =   "UnitsOnOrder"
      DataMember      =   "Products"
      DataSource      =   "deb_nwind"
      Height          =   285
      Left            =   1470
      TabIndex        =   15
      Top             =   3195
      Width           =   330
   End
   Begin VB.TextBox txtUnitsInStock 
      DataField       =   "UnitsInStock"
      DataMember      =   "Products"
      DataSource      =   "deb_nwind"
      Height          =   285
      Left            =   1470
      TabIndex        =   13
      Top             =   2820
      Width           =   330
   End
   Begin VB.TextBox txtUnitPrice 
      DataField       =   "UnitPrice"
      DataMember      =   "Products"
      DataSource      =   "deb_nwind"
      Height          =   285
      Left            =   1470
      TabIndex        =   11
      Top             =   2445
      Width           =   1320
   End
   Begin VB.TextBox txtQuantityPerUnit 
      DataField       =   "QuantityPerUnit"
      DataMember      =   "Products"
      DataSource      =   "deb_nwind"
      Height          =   285
      Left            =   1470
      TabIndex        =   9
      Top             =   2055
      Width           =   3300
   End
   Begin VB.TextBox txtCategoryID 
      DataField       =   "CategoryID"
      DataMember      =   "Products"
      DataSource      =   "deb_nwind"
      Height          =   285
      Left            =   1470
      TabIndex        =   7
      Top             =   1680
      Width           =   660
   End
   Begin VB.TextBox txtSupplierID 
      DataField       =   "SupplierID"
      DataMember      =   "Products"
      DataSource      =   "deb_nwind"
      Height          =   285
      Left            =   1470
      TabIndex        =   5
      Top             =   1305
      Width           =   660
   End
   Begin VB.TextBox txtProductName 
      DataField       =   "ProductName"
      DataMember      =   "Products"
      DataSource      =   "deb_nwind"
      Height          =   285
      Left            =   1470
      TabIndex        =   3
      Top             =   915
      Width           =   3375
   End
   Begin VB.TextBox txtProductID 
      DataField       =   "ProductID"
      DataMember      =   "Products"
      DataSource      =   "deb_nwind"
      Height          =   285
      Left            =   1470
      TabIndex        =   1
      Top             =   540
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Discontinued:"
      Height          =   255
      Index           =   9
      Left            =   465
      TabIndex        =   18
      Top             =   4005
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ReorderLevel:"
      Height          =   255
      Index           =   8
      Left            =   345
      TabIndex        =   16
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UnitsOnOrder:"
      Height          =   255
      Index           =   7
      Left            =   345
      TabIndex        =   14
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UnitsInStock:"
      Height          =   255
      Index           =   6
      Left            =   465
      TabIndex        =   12
      Top             =   2865
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UnitPrice:"
      Height          =   255
      Index           =   5
      Left            =   705
      TabIndex        =   10
      Top             =   2490
      Width           =   735
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "QuantityPerUnit:"
      Height          =   255
      Index           =   4
      Left            =   225
      TabIndex        =   8
      Top             =   2100
      Width           =   1215
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CategoryID:"
      Height          =   255
      Index           =   3
      Left            =   465
      TabIndex        =   6
      Top             =   1725
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SupplierID:"
      Height          =   255
      Index           =   2
      Left            =   585
      TabIndex        =   4
      Top             =   1350
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ProductName:"
      Height          =   255
      Index           =   1
      Left            =   345
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ProductID:"
      Height          =   255
      Index           =   0
      Left            =   585
      TabIndex        =   0
      Top             =   585
      Width           =   855
   End
End
Attribute VB_Name = "Frm_Produtos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
