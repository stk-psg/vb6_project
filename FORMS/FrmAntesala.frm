VERSION 5.00
Begin VB.Form FrmAntesala 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Antesala"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Ventas"
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Inventario"
      Height          =   975
      Left            =   5040
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Productos"
      Height          =   975
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Facturas"
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "FrmAntesala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FrmAntesala.Visible = False
FrmClientes.Show
End Sub

Private Sub Command2_Click()
FrmAntesala.Visible = False
FrmProductos.Show
End Sub
