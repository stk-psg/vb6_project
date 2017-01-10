VERSION 5.00
Begin VB.Form FrmBus 
   Caption         =   "BuscarVentas"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Buscar"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Un solo Producto"
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Todos los Productos"
         Height          =   495
         Index           =   0
         Left            =   2160
         TabIndex        =   1
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Seleccione el tipo de Busqueda. "
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click(Index As Integer)
  Select Case Index
       Case 0
            If Check1(0).Value = 1 Then
            Check1(1).Value = 0
            Else
            Check1(1).Value = 1
            End If
       Case 1
            If Check1(1).Value = 1 Then
            Check1(0).Value = 0
            Else
            Check1(0).Value = 1
            End If
   End Select
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Check1(0).Value = 1
    CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Check1(0).Value Then
todosprod = True
Else
todosprod = False
End If
End Sub
