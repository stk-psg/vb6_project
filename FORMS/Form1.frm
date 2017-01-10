VERSION 5.00
Begin VB.Form FrmInventario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6330
   Begin VB.Frame Frame1 
      Caption         =   "Inventario"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Agregar Unidades:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Inventario Actual:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Producto:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim recproductos As Variant
Dim recinvent As Variant


Private Sub Combo1_Click()
 Dim admin As New ClsAdmin
 Dim prod As Integer
   Set admin = New ClsAdmin
       admin.RecuperarInvent recinvent
   Set admin = Nothing
   
   
   If recinvent.RecordCount > 0 Then
      If Combo1.Text = "" Then Exit Sub
      prod = IDCombo(Combo1.Text)
      recinvent.Filter = "Idproducto = " & prod & ""
      If recinvent.RecordCount > 0 Then
         Text2.Text = recinvent(1)
      Else
         Text2.Text = ""
      'End If
      End If
End If
End Sub

Private Sub Command1_Click()
Dim admin As ClsAdmin

Set admin = New ClsAdmin
    admin.Alta_Invent IDCombo(Combo1), Text2.Text, Format$(Now, "dd/mm/yyyy"), _
          CInt(Text3.Text)
Set admin = Nothing
End Sub

Private Sub Form_Load()
    Dim admin As ClsAdmin
    
    Set admin = New ClsAdmin
        admin.Recuperar recproductos
    Set admin = Nothing
     If recproductos.RecordCount > 0 Then
        ''Beep
        recproductos.MoveFirst
          Combo1.Clear
          Combo1.AddItem ""
       Do While Not recproductos.EOF
           Combo1.AddItem Trim$(recproductos(0)) & ":" & "" & (recproductos(1))
          recproductos.MoveNext
       Loop
     
     End If
     CenterForm Me
End Sub
