VERSION 5.00
Begin VB.Form FrmProductos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   7770
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Productos"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6720
         TabIndex        =   19
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton CmdCerrar 
         Caption         =   "Home"
         Height          =   375
         Left            =   6720
         TabIndex        =   17
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   6720
         TabIndex        =   16
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   6720
         TabIndex        =   15
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton CmdGuardarProd 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   6720
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   13
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         Height          =   405
         Index           =   3
         Left            =   720
         TabIndex        =   11
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1920
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   4935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Regresar"
         Height          =   615
         Left            =   6720
         TabIndex        =   3
         Top             =   3360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label7 
         Caption         =   "M.N."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   18
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Precio:   $"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "CLAVE:"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Multiplo de Conversión:"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre Corto:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Productos:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const adUseClient = 3
Const adOpenForwardOnly = 3

Dim StrQuery As String
Dim recproductos As Variant

Private Sub CmdCerrar_Click()
'FrmAntesala.Visible = True
Unload Me

End Sub

Private Sub CmdEliminar_Click()
If Text1(3).Text <> "" Then
    Dim admin As ClsAdmin
   Set admin = New ClsAdmin
       admin.Baja CInt(Text1(3).Text)
   Set admin = Nothing
   MsgBox "Se elimino un producto a la base de datos", vbOKOnly + vbInformation, "Aviso-Productos"
End If
End Sub

Private Sub CmdGuardarProd_Click()
Dim resp1 As Boolean
Dim precio As Double
   precio = Format(Trim$(Text2.Text), "##########.##")
   
   If Text1(0).Text = "" Or Text1(1).Text = "" Or Text1(2).Text = "" Then
      MsgBox "Por favor indique el nombre del producto, el nombre corto y el precio", vbOKOnly, "Aviso-MC"
      Exit Sub
   End If
   Dim admin As ClsAdmin
   
   Set admin = New ClsAdmin
   If Text1(3).Text <> "" Then
      resp1 = MsgBox("Desea Modificar el producto en la base de datos", vbYesNo + vbExclamation, "Aviso-Productos") = vbYes
      If resp1 Then
            admin.ModificacionProd CInt(Text1(3).Text), Trim$(Text1(0).Text), Trim$(Text1(1).Text), CDbl(Text1(2).Text), CDbl(precio)
      Else
         Exit Sub
      End If
   Else
       admin.Alta_Productos Trim$(Text1(0).Text), Trim$(Text1(1).Text), CDbl(Text1(2).Text), CDbl(precio)
   End If
   Set admin = Nothing
   MsgBox "Se añadio un producto a la base de datos", vbOKOnly + vbInformation, "Aviso-Productos"
End Sub

Private Sub CmdNuevo_Click()
         Combo1.Text = ""
         Text1(0).Text = ""
         Text1(1).Text = ""
         Text1(2).Text = ""
         Text1(3).Text = ""
         Text2.Text = ""
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
End Sub
Private Sub Combo1_Click()

         
   Dim prod As Integer
   
   If recproductos.RecordCount > 0 Then
      If Combo1.Text = "" Then Exit Sub
      prod = IDCombo(Combo1.Text)
      recproductos.Filter = "Idproducto = " & prod & ""
      If recproductos.RecordCount > 0 Then
         Text1(0).Text = recproductos(1)
         Text1(1).Text = recproductos(2)
         Text1(2).Text = recproductos(3)
         Text1(3).Text = recproductos(0)
      'If Text2.Text <> "" Then
         If recproductos(5) <> "" Then
            Text2.Text = recproductos(5)
         Else
            Text2.Text = ""
         End If
      'End If
      End If
End If
End Sub

Private Sub Command1_Click()
'FrmAntesala.Visible = True
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
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

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
          Case 2
          If KeyAscii = 46 Then Exit Sub
             EsEntero Text2, 9, KeyAscii
          End Select

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 46 Then Exit Sub
           EsEntero Text2, 9, KeyAscii
        
End Sub

