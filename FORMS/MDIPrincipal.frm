VERSION 5.00
Begin VB.MDIForm MDIPrincipal 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Facturación Madereria Cabreras"
   ClientHeight    =   6270
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPrincipal.frx":0000
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuPrincipal 
      Caption         =   "&Facturación"
      Index           =   0
      Begin VB.Menu mnuFacturación 
         Caption         =   "&Facturas"
         Index           =   0
      End
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "&Seguridad"
      Index           =   1
      Begin VB.Menu mnuSeguridad 
         Caption         =   "&Usuarios"
         Index           =   0
      End
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "&Productos"
      Index           =   2
      Begin VB.Menu mnuProductos 
         Caption         =   "&Productos"
         Index           =   0
      End
      Begin VB.Menu mnuProductos 
         Caption         =   "&Inventario"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "&Ventas"
      Index           =   3
      Begin VB.Menu mnuVentas 
         Caption         =   "&Ventas"
         Index           =   0
      End
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "&Salir"
      Index           =   4
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuFacturación_Click(Index As Integer)
     Select Case Index
            Case 0
                  FrmClientes.Show
     End Select
End Sub

Private Sub mnuPrincipal_Click(Index As Integer)
     If Index = 4 Then Unload Me
End Sub

Private Sub mnuProductos_Click(Index As Integer)
     Select Case Index
            Case 0
                  FrmProductos.Show
            Case 1
                 FrmInventario.Show
     End Select
End Sub

Private Sub mnuSeguridad_Click(Index As Integer)
Select Case Index
            Case 0
                  FrmUsuarios.Show
     End Select
End Sub

Private Sub mnuVentas_Click(Index As Integer)
     Select Case Index
            Case 0
                  FrmVentas.Show
     End Select
End Sub
