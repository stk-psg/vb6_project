VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   11775
   Begin VB.Frame Frame1 
      Caption         =   "Ventas"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   11535
      Begin VB.TextBox Txtm3 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   8760
         TabIndex        =   22
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "RecuperaTodo"
         Height          =   375
         Left            =   9720
         TabIndex        =   20
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Borrar Todo"
         Height          =   495
         Left            =   10440
         TabIndex        =   19
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox TxtCIVA 
         BackColor       =   &H00C0FFC0&
         Height          =   405
         Left            =   6120
         TabIndex        =   17
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox TxtSIVA 
         BackColor       =   &H00C0FFC0&
         Height          =   405
         Left            =   3480
         TabIndex        =   15
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox TxtUV 
         BackColor       =   &H00C0FFC0&
         Height          =   405
         Left            =   840
         TabIndex        =   13
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   9720
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox TxtNC 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton CmdBP2 
         Caption         =   "B.P."
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton CmdProd 
         Caption         =   "C.P."
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPFin 
         Height          =   375
         Left            =   7680
         TabIndex        =   7
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16842753
         CurrentDate     =   36520
      End
      Begin MSComCtl2.DTPicker DTPIni 
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16842753
         CurrentDate     =   36520
      End
      Begin VB.Label Label9 
         Caption         =   "Metros Cubicos:"
         Height          =   375
         Left            =   8040
         TabIndex        =   23
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Total Con IVA:"
         Height          =   375
         Left            =   5400
         TabIndex        =   18
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Total Sin IVA:"
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Unidades Vendidas:"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fin:"
         Height          =   255
         Left            =   7320
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio:"
         Height          =   375
         Left            =   4800
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre Corto:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Venta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "MontoSinIVA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "MontoConIVA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "CantidadProd"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "CostoUnitario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "FechaVenta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Conversión"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Metros Cúbicos"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "HISTÓRICO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "FrmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim recproductos As Variant
Dim recproductos1 As Variant
Dim recproductos2 As Variant
Dim RecVentas As Variant
Dim RecVtasFecha As Variant

Private Sub CmdBP2_Click()
    Combo1.Clear
        
    If TxtNC.Text <> "" Then
    Dim recupventas As ClsAdmin
       Set recupventas = New ClsAdmin
           recupventas.Recuperar recproductos1
       Set recupventas = Nothing
                  
          If recproductos1.RecordCount > 0 Then
             recproductos1.Filter = "UnidadMedida = '" & Trim(TxtNC.Text) & "' "
                If recproductos1.RecordCount > 0 Then
                   Combo1.Text = recproductos1(0) & ":" & recproductos1(1)
                Else
                   MsgBox "No se encontro el producto", vbExclamation + vbOKOnly, "Aviso - MC"
                End If
           End If
    End If
End Sub

Private Sub CmdBuscar_Click()
Dim itmx As ListItem
Dim i As Integer
FrmBus.Show 1
    
    If todosprod Then
            Dim recupventas As ClsAdmin
            
            Set recupventas = New ClsAdmin
                recupventas.RecuperarVentasTodo RecVtasFecha, DTPIni.Value, _
                DTPFin.Value
            Set recupventas = Nothing
            
            If RecVtasFecha.RecordCount > 0 Then
               RecVtasFecha.Sort = "IdVenta Desc"
               RecVtasFecha.MoveFirst
                     
               
                     
                     ListView1.ListItems.Clear
                     Do While Not RecVtasFecha.EOF
                        Set itmx = ListView1.ListItems.Add(, , Trim$(RecVtasFecha(0)))
                            itmx.SubItems(1) = Trim$(RecVtasFecha(1))
                            itmx.SubItems(2) = Trim$(RecVtasFecha(2))
                            itmx.SubItems(3) = Trim$(RecVtasFecha(3))
                            itmx.SubItems(4) = Trim$(RecVtasFecha(4))
                            itmx.SubItems(5) = Trim$(RecVtasFecha(5))
                            itmx.SubItems(6) = Trim$(RecVtasFecha(6))
                            itmx.SubItems(7) = Trim$(RecVtasFecha(7))
                            
                           ' If Not IsNull(RecVentas(8)) Then
                           '     itmx.SubItems(8) = Trim$(RecVentas(8))
                           ' End If
                           ' If Not IsNull(RecVentas(9)) Then
                           '     itmx.SubItems(9) = Trim$(RecVentas(9))
                           ' End If
                            
                            
                            RecVtasFecha.MoveNext
                     Loop
            Else
                 MsgBox "No se encontro el registro del Producto", vbOKOnly + vbInformation, "Aviso-MC"
                 Exit Sub
            End If


    Else
            Set recupventas = New ClsAdmin
                recupventas.RecuperarVentasFecha Mid$(Combo1.Text, InStr(Combo1.Text, ":") + 1), DTPIni.Value, _
                DTPFin.Value, RecVtasFecha
            Set recupventas = Nothing
            
            If RecVtasFecha.RecordCount > 0 Then
               RecVtasFecha.Sort = "IdVenta Desc"
               RecVtasFecha.MoveFirst
                     
               
                     
                     ListView1.ListItems.Clear
                     Do While Not RecVtasFecha.EOF
                        Set itmx = ListView1.ListItems.Add(, , Trim$(RecVtasFecha(0)))
                            itmx.SubItems(1) = Trim$(RecVtasFecha(1))
                            itmx.SubItems(2) = Trim$(RecVtasFecha(2))
                            itmx.SubItems(3) = Trim$(RecVtasFecha(3))
                            itmx.SubItems(4) = Trim$(RecVtasFecha(4))
                            itmx.SubItems(5) = Trim$(RecVtasFecha(5))
                            itmx.SubItems(6) = Trim$(RecVtasFecha(6))
                            itmx.SubItems(7) = Trim$(RecVtasFecha(7))
                            
                            If Not IsNull(RecVtasFecha(8)) Then
                itmx.SubItems(8) = Trim$(RecVtasFecha(8))
                End If
                If Not IsNull(RecVtasFecha(9)) Then
                itmx.SubItems(9) = Trim$(RecVtasFecha(9))
                End If
                            
                            RecVtasFecha.MoveNext
                     Loop
            Else
                 MsgBox "No se encontro el registro del Producto", vbOKOnly + vbInformation, "Aviso-MC"
                 Exit Sub
            End If
    End If
Label8.Caption = "VENTAS POR PRODUCTO"
DameTotales

End Sub

Private Sub CmdProd_Click()

Dim recupventas As ClsAdmin

Set recupventas = New ClsAdmin
                recupventas.Recuperar recproductos
            Set recupventas = Nothing
            If recproductos.RecordCount > 0 Then
                    
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
    TxtNC.Text = ""
    If Combo1.Text <> "" Then
        Dim recupventas As ClsAdmin
        
        Set recupventas = New ClsAdmin
               recupventas.Recuperar recproductos2
        Set recupventas = Nothing
        If recproductos2.RecordCount > 0 Then
             recproductos2.Filter = "IDPRODUCTO = " & IDCombo(Combo1.Text) & ""
                If recproductos2.RecordCount > 0 Then
                   TxtNC.Text = recproductos2(2)
                End If
        End If
    End If
End Sub

Private Sub Command1_Click()
If MsgBox("¿Esta seguro  de Eliminar Todas las Ventas?", vbCritical + vbYesNo, "Aviso - MC") = vbYes Then
    Dim recupventas As ClsAdmin
    
   Set recupventas = New ClsAdmin
       recupventas.BajaTodoVenta
   Set recupventas = Nothing

   MsgBox "Se eliminaron todas las ventas", vbOKOnly, "Aviso - MC"
ListView1.ListItems.Clear
Else
   Exit Sub
End If

End Sub

Private Sub Form_Load()
    Dim itmx As ListItem
            Dim recupventas As ClsAdmin
            
            Set recupventas = New ClsAdmin
                recupventas.Recuperar recproductos
            Set recupventas = Nothing
            If recproductos.RecordCount > 0 Then
                    
                    recproductos.MoveFirst
                      Combo1.Clear
                      Combo1.AddItem ""
                   Do While Not recproductos.EOF
                       Combo1.AddItem Trim$(recproductos(0)) & ":" & "" & (recproductos(1))
                      recproductos.MoveNext
                   Loop
            End If
        
        
        
    Set recupventas = New ClsAdmin
        recupventas.RecuperarVentas RecVentas
    Set recupventas = Nothing

    If RecVentas.RecordCount > 0 Then
       RecVentas.Sort = "IdVenta Desc"
       RecVentas.MoveFirst
         
        
         
         ListView1.ListItems.Clear
         Do While Not RecVentas.EOF
            Set itmx = ListView1.ListItems.Add(, , Trim$(RecVentas(0)))
                itmx.SubItems(1) = Trim$(RecVentas(1))
                itmx.SubItems(2) = Trim$(RecVentas(2))
                itmx.SubItems(3) = Trim$(RecVentas(3))
                itmx.SubItems(4) = Trim$(RecVentas(4))
                itmx.SubItems(5) = Trim$(RecVentas(5))
                itmx.SubItems(6) = Trim$(RecVentas(6))
                itmx.SubItems(7) = Trim$(RecVentas(7))
                If Not IsNull(RecVentas(8)) Then
                itmx.SubItems(8) = Trim$(RecVentas(8))
                End If
                If Not IsNull(RecVentas(9)) Then
                itmx.SubItems(9) = Trim$(RecVentas(9))
                End If
                RecVentas.MoveNext
         Loop
         
         'Label3.Caption = "Registros encontrados: " & reg
    End If
DTPFin.Value = Format(Date, "dd/MM/yyyy")
DTPIni.Value = Format(Date, "dd/MM/yyyy")
CenterForm Me



End Sub



Private Sub TxtNC_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
         CmdBP2_Click
      End If

End Sub
Private Sub DameTotales()
Dim i As Integer
Dim y As Integer
Dim c As Long
Dim d As Double
Dim e As Double
Dim t As Double
Dim x As Double
Dim m As Double
Dim m3 As Double

If ListView1.ListItems.Count > 0 Then
  For i = 1 To ListView1.ListItems.Count
   y = ListView1.ListItems(i).ListSubItems(4).Text
         ''ListView(Index).ListItems.Remove ListView(Index).ListItems(i).Index
         ''Exit Sub
   c = c + y

  Next i
  
End If

TxtUV = c

If ListView1.ListItems.Count > 0 Then
  For i = 1 To ListView1.ListItems.Count
   d = ListView1.ListItems(i).ListSubItems(2).Text
         ''ListView(Index).ListItems.Remove ListView(Index).ListItems(i).Index
         ''Exit Sub
   e = e + d

  Next i
  
End If

TxtSIVA = Format(e, "##########.##")

If ListView1.ListItems.Count > 0 Then
  For i = 1 To ListView1.ListItems.Count
   x = ListView1.ListItems(i).ListSubItems(3).Text
         ''ListView(Index).ListItems.Remove ListView(Index).ListItems(i).Index
         ''Exit Sub
   t = t + x

  Next i
  
End If

TxtCIVA = Format(t, "##########.##")

''**********metros
If ListView1.ListItems.Count > 0 Then
  
  
  For i = 1 To ListView1.ListItems.Count
   If ListView1.ListItems(i).SubItems(9) <> "" Then
   m = ListView1.ListItems(i).ListSubItems(9).Text
         ''ListView(Index).ListItems.Remove ListView(Index).ListItems(i).Index
         ''Exit Sub
   m3 = m3 + m
   End If
  Next i
  
End If

Txtm3 = m3
''**********metros

End Sub

Private Sub Command2_Click()
Dim itmx As ListItem
Dim recupventas As ClsAdmin

Set recupventas = New ClsAdmin
        recupventas.RecuperarVentas RecVentas
    Set recupventas = Nothing

    If RecVentas.RecordCount > 0 Then
       RecVentas.Sort = "IdVenta Desc"
       RecVentas.MoveFirst
         
        
         
         ListView1.ListItems.Clear
         Do While Not RecVentas.EOF
            Set itmx = ListView1.ListItems.Add(, , Trim$(RecVentas(0)))
                itmx.SubItems(1) = Trim$(RecVentas(1))
                itmx.SubItems(2) = Trim$(RecVentas(2))
                itmx.SubItems(3) = Trim$(RecVentas(3))
                itmx.SubItems(4) = Trim$(RecVentas(4))
                itmx.SubItems(5) = Trim$(RecVentas(5))
                itmx.SubItems(6) = Trim$(RecVentas(6))
                itmx.SubItems(7) = Trim$(RecVentas(7))
                
                If Not IsNull(RecVentas(8)) Then
                itmx.SubItems(8) = Trim$(RecVentas(8))
                End If
                If Not IsNull(RecVentas(9)) Then
                itmx.SubItems(9) = Trim$(RecVentas(9))
                End If
                
                
                RecVentas.MoveNext
         Loop
    
    Else
        MsgBox "No  se encontro ningun registro", vbInformation + vbOKOnly, "Aviso-MC"
        Exit Sub
    End If
    DameTotales
    Label8.Caption = "HISTÓRICO"
End Sub

