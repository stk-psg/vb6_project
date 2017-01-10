VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   9750
   Begin VB.TextBox Text21 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   63
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox Text20 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   61
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox Text19 
      Height          =   615
      Left            =   120
      TabIndex        =   57
      Text            =   "Text19"
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text18 
      Height          =   615
      Left            =   120
      TabIndex        =   56
      Text            =   "Text18"
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton CmdAgregar2 
      Caption         =   "agregar spread"
      Height          =   375
      Left            =   120
      TabIndex        =   46
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   495
      Left            =   8520
      TabIndex        =   45
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton CmdTotal 
      Caption         =   "&Total"
      Height          =   495
      Left            =   8520
      TabIndex        =   44
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   495
      Left            =   8520
      TabIndex        =   43
      Top             =   3960
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHventa 
      Height          =   2535
      Left            =   120
      TabIndex        =   42
      Top             =   5040
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4471
      _Version        =   393216
      Rows            =   10
      Cols            =   9
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ventas"
      Height          =   1335
      Left            =   120
      TabIndex        =   29
      Top             =   3600
      Width           =   8175
      Begin VB.CommandButton Command1 
         Caption         =   "C.P."
         Height          =   375
         Left            =   4320
         TabIndex        =   60
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdBP 
         Caption         =   "B.P."
         Height          =   375
         Left            =   3720
         TabIndex        =   59
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         TabIndex        =   55
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   41
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   37
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         TabIndex        =   33
         Text            =   "Combo2"
         Top             =   960
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         TabIndex        =   31
         Text            =   "Combo1"
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label18 
         Caption         =   "Total Iva:"
         Height          =   375
         Left            =   6720
         TabIndex        =   40
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Total:"
         Height          =   375
         Left            =   6720
         TabIndex        =   38
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "Num. Unidades:"
         Height          =   615
         Left            =   4800
         TabIndex        =   36
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Precio Unitario:"
         Height          =   375
         Left            =   4920
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "N&ombre Corto:"
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdMostrar 
      Caption         =   "&Clientes"
      Height          =   495
      Left            =   8520
      TabIndex        =   27
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar"
      Height          =   495
      Left            =   8520
      TabIndex        =   26
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton CmdRegresar 
      Caption         =   "Home"
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nue&vo"
      Height          =   495
      Left            =   8520
      TabIndex        =   24
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Guardar"
      Height          =   495
      Left            =   8520
      TabIndex        =   23
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Personales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.TextBox txtdia19 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   54
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox Txtmes18 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   52
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox Txtaño17 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   50
         Top             =   2880
         Width           =   375
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   48
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton CmdRFC 
         Caption         =   "rfc"
         Height          =   195
         Left            =   5160
         TabIndex        =   47
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   28
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   2280
         Width           =   6735
      End
      Begin VB.TextBox Text10 
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
         Height          =   375
         Left            =   5880
         TabIndex        =   20
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   18
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   15
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   13
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         TabIndex        =   10
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label22 
         Caption         =   "Clave:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "DIA"
         Height          =   255
         Left            =   4680
         TabIndex        =   53
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "MES"
         Height          =   255
         Left            =   3840
         TabIndex        =   51
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label19 
         Caption         =   "&AÑO"
         Height          =   255
         Left            =   2760
         TabIndex        =   49
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "Observaciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "RFC:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Teléfono:"
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Código Postal:"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Delegación o Municipio:"
         Height          =   495
         Left            =   4080
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Ciudad:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "Lucida Sans Unicode"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Colonia:"
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Calle y Número:"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Apellido Materno:"
         Height          =   495
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Apellido Paterno:"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "&Nombre:"
         Height          =   255
         Left            =   5520
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label24 
      Caption         =   "Conversión"
      Height          =   255
      Left            =   8760
      TabIndex        =   64
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label23 
      Caption         =   "Copias:"
      Height          =   255
      Left            =   8520
      TabIndex        =   62
      Top             =   5760
      Width           =   615
   End
End
Attribute VB_Name = "FrmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim d As Integer
Dim RecClientes As Variant
Dim RecUnidades As Variant
Dim recproductos As Variant

Dim StrQuery As String
Dim ventatotal As Double
    Dim metrostot As Double
    Dim ventatotaliva As Double
    Dim nuevo As Boolean
    Dim iva As Double
    Dim ventagrantotiva As Integer
Dim totcant As Double
Dim totcant1 As Double
Dim totcant2 As Long
Dim L As String
Dim Numcopias As Integer
Dim metroscub As Double
Private Sub CmdAgregar_Click()

Dim x As Integer
Dim y As Integer
Dim i As Integer

If Text15.Text <> "" And Text16.Text <> "" Then
    With MSHventa
     
        
'    If .Row > 0 Then
'       X = .Row
'       .Row = X
'    Else
'       .Row = 1
'    End If
     .Row = 1
     .Col = 1
     For i = 1 To .Rows
         
         If .Text = "" Then
            x = i
            i = .Rows
        Else
            If .Row = 9 Then
               Exit Sub
            Else
            .Row = i + 1
            x = i
            End If
        End If
     Next i
      .Row = x
      
      .Col = 1
      ''.Text = Mid$(Combo1.Text, 4)
      .Text = Mid$(Combo1.Text, InStr(Combo1.Text, ":") + 1)
      .Col = 2
      .Text = Text17.Text
      .Col = 3
      .Text = Text13.Text
      .Col = 4
      .Text = Text14.Text
      .Col = 5
      .Text = Text15.Text
      .Col = 6
      .Text = Text16.Text
      .Col = 7
      .Text = Text21.Text
      .Col = 8
      .Text = metroscub
      End With
      
Else
      Beep
End If
   Text14.Text = ""
   Text15.Text = ""
   Text16.Text = ""
   
End Sub

'Private Sub CmdAgregar2_Click()
'Dim x As Integer
'Dim y As Integer
'y = 1
'If Text15.Text <> "" And Text16.Text <> "" Then
'   With SpVenta
'     If .Row > 0 Then
'        x = .Row + 1
'        .Row = x
'
'    Else
'
'      .Row = 1
'    End If
'      .Col = 1
'      .Text = Combo1.Text
'      .Col = 2
'      .Text = Combo2.Text
'      .Col = 3
''      .Text = Text13.Text
  '    .Col = 4
 ''     .Text = Text14.Text
  '    .Col = 5
  '    .Text = Text15.Text
  '    .Col = 6
  '    .Text = Text16.Text
  '
  '    End With
  ' Else
  ' Beep
  ' End If
'End Sub

Private Sub CmdEliminar_Click()
    Dim resp As Boolean
    If Text12.Text <> "" Then
       resp = MsgBox("¿Desea eliminar el registro?", vbQuestion + vbYesNo, "Aviso-MC") = vbYes
       
    If resp Then
    Dim admin As ClsAdmin
    
    Set admin = New ClsAdmin
        admin.BajaCliente CInt(Text12.Text)
    Set admin = Nothing
    MsgBox "Se Elimino un registro", vbOKOnly + vbInformation, "Aviso-MC"
    End If
    End If
End Sub

Private Sub CmdGuardar_Click()
Dim resp As Boolean
On Error GoTo Guardar_err:
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" _
       Or Text4.Text = "" Or Text5.Text = "" Or Text10.Text = "" Then
       MsgBox "Favor de llenar el Nombre completo,y dirección así como RFC", vbOKOnly + vbExclamation, "Aviso MC"
       Exit Sub
    End If
    Dim admin As ClsAdmin
    
    Set admin = New ClsAdmin
        If Text12.Text <> "" Then
        ''If modific Then
           resp = MsgBox("Desea modificar el registro", vbYesNo + vbQuestion, "aviso") = vbYes
           If resp Then
           admin.Modificacion Text12.Text, Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, _
              Text6.Text, Text7.Text, Text8.Text, Text9.Text, Text11.Text, Text10.Text
           Else
           Exit Sub
           End If
        Else
        admin.Alta_Movimiento Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, _
              Text6.Text, Text7.Text, Text8.Text, Text9.Text, Text11.Text, Text10.Text
        End If
    Set admin = Nothing
    
    MsgBox "Se añadio un cliente a la base de datos", vbOKOnly + vbInformation, "Aviso-Clientes"
    
Exit Sub

Guardar_err:
Set admin = Nothing
MsgBox "No se guardo", vbOKOnly + vbInformation, "aviso MC"
End Sub

Private Sub CmdNuevo_Click()
   Combo1.Text = ""
   Text1.Text = ""
   Text2.Text = ""
   Text3.Text = ""
   Text4.Text = ""
   Text5.Text = ""
   Text6.Text = ""
   Text7.Text = ""
   Text8.Text = ""
   Text9.Text = ""
   Text10.Text = ""
   Text11.Text = ""
   Text12.Text = ""
   Text13.Text = ""
   Text14.Text = ""
   Text15.Text = ""
   Text16.Text = ""
   Text17.Text = ""
   Txtaño17.Text = ""
   Txtmes18.Text = ""
   txtdia19.Text = ""
 End Sub

Private Sub CmdRegresar_Click()
'FrmAntesala.Visible = True
Unload Me
End Sub

Private Sub CmdBuscar_Click()
    Dim admin As ClsAdmin
    
   Set admin = New ClsAdmin
       admin.RecuperarClientes RecClientes
   Set admin = Nothing
   If RecClientes.RecordCount > 0 And Text2.Text <> "" Then
      RecClientes.Filter = "ApellidoPat like '%" & Trim$(Text2.Text & "%") & "'"
      ''RecMovimientos.Filter = "Opcion Like '%" & Trim$(Text1.Text & "%") & "'"
      If RecClientes.Filter > 0 Then
         MsgBox "El cliente ya existe", vbOKOnly + vbExclamation, "Aviso MC"
     
     End If
   End If
End Sub

Private Sub CmdMostrar_Click()
  Form2.Show 1
End Sub

Private Sub CmdRFC_Click()
    If Text2.Text <> "" And Text1.Text <> "" And Text3.Text <> "" And Txtaño17.Text <> "" And Txtmes18.Text <> "" And txtdia19.Text <> "" Then
       If CInt(Txtmes18.Text) > 12 Then
          MsgBox "El mes de nacimiento no puede ser mayor de 12", vbInformation + vbOKOnly, "Aviso-MC"
          Txtmes18.Text = ""
       Else
          If txtdia19.Text > 31 Then
             MsgBox "El día de nacimiento no puede ser mayor de 31", vbInformation + vbOKOnly, "Aviso-MC"
             txtdia19.Text = ""
          Else
             If Len(Txtaño17.Text) = 1 Then
                Txtaño17.Text = 0 & Txtaño17.Text
             End If
             If Len(txtdia19.Text) = 1 Then
                txtdia19.Text = 0 & txtdia19.Text
             End If
             If Len(Txtmes18.Text) = 1 Then
                Txtmes18.Text = 0 & Txtmes18.Text
             End If
             Text10.Text = UCase$(Mid$(Text2.Text, 1, 2)) & UCase$(Mid$(Text3.Text, 1, 1)) & UCase$(Mid$(Text1.Text, 1, 1)) & Mid$(Txtaño17.Text, 1, 2) & Mid$(Txtmes18.Text, 1, 2) & Mid$(txtdia19.Text, 1, 2)
          End If
       End If
    Else
       MsgBox "Favor de introducir la Fecha de Nacimiento y el nombre del cliente", vbOKOnly + vbInformation, "Aviso-PROA2Mil"
    End If
End Sub
 
Private Sub CmdTotal_Click()
    Dim t As Integer
    Dim totalSin1 As Double, totalsin4 As Double, totalsin5 As Double
    Dim totalsin2 As Double, totalsin6 As Double, totalsin7 As Double
    Dim totalsin3 As Double, totalsin8 As Double, totalsin9 As Double
    Dim con1 As Double, con4 As Double, con5 As Double
    Dim con2 As Double, con6 As Double, con7 As Double
    Dim con3 As Double, con8 As Double, con9 As Double
    Dim metros1 As Double, metros4 As Double, metros5 As Double
    Dim metros2 As Double, metros6 As Double, metros7 As Double
    Dim metros3 As Double, metros8 As Double, metros9 As Double
    Dim dig As Integer
    Dim arreglo() As Variant
    Dim numelem As Integer
    On Error GoTo Impre_Err
    
    With MSHventa
                   
        For t = 1 To 9
        
        .Row = t
    
        .Col = 5
       If .Row = 1 And .Text <> "" Then
          totalSin1 = .Text
       End If
       If .Row = 2 And .Text <> "" Then
          totalsin2 = .Text
       End If
       If .Row = 3 And .Text <> "" Then
          totalsin3 = .Text
       End If
       If .Row = 4 And .Text <> "" Then
          totalsin4 = .Text
       End If
       If .Row = 5 And .Text <> "" Then
          totalsin5 = .Text
       End If
       If .Row = 6 And .Text <> "" Then
          totalsin6 = .Text
       End If
       If .Row = 7 And .Text <> "" Then
          totalsin7 = .Text
       End If
       If .Row = 8 And .Text <> "" Then
          totalsin8 = .Text
       End If
       If .Row = 9 And .Text <> "" Then
          totalsin9 = .Text
       End If
       
       .Col = 6
       If .Row = 1 And .Text <> "" Then
           con1 = .Text
       End If
       If .Row = 2 And .Text <> "" Then
          con2 = .Text
       End If
       If .Row = 3 And .Text <> "" Then
          con3 = .Text
       End If
       If .Row = 4 And .Text <> "" Then
           con4 = .Text
       End If
        If .Row = 5 And .Text <> "" Then
           con5 = .Text
       End If
       If .Row = 6 And .Text <> "" Then
           con6 = .Text
       End If
       If .Row = 7 And .Text <> "" Then
           con7 = .Text
       End If
       If .Row = 8 And .Text <> "" Then
           con8 = .Text
       End If
       If .Row = 9 And .Text <> "" Then
           con9 = .Text
       End If

    .Col = 8
       If .Row = 1 And .Text <> "" Then
           metros1 = .Text
       End If
       If .Row = 2 And .Text <> "" Then
          metros2 = .Text
       End If
       If .Row = 3 And .Text <> "" Then
          metros3 = .Text
       End If
       If .Row = 4 And .Text <> "" Then
           metros4 = .Text
       End If
        If .Row = 5 And .Text <> "" Then
           metros5 = .Text
       End If
       If .Row = 6 And .Text <> "" Then
           metros6 = .Text
       End If
       If .Row = 7 And .Text <> "" Then
           metros7 = .Text
       End If
       If .Row = 8 And .Text <> "" Then
           metros8 = .Text
       End If
       If .Row = 9 And .Text <> "" Then
           metros9 = .Text
       End If
               
       ' If nuevo Then
         '  ventagrantot = ventatotal = .Text
        'End If
        '.Col = 6
       ' ventatotaliva = .Text
       ' If t <> .Rows Then
       '   nuevo = True
       ' Else
       '  nuevo = False
       ' End If
        Next t
        metrostot = CDbl(metros1) + CDbl(metros2) + CDbl(metros3) + CDbl(metros4) + CDbl(metros5) + CDbl(metros6) + CDbl(metros7) + CDbl(metros8) + CDbl(metros9)
        ventatotal = CDbl(totalSin1) + CDbl(totalsin2) + CDbl(totalsin3) + CDbl(totalsin4) + CDbl(totalsin5) + CDbl(totalsin6) + CDbl(totalsin7) + CDbl(totalsin8) + CDbl(totalsin9)
        ventatotaliva = CDbl(con1) + CDbl(con2) + CDbl(con3) + CDbl(con4) + CDbl(con5) + CDbl(con6) + CDbl(con7) + CDbl(con8) + CDbl(con9)
        If ventatotaliva - ventatotal <> 0 Then
           iva = Format(ventatotaliva - ventatotal, "#########.#####")
        End If
        
       If ventatotal = 0 Then
          MsgBox "No se puede generar una factura con Cero pesos", vbOKOnly, "Aviso-MC"
          Exit Sub
       End If
       LetrasNumero ventatotaliva, L
        
       
   End With
   
   If MsgBox("El total sin iva es: " & ventatotal & " pesos. " & Chr(13) & "El total con iva es: " & ventatotaliva & " pesos. " & Chr(13) & _
      Chr(13) & "¿Desea imprimir la factura?", vbYesNo + vbInformation, "RESULTADO TOTAL ") = vbYes Then
      If Text20.Text = "" Then Text20.Text = "3"
      Numcopias = CInt(Text20.Text)
      For d = 0 To Numcopias - 1
          ImprimeFactura
          Printer.NewPage
      Next d
   
   End If
   If MsgBox("¿Desea que se guarde la venta?", vbYesNo + vbQuestion, "Aviso-MC") = vbYes Then
        If Text12.Text <> "" Then
           LeeMSHVenta arreglo, numelem
        Else
           MsgBox "Para guardar la venta se necesita el cliente", vbOKOnly + vbInformation, "Aviso MC"
           Exit Sub
        End If
         Dim admin As ClsAdmin
         
         Set admin = New ClsAdmin
             ''admin.ALta_Venta 1, 1, 1, 1, Text12.Text, Format$(Date, "dd/MM/yyyy"), 1, arreglo, numelem
              admin.ALta_Venta Text12.Text & ":" & " " & Text1.Text & " " & Text2.Text, Date, arreglo, numelem
         Set admin = Nothing
    '  End If
   End If
    'With SpVenta
    '.Row = 1
    '.Col = 5
    'totalSin1 = .Text
    '.Row = 2
    '.Text = totalsin2
    
   ' End With
Exit Sub
Impre_Err:
If Err.Number = 482 Then
MsgBox "No existe una impresora, por favor cheque sus conexiones", vbOKOnly + vbExclamation, "Aviso - MC"
End If
Resume Next
End Sub

Private Sub Combo1_Click()
    
    
    Combo2.Clear
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text21.Text = ""
    If Combo1.Text <> "" Then
        Dim admin As ClsAdmin
        
        Set admin = New ClsAdmin
            
            ''            admin.RecuperarUnidades  IDCombo(Combo1.Text), RecUnidades
            
        Set admin = Nothing
        ''If RecUnidades.RecordCount > 0 Then
                
          ''      RecUnidades.MoveFirst
            ''      Combo2.Clear
                  ''Combo2.AddItem ""
        ''RecUnidades.Filter = "IdUNIDAD = " & uni & ""
        recproductos.Filter = "IDPRODUCTO = " & IDCombo(Combo1.Text) & ""
        If recproductos.RecordCount > 0 Then
           Text17.Text = recproductos(2)
           Text21.Text = recproductos(3)
           If Not IsNull(recproductos(5)) Then
              Text13.Text = recproductos(5)
              
           Else
           MsgBox "No se ha definido el precio producto", vbExclamation + vbOKOnly, "Aviso - MC"
           End If
               'Do While Not RecUnidades.EOF
                '   Combo2.AddItem Trim$(RecUnidades(0)) & ":" & "" & (RecUnidades(1))
                '  RecUnidades.MoveNext
               'Loop
        Else
           MsgBox "No se ha definido la unidad del producto", vbExclamation + vbOKOnly, "Aviso - MC"
        
        End If
    End If
    
End Sub

Private Sub Combo2_Click()
   Dim uni As Integer
   Text13.Text = ""
   Text14.Text = ""
   Text15.Text = ""
   Text16.Text = ""
   If RecUnidades.RecordCount > 0 Then
      uni = IDCombo(Combo2.Text)
      RecUnidades.Filter = "IdUNIDAD = " & uni & ""
  '' RecMovimientos.Filter = "IdMovimiento = " & Trim$(Text1.Text) & ""
      If RecUnidades.RecordCount > 0 Then
      Text13.Text = RecUnidades(5)
      Else
      Exit Sub
      End If
   End If
   
End Sub

Private Sub CmdBP_Click()
    Combo1.Clear
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text21.Text = ""
    If Text17.Text <> "" Then
        Dim admin As ClsAdmin
        
       Set admin = New ClsAdmin
           admin.Recuperar recproductos
       Set admin = Nothing
                  
          If recproductos.RecordCount > 0 Then
            '' recproductos.Filter = "IDPRODUCTO = " & IDCombo(Combo1.Text) & ""
             recproductos.Filter = "UnidadMedida = '" & Trim(Text17.Text) & "' "
                If recproductos.RecordCount > 0 Then
                   'Text17.Text = recproductos(2)
                   Combo1.Text = recproductos(0) & ":" & recproductos(1)
                   Text21.Text = recproductos(3)
                   If Not IsNull(recproductos(5)) Then
                      Text13.Text = recproductos(5)
                      
                   Else
                   MsgBox "No se ha definido el precio producto", vbExclamation + vbOKOnly, "Aviso - MC"
                   End If
                       'Do While Not RecUnidades.EOF
                        '   Combo2.AddItem Trim$(RecUnidades(0)) & ":" & "" & (RecUnidades(1))
                        '  RecUnidades.MoveNext
                       'Loop
                Else
                   MsgBox "No se encontro el producto", vbExclamation + vbOKOnly, "Aviso - MC"
                
                End If
           End If
    End If
End Sub

Private Sub Command1_Click()
Dim admin As ClsAdmin

  Set admin = New ClsAdmin
        admin.Recuperar recproductos
        
    Set admin = Nothing
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

Private Sub Form_Load()
   
   MSHventa.Row = 0
   MSHventa.Col = 1
   MSHventa.Text = "Producto"
   MSHventa.Col = 2
   MSHventa.Text = "NombreCorto"
   MSHventa.Col = 3
   MSHventa.Text = "PrecioUnitario"
   MSHventa.Col = 4
   MSHventa.Text = "NumUnidades"
   MSHventa.Col = 5
   MSHventa.Text = "Precio"
   MSHventa.Col = 6
   MSHventa.Text = "PrecioIVA"
   MSHventa.Col = 7
   MSHventa.Text = "Conversión"
   MSHventa.Col = 8
   MSHventa.Text = "Metros Cubicos"
   
   MSHventa.Row = 1
   MSHventa.Col = 0
   MSHventa.Text = "1"
   MSHventa.Row = 2
   MSHventa.Text = "2"
   MSHventa.Row = 3
   MSHventa.Text = "3"
   MSHventa.Row = 4
   MSHventa.Text = "4"
   MSHventa.Row = 5
   MSHventa.Text = "5"
   MSHventa.Row = 6
   MSHventa.Text = "6"
   MSHventa.Row = 7
   MSHventa.Text = "7"
   MSHventa.Row = 8
   MSHventa.Text = "8"
   MSHventa.Row = 9
   MSHventa.Text = "9"
   
   Dim admin As ClsAdmin
   
    Set admin = New ClsAdmin
        admin.Recuperar recproductos
        
    Set admin = Nothing
    If recproductos.RecordCount > 0 Then
            
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

Private Sub MSHventa_Click()
 With MSHventa
      
      .Col = 1
      .Text = ""
      .Col = 2
      .Text = ""
      .Col = 3
      .Text = ""
      .Col = 4
      .Text = ""
      .Col = 5
      .Text = ""
      .Col = 6
      .Text = ""
      ''.Row = 0
 End With
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Dim admin As ClsAdmin
   
   Set admin = New ClsAdmin
       admin.RecuperarClientes RecClientes
   Set admin = Nothing
   If RecClientes.RecordCount > 0 Then
      RecClientes.Filter = "ApellidoPat like '%" & Trim$(Text1.Text & "%") & "'"
      ''RecMovimientos.Filter = "Opcion Like '%" & Trim$(Text1.Text & "%") & "'"
      If RecClientes.Filter > 0 Then
         MsgBox "se encontro", vbOKOnly, "aviso"
     End If
   End If
Else
End If
   
End Sub


Private Sub Text13_KeyPress(KeyAscii As Integer)
        If KeyAscii = 46 Then Exit Sub
        EsEntero Text13, 14, KeyAscii
End Sub

Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 13 And Text13.Text <> "" Then
   
   totcant1 = Trim$(Text13.Text)
   totcant2 = Trim$(Text14.Text)
   totcant = totcant1 * totcant2
   metroscub = Trim(CDbl(Text21.Text)) * m3 * totcant2
   'Text15.Text = CInt(Text13.Text) * CInt(Text14.Text)
   'Text16.Text = (CInt(Text15.Text) * 0.15) + (CInt(Text15.Text))
   Text15.Text = totcant
   Text16.Text = (totcant * 0.15) + totcant
End If
   
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
       EsEntero Text13, 14, KeyAscii
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
     If KeyAscii = 46 Then Exit Sub
     EsEntero Text15, 13, KeyAscii
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then Exit Sub
    EsEntero Text16, 13, KeyAscii
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
         CmdBP_Click
      End If

End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
    EsEntero Text20, 1, KeyAscii
End Sub

Private Sub Txtaño17_KeyPress(KeyAscii As Integer)
  EsEntero Txtaño17, 2, KeyAscii
End Sub
Private Sub Txtmes18_KeyPress(KeyAscii As Integer)
  EsEntero Txtmes18, 2, KeyAscii
End Sub
Private Sub Txtdia19_KeyPress(KeyAscii As Integer)
  EsEntero txtdia19, 2, KeyAscii
End Sub
Private Sub ImprimeFactura()
Dim i As Integer
Dim y As Double
y = 0.4

                    Printer.ScaleMode = 7 'CMS
                    Printer.Orientation = 1
                    Printer.Font.Size = 12
                    Printer.FontBold = True
                    Printer.Font = "arial"
                    Printer.CurrentY = 4.5
                    Printer.CurrentX = 4
                    Printer.Print Text2.Text & " " & Text3.Text & " " & Text1.Text
                    Printer.FontSize = 12
                    Printer.CurrentY = 5
                    Printer.CurrentX = 4
                    Printer.Print Text4.Text & "  " & "Col." & Text5.Text & " " & Text6.Text & " " & Text7.Text & Text8.Text & " " & Text9.Text
                    Printer.FontSize = 12
                    Printer.CurrentY = 5.5
                    Printer.CurrentX = 4
                    Printer.Print Text10.Text
                    Printer.FontSize = 10
                    Printer.CurrentY = 5.5
                    Printer.CurrentX = 16
                    Printer.Print Format(Date, "dd/MM/yyyy")
                    
                    ''ROW 1
                    
                    Printer.CurrentY = 7.4
                    Printer.CurrentX = 1.5
                    
                    With MSHventa
                    ''Do While Not .Row = 10
                    ''For i = 1 To 9
                    .Row = 1
                    .Col = 4
                    If .Text <> "" Then    ''Unidaes
                       Printer.Print .Text
                    Else
                    Printer.FontSize = 7
                    Printer.CurrentY = 11.9
                    Printer.CurrentX = 1.5
                    Printer.Print
                    
                    
                    Printer.FontSize = 9
                    
                    'Printer.CurrentY = 11.9
                    'Printer.CurrentX = 17.5
                    'Printer.Print ventatotal
                    Printer.CurrentY = 12.5
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotal
                    Printer.CurrentY = 13
                    Printer.CurrentX = 17.5
                    Printer.Print iva
                    Printer.CurrentY = 13.5
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotaliva
                                             
                       Printer.EndDoc
                       Exit Sub
                    End If
                    Printer.CurrentY = 7.4
                    Printer.CurrentX = 3.5
                    .Col = 1
                    Printer.Print .Text     ''NombreProd
                    
                    Printer.CurrentY = 7.4
                    Printer.CurrentX = 13
                    .Col = 8
                    Printer.Print .Text
                    
                    ''Printer.Print .Text; space(
                    Printer.CurrentY = 7.4
                    Printer.CurrentX = 15
                    .Col = 3
                    Printer.Print .Text     ''PrecioUnit
                    Printer.CurrentY = 7.4
                    Printer.CurrentX = 17.5
                    .Col = 5
                    Printer.Print .Text     ''Preciosin iva
                    
                    ''''ROW 2
                    Printer.CurrentY = 7.9
                    Printer.CurrentX = 1.5
                    .Row = 2
                    .Col = 4
                    If .Text <> "" Then
                       Printer.Print .Text
                    Else
                       Printer.FontSize = 7
                       Printer.CurrentY = 12.4
                       Printer.CurrentX = 1
                       Printer.Print L           ''CANTIDAD EN LETRAS
                       
                       Printer.CurrentX = 13
                       Printer.CurrentY = 12.4
                       Printer.Print metrostot & " m3"
                       
                       Printer.FontSize = 9
                       'Printer.CurrentY = 11.9 + Y
                    'Printer.CurrentX = 17.5
                    'Printer.Print ventatotal     '' VENTA SIN IVA
                    Printer.CurrentY = 12.6 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotal      '' VENTA SIN IVA
                    Printer.CurrentY = 13.2 + y
                    Printer.CurrentX = 17.5
                    Printer.Print iva           '' PURO IVA DE TODO
                    Printer.CurrentY = 13.7 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotaliva   '' CANTIDAD A PAGAR
                       Printer.EndDoc
                       Exit Sub
                    End If
                    Printer.CurrentY = 7.9
                    Printer.CurrentX = 3.5
                    .Col = 1
                    Printer.Print .Text
                    
                    Printer.CurrentY = 7.9
                    Printer.CurrentX = 13
                    .Col = 8
                    Printer.Print .Text
                    
                    Printer.CurrentY = 7.9
                    Printer.CurrentX = 15
                    
                                       
                    .Col = 3
                    Printer.Print .Text
                    Printer.CurrentY = 7.9
                    Printer.CurrentX = 17.5
                    .Col = 5
                    Printer.Print .Text
                    
                    ''ROW3
                    Printer.CurrentY = 8.4
                    Printer.CurrentX = 1.5
                    .Row = 3
                    .Col = 4
                    If .Text <> "" Then
                       Printer.Print .Text
                    Else
                        Printer.FontSize = 7
                       Printer.CurrentY = 12.4
                       Printer.CurrentX = 1
                       Printer.Print L           ''CANTIDAD EN LETRAS
                       
                       Printer.CurrentX = 13
                       Printer.CurrentY = 12.4
                       Printer.Print metrostot & " m3"
                       
                       Printer.FontSize = 9
                       'Printer.CurrentY = 11.9 + Y
                    'Printer.CurrentX = 17.5
                    'Printer.Print ventatotal     '' VENTA SIN IVA
                    Printer.CurrentY = 12.6 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotal      '' VENTA SIN IVA
                    Printer.CurrentY = 13.2 + y
                    Printer.CurrentX = 17.5
                    Printer.Print iva           '' PURO IVA DE TODO
                    Printer.CurrentY = 13.7 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotaliva   '' CANTIDAD A PAGAR
                       Printer.EndDoc
                       Exit Sub
                    End If
                    Printer.CurrentY = 8.4
                    Printer.CurrentX = 3.5
                    .Col = 1
                    Printer.Print .Text
                    
                    Printer.CurrentY = 8.4
                    Printer.CurrentX = 13
                    .Col = 8
                    Printer.Print .Text
                    
                    Printer.CurrentY = 8.4
                    Printer.CurrentX = 15
                    .Col = 3
                    Printer.Print .Text
                    Printer.CurrentY = 8.4
                    Printer.CurrentX = 17.5
                    .Col = 5
                    Printer.Print .Text
                    
                    ''ROW 4
                    Printer.CurrentY = 8.9
                    Printer.CurrentX = 1.5
                    .Row = 4
                    .Col = 4
                    If .Text <> "" Then
                       Printer.Print .Text
                    Else
                      Printer.FontSize = 7
                       Printer.CurrentY = 12.4
                       Printer.CurrentX = 1
                       Printer.Print L           ''CANTIDAD EN LETRAS
                       
                       Printer.CurrentX = 13
                       Printer.CurrentY = 12.4
                       Printer.Print metrostot & " m3"
                       
                       
                       
                       Printer.FontSize = 9
                       'Printer.CurrentY = 11.9 + Y
                    'Printer.CurrentX = 17.5
                    'Printer.Print ventatotal     '' VENTA SIN IVA
                    Printer.CurrentY = 12.6 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotal      '' VENTA SIN IVA
                    Printer.CurrentY = 13.2 + y
                    Printer.CurrentX = 17.5
                    Printer.Print iva           '' PURO IVA DE TODO
                    Printer.CurrentY = 13.7 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotaliva   '' CANTIDAD A PAGAR
                       Printer.EndDoc
                       Exit Sub
                    End If
                    Printer.CurrentY = 8.9
                    Printer.CurrentX = 3.5
                    .Col = 1
                    Printer.Print .Text
                    
                    Printer.CurrentY = 8.9
                    Printer.CurrentX = 13
                    .Col = 8
                    Printer.Print .Text
                    
                    Printer.CurrentY = 8.9
                    Printer.CurrentX = 15
                    .Col = 3
                    Printer.Print .Text
                    Printer.CurrentY = 8.9
                    Printer.CurrentX = 17.5
                    .Col = 5
                    Printer.Print .Text
                    
                    ''ROW 5
                    Printer.CurrentY = 9.4
                    Printer.CurrentX = 1.5
                    .Row = 5
                    .Col = 4
                    If .Text <> "" Then
                       Printer.Print .Text
                    Else
                      Printer.FontSize = 7
                       Printer.CurrentY = 12.4
                       Printer.CurrentX = 1
                       Printer.Print L           ''CANTIDAD EN LETRAS
                       
                       Printer.CurrentX = 13
                       Printer.CurrentY = 12.4
                       Printer.Print metrostot & " m3"
                       
                       Printer.FontSize = 9
                       'Printer.CurrentY = 11.9 + Y
                    'Printer.CurrentX = 17.5
                    'Printer.Print ventatotal     '' VENTA SIN IVA
                    Printer.CurrentY = 12.6 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotal      '' VENTA SIN IVA
                    Printer.CurrentY = 13.2 + y
                    Printer.CurrentX = 17.5
                    Printer.Print iva           '' PURO IVA DE TODO
                    Printer.CurrentY = 13.7 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotaliva   '' CANTIDAD A PAGAR
                       Printer.EndDoc
                       Exit Sub
                    End If
                    Printer.CurrentY = 9.4
                    Printer.CurrentX = 3.5
                    .Col = 1
                    Printer.Print .Text
                    
                    Printer.CurrentY = 9.4
                    Printer.CurrentX = 13
                    .Col = 8
                    Printer.Print .Text
                    
                    Printer.CurrentY = 9.4
                    Printer.CurrentX = 15
                    .Col = 3
                    Printer.Print .Text
                    Printer.CurrentY = 9.4
                    Printer.CurrentX = 17.5
                    .Col = 5
                    Printer.Print .Text
                    
                    ''ROW 6
                    
                    Printer.CurrentY = 10
                    Printer.CurrentX = 1.5
                    .Row = 6
                    .Col = 4
                    If .Text <> "" Then
                       Printer.Print .Text
                    Else
                       Printer.FontSize = 7
                       Printer.CurrentY = 12.4
                       Printer.CurrentX = 1
                       Printer.Print L           ''CANTIDAD EN LETRAS
                       
                       Printer.CurrentX = 13
                       Printer.CurrentY = 12.4
                       Printer.Print metrostot & " m3"
                       
                       Printer.FontSize = 9
                       'Printer.CurrentY = 11.9 + Y
                    'Printer.CurrentX = 17.5
                    'Printer.Print ventatotal     '' VENTA SIN IVA
                    Printer.CurrentY = 12.6 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotal      '' VENTA SIN IVA
                    Printer.CurrentY = 13.2 + y
                    Printer.CurrentX = 17.5
                    Printer.Print iva           '' PURO IVA DE TODO
                    Printer.CurrentY = 13.7 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotaliva   '' CANTIDAD A PAGAR
                       Printer.EndDoc
                       Exit Sub
                    End If
                    Printer.CurrentY = 10
                    Printer.CurrentX = 3.5
                    .Col = 1
                    Printer.Print .Text
                    
                    Printer.CurrentY = 10
                    Printer.CurrentX = 13
                    .Col = 8
                    Printer.Print .Text
                    
                    
                    Printer.CurrentY = 10
                    Printer.CurrentX = 15
                    .Col = 3
                    Printer.Print .Text
                    Printer.CurrentY = 10
                    Printer.CurrentX = 17.5
                    .Col = 5
                    Printer.Print .Text
                    
                    ''ROW 7
                    
                    Printer.CurrentY = 10.6
                    Printer.CurrentX = 1.5
                    .Row = 7
                    .Col = 4
                    If .Text <> "" Then
                       Printer.Print .Text
                    Else
                        Printer.FontSize = 7
                       Printer.CurrentY = 12.4
                       Printer.CurrentX = 1
                       Printer.Print L           ''CANTIDAD EN LETRAS
                       
                       Printer.CurrentX = 13
                       Printer.CurrentY = 12.4
                       Printer.Print metrostot & " m3"
                       
                       
                       Printer.FontSize = 9
                       'Printer.CurrentY = 11.9 + Y
                    'Printer.CurrentX = 17.5
                    'Printer.Print ventatotal     '' VENTA SIN IVA
                    Printer.CurrentY = 12.6 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotal      '' VENTA SIN IVA
                    Printer.CurrentY = 13.2 + y
                    Printer.CurrentX = 17.5
                    Printer.Print iva           '' PURO IVA DE TODO
                    Printer.CurrentY = 13.7 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotaliva   '' CANTIDAD A PAGAR
                       Printer.EndDoc
                       Exit Sub
                    End If
                    Printer.CurrentY = 10.6
                    Printer.CurrentX = 3.5
                    .Col = 1
                    Printer.Print .Text
                    
                    Printer.CurrentY = 10.6
                    Printer.CurrentX = 13
                    .Col = 8
                    Printer.Print .Text
                    
                    Printer.CurrentY = 10.6
                    Printer.CurrentX = 15
                    .Col = 3
                    Printer.Print .Text
                    Printer.CurrentY = 10.6
                    Printer.CurrentX = 17.5
                    .Col = 5
                    Printer.Print .Text
                    
                    ''ROW 8
                    
                    Printer.CurrentY = 11.1
                    Printer.CurrentX = 1.5
                    .Row = 8
                    .Col = 4
                    If .Text <> "" Then
                       Printer.Print .Text
                    Else
                     Printer.FontSize = 7
                       Printer.CurrentY = 12.4
                       Printer.CurrentX = 1
                       Printer.Print L           ''CANTIDAD EN LETRAS
                       
                       Printer.CurrentX = 13
                       Printer.CurrentY = 12.4
                       Printer.Print metrostot & " m3"
                       
                       Printer.FontSize = 9
                       'Printer.CurrentY = 11.9 + Y
                    'Printer.CurrentX = 17.5
                    'Printer.Print ventatotal     '' VENTA SIN IVA
                    Printer.CurrentY = 12.6 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotal      '' VENTA SIN IVA
                    Printer.CurrentY = 13.2 + y
                    Printer.CurrentX = 17.5
                    Printer.Print iva           '' PURO IVA DE TODO
                    Printer.CurrentY = 13.7 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotaliva   '' CANTIDAD A PAGAR
                       Printer.EndDoc
                       Exit Sub
                    End If
                    Printer.CurrentY = 11.1
                    Printer.CurrentX = 3.5
                    .Col = 1
                    Printer.Print .Text
                    
                    Printer.CurrentY = 11.1
                    Printer.CurrentX = 13
                    .Col = 8
                    Printer.Print .Text
                    
                    
                    Printer.CurrentY = 11.1
                    Printer.CurrentX = 15
                    .Col = 3
                    Printer.Print .Text
                    Printer.CurrentY = 11.1
                    Printer.CurrentX = 17.5
                    .Col = 5
                    Printer.Print .Text
                    
                    ''ROW 9
                    
                    Printer.CurrentY = 11.65
                    Printer.CurrentX = 1.5
                    .Row = 9
                    .Col = 4
                    If .Text <> "" Then
                       Printer.Print .Text
                    Else
                     Printer.FontSize = 7
                       Printer.CurrentY = 12.4
                       Printer.CurrentX = 1
                       Printer.Print L           ''CANTIDAD EN LETRAS
                       
                       Printer.CurrentX = 13
                       Printer.CurrentY = 12.4
                       Printer.Print metrostot & " m3"
                       
                       Printer.FontSize = 9
                       'Printer.CurrentY = 11.9 + Y
                    'Printer.CurrentX = 17.5
                    'Printer.Print ventatotal     '' VENTA SIN IVA
                    Printer.CurrentY = 12.6 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotal      '' VENTA SIN IVA
                    Printer.CurrentY = 13.2 + y
                    Printer.CurrentX = 17.5
                    Printer.Print iva           '' PURO IVA DE TODO
                    Printer.CurrentY = 13.7 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotaliva   '' CANTIDAD A PAGAR
                       Printer.EndDoc
                       Exit Sub
                    End If
                    Printer.CurrentY = 11.65
                    Printer.CurrentX = 3.5
                    .Col = 1
                    Printer.Print .Text
                    
                    Printer.CurrentY = 11.65
                    Printer.CurrentX = 13
                    .Col = 8
                    Printer.Print .Text
                    
                    
                    Printer.CurrentY = 11.65
                    Printer.CurrentX = 15
                    .Col = 3
                    Printer.Print .Text
                    Printer.CurrentY = 11.65
                    Printer.CurrentX = 17.5
                    .Col = 5
                    Printer.Print .Text
                    
                    End With
                    
                    
                     Printer.FontSize = 7
                       Printer.CurrentY = 12.4
                       Printer.CurrentX = 1
                       Printer.Print L                       ''CANTIDAD EN LETRAS
                       
                       Printer.CurrentX = 13
                       Printer.CurrentY = 12.4
                       Printer.Print metrostot & " m3"
                        
                       Printer.FontSize = 9
                    '   Printer.CurrentY = 11.9 + Y
                    'Printer.CurrentX = 17.5
                    'Printer.Print ventatotal     '' VENTA SIN IVA
                    Printer.CurrentY = 12.6 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotal      '' VENTA SIN IVA
                    Printer.CurrentY = 13.2 + y
                    Printer.CurrentX = 17.5
                    Printer.Print iva           '' PURO IVA DE TODO
                    Printer.CurrentY = 13.7 + y
                    Printer.CurrentX = 17.5
                    Printer.Print ventatotaliva   '' CANTIDAD A PAGAR
                    
                    
                    
                    Printer.EndDoc
                    
End Sub
Private Sub LeeMSHVenta(arr() As Variant, NumElementos As Integer)
    Dim Lee     As Boolean
    Dim cx      As Long
    Dim cy      As Long
    Dim y       As Integer
    Dim x       As Integer
    
   On Error GoTo Error_Handler
    NumElementos = 0
    'cy = Spread.MaxRows
     ''cy = MSHventa.Rows
       cy = 9
    ''cx = Spread.MaxCols
     cx = 8
  ReDim arr(cy, cx)
      
  For y = 0 To cy - 1
      MSHventa.Row = y + 1
      MSHventa.Col = 1
     If MSHventa.Text <> "" Then
        Lee = True
     Else
        Lee = False
     End If
     If Lee Then
          For x = 0 To cx - 1
          
             MSHventa.Row = y + 1
             MSHventa.Col = x + 1
                If Lee Then
                   arr(NumElementos, x) = MSHventa.Text
                End If
          Next
      End If
      If Lee Then
           NumElementos = NumElementos + 1
      End If
  Next
Exit Sub
Error_Handler:
  ''Manejador_Errores Err.Number, sblab
End Sub
