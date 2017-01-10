VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   12
      SelectionMode   =   1
      AllowUserResizing=   1
      MousePointer    =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   MSHFlexGrid1.Row = 0
   MSHFlexGrid1.Col = 1
   MSHFlexGrid1.Text = "Nombre"
   MSHFlexGrid1.Col = 2
   MSHFlexGrid1.Text = "ApellidoPat"
   MSHFlexGrid1.Col = 3
   MSHFlexGrid1.Text = "ApellidoMat"
   MSHFlexGrid1.Col = 4
   MSHFlexGrid1.Text = "Calle"
   MSHFlexGrid1.Col = 5
   MSHFlexGrid1.Text = "Colonia"
   MSHFlexGrid1.Col = 6
   MSHFlexGrid1.Text = "Ciudad"
   MSHFlexGrid1.Col = 7
   MSHFlexGrid1.Text = "Deleg/Municipio"
   MSHFlexGrid1.Col = 8
   MSHFlexGrid1.Text = "C.P."
   MSHFlexGrid1.Col = 9
   MSHFlexGrid1.Text = "Teléfono"
   MSHFlexGrid1.Col = 10
   MSHFlexGrid1.Text = "Observaciones"
   MSHFlexGrid1.Col = 11
   MSHFlexGrid1.Text = "RFC"
   Set admin = New ClsAdmin
       admin.RecuperarClientes RecClientes
   Set admin = Nothing
   If RecClientes.RecordCount > 0 Then
      With MSHFlexGrid1
        For i = 1 To RecClientes.RecordCount
            If RecClientes.BOF = True Or RecClientes.EOF = True Then
               Exit Sub
            End If
        .AddItem (Row)
        .Row = i
          .Col = 0
               .Text = (RecClientes(0))
            
          .Col = 1
               .Text = (RecClientes(1))
          .Col = 2
               .Text = (RecClientes(2))
          .Col = 3
               .Text = (RecClientes(3))
          .Col = 4
               .Text = (RecClientes(4))
          .Col = 5
               .Text = (RecClientes(5))
          .Col = 6
               If RecClientes(6) <> "" Then
                  .Text = (RecClientes(6))
               End If
                  
          .Col = 7
               If RecClientes(7) <> "" Then
               .Text = (RecClientes(7))
               End If
          .Col = 8
               If RecClientes(8) <> "" Then
               .Text = (RecClientes(8))
               End If
          .Col = 9
               If RecClientes(9) <> "" Then
               .Text = (RecClientes(9))
               End If
          .Col = 10
               If RecClientes(10) <> "" Then
               .Text = (RecClientes(10))
               End If
          .Col = 11
               .Text = (RecClientes(11))
          RecClientes.MoveNext
        Next i
      End With
   End If
End Sub

Private Sub MSHFlexGrid1_DblClick()
    MSHFlexGrid1.Col = 0
    FrmClientes.Text12.Text = MSHFlexGrid1.Text
    MSHFlexGrid1.Col = 1
    FrmClientes.Text1.Text = MSHFlexGrid1.Text
    MSHFlexGrid1.Col = 2
    FrmClientes.Text2.Text = MSHFlexGrid1.Text
    MSHFlexGrid1.Col = 3
    FrmClientes.Text3.Text = MSHFlexGrid1.Text
    MSHFlexGrid1.Col = 4
    FrmClientes.Text4.Text = MSHFlexGrid1.Text
    MSHFlexGrid1.Col = 5
    FrmClientes.Text5.Text = MSHFlexGrid1.Text
    MSHFlexGrid1.Col = 6
    FrmClientes.Text6.Text = MSHFlexGrid1.Text
    MSHFlexGrid1.Col = 7
    FrmClientes.Text7.Text = MSHFlexGrid1.Text
    MSHFlexGrid1.Col = 8
    FrmClientes.Text8.Text = MSHFlexGrid1.Text
    MSHFlexGrid1.Col = 9
    FrmClientes.Text9.Text = MSHFlexGrid1.Text
    MSHFlexGrid1.Col = 10
    FrmClientes.Text11.Text = MSHFlexGrid1.Text
    MSHFlexGrid1.Col = 11
    FrmClientes.Text10.Text = MSHFlexGrid1.Text
    FrmClientes.Visible = True
    modific = True
    Unload Me
End Sub
