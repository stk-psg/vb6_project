VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3870
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5685
   ControlBox      =   0   'False
   DrawMode        =   15  'Merge Pen Not
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   3870
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   240
      Top             =   3240
   End
   Begin VB.Label Label4 
      Caption         =   "Systems..XXX/I/MM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Facturación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3675
      Left            =   120
      Picture         =   "Form11.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   " Cabreras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1125
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   " Madereria"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1125
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4830
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cuantos As Integer
Private Sub Explosion()
    Dim i As Integer, CodigoDeColor As Integer
    Dim Columna As Single, Fila As Single
    Randomize
    Scale (-320, 240)-(320, -240)
    For i = 1 To 75
        Columna = 320 * Rnd
        If Rnd < 0.5 Then Columna = -Columna
        Fila = 240 * Rnd
        If Rnd < 0.5 Then Fila = -Fila
        CodigoDeColor = 15 * Rnd
        Line (0, 0)-(Columna, Fila), QBColor(CodigoDeColor)
    Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    FrmLogin.Show
End Sub

Private Sub Timer1_Timer()
    Cuantos = Cuantos + 1
    If Cuantos > 70 And Cuantos < 450 Then Explosion
    If Cuantos = 450 Then
        Form1.AutoRedraw = True
        Image1.Visible = True
        Label4.Visible = True
        Label3.Visible = True
        Label2.Visible = True
        'Unload Me
        'Label1(0).BackStyle = 1
        'Label1(1).BackStyle = 1
        'Label1(0).BackColor = &HFFFFFF
        'Label1(1).BackColor = &HFFFFFF
    End If
    If Cuantos = 650 Then
        Timer1.Enabled = False
        Form1.BackColor = &HE0E0E0
        Unload Me
    End If
End Sub
