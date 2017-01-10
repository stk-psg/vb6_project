VERSION 5.00
Begin VB.Form FrmLogin 
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Contraseña"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton Command2 
         Caption         =   "&Salir"
         Height          =   615
         Left            =   4560
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Entrar"
         Height          =   615
         Left            =   4560
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   1545
         Left            =   6000
         Picture         =   "FrmLogin.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Clave:"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecUsuarios As Variant
Dim RecUsuariounico As Variant
Dim enter As Boolean

Private Sub Combo1_Click()
Text1.Text = ""
End Sub

Private Sub Command1_Click()
If Combo1.Text <> "" Then
    Dim admincls As ClsAdmin
    Set admincls = New ClsAdmin
        admincls.RecuperarUsuarios RecUsuarios, RecUsuariounico, Trim$(Combo1.Text), True
    Set admincls = Nothing
    If RecUsuariounico.RecordCount > 0 Then
       If Text1.Text = UCase(RecUsuariounico(2)) Then
          enter = True
           Unload Me
           Exit Sub
        Else
           MsgBox "La contraseña es incorrecta", vbInformation + vbOKOnly, "Aviso-MC"
           enter = False
           Exit Sub
       End If
    End If
    MsgBox "Usuario desconocido", vbExclamation + vbOKOnly, "Aviso-MC"
    enter = False
    
Else
   MsgBox "Indique el Nombre Por Favor", vbExclamation + vbOKOnly, "Aviso-MC"
   enter = False
End If

End Sub

Private Sub Command2_Click()
   End
End Sub

Private Sub Form_Load()
Dim admincls As ClsAdmin
  Set admincls = New ClsAdmin
       admincls.RecuperarUsuarios RecUsuarios, RecUsuariounico
   Set admincls = Nothing
   If RecUsuarios.RecordCount > 0 Then
      Do While Not RecUsuarios.EOF
             Combo1.AddItem RecUsuarios(1)
             RecUsuarios.MoveNext
      Loop
   End If
 ' Combo1.AddItem "Erón Cabrera"
 ' Combo1.AddItem "Usuario"
CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If enter = True Then
   MDIPrincipal.Show
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1_Click
   
   
Else

  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
