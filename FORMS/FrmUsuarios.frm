VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   7695
   Begin VB.Frame Frame1 
      Caption         =   "Accesos"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   1800
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H0080FF80&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         TabIndex        =   10
         Top             =   3240
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Eliminar"
         Height          =   495
         Index           =   3
         Left            =   5640
         TabIndex        =   9
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         Height          =   495
         Index           =   2
         Left            =   5640
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   495
         Index           =   1
         Left            =   5640
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nuevo"
         Height          =   495
         Index           =   0
         Left            =   5640
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxtCU 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox TxtNU 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   840
         Width           =   2775
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   3000
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   873
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label Label5 
         Caption         =   "Confirmar Contraseña:"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Usuarios:"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Clave:"
         Height          =   255
         Left            =   5880
         TabIndex        =   11
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Contraseña:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre Usuario:"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RecUsuarios As Variant
Dim RecUsuariounico As Variant
Private Sub Command2_Click()

End Sub

Private Sub Combo1_Click()
       If Combo1.Text = "" Then Exit Sub
       Dim admincls As ClsAdmin
       Set admincls = New ClsAdmin
           admincls.RecuperarUsuarios RecUsuarios, RecUsuariounico, Trim$(Combo1.Text), True
       Set admincls = Nothing
       If RecUsuariounico.RecordCount > 0 Then
          TxtNU = RecUsuariounico(1)
          TxtCU = RecUsuariounico(2)
          Text1 = RecUsuariounico(0)
       End If
End Sub

Private Sub Command1_Click(Index As Integer)

     Select Case Index
            Case 0
                 Text2.Text = ""
                 Text1.Text = ""
                 TxtNU.Text = ""
                 TxtCU.Text = ""
                 Combo1.Clear
                 
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
            Case 1
                 If TxtNU.Text <> "" And TxtCU.Text <> "" Then
                    Set admincls = New ClsAdmin
                        If Text1.Text <> "" Then
                           If TxtCU.Text = Text2.Text Then
                              admincls.Alta_Usuario Trim$(TxtNU.Text), Trim$(TxtCU.Text), , CInt(Text1.Text)
                           Else
                              MsgBox "Las contraseñas no son correctas", vbOKOnly + vbExclamation, "Aviso - MC"
                              Exit Sub
                           End If
                            
                        Else
                           If TxtCU.Text = Text2.Text Then
                           admincls.Alta_Usuario Trim$(TxtNU.Text), Trim$(TxtCU.Text)
                           Else
                           MsgBox "Las contraseñas no son correctas", vbOKOnly + vbExclamation, "Aviso - MC"
                           Exit Sub
                           End If
                        End If
                    Set admincls = Nothing
                 Else
                    MsgBox "Por Favor indique el usuario y contraseña", vbExclamation + vbOKOnly, "Aviso - MC"
                    Exit Sub
                 End If
                 MsgBox "Se guardo un registro en la base de datos", vbInformation + vbOKOnly, "Aviso-MC"
                 Command1_Click (0)
           Case 2
           Case 3
           If TxtNU.Text <> "" And TxtCU.Text <> "" And Text1.Text <> "" Then
              Set admincls = New ClsAdmin
                  admincls.BorraUsuario Text1.Text
              Set admincls = Nothing
              MsgBox "Se Elimino un registro", vbInformation + vbOKOnly, "Aviso-MC"
           End If
  End Select
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
   CenterForm Me
End Sub
