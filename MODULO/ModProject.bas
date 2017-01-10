Attribute VB_Name = "Module1"
Public modific As Boolean
Public todosprod As Boolean
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal aAction As Long, ByVal aParam As Long, R As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48
'Global Const DSN = "DSN=DBAccess;uid=;pwd="
Public Const m3 = 0.00236
Public Sub AsignaTexto(ByVal Inicio As Integer, ByVal Final As Integer, ByVal Reg As Variant, Forma As Form)
Dim i As Integer

    For i = Inicio To Final
         On Error Resume Next
              Forma.txtlab(i - 1) = Trim$(Reg(i - 1))
         On Error GoTo 0
    Next i

End Sub
Public Function IDCombo(ByVal strCombo As String) As String

    If Trim$(strCombo) <> "" Then
       IDCombo = Mid$(strCombo, 1, InStr(strCombo, ":") - 1)
    Else
       IDCombo = ""
    End If

End Function
Public Sub EsEntero(ByVal ctrlText As Control, ByVal longitud As Integer, ByRef KeyAscii As Integer)
    
    ' Rutina para validar valores enteros, se requiere pasar el nombre de la caja de
    ' texto que se quiere validar.
    ' El valor de KeyAscii debe ser por referencia.
    
    If Len(ctrlText.Text) < longitud Then
       If Not ValidaEntero(KeyAscii) Then KeyAscii = 0
    Else
       If KeyAscii <> 8 Then
          KeyAscii = 0
          Beep
       End If
    End If
    
End Sub
Public Function ValidaEntero(KeyAscii As Integer) As Boolean

    If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        ValidaEntero = True
    Else
        ValidaEntero = False
        Beep
    End If
 
End Function
Public Function LetraCantidad(Numero As Integer) As String
    Dim cantidad    As String
    cantidad = ""
    Select Case Numero
        Case 0: cantidad = ""
        Case 100: cantidad = "CIEN"
        Case Else
            Select Case Numero \ 100
                Case 1: cantidad = "CIENTO "
                Case 2: cantidad = "DOSCIENTOS "
                Case 3: cantidad = "TRESCIENTOS "
                Case 4: cantidad = "CUATROCIENTOS "
                Case 5: cantidad = "QUINIENTOS "
                Case 6: cantidad = "SEISCIENTOS "
                Case 7: cantidad = "SETECIENTOS "
                Case 8: cantidad = "OCHOCIENTOS "
                Case 9: cantidad = "NOVECIENTOS "
            End Select
            Select Case Numero Mod 100
                Case 11: cantidad = cantidad + "ONCE"
                Case 12: cantidad = cantidad + "DOCE"
                Case 13: cantidad = cantidad + "TRECE"
                Case 14: cantidad = cantidad + "CATORCE"
                Case 15: cantidad = cantidad + "QUINCE"
                Case Else
                    Select Case (Numero Mod 100) \ 10
                        Case 1: cantidad = cantidad + "DIEZ Y "
                        Case 2: cantidad = cantidad + "VEINTE Y "
                        Case 3: cantidad = cantidad + "TREINTA Y "
                        Case 4: cantidad = cantidad + "CUARENTA Y "
                        Case 5: cantidad = cantidad + "CINCUENTA Y "
                        Case 6: cantidad = cantidad + "SESENTA Y "
                        Case 7: cantidad = cantidad + "SETENTA Y "
                        Case 8: cantidad = cantidad + "OCHENTA Y "
                        Case 9: cantidad = cantidad + "NOVENTA Y "
                    End Select
                    Select Case Numero Mod 10
                        Case 1: cantidad = cantidad + "UN"
                        Case 2: cantidad = cantidad + "DOS"
                        Case 3: cantidad = cantidad + "TRES"
                        Case 4: cantidad = cantidad + "CUATRO"
                        Case 5: cantidad = cantidad + "CINCO"
                        Case 6: cantidad = cantidad + "SEIS"
                        Case 7: cantidad = cantidad + "SIETE"
                        Case 8: cantidad = cantidad + "OCHO"
                        Case 9: cantidad = cantidad + "NUEVE"
                        Case 0:
                            If (Numero Mod 100) \ 10 = 0 Then
                                cantidad = Mid$(cantidad, 1, Len(cantidad) - 1)
                            Else
                                cantidad = Mid$(cantidad, 1, Len(cantidad) - 3)
                            End If
                    End Select
            End Select
    End Select
    LetraCantidad = cantidad
End Function
Public Function LetrasNumero(Numero As Double, ByRef L As String) As String
    Dim Cienes          As Integer
    Dim Miles           As Integer
    Dim Millones        As Integer
    Dim MilesMillones   As Integer

    Dim CantidadCien            As String
    Dim CantidadMil             As String
    Dim CantidadMillon          As String
    Dim CantidadMilesMillon     As String
    Dim TextoNumero             As String
    
    MilesMillones = CInt(Int(Numero / 1000000000))
    Millones = CInt(Int(Numero / 1000000) - MilesMillones * 1000)
    Miles = CInt(Int(Numero / 1000) - Int(Numero / 1000000) * 1000)
    Cienes = Int(Numero - Int(Numero / 1000) * 1000)
    
    CantidadCien = LetraCantidad(Cienes)
    CantidadMil = LetraCantidad(Miles)
    CantidadMillon = LetraCantidad(Millones)
    CantidadMilesMillon = LetraCantidad(MilesMillones)

    If CantidadMilesMillon <> "" Then
        If MilesMillones = 1 Then
            CantidadMilesMillon = CantidadMilesMillon + " MIL "
        Else
            CantidadMilesMillon = CantidadMilesMillon + " MIL "
        End If
    End If

    If CantidadMillon <> "" Then
        If Millones = 1 Then
            If CantidadMilesMillon <> "" Then
                CantidadMillon = CantidadMillon + " MILLONES "
            Else
                CantidadMillon = CantidadMillon + " MILLON "
            End If
        Else
            CantidadMillon = CantidadMillon + " MILLONES "
        End If
    End If
    If CantidadMilesMillon <> "" And CantidadMillon = "" Then
        CantidadMilesMillon = CantidadMilesMillon + " MILLONES "
    End If
    If CantidadMil <> "" Then
        CantidadMil = CantidadMil + " MIL "
    End If
    If Numero = 0 Then
        TextoNumero = " CERO "
    Else
        TextoNumero = CantidadMilesMillon + CantidadMillon + CantidadMil + CantidadCien
    End If
    LetrasNumero = TextoNumero + " PESOS " + Format$(Str((Numero - Int(Numero)) * 100 + 0.1), "00") + "/100"
    L = LetrasNumero
End Function
Public Sub CenterForm(frm As Form)
' Funcion que centra una forma Child

    Dim R As RECT, lRes As Long
    Dim lW As Long, lH As Long
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, R, 0)
    If lRes Then
       With R
            .Left = Screen.TwipsPerPixelX * .Left
            .Top = Screen.TwipsPerPixelY * .Top
            .Right = Screen.TwipsPerPixelX * .Right
            .Bottom = Screen.TwipsPerPixelY * .Bottom
            lW = .Right - .Left
            lH = .Bottom - .Top
            frm.Move .Left + (lW - frm.Width) / 2, .Top + (lH - frm.Height) / 2
       End With
    End If
End Sub

