Attribute VB_Name = "MdlFuncao"
Option Explicit
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long



Public Function sUserID() As String
'Busca o usuário logado na rede
     Dim sBuffer As String
     Dim lSize As Long
     sBuffer = Space$(255)
     lSize = Len(sBuffer)
     GetUserName sBuffer, lSize
     If lSize > 0 Then
          sUserID = Left$(sBuffer, lSize - 1)
     End If
End Function
'seleciona texto do controle
Public Sub SelecionaTudo()
On Error Resume Next
   With Screen.ActiveControl
      .SelStart = 0
      If TypeOf Screen.ActiveControl Is TextBox Then
         .SelLength = Len(Screen.ActiveControl.Text)
      ElseIf TypeOf Screen.ActiveControl Is Masked Then
         .SelLength = Len(Screen.ActiveControl.Texto)
      ElseIf TypeOf Screen.ActiveControl Is Combo_DB Then
         .SelLength = Len(Screen.ActiveControl.Criterio)
      End If
   End With
End Sub
Public Function ChecaInscrE(pUF As String, pInscr As String)
'Função para verificar inscrições estaduais (todos os estados)
'CORRIGIDA PARA OS ESTADOS AC,AP,TO,PE,MT E RR POR CARLOS FABIANO EM 04/12/2003
On Error GoTo Trata_erro
    ChecaInscrE = False
    Dim strBase As String
    Dim strBase2 As String
    Dim strOrigem As String
    Dim strDigito1 As String
    Dim strDigito2 As String
    Dim intPos As Integer
    Dim intValor As Integer
    Dim intSoma As Integer
    Dim intResto As Integer
    Dim intNumero As Long
    Dim intPeso As Integer
    Dim intDig As Integer
    Dim ilPeso(1 To 14) As Integer
    strBase = ""
    strBase2 = ""
    strOrigem = ""
    
    If Trim(pInscr) = "ISENTO" Then
        ChecaInscrE = True
        Exit Function
    End If
    For intPos = 1 To Len(Trim(pInscr))
        If InStr(1, "0123456789P", Mid$(pInscr, intPos, 1), vbTextCompare) > 0 Then
            strOrigem = strOrigem & Mid$(pInscr, intPos, 1)
        End If
    Next
    Select Case UCase(pUF)
        Case "AC" ' Acre
            If Len(pInscr) = 9 Then
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                If Left(strBase, 2) = "01" And Mid$(strBase, 3, 2) <> "00" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                        ChecaInscrE = True
                    End If
                End If
            ElseIf Len(pInscr) = 13 Then 'NOVA INSCRIÇÃO
                strBase = Left(Trim(strOrigem) & "0000000000000", 13)
                If Left(strBase, 2) = "01" Then
                    intSoma = 0
                    intPeso = 2
                    For intPos = 11 To 1 Step -1
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso > 9 Then
                            intPeso = 2
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 11) & strDigito1
                    intSoma = 0
                    intPeso = 2
                    For intPos = 12 To 1 Step -1
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso > 9 Then
                            intPeso = 2
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 12) & strDigito2
                    If strBase2 = strOrigem Then
                        ChecaInscrE = True
                    End If
                End If
            End If
        Case "AL" ' Alagoas
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            If Left(strBase, 2) = "24" Then
                intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intSoma = intSoma * 10
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto = 10, "0", Str(intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                        ChecaInscrE = True
                    End If
            End If
        Case "AM" ' Amazonas
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            intSoma = 0
            For intPos = 1 To 8
                intValor = Val(Mid$(strBase, intPos, 1))
                intValor = intValor * (10 - intPos)
                intSoma = intSoma + intValor
            Next
            If intSoma < 11 Then
                strDigito1 = Right(Str(11 - intSoma), 1)
            Else
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
            End If
            strBase2 = Left(strBase, 8) & strDigito1
            If strBase2 = strOrigem Then
                ChecaInscrE = True
            End If
        Case "AP" ' Amapa
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            intPeso = 0
            intDig = 0
            If Left(strBase, 2) = "03" Then
                intNumero = Val(Left(strBase, 8))
                If intNumero >= 3000001 And _
                    intNumero <= 3017009 Then
                    intPeso = 5
                    intDig = 0
                ElseIf intNumero >= 3017001 And _
                    intNumero <= 3019022 Then
                    intPeso = 9
                    intDig = 1
                ElseIf intNumero >= 3019023 Then
                    intPeso = 0
                    intDig = 0
                End If
                intSoma = intPeso
                For intPos = 1 To 8
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                intValor = 11 - intResto
                If intValor = 10 Then
                    intValor = 0
                ElseIf intValor = 11 Then
                    intValor = intDig
                End If
                strDigito1 = Right(Str(intValor), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    ChecaInscrE = True
                End If
            End If
        Case "BA" ' Bahia
            strBase = Left(Trim(strOrigem) & "00000000", 8)
            If InStr(1, "0123458", Left(strBase, 1), vbTextCompare) > 0 Then
                intSoma = 0
                For intPos = 1 To 6
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * (8 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 10
                strDigito2 = Right(IIf(intResto = 0, "0", Str(10 - intResto)), 1)
                strBase2 = Left(strBase, 6) & strDigito2
                intSoma = 0
                For intPos = 1 To 7
                    intValor = Val(Mid$(strBase2, intPos, 1))
                    intValor = intValor * (9 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 10
                strDigito1 = Right(IIf(intResto = 0, "0", Str(10 - intResto)), 1)
            Else
                intSoma = 0
                For intPos = 1 To 6
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * (8 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                strBase2 = Left(strBase, 6) & strDigito2
                intSoma = 0
                For intPos = 1 To 7
                    intValor = Val(Mid$(strBase2, intPos, 1))
                    intValor = intValor * (9 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
            End If
            strBase2 = Left(strBase, 6) & strDigito1 & strDigito2
            If strBase2 = strOrigem Then
                ChecaInscrE = True
            End If
        Case "CE" ' Ceara
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            intSoma = 0
            For intPos = 1 To 8
                intValor = Val(Mid$(strBase, intPos, 1))
                intValor = intValor * (10 - intPos)
                intSoma = intSoma + intValor
            Next
            intResto = intSoma Mod 11
            intValor = 11 - intResto
            If intValor > 9 Then
                intValor = 0
            End If
            strDigito1 = Right(Str(intValor), 1)
            strBase2 = Left(strBase, 8) & strDigito1
            If strBase2 = strOrigem Then
                ChecaInscrE = True
            End If
        Case "DF" ' Distrito Federal
        strBase = Left(Trim(strOrigem) & "0000000000000", 13)
        If Left(strBase, 3) = "073" Then
            intSoma = 0
            intPeso = 2
            For intPos = 11 To 1 Step -1
                intValor = Val(Mid$(strBase, intPos, 1))
                intValor = intValor * intPeso
                intSoma = intSoma + intValor
                intPeso = intPeso + 1
                If intPeso > 9 Then
                    intPeso = 2
                End If
            Next
            intResto = intSoma Mod 11
            strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
            strBase2 = Left(strBase, 11) & strDigito1
            intSoma = 0
            intPeso = 2
            For intPos = 12 To 1 Step -1
                intValor = Val(Mid$(strBase, intPos, 1))
                intValor = intValor * intPeso
                intSoma = intSoma + intValor
                intPeso = intPeso + 1
                If intPeso > 9 Then
                    intPeso = 2
                End If
            Next
            intResto = intSoma Mod 11
            strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
            strBase2 = Left(strBase, 12) & strDigito2
            If strBase2 = strOrigem Then
                ChecaInscrE = True
            End If
        End If
        Case "ES" ' Espirito Santo
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            intSoma = 0
            For intPos = 1 To 8
                intValor = Val(Mid$(strBase, intPos, 1))
                intValor = intValor * (10 - intPos)
                intSoma = intSoma + intValor
            Next
            intResto = intSoma Mod 11
            strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
            strBase2 = Left(strBase, 8) & strDigito1
            If strBase2 = strOrigem Then
                ChecaInscrE = True
            End If
        Case "GO" ' Goias
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            If InStr(1, "10,11,15", Left(strBase, 2), vbTextCompare) > 0 Then
                intSoma = 0
                For intPos = 1 To 8
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                If intResto = 0 Then
                    strDigito1 = "0"
                ElseIf intResto = 1 Then
                    intNumero = Val(Left(strBase, 8))
                    strDigito1 = Right(IIf(intNumero >= 10103105 And intNumero <= 10119997, "1", "0"), 1)
                Else
                    strDigito1 = Right(Str(11 - intResto), 1)
                End If
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    ChecaInscrE = True
                End If
            End If
        Case "MA" ' Maranhão
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            If Left(strBase, 2) = "12" Then
                intSoma = 0
                For intPos = 1 To 8
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    ChecaInscrE = True
                End If
            End If
        Case "MT" ' Mato Grosso
            strBase = Left(Trim(strOrigem) & "00000000000", 11)
                strBase2 = Left(strBase, 2) & Mid$(strBase, 5, 6)
                intSoma = 0
                For intPos = 1 To 8
                    intValor = Val(Mid$(strBase2, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                strBase2 = Left(strBase, 10) & strDigito1
                If strBase2 = strOrigem Then
                    ChecaInscrE = True
                End If
        Case "MS" ' Mato Grosso do Sul
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            If Left(strBase, 2) = "28" Then
                intSoma = 0
                For intPos = 1 To 8
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    ChecaInscrE = True
                End If
            End If
        Case "MG" ' Minas Gerais
            strBase = Left(Trim(strOrigem) & "0000000000000", 13)
            strBase2 = Left(strBase, 3) & "0" & Mid$(strBase, 4, 8)
            intNumero = 2
            For intPos = 1 To 12
                intValor = Val(Mid$(strBase2, intPos, 1))
                intNumero = IIf(intNumero = 2, 1, 2)
                intValor = intValor * intNumero
                If intValor > 9 Then
                    strDigito1 = Format(intValor, "00")
                    intValor = Val(Left(strDigito1, 1)) + _
                    Val(Right(strDigito1, 1))
                End If
                intSoma = intSoma + intValor
            Next
            intValor = intSoma
            While Right(Format(intValor, "000"), 1) <> "0"
                intValor = intValor + 1
            Wend
            strDigito1 = Right(Format(intValor - intSoma, "00"), 1)
            strBase2 = Left(strBase, 11) & strDigito1
            intSoma = 0
            intPeso = 2
            For intPos = 12 To 1 Step -1
                intValor = Val(Mid$(strBase2, intPos, 1))
                intValor = intValor * intPeso
                intSoma = intSoma + intValor
                intPeso = intPeso + 1
                If intPeso > 11 Then
                    intPeso = 2
                End If
            Next
            intResto = intSoma Mod 11
            strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
            strBase2 = strBase2 & strDigito2
            If strBase2 = strOrigem Then
               ChecaInscrE = True
            End If
        Case "PA" ' Para
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            If Left(strBase, 2) = "15" Then
                intSoma = 0
                For intPos = 1 To 8
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                  ChecaInscrE = True
                End If
            End If
        Case "PB" ' Paraiba
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            intSoma = 0
            For intPos = 1 To 8
                intValor = Val(Mid$(strBase, intPos, 1))
                intValor = intValor * (10 - intPos)
                intSoma = intSoma + intValor
            Next
            intResto = intSoma Mod 11
            intValor = 11 - intResto
            If intValor > 9 Then
                intValor = 0
            End If
            strDigito1 = Right(Str(intValor), 1)
            strBase2 = Left(strBase, 8) & strDigito1
            If strBase2 = strOrigem Then
                ChecaInscrE = True
            End If
        Case "PE" ' Pernambuco
            'Multiplique cada algarismo principal pelo seu respectivo peso
            ilPeso(1) = Val(Mid(Trim(pInscr), 14, 1))
            ilPeso(2) = Val(Mid(Trim(pInscr), 13, 1)) * 2
            ilPeso(3) = Val(Mid(Trim(pInscr), 12, 1)) * 3
            ilPeso(4) = Val(Mid(Trim(pInscr), 11, 1)) * 4
            ilPeso(5) = Val(Mid(Trim(pInscr), 10, 1)) * 5
            ilPeso(6) = Val(Mid(Trim(pInscr), 9, 1)) * 6
            ilPeso(7) = Val(Mid(Trim(pInscr), 8, 1)) * 7
            ilPeso(8) = Val(Mid(Trim(pInscr), 7, 1)) * 8
            ilPeso(9) = Val(Mid(Trim(pInscr), 6, 1)) * 9
            ilPeso(10) = Val(Mid(Trim(pInscr), 5, 1)) * 1
            ilPeso(11) = Val(Mid(Trim(pInscr), 4, 1)) * 2
            ilPeso(12) = Val(Mid(Trim(pInscr), 3, 1)) * 3
            ilPeso(13) = Val(Mid(Trim(pInscr), 2, 1)) * 4
            ilPeso(14) = Val(Mid(Trim(pInscr), 1, 1)) * 5
            'Some os produtos obtidos para encontrar o total:
            For intNumero = 2 To 14
                intSoma = intSoma + ilPeso(intNumero)
            Next intNumero
            'Divida esse total pela constante "11" para obter o resto
            'Subtraia esse resto da constante "11" para encontrar o dígito verificador:
            intDig = 11 - (intSoma Mod 11)
            'Quando essa diferença for maior que "9", subtraia "10" unidades para obter
            'o valor do dígito verificador, uma vez que o mesmo deve ser sempre representado
            'por apenas um algarismo
            Select Case intDig
                Case Is > 9
                    intDig = intDig - 10
                    If intDig = ilPeso(1) Then
                        ChecaInscrE = True
                    Else
                        ChecaInscrE = False
                    End If
                Case Is <> ilPeso(1)
                    ChecaInscrE = False
                Case Else
                    ChecaInscrE = True
            End Select
        Case "PI" ' Piaui
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            intSoma = 0
            For intPos = 1 To 8
                intValor = Val(Mid$(strBase, intPos, 1))
                intValor = intValor * (10 - intPos)
                intSoma = intSoma + intValor
            Next
            intResto = intSoma Mod 11
            strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
            strBase2 = Left(strBase, 8) & strDigito1
            If strBase2 = strOrigem Then
                ChecaInscrE = True
            End If
        Case "PR" ' Parana
            strBase = Left(Trim(strOrigem) & "0000000000", 10)
            intSoma = 0
            intPeso = 2
            For intPos = 8 To 1 Step -1
                intValor = Val(Mid$(strBase, intPos, 1))
                intValor = intValor * intPeso
                intSoma = intSoma + intValor
                intPeso = intPeso + 1
                If intPeso > 7 Then
                    intPeso = 2
                End If
            Next
            intResto = intSoma Mod 11
            strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
            strBase2 = Left(strBase, 8) & strDigito1
            intSoma = 0
            intPeso = 2
            For intPos = 9 To 1 Step -1
                intValor = Val(Mid$(strBase2, intPos, 1))
                intValor = intValor * intPeso
                intSoma = intSoma + intValor
                intPeso = intPeso + 1
                If intPeso > 7 Then
                    intPeso = 2
                End If
            Next
            intResto = intSoma Mod 11
            strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
            strBase2 = strBase2 & strDigito2
            If strBase2 = strOrigem Then
                ChecaInscrE = True
            End If
        Case "RJ" ' Rio de Janeiro
            strBase = Left(Trim(strOrigem) & "00000000", 8)
            intSoma = 0
            intPeso = 2
            For intPos = 7 To 1 Step -1
                intValor = Val(Mid$(strBase, intPos, 1))
                intValor = intValor * intPeso
                intSoma = intSoma + intValor
                intPeso = intPeso + 1
                If intPeso > 7 Then
                    intPeso = 2
                End If
            Next
            intResto = intSoma Mod 11
            strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
            strBase2 = Left(strBase, 7) & strDigito1
            If strBase2 = strOrigem Then
                ChecaInscrE = True
            End If
        Case "RN" ' Rio Grande do Norte
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            If Left(strBase, 2) = "20" Then
                intSoma = 0
                For intPos = 1 To 8
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intSoma = intSoma * 10
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto > 9, "0", Str(intResto)), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    ChecaInscrE = True
                End If
        End If
        Case "RO" ' Rondonia
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            strBase2 = Mid$(strBase, 4, 5)
            intSoma = 0
            For intPos = 1 To 5
                intValor = Val(Mid$(strBase2, intPos, 1))
                intValor = intValor * (7 - intPos)
                intSoma = intSoma + intValor
            Next
            intResto = intSoma Mod 11
            intValor = 11 - intResto
            If intValor > 9 Then
                intValor = intValor - 10
            End If
            strDigito1 = Right(Str(intValor), 1)
            strBase2 = Left(strBase, 8) & strDigito1
            If strBase2 = strOrigem Then
                ChecaInscrE = True
            End If
        Case "RR" ' Roraima
            strBase = Left(Trim(strOrigem) & "000000000", 9)
            If Left(strBase, 2) = "24" Then
                intSoma = 0
                For intPos = 1 To 8
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * intPos
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 9
                strDigito1 = Right(Str(intResto), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    ChecaInscrE = True
                End If
            End If
        Case "RS" ' Rio Grande do Sul
            strBase = Left(Trim(strOrigem) & "0000000000", 10)
            intNumero = Val(Left(strBase, 3))
            If intNumero > 0 And intNumero < 468 Then
                intSoma = 0
                intPeso = 2
                For intPos = 9 To 1 Step -1
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 9 Then
                        intPeso = 2
                    End If
                Next
                intResto = intSoma Mod 11
                intValor = 11 - intResto
                If intValor > 9 Then
                    intValor = 0
                End If
                strDigito1 = Right(Str(intValor), 1)
                strBase2 = Left(strBase, 9) & strDigito1
                If strBase2 = strOrigem Then
                    ChecaInscrE = True
                End If
            End If
        Case "SC" ' Santa Catarina
        strBase = Left(Trim(strOrigem) & "000000000", 9)
        intSoma = 0
        For intPos = 1 To 8
            intValor = Val(Mid$(strBase, intPos, 1))
            intValor = intValor * (10 - intPos)
            intSoma = intSoma + intValor
        Next
        intResto = intSoma Mod 11
        strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
        strBase2 = Left(strBase, 8) & strDigito1
        If strBase2 = strOrigem Then
            ChecaInscrE = True
        End If
        Case "SE" ' Sergipe
        strBase = Left(Trim(strOrigem) & "000000000", 9)
        intSoma = 0
        For intPos = 1 To 8
            intValor = Val(Mid$(strBase, intPos, 1))
            intValor = intValor * (10 - intPos)
            intSoma = intSoma + intValor
        Next
        intResto = intSoma Mod 11
        intValor = 11 - intResto
        If intValor > 9 Then
            intValor = 0
        End If
        strDigito1 = Right(Str(intValor), 1)
        strBase2 = Left(strBase, 8) & strDigito1
        If strBase2 = strOrigem Then
            ChecaInscrE = True
        End If
        Case "SP" ' São Paulo
            If Left(strOrigem, 1) = "P" Then
                strBase = Left(Trim(strOrigem) & "0000000000000", 13)
                strBase2 = Mid$(strBase, 2, 8)
                intSoma = 0
                intPeso = 1
                For intPos = 1 To 8
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso = 2 Then
                        intPeso = 3
                    End If
                    If intPeso = 9 Then
                        intPeso = 10
                    End If
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(Str(intResto), 1)
                strBase2 = Left(strBase, 8) & strDigito1 & Mid$(strBase, 11, 3)
            Else
                strBase = Left(Trim(strOrigem) & "000000000000", 12)
                intSoma = 0
                intPeso = 1
                For intPos = 1 To 8
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso = 2 Then
                        intPeso = 3
                    End If
                    If intPeso = 9 Then
                        intPeso = 10
                    End If
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(Str(intResto), 1)
                strBase2 = Left(strBase, 8) & strDigito1 & Mid$(strBase, 10, 2)
                intSoma = 0
                intPeso = 2
                For intPos = 11 To 1 Step -1
                    intValor = Val(Mid$(strBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 10 Then
                        intPeso = 2
                    End If
                Next
                intResto = intSoma Mod 11
                strDigito2 = Right(Str(intResto), 1)
                strBase2 = strBase2 & strDigito2
            End If
            If strBase2 = strOrigem Then
                ChecaInscrE = True
            End If
        Case "TO" ' Tocantins
            strBase = Left(Trim(strOrigem) & "00000000000", 11)
                If InStr(1, "01,02,03,99", Mid$(strBase, 3, 2), vbTextCompare) > 0 Then
                    strBase2 = Left(strBase, 2) & Mid$(strBase, 5, 6)
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = Val(Mid$(strBase2, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
                    strBase2 = Left(strBase, 10) & strDigito1
                    If strBase2 = strOrigem Then
                        ChecaInscrE = True
                    End If
                End If
        End Select
    Exit Function
Trata_erro:
    MsgBox "ERRO" & Err.Number & " - " & Err.Description, vbCritical
End Function
Function ConfereCPFCGC(PfPj As Integer, CódIgo As Double) As Integer
'***********************************************************************
'*   FUNCAO PARA VERIFICAR SE O CNPJ OU CPF E VALIDO                   *
'*   PFPJ = 1(TRUE) TESTA O CNPJ                                       *
'*   PFPJ = 0(FALSE) TESTA O CPF                                       *
'*   CODIGO = NUMERO CNPJ/CPF QUE SERÁ TESTADO                         *
'*   SE A FUNCAO RETORNAR ZERO, CNPJ/CPF INVALIDO                      *
'*   SE A FUNCAO RETORNAR ALGO DIFERENTE DE ZERO O CNPJ/CPF É VALIDO   *
'***********************************************************************
    ReDim DIg(1 To 14) As Integer
                                   
    Dim i As Integer, soma As Integer, Resto As Integer, Tamanho As Integer
    
    On Error GoTo Trata_erro
    
    ConfereCPFCGC = False
    soma = 0
    Tamanho = Len(CStr(CódIgo))
    
    If CódIgo = 0 Then
       ConfereCPFCGC = False
       Exit Function
    End If
    
    If (PfPj = True) And (Tamanho <= 11) Then            '*** CPF
        For i = 11 To 1 Step -1
              DIg(i) = SelecDIg(i - 1, CódIgo)
              If i > 2 Then
                soma = soma + DIg(i) * (i - 1)
              End If
        Next i
    
        Resto = soma Mod 11
      
        If (Resto <> 0) Then
          Resto = 11 - Resto
        End If
      
        If Resto = 10 Then
           Resto = 0
        ElseIf Resto = 11 Then
           Resto = 1
        End If
    
        If (Resto = DIg(2)) Then
            soma = 0
            For i = 11 To 2 Step -1
              soma = soma + (DIg(i) * i)
            Next i
        
            Resto = soma Mod 11
            
            If (Resto <> 0) Then
              Resto = 11 - Resto
            End If
    
            If Resto = 10 Then
                Resto = 0
            ElseIf Resto = 11 Then
                Resto = 1
            End If
    
            If (Resto = DIg(1)) Then
              ConfereCPFCGC = True
            End If
        End If
    Else                                    '*** CGC
    
          If (Tamanho <= 14) Then
                For i = 14 To 1 Step -1
                    DIg(i) = SelecDIg(i - 1, CódIgo)
                    If (i > 10) Then
                      soma = soma + DIg(i) * (i - 9)
                    Else
                        If (i > 2) Then
                          soma = soma + DIg(i) * (i - 1)
                        End If
                    End If
                Next i
        
                Resto = soma Mod 11
      
                If (Resto <> 0) Then
                  Resto = 11 - Resto
                End If
        
                If (Resto = 10) Then
                  Resto = 0
                End If
            
                If (Resto = DIg(2)) Then
                      soma = 0
                      For i = 14 To 2 Step -1
                          If (i > 9) Then
                            soma = soma + (DIg(i) * (i - 8))
                          Else
                            soma = soma + (DIg(i) * i)
                          End If
                      Next i
                
                      Resto = soma Mod 11
                    
                      If (Resto <> 0) Then
                        Resto = 11 - Resto
                      End If
                
                      If (Resto = 10) Then
                        Resto = 0
                      End If
                  
                    If (Resto = DIg(1)) Then
                      ConfereCPFCGC = True
                    End If
              End If
        End If
    End If

    Exit Function
 
Trata_erro:
    Rotina_Erro "ConfereCPFCGC"
End Function
Function SelecDIg(ByVal Tam As Integer, Dado As Double) As Integer
'Função usada na rotina que verifica CGC e CPF
    Dim Conv As String
    Dim AUX As Double
    Dim Selecionado As Integer
    
    On Error GoTo Trata_erro
    
    SelecDIg = 0
    AUX = 0
    
    If (Tam >= 0) Or (Dado = 0) Then
         Conv = "1" & String$(Tam, "0")
         AUX = Dado / CDbl(Conv)
        
         Selecionado = AUX - (AUX - Fix(AUX))
         Dado = Dado - (Selecionado * CDbl(Conv))  '*** ATUALIZA NUMERO INICIAL
        
         SelecDIg = Selecionado
    End If

    Exit Function
 
Trata_erro:
  
  Rotina_Erro "SelecDIg"

End Function
Public Function ConsisteUF(ByVal vUF As String) As Boolean
'*********************************************************
'**        Rotina para Verificar se a UF é válida       **
'**      Criado por Randerson Maurilio em 28/08/2003    **
'**                     USIFAST - CPD                   **
'*********************************************************
   ConsisteUF = False
   If vUF <> "" And Len(vUF) = 2 Then
      ConsisteUF = InStr(1, sgEstados, UCase$(vUF)) > 0
   End If
End Function
Public Sub SetaConsisteUF()
On Error Resume Next
'***************************************************************
'** Rotina para auxiliar a verificação de cosnsistencia da UF **
'**        Criado por Randerson Maurilio em 28/08/2003        **
'**                        USIFAST - CPD                      **
'***************************************************************
   Call consulta("Select CodUF From Estado")
   With Rs
      Do While Not Rs.EOF
         sgEstados = sgEstados & (!CodUf & "-")
         .MoveNext
      Loop
      .Close
      Set Rs = Nothing
   End With
End Sub
Public Function RetiraFormatacao(Texto As String)
'Retira caracteres especiais deixando somente números
On Error GoTo Trataerro
    If Texto = "" Then Exit Function
    Dim blIndex As Byte
    RetiraFormatacao = vbNullString
    For blIndex = 1 To Len(Texto)
        If Asc(Mid(Texto, blIndex, 1)) >= 48 And Asc(Mid(Texto, blIndex, 1)) <= 57 Then
            RetiraFormatacao = RetiraFormatacao & Mid(Texto, blIndex, 1)
        End If
    Next blIndex
    Exit Function
Trataerro:
    Rotina_Erro "RetiraFormatacao"
End Function
Public Sub LimpaForm(frmForm As Form)
'*****************************************************
' Rotina para limpar todos os controles do form q foram solicitados conforme abaixo:
' Se TAG = 99 => ""
' Se TAG = 90 => 0
' Se TAG = 91 => Data formatada "__/__/____"
' Se TAG = 92 => número formatado com duas casas "0,00"
' Se TAG = 93 => somente o mes e o ano "__/____"
' Se TAG = 94 => hora e minuto "__:__"
' Se TAG = 95 => Data e Hora "__/__/____ __:__:__"
'   Randerson Maurilio - 29/08/2003
'*****************************************************
On Error Resume Next
Dim i As Integer
Dim vMaskTmp As String
   With frmForm
      For i = 1 To .Controls.Count - 1
         If .Controls(i).Tag <> "" Then
            If TypeOf .Controls(i) Is Label Then
               If .Controls(i).Tag = 99 Then
                  .Controls(i).Caption = ""
               ElseIf .Controls(i).Tag = 90 Then
                  .Controls(i).Caption = 0
               ElseIf .Controls(i).Tag = 91 Then
                  .Controls(i).Caption = sgStrData
               ElseIf .Controls(i).Tag = 92 Then
                  .Controls(i).Caption = "0,00"
               ElseIf .Controls(i).Tag = 93 Then
                  .Controls(i).Caption = sgStrMesAno
               ElseIf .Controls(i).Tag = 94 Then
                  .Controls(i).Caption = sgStrHora
               ElseIf .Controls(i).Tag = 95 Then
                  .Controls(i).Caption = sgStrDataHora
               End If
            ElseIf (TypeOf .Controls(i) Is TextBox) Or (TypeOf .Controls(i) Is ComboBox) Or _
                  (TypeOf .Controls(i) Is MaskEdBox) Then
               If .Controls(i).Tag = 90 Then
                  .Controls(i).Text = 0
               ElseIf .Controls(i).Tag = 91 Then
                  .Controls(i).Text = sgStrData
               ElseIf .Controls(i).Tag = 92 Then
                  .Controls(i).Text = "0,00"
               ElseIf .Controls(i).Tag = 93 Then
                  .Controls(i).Text = sgStrMesAno
               ElseIf .Controls(i).Tag = 94 Then
                  .Controls(i).Text = sgStrHora
               ElseIf .Controls(i).Tag = 95 Then
                  .Controls(i).Text = sgStrDataHora
               End If
            ElseIf (TypeOf .Controls(i) Is Masked) Then
               If .Controls(i).Tag = 90 Then
                  .Controls(i).Texto = 0
               ElseIf .Controls(i).Tag = 91 Then
                  .Controls(i).Texto = sgStrData
               ElseIf .Controls(i).Tag = 92 Then
                  .Controls(i).Texto = "0,00"
               ElseIf .Controls(i).Tag = 93 Then
                  .Controls(i).Texto = sgStrMesAno
               ElseIf .Controls(i).Tag = 94 Then
                  .Controls(i).Texto = sgStrHora
               ElseIf .Controls(i).Tag = 95 Then
                  .Controls(i).Texto = sgStrDataHora
               End If
            End If
         Else
            If TypeOf .Controls(i) Is TextBox Then .Controls(i).Text = ""
            If TypeOf .Controls(i) Is ComboBox Then .Controls(i).Text = ""
            If TypeOf .Controls(i) Is CheckBox Then .Controls(i).Value = vbUnchecked
            If TypeOf .Controls(i) Is OptionButton Then .Controls(i).Value = False
            'Como a MaskEdBox não aceita vazio, tem esta forma de enganá-la
            If TypeOf .Controls(i) Is MaskEdBox Then
               vMaskTmp = .Controls(i).Mask
               .Controls(i).Mask = ""
               .Controls(i).Text = ""
               .Controls(i).Mask = vMaskTmp
            End If
            If TypeOf .Controls(i) Is Combo_DB Then .Controls(i).Criterio = ""
            If TypeOf .Controls(i) Is Combo_DB Then .Controls(i).Codigo = ""
            If TypeOf .Controls(i) Is Masked Then .Controls(i).Texto = ""
         End If
      Next
   End With
End Sub
Public Function SeeDados(vTabela As String, vWhere As String, vCampoRet As String, Optional vTipo As Integer) As Variant
'*********************************************************
'**     Rotina para Ler uma informação em uma tabela    **
'**      Criado por Randerson Maurilio em 30/08/2003    **
'**                     USIFAST - CPD                   **
'*********************************************************
'Ex:    lsFlgNumPrg = SeeDados("FILIAL", "codfil = " & igCodFil, "flgnumprg", 2)
'Pesquisa na tabela FILIAL pelo campo CODFIL e retorna o valor contido no campo FlgNumPrg como S/N
' e o joga o resultado da query no flag lsFlgNumPrg
On Error Resume Next
Dim vSql As String
Dim TBCursor As ADODB.Recordset
      vSql = "SELECT " & vCampoRet & " FROM " & vTabela & vbCr
      If Not vWhere = "" Then _
         vSql = vSql & " WHERE " & vWhere & vbCr
      vSql = vSql & " ORDER BY " & vCampoRet & "; "
      Set TBCursor = New ADODB.Recordset
      With TBCursor
         .CursorType = adOpenKeyset 'TIPO DE NAVEGACAO
         .LockType = adLockReadOnly  'FAZ A CONSULTA "SOMENTE LEITURA NA TABELA"
         .Source = vSql
         .ActiveConnection = Conexao ' ATIVA A CONEXAO COM O BANCO DE DADOS
         .Open 'ABRE O BANCO DE DADOS
         If .EOF Then
            Select Case vTipo
               Case 0: SeeDados = ""   'Para tipos de dados string
               Case 1: SeeDados = 0    'Para tipos de dados inteiros
               Case 2: SeeDados = "N"  'Para tipos de dados "S/N"
            End Select
         Else
            If InStr(1, vCampoRet, ".") > 0 Then
               vCampoRet = Right(vCampoRet, Len(vCampoRet) - 2)
            End If
            SeeDados = Trim(TBCursor(vCampoRet))
            If IsNull(SeeDados) Then
               Select Case vTipo
                  Case 0: SeeDados = ""   'Para tipos de dados string
                  Case 1: SeeDados = 0    'Para tipos de dados inteiros
                  Case 2: SeeDados = "N"  'Para tipos de dados "S/N"
               End Select
            End If
         End If
      End With
      TBCursor.Close
      Set TBCursor = Nothing
End Function
Public Sub PreecheComboUF(oCombo As ComboBox)
'Procedure criada para preencher um ComboBox com os Estados da tabela Estado
    Dim i As Integer
    consulta "Select CodUF From Estado"
    oCombo.Clear
    With Rs
      .MoveFirst
      Do While Not .EOF
          oCombo.AddItem !CodUf
          .MoveNext
      Loop
      If .State <> adStateClosed Then .Close
   End With
End Sub
Public Sub EventoEnter(vKeyAscii As Integer)
    'Isto é para o caso dos campos de Observaçoes não pegar o enter
    If TypeOf Screen.ActiveControl Is TextBox Then _
      If Screen.ActiveControl.MultiLine Then Exit Sub
    Select Case vKeyAscii
      Case vbKeyReturn
            vKeyAscii = 0
            SendKeys "{Tab}"
      Case vbKeyEscape
            vKeyAscii = 0
            SendKeys "+{Tab}"
      End Select
End Sub
Public Sub ConsultaQuery(cboTemp As Combo_DB, slCampos As String, SlTabela As String, slCampoQuery As String, Optional vWhere As String)
On Error Resume Next
Dim sgQuery As String
   With cboTemp
      sgQuery = "SELECT " & slCampos & vbCr & _
             "FROM " & SlTabela & vbCr
            If .Criterio <> "" Then
               If vWhere = "" Then
                  sgQuery = sgQuery & "WHERE " & slCampoQuery & " LIKE '" & .Criterio & "%' " & vbCr
               Else
                  sgQuery = sgQuery & "WHERE " & vWhere
               End If
            End If
      sgQuery = sgQuery & "ORDER BY " & slCampoQuery
      .query = sgQuery
   End With
End Sub
Public Function ConsultaGrid(dGrid As DataGrid, sgQuery As String) As Boolean
'*********************************************************
'**    Rotina para Preencher um DataGrid de Consulta    **
'**      Criado por Randerson Maurilio em 11/09/2003    **
'**                     USIFAST - CPD                   **
'*********************************************************
'Exemplo de Como utilizar :
'   If Not ConsultaGrid(grdConsulta, sgQuery) Then txtPesquisa.SetFocus

On Error GoTo Trataerro
   ConsultaGrid = False
   Call consulta(sgQuery)
   If Rs.EOF Then
      MsgBox "Nenhum registro a apresentar.", vbExclamation, "Atenção!"
   ElseIf Rs.RecordCount > LIMITE_PESQUISA Then
      MsgBox "Especifique um critério de pesquisa para que os dados sejam exibidos.", vbExclamation, "Atenção!"
   Else
      Set dGrid.DataSource = Rs
      'Call AjustaColWidthDB(dGrid)
      ConsultaGrid = True
      Exit Function
   End If
   Rs.Close
   Set Rs = Nothing
   Set dGrid.DataSource = Nothing
Exit Function
Trataerro:
    Rotina_Erro "ConsultaGrid"
End Function
'Public Sub CarregaForm(vForm As Form)
''*********************************************************
''**      Rotina para Exibir os forms no MDISisTrans     **
''**     Criado por Randerson Maurilio em 18/09/2003     **
''**     ALTERADO POR CARLOS FABIANO EM 26/01/2004       **
''**                     USIFAST - CPD                   **
''*********************************************************
'On Error Resume Next
'   If Forms.Count >= 2 Then
'       MsgBox "Existe Mais de uma Tela Aberta no Sistema" + Chr(13) + "Feche-a antes de Abrir Outra", vbInformation
'   Else
'        With vForm
'           .Show
'           .SetFocus
'        End With
'   End If
'End Sub
Public Sub CarregaForm(vForm As Form, vMenu As String, Optional vModal As Boolean)
'*********************************************************
'**      Rotina para Exibir os forms no MDISisTrans     **
'**     Criado por Randerson Maurilio em 18/09/2003     **
'**     ALTERADO POR CARLOS FABIANO EM 26/01/2004       **
'**     Alterado por Randerson Maurilio em 16/02/2004   **
'**     ALTERADO POR CARLOS FABIANO EM 01/03/2004       **
'**                     USIFAST - CPD                   **
'*********************************************************
On Error Resume Next
'   If Forms.Count >= 3 Then
'       MsgBox "Existe uma Tela Aberta no Sistema" + Chr(13) + "Feche-a antes de Abrir Outra", vbInformation
'   Else
'        With FrmAcessoMenu2
'            For blI = 1 To .GrdMenu.Rows - 1
'                If UCase(vMenu) = UCase(.GrdMenu.TextMatrix(blI, 0)) Then
'                    sgflgmanut = .GrdMenu.TextMatrix(blI, 1)
'                    Exit For
'                End If
'            Next
'        End With
        With vForm
           If vModal Then
              .Show vbModal, Mdi_ProjUno
           Else
              .Show
              .SetFocus
           End If
        End With
'   End If

End Sub
'**********FABIANO ***************
Public Sub ajusta_combo(comb_recebe As ComboBox, comb_valor As ComboBox)
'*******************PROCEDURE CRIADA POR CARLOS FABIANO
'ROTINA PARA AJUSTAR O INDICE DA LISTA ENTRE DOIS COMBOS
'********************************************************************
    If comb_valor.ListIndex = -1 Then Exit Sub
    comb_recebe.Text = comb_recebe.List(comb_valor.ListIndex)
End Sub
Public Sub ajustajanela(formulario As Form, h As Long, l As Long, pos_sup As Long, pos_lat As Long)
'**************************************************************************************************
'ROTINA PARA AJUSTAR JANELAS CONFIGURADAS COMO MDICHILD
'**************************************************************************************************
    formulario.Height = l
    formulario.Width = h
    formulario.Top = pos_sup
    formulario.Left = pos_lat
End Sub
Public Function ProcuraForm(ByVal form_name As String) As Boolean
'*********************************************************
'**    Rotina para verificar se um form está aberto     **
'**     Criado por Randerson Maurilio em 30/09/2003     **
'**                     USIFAST - CPD                   **
'*********************************************************
On Error Resume Next
Dim i As Integer
ProcuraForm = False
' Procura pelos forms carregados.
For i = 0 To Forms.Count - 1
   If Forms(i).Name = form_name Then
      'se encontramos retorna o form.
      ProcuraForm = True
     Exit For
   End If
Next i
End Function
Public Function FormataCPFCNPJ(ByVal vCPFCNPJ As String) As String
On Error Resume Next
Dim x As Integer
Dim vAux, vCPFCNPJAux As String
   vCPFCNPJAux = ""
   vCPFCNPJ = Trim(vCPFCNPJ)
   For x = 1 To Len(vCPFCNPJ)
      If Len(vCPFCNPJ) > 11 Then 'CNPJ
         Select Case x
            Case 3, 6: vCPFCNPJAux = vCPFCNPJAux & "."
            Case 9: vCPFCNPJAux = vCPFCNPJAux & "/"
            Case 13: vCPFCNPJAux = vCPFCNPJAux & "-"
         End Select
      Else 'CPF
         Select Case x
            Case 4, 7: vCPFCNPJAux = vCPFCNPJAux & "."
            Case 10: vCPFCNPJAux = vCPFCNPJAux & "-"
         End Select
      End If
      vAux = Mid(vCPFCNPJ, x, 1)
      vCPFCNPJAux = vCPFCNPJAux & vAux
   Next
   FormataCPFCNPJ = vCPFCNPJAux
End Function
Public Function DataSQL(ByVal vData As String) As String
On Error Resume Next
   DataSQL = Format(CDate(vData), "mm/dd/yyyy")
End Function

'Valida o digito verificador
'Jeferson C. Oliveira
Function DigitoVerificador(ByVal vValor As Double, Optional ByVal vDigit As Integer, Optional ByVal vValidar As Boolean = False)
    Dim iCount  As Integer
    Dim vValores(10)
    Dim vSoma   As Double
    
    'guarda os valores
    For iCount = 10 To 1 Step -1
        If Len(Trim(vValor)) > 10 Then Exit For
        vValores(iCount) = Val(Mid$(StrReverse(vValor), (11 - iCount), 1))
    Next
    
    'Somar todos os valores
    vSoma = (vValores(1) * 3) + (vValores(2) * 2) + (vValores(3) * 9) + (vValores(4) * 8) + (vValores(5) * 7) + (vValores(6) * 6) + (vValores(7) * 5) + (vValores(8) * 4) + (vValores(9) * 3) + (vValores(10) * 2)
    
    
    If (vSoma Mod 11) = 0 Then
        vSoma = 0
    Else
        vSoma = Abs(11 - (vSoma Mod 11))
        If vSoma = 10 Then vSoma = 0
    End If
    
    If vValidar Then
        If vSoma = vDigit Then
            DigitoVerificador = True
        Else
            DigitoVerificador = False
        End If
    Else
        DigitoVerificador = vSoma
    End If
        
End Function

Public Sub procuraCombo(VALOR As String, comb As ComboBox)
    'PROCEDURE CRIADA POR CARLOS FABIANO
    'PROCEDURE PARA PROCURAR UM ITEM DENTRO DE COMBOBOX
    If VALOR = "" Then
        comb.ListIndex = -1
        Exit Sub
    Else
        For blI = 0 To comb.ListCount
            If UCase(comb.List(blI)) = UCase(VALOR) Then
                comb.Text = comb.List(blI)
                Exit For
            End If
        Next blI
    End If
End Sub
Public Function testa_intervalo(vlIni As Integer, vlfim As Integer, comprimento As Boolean, grd As MSFlexGrid) As Boolean
    'FUNCAO PARA TESTAR SE O VALOR ESTÁ ENTRE O INTERVALO DIGITADO NO GRID
    'SE ESTIVER NO INTERVALO RETORNA TRUE, SE NAO ESTIVER RETORNA FALSE
    testa_intervalo = False
    For blI = 1 To grd.Rows - 1
        If comprimento Then
            If vlIni >= grd.TextMatrix(blI, 2) And vlIni <= grd.TextMatrix(blI, 3) Then
                testa_intervalo = True
                MsgBox "Comprimento Mínimo Dentro de um Intervalo Cadastrado", vbInformation
                Exit For
            End If
            If vlfim >= grd.TextMatrix(blI, 2) And vlfim <= grd.TextMatrix(blI, 3) Then
                testa_intervalo = True
                MsgBox "Comprimento Máximo Dentro de um Intervalo Cadastrado", vbInformation
                Exit For
            End If
        Else
            If vlIni >= grd.TextMatrix(blI, 0) And vlIni <= grd.TextMatrix(blI, 1) Then
                testa_intervalo = True
                MsgBox "Largura Mínima Dentro de um Intervalo Cadastrado", vbInformation
                Exit For
            End If
            If vlfim >= grd.TextMatrix(blI, 0) And vlfim <= grd.TextMatrix(blI, 1) Then
                testa_intervalo = True
                MsgBox "Largura Máxima Dentro de um Intervalo Cadastrado", vbInformation
                Exit For
            End If
        End If
    Next blI
    
End Function
'Public Sub calcula_dimensao(largmin As Masked, largmax As Masked, compmin As Masked, compmax As Masked, padrao As CheckBox, grd As MSFlexGrid)
'    Select Case grd.Rows
'        Case 1
'            largmin.Texto = 0
'            largmax.Texto = 2600
'            compmin.Texto = 0
'            compmax.Texto = 13000
'        Case 2
'            largmin.Texto = 2601
'            largmax.Texto = 3140
'            compmin.Texto = 0
'            compmax.Texto = 13000
'        Case 3
'            largmin.Texto = 3141
'            largmax.Texto = 3700
'            compmin.Texto = 0
'            compmax.Texto = 13000
'        Case 4
'            largmin.Texto = 3701
'            largmax.Texto = 4000
'            compmin.Texto = 0
'            compmax.Texto = 13000
'        Case 5
'            largmin.Texto = 0
'            largmax.Texto = 2600
'            compmin.Texto = 13001
'            compmax.Texto = 18000
'        Case 6
'            largmin.Texto = 2601
'            largmax.Texto = 3140
'            compmin.Texto = 13001
'            compmax.Texto = 18000
'        Case 7
'            largmin.Texto = 3141
'            largmax.Texto = 3700
'            compmin.Texto = 13001
'            compmax.Texto = 18000
'        Case 8
'            largmin.Texto = 3701
'            largmax.Texto = 4000
'            compmin.Texto = 13001
'            compmax.Texto = 18000
'        Case Else
'            padrao.Value = 0
'            bgpadrao = False
'    End Select
'End Sub
Public Sub calcula_dimensao(largmin As Masked, largmax As Masked, compmin As Masked, compmax As Masked, padrao As CheckBox, grd As MSFlexGrid, FlgPadrao As String, dblclick As Boolean)
    'ROTINA CRIADA PARA CALCULAR AS DIMENSÕES  PADRÕES DE CHAPAS GROSSAS
    'DA USIMINAS (U), COSIPA (C) E NENHUM (N)
    If dblclick Then Exit Sub
     If UCase(FlgPadrao) = "U" Then  'CHAPA PADRÃO USIMINAS
        Select Case grd.Rows
            Case 1
                largmin.Texto = 0
                largmax.Texto = 2600
                compmin.Texto = 0
                compmax.Texto = 13000
            Case 2
                largmin.Texto = 2601
                largmax.Texto = 3140
                compmin.Texto = 0
                compmax.Texto = 13000
            Case 3
                largmin.Texto = 3141
                largmax.Texto = 3700
                compmin.Texto = 0
                compmax.Texto = 13000
            Case 4
                largmin.Texto = 3701
                largmax.Texto = 4000
                compmin.Texto = 0
                compmax.Texto = 13000
            Case 5
                largmin.Texto = 0
                largmax.Texto = 2600
                compmin.Texto = 13001
                compmax.Texto = 18000
            Case 6
                largmin.Texto = 2601
                largmax.Texto = 3140
                compmin.Texto = 13001
                compmax.Texto = 18000
            Case 7
                largmin.Texto = 3141
                largmax.Texto = 3700
                compmin.Texto = 13001
                compmax.Texto = 18000
            Case 8
                largmin.Texto = 3701
                largmax.Texto = 4000
                compmin.Texto = 13001
                compmax.Texto = 18000
            Case Else
                padrao.Value = 0
                bgpadrao = False
                largmin.Texto = ""
                largmax.Texto = ""
                compmin.Texto = ""
                compmax.Texto = ""
        End Select
    ElseIf UCase(FlgPadrao) = "C" Then 'CHAPA PADRÃO COSIPA
        Select Case grd.Rows
            Case 1
                largmin.Texto = 0
                largmax.Texto = 2600
                compmin.Texto = 0
                compmax.Texto = 12600
            Case 2
                largmin.Texto = 0
                largmax.Texto = 2600
                compmin.Texto = 12601
                compmax.Texto = 18000
            Case 3
                largmin.Texto = 0
                largmax.Texto = 2600
                compmin.Texto = 18001
                compmax.Texto = 99999
            Case 4
                largmin.Texto = 2601
                largmax.Texto = 3500
                compmin.Texto = 0
                compmax.Texto = 12600
            Case 5
                largmin.Texto = 2601
                largmax.Texto = 3500
                compmin.Texto = 12601
                compmax.Texto = 18000
            Case 6
                largmin.Texto = 2601
                largmax.Texto = 3500
                compmin.Texto = 18001
                compmax.Texto = 99999
            Case 7
                largmin.Texto = 3501
                largmax.Texto = 3800
                compmin.Texto = 0
                compmax.Texto = 12600
            Case 8
                largmin.Texto = 3501
                largmax.Texto = 3800
                compmin.Texto = 12601
                compmax.Texto = 18000
            Case 9
                largmin.Texto = 3501
                largmax.Texto = 3800
                compmin.Texto = 18001
                compmax.Texto = 99999
            Case 10
                largmin.Texto = 3801
                largmax.Texto = 99999
                compmin.Texto = 0
                compmax.Texto = 12600
            Case 11
                largmin.Texto = 3801
                largmax.Texto = 99999
                compmin.Texto = 12601
                compmax.Texto = 18000
            Case 12
                largmin.Texto = 3801
                largmax.Texto = 99999
                compmin.Texto = 18001
                compmax.Texto = 99999
            Case Else
                padrao.Value = 0
                bgpadrao = False
                largmin.Texto = ""
                largmax.Texto = ""
                compmin.Texto = ""
                compmax.Texto = ""
        End Select
    End If
End Sub

'Retorna a Sigla do Item Selecionado na combo
Function RetSigla(ByRef cboSigla As ComboBox) As String
Dim avSiglas As Variant

avSiglas = Array("", "AR", "TR", "EX", "SU", "OP", "DF", "IN", "MN")

RetSigla = avSiglas(cboSigla.ItemData(cboSigla.ListIndex))
'AR = Aguardando recolhimento
'TR = Transito
'EX = Extraviado
'SU = Sucateado
'OP = Operação
'DF = Danificado
'IN = Inoperante
'MN = Manutenção

End Function

'Retorna a Sigla do Item Selecionado na combo
Function RetDescSigla(ByVal strSigla As String) As String
    Select Case strSigla
            Case "AR" 'AR = Aguardando recolhimento
                RetDescSigla = "Aguardando recolhimento"
            Case "TR" 'TR = Transito
                RetDescSigla = "Transito"
            Case "EX" 'EX = Extraviado
                RetDescSigla = "Extraviado"
            Case "SU" 'SU = Sucateado
                RetDescSigla = "Sucateado"
            Case "OP" 'OP = Operação
                RetDescSigla = "Operação"
            Case "DF" 'DF = Danificado
                RetDescSigla = "Danificado"
            Case "IN" 'IN = Inoperante
                RetDescSigla = "Inoperante"
            Case "MN" 'MN = Manutenção
                RetDescSigla = "Manutenção"
            Case Else
                RetDescSigla = "INDETERMINADO"
    End Select
End Function
Public Function Extenso(nValor)
'*********************************************************
'**      Rotina para Exibir os forms no MDISisTrans     **
'**   Criado por Carlos Fabiano - O Bom em 19/11/2003   **
'**                     USIFAST - CPD                   **
'*********************************************************
On Error Resume Next
  'Faz a validação do argumento
  If IsNull(nValor) Or nValor <= 0 Or nValor > 999999999.99 Then
    Exit Function
  End If

  'Declara as variáveis da função
  Dim nContador, nTamanho As Integer
  Dim cValor, cParte, cFinal As String
  ReDim agrupo(4), ATEXTO(4) As String

  'Define matrizes com extensos parciais
  ReDim aUnid(19) As String
  aUnid(1) = "UM ": aUnid(2) = "DOIS ": aUnid(3) = "TRES "
  aUnid(4) = "QUATRO ": aUnid(5) = "CINCO ": aUnid(6) = "SEIS "
  aUnid(7) = "SETE ": aUnid(8) = "OITO ": aUnid(9) = "NOVE "
  aUnid(10) = "DEZ ": aUnid(11) = "ONZE ": aUnid(12) = "DOZE "
  aUnid(13) = "TREZE ": aUnid(14) = "QUATORZE ": aUnid(15) = "QUINZE "
  aUnid(16) = "DEZESSEIS ": aUnid(17) = "DEZESSETE ": aUnid(18) = "DEZOITO "
  aUnid(19) = "DEZENOVE "

  ReDim aDezena(9) As String
  aDezena(1) = "DEZ ": aDezena(2) = "VINTE ": aDezena(3) = "TRINTA "
  aDezena(4) = "QUARENTA ": aDezena(5) = "CINQUENTA "
  aDezena(6) = "SESSENTA ": aDezena(7) = "SETENTA ": aDezena(8) = "OITENTA "
  aDezena(9) = "NOVENTA "

  ReDim aCentena(9) As String
  aCentena(1) = "CENTO ":  aCentena(2) = "DUZENTOS "
  aCentena(3) = "TREZENTOS ": aCentena(4) = "QUATROCENTOS "
  aCentena(5) = "QUINHENTOS ": aCentena(6) = "SEISCENTOS "
  aCentena(7) = "SETECENTOS ": aCentena(8) = "OITOCENTOS "
  aCentena(9) = "NOVECENTOS "
  
  'Divide o valor em vários grupos
  cValor = Format$(nValor, "0000000000.00")
  agrupo(1) = Mid$(cValor, 2, 3)
  agrupo(2) = Mid$(cValor, 5, 3)
  agrupo(3) = Mid$(cValor, 8, 3)
  agrupo(4) = "0" & Mid$(cValor, 12, 2)
  
  'Processa cada grupo
  For nContador = 1 To 4
    cParte = agrupo(nContador)
    nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
    If nTamanho = 3 Then
      If Right$(cParte, 2) <> "00" Then
        ATEXTO(nContador) = ATEXTO(nContador) + aCentena(Left(cParte, 1)) & "E "
        nTamanho = 2
      Else
        ATEXTO(nContador) = ATEXTO(nContador) + IIf(Left$(cParte, 1) = "1", "CEM ", aCentena(Left(cParte, 1)))
      End If
    End If
    If nTamanho = 2 Then
      If Val(Right(cParte, 2)) < 20 Then
        ATEXTO(nContador) = ATEXTO(nContador) + aUnid(Right(cParte, 2))
      Else
        ATEXTO(nContador) = ATEXTO(nContador) + aDezena(Mid(cParte, 2, 1))
        If Right$(cParte, 1) <> "0" Then
          ATEXTO(nContador) = ATEXTO(nContador) & "E "
          nTamanho = 1
        End If
      End If
    End If
    If nTamanho = 1 Then
      ATEXTO(nContador) = ATEXTO(nContador) + aUnid(Right(cParte, 1))
    End If
  Next

  'Gera o formato final do texto
  If Val(agrupo(1) + agrupo(2) + agrupo(3)) = 0 And Val(agrupo(4)) <> 0 Then
    cFinal = ATEXTO(4) + IIf(Val(agrupo(4)) = 1, "CENTAVO", "CENTAVOS")
  Else
    cFinal = ""
    cFinal = cFinal + IIf(Val(agrupo(1)) <> 0, ATEXTO(1) + IIf(Val(agrupo(1)) > 1, "MILHÕES ", "MILHÃO "), "")
    If Val(agrupo(2) + agrupo(3)) = 0 Then
      cFinal = cFinal & "DE "
    Else
      cFinal = cFinal + IIf(Val(agrupo(2)) <> 0, ATEXTO(2) & "MIL ", "")
    End If
    cFinal = cFinal + ATEXTO(3) + IIf(Val(agrupo(1) + agrupo(2) + agrupo(3)) = 1, "REAL ", "REAIS ")
    cFinal = cFinal + IIf(Val(agrupo(4)) <> 0, "E " & ATEXTO(4) + IIf(Val(agrupo(4)) = 1, "CENTAVO", "CENTAVOS"), "")
  End If
  Extenso = cFinal
End Function

Public Sub SetaFlagFilial()
'ROTINA PARA CARREGAR OS PARAMETROS DA FILIAL EM VARIAVEIS ALOCADAS EM MEMORIA
'CRIADA POR RANDERSON
'ALTERADA POR CARLOS FABIANO - 25/11/2003
On Error Resume Next
   sgQuery = " select *,b.codUF from filial a, cidade b" & _
             " Where a.codcid = b.codcid and CodFil = " & igCodFil
   Call consulta(sgQuery)
   With Rs
      If Not .EOF Then
         igCodFilAnt = IIf(IsNull(!CodFilAnt), 0, !CodFilAnt)
         sgCodCCFil = "" & !CodCCFil
         sgCNPJFil = "" & !CNPJFIL
         sgRepomFil = "" & !RepomFil
         sgDigAprFil = "" & !DigAprFil
         sgDigLibFil = "" & !DigLibFil
         sgBerUsiFil = "" & !BerUsiFil
         igForConFil = IIf(IsNull(!ForConFil), 0, !ForConFil)
         igForNFFil = IIf(IsNull(!ForNFFil), 0, !ForNFFil)
         igForRFFil = IIf(IsNull(!ForRFFil), 0, !ForRFFil)
         igForFatFil = IIf(IsNull(!ForFatFil), 0, !ForFatFil)
         igForOCFil = IIf(IsNull(!ForOCFil), 0, !ForOCFil)
         lgPerDesPneuFil = IIf(IsNull(!PerDesPneuFil), 0, !PerDesPneuFil)
         sgImpISSQNFil = "" & !ImpISSQNFil
         sgImpINSSFil = "" & !ImpINSSFil
         sgFreRetFrota = "" & !FreRetFrota
         sgGuiCarFil = "" & !GuiCarFil
         sgProGuiCarFil = "" & !ProGuiCarFil
         dgDatProFil = !DatProFil
         igAlqINSSFil = IIf(IsNull(!AlqINSSFil), 0, !AlqINSSFil)
         igAlqISSFil = IIf(IsNull(!AlqISSFil), 0, !AlqISSFil)
         igEstMinBerUsi = IIf(IsNull(!EstMinBerUsi), 0, !EstMinBerUsi)
         igEstMinBerInt = IIf(IsNull(!EstMinBerInt), 0, !EstMinBerInt)
         sgFlgDigPed = "" & !FlgDigPed
         sgPedIncFre = "" & !PedIncFre
         sgFlgPrgUsi = "" & !FlgPrgUsi
         sgFlgNumBol = "" & !FlgNumBol
         sgFlgNumPrg = "" & !FlgNumPrg
         sgFlgDimPrg = "" & !FlgDimPrg
         sgFlgItePrg = "" & !FlgItePrg
         sgFlgRotMaxion = "" & !FlgRotMaxion
         sgFlgCgaNfUsi = "" & !FlgCgaNfUsi
         igPerAdiFro = IIf(IsNull(!PerAdiFro), 0, !PerAdiFro)
         igPerAdiTer = IIf(IsNull(!PerAdiTer), 0, !PerAdiTer)
         sgFlgAltPso = "" & !FlgAltPso
         sgUfFil = !CodUf
      End If
      sgQuery = "Select NomEmp,PerSegMet from empresa where codemp = 1"
      consulta sgQuery
      If Not Rs.EOF Then
        sgNomeEmp = Rs("NomEmp")
        dgPerSegMet = Rs("PerSegMet")
      Else
        sgNomeEmp = ""
        dgPerSegMet = 0
      End If
   End With
End Sub
'Public Function NoRound(vValor As Double, vCasasDec As Integer) As Double
''*********************************************************
''** Rotina para Truncar casas decimais de um valor Real **
''**      Criado por Randerson Maurilio em 26/11/2003    **
''**                     USIFAST - CPD                   **
''*********************************************************
''Exemplo de Como utilizar :
'' label1.caption = NoRound(Text1.text,2)
''   Onde Text1.text = 15,339
'' --> Retorna 15,33
''Dim vValAux As Double
''   vValAux = vValor * (10 ^ vCasasDec)
''   NoRound = Format(Int(vValAux) / (10 ^ vCasasDec), "#,##0." & String(vCasasDec, "0"))
'    NoRound = Format(Mid(vValor, 1, InStr(vValor, ",") + vCasasDec), "#,##0." & String(vCasasDec, "0"))
'End Function
Public Function NoRound(vValor As Double, vCasasDec As Integer) As Double
'*********************************************************
'** Rotina para Truncar casas decimais de um valor Real **
'**      Criado por Randerson Maurilio em 26/11/2003    **
'**                     USIFAST - CPD                   **
'*********************************************************
'Exemplo de Como utilizar :
' label1.caption = NoRound(Text1.text,2)
'   Onde Text1.text = 15,339
' --> Retorna 15,33
    If InStr(vValor, ",") = 0 Then
       NoRound = vValor
    Else
       NoRound = Format(Mid(vValor, 1, InStr(vValor, ",") + vCasasDec), "#,##0." & String(vCasasDec, "0"))
    End If
End Function


Public Function UltimoDiaMes(vData As Date) As Integer
'*********************************************************
'** Rotina para retornar o último dia do mes informado  **
'**     Criado por Randerson Maurilio em 07/01/2004     **
'**                     USIFAST - CPD                   **
'*********************************************************
On Error Resume Next
   Dim ilUltDia
   'Alimenta o vetor com os últimos dias dos meses
   ilUltDia = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
   'Verifica se o ano é bissexto
   If Month(vData) = 2 And (Year(vData) Mod 4) = 0 Then
      UltimoDiaMes = 29
   Else
      UltimoDiaMes = ilUltDia(Month(vData) - 1)
   End If
End Function
Public Function FormataProposta(vProp As String) As String
'*********************************************************
'**      Rotina para formatar o número da proposta      **
'**     Criado por Randerson Maurilio em 16/01/2004     **
'**                     USIFAST - CPD                   **
'*********************************************************
On Error Resume Next
   FormataProposta = Mid(vProp, 1, 4) & "-" & Format(Mid(vProp, 5, 4), "0000") & _
   "/" & Format(Mid(vProp, 9, 2), "00")
End Function

