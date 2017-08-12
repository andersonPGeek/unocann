Attribute VB_Name = "MdlConectaDb"
'*****************************************************************************************
'* DESCRIPTION:
'* This is an example of how to make a TextBox or a RichTextBox scroll when Disabled. This file
'* can also be used if you want to scroll an Enabled TextBox without using the scrollbars (From
'* outside of the control.), and without moving the Caret.
'*****************************************************************************************
'*****************************************************************************************
'* USAGE:
'* Num& = ScrollText&(OBJECT, LINES#)
'* #=Number of lines to scroll. Positive Numbers
'* (eg. 2) go down, and Negative Numbers (eg. -2)
'* go up.
'*****************************************************************************************

#If Win32 Then
    Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
#Else
    Declare Function PutFocus% Lib "user" Alias "SetFocus" (ByVal hWd%)
    Declare Function SendMessage& Lib "user" (ByVal hWd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
#End If

Function ScrollText&(TextBox As Control, vLines As Integer)
    
    'The Windows Version Stuff...
    
    #If Win32 Then
        Dim Success As Long
        Dim SavedWnd As Long
        Dim R As Long
    #Else
        Dim Success As Integer
        Dim SavedWnd As Integer
        Dim R As Integer
    #End If
    
    Const EM_LINESCROLL = &HB6
    
    'Get the window handle of the control that
    'currently has the focus (eg. Command1). This
    'is so that the control that had the focus when
    'clicked gets the focus back. In this case on of the
    '[U] of [D] buttons.
    
    SavedWnd = Screen.ActiveControl.hwnd
    Lines& = vLines
    
    'Remove the comment (') if your TextBox is !ENABLED!
    '# TextBox.SetFocus
    
    'Scroll the lines, using the SendMessage.
    Success = SendMessage(TextBox.hwnd, EM_LINESCROLL, 0, Lines&)
    
    'Restore the focus to the original control which had
    'the previous focus, which we previously recorded before.
    R = PutFocus(SavedWnd)
    
    'Return the number of lines actually scrolled.
    ScrollText& = Success

End Function

Public Sub AbreConexao()

    '*****************************************************************************
    'Abre a conexão com banco de dados. Utiliza uma conexão ODBC chamada "Unocann"
    'e o usuário e senha do SQL Server.
    '*****************************************************************************
    
    On Error GoTo Erro
    
    sgservidor = "Ambiente Desenvolvimento"
    
    Set Conexao = New ADODB.Connection
'
'     Conexao.Open "Provider=SQLNCLI; " & _
'              "Initial Catalog=UNOCANN; " & _
'              "Data Source=UNOMOBILE010\SQLEXPRESS; " & _
'              "Integrated security=SSPI; Persist Security info=true;"
'''''
    With Conexao
        
        Select Case sgRepresentante
           Case 2 'Evaldo
                If Environ("COMPUTERNAME") = "UNOCANN-PC" Then
                    .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=UNOCANN-PC;DATABASE=Unocann;UID=sa;PWD=unocann2017"
                    '.ConnectionString = "PROVIDER=SQLOLEDB;SERVER=201.65.158.20;DATABASE=Unocann;UID=sa;PWD=#unoforte5600!"
                Else
                    .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=UNOMOBILE001\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
                End If
           Case 7 'Galvão
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=UNOMOBILE05;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case 8 'Luciano
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=LUCIANO-PC\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
               '  .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=UNOMOBIL13\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case 10 'Antonio Madureira
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=LUCASRAMOS-PC;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case 600 'Baiao
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=UNOMOBILE04\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case 800 'Adalton
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=ADAUTON-PC\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
               '  .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=UNOMOBILE13\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case 905 'Mara
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=UNOMOBILE11\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case 1001 'Marcio Ribas
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=UNOMOBILE002\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
            Case 1900 'Henrique
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=HENRIQUE-PC\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
            Case 2100 'Jose Julio
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=UNOMOBILE08\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case 5000 'Jorge Leandro
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=JORGELEANDRO-PC\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case 5001 'Everaldo
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=UNOMOBILE015\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case 6000 'Reis
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=UNOMOBILE09\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case 7050 'Freitas
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=JOSECARLOSREP\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case 7075 'CEBE
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=CEBE-NOTE\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case 7060 'Alcir
                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=NOTE-ALACIR\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=sysadmpss1"
           Case Else
'                .ConnectionString = "PROVIDER=SQLOLEDB;SERVER=ti01\SQLEXPRESS;DATABASE=Unocann"
.ConnectionString = "PROVIDER=SQLOLEDB;SERVER=UNOCANN-PC;DATABASE=Unocann;UID=sa;PWD=unocann2017"
'.ConnectionString = "PROVIDER=SQLOLEDB;SERVER=192.168.254.11\SQLEXPRESS;DATABASE=Unocann;UID=sa;PWD=#unoforte5600!"
           End Select
        .Open
    End With

'Conexao.Open "Provider=SQLNCLI; " & _
'              "Initial Catalog=UNOCANN; " & _
'              "Data Source=(local)\SQLEXPRESS; " & _
'              "Integrated security=SSPI; Persist Security info=true;"


    Exit Sub
    
Erro:

    If Err.Number = 53 Then
        
        MsgBox "Não Existe o Arquivo de Acesso ao Banco de Dados" & Chr(13) & "Consulte o Administrador do Sistema", vbCritical
        
        End
        
    Else
                       
        End
        
    End If
    
End Sub

Public Sub FechaConexao()

    '*****************************************************************************
    'Encerra a conexão e o recordset aberto.
    '*****************************************************************************

    Conexao.Close
    
    Set Rs = Nothing
    Set Conexao = Nothing
    
End Sub

Public Sub Consulta(sql As String)

    '*****************************************************************************************
    'Executa consultas do tipo somente leitura na base de dados.
    '*****************************************************************************************

    Set Rs = Nothing
    Set Rs = New ADODB.Recordset

    With Rs
        .ActiveConnection = Conexao
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Source = sql
        .Open
    End With

End Sub

Public Sub Consulta2(sql As String)
    
    Set Rs2 = Nothing
    Set Rs2 = New ADODB.Recordset ' INSTANCIAR VARIAVEL ASSOCIADA AO BANCO DE DADOS
    
    With Rs2
        .CursorType = adOpenKeyset 'TIPO DE NAVEGACAO
        .LockType = adLockReadOnly  'FAZ A CONSULTA "SOMENTE LEITURA NA TABELA"
        .Source = sql 'String de comandos ao Banco de Dados
        .ActiveConnection = Conexao ' ATIVA A CONEXAO COM O BANCO DE DADOS
        .Open 'ABRE O BANCO DE DADOS
    End With
    
End Sub

Public Sub Rotina_Erro(sl_Lugar As String)

    'Mostra todas as mensagens de erro
    
    Dim sl_strErro As String, il_Msg As VbMsgBoxStyle, Erro As ADODB.Error

    If Conexao.Errors.Count = 0 Then
    
     '   Beep
        
        Select Case Err.Number
        
            Case 76:
                
                MsgBox sl_Lugar & vbCr & "A IMPRESSORA não foi encontrada pelo sistema!" & vbCr & Err.Number & " - " & Err.Description & vbCr & "Anote este número e avise o Administrador do Sistema imediatamente.", vbExclamation
                
            Case Else
                
                MsgBox sl_Lugar & vbCr & Err.Number & " - " & Err.Description & vbCr & "Anote este número e avise o Administrador do Sistema imediatamente.", vbExclamation
                
        End Select
        
    Else
    
        For Each Erro In Conexao.Errors
        
            sl_strErro = ""
            
            il_Msg = vbExclamation
            
            Select Case Erro.NativeError
            
                Case 547
                
                    If InStr(1, Erro.Description, "UPDATE", vbTextCompare) > 0 Then
                        sl_strErro = sl_strErro & "Não foi possível atualizar esse registro. Existem movimentações relacionadas a ele em outra(s) tabela(s)"
                    ElseIf InStr(1, Erro.Description, "DELETE", vbTextCompare) > 0 Then
                        sl_strErro = sl_strErro & "Não foi possível deletar esse registro. Existem movimentações relacionadas a ele em outra(s) tabela(s)"
                    End If
                
                Case 2627
                    
                    sl_strErro = sl_strErro & "Não foi possível incluir esse registro. Tentativa de inclusão de um registro já existente"
                
                Case Else
                
                    sl_strErro = "Erro " & Erro.NativeError & " em " & sl_Lugar & ". Operação cancelada. " & vbCrLf & vbCrLf
                    sl_strErro = sl_strErro & Erro.Description & vbCrLf & vbCrLf
                    sl_strErro = sl_strErro & "Anote este número e avise o Administrador do Sistema imediatamente."
                    
                    il_Msg = vbCritical
                    
            End Select
            
            Exit For
            
        Next
        
        Beep
        
        MsgBox sl_strErro, il_Msg, "A T E N Ç Ã O"
        
    End If
    
    Screen.MousePointer = 0
    
    On Error Resume Next
    
    Conexao.RollbackTrans
    
    Set Rs = Nothing
    Set Cmd = Nothing
    
End Sub

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

Public Sub EventoEnter(vKeyAscii As Integer)

    '*****************************************************************************
    'Rotina para evitar que a ação da tecla ENTER não prejudique o preenchimento
    'do campo "Observações" do programa. Como ENTER emula TAB em todo o
    'formulário, sua utilização num campo com múltiplas linhas apenas moveria o
    'foco, ao invés de aplicar uma quebra de linha.
    '*****************************************************************************
    
    If TypeOf Screen.ActiveControl Is TextBox Then
    
        If Screen.ActiveControl.MultiLine = True Then
            Exit Sub
        End If
        
    End If
        
    Select Case vKeyAscii
            
        Case vbKeyReturn
            
            vKeyAscii = 0
            SendKeys "{Tab}"
            
        Case vbKeyEscape
            
            vKeyAscii = 0
            SendKeys "+{Tab}"
        
    End Select
        
End Sub

Public Sub AjustaJanela(Formulario As Form, h As Long, l As Long, pos_sup As Long, pos_lat As Long)
    
    '*****************************************************************************************
    'Ajusta MDI Childs na área do MDI de acordo com as medidas informadas.
    '*****************************************************************************************
    
    Formulario.Height = l
    Formulario.Width = h
    Formulario.Top = pos_sup
    Formulario.Left = pos_lat
    
End Sub


