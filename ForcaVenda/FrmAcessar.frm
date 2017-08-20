VERSION 5.00
Begin VB.Form FrmAcessar 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Unocann - For�a de Venda"
   ClientHeight    =   3330
   ClientLeft      =   1185
   ClientTop       =   5490
   ClientWidth     =   7140
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "FrmAcessar.frx":0000
   ScaleHeight     =   3330
   ScaleWidth      =   7140
   Begin VB.TextBox MskOperador 
      Height          =   435
      Left            =   960
      MaxLength       =   9
      TabIndex        =   0
      Top             =   420
      Width           =   1380
   End
   Begin VB.CommandButton CmdAltSenha 
      BackColor       =   &H80000013&
      Caption         =   "Alterar &Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton CmdSair 
      BackColor       =   &H80000013&
      Caption         =   "Sai&r"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton CmdAcessar 
      BackColor       =   &H80000013&
      Caption         =   "&Acessar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox TxtPwdoper 
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   960
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1140
      Width           =   1380
   End
   Begin VB.Label lblManut 
      BackColor       =   &H000000FF&
      Caption         =   "Manuten��o do banco de dados - Favor aguardar um momento..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label LblNomeOper 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   2430
      TabIndex        =   7
      Top             =   360
      Width           =   2790
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   60
      TabIndex        =   6
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usu�rio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   825
   End
End
Attribute VB_Name = "FrmAcessar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sgSenUsuSis  As String
Dim sgDatUltAce As String
Dim blErro As Byte

Private Sub CompactaBanco()

    '*****************************************************************************
    'A compacta��o acontece uma vez por m�s.
    '*****************************************************************************
    
    '*****************************************************************************
    'Consulta o m�s de atualiza��o de todas as seq��ncias de pedido atribu�das ao
    'representante.
    '*****************************************************************************
    
    sgQuery = "select month(datatu) as Mes from SEQUENCIA_PEDIDO"
    
    Consulta sgQuery
    
    '*****************************************************************************
    'Se n�o encontrar nenhuma seq��ncia, o representante n�o pode emitir pedidos.
    'A rotina � abandonada.
    '*****************************************************************************
    
    If Rs.EOF = True Then
        Exit Sub
    End If
    
    '*****************************************************************************
    'Se o usu�rio logado for o representante e se o m�s de atualiza��o da
    'seq��ncia for diferente do m�s atual, o programa exibe um r�tulo avisando
    'sobre a manuten��o e executa as senten�as que fazem a compacta��o. A quest�o
    'do m�s diferente envolve o fato de a opera��o ocorrer uma vez por m�s.
    '*****************************************************************************
    
    If APLICA = 1 Then
    
        If Rs!Mes <> Month(Date) Then
        
'            lblManut.Visible = True
'
'            sgQuery = "update SEQUENCIA_PEDIDO set datatu = convert(datetime,getdate(),103)"
'            Conexao.Execute sgQuery
'
'            sgQuery = "BACKUP LOG unocann WITH TRUNCATE_ONLY"
'            Conexao.Execute sgQuery
'
'            sgQuery = "DBCC SHRINKFILE (N'unocann_log' , 0, TRUNCATEONLY)"
'            Conexao.Execute sgQuery
'
'            sgQuery = "DBCC SHRINKDATABASE (N'unocann' , 0, TRUNCATEONLY)"
'            Conexao.Execute sgQuery
          
        End If
        
        Rs.Close
        
        Set Rs = Nothing
    
    End If

End Sub

Private Sub CmdAcessar_Click()
    
    '*****************************************************************************
    'Informa a senha digitada pelo usu�rio como par�metro.
    '*****************************************************************************
    
    Acessar sgSenUsuSis
    
End Sub

Private Sub CmdAltSenha_Click()

    If Trim(MskOperador.Text) <> "" Then
        
        FrmAlterarSenha.Show
        
    Else
        
        MsgBox "Digite o Nome Usu�rio", vbInformation
        MskOperador.SetFocus
        
    End If
    
End Sub

Private Sub CmdSair_Click()
    
    FechaConexao
    
    End
    
End Sub

Private Sub Form_Load()

'strSenha = Crypt(ReadINI("Geral", "Primeira", App.Path & "\ProjUno.ini"))


    '*****************************************************************************************
    'Primeira rotina a ser executada no sistema.
    '*****************************************************************************************
    
    '*****************************************************************************************
    'Se o programa j� estiver sendo executado, a nova tentativa de abertura � abortada.
    '*****************************************************************************************
    
    If App.PrevInstance Then
        End
    End If

    '*****************************************************************************************
    'Abre o arquivo Dba.sys, que armazena o c�digo do representante que est� se logando.
    '*****************************************************************************************
    
    Open "C:\Windows\Dba.sys" For Input As #1
   
    If Not EOF(1) Then
        
        Line Input #1, sgRepresentante
    
    Else
        
        MsgBox "Arquivo de configura��o do sistema inexistente, Consulte o Administrador do Sistema", vbCritical
        
        Unload Me
        
    End If
    
    Close #1
    
    'sgRepresentante = 999
    
    '*****************************************************************************************
    'Se o c�digo encontrado no arquivo for 999, significa que o usu�rio � um Administrador e o
    'valor 0 (zero) da vari�vel APLICA indica isso. Se o valor for qualquer outro, entende-se
    'que o usu�rio � um representante qualquer.
    '*****************************************************************************************
   
    If sgRepresentante = 999 Then
        APLICA = 0
    End If
    
    If (sgRepresentante = 2 Or sgRepresentante = 7 Or sgRepresentante = 8 Or sgRepresentante = 10 Or sgRepresentante = 600 Or sgRepresentante = 800 Or sgRepresentante = 905 Or sgRepresentante = 1001 Or sgRepresentante = 1900 Or sgRepresentante = 2100 Or sgRepresentante = 5000 Or sgRepresentante = 5001 Or sgRepresentante = 6000 Or sgRepresentante = 7050 Or sgRepresentante = 7060) Then
        APLICA = 1
    Else
        APLICA = 2
    End If
    
    'sgRepresentante = 9999
    
    '*****************************************************************************************
    'Armazena datas de modo que seja poss�vel calcular a diferen�a de dias entre a data atual
    'e o in�cio do m�s corrente.
    '*****************************************************************************************
    
    datah = Date
    data1 = "01/" & Month(datah) & "/" & Year(datah)

    '*****************************************************************************************
    'Executa fun��o que ajusta a janela no canto superior esquerdo do MDI.
    '*****************************************************************************************

    AjustaJanela Me, 7260, 3840, 3000, 3500
    
    '*****************************************************************************************
    '
    '*****************************************************************************************
    
    bgSenOK = False
    
    If bgSenComi = False Then
        
        AbreConexao
        
        LblNomeOper = ""
        
    Else
        
        MskOperador.Text = LgCodUsuSis
        LblNomeOper.Caption = sgNomUsuSis
        MskOperador.Enabled = False
        CmdAltSenha.Enabled = False
                
    End If
            
    blErro = 1
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    '*****************************************************************************
    'Emula o pressionamento da tecla TAB cada vez que ENTER for pressionada com o
    'formul�rio ativo.
    '*****************************************************************************
    
    On Error GoTo TrataErro
    
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
    End If
    
    Exit Sub
    
TrataErro:

    Rotina_Erro "Form_KeyPress"
    
End Sub

Private Sub MskOperador_Change()

    If MskOperador.Text = "" Then
        LblNomeOper.Caption = ""
        LgCodUsuSis = 0
    End If
    
End Sub

Private Sub MskOperador_KeyPress(KeyAscii As Integer)
    
    If MskOperador.Text = "" Then
        LblNomeOper.Caption = ""
        LgCodUsuSis = 0
        MskOperador.TabIndex = 0
    End If
    
End Sub

Private Sub MskOperador_LostFocus()
    
    If bgSenComi = True Then
        Exit Sub
    End If
    
    '*****************************************************************************
    'O campo de identifica��o do usu�rio suporta logon ou c�digo do cliente.
    '*****************************************************************************
    
    '*****************************************************************************
    'Se houver algo digitado no campo de identifica��o do usu�rio, o programa vai
    'consultar essa informa��o na tabela de usu�rios a fim de tentar encontrar
    'registro correspondente. Se n�o houver nada digitado no campo, o programa
    'apenas limpa o campo que deveria trazer o nome do usu�rio reconhecido.
    '*****************************************************************************
    
    If Trim(MskOperador.Text) <> "" Then
        
        sgQuery = "select * from usuario"
        sgQuery = sgQuery & " Where  (codusu='" & Val(MskOperador.Text) & "' or logUsu = '" & Trim(MskOperador.Text) & "')"
        
        Consulta sgQuery
        
        '*************************************************************************
        'Se o usu�rio tiver sido identificado, seus dados s�o armazenados pelo
        'programa e seu nome � exibido na tela. Caso contr�rio surgir� uma
        'mensagem avisando que tal usu�rio n�o est� cadastrado.
        '*************************************************************************
        
        '*************************************************************************
        'Pode acontecer de o usu�rio ser identificado e estar com seu acesso
        'bloqueado. Nesse caso mostra-se uma mensagem que informa a situa��o e a
        'identifica��o � interrompida.
        '*************************************************************************
        
        If Not Rs.EOF Then
            
            LblNomeOper.Caption = Trim(Rs("NomUsu"))
            
            sgNomUsuSis = Trim(Rs("NomUsu"))
            sgFlgUsu = Trim(Rs("FlgUsu"))
            LgCodUsuSis = Rs("CodUsu")
            
            If UCase(Rs("SitUsu")) <> "S" Then
                
                MsgBox "Usu�rio Com Acesso Bloqueado", vbInformation
                
                Exit Sub
                
            End If
            
            DoEvents
            
        Else
            
            MsgBox "Usu�rio N�o Cadastrado", vbInformation
            
            Unload Me
            
            Me.Show
            
            Exit Sub
            
        End If
        
    Else
        
        LblNomeOper.Caption = ""
        
    End If
    
End Sub

Private Sub Acessar(slsenhacript As String)
    
    '*****************************************************************************
    'A fun��o recebe como par�metro a senha digitada pelo usu�rio.
    '*****************************************************************************
    
    Dim slletra As String
    Dim ilok As Integer
    
    '*****************************************************************************
    'O evento LostFocus de mskOperador identifica o usu�rio que est� se logando.
    '*****************************************************************************
    
    MskOperador_LostFocus
    
    '*****************************************************************************
    'Consiste os campos "Usu�rio" e "Senha".
    '*****************************************************************************
            
    If Trim(MskOperador.Text) = "" Then
    
        MsgBox "Digite um Usu�rio", vbInformation
        
        MskOperador.SetFocus
        
        Exit Sub
        
    ElseIf Trim(TxtPwdoper.Text) = "" Then
    
        MsgBox "Digite uma Senha", vbInformation
        
        TxtPwdoper.SetFocus
        
        Exit Sub
    
    End If
    
    '*****************************************************************************
    'Pesquisa o banco de dados para saber se a senha digitada para o usu�rio
    'informado � correta.
    '*****************************************************************************
    
    If slsenhacript <> "" Then
    
        sgQuery = "SELECT PWDCOMPARE('" & Trim(slsenhacript) & "',SenUsu, 0) AS Senha_OK, DatUltAce from usuario where codusu = " & LgCodUsuSis
        
        Consulta sgQuery
                        
        ilok = Rs("Senha_OK")
        
        '*************************************************************************
        'Se a senha estiver correta, o sistema verifica se � o primeiro acesso do
        'usu�rio. Se for, abre uma janela para que a senha seja trocada; sen�o,
        'atualiza a data do �ltimo acesso, compacta o banco de dados e abre o MDI.
        '*************************************************************************
        
        If Rs("Senha_OK") = 1 Then
        
            If bgSenComi = False Then
                
                If Not IsNull(Rs("DatUltAce")) Then
                    
                    sgQuery = "update usuario set DatUltAce  = convert(datetime,'" & Date & "',103)"
                    sgQuery = sgQuery & " where CodUsu = " & LgCodUsuSis
                    
                    Set Rs = Conexao.Execute(sgQuery)
                    Set Rs = Nothing
                    
                    Unload FrmAcessar
                    
                    DoEvents
                    
                    CompactaBanco
                    
                    MDIProjUNO.Show
                    
                Else
                    
                    MsgBox "Primeiro Acesso ao Sistema. Altere a Sua Senha", vbInformation
                    
                    FrmAlterarSenha.Show
                    
                End If
                
            End If
        
        Else
        
            '*********************************************************************
            'A cada tentativa errada de logon o sistema emite uma mensagem. Na
            'quinta tentativa o programa bloqueia o acesso do usu�rio informado e
            'apenas o Administrador poder� liber�-lo novamente.
            '*********************************************************************
        
            If blErro < 5 Then
            
                MsgBox "Senha Inv�lida", vbInformation
                
                TxtPwdoper.Text = ""
                TxtPwdoper.SetFocus
                
                blErro = blErro + 1
                
            Else
            
                sgQuery = "update usuario set DatUltAce  = convert(datetime,'" & Date & "',103),"
                sgQuery = sgQuery & " SitUsu = 'N' where CodUsu = " & LgCodUsuSis
                
                Set Rs = Conexao.Execute(sgQuery)
                Set Rs = Nothing
                
                MsgBox "Seu Acesso foi Bloqueado, Entre em Contato com o Adminstrador do Sistema ", vbInformation
                
                If bgSenComi = False Then
                    
                    TxtPwdoper.Text = ""
                    TxtPwdoper.SetFocus
                    
                Else
                    
                    Unload Me
                    
                    bgSenComi = False
                    
                    Set FrmAcessar = Nothing
                    
                    Exit Sub
                    
                End If
                
            End If
            
        End If
        
    End If
    
End Sub

Private Sub TxtPwdoper_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And TxtPwdoper.Text <> "" Then
        
        sgSenUsuSis = UCase(TxtPwdoper.Text)
        
        Acessar sgSenUsuSis
        
    End If
    
End Sub

Private Sub TxtPwdoper_LostFocus()

    sgSenUsuSis = UCase(TxtPwdoper.Text)
    
End Sub
