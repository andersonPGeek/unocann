VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmNotificacoes 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de Notificações"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   10335
   Begin InetCtlsObjects.Inet intTransfer 
      Left            =   9240
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSair 
      Height          =   735
      Left            =   9120
      Picture         =   "frmNotificacoes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Sair"
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdLimpar 
      Height          =   735
      Left            =   9120
      Picture         =   "frmNotificacoes.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Limpar"
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdNotificar 
      Height          =   735
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Notificar"
      Top             =   240
      Width           =   855
   End
   Begin VB.CheckBox chkExigirAlteracaoPedido 
      BackColor       =   &H80000018&
      Caption         =   "Esta notificação exige alteração no pedido?"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   5040
      Width           =   3495
   End
   Begin VB.TextBox txtSolucao 
      Height          =   1035
      Left            =   4200
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3720
      Width           =   4215
   End
   Begin VB.TextBox txtRazaoDetalhada 
      Height          =   1035
      Left            =   4200
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Frame fraRazaoNotificacao 
      BackColor       =   &H80000018&
      Caption         =   "Razão da Notificação"
      Height          =   2895
      Left            =   360
      TabIndex        =   19
      Top             =   1920
      Width           =   3495
      Begin VB.CheckBox chkForaPoliticaComercial 
         BackColor       =   &H80000018&
         Caption         =   "Descontos, prazos ou preços fora da política comercial"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Tag             =   "2"
         Top             =   750
         Width           =   2895
      End
      Begin VB.CheckBox chkDuplicatasAtraso 
         BackColor       =   &H80000018&
         Caption         =   "Duplicatas em atraso"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Tag             =   "3"
         Top             =   1380
         Width           =   2655
      End
      Begin VB.CheckBox chkRestricoes 
         BackColor       =   &H80000018&
         Caption         =   "Restrições (cheques sem fundos, protestos ou ações judiciais)"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Tag             =   "4"
         Top             =   1770
         Width           =   2655
      End
      Begin VB.CheckBox chkOutras 
         BackColor       =   &H80000018&
         Caption         =   "Outras"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Tag             =   "5"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CheckBox chkDebitoJuros 
         BackColor       =   &H80000018&
         Caption         =   "Débito de Juros"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Tag             =   "1"
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.TextBox txtDataEnvio 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7200
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtRepresentante 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
      Width           =   3975
   End
   Begin VB.TextBox txtCliente 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.Frame fraPedido 
      BackColor       =   &H80000018&
      Caption         =   "Pedido"
      Height          =   1335
      Left            =   360
      TabIndex        =   16
      Top             =   240
      Width           =   2175
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         MaxLength       =   6
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Line Line1 
      X1              =   8760
      X2              =   8760
      Y1              =   240
      Y2              =   5280
   End
   Begin VB.Label lblSolucao 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Solução"
      Height          =   195
      Left            =   4200
      TabIndex        =   21
      Top             =   3480
      Width           =   585
   End
   Begin VB.Label lblRazaoDetalhada 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Razão"
      Height          =   195
      Left            =   4200
      TabIndex        =   20
      Top             =   1920
      Width           =   465
   End
   Begin VB.Label lblDataEnvio 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Data de Envio"
      Height          =   195
      Left            =   7200
      TabIndex        =   18
      Top             =   960
      Width           =   1020
   End
   Begin VB.Label lblRepresentante 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Representante"
      Height          =   195
      Left            =   2880
      TabIndex        =   17
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label lblCliente 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Cliente"
      Height          =   195
      Left            =   2880
      TabIndex        =   15
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmNotificacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDados As New ADODB.Recordset
Private mSQL As String

Private Sub cmdLimpar_Click()

    txtPedido.Text = ""
    txtCliente.Text = ""
    txtRepresentante.Text = ""
    txtDataEnvio.Text = ""
    txtRazaoDetalhada.Text = ""
    txtSolucao.Text = ""
    chkDebitoJuros.Value = Unchecked
    chkForaPoliticaComercial.Value = Unchecked
    chkDuplicatasAtraso.Value = Unchecked
    chkRestricoes.Value = Unchecked
    chkOutras.Value = Unchecked
    chkExigirAlteracaoPedido.Value = Unchecked
    
    txtPedido.SetFocus

End Sub

Private Sub cmdNotificar_Click()
    
    Dim lCodigoRepresentante As String
    Dim lTipo As String
    Dim lStatus As String
    Dim lPedido As String
    Dim lExigeAlteracao As Integer
    Dim lTextoCabecalho As String
    Dim lTextoRazao As String
    Dim lTextoSolucao As String
    Dim lNomeArquivo As String
    Dim lArquivoID As Long
    Dim lNotificacaoAntiga As String
    
    '****************************************************************************
    'Faz as consistências.
    '****************************************************************************
    
    If txtRazaoDetalhada.Text = "" Then
    
        MsgBox "A informação 'Razão Detalhada' é obrigatória e deve ser preenchida.", vbCritical, Me.Caption
        
        txtRazaoDetalhada.SetFocus
        
        Exit Sub
        
    End If
    
    If txtSolucao.Text = "" Then
    
        MsgBox "A informação 'Solução' é obrigatória e deve ser preenchida.", vbCritical, Me.Caption
        
        txtSolucao.SetFocus
        
        Exit Sub
        
    End If
    
    '****************************************************************************
    'Define os parâmetros a serem gravados no banco.
    '****************************************************************************
    
    If chkExigirAlteracaoPedido.Value = Checked Then
        
        lStatus = "L"
        lExigeAlteracao = 1
        
    ElseIf chkExigirAlteracaoPedido.Value = Unchecked Then
        
        lStatus = "N"
        lExigeAlteracao = 0
        
    End If
    
    lTipo = ""
    
    If chkDebitoJuros.Value = Checked Then
        lTipo = lTipo & 1
    End If
    
    If chkForaPoliticaComercial.Value = Checked Then
        lTipo = lTipo & 2
    End If
    
    If chkDuplicatasAtraso.Value = Checked Then
        lTipo = lTipo & 3
    End If
    
    If chkRestricoes.Value = Checked Then
        lTipo = lTipo & 4
    End If
    
    If chkOutras.Value = Checked Then
        lTipo = lTipo & 5
    End If
    
    If lTipo = "" Then
    
        MsgBox "A informação 'Razão da Notificação' é obrigatória e deve ser preenchida.", vbCritical, Me.Caption
        
        chkDebitoJuros.SetFocus
        
        Exit Sub
        
    End If
    
    '****************************************************************************
    'Recupera notificações anteriores. O objetivo é não perder recados anteriores
    'quando houverem novas atualizações.
    '****************************************************************************
    
    mSQL = "SELECT TexNeg FROM Pedido WHERE NroPed = " & txtPedido.Text
    
    mDados.Open mSQL, Conexao, adOpenForwardOnly, adLockOptimistic
    
    lNotificacaoAntiga = IIf(IsNull(mDados("TexNeg")) = True, "", mDados("TexNeg"))
    
    mDados.Close
    
    '****************************************************************************
    'Salva os dados.
    '****************************************************************************
    
    lArquivoID = FreeFile
    lTextoRazao = "RAZÃO - " & txtRazaoDetalhada.Text
    lTextoSolucao = "SOLUÇÃO - " & txtSolucao.Text
    lCodigoRepresentante = Format(txtRepresentante.Tag, "0000")
    lNomeArquivo = "MOVRALT" & lCodigoRepresentante & ".TXT"
    lPedido = Format(txtPedido.Text, "0000000")
    
    If lExigeAlteracao = 0 Then
        lTextoCabecalho = "Notificação de pendência enviada por " & sgNomUsuSis & " em " & Now
    ElseIf lExigeAlteracao = 1 Then
        lTextoCabecalho = "Notificação de pendência - PASSÍVEL DE ALTERAÇÃO - enviada por " & sgNomUsuSis & " em " & Now
    End If
    
    mSQL = "INSERT INTO Notifica_Pedido (NroPed, DatNot, CodUsu, TexNeg, TexSol, TipNeg, Status) "
    mSQL = mSQL & "VALUES (" & txtPedido.Text & ", GETDATE(), " & LgCodUsuSis & ", '" & txtRazaoDetalhada.Text & "', '" & txtSolucao.Text & "', " & lTipo & ", '" & lStatus & "')"
    
    Conexao.Execute mSQL
    
    If chkExigirAlteracaoPedido.Value = Checked Then
        
        If lNotificacaoAntiga = "" Then
            mSQL = "UPDATE Pedido SET TexNeg = '" & lTextoCabecalho & vbCrLf & lTextoRazao & vbCrLf & lTextoSolucao & vbCrLf & vbCrLf & "', FlgAlt = 'L' WHERE NroPed = " & txtPedido.Text
        Else
            mSQL = "UPDATE Pedido SET TexNeg = '" & lNotificacaoAntiga & vbCrLf & vbCrLf & lTextoCabecalho & vbCrLf & lTextoRazao & vbCrLf & lTextoSolucao & vbCrLf & vbCrLf & "', FlgAlt = 'L' WHERE NroPed = " & txtPedido.Text
        End If
        
    ElseIf chkExigirAlteracaoPedido.Value = Unchecked Then
        
        If lNotificacaoAntiga = "" Then
            mSQL = "UPDATE Pedido SET TexNeg = '" & lTextoCabecalho & vbCrLf & lTextoRazao & vbCrLf & lTextoSolucao & vbCrLf & vbCrLf & "', FlgAlt = 'N' WHERE NroPed = " & txtPedido.Text
        Else
            mSQL = "UPDATE Pedido SET TexNeg = '" & lNotificacaoAntiga & vbCrLf & vbCrLf & lTextoCabecalho & vbCrLf & lTextoRazao & vbCrLf & lTextoSolucao & vbCrLf & vbCrLf & "', FlgAlt = 'N' WHERE NroPed = " & txtPedido.Text
        End If
        
    End If
    
    Conexao.Execute mSQL
    
    '****************************************************************************
    'Cria arquivo-texto.
    '****************************************************************************
    
    Open "C:\Interface\" & lNomeArquivo For Output As #lArquivoID
    Print #lArquivoID, lPedido & lExigeAlteracao & lTextoCabecalho
    Print #lArquivoID, lPedido & lExigeAlteracao & lTextoRazao
    Print #lArquivoID, lPedido & lExigeAlteracao & lTextoSolucao
    Close #lArquivoID
    
    '****************************************************************************
    'Transfere arquivo para o FTP e apaga-o no endereço local.
    '****************************************************************************
    
    With intTransfer
    
        .URL = "ftp://201.48.31.34"
        .Protocol = icFTP
        .RequestTimeout = 100
        .RemotePort = 21
        .AccessType = icDirect
        .UserName = "unocann"
        .Password = "unodataac5621"
      '   .Password = "u2n4o5c3a1n4n4"
        If .StillExecuting = True Then
            .Cancel
        End If
        
        .Execute , "SEND C:\Interface\" & lNomeArquivo & " " & lNomeArquivo
        
        Do While .StillExecuting = True
            DoEvents
        Loop
    
    End With
    
    '****************************************************************************
    'Limpa o formulário e exibe mensagem confirmando o sucesso da operação.
    '****************************************************************************
    
    txtPedido.Text = ""
    txtCliente.Text = ""
    txtRepresentante.Text = ""
    txtDataEnvio.Text = ""
    txtRazaoDetalhada.Text = ""
    txtSolucao.Text = ""
    chkDebitoJuros.Value = Unchecked
    chkForaPoliticaComercial.Value = Unchecked
    chkDuplicatasAtraso.Value = Unchecked
    chkRestricoes.Value = Unchecked
    chkOutras.Value = Unchecked
    chkExigirAlteracaoPedido.Value = Unchecked
    
    MsgBox "Notificação enviada com sucesso!", vbInformation, Me.Caption
    
    txtPedido.SetFocus
    
End Sub

Private Sub CmdSair_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Me.Left = (MDIProjUNO.Width - Me.Width) / 2
    Me.Top = (MDIProjUNO.Height - Me.Height) / 4

End Sub

Private Sub txtPedido_GotFocus()

    With txtPedido
        .SelStart = 0
        .SelLength = Len(txtPedido.Text)
        .SetFocus
    End With

End Sub

Private Sub txtPedido_LostFocus()
    
    If txtPedido.Text <> "" Then
    
        '************************************************************************
        'Limpa o formulário.
        '************************************************************************
    
        txtCliente.Text = ""
        txtRepresentante.Text = ""
        txtDataEnvio.Text = ""
        txtRazaoDetalhada.Text = ""
        txtSolucao.Text = ""
        chkDebitoJuros.Value = Unchecked
        chkForaPoliticaComercial.Value = Unchecked
        chkDuplicatasAtraso.Value = Unchecked
        chkRestricoes.Value = Unchecked
        chkOutras.Value = Unchecked
        chkExigirAlteracaoPedido.Value = Unchecked
    
        '************************************************************************
        'Consulta o pedido e exibe seus dados.
        '************************************************************************
    
        mSQL = "SELECT P.*, C.NomCli, R.CodRep, R.NomRep "
        mSQL = mSQL & "FROM Pedido P "
        mSQL = mSQL & "INNER JOIN Cliente C ON C.CodCli = P.CodCli "
        mSQL = mSQL & "INNER JOIN Representante R ON R.CodRep = P.CodRep "
        mSQL = mSQL & "WHERE P.NroPed = " & txtPedido.Text
    
        mDados.Open mSQL, Conexao, adOpenForwardOnly, adLockOptimistic
        
        If mDados.EOF = True Then
    
            MsgBox "Não existe pedido com o número informado.", vbCritical, Me.Caption
        
            mDados.Close
            
            txtPedido.SetFocus
        
            Exit Sub
        
        End If
    
        txtCliente.Text = mDados("NomCli")
        txtRepresentante.Text = mDados("NomRep")
        txtRepresentante.Tag = mDados("CodRep")
        txtDataEnvio.Text = Format(mDados("DatEnv"), "dd/mm/yyyy")
        
        '************************************************************************
        'Emite alertas caso pedido consultado já tenha sido faturado ou
        'cancelado.
        '************************************************************************
    
        If IsNull(mDados("NroNot")) = False Then
    
            MsgBox "O pedido " & txtPedido.Text & " já foi faturado.", vbCritical, Me.Caption
        
            mDados.Close
        
            Exit Sub
    
        End If
    
        If IsNull(mDados("DatCan")) = False Then
    
            MsgBox "O pedido " & txtPedido.Text & " foi cancelado.", vbCritical, Me.Caption
        
            mDados.Close
        
            Exit Sub
    
        End If
    
        txtRazaoDetalhada.SetFocus
    
        mDados.Close
        
    End If
    
End Sub

Private Sub txtRazaoDetalhada_GotFocus()

    With txtRazaoDetalhada
        .SelStart = 0
        .SelLength = Len(txtRazaoDetalhada.Text)
        .SetFocus
    End With

End Sub

Private Sub txtRazaoDetalhada_LostFocus()

    txtRazaoDetalhada.Text = UCase(txtRazaoDetalhada.Text)

End Sub

Private Sub txtSolucao_GotFocus()

    With txtSolucao
        .SelStart = 0
        .SelLength = Len(txtSolucao.Text)
        .SetFocus
    End With

End Sub

Private Sub txtSolucao_LostFocus()

    txtSolucao.Text = UCase(txtSolucao.Text)

End Sub
