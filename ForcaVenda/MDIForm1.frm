VERSION 5.00
Begin VB.MDIForm MDIProjUNO 
   BackColor       =   &H8000000C&
   Caption         =   "V1.05 -  U N O C A N N   T U B O S   E   C O N E X � E S  -  For�a de Venda    "
   ClientHeight    =   7860
   ClientLeft      =   165
   ClientTop       =   -690
   ClientWidth     =   13035
   LinkTopic       =   "MDIForm1"
   MousePointer    =   99  'Custom
   Picture         =   "MDIForm1.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   3360
      Top             =   360
   End
   Begin VB.Menu mnu_pedido 
      Caption         =   "&Simula��o de Pedido"
   End
   Begin VB.Menu mnuLiberaPed 
      Caption         =   "&Libera Pedidos"
   End
   Begin VB.Menu mnuInterface 
      Caption         =   "&Interface"
   End
   Begin VB.Menu mnuComunica 
      Caption         =   "&Comunica��o"
      Begin VB.Menu mnuChave 
         Caption         =   "&Gera Chave"
      End
      Begin VB.Menu mnuLiberaAlt 
         Caption         =   "&Notifica��o de Pend�ncia "
      End
   End
   Begin VB.Menu mnu_indica 
      Caption         =   "In&dicadores"
   End
   Begin VB.Menu mnuConsultarPedidos 
      Caption         =   "&Consultar Pedidos"
   End
   Begin VB.Menu mnuProdutosImportados 
      Caption         =   "Produtos Importados"
   End
   Begin VB.Menu mnu_PosPed 
      Caption         =   "P&osi��o de Pedidos"
   End
   Begin VB.Menu mnuImprime 
      Caption         =   "&Monitora Pedidos"
   End
   Begin VB.Menu mnuCancelaPedido 
      Caption         =   "Cancela P&edido"
   End
   Begin VB.Menu mnuRel 
      Caption         =   "&Relat�rios Operacionais"
      Begin VB.Menu mnurelfat 
         Caption         =   "Faturamento"
      End
      Begin VB.Menu mnuRelTitVen 
         Caption         =   "T�tulos Vencidos"
      End
      Begin VB.Menu itmSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu itmPedidosSemDigitacao 
         Caption         =   "Pedidos Sem Digita��o"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuGerenciais 
      Caption         =   "Relat�rios &Gerenciais"
      Begin VB.Menu mnuRelCliCid 
         Caption         =   "&Clientes por Cidade"
      End
      Begin VB.Menu mnuVendasRepre 
         Caption         =   "&Vendas por Representante"
      End
      Begin VB.Menu itmRelatorioQualidadeVendas 
         Caption         =   "&Qualidade das Vendas"
      End
      Begin VB.Menu itmRelatorioPrecosVendasCompostos 
         Caption         =   "&Pre�os de Venda dos Compostos"
      End
      Begin VB.Menu itmRelatorioProdutosMargensNegativas 
         Caption         =   "P&rodutos com Margens Negativas"
      End
      Begin VB.Menu itmRelatorioVendasRepresentantes 
         Caption         =   "Vendas por Representantes (Sint�tico)"
      End
   End
   Begin VB.Menu mnuTlmk 
      Caption         =   "Telemarketing"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDIProjUNO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ControleTempo As Integer

Private Sub itmPedidosSemDigitacao_Click()

    frmRelatorioPedidosSemDigitacao.Show

End Sub

Private Sub itmRelatorioPrecosVendasCompostos_Click()

    frmRelatorioPrecoVendaComposto.Show

End Sub

Private Sub itmRelatorioProdutosMargensNegativas_Click()

    frmRelatorioMargensNegativas.Show

End Sub

Private Sub itmRelatorioQualidadeVendas_Click()

    frmRelatorioQualidadeVendasProduto.Show

End Sub

Private Sub itmRelatorioVendasRepresentantes_Click()

    frmRelatorioVendasRepresentante.Show

End Sub

Private Sub MDIForm_Activate()

    '*****************************************************************************
    'Se o MDI foi aberto pela primeira vez e se o usu�rio logado for um
'    'representante, o programa abre a janela com indicadores.
'    '*****************************************************************************
'
    If bgabertura = False And APLICA = 1 Then
    
    
    sgQuery = "select a.nroped, a.datped, a.codcli, c.NomCli, sum(b.vlrite) - sum(distinct a.vlrsimples) as Valor, d.DscCnd,"
    sgQuery = sgQuery & " a.NomTra , a.datenv, a.flgalt "
    sgQuery = sgQuery & " from Pedido a, Item_pedido b, Cliente c, Condicao d"
  '  sgQuery = sgQuery & " Where a.datlib is not null " 'and a.FlgAlt is not null and a.FlgAlt <> 'N'"
    sgQuery = sgQuery & " Where a.FlgAlt <> 'N' "
    sgQuery = sgQuery & " and a.sitped = 'N' and a.nronot is null"
    sgQuery = sgQuery & " and a.nroped = b.nroped"
    sgQuery = sgQuery & " and a.codcli = c.codcli"
    sgQuery = sgQuery & " and a.codcnd = d.codcnd"
    sgQuery = sgQuery & " and a.codrep = " & sgRepresentante
    sgQuery = sgQuery & " and a.mgrtot < 9 "
    sgQuery = sgQuery & " and b.idxdsc > a.dscpdr"
    sgQuery = sgQuery & " and a.datped >= convert(datetime, '01/01/2013', 103)"
    sgQuery = sgQuery & " group by a.nroped, a.datped, a.codcli, c.NomCli, d.DscCnd, a.NomTra, a.datenv, a.flgalt"
    sgQuery = sgQuery & " order by 2, 1 desc"
        
    Consulta sgQuery

    
    If Rs.RecordCount = 0 Then
        Rs.Close
        Set Rs = Nothing
    Else
         FrmMisc.Show (modal)
    End If

    '*****************************************************************************
    'A vari�vel a seguir define que o formul�rio j� foi aberto uma vez, e que a
    'janela com os indicadores n�o deve ser exibida novamente.
    '*****************************************************************************

End If

    bgabertura = True

End Sub

Private Sub MDIForm_Load()
       
    '*****************************************************************************
    'Zera o c�digo do representante se o usu�rio logado for um Administrador.
    '*****************************************************************************
    
    If APLICA = 0 Then
        sgRepresentante = 0
    End If
    
    '*****************************************************************************
    'Informa que n�o haver� nenhum pedido envolvido com telemarketing.
    '*****************************************************************************
    
    bgPedMKT = False
    
    '*****************************************************************************
    'Define os menus que ser�o exibidos no MDI, de acordo com o perfil do usu�rio
    'logado. Administradores e representantes n�o podem ter os mesmos acessos.
    '*****************************************************************************
    
    '*****************************************************************************
    'APLICA = 1 quando o usu�rio logado � o representante, Quando o valor for 0, o
    'usu�rio atual � um administrador.
    '*****************************************************************************
    
    If APLICA = 1 Then
        mnuComunica.Visible = False
        mnuChave.Visible = False
        mnuImprime.Visible = False
        mnuCancelaPedido.Visible = False
        mnuGerenciais.Visible = False
        mnuProdutosImportados.Visible = False
    Else
        mnu_pedido.Visible = False
        mnuLiberaPed.Visible = False
        mnuInterface.Visible = False
        mnu_indica.Visible = False
        mnu_PosPed.Visible = False
        mnuConsultarPedidos.Visible = False
        mnuProdutosImportados.Visible = False
    
    End If

End Sub

Private Sub mnu_indica_Click()
    
    '*****************************************************************************
    'Indicadores.
    '*****************************************************************************
    
    FrmMisc.Show (modal)
    
End Sub

Private Sub mnu_pedido_Click()
    
    '*****************************************************************************
    'Pedidos.
    '*****************************************************************************
    
    FrmConhecimento.Show
    
End Sub

Private Sub mnu_PosPed_Click()
    
    '*****************************************************************************
    'Posi��es de Pedidos (Representantes).
    '*****************************************************************************
    
    FrmPosiPed.Show
    
End Sub

Private Sub mnuaviso_Click()
  
      sgQuery = "select a.nroped, a.datped, a.codcli, c.NomCli, sum(b.vlrite) - sum(distinct a.vlrsimples) as Valor, d.DscCnd,"
    sgQuery = sgQuery & " a.NomTra , a.datenv, a.flgalt "
    sgQuery = sgQuery & " from Pedido a, Item_pedido b, Cliente c, Condicao d"
  '  sgQuery = sgQuery & " Where a.datlib is not null " 'and a.FlgAlt is not null and a.FlgAlt <> 'N'"
    sgQuery = sgQuery & " Where a.FlgAlt <> 'N' "
    sgQuery = sgQuery & " and a.sitped = 'N' and a.nronot is null"
    sgQuery = sgQuery & " and a.nroped = b.nroped"
    sgQuery = sgQuery & " and a.codcli = c.codcli"
    sgQuery = sgQuery & " and a.codcnd = d.codcnd"
    sgQuery = sgQuery & " and a.codrep = " & sgRepresentante
    sgQuery = sgQuery & " and a.mgrtot < 9 "
    sgQuery = sgQuery & " and b.idxdsc > a.dscpdr"
    sgQuery = sgQuery & " and a.datped >= convert(datetime, '01/01/2013', 103)"
    sgQuery = sgQuery & " group by a.nroped, a.datped, a.codcli, c.NomCli, d.DscCnd, a.NomTra, a.datenv, a.flgalt"
    sgQuery = sgQuery & " order by 2, 1 desc"
        
    Consulta sgQuery

    'If bgabertura = False And APLICA = 1 Then
    
    If Rs.RecordCount = 0 Then
        Rs.Close
        Set Rs = Nothing
    Else
        FrmMisc.Show (modal)
    End If

    '*****************************************************************************
    'A vari�vel a seguir define que o formul�rio j� foi aberto uma vez, e que a
    'janela com os indicadores n�o deve ser exibida novamente.
    '*****************************************************************************

  '  bgabertura = True

    
End Sub

Private Sub mnuCancelaPedido_Click()
    
    '*****************************************************************************
    'Cancelamento de Pedidos.
    '*****************************************************************************
    
    FrmCancelaPedido.Show
    
End Sub

Private Sub mnuChave_Click()
    
    '*****************************************************************************
    'Gerador de Chaves de Descontos.
    '*****************************************************************************
    
    FrmGeraChave.Show
    
End Sub

Private Sub mnuConsultarPedidos_Click()

    frmConsultarPedidos.Show

End Sub

Private Sub mnuImprime_Click()
    
    '*****************************************************************************
    'Posi��es de Pedidos (Administradores).
    '*****************************************************************************
    
    FrmPosMonit.Show
    
End Sub

Private Sub mnuInterface_Click()

   

    If bgabertura = False And APLICA = 1 Then
    
    
    '*****************************************************************************
    'A vari�vel a seguir define que o formul�rio j� foi aberto uma vez, e que a
    'janela com os indicadores n�o deve ser exibida novamente.
    '*****************************************************************************

    bgabertura = True

    '*****************************************************************************
    'Interface. A janela s� � aberta se n�o houver nenhum MDI Child aberto.
    '*****************************************************************************
    
    If Forms.Count > 1 Then
        
        MsgBox "Feche todas as telas Antes de efetuar a interface", vbInformation
        
        Exit Sub
        
    End If

End If
   'FrmInterface.Show
    
End Sub

Private Sub mnuLiberaAlt_Click()

    '*****************************************************************************
    'Emite notifica��es de pedidos.
    '*****************************************************************************
    
    frmNotificacoes.Show
    
End Sub

Private Sub mnuLiberaPed_Click()
    
    '*****************************************************************************
    'Libera��o de Pedidos.
    '*****************************************************************************
    
    FrmLiberaPedido.Show
    
End Sub

Private Sub mnuProdutosImportados_Click()

    '*****************************************************************************
    'Rela��o de produtos importados da Polyvin.
    '*****************************************************************************
    
    frmRelacaoProdutosImportados.Show

End Sub

Private Sub mnuRelCliCid_Click()
    
    '*****************************************************************************
    'Relat�rio de Clientes por Cidades.
    '*****************************************************************************
    
    FrmRelCliCid.Show
    
End Sub

Private Sub mnurelfat_Click()
    
    '*****************************************************************************
    'Relat�rio de Faturamento.
    '*****************************************************************************
    
    FrmRelFat.Show
    
End Sub

Private Sub mnuRelTitVen_Click()
    
    '*****************************************************************************
    'Relat�rio de T�tulos Vencidos.
    '*****************************************************************************
    
    FrmRelTitVen.Show
    
End Sub

Private Sub mnuTlmk_Click()
    
    '*****************************************************************************
    'Telemarketing.
    '*****************************************************************************
    
    lgSeqLig = 0
    
    FrmTMKPrincipal.Show
    
End Sub

Private Sub mnuVendasRepre_Click()
    
    '*****************************************************************************
    'Relat�rio de Vendas por Representante.
    '*****************************************************************************
    
    FrmRelRepreProduto.Show '1
    
End Sub

'Private Sub Timer1_Timer()
'
'If ControleTempo = 5 Then
'
'    If FrmAviso.Visible = False Then
'
'        FrmAviso.Show vbModal
'
'        ControleTempo = 0
'    End If
'
'Else
'
'        ControleTempo = ControleTempo + 1
'
'End If
'
'End Sub
