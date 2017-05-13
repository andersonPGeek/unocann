VERSION 5.00
Begin VB.Form Frm_Clientes 
   Caption         =   "CAD003 - Manutenção no Cadastro de Propecção"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Status do Cliente"
      Height          =   855
      Left            =   6120
      TabIndex        =   7
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox Cbo_CodigoStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Frm_Clientes.frx":0000
         Left            =   120
         List            =   "Frm_Clientes.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox Cbo_NomeStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Frm_Clientes.frx":0004
         Left            =   120
         List            =   "Frm_Clientes.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pesquisa"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6000
      Begin VB.TextBox txt_Pesquisa 
         Height          =   375
         Left            =   120
         MaxLength       =   35
         TabIndex        =   1
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.PictureBox Grd_Cliente 
      BackColor       =   &H00FFFFFF&
      Height          =   3120
      Left            =   120
      ScaleHeight     =   3060
      ScaleWidth      =   10260
      TabIndex        =   2
      Top             =   1080
      Width           =   10320
   End
   Begin VB.PictureBox Grd_ClienteMkt 
      BackColor       =   &H00FFFFFF&
      Height          =   2640
      Left            =   120
      ScaleHeight     =   2580
      ScaleWidth      =   10260
      TabIndex        =   10
      Top             =   4440
      Width           =   10320
   End
   Begin VB.Label Label1 
      Caption         =   "Telemarketing já realizado"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   3255
   End
End
Attribute VB_Name = "Frm_Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'+----------------------------------------------------+
'| Rotina de manutencao tabela de Tba_CLIENTE       |
'| Wagner Estefanio Cardoso                           |
'| Fev/97                                       V.01  |
'+----------------------------------------------------+
Dim Sl_Desc As String
    
Dim sl_Num_LinhasSel            As String
Dim sl_Cod_Delecao              As String
Dim Sl_Desc_Mensagem1           As String
Dim Sl_Desc_Mensagem2           As String
Dim sl_Texto_Atributo           As String
Dim sl_Cod_Retorno              As String

Dim il_Valor_Linha              As Integer
Dim il_Valor_loop               As Integer
Dim Il_Num_Gridrow              As Integer
Dim il_Num_LinhaMaior           As Integer
Dim il_Num_LinhaMenor           As Integer
Dim il_Num_LinhaAtual           As Integer
Dim nl_num_gridrow              As Integer
Dim Rst_Cliente                 As rdoResultset
Dim Rst_Contato                 As rdoResultset


Dim Rst_Ficha As rdoResultset


Private Sub PosicionarTransp()
       
    If Trim(UCase$(Txt_pesquisa.Text)) > Trim(UCase$(Grd_Cliente.Text)) Then
       il_Num_LinhaMaior = Grd_Cliente.Rows
       il_Num_LinhaMenor = Grd_Cliente.Row
    Else
       il_Num_LinhaMaior = Grd_Cliente.Row
       il_Num_LinhaMenor = 0
    End If
    
Procurar:

    If (il_Num_LinhaMaior - il_Num_LinhaMenor) < 11 Then
       nl_num_gridrow = 0
       GoTo Posicionar
    End If

    il_Num_LinhaAtual = il_Num_LinhaMenor + Int((il_Num_LinhaMaior - il_Num_LinhaMenor) / 2)

    If il_Num_LinhaAtual = 0 Then
       il_Num_LinhaMenor = 0
       nl_num_gridrow = 0
       GoTo Posicionar
    End If

    Grd_Cliente.Row = il_Num_LinhaAtual
    Grd_Cliente.Col = 1

    If Trim(UCase$(Txt_pesquisa.Text)) > Trim(UCase$(Grd_Cliente.Text)) Then
       il_Num_LinhaMenor = il_Num_LinhaAtual
       GoTo Procurar
    Else
       If UCase$(Txt_pesquisa) < Grd_Cliente Then
          il_Num_LinhaMaior = il_Num_LinhaAtual
          GoTo Procurar
       Else
          il_Num_LinhaMenor = il_Num_LinhaAtual - 1
          nl_num_gridrow = 0
          GoTo Posicionar
       End If
    End If

Posicionar:
    
    If Grd_Cliente.Rows = nl_num_gridrow Then
       Grd_Cliente.Col = 1
       If Grd_Cliente.Rows = 2 And Grd_Cliente.Text = "" Then
          Exit Sub
       End If
    End If

    Grd_Cliente.Row = nl_num_gridrow
    Grd_Cliente.Col = 1
    
    If UCase$(Txt_pesquisa) > UCase$(Grd_Cliente.Text) Then
       nl_num_gridrow = nl_num_gridrow + 1
       If nl_num_gridrow = Grd_Cliente.Rows Then
          Exit Sub
       Else
          GoTo Posicionar
       End If
    End If

End Sub

Private Sub Bto_Alterar_Click()
    If Sg_FlagGrid = "1" Then
        If Grd_Cliente.SelStartRow <> Grd_Cliente.SelEndRow Then
            MsgBox "Selecione somente um cliente para alteração.", vbOK + vbExclamation, "A T E N Ç Ã O"
            Exit Sub
        End If
        flag_Cliente = "A"
        Grd_Cliente.Col = 0
        
        If Grd_Cliente.Text = "" Then
           MsgBox "Pelo menos um cliente deve ser selecionado para alteração.", vbOK + vbExclamation, "A T E N Ç Ã O"
           Exit Sub
        End If
    
        Sg_CodCli = Grd_Cliente.Text
        
        Screen.MousePointer = vbHourglass              ' Ativa a ampulheta
       
        Frm_IncCliente.Show
        Screen.MousePointer = vbDefault                ' Desativa a ampulheta
    End If
    If Sg_FlagGrid = "2" Then
        If Grd_ClienteMkt.SelStartRow <> Grd_ClienteMkt.SelEndRow Then
            MsgBox "Selecione somente um cliente para alteração.", vbOK + vbExclamation, "A T E N Ç Ã O"
            Exit Sub
        End If
        flag_Cliente = "A"
        Grd_ClienteMkt.Col = 0
        
        If Grd_ClienteMkt.Text = "" Then
           MsgBox "Pelo menos um cliente deve ser selecionado para alteração.", vbOK + vbExclamation, "A T E N Ç Ã O"
           Exit Sub
        End If
    
        Sg_CodCli = Grd_ClienteMkt.Text
        
        Screen.MousePointer = vbHourglass              ' Ativa a ampulheta
       
        Frm_IncCliente.Show
        Screen.MousePointer = vbDefault                ' Desativa a ampulheta
    End If
 
End Sub



Private Sub Bto_Excluir_Click()
    Dim erro As Integer
    
 
    If Trim(Grd_Cliente.Text) = "" Then
        MsgBox "Pelo menos um documento deve ser selecionada para exclusão.", vbOK + vbExclamation, "A T E N Ç Ã O"
        Exit Sub
    End If
    flag_Cliente = "E"
    sl_Num_LinhasSel = ""
    sl_Cod_Delecao = ""
    
    If Grd_Cliente.SelStartRow = 0 Then
       Grd_Cliente.SelStartRow = 1
    End If
    
    il_Valor_Linha = Grd_Cliente.SelStartRow
    
    If Grd_Cliente.SelStartRow <> Grd_Cliente.SelEndRow Then
        sl_Num_LinhasSel = "*"
    End If

     
    For il_Valor_loop = Grd_Cliente.SelStartRow To Grd_Cliente.SelEndRow
        
        Grd_Cliente.Row = il_Valor_Linha

        Grd_Cliente.Col = 1
        Sl_Desc_Mensagem1 = "do CLIENTE '" & Trim(Grd_Cliente.Text) & "' ?"
        Sl_Desc_Mensagem2 = "de todos os clientes selecionados ?"
        sl_Texto_Atributo = Grd_Cliente.Text
        Call ConfirmarExclusao(Sl_Desc_Mensagem1, Sl_Desc_Mensagem2, sl_Texto_Atributo, sl_Cod_Delecao, sl_Num_LinhasSel, il_Valor_Linha, sl_Cod_Retorno)
         
        Grd_Cliente.Col = 0
        
        If sl_Cod_Retorno = "S" Then
          Sl_Desc = "SELECT * FROM Tba_CLIENTES where Cli_Codigo  = '" & Trim(Grd_Cliente.Text) & "'"
          Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenKeyset, rdConcurRowVer)
          While Rst_Cliente.EOF = False
             Rst_Cliente.Edit
             Rst_Cliente.Delete
             Rst_Cliente.MoveNext
          Wend
          Rst_Cliente.Close

 
           If Grd_Cliente.Rows = 1 Then
              Grd_Cliente.Row = 0
              Grd_Cliente.Col = 0
              Grd_Cliente.Text = ""
              Grd_Cliente.Col = 1
              Grd_Cliente.Text = ""
           Else
              If Grd_Cliente.SelStartRow = Grd_Cliente.SelEndRow Then
                 Grd_Cliente.Col = 0
                 Grd_Cliente.Text = ""
                 Grd_Cliente.Col = 1
                 Grd_Cliente.Text = ""
              Else
                 Grd_Cliente.RemoveItem Grd_Cliente.Row
              End If
           End If
        End If
'        CN.CommitTrans
    Next

   'Atualização do Grid
    Sl_Desc = "SELECT * "
    Sl_Desc = Sl_Desc & " FROM Tba_ClientesStatus b,"
    Sl_Desc = Sl_Desc & "      Tba_CLIENTES a"
    Sl_Desc = Sl_Desc & " where b.Cli_CodigoStatus  = a.Cli_CodigoStatus"
    Sl_Desc = Sl_Desc & " ORDER BY CLI_NOME "
    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenKeyset, rdConcurRowVer)
    Il_Num_Gridrow = 1
    Do While Rst_Cliente.EOF = False
       Grd_Cliente.Rows = Il_Num_Gridrow + 1
       Grd_Cliente.Row = Il_Num_Gridrow
       Grd_Cliente.Col = 0
       Grd_Cliente.Text = Rst_Cliente("Cli_Codigo")
       Grd_Cliente.Col = 1
       Grd_Cliente.Text = Rst_Cliente("Cli_Nome")
       Grd_Cliente.Col = 2
       Grd_Cliente.Text = Rst_Cliente("Cli_NomeStatus")
       
       Rst_Cliente.MoveNext
       Il_Num_Gridrow = Il_Num_Gridrow + 1
    Loop
    Grd_Cliente.Row = Grd_Cliente.SelStartRow
    Rst_Cliente.Close
  Exit Sub
    
Trata_Erro:
  Rotina_erro ("A")
  Beep
End Sub


Private Sub bto_Incluir_Click()
    flag_Cliente = "I"
    Screen.MousePointer = vbHourglass          ' Ativa a ampulheta
    Set Cliente_objt = Frm_Clientes
    Frm_IncCliente.Show
    Screen.MousePointer = vbDefault            ' Desativa a ampulheta
End Sub

Private Sub Bto_Sair_Click()
 
   Unload Me
End Sub

Private Sub Cbo_NomeStatus_Click()
        Cbo_CodigoStatus.ListIndex = Cbo_NomeStatus.ListIndex
        Busca_Cliente
End Sub

Private Sub Form_Activate()
 
 
    Cbo_NomeStatus.Clear
    Cbo_CodigoStatus.Clear
    Sl_Desc = "select * from Tba_ClientesStatus"
    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    Cbo_NomeStatus.AddItem "TODOS"
    Cbo_CodigoStatus.AddItem "999"
    While Rst_Cliente.EOF = False
          Cbo_CodigoStatus.AddItem Rst_Cliente!Cli_CodigoStatus
          Cbo_NomeStatus.AddItem Rst_Cliente!Cli_NomeStatus
          Rst_Cliente.MoveNext
    Wend
    Rst_Cliente.Close
    Cbo_NomeStatus.ListIndex = 3
    Cbo_CodigoStatus.ListIndex = 3
    
   
 
    Busca_ClienteMkt
 
 

   Exit Sub
    
Trata_Erro:
  Rotina_erro ("")
  Beep
End Sub
Private Sub Busca_Cliente()
    Grd_Cliente.Rows = 2
    Grd_Cliente.Col = 0
    Grd_Cliente.Row = 0
    Grd_Cliente.ColWidth(0) = 1300
    Grd_Cliente.Text = "CGC"
    Grd_Cliente.Col = 1
    Grd_Cliente.ColWidth(1) = 4700
    Grd_Cliente.Text = " Nome do Cliente"
    Grd_Cliente.Col = 2
    Grd_Cliente.ColWidth(2) = 3950
    Grd_Cliente.Text = " Status"
    Grd_Cliente.Row = 1
    Grd_Cliente.Col = 0
    Grd_Cliente.Text = " "
    Grd_Cliente.Col = 1
    Grd_Cliente.Text = " "
    Grd_Cliente.Col = 2
    Grd_Cliente.Text = " "
    
    Sl_Desc = "SELECT * "
    Sl_Desc = Sl_Desc & " FROM Tba_ClientesStatus b,"
    Sl_Desc = Sl_Desc & "      Tba_CLIENTES a"
    Sl_Desc = Sl_Desc & " where b.Cli_CodigoStatus  = a.Cli_CodigoStatus"
    If Cbo_CodigoStatus.Text <> "999" Then
       Sl_Desc = Sl_Desc & "   and a.Cli_CodigoStatus = " & Cbo_CodigoStatus
    End If
    Sl_Desc = Sl_Desc & "   and isnull(a.Cli_status)"
     
    Sl_Desc = Sl_Desc & " ORDER BY CLI_NOME "
    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    
    Il_Num_Gridrow = 1
    Do While Rst_Cliente.EOF = False
       Grd_Cliente.Rows = Il_Num_Gridrow + 1
       Grd_Cliente.Row = Il_Num_Gridrow
       Grd_Cliente.Col = 0
       Grd_Cliente.Text = Rst_Cliente("Cli_Codigo")
       Grd_Cliente.Col = 1
       Grd_Cliente.Text = Rst_Cliente("Cli_Nome")
       Grd_Cliente.Col = 2
       Grd_Cliente.Text = Rst_Cliente("Cli_NomeStatus")

       Rst_Cliente.MoveNext
       Il_Num_Gridrow = Il_Num_Gridrow + 1
    Loop

    Grd_Cliente.Row = Grd_Cliente.SelStartRow
    Rst_Cliente.Close

End Sub
Private Sub Form_Load()
    Left = 0
    Top = 0
    Height = 7880
    Width = 11850
  
' Montagem dos Cabeçalhos do Grid
    Sg_FlagGrid = "1"
    
End Sub



 
 

Private Sub Grd_CLIENTE_DblClick()
   
    Sg_FlagGrid = "1"
   ' Ativa Alteracao na ROW corrente no evento CLICK
    If Grd_Cliente.SelStartRow <> Grd_Cliente.SelEndRow Then
        MsgBox "Selecione somente um cliente para alteração.", vbOK + vbExclamation, "A T E N Ç Ã O"
        Exit Sub
    End If
    flag_Cliente = "A"
    Grd_Cliente.Col = 0
    
    If Grd_Cliente.Text = "" Then
       MsgBox "Pelo menos um documento deve ser selecionado para alteração.", vbOK + vbExclamation, "A T E N Ç Ã O"
       Exit Sub
    End If
      
      Grd_Cliente.Col = 0
      Sg_CodCli = Grd_Cliente.Text
      
      Grd_Cliente.Col = 1
      Sg_NomCli = Grd_Cliente.Text
  
    Screen.MousePointer = vbHourglass              ' Ativa a ampulheta
     
    Frm_IncCliente.Show
    Screen.MousePointer = vbDefault                ' Desativa a ampulheta

End Sub

Private Sub Grd_ClienteMkt_Click()
   Sg_FlagGrid = "2"
   ' Ativa Alteracao na ROW corrente no evento CLICK
     flag_Cliente = "A"
    Grd_ClienteMkt.Col = 0
    
       
      Grd_ClienteMkt.Col = 0
      Sg_CodCli = Grd_ClienteMkt.Text
      
      Grd_ClienteMkt.Col = 1
      Sg_NomCli = Grd_ClienteMkt.Text
  
          
     
 
End Sub

Private Sub Grd_ClienteMkt_DblClick()
   Sg_FlagGrid = "2"
   ' Ativa Alteracao na ROW corrente no evento CLICK
    If Grd_ClienteMkt.SelStartRow <> Grd_ClienteMkt.SelEndRow Then
        MsgBox "Selecione somente um cliente para alteração.", vbOK + vbExclamation, "A T E N Ç Ã O"
        Exit Sub
    End If
    flag_Cliente = "A"
    Grd_ClienteMkt.Col = 0
    
    If Grd_ClienteMkt.Text = "" Then
       MsgBox "Pelo menos um documento deve ser selecionado para alteração.", vbOK + vbExclamation, "A T E N Ç Ã O"
       Exit Sub
    End If
      
      Grd_ClienteMkt.Col = 0
      Sg_CodCli = Grd_ClienteMkt.Text
      
      Grd_ClienteMkt.Col = 1
      Sg_NomCli = Grd_ClienteMkt.Text
  
    Screen.MousePointer = vbHourglass              ' Ativa a ampulheta
     
    Frm_IncCliente.Show
    Screen.MousePointer = vbDefault                ' Desativa a ampulheta
End Sub

Private Sub txt_Pesquisa_Change()
   If Trim(Txt_pesquisa.Text) = "" Then
       Grd_Cliente.SetFocus
       SendKeys "^{home}", True
       Txt_pesquisa.SetFocus
       Exit Sub
    End If
    
    Grd_Cliente.Col = 1

    PosicionarTransp

    If Trim(Txt_pesquisa.Text) <> "" Then
       Grd_Cliente.SetFocus
       SendKeys "{home}", True
       Txt_pesquisa.SetFocus
    End If

End Sub



Private Sub Busca_ClienteMkt()
    Grd_ClienteMkt.Rows = 2
    Grd_ClienteMkt.Col = 0
    Grd_ClienteMkt.Row = 0
    Grd_ClienteMkt.ColWidth(0) = 1300
    Grd_ClienteMkt.Text = "CGC"
    Grd_ClienteMkt.Col = 1
    Grd_ClienteMkt.ColWidth(1) = 4700
    Grd_ClienteMkt.Text = " Nome do Cliente"
    Grd_ClienteMkt.Col = 2
    Grd_ClienteMkt.ColWidth(2) = 2000
    Grd_ClienteMkt.Text = " Data do Contato"
    
    Grd_ClienteMkt.Col = 3
    Grd_ClienteMkt.ColWidth(3) = 1950
    Grd_ClienteMkt.Text = " Nome do Contato"
    
    Grd_ClienteMkt.Row = 1
    Grd_ClienteMkt.Col = 0
    Grd_ClienteMkt.Text = " "
    Grd_ClienteMkt.Col = 1
    Grd_ClienteMkt.Text = " "
    Grd_ClienteMkt.Col = 2
    Grd_ClienteMkt.Text = " "
    Grd_ClienteMkt.Col = 3
    Grd_ClienteMkt.Text = " "
    
    Sl_Desc = "SELECT DISTINCT  [Tba_Clientes].[Cli_Tipo],"
    Sl_Desc = Sl_Desc & "          [Tba_FichaCliente].[Cli_Codigo] AS Tba_FichaCliente_Cli_Codigo,"
    Sl_Desc = Sl_Desc & "          [Tba_FichaCliente].[Fix_Data] AS Tba_FichaCliente_Fix_Data,"
    Sl_Desc = Sl_Desc & "          [Tba_FichaCliente].[CodUsuSis] AS Tba_FichaCodUsuSis,"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_Codigo] AS Tba_Clientes_Cli_Codigo,"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_Nome] as nome,"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_Endereco],"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_Bairro],"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_Municipio],"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_UF],"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_Pais],"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_CEP],"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_Telcor],"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_TelcorRamal],"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_Fax],"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_FaxRamal],"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_homepage],"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_CodigoStatus],"
    Sl_Desc = Sl_Desc & "          [Tba_Clientes].[Cli_Status] as status"
    Sl_Desc = Sl_Desc & " FROM Tba_Clientes INNER JOIN Tba_FichaCliente ON [Tba_Clientes].[Cli_Codigo] =[Tba_FichaCliente].[Cli_Codigo]"
    Sl_Desc = Sl_Desc & " WHERE   not isnull([Tba_Clientes].[Cli_Status]) "
     
    Sl_Desc = Sl_Desc & " ORDER BY  [Tba_FichaCliente].[Fix_Data] desc;"
    
    
    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    
    Il_Num_Gridrow = 1
    Do While Rst_Cliente.EOF = False
          Grd_ClienteMkt.Rows = Il_Num_Gridrow + 1
          Grd_ClienteMkt.Row = Il_Num_Gridrow
          Grd_ClienteMkt.Col = 0
          Grd_ClienteMkt.Text = Rst_Cliente("Tba_FichaCliente_Cli_Codigo")
          Grd_ClienteMkt.Col = 1
          Grd_ClienteMkt.Text = Rst_Cliente("nome")
          Grd_ClienteMkt.Col = 2
          Grd_ClienteMkt.Text = Rst_Cliente("Tba_FichaCliente_Fix_Data")
          
          Sl_Desc = "select * from Tba_Usuarios where codususis =  " & Rst_Cliente("Tba_FichaCodUsuSis")
          Set Rst_Contato = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
          If Rst_Contato.EOF = False Then
             Grd_ClienteMkt.Col = 3
             Grd_ClienteMkt.Text = Rst_Contato("nomususis")
          End If
          Rst_Contato.Close

          
          
          
          Il_Num_Gridrow = Il_Num_Gridrow + 1
     
   
       
       Rst_Cliente.MoveNext

    Loop

    Grd_ClienteMkt.Row = Grd_ClienteMkt.SelStartRow
    Rst_Cliente.Close

End Sub







