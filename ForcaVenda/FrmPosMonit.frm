VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{368CC970-FF03-11D7-9B5A-000B6A03449D}#1.1#0"; "Combo_DB.ocx"
Object = "{BEACF734-D8AC-11D7-9B57-000B6A03449D}#2.0#0"; "Masked.ocx"
Object = "{F454059D-91FE-11D2-8865-AD1268A0A52F}#2.0#0"; "ActiveDate.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmPosMonit 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Monitoramento de Pedidos"
   ClientHeight    =   10530
   ClientLeft      =   -210
   ClientTop       =   1515
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame fraAtualizacoes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Atualizações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10320
      TabIndex        =   19
      Top             =   240
      Width           =   4695
      Begin VB.Label lblDadosProximaAtualizacao 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   23
         Top             =   720
         Width           =   45
      End
      Begin VB.Label lblDadosUltimaAtualizacao 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   960
         TabIndex        =   22
         Top             =   360
         Width           =   45
      End
      Begin VB.Label lblProximaAtualizacao 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Próxima: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblUltimaAtualizacao 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Última:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame fraOrdenacao 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ordenação dos Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6120
      TabIndex        =   14
      Top             =   240
      Width           =   3855
      Begin VB.OptionButton optPedido 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pedido"
         Height          =   255
         Left            =   2520
         TabIndex        =   25
         Top             =   400
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optCliente 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cliente"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton optDigitacao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Digitados"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   1100
         Width           =   975
      End
      Begin VB.OptionButton optDataRecebimento 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data de Recebimento"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1100
         Width           =   1935
      End
      Begin VB.OptionButton optRepresentante 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Representante"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   400
         Width           =   1575
      End
   End
   Begin VB.Frame fraPesquisa 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pesquisa de Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   5415
      Begin VB.CheckBox chkNaoDigitados 
         BackColor       =   &H80000009&
         Caption         =   "Somente não-digitados"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3210
         TabIndex        =   26
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CheckBox Chkimpr 
         BackColor       =   &H80000009&
         Caption         =   "Somente não-impressos"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin Project_Masked.Masked MskNroPedido 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         FormatoString   =   "000000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         AutoTab         =   -1  'True
         ValInteiro      =   6
      End
      Begin rdActiveDate.ActiveDate ActDtfim 
         Height          =   315
         Left            =   3720
         TabIndex        =   7
         Top             =   600
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveDate.ActiveDate ActDtini 
         Height          =   315
         Left            =   1860
         TabIndex        =   8
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin Project_Combo_DB.Combo_DB CboRemet 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   661
         Cols            =   0
         Cabecalho       =   -1  'True
      End
      Begin VB.Label Lbl100 
         BackColor       =   &H000000FF&
         Caption         =   "+100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2505
         TabIndex        =   24
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Representante"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data Final"
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data Inicial"
         Height          =   255
         Index           =   1
         Left            =   1860
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pedido"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdImprimirTodos 
      BackColor       =   &H000080FF&
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   12000
      Picture         =   "FrmPosMonit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton CmdAtualizar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Atualizar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   10320
      Picture         =   "FrmPosMonit.frx":5C12
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   240
      Top             =   10560
   End
   Begin VB.CommandButton BtoSair 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   13680
      Picture         =   "FrmPosMonit.frx":68DC
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid GrdPedido 
      Height          =   7335
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   12938
      _Version        =   393216
      Rows            =   1
      Cols            =   15
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   -2147483640
      GridColorFixed  =   -2147483640
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport rptcontprop 
      Left            =   840
      Top             =   10560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmPosMonit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim slremet As String
Dim ilCodRep As Integer
Dim blLoad As Boolean

'*****************************************************************************
'Variável usada para impedir que a rotina CompoeGridPed exiba uma mensagem
'acusando pesquisa sem retorno logo na abertura do formulário.
'*****************************************************************************
Dim mPrimeiraExecucao As Boolean

'*****************************************************************************
'Variável usada para permitir que as atualizações do grid de pedidos aconteçam
'a cada cinco minutos. Sozinho, o controle Timer suporta intervalor de pouco
'mais de um minuto.
'*****************************************************************************
Dim mControleTempo As Integer

Function CompoeGridPed() As Boolean

    Dim blI As Integer
    
    On Error GoTo TratarErro
    
    If CboRemet.Criterio <> "" Then
        ilCodRep = CboRemet.Codigo
    End If

    If MskNroPedido.Texto > 0 Then
        
        ilCodRep = 0
        
        CboRemet.Criterio = ""
        
    End If

    sgQuery = "select a.nroped, a.datped, a.codcli,c.NomCli, sum(b.vlrite) as Valor, a.codcnd,a.codrep,d.DscCnd, a.clascor, a.flgAlt, "
    sgQuery = sgQuery & "  a.FlgImpr, a.codrep,e.NomRep , a.datrecuno, a.DatLibUno, f.NomUsu, a.cifob,a.sitped, a.datlib, a.datenv, a.seqenvinc, a.flgdig "
    sgQuery = sgQuery & "  from Pedido a, Item_pedido b, Cliente c, Condicao d, representante e, usuario f"
    sgQuery = sgQuery & "  Where a.nroped = b.nroped"
    sgQuery = sgQuery & "    and a.nroped >= 840000"
    sgQuery = sgQuery & "    and a.codcli = c.codcli"
    sgQuery = sgQuery & "    and a.codcnd = d.codcnd"
    sgQuery = sgQuery & "    and a.codrep = e.codrep"
    sgQuery = sgQuery & "    and a.codUsuLib *= f.codusu"
    
    If MskNroPedido.Texto > 0 Then
        sgQuery = sgQuery & "    and a.nroped = " & Trim(MskNroPedido.Texto)
    End If

    If Trim(ActDtini.Text) <> "" Then
        sgQuery = sgQuery & " and a.datped between convert(datetime,'" & Trim(ActDtini.Text) & "',103)"
        sgQuery = sgQuery & " and convert(datetime,'" & Trim(ActDtfim.Text) & "',103)"
    End If
    
    If ilCodRep <> 0 Then
        sgQuery = sgQuery & " and a.codrep = " & ilCodRep
    End If

    If Chkimpr.Value = 1 And MskNroPedido.Texto = 0 Then
        sgQuery = sgQuery & " and a.FlgImpr is null "
    End If
    
    If chkNaoDigitados.Value = Checked Then
        sgQuery = sgQuery & " and a.FlgDig is null "
    End If

    sgQuery = sgQuery & "    group by a.nroped, a.datped, a.codcli, c.NomCli,a.codcnd,a.codrep, d.DscCnd,"
    sgQuery = sgQuery & "              a.cifob, a.sitped, a.datrecuno, a.datlib, a.datenv, a.clascor, "
    sgQuery = sgQuery & "              a.FlgImpr, a.codrep, e.NomRep , a.DatLibUno, f.NomUsu, a.flgAlt, a.seqenvinc, a.flgdig "
    
    If optRepresentante.Value = True Then
        sgQuery = sgQuery & "    order by e.nomrep, a.nroped desc"
    ElseIf optCliente.Value = True Then
        sgQuery = sgQuery & "    order by c.nomcli, a.nroped desc"
    ElseIf optDataRecebimento.Value = True Then
        sgQuery = sgQuery & "    order by a.datrecuno, a.nroped desc"
    ElseIf optDigitacao.Value = True Then
        sgQuery = sgQuery & "    order by a.flgdig desc, a.nroped desc"
    ElseIf optPedido.Value = True Then
        sgQuery = sgQuery & "    order by a.nroped desc"
    End If
    
    Consulta sgQuery
    
    If blLoad = False And Rs.RecordCount > 100 Then
        Lbl100.Visible = True
    Else
        Lbl100.Visible = False
    End If
    
    GrdPedido.Visible = False
    GrdPedido.Rows = 1
    
    blI = 0
    
    If Rs.EOF = False Then
    
        Do While Not Rs.EOF
            
            If blI > 100 Then
                Exit Do
            End If
   
            blI = blI + 1
            
            '*****************************************************************************
            'O grid é preenchido com os dados do pedido. Inicialmente, a cor do status é
            'amarela, podendo mudar de acordo com os dados levantados nas linhas a seguir.
            '*****************************************************************************
   
            GrdPedido.Rows = GrdPedido.Rows + 1
            GrdPedido.TextMatrix(blI, 0) = Format(Trim(Rs!NroPed), "000000")
            GrdPedido.TextMatrix(blI, 1) = Format(Trim(Rs!Datped), "dd/mm/yyyy")
            GrdPedido.TextMatrix(blI, 2) = Format(Trim(Rs!Codcli), "00000")
            GrdPedido.TextMatrix(blI, 3) = Rs!NomCli
            GrdPedido.TextMatrix(blI, 4) = Format(Rs!Valor, "##,###,##0.00")
            GrdPedido.TextMatrix(blI, 5) = Rs!DscCnd
            GrdPedido.TextMatrix(blI, 6) = Rs!NomRep
            GrdPedido.TextMatrix(blI, 7) = IIf(Rs!CIFOB = "C", "CIF", "FOB")
            GrdPedido.TextMatrix(blI, 8) = IIf(IsNull(Rs!datlib), "", Format(Trim(Rs!datlib), "dd/mm/yyyy"))
            GrdPedido.TextMatrix(blI, 9) = IIf(IsNull(Rs!Datrecuno), "", Format(Trim(Rs!Datrecuno), "dd/mm/yyyy"))
            GrdPedido.TextMatrix(blI, 10) = IIf(IsNull(Rs!nomusu), "", Rs!nomusu)
            GrdPedido.TextMatrix(blI, 11) = IIf(IsNull(Rs!datlibuno), "", Format(Trim(Rs!datlibuno), "dd/mm/yyyy"))
            GrdPedido.TextMatrix(blI, 12) = Rs!codrep
            GrdPedido.Row = blI
            GrdPedido.Col = 0
            GrdPedido.CellBackColor = &HFFFF&
            
            '*****************************************************************
            'Marca as primeiras colunas do grid, que informam os status dos
            'pedidos.
            '*****************************************************************

            '*****************************************************************
            'Marca de verde os pedidos que já foram impressos.
            '*****************************************************************

            If Rs!FlgImpr = "S" Then
                GrdPedido.Col = 0
                GrdPedido.CellBackColor = &HFF00&
            End If

            '*****************************************************************
            'Marca de vermelho os pedidos com pendência de liberação.
            '*****************************************************************

            If Rs!ClasCor = "R" And IsNull(Rs!nomusu) Then
                GrdPedido.Col = 0
                GrdPedido.CellBackColor = &HC0&
                GrdPedido.CellForeColor = &HFFFF&
            End If
            
            '*****************************************************************
            'Marca de roxo os pedidos notificados com pendência de alteração.
            '*****************************************************************
            
            If Trim(Rs!flgalt) = "L" Then
                
                GrdPedido.CellBackColor = &H800080
                GrdPedido.CellForeColor = &HFFFF&
                
            End If

            '*****************************************************************
            'Marca de preto os pedidos cancelados.
            '*****************************************************************
            
            If Rs!NroPed = 865515 Then
                MsgBox 5, vbInformation
            End If
            
            If Trim(Rs!sitPed) = "C" Or Trim(Rs!sitPed) = "U" Then
                
                GrdPedido.Col = 0
                GrdPedido.CellBackColor = &H80000012
                GrdPedido.CellForeColor = &HFFFF&
                
            End If
      
            '*****************************************************************
            'Marca as últimas colunas do grid, com alertas do pedido.
            '*****************************************************************
      
            '*****************************************************************
            'Marca de roxo pedido que recebeu notificação de alteração que já
            'foi modificado.
            '*****************************************************************
      
            If Trim(Rs!flgalt) = "A" Then
            
                GrdPedido.TextMatrix(blI, 13) = "A"
                GrdPedido.Col = 13
                GrdPedido.CellBackColor = &H800080
                GrdPedido.CellForeColor = &HFFFF&
            
            Else
            
                '*************************************************************
                'Marca de marrom pedido que recebeu notificação de advertência
                'já lida.
                '*************************************************************
            
                If Trim(Rs!flgalt) = "N" Or Trim(Rs!flgalt) = "O" Then
                    GrdPedido.TextMatrix(blI, 13) = "N"
                    GrdPedido.Col = 13
                    GrdPedido.CellBackColor = &H4080&
                    GrdPedido.CellForeColor = &HFFFFFF
                End If
            
            End If
            
            '*****************************************************************
            'Marca de azul piscina os pedidos que já foram digitados.
            '*****************************************************************
   
            If Trim(Rs!FlgDig) = "S" Then
            
                GrdPedido.TextMatrix(blI, 14) = "D"
                GrdPedido.Col = 14
                GrdPedido.CellBackColor = &HFFFF00
            
                If Rs!SeqEnvInc > 0 Then
                    GrdPedido.CellBackColor = &H80000012
                    GrdPedido.CellForeColor = &HFFFFFF
                End If
            
            End If
      
            Rs.MoveNext
    
        Loop
        
        GrdPedido.Visible = True

        If MskNroPedido.Texto > 0 And blI = 0 Then
            MsgBox "Pedido inexistente", vbExclamation + vbOKOnly, "Atenção!"
        End If
        
    Else
    
        If mPrimeiraExecucao = False Then
            MsgBox "Não há registros compatíveis com os filtros informados.", vbExclamation, "Força de Venda"
        End If
    
    End If
    
    lblDadosUltimaAtualizacao.Caption = "Em " & Format(Now, "dd/mm/yyyy") & ", às " & Format(Now, "hh:mm")
    lblDadosProximaAtualizacao.Caption = "Em " & Format(Now, "dd/mm/yyyy") & ", às " & Format(DateAdd("n", 5, Now), "hh:mm")
    
    Rs.Close
    
    blLoad = False

    Set Rs = Nothing
    
    DoEvents
    
    mPrimeiraExecucao = False
    
    Exit Function

TratarErro:

    Rotina_Erro "CompoeGridPed"

End Function

Private Sub BtoSair_Click()

    Unload Me
  
End Sub

Private Sub CboRemet_Consultar()

    slremet = ""
    
    CboRemet.query = "Select NomRep As Representante, CodRep As Código From Representante Where " & IIf(IsNumeric(CboRemet.Criterio), "Codrep", "Nomrep") & " Like '" & CboRemet.Criterio & "%' order by " & IIf(IsNumeric(CboRemet.Criterio), "Codrep", "Nomrep")
    
End Sub

Private Sub CboRemet_GotFocus()

    Call SelecionaTudo

End Sub

Private Sub CboRemet_LostFocus()

    If CboRemet.Criterio <> "" Then
        'ilCodRep = CboRemet.Codigo
        slremet = CboRemet.Criterio
    Else
        ilCodRep = 0
        slremet = ""
    End If
    
End Sub

Private Sub Check1_Click()

End Sub

Private Sub CmdAtualizar_Click()

    Screen.MousePointer = vbHourglass
    
    DoEvents
    
    If MskNroPedido.Texto = "" Then
        MskNroPedido.Texto = 0
    End If
    
    If MskNroPedido.Texto > 0 Then
        
        ActDtini.Text = ""
        ActDtfim.Text = ""
        ilCodRep = 0
        
        GoTo Compoe
        
    End If
    
    If ActDtini.Text = "" And ActDtini.Text = "" And ilCodRep = 0 Then
        GoTo Compoe
    End If

    If ActDtini.Text <> "" Or ActDtini.Text <> "" Then
    
        If CDate(ActDtini.Text) > CDate(ActDtfim.Text) Or Year(CDate(ActDtini.Text)) < 1950 Or Year(CDate(ActDtfim.Text)) < 1950 Then
            
            MsgBox "Intervalo de datas inválido", vbInformation
            
            ActDtini.SetFocus
            
            Exit Sub
            
        End If
        
    End If
    
Compoe:

    CompoeGridPed
    
    Screen.MousePointer = vbDefault
   
End Sub

Private Sub cmdImprimirTodos_Click()

    '*************************************************************************
    'Rotina desenvolvida por André Corrêa em 23/01/2009.
    '*************************************************************************

    Dim lDados1 As New ADODB.Recordset
    Dim lDados2 As New ADODB.Recordset
    Dim lSQL As String
    Dim lPathRelatorio As String

    If MsgBox("Você está prestes a imprimir TODOS os pedidos que ainda não foram impressos. Deseja continuar?", vbQuestion + vbYesNo, "Força de Venda") = vbYes Then
    
        '*********************************************************************
        'Seleciona todos os pedidos não-impressos, exceto aqueles que ainda
        'dependem de liberação.
        '*********************************************************************
    
        lSQL = "SELECT D.Pedido FROM (SELECT NroPed As Pedido FROM Pedido WHERE NroPed > 840000 And FlgImpr Is Null And ClasCor <> 'R' UNION SELECT NroPed As Pedido FROM Pedido WHERE NroPed > 840000 And FlgImpr Is Null And ClasCor = 'R' And DatLibUno Is Not Null) AS D ORDER BY D.Pedido"
        
        lDados1.Open lSQL, Conexao
        
        If lDados1.EOF = True Then
        
            MsgBox "Não há nenhum pedido com impressão em aberto.", vbExclamation, "Força de Venda"
            
            lDados1.Close
            
            Set lDados1 = Nothing
            Set lDados2 = Nothing
                        
            Exit Sub
        
        End If
        
        Do While lDados1.EOF = False
                                    
            sgQuery = " SELECT"
            sgQuery = sgQuery & " PEDIDO.NroPed, PEDIDO.Datped, PEDIDO.CIFOB, PEDIDO.NomTra, PEDIDO.TexObs, PEDIDO.ChvDsc,"
            sgQuery = sgQuery & " ITEM_PEDIDO.VlrIte, CONDICAO.DscCnd,CLIENTE.CodCli, CLIENTE.NomCli, CLIENTE.EndCli, CLIENTE.BaiCli,CLIENTE.CidCli, CLIENTE.CepCli,"
            sgQuery = sgQuery & " CLIENTE.CgcCli , CLIENTE.InsCli, CLIENTE.UFCli, CLIENTE.FlgContr, CLIENTE.FlgSIMBa, REPRESENTANTE.NomRep"
            sgQuery = sgQuery & " From"
            sgQuery = sgQuery & " PEDIDO , ITEM_PEDIDO , CONDICAO , CLIENTE, REPRESENTANTE, USUARIO  "
            sgQuery = sgQuery & " where PEDIDO.nroped = ITEM_PEDIDO.nroped"
            sgQuery = sgQuery & "   and PEDIDO.codcnd = CONDICAO.codcnd"
            sgQuery = sgQuery & "   and PEDIDO.codcli = CLIENTE.codcli"
            sgQuery = sgQuery & "   and PEDIDO.CodRep = REPRESENTANTE.CodRep"
            sgQuery = sgQuery & "   and PEDIDO.CodUsuLib *= USUARIO.CodUsu"
            sgQuery = sgQuery & "   and PEDIDO.nroped = " & lDados1("Pedido")
            
            lPathRelatorio = App.Path & "\Relatorios\PedidoMatriz.rpt"
            
            With rptcontprop
            
                .ReportFileName = lPathRelatorio
                .SQLQuery = sgQuery
                If APLICA = 1 Then
                    .Connect = "DSN=" & "unocann" & ";UID=" & "sa" & ";PWD=" & "sysadmpss1"
                Else
                    .Connect = "DSN=" & "unocann" & ";UID=" & "sa" & ";PWD=" & "#unoforte5600!"
                End If
                .Destination = crptToPrinter
                .DiscardSavedData = True
                .Action = 1
        
            End With
        
            lSQL = "UPDATE Pedido SET FlgImpr = 'S' WHERE NroPed = " & lDados1("Pedido")
        
            Set lDados2 = Conexao.Execute(lSQL)
            
            lDados1.MoveNext
            
        Loop
        
        lDados1.Close
        
        MsgBox "Impressão realizada com sucesso!", vbInformation, "Força de Venda"
        
    End If
    
    Set lDados1 = Nothing
    Set lDados2 = Nothing
    
End Sub

Private Sub Form_Activate()
    
    Me.WindowState = 2
    
    blLoad = True
    
    CompoeGridPed
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Call EventoEnter(KeyAscii)
    
End Sub

Private Sub Form_Load()

    Me.Left = 0
    Me.Top = 0
    Me.Height = 10750
    Me.Width = 15260
    
    Set CboRemet.Conexao = Conexao

    ilCodRep = 0
    slremet = ""
    blLoad = True
    bgBloqPed = False
    bgConsultaPed = False
    Lbl100.Visible = False
    MskNroPedido.Texto = 0

    GrdPedido.TextMatrix(0, 0) = "Pedido"
    GrdPedido.ColWidth(0) = 700
    GrdPedido.TextMatrix(0, 1) = "Dt.Emissão"
    GrdPedido.ColWidth(1) = 900
    GrdPedido.TextMatrix(0, 2) = "Cod.Cli"
    GrdPedido.ColWidth(2) = 600
    GrdPedido.TextMatrix(0, 3) = "Cliente"
    GrdPedido.ColWidth(3) = 3600
    GrdPedido.TextMatrix(0, 4) = "Val.Pedido"
    GrdPedido.ColWidth(4) = 1000
    GrdPedido.TextMatrix(0, 5) = "Cond.Pagto"
    GrdPedido.ColWidth(5) = 1800
    GrdPedido.TextMatrix(0, 6) = "Representante"
    GrdPedido.ColWidth(6) = 1610
    GrdPedido.TextMatrix(0, 7) = "Frete"
    GrdPedido.ColWidth(7) = 450
    GrdPedido.TextMatrix(0, 8) = "Dt. Envio"
    GrdPedido.ColWidth(8) = 1000
    GrdPedido.TextMatrix(0, 9) = "Dt. Receb."
    GrdPedido.ColWidth(9) = 900
    GrdPedido.TextMatrix(0, 10) = "Liberado Por"
    GrdPedido.ColWidth(10) = 1000
    GrdPedido.TextMatrix(0, 11) = "Dt.Liberação"
    GrdPedido.ColWidth(11) = 1000
    GrdPedido.TextMatrix(0, 12) = ""
    GrdPedido.ColWidth(12) = 0
    GrdPedido.TextMatrix(0, 13) = ""
    GrdPedido.ColWidth(13) = 190
    GrdPedido.TextMatrix(0, 14) = ""
    GrdPedido.ColWidth(14) = 190
    
    mPrimeiraExecucao = True
    mControleTempo = 0
    
End Sub

Private Sub GrdPedido_DblClick()

    bgBloqPed = False
    
    If GrdPedido.RowSel = 0 Then
        Exit Sub
    End If
    
    bgConsultaPed = True
    
    Me.Enabled = False
    
    igNroPed = GrdPedido.TextMatrix(GrdPedido.RowSel, 0)
    sgRepresentante = Trim(GrdPedido.TextMatrix(GrdPedido.RowSel, 12))
    bgBloqPed = True
    igTela = "Monit"
  
    FrmConhecimento.Show
    
End Sub

Private Sub MskNroPedido_GotFocus()
    
    Call SelecionaTudo
    
End Sub

Private Sub Timer1_Timer()

    If mControleTempo = 5 Then
    
        Beep
        CompoeGridPed
        
        mControleTempo = 0
        
    Else
    
        mControleTempo = mControleTempo + 1
        
    End If
    
End Sub
