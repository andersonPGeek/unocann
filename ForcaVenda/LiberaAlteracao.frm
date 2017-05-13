VERSION 5.00
Object = "{BEACF734-D8AC-11D7-9B57-000B6A03449D}#2.0#0"; "MASKED.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmLiberaAlteracao 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notificação de pendência para aprovação de pedidos"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10650
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10650
   Begin VB.CheckBox ChkAltera 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      Caption         =   "Exige Alteração"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   120
      MaskColor       =   &H00800080&
      TabIndex        =   17
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton BtoLimpar 
      BackColor       =   &H80000016&
      Caption         =   "&Limpar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8040
      Picture         =   "LiberaAlteracao.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Selecione uma opção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   6360
      TabIndex        =   9
      Top             =   0
      Width           =   4215
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Outros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Restrições (Cheque SF/Protesto/Ações Judiciais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   3855
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Duplicatas em atraso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Descontos, prazos ou preços fora da política comercial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Débito de Juros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.TextBox TxtSolu 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1215
      Left            =   1200
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4800
      Width           =   8415
   End
   Begin VB.TextBox TxtObserva 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1215
      Left            =   1200
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3120
      Width           =   8415
   End
   Begin VB.TextBox txtMensa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6120
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.CommandButton BtoSair 
      BackColor       =   &H80000016&
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
      Height          =   975
      Left            =   9360
      Picture         =   "LiberaAlteracao.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1095
   End
   Begin Project_Masked.Masked MskSenha 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      FormatoString   =   "000000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   65535
      ForeColor       =   8388608
      ValInteiro      =   7
   End
   Begin VB.CommandButton CmdGerar 
      BackColor       =   &H80000016&
      Caption         =   "Notificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      MaskColor       =   &H00E0E0E0&
      Picture         =   "LiberaAlteracao.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet itcFTP 
      Left            =   840
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.Label LblDatEnv 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label LblCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   5415
   End
   Begin VB.Label LblRep 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FFFF&
      Caption         =   "Data  Envio"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cliente"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Representante"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Solução"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Razão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "FrmLiberaAlteracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sss As Boolean
Dim slExiste As Boolean
Dim slPasso As String
Dim ilCodRep As Integer
Dim vFileName As String
Dim slArqPed As String
Dim slDatLib As String
Dim slString As String
Dim dlNroPed As Double
Dim slObserva As String
Dim slSolu As String
Dim operacao As String
Dim slTexNeg As String
Dim slTipo As String
Dim slFlgDig As String
Dim slNomUsuAltLib As String

Private Sub BtoLimpar_Click()

    LimpaGeral

End Sub

Private Sub BtoSair_Click()

    Unload Me
    
    Set FrmLiberaAlteracao = Nothing

End Sub

Private Sub ChkAltera_Click()

    If ChkAltera.Value = 1 Then
        ChkAltera.BackColor = &H800080
    Else
        ChkAltera.BackColor = &H40C0&
    End If
    
End Sub

Private Sub cmdGerar_Click()
    
    Dim sldata As String
    Dim slJuntaObs As String
    Dim slOper As Integer

    If Trim(MskSenha.Texto) = "" Or Trim(MskSenha.Texto) = 0 Then
        
        MsgBox "Informe o Número do Pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        MskSenha.SetFocus
        
        Exit Sub
        
    End If

    If Trim(TxtObserva.Text) = "" And Trim(TxtSolu.Text) = "" Then
        
        MsgBox "Informe a razão e/ou solução para a notificação", vbExclamation + vbOKOnly, "Atenção!"
        
        TxtObserva.SetFocus
        
        Exit Sub
        
    End If

    slTipo = ""

    If Check1.Value = 1 Then
        slTipo = slTipo & "1"
    End If

    If Check2.Value = 1 Then
        slTipo = slTipo & "2"
    End If

    If Check3.Value = 1 Then
        slTipo = slTipo & "3"
    End If

    If Check4.Value = 1 Then
        slTipo = slTipo & "4"
    End If
   
    If Check5.Value = 1 Then
        slTipo = slTipo & "5"
    End If

    If ChkAltera.Value = 1 Then
        slOper = 1
    Else
        slOper = 0
    End If

    On Error GoTo TrataErro

    dlNroPed = Trim(MskSenha.Texto)
    slObserva = Trim(TxtObserva.Text)
    slObserva = Replace(slObserva, "'", "´")
    slObserva = Replace(slObserva, """", "§")
    slObserva = Replace(slObserva, "§", "´")
    slSolu = Trim(TxtSolu)
    slSolu = Replace(slSolu, "'", "´")
    slSolu = Replace(slSolu, """", "§")
    slSolu = Replace(slSolu, "§", "´")

    If Trim(slTipo) = "" Then
        
        MsgBox "Selecione ao menos uma opção", vbExclamation + vbOKOnly, "Atenção!"
        
        Exit Sub
        
    End If

    CmdGerar.Enabled = False
    TxtObserva.Enabled = False
    TxtSolu.Enabled = False
    Frame1.Enabled = False
    BtoLimpar.Enabled = False
    BtoSair.Enabled = False

    On Error Resume Next

    slExiste = True
    sss = True
    igFileNumber = FreeFile
    slArqPed = "MOVRALT" & Format(ilCodRep, "0000") & ".TXT"
    vFileName = "c:\INTERFACE\" & Trim(slArqPed)
    
    If Dir(vFileName) <> "" Then
        Kill vFileName
    End If

    txtMensa.Visible = True

    operacao = "get " & slArqPed & " " & vFileName  ' copia MOVRALT9999 do site FTP para c:\interface
    
    executaComando operacao

    If sss = False Then
        
        MsgBox "Erro na transmissão da liberação do pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        MskSenha.SetFocus
        
        Exit Sub
        
    End If

    slExiste = True
    sss = True
    operacao = "delete " & slArqPed  ' Deleta MOVRALT9999 no site FTP
    
    executaComando operacao

    If sss = False Then
        
        MsgBox "Erro na transmissão da liberação do pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        MskSenha.SetFocus
        
        Exit Sub
        
    End If

    sldata = Format(CDate(Date) & " " & Time, "dd/mm/yyyy hh:mm:ss")

    If Trim(slTexNeg) <> "" Then
        slJuntaObs = Trim(slTexNeg) & vbCrLf
    End If

    If slOper = 1 Then
        slJuntaObs = slJuntaObs & " Notificação de pendência -  PASSÍVEL DE ALTERAÇÃO - enviado(a) por " & sgNomUsuSis & " em " & sldata & vbCrLf
    Else
        slJuntaObs = slJuntaObs & " Notificação de pendência enviado(a) por " & sgNomUsuSis & " em " & sldata & vbCrLf
    End If

    If Trim(slObserva) <> "" Then
        slJuntaObs = slJuntaObs & "RAZÃO - " & Trim(slObserva) & vbCrLf
    End If

    If Trim(slSolu) <> "" Then
        slJuntaObs = slJuntaObs & "SOLUÇÃO - " & Trim(slSolu)
    End If

    slJuntaObs = Replace(slJuntaObs, "'", "´")
    slJuntaObs = Replace(slJuntaObs, """", "§")
    slJuntaObs = Replace(slJuntaObs, "§", "´")

    'CRIA ARQUIVO AUXILIAR

    If Dir("c:\INTERFACE\MOVALT.TXT") <> "" Then
        Kill "c:\INTERFACE\MOVALT.TXT"
    End If

    igFileNumber = FreeFile
    
    Open "c:\INTERFACE\MOVALT.TXT" For Output As #igFileNumber
    
    Print #igFileNumber, Trim(slJuntaObs)
    
    Close #igFileNumber

'----------------------------------------------------------------------

    igFileNumber = FreeFile
    
    If Dir(vFileName) <> "" Then
        Open vFileName For Append As #igFileNumber
    Else
        Open vFileName For Output As #igFileNumber
    End If

    Open "c:\INTERFACE\MOVALT.TXT" For Input As #2
    
    Do While Not EOF(2)
        
        Input #2, sglinha
        slString = Format(MskSenha.Texto, "0000000") & Trim(slOper) & Trim(sglinha)
        Print #igFileNumber, slString
        
    Loop

    Close #2
    
    Close #igFileNumber

    slExiste = True
    sss = True
    operacao = "send " & vFileName & " " & slArqPed  ' Grava MOVRALT9999 no site FTP
    
    executaComando operacao

    If sss = False Then
        
        MsgBox "Erro na transmissão da liberação do pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        MskSenha.SetFocus
        
        Exit Sub
        
    End If

    On Error GoTo TrataErro

    If slOper = 1 Then
        
        sgQuery = "Update transito_pedido set DatLibAlt = convert(datetime,'" & sldata & "',103), "
        sgQuery = sgQuery & " CodUsuLibAlt = " & LgCodUsuSis
        sgQuery = sgQuery & " where NroPed = " & dlNroPed
        sgQuery = sgQuery & "   and Datlib = convert(datetime,'" & slDatLib & "',103)"
        
        Set Rs = Conexao.Execute(sgQuery)
        Set Rs = Nothing
    
        sgQuery = "Insert Notifica_Pedido Values ("
        sgQuery = sgQuery & dlNroPed & ", convert(datetime,'" & sldata & "',103), " & LgCodUsuSis & ", "
        sgQuery = sgQuery & " '" & Trim(slObserva) & "', '" & Trim(slSolu) & "', '" & Trim(slTipo) & "', 'L')"
        
        Set Rs = Conexao.Execute(sgQuery)
        Set Rs = Nothing
    
        sgQuery = "Update pedido set FlgAlt = 'L', TexNeg = '" & Trim(slJuntaObs) & "'"
        sgQuery = sgQuery & " where NroPed = " & dlNroPed
        
        Set Rs = Conexao.Execute(sgQuery)
        Set Rs = Nothing
        
    Else
   
        sgQuery = "Update pedido set FlgAlt = 'N', TexNeg = '" & Trim(slJuntaObs) & "'"
        sgQuery = sgQuery & " where NroPed = " & dlNroPed
        
        Set Rs = Conexao.Execute(sgQuery)
        Set Rs = Nothing
   
        sgQuery = "Insert Notifica_Pedido Values ("
        sgQuery = sgQuery & dlNroPed & ", convert(datetime,'" & sldata & "',103), " & LgCodUsuSis & ", "
        sgQuery = sgQuery & " '" & Trim(slObserva) & "', '" & Trim(slSolu) & "', '" & Trim(slTipo) & "', 'N')"
        
        Set Rs = Conexao.Execute(sgQuery)
        Set Rs = Nothing
    
    End If

    txtMensa.Visible = False

    LimpaGeral

    Exit Sub

TrataErro:

    CmdGerar.Enabled = True
    TxtObserva.Enabled = True
    TxtSolu.Enabled = True
    Frame1.Enabled = True
    BtoLimpar.Enabled = True
    BtoSair.Enabled = True

    Rotina_Erro "LiberaAltPedido"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Call EventoEnter(KeyAscii)
    
End Sub

Private Sub Form_Load()

    Me.Left = 0
    Me.Top = 0
    Me.Height = 8070
    Me.Width = 10740

    MskSenha.TipodeDados numero
    
    'itcFTP.URL = "ftp://ftp.unocann.com.br"
    'itcFTP.Protocol = icFTP
    'itcFTP.RequestTimeout = 100
    'itcFTP.RemotePort = 21
    'itcFTP.AccessType = icDirect
    'itcFTP.UserName = "unocann"
    'itcFTP.Password = "unodataac5621"
   
    itcFTP.URL = "ftp://172.21.0.3"
    itcFTP.Protocol = icFTP
    itcFTP.RequestTimeout = 100
    itcFTP.RemotePort = 21
    itcFTP.AccessType = icDirect
    itcFTP.UserName = "unocann"
    itcFTP.Password = "unodataac5621"
    
End Sub

Private Sub MskSenha_GotFocus()
    
    Call SelecionaTudo
    
End Sub

Private Sub MskSenha_LostFocus()
    
    Dim lExecutor As New ADODB.Command
    Dim lParametro As New ADODB.Parameter
    Dim ExisteN As Boolean

    If Me.ActiveControl.Name = "BtoSair" Or Me.ActiveControl.Name = "BtoLimpar" Then
        
        Exit Sub
        
    End If

    If Trim(MskSenha.Texto) = "" Or Trim(MskSenha.Texto) = 0 Then
        
        MsgBox "Informe o Número do Pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        MskSenha.SetFocus
        
        Exit Sub
        
    End If

    On Error GoTo TrataErro

    dlNroPed = Trim(MskSenha.Texto)
    
    With lExecutor
    
        .ActiveConnection = Conexao
        .CommandType = adCmdStoredProc
        .CommandText = "sp_notificacao"
        
        Do While .Parameters.Count > 0
            .Parameters.Delete 0
        Loop
        
        Set lParametro = .CreateParameter("@Pedido", adInteger, adParamInput, 4, dlNroPed)
        .Parameters.Append lParametro
        
        Set Rs = .Execute
    
    End With
    
    Set lParametro = Nothing
    Set lExecutor = Nothing
    
    'sgQuery = "select a.DatLib, a.DatRecUno, a.DatLibAlt, b.codrep, b.ClasCor, b.datlibuno,"
    'sgQuery = sgQuery + "        b.SitPed, b.texneg, b.flgdig, f.texneg as texnegtra, f.texsol, f.tipneg,"
    'sgQuery = sgQuery + "        b.DatEnv , c.NomRep, d.NomCli, e.nomusu"
    'sgQuery = sgQuery + " from transito_pedido a, pedido b, representante c, cliente d, usuario e, notifica_pedido f"
    'sgQuery = sgQuery + " Where a.NroPed = " & MskSenha.Texto
    'sgQuery = sgQuery + "   and a.nroped = b.nroped"
    'sgQuery = sgQuery + "   and b.codrep = c.codrep"
    'sgQuery = sgQuery + "   and b.codcli = d.codcli"
    'sgQuery = sgQuery + "   and a.codusulibalt *= e.codusu"
    'sgQuery = sgQuery + "   and a.datlib = (select max(datlib) from transito_pedido"
    'sgQuery = sgQuery + "                     where nroped = " & MskSenha.Texto & ")"
    'sgQuery = sgQuery + "   and a.nroped *= f.nroped"
    'sgQuery = sgQuery + "   and a.datlibalt *= f.DatNot"
    
    'Consulta sgQuery
    
    If Rs.EOF Then
        
        MsgBox "Pedido inexistente", vbExclamation + vbOKOnly, "Atenção!"
        
        MskSenha.SetFocus
        
        Exit Sub
        
    End If

    LblRep.Caption = Rs!NomRep
    LblCli.Caption = Rs!NomCli
    LblDatEnv.Caption = Format(Rs!DatEnv, "dd/mm/yyyy hh:mm:ss")

    slFlgDig = IIf(IsNull(Rs!FlgDig), "", Rs!FlgDig)
    ilCodRep = Rs!codrep
    slDatLib = Rs!Datlib
    slTipo = "2"
    ExisteN = False

    If IsNull(Rs!DatLibAlt) Then
        
        sgQuery = "select b.texneg, f.datnot, f.texneg as texnegtra, f.texsol, f.tipneg, b.DatEnv, c.nomusu"
        sgQuery = sgQuery + " from pedido b, notifica_pedido f, usuario c"
        sgQuery = sgQuery + " Where b.NroPed = " & MskSenha.Texto
        sgQuery = sgQuery + "   and b.nroped = f.nroped"
        sgQuery = sgQuery + "   and f.codusu = c.codusu"
        sgQuery = sgQuery + "   and f.datNot = (select max(datNot) from Notifica_pedido"
        sgQuery = sgQuery + "                     where nroped = " & MskSenha.Texto & ")"
        sgQuery = sgQuery + "   and f.status <> 'C'"
        
        Consulta2 sgQuery
        
        If Not Rs2.EOF Then
            
            slTexNeg = IIf(IsNull(Rs2!texnegtra), "", Rs2!texneg)
            
            If IsNull(Rs2!tipneg) Then
                slTipo = "2"
            Else
                slTipo = Rs2!tipneg
            End If
            
            TxtObserva.Text = IIf(IsNull(Rs2!texnegtra), "", Rs2!texnegtra)
            TxtSolu.Text = IIf(IsNull(Rs2!texsol), "", Rs2!texsol)
            
            slNomUsuAltLib = IIf(IsNull(Rs2!nomusu), "", Rs2!nomusu)
            ExisteN = True
            
        End If
        
    Else
    
        slTexNeg = IIf(IsNull(Rs!texneg), "", Rs!texneg)
        
        If IsNull(Rs!tipneg) Then
            slTipo = "2"
        Else
            slTipo = Rs!tipneg
        End If
        
        TxtObserva.Text = IIf(IsNull(Rs!texnegtra), "", Rs!texnegtra)
        TxtSolu.Text = IIf(IsNull(Rs!texsol), "", Rs!texsol)
        
        slNomUsuAltLib = IIf(IsNull(Rs!nomusu), "", Rs!nomusu)
        
    End If

    If InStr(slTipo, "1") > 0 Then
        Check1.Value = 1
    End If

    If InStr(slTipo, "2") > 0 Then
        Check2.Value = 1
    End If

    If InStr(slTipo, "3") > 0 Then
        Check3.Value = 1
    End If

    If InStr(slTipo, "4") > 0 Then
        Check4.Value = 1
    End If
   
    If InStr(slTipo, "5") > 0 Then
        Check5.Value = 1
    End If

    If Trim(Rs!SitPed) = "C" Or Trim(Rs!SitPed) = "U" Then
        
        MsgBox "Este pedido está cancelado no sistema", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        LimpaGeral
        
        Exit Sub
        
    End If

    If Not IsNull(Rs!DatLibAlt) Then
        
        TxtObserva.Enabled = False
        TxtSolu.Enabled = False
        
        MsgBox "Este pedido tem notificação passível de alteração aberta no dia " & Format(Rs!DatLibAlt, "dd/mm/yyyy hh:mm:ss") & vbCrLf & "Por " & slNomUsuAltLib, vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        CmdGerar.Enabled = False
        MskSenha.Enabled = False
        
        Exit Sub
    
    Else
        
        If ExisteN = True Then
            
            sgQuery = MsgBox("Este pedido tem notificação aberta no dia " & Format(Rs2!Datnot, "dd/mm/yyyy hh:mm:ss") & vbCrLf & "Por " & slNomUsuAltLib & ", Deseja emitir nova Notificação ?", vbQuestion + vbYesNo + vbDefaultButton1, "Atenção!")
            
            If sgQuery = vbYes Then
                
                TxtObserva.Text = ""
                TxtSolu.Text = ""
                Check1.Value = 0
                Check2.Value = 0
                Check3.Value = 0
                Check4.Value = 0
                Check5.Value = 0
                
            Else
            
                Rs.Close
                
                Set Rs = Nothing
                
                LimpaGeral
                
                Exit Sub
                
            End If
        
        End If
    
    End If

    If Trim(Rs!ClasCor) = "R" And IsNull(Rs!datlibuno) And sgFlgUsu <> "L" Then
        
        MsgBox "Acesso negado a este pedido. Passível de liberação pelo Depto Comercial", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        LimpaGeral
        
        Exit Sub
    
    End If

    Rs.Close
    
    Set Rs = Nothing

    If Trim(slFlgDig) = "S" Then
        ChkAltera.Value = 0
        ChkAltera.Enabled = False
    End If

    MskSenha.Enabled = False

    Exit Sub

TrataErro:

    Rotina_Erro "MskSenha_LostFocus"

End Sub

Private Sub TxtObserva_Change()

    'Call SelecionaTudo
    
End Sub

Private Sub LimpaGeral()

    LblRep.Caption = ""
    LblCli.Caption = ""
    LblDatEnv.Caption = ""
    MskSenha.Texto = ""
    TxtObserva.Text = ""
    TxtSolu.Text = ""

    ilCodRep = 0
    dlNroPed = 0
    vFileName = ""
    slArqPed = ""
    slDatLib = ""
    slString = ""
    slObserva = ""
    slSolu = ""
    operacao = ""
    slTexNeg = ""
    slNomUsuAltLib = ""
    
    ChkAltera.Value = 0
    ChkAltera.Enabled = True
    ChkAltera.BackColor = &H40C0&
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    Check5.Value = 0
    CmdGerar.Enabled = True
    TxtObserva.Enabled = True
    TxtSolu.Enabled = True
    CmdGerar.Enabled = True
    Frame1.Enabled = True
    BtoLimpar.Enabled = True
    BtoSair.Enabled = True
    MskSenha.Enabled = True
    MskSenha.SetFocus

End Sub

Private Sub executaComando(ByVal op As String)
 
    On Error Resume Next

    If itcFTP.StillExecuting Then
        itcFTP.Cancel
    End If
    
    txtMensa.Text = txtMensa.Text & "Comando: " & op & vbCrLf
    itcFTP.Execute , op
    
    terminaComando
    
End Sub

Private Sub terminaComando()
    
    Do While itcFTP.StillExecuting
        DoEvents
    Loop
    
End Sub

Private Sub itcFTP_StateChanged(ByVal State As Integer)
  
    On Error Resume Next
    
    Select Case State
    
        Case icResolvingHost:
            
            slPasso = "Resolvendo Host"
            
        Case icHostResolved:
            
            slPasso = "Host Resolvido"
            
        Case icConnecting:
            
            slPasso = "Conectando ..."
            
        Case icConnected:
            
            slPasso = "Conectado"
            
        Case icRequesting:
            
            slPasso = "Requisitando ..."
            
        Case icRequestSent:
            
            slPasso = "Requisição enviada"
            
        Case icReceivingResponse
            
            slPasso = "Recebendo ..."
            
        Case icResponseReceived:
            
            slPasso = "Resposta recebida"
            
        Case icDisconnecting:
            
            slPasso = "Desconectando ..."
            
        Case icDisconnected:
            
            slPasso = "Desconectado"
            
        Case icError:
            
            Call GravarLog("Erro: " & itcFTP.ResponseCode & " " & itcFTP.ResponseInfo, 0)
            
    End Select
    
    txtMensa = txtMensa & slPasso & vbCrLf
    txtMensa.SelStart = Len(txtMensa.Text)
    
    Err.Clear
    
End Sub

Private Sub GravarLog(slErro As String, Tipo As Integer)
    
    Dim slArq As String
    Dim slString As String
    Dim ilFile As Long
    Dim ilAcha As Integer
    
    'If Tipo = 1 Then
        'GoTo Imprime_erro
    'End If

    ilAcha = InStr(1, slErro, "cannot find", 1)

    If ilAcha > 0 Then
        
        slExiste = False
        
        Exit Sub
        
    End If

    ilAcha = InStr(1, slErro, "no such file", 1)
    
    If ilAcha > 0 Then
        
        slExiste = False
        
        Exit Sub
        
    End If

    ilAcha = InStr(1, slErro, "file doesn't exist", 1)
    
    If ilAcha > 0 Then
        
        slExiste = False
        
        Exit Sub
        
    End If
    
Imprime_erro:

    sss = False
    
    txtMensa = txtMensa & slErro & vbCrLf
    
    'slArq = "c:\LOG_INTERFACE\LOGFTP & Format(Date, "ddmmyyyy") & ".txt"
    'ilFile = FreeFile
    
    'If Dir(slArq) <> "" Then
        'Open slArq For Append As #ilFile
    'Else
        'Open slArq For Output As #ilFile
    'End If
    
    'If Tipo = 1 Then
        'slString = aponta & " *****" & slPasso & ">>>>>  " & slErro & " ----> " & Format(CDate(Date) & " " & Time, "dd/mm/yyyy hh:mm:ss")
    'Else
        'slString = aponta & " *****" & slPasso & ">>>>>  " & operacao & " --- " & slErro & "  <<<<<" & slOperMens & "----> " & Format(CDate(Date) & " " & Time, "dd/mm/yyyy hh:mm:ss")
    'End If
    
    'Print #ilFile, slString
    
    'Close #ilFile

End Sub
