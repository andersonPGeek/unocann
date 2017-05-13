VERSION 5.00
Object = "{BEACF734-D8AC-11D7-9B57-000B6A03449D}#2.0#0"; "Masked.ocx"
Begin VB.Form FrmCancelaPedido 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelamento de Pedido"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10650
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10650
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
      Picture         =   "CancelaPedido.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox TxtObserva 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00400000&
      Height          =   1575
      Left            =   1080
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3600
      Width           =   8655
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
      Picture         =   "CancelaPedido.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1095
   End
   Begin Project_Masked.Masked MskSenha 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   240
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
      BackColor       =   16777215
      ForeColor       =   8388608
      ValInteiro      =   7
   End
   Begin VB.CommandButton CmdGerar 
      BackColor       =   &H0000FFFF&
      Caption         =   " &Cancelar Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      MaskColor       =   &H000000FF&
      Picture         =   "CancelaPedido.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label LblDatCan 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7680
      TabIndex        =   15
      Top             =   960
      Width           =   975
   End
   Begin VB.Label LblDatEmi 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label LblDatEnv 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6360
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label LblCli 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   5415
   End
   Begin VB.Label LblRep 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dat. Envio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Representante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   5
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dat.Cancel."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7680
      TabIndex        =   16
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "FrmCancelaPedido"
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
Dim ilTipo As Integer
Dim slNomUsuAltLib As String
Dim mCanceladoAposNotificacao As Boolean

Private Sub BtoLimpar_Click()

    LimpaGeral

End Sub

Private Sub BtoSair_Click()

    Unload Me
    
    Set FrmCancelaPedido = Nothing

End Sub

Private Sub cmdGerar_Click()

    Dim sldata As String
    Dim slJuntaObs As String

    If Trim(MskSenha.Texto) = "" Or Trim(MskSenha.Texto) = 0 Then
        
        MsgBox "Informe o Número do Pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        MskSenha.SetFocus
        
        Exit Sub
        
    End If

    If Trim(TxtObserva.Text) = "" Then
        
        MsgBox "Informe a razão para o cancelamento do pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        TxtObserva.SetFocus
        
        Exit Sub
    
    End If

    On Error GoTo TrataErro

    dlNroPed = Trim(MskSenha.Texto)
    
    slObserva = Trim(TxtObserva.Text)
    slObserva = Replace(slObserva, "'", "´")
    slObserva = Replace(slObserva, """", "§")
    slObserva = Replace(slObserva, "§", "´")
    
    cmdGerar.Enabled = False
    TxtObserva.Enabled = False
    BtoLimpar.Enabled = False
    BtoSair.Enabled = False

    sldata = Format(CDate(Date) & " " & Time, "dd/mm/yyyy hh:mm:ss")

    If Trim(slTexNeg) <> "" Then
        slJuntaObs = Trim(slTexNeg)
    End If

    slJuntaObs = vbCrLf & slJuntaObs & " Pedido Cancelado por " & sgNomUsuSis & " em " & sldata & vbCrLf

    If Trim(slObserva) <> "" Then
        slJuntaObs = slJuntaObs & "RAZÃO - " & Trim(slObserva) & vbCrLf
    End If

    slJuntaObs = Replace(slJuntaObs, "'", "´")
    slJuntaObs = Replace(slJuntaObs, """", "§")
    slJuntaObs = Replace(slJuntaObs, "§", "´")
    
    If mCanceladoAposNotificacao = True Then
        slJuntaObs = slJuntaObs & vbCrLf & vbCrLf & "PEDIDO CANCELADO PELA UNOCANN; REPRESENTANTE NÃO RESPONDEU A NOTIFICAÇÃO ENVIADA A ESTE PEDIDO."
    End If
    
    On Error GoTo TrataErro
    
    sgQuery = "Update pedido set Datcan = convert(datetime,'" & sldata & "',103), "
    sgQuery = sgQuery & " TexNeg = '" & Trim(slJuntaObs) & "', sitped = 'U'"
    sgQuery = sgQuery & " where NroPed = " & dlNroPed
    
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
    
    sgQuery = "Insert Notifica_Pedido Values ("
    sgQuery = sgQuery & dlNroPed & ", convert(datetime,'" & sldata & "',103), " & LgCodUsuSis & ", "
    sgQuery = sgQuery & " '" & Trim(slObserva) & "', ' ', '999', 'C')"
    
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
    
    LimpaGeral

    Exit Sub

TrataErro:

    cmdGerar.Enabled = True
    TxtObserva.Enabled = True
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

End Sub

Private Sub MskSenha_GotFocus()
    
    Call SelecionaTudo
    
End Sub

Private Sub MskSenha_LostFocus()
    
    Dim lExecutor As New ADODB.Command
    Dim lParametro As New ADODB.Parameter
    Dim lDiferencaData As Integer
    Dim lPeriodoAtraso As String
    
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
        .CommandText = "sp_cancelamento"
        
        Do While .Parameters.Count > 0
            .Parameters.Delete 0
        Loop
        
        Set lParametro = .CreateParameter("@Pedido", adInteger, adParamInput, 4, dlNroPed)
        .Parameters.Append lParametro
        
        Set Rs = .Execute
    
    End With
    
    Set lParametro = Nothing
    Set lExecutor = Nothing
    
    If Rs.EOF Then
        
        MsgBox "Pedido inexistente", vbExclamation + vbOKOnly, "Atenção!"
        
        MskSenha.SetFocus
        
        Exit Sub
        
    End If
    
    LblRep.Caption = Rs!NomRep
    LblCli.Caption = Rs!NomCli
    LblDatEnv.Caption = Format(Rs!DatEnv, "dd/mm/yyyy hh:mm:ss")
    LblDatEmi.Caption = Format(Rs!Datped, "dd/mm/yyyy hh:mm:ss")
    LblDatCan.Caption = Format(Rs!DatCan, "dd/mm/yyyy hh:mm:ss")

    ilCodRep = Rs!codrep
    slDatLib = Rs!Datlib
    slTexNeg = Rs!texneg
    slNomUsuAltLib = IIf(IsNull(Rs!nomusu), "", Rs!nomusu)

    If Trim(Rs!SitPed) = "C" Or Trim(Rs!SitPed) = "U" Then
        
        MsgBox "Este pedido está cancelado no sistema", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        MskSenha.SetFocus
        
        Exit Sub
        
    End If

    If Trim(Rs!ClasCor) = "R" And IsNull(Rs!datlibuno) And sgFlgUsu <> "L" Then
        
        MsgBox "Acesso negado a este pedido. Passível de liberação pelo Depto Comercial", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
   
        MskSenha.SetFocus
        
        Exit Sub
        
    End If

    If Trim(Rs!NroNot) <> "" Then
        
        MsgBox "Pedido Faturado não pode ser cancelado, Nota Fiscal - " & Format(Rs!NroNot, "000000"), vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        MskSenha.SetFocus
        
        Exit Sub
        
    End If

    If Trim(Rs!FlgDig) = "S" Then
        
        MsgBox "Acesso negado, este pedido já foi digitado", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        MskSenha.SetFocus
        
        Exit Sub
        
    End If
    
    '*************************************************************************
    'A partir de 26/01/2009, os representantes passam a ter prazo de sete
    'dias corridos para responder às notificações emitidas a seus pedidos. Se
    'não houver resposta nesse período, a empresa poderá cancelar esses
    'pedidos. (André Corrêa)
    '*************************************************************************
    
    lPeriodoAtraso = ""
    
    If Trim(Rs!DatLibAlt) <> "" Then
        
        lDiferencaData = DateDiff("d", Rs("DatLibAlt"), Now)
        
        If lDiferencaData < 3 Then
        
            MsgBox "Este pedido tem notificação para alteração aberta no dia " & Format(Rs!DatLibAlt, "dd/mm/yyyy") & ", às " & Format(Now, "hh:mm") & ", por " & slNomUsuAltLib & "." & vbCrLf & "O representante ainda tem (têm) " & 7 - lDiferencaData & " dia(s) para responder.", vbExclamation, "Atenção!"
            
            TxtObserva.Enabled = False
            
            Rs.Close
        
            Set Rs = Nothing
                
            cmdGerar.Enabled = False
            
            mCanceladoAposNotificacao = False
        
            Exit Sub
            
        Else
            
            If lDiferencaData - 7 = 0 Then
                lPeriodoAtraso = "hoje."
            Else
                lPeriodoAtraso = "há " & lDiferencaData - 7 & " dia(s)."
            End If
            
            If MsgBox("Este pedido tem notificação para alteração aberta no dia " & Format(Rs!DatLibAlt, "dd/mm/yyyy") & ", às " & Format(Now, "hh:mm") & ", por " & slNomUsuAltLib & "." & vbCrLf & "O prazo para que o representante emitisse resposta venceu " & lPeriodoAtraso & " Deseja cancelar este pedido?", vbQuestion + vbYesNo, "Atenção!") = vbNo Then
                
                TxtObserva.Enabled = False
            
                Rs.Close
        
                Set Rs = Nothing
                
                cmdGerar.Enabled = False
            
                mCanceladoAposNotificacao = False
        
                Exit Sub
            
            Else
            
                mCanceladoAposNotificacao = True
            
            End If
        
        End If

    End If

    Rs.Close
    
    Set Rs = Nothing

    MskSenha.Enabled = False

    Exit Sub

TrataErro:
    
    Rotina_Erro "MskSenha_LostFocus"

End Sub

Private Sub LimpaGeral()

    LblRep.Caption = ""
    LblCli.Caption = ""
    LblDatEnv.Caption = ""
    LblDatEmi.Caption = ""
    LblDatCan.Caption = ""
    MskSenha.Texto = ""
    TxtObserva.Text = ""

    ilCodRep = 0
    dlNroPed = 0
    vFileName = ""
    slArqPed = ""
    slDatLib = ""
    slString = ""
    slObserva = ""
    operacao = ""
    slTexNeg = ""
    slNomUsuAltLib = ""

    cmdGerar.Enabled = True
    TxtObserva.Enabled = True
    cmdGerar.Enabled = True
    BtoLimpar.Enabled = True
    BtoSair.Enabled = True
    MskSenha.Enabled = True

    MskSenha.SetFocus

End Sub
