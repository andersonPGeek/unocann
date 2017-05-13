VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultarPedidos 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar Pedidos"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   8595
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdSaldos 
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2566
      _Version        =   393216
      BackColorBkg    =   -2147483624
      Enabled         =   0   'False
      ScrollBars      =   2
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtCliente 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Frame fraTipo 
      BackColor       =   &H80000018&
      Caption         =   "Tipo"
      Height          =   1335
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1575
      Begin VB.OptionButton optSaldo 
         BackColor       =   &H80000018&
         Caption         =   "Saldo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optPedido 
         BackColor       =   &H80000018&
         Caption         =   "Pedido"
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.TextBox txtPedido 
      Height          =   315
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   6600
      X2              =   6600
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Label lblMensagem 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   6015
   End
   Begin VB.Label lblSaldos 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Outros Saldos"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label lblCliente 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Cliente"
      Height          =   195
      Left            =   2040
      TabIndex        =   10
      Top             =   960
      Width           =   480
   End
   Begin VB.Label lblPedido 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Pedido"
      Height          =   195
      Left            =   2040
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmConsultarPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConsultar_Click()
    
    Dim lDados As New ADODB.Recordset
    Dim lSQL As String
    Dim lPedido As String
    
    ConfigurarGrid
    
    '************************************************************************************
    'Faz a consistência.
    '************************************************************************************
    
    If txtPedido.Text = "" Then
        
        MsgBox "Informe o pedido que você deseja consultar.", vbCritical, Me.Caption
        
        Exit Sub
        
    End If
    
    '************************************************************************************
    'Consulta para saber se o número é de pedido. Se for, obtém os dados a serem
    'exibidos.
    '************************************************************************************
    
    lSQL = "SELECT P.*, C.NomCli FROM Pedido P INNER JOIN Cliente C ON C.CodCli = P.CodCli WHERE P.NroPed = " & txtPedido.Text
    
    lDados.Open lSQL, Conexao, adOpenStatic, adLockOptimistic
    
    If lDados.RecordCount = 0 Then
    
        lDados.Close
        
        '********************************************************************************
        'Como o número informado não é de um pedido, consulta para saber se o mesmo
        'refere-se a um saldo. Se for, obtém o nome do cliente e o exibe.
        '********************************************************************************
        
        lSQL = "SELECT PS.*, C.NomCli FROM Pedido_Saldo PS INNER JOIN Pedido P ON P.NroPed = PS.NroPed INNER JOIN Cliente C ON C.CodCli = P.CodCli WHERE NroPedSdo = " & txtPedido.Text
    
        lDados.Open lSQL, Conexao, adOpenStatic, adLockOptimistic
        
        If lDados.RecordCount = 0 Then
            
            MsgBox "O pedido informado não foi encontrado.", vbExclamation, Me.Caption
            
        Else
        
            optPedido.Value = False
            optSaldo.Value = True
            txtCliente.Text = UCase(Trim(lDados("NomCli")))
            
            lDados.Close
            
            '****************************************************************************
            'O pedido é um saldo: consulta para obter os dados a serem exibidos.
            '****************************************************************************
            
            lSQL = "SELECT * FROM Pedido_Saldo WHERE NroPedSdo = " & txtPedido.Text
            
            lDados.Open lSQL, Conexao, adOpenStatic, adLockOptimistic
            
            With grdSaldos
            
                lPedido = lDados("NroPed")
            
                Do While lDados.EOF = False
                
                    If .TextMatrix(1, 0) = "" Then
                        
                        .TextMatrix(1, 0) = lDados("NroPed")
                        .TextMatrix(1, 1) = lDados("NroPedSdo")
                        .TextMatrix(1, 2) = lDados("DatPed")
                        
                        If IsNull(lDados("DatEmiNot")) = True Then
                            .TextMatrix(1, 3) = "Não"
                        Else
                            .TextMatrix(1, 3) = "Sim"
                        End If
                        
                    Else
                    
                        .AddItem lDados("NroPed")
                        
                        .Row = .Rows - 1
                        
                        .TextMatrix(.Row, 1) = lDados("NroPedSdo")
                        .TextMatrix(.Row, 2) = lDados("DatPed")
                        
                        If IsNull(lDados("DatEmiNot")) = True Then
                            .TextMatrix(.Row, 3) = "Não"
                        Else
                            .TextMatrix(.Row, 3) = "Sim"
                        End If
                        
                    End If
                    
                    lDados.MoveNext
                
                Loop
                
                lblMensagem.Caption = txtPedido.Text & " É SALDO DO PEDIDO " & lPedido & "."
                
            End With
            
        End If
        
    Else
    
        optPedido.Value = True
        optSaldo.Value = False
        txtCliente.Text = lDados("NomCli")
        
        lDados.Close
        
        '********************************************************************************
        'Consulta para determinar se o pedido têm ou não algum saldo. Se tiver, os dados
        'desses saldos serão exibidos.
        '********************************************************************************
        
        lSQL = "SELECT * FROM Pedido_Saldo WHERE NroPed = " & txtPedido.Text
        
        lDados.Open lSQL, Conexao, adOpenStatic, adLockOptimistic
        
        If lDados.RecordCount = 0 Then
            
            lblMensagem.Caption = txtPedido.Text & " É PEDIDO SEM SALDO."
            
        Else
            
            lblMensagem.Caption = txtPedido.Text & " É PEDIDO COM SALDO(S)."
            
            With grdSaldos
            
                Do While lDados.EOF = False
                
                    If .TextMatrix(1, 0) = "" Then
                        
                        .TextMatrix(1, 0) = lDados("NroPed")
                        .TextMatrix(1, 1) = lDados("NroPedSdo")
                        .TextMatrix(1, 2) = lDados("DatPed")
                        
                        If IsNull(lDados("DatEmiNot")) = True Then
                            .TextMatrix(1, 3) = "Não"
                        Else
                            .TextMatrix(1, 3) = "Sim"
                        End If
                        
                    Else
                    
                        .AddItem lDados("NroPed")
                        
                        .Row = .Rows - 1
                        
                        .TextMatrix(.Row, 1) = lDados("NroPedSdo")
                        .TextMatrix(.Row, 2) = lDados("DatPed")
                        
                        If IsNull(lDados("DatEmiNot")) = True Then
                            .TextMatrix(.Row, 3) = "Não"
                        Else
                            .TextMatrix(.Row, 3) = "Sim"
                        End If
                        
                    End If
                    
                    lDados.MoveNext
                
                Loop
                
            End With
            
        End If
        
    End If
    
    lDados.Close
    
    Set lDados = Nothing
    
    With txtPedido
        .SelStart = 0
        .SelLength = Len(txtPedido.Text)
        .SetFocus
    End With
    
End Sub

Private Sub cmdFechar_Click()

    Unload Me

End Sub

Private Sub ConfigurarGrid()

    With grdSaldos
    
        .Clear
        
        .Rows = 2
        .Cols = 4
        
        .FixedRows = 1
        .FixedCols = 0
        
        .Row = 0
            
        .Col = 0
        .ColWidth(0) = 1500
        .ColAlignment(0) = 3
        .Text = "PEDIDO"
            
        .Col = 1
        .ColWidth(1) = 1500
        .ColAlignment(1) = 3
        .Text = "SALDO"
            
        .Col = 2
        .ColWidth(2) = 1500
        .ColAlignment(2) = 3
        .Text = "DATA"
            
        .Col = 3
        .ColWidth(3) = 1430
        .ColAlignment(3) = 3
        .Text = "FATURADO?"
    
    End With

End Sub

Private Sub cmdLimpar_Click()

    ConfigurarGrid
    
    optPedido.Value = False
    optSaldo.Value = False
    txtPedido.Text = ""
    txtCliente.Text = ""
    
    txtPedido.SetFocus

End Sub

Private Sub Form_Load()

    ConfigurarGrid
    
    Me.Left = (MDIProjUNO.Width - Me.Width) / 2
    Me.Top = (MDIProjUNO.Height - Me.Height) / 4

End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub
