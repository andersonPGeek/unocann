VERSION 5.00
Object = "{368CC970-FF03-11D7-9B5A-000B6A03449D}#1.1#0"; "Combo_DB.ocx"
Object = "{BEACF734-D8AC-11D7-9B57-000B6A03449D}#2.0#0"; "Masked.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRelCliCid 
   BackColor       =   &H00400040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RELA��O  DE  CLIENTES  POR  CIDADE"
   ClientHeight    =   4380
   ClientLeft      =   -225
   ClientTop       =   1500
   ClientWidth     =   7695
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7695
   Begin Project_Combo_DB.Combo_DB CboCidade 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   661
      Cols            =   0
      Cabecalho       =   -1  'True
   End
   Begin Crystal.CrystalReport rptcontprop 
      Left            =   600
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Project_Combo_DB.Combo_DB CboRepre 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   661
      Cols            =   0
      Cabecalho       =   -1  'True
   End
   Begin VB.CommandButton CmdAtualizar 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Gerar"
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
      Left            =   4920
      MaskColor       =   &H000080FF&
      Picture         =   "FrmRelCliCid.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   600
      Top             =   10200
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
      Left            =   6480
      Picture         =   "FrmRelCliCid.frx":0E52
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1095
   End
   Begin Project_Masked.Masked MskDias 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Texto           =   "0"
      CampoDb         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoTab         =   -1  'True
      ForeColor       =   0
      ValInteiro      =   4
      Texto           =   "0"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400040&
      Caption         =   "Cidade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   930
   End
   Begin VB.Label LblRepre 
      BackColor       =   &H00400040&
      Caption         =   "Representante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400040&
      Caption         =   "�ltimo    Pedido (Dias)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "FrmRelCliCid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim slCidade As String
Dim slrepre As String
Dim slNomeImpr  As String
Dim ilCodRep As Integer
Dim ilCodCli As Integer
Dim ilDias As Integer
Dim blLoad As Boolean
Dim logon As Integer

Private Sub BtoSair_Click()

    Unload Me

End Sub

Private Sub CboCidade_Consultar()

    slCidade = " "
    
    CboCidade.query = "Select distinct CidCli As Cidade, UfCli As UF From Cliente Where CidCli Like '" & CboCidade.Criterio & "%' order by cidcli"
    
End Sub

Private Sub CboCidade_GotFocus()
    
    Call SelecionaTudo
    
End Sub

Private Sub CboCidade_LostFocus()

    If CboCidade.Criterio <> "" Then
        slCidade = CboCidade.Criterio
    Else
        slCidade = "  "
    End If
    
End Sub

Private Sub CboRepre_Consultar()

    slrepre = ""
    
    CboRepre.query = "Select NomRep As Representante, CodRep As C�digo From Representante Where " & IIf(IsNumeric(CboRepre.Criterio), "Codrep", "Nomrep") & " Like '" & CboRepre.Criterio & "%' order by " & IIf(IsNumeric(CboRepre.Criterio), "Codrep", "Nomrep")
    
End Sub

Private Sub CboRepre_GotFocus()
    
    Call SelecionaTudo

End Sub

Private Sub CboRepre_LostFocus()

    If CboRepre.Criterio <> "" Then
        slrepre = CboRepre.Criterio
        'ilCodRep = CboRepre.Codigo
    Else
        ilCodRep = 0
        slrepre = ""
    End If
    
End Sub

Private Sub CmdAtualizar_Click()

    If Trim(MskDias.Texto) = "" Then
        ilDias = 0
    Else
        ilDias = Trim(MskDias.Texto)
    End If

    DoEvents

    If CboRepre.Criterio <> "" And APLICA = 0 Then
        ilCodRep = CboRepre.Codigo
    End If
  
    If APLICA = 1 Then
        slNomeImpr = App.Path & "\Relatorios\RelClienteCidade.rpt"
    Else
        slNomeImpr = App.Path & "\Relatorios\RelClienteCidade.rpt"
    End If

    With rptcontprop
        .ReportFileName = slNomeImpr
        .StoredProcParam(0) = ilCodRep
        .StoredProcParam(1) = ilDias
        .StoredProcParam(2) = slCidade
        'logon = .LogOnServer("pdsodbc.dll", "unocann", "unocann", "sa", "sysadmpss1")
        If APLICA = 1 Then
            .Connect = "DSN=" & "unocann" & ";UID=" & "sa" & ";PWD=" & "sysadmpss1"
        Else
            .Connect = "DSN=" & "unocann" & ";UID=" & "sa" & ";PWD=" & "#unoforte5600!"
        End If
        .DiscardSavedData = True
        .WindowState = crptMaximized
        .Action = 1
    End With
  
End Sub

Private Sub Form_Activate()

    'If APLICA = 1 Then
        'CboRepre.Codigo = sgRepresentante
        'ilCodRep = sgRepresentante
        'CboRepre.Habilitado = False
    'End If

    'DoEvents
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    'If Me.ActiveControl.Name = "GrdPedido" Then
        
        'If KeyAscii = 13 Then
            'GrdPedido_DblClick
        'End If
    
    'Else
        
        Call EventoEnter(KeyAscii)
        
    'End If
    
End Sub

Private Sub Form_Load()

    Me.Left = 0
    Me.Top = 0
    Me.Height = 4890
    Me.Width = 7815
    
    Set CboRepre.Conexao = Conexao
    Set CboCidade.Conexao = Conexao

    MskDias.TipodeDados numero

    If APLICA = 0 Then
        ilCodRep = 0
    Else
        ilCodRep = sgRepresentante
    End If

    slrepre = ""
    slCidade = " "
    ilCodCli = 0
    ilDias = 0
    blLoad = True

    If APLICA = 1 Then
        CboRepre.Visible = False
        LblRepre.Visible = False
    End If

    DoEvents

End Sub
