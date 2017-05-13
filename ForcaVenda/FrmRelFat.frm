VERSION 5.00
Object = "{368CC970-FF03-11D7-9B5A-000B6A03449D}#1.1#0"; "Combo_DB.ocx"
Object = "{F454059D-91FE-11D2-8865-AD1268A0A52F}#2.0#0"; "ActiveDate.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRelFat 
   BackColor       =   &H00400040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RELATÓRIO  DE  FATURAMENTO"
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
   Begin Crystal.CrystalReport rptcontprop 
      Left            =   600
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin Project_Combo_DB.Combo_DB CboRepre 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   6105
      _ExtentX        =   10769
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
      Picture         =   "FrmRelFat.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Picture         =   "FrmRelFat.frx":0E52
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1095
   End
   Begin rdActiveDate.ActiveDate ActDtfim 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1305
      _ExtentX        =   2302
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
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
      TabIndex        =   2
      Top             =   1320
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   661
      Cols            =   0
      Cabecalho       =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400040&
      Caption         =   "Cliente"
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
      TabIndex        =   9
      Top             =   1080
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
      TabIndex        =   8
      Top             =   1920
      Width           =   1530
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400040&
      Caption         =   "Data Final"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400040&
      Caption         =   "Data Inicial"
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
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmRelFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim slremet As String
Dim slrepre As String
Dim slNomeImpr  As String
Dim ilCodRep As Integer
Dim ilCodCli As Integer
Dim blLoad As Boolean
Dim logon As Integer

Private Sub BtoSair_Click()

    Unload Me
  
End Sub

Private Sub CboRemet_Consultar()

    slremet = ""
    
    CboRemet.query = "Select NomCli As Cliente, CodCli As Código, CgcCli as CNPJ From Cliente Where " & IIf(IsNumeric(CboRemet.Criterio), "CodCli", "NomCli") & " Like '" & CboRemet.Criterio & "%' order by " & IIf(IsNumeric(CboRemet.Criterio), "CodCli", "NomCli")
    
End Sub

Private Sub CboRemet_GotFocus()

    Call SelecionaTudo

End Sub

Private Sub CboRemet_LostFocus()

    If CboRemet.Criterio <> "" Then
        slremet = CboRemet.Criterio
        'ilCodCli = CboRemet.Codigo
    Else
        ilCodCli = 0
        slremet = ""
    End If
    
End Sub

Private Sub CboRepre_Consultar()

    slrepre = ""
    
    CboRepre.query = "Select NomRep As Representante, CodRep As Código From Representante Where " & IIf(IsNumeric(CboRepre.Criterio), "Codrep", "Nomrep") & " Like '" & CboRepre.Criterio & "%' order by " & IIf(IsNumeric(CboRepre.Criterio), "Codrep", "Nomrep")

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

    DoEvents
    
    If ActDtini.Text = "" Or ActDtfim.Text = "" Then
        
        MsgBox "Intervalo de datas inválido", vbInformation
        
        ActDtini.SetFocus
        
        Exit Sub
        
    End If

    If CDate(ActDtini.Text) > CDate(ActDtfim.Text) Or Year(CDate(ActDtini.Text)) < 1950 Or Year(CDate(ActDtfim.Text)) < 1950 Then
        
        MsgBox "Intervalo de datas inválido", vbInformation
        
        ActDtini.SetFocus
        
        Exit Sub
        
    End If
        
    If CboRepre.Criterio <> "" And APLICA = 0 Then
        ilCodRep = CboRepre.Codigo
    End If

    If CboRemet.Criterio <> "" Then
        
        ilCodCli = CboRemet.Codigo
        
        If APLICA = 0 Then
            
            CboRepre.Criterio = ""
            
            ilCodRep = 0
            
        End If
        
    End If
    
    If APLICA = 1 Then
        slNomeImpr = App.Path & "\Relatorios\RelFat.rpt"
    Else
        slNomeImpr = App.Path & "\Relatorios\RelFat.rpt"
    End If

    With rptcontprop
        .ReportFileName = slNomeImpr
        .StoredProcParam(0) = ilCodRep
        .StoredProcParam(1) = ilCodCli
        .StoredProcParam(2) = Format(ActDtini.Text, "dd/mm/yyyy")
        .StoredProcParam(3) = Format(ActDtfim.Text, "dd/mm/yyyy")
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

    If APLICA = 1 Then
        CboRepre.Codigo = ilCodRep
        CboRepre.Habilitado = False
    End If

    DoEvents
    
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
    
    Set CboRemet.Conexao = Conexao
    Set CboRepre.Conexao = Conexao

    If APLICA = 0 Then
        ilCodRep = 0
    Else
        ilCodRep = sgRepresentante
    End If

    slremet = ""
    slrepre = ""
    ilCodCli = 0
    blLoad = True

    If APLICA = 1 Then
        CboRepre.Visible = False
        LblRepre.Visible = False
    End If

End Sub
