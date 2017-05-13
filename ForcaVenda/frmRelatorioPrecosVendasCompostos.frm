VERSION 5.00
Object = "{F454059D-91FE-11D2-8865-AD1268A0A52F}#2.0#0"; "ActiveDate.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRelatorioPrecoVendaComposto 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Preços de Venda dos Compostos"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   9450
   Begin VB.ComboBox cboRepresentantes 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   750
      Width           =   1215
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Default         =   -1  'True
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox cboGruposProdutos 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   3495
   End
   Begin VB.ComboBox cboProdutos 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1320
      Width           =   3495
   End
   Begin Crystal.CrystalReport CR 
      Left            =   6960
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin rdActiveDate.ActiveDate adtDataInicial 
      Height          =   315
      Left            =   4080
      TabIndex        =   1
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
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
   Begin rdActiveDate.ActiveDate adtDataFinal 
      Height          =   315
      Left            =   6000
      TabIndex        =   2
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.Label lblRepresentante 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Representante"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   1050
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "a"
      Height          =   195
      Left            =   5760
      TabIndex        =   11
      Top             =   540
      Width           =   90
   End
   Begin VB.Label lblPeriodo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Período"
      Height          =   195
      Left            =   4080
      TabIndex        =   10
      Top             =   240
      Width           =   570
   End
   Begin VB.Label lblGrupoProduto 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Grupos de Produtos"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   1410
   End
   Begin VB.Label lblProduto 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Produtos"
      Height          =   195
      Left            =   4080
      TabIndex        =   8
      Top             =   1080
      Width           =   630
   End
End
Attribute VB_Name = "frmRelatorioPrecoVendaComposto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDados As New ADODB.Recordset
Private mExecutor As New ADODB.Command
Private mParametro As New ADODB.Parameter
Private mSQL As String
Private i As Integer

Private Sub cboGruposProdutos_Click()

    '****************************************************************************
    'Carrega produtos.
    '****************************************************************************
    
    i = 0
    
    If cboGruposProdutos.ListIndex > -1 Then
        
        cboProdutos.Clear
        
        mSQL = "SELECT CodPrd, DscPrd FROM Produto WHERE IdeGrp = " & cboGruposProdutos.ItemData(cboGruposProdutos.ListIndex) & " ORDER BY DscPrd"
        
        mDados.Open mSQL, Conexao, adOpenForwardOnly, adLockOptimistic
        
        Do While mDados.EOF = False
            
            cboProdutos.AddItem mDados("DscPrd")
            cboProdutos.ItemData(i) = mDados("CodPrd")
            
            i = i + 1
            
            mDados.MoveNext
            
        Loop
        
        mDados.Close
        
    End If

End Sub

Private Sub cmdFechar_Click()

    Unload Me

End Sub

Private Sub cmdGerar_Click()

    Dim lReportApplication As New CRAXDRT.Application
    Dim lReport As New CRAXDRT.Report
    Dim i As Integer
    
    '****************************************************************************
    'Inicializa o relatório e loga na base de dados.
    '****************************************************************************
    
    Screen.MousePointer = vbHourglass
    
    Set lReport = lReportApplication.OpenReport(App.Path & "\Relatorios\relPrecosVendasCompostos.rpt")
    
    If (adtDataInicial.Text <> "" And adtDataFinal.Text = "") Or (adtDataInicial.Text = "" And adtDataFinal.Text <> "") Then
    
        MsgBox "Preencha todos os campos do filtro 'Período'.", vbCritical, Me.Caption
        
        adtDataInicial.SetFocus
        
        Exit Sub
        
    End If
    
    If adtDataInicial.Text <> "" And adtDataFinal.Text <> "" Then
    
        If CDate(adtDataInicial.Text) > CDate(adtDataFinal.Text) Then
    
            MsgBox "No filtro 'Período', a data inicial deve ser menor que a final.", vbCritical, Me.Caption
        
            adtDataInicial.SetFocus
        
            Exit Sub
    
        End If
        
    End If
    
    With lReport
        
        '************************************************************************
        'Loga na base de dados.
        '************************************************************************
        
        If APLICA = 1 Then
            .Database.LogOnServer "pdsodbc.dll", "Unocann", "", "sa", "sysadmpss1"
        Else
            .Database.LogOnServer "pdsodbc.dll", "Unocann", "", "sa", "#unoforte5600!"
        End If
        
        For i = 1 To .Database.Tables.Count
            If APLICA = 1 Then
                .Database.Tables(i).SetLogOnInfo "Unocann", "", "sa", "sysadmpss1"
            Else
                .Database.Tables(i).SetLogOnInfo "Unocann", "", "sa", "#unoforte5600!"
            End If
        Next
        
        .DiscardSavedData
        
        '************************************************************************
        'Preenche parâmetros.
        '************************************************************************
    
        .ParameterFields(1).ClearCurrentValueAndRange
        
        If cboRepresentantes.ListIndex > -1 Then
            .ParameterFields(1).AddCurrentValue cboRepresentantes.ItemData(cboRepresentantes.ListIndex)
        Else
            .ParameterFields(1).AddCurrentValue 0
        End If
        
        .ParameterFields(2).ClearCurrentValueAndRange
        
        If cboGruposProdutos.ListIndex > -1 Then
            .ParameterFields(2).AddCurrentValue cboGruposProdutos.ItemData(cboGruposProdutos.ListIndex)
        Else
            .ParameterFields(2).AddCurrentValue 0
        End If
        
        .ParameterFields(3).ClearCurrentValueAndRange
        
        If cboProdutos.ListIndex > -1 Then
            .ParameterFields(3).AddCurrentValue cboProdutos.ItemData(cboProdutos.ListIndex)
        Else
            .ParameterFields(3).AddCurrentValue 0
        End If
        
        .ParameterFields(4).ClearCurrentValueAndRange
        
        If IsDate(adtDataInicial.Text) = True Then
            .ParameterFields(4).AddCurrentValue CDate(adtDataInicial.Text)
        Else
            .ParameterFields(4).AddCurrentValue CDate("01/01/2000")
        End If
        
        .ParameterFields(5).ClearCurrentValueAndRange
        
        If IsDate(adtDataFinal.Text) = True Then
            .ParameterFields(5).AddCurrentValue CDate(adtDataFinal.Text)
        Else
            .ParameterFields(5).AddCurrentValue CDate("01/01/2000")
        End If
        
        '************************************************************************
        'Preenche fórmulas.
        '************************************************************************
        
        If cboRepresentantes.ListIndex >= 0 Then
            .FormulaFields(1).Text = "'" & cboRepresentantes.Text & "'"
        Else
            .FormulaFields(1).Text = "'Indefinido'"
        End If
        
        If IsDate(adtDataInicial.Text) = True Then
            .FormulaFields(2).Text = "'De " & adtDataInicial.Text & " a " & adtDataFinal.Text & "'"
        Else
            .FormulaFields(2).Text = "'Indefinido'"
        End If
        
    End With
    
    '****************************************************************************
    'Visualiza relatório.
    '****************************************************************************
    
    With frmViewer
        
        .Caption = "Relatório de Preços de Venda dos Compostos"
        .CRViewer.ReportSource = lReport
        .CRViewer.ViewReport
        .Show
        
    End With
    
    '****************************************************************************
    'Finaliza a execução da rotina.
    '****************************************************************************
    
    Screen.MousePointer = vbDefault
    
    Set lReport = Nothing
    Set lReportApplication = Nothing

End Sub

Private Sub cmdLimpar_Click()

    cboRepresentantes.ListIndex = -1
    cboGruposProdutos.ListIndex = -1
    adtDataInicial.Text = ""
    adtDataFinal.Text = ""
    cboProdutos.Clear

End Sub

Private Sub Form_Load()

    '****************************************************************************
    'Carrega representantes.
    '****************************************************************************
    
    i = 0
    mSQL = "SELECT CodRep, NomRep FROM Representante ORDER BY NomRep"
    
    mDados.Open mSQL, Conexao, adOpenForwardOnly, adLockOptimistic
    
    Do While mDados.EOF = False
        
        cboRepresentantes.AddItem mDados("NomRep")
        cboRepresentantes.ItemData(i) = mDados("CodRep")
        
        i = i + 1
        
        mDados.MoveNext
        
    Loop
    
    mDados.Close
    
    '****************************************************************************
    'Carrega grupos de produtos.
    '****************************************************************************
    
    i = 0
    mSQL = "SELECT IdeGrp, NomGrp FROM Grupo_Produto ORDER BY NomGrp"
    
    mDados.Open mSQL, Conexao, adOpenForwardOnly, adLockOptimistic
    
    Do While mDados.EOF = False
        
        cboGruposProdutos.AddItem mDados("NomGrp")
        cboGruposProdutos.ItemData(i) = mDados("IdeGrp")
        
        i = i + 1
        
        mDados.MoveNext
        
    Loop
    
    i = 0
    
    mDados.Close
    
    '****************************************************************************
    'Centraliza o formulário.
    '****************************************************************************
    
    Me.Left = (MDIProjUNO.Width - Me.Width) / 2
    Me.Top = (MDIProjUNO.Height - Me.Height) / 4

End Sub
