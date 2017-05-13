VERSION 5.00
Object = "{F454059D-91FE-11D2-8865-AD1268A0A52F}#2.0#0"; "ActiveDate.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRelatorioVendasRepresentante 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Vendas"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3960
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   1500
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin Crystal.CrystalReport CR 
      Left            =   2880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin rdActiveDate.ActiveDate adtDataInicial 
      Height          =   315
      Left            =   240
      TabIndex        =   3
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
      Left            =   2160
      TabIndex        =   4
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
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "a"
      Height          =   195
      Left            =   1920
      TabIndex        =   6
      Top             =   540
      Width           =   90
   End
   Begin VB.Label lblPeriodo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Período"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   570
   End
End
Attribute VB_Name = "frmRelatorioVendasRepresentante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    
    Set lReport = lReportApplication.OpenReport(App.Path & "\Relatorios\relVendasRepresentantes.rpt")
    
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
        
        If IsDate(adtDataInicial.Text) = True Then
            .ParameterFields(1).AddCurrentValue CDate(adtDataInicial.Text)
        Else
            .ParameterFields(1).AddCurrentValue CDate("01/01/2000")
        End If
        
        .ParameterFields(2).ClearCurrentValueAndRange
        
        If IsDate(adtDataFinal.Text) = True Then
            .ParameterFields(2).AddCurrentValue CDate(adtDataFinal.Text)
        Else
            .ParameterFields(2).AddCurrentValue CDate("01/01/2000")
        End If
        
        '************************************************************************
        'Preenche fórmulas.
        '************************************************************************
        
        If IsDate(adtDataInicial.Text) = True Then
            .FormulaFields(1).Text = "'De " & adtDataInicial.Text & " a " & adtDataFinal.Text & "'"
        Else
            .FormulaFields(1).Text = "'Indefinido'"
        End If
        
    End With
    
    '****************************************************************************
    'Visualiza relatório.
    '****************************************************************************
    
    With frmViewer
        
        .Caption = "Relatório de Vendas por Representantes"
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

    adtDataInicial.Text = ""
    adtDataFinal.Text = ""
    adtDataInicial.SetFocus

End Sub

Private Sub Form_Activate()

    adtDataInicial.SetFocus

End Sub

Private Sub Form_Load()

    '****************************************************************************
    'Centraliza o formulário.
    '****************************************************************************
    
    Me.Left = (MDIProjUNO.Width - Me.Width) / 2
    Me.Top = (MDIProjUNO.Height - Me.Height) / 4

End Sub
