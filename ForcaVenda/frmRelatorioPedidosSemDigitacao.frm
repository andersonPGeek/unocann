VERSION 5.00
Object = "{F454059D-91FE-11D2-8865-AD1268A0A52F}#2.0#0"; "ActiveDate.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRelatorioPedidosSemDigitacao 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Pedidos Sem Digitação"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5595
   Begin Crystal.CrystalReport CrystalReport 
      Left            =   2160
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCancelBtn=   0   'False
      WindowShowExportBtn=   0   'False
      WindowShowProgressCtls=   0   'False
      WindowShowPrintSetupBtn=   -1  'True
   End
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
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin rdActiveDate.ActiveDate adtDataInicio 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1200
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
   Begin rdActiveDate.ActiveDate adtDataFim 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
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
      TabIndex        =   8
      Top             =   240
      Width           =   1050
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "a"
      Height          =   195
      Left            =   1920
      TabIndex        =   7
      Top             =   1260
      Width           =   90
   End
   Begin VB.Label lblPeriodo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Período"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   570
   End
End
Attribute VB_Name = "frmRelatorioPedidosSemDigitacao"
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
    
    Screen.MousePointer = vbHourglass
    
    Set lReport = lReportApplication.OpenReport(App.Path & "\Relatorios\rptPedidosSemDigitacao.rpt")
    
    If (adtDataInicio.Text <> "" And adtDataFim.Text = "") Or (adtDataInicio.Text = "" And adtDataFim.Text <> "") Then
    
        MsgBox "Preencha todos os campos do filtro 'Período'.", vbCritical, Me.Caption
        
        adtDataInicio.SetFocus
        
        Exit Sub
        
    End If
    
    If adtDataInicio.Text <> "" And adtDataFim.Text <> "" Then
    
        If CDate(adtDataInicio.Text) > CDate(adtDataFim.Text) Then
    
            MsgBox "No filtro 'Período', a data inicial deve ser menor que a final.", vbCritical, Me.Caption
        
            adtDataInicio.SetFocus
        
            Exit Sub
    
        End If
        
    End If
    
    With lReport
        
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
        
        .ParameterFields(1).ClearCurrentValueAndRange
        
        If cboRepresentantes.ListIndex > -1 Then
            .ParameterFields(1).AddCurrentValue cboRepresentantes.ItemData(cboRepresentantes.ListIndex)
        Else
            .ParameterFields(1).AddCurrentValue 0
        End If
        
        .ParameterFields(2).ClearCurrentValueAndRange
        
        If IsDate(adtDataInicio.Text) = True Then
            .ParameterFields(2).AddCurrentValue Format(adtDataInicio.Text, "dd/mm/yyyy")
        Else
            .ParameterFields(2).AddCurrentValue "0"
        End If
        
        .ParameterFields(3).ClearCurrentValueAndRange
        
        If IsDate(adtDataFim.Text) = True Then
            .ParameterFields(3).AddCurrentValue Format(adtDataFim.Text, "dd/mm/yyyy")
        Else
            .ParameterFields(3).AddCurrentValue "0"
        End If
        
        If IsDate(adtDataInicio.Text) = True Then
            .FormulaFields(1).Text = "'De " & adtDataInicio.Text & " a " & adtDataFim.Text & "'"
        Else
            .FormulaFields(1).Text = "'Indefinido'"
        End If
        
    End With
    
    With frmViewer
    
        .Caption = "Relatório de Pedidos Não-Digitados"
        .CRViewer.ReportSource = lReport
        .CRViewer.ViewReport
        
        .Show
    
    End With
    
    Screen.MousePointer = vbDefault
    
    Set lReport = Nothing
    Set lReportApplication = Nothing
    
End Sub
Private Sub cmdLimpar_Click()

    cboRepresentantes.ListIndex = -1
    adtDataInicio.Text = ""
    adtDataFim.Text = ""
    cboRepresentantes.SetFocus

End Sub

Private Sub Form_Load()

    Dim lDados As New ADODB.Recordset
    Dim i As Integer
    
    With lDados
        .ActiveConnection = Conexao
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .Source = "SELECT CodRep, NomRep FROM Representante ORDER BY NomRep"
        .Open
    End With

    For i = 0 To lDados.RecordCount - 1
    
        cboRepresentantes.AddItem lDados("NomRep")
        cboRepresentantes.ItemData(i) = lDados("CodRep")
        
        lDados.MoveNext
    
    Next
    
    lDados.Close
        
    Set lDados = Nothing
    
    With CrystalReport
            
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowPrintSetupBtn = True
    
    End With
    
    Me.Left = (MDIProjUNO.Width - Me.Width) / 2
    Me.Top = (MDIProjUNO.Height - Me.Height) / 4

End Sub
