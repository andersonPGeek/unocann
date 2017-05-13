VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{BEACF734-D8AC-11D7-9B57-000B6A03449D}#2.0#0"; "Masked.ocx"
Object = "{F454059D-91FE-11D2-8865-AD1268A0A52F}#2.0#0"; "ActiveDate.ocx"
Begin VB.Form FrmPosiLig 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Posição de Ligações"
   ClientHeight    =   7590
   ClientLeft      =   -225
   ClientTop       =   1500
   ClientWidth     =   10935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   10935
   Begin VB.CommandButton BtoGrava 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Consultar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      MaskColor       =   &H00FF0000&
      Picture         =   "FrmPosPed.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1095
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
      Height          =   735
      Left            =   9840
      Picture         =   "FrmPosPed.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid GrdLigacao 
      Height          =   6375
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      BackColor       =   14737632
      ForeColor       =   12582912
      BackColorFixed  =   192
      ForeColorFixed  =   16777215
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   255
      GridColorFixed  =   128
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin rdActiveDate.ActiveDate ActDtfim 
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   360
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
      Left            =   2280
      TabIndex        =   1
      Top             =   360
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
   Begin Project_Masked.Masked MskNroLigacao 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
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
      BackColor       =   -2147483628
      AutoTab         =   -1  'True
      ValInteiro      =   6
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
      Left            =   8760
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "Ligação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   8295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Data Inicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Data Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmPosiLig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ilCodCli As Integer
Dim blLoad As Boolean

Function CompoeGridPed() As Boolean

    Dim blI As Integer
    
    On Error GoTo TratarErro

    'Pedidos não Liberados
    sgQuery = "select a.SeqLig, a.DatIniLig, a.DatFimLig, a.tipLig, a.codcli, b.NomUsu, c.NomCli"
    sgQuery = sgQuery & "  from LIGACAO a, USUARIO b, CLIENTE c"
    sgQuery = sgQuery & "  Where a.codusu = b.codusu"
    sgQuery = sgQuery & "    and a.codcli *= c.codcli"
    sgQuery = sgQuery & "    and a.datfimlig is not null"
    
    If MskNroLigacao.Texto > 0 Then
        sgQuery = sgQuery & "    and a.seqlig = " & Trim(MskNroLigacao.Texto)
    End If
    
    If Trim(ActDtini.Text) <> "" Then
        sgQuery = sgQuery & "    and a.datinilig between convert(datetime,'" & Trim(ActDtini.Text) & "',103)"
        sgQuery = sgQuery & "                     and convert(datetime,'" & Trim(ActDtfim.Text) & " 23:59:59',103)"
    Else
        sgQuery = sgQuery & "    and a.datinilig between (getdate() - 180) and Getdate()"
    End If
    
    If igCodCli <> 0 And MskNroLigacao.Texto <= 0 Then
        sgQuery = sgQuery & "    and a.codcli = " & igCodCli
    End If
    
    sgQuery = sgQuery & "    order by 1 desc"

    Consulta sgQuery
    
    If Rs.RecordCount > 100 Then
        Lbl100.Visible = True
    Else
        Lbl100.Visible = False
    End If

    GrdLigacao.Rows = 1
    GrdLigacao.Visible = False
    blI = 0
    
    Do While Not Rs.EOF
    
        If blI > 100 Then
            Exit Do
        End If
   
        blI = blI + 1

        GrdLigacao.Rows = GrdLigacao.Rows + 1
        GrdLigacao.TextMatrix(blI, 0) = Format(Trim(Rs!seqlig), "000000")
        GrdLigacao.TextMatrix(blI, 1) = IIf(IsNull(Rs!Codcli), "", Format(Trim(Rs!Codcli), "00000"))
        GrdLigacao.TextMatrix(blI, 2) = IIf(IsNull(Rs!NomCli), "", Rs!NomCli)
        GrdLigacao.TextMatrix(blI, 3) = Format(Trim(Rs!Datinilig), "dd/mm/yyyy hh:mm:ss")
        GrdLigacao.TextMatrix(blI, 4) = Format(Trim(Rs!Datfimlig), "dd/mm/yyyy hh:mm:ss")
        GrdLigacao.TextMatrix(blI, 5) = IIf(Trim(Rs!tiplig) = 1, "Ativa", "Receptiva")
        GrdLigacao.TextMatrix(blI, 6) = Rs!nomusu
        GrdLigacao.Row = blI
   
        Rs.MoveNext
    
    Loop

    Rs.Close
    
    Set Rs = Nothing
    
    GrdLigacao.Visible = True

    DoEvents

    If MskNroLigacao.Texto > 0 And blI = 0 Then
        MsgBox "Ligação inexistente", vbExclamation + vbOKOnly, "Atenção!"
    End If

    Exit Function

TratarErro:

    Rotina_Erro "CompoeGridPed"

End Function

Private Sub BtoGrava_Click()

    DoEvents

    If MskNroLigacao.Texto = "" Then
        MskNroLigacao.Texto = 0
    End If

    If MskNroLigacao.Texto > 0 Then
        
        ActDtini.Text = ""
        ActDtfim.Text = ""
        
        GoTo Compoe
        
    End If

    If Trim(ActDtini.Text) = "" And Trim(ActDtfim.Text) = "" Then
        GoTo Compoe
    End If
   
    If Trim(ActDtini.Text) <> "" Or Trim(ActDtfim.Text) <> "" Then
        
        If Trim(ActDtini.Text) = "" Or Trim(ActDtfim.Text) = "" Then
            
            MsgBox "Intervalo de datas inválido", vbInformation
            
            ActDtini.SetFocus
            
            Exit Sub
        
        ElseIf CDate(ActDtini.Text) > CDate(ActDtfim.Text) Or Year(CDate(ActDtini.Text)) < 1950 Or Year(CDate(ActDtfim.Text)) < 1950 Then
            
            MsgBox "Intervalo de datas inválido", vbInformation
            
            ActDtini.SetFocus
            
            Exit Sub
            
        End If
        
    End If

Compoe:

    CompoeGridPed

End Sub

Private Sub BtoSair_Click()

    lgSeqLig = 0
    
 '   FrmTMKPrincipal.Enabled = True
    
    Unload Me
  
End Sub

Private Sub Form_Activate()

    blLoad = True
    bgPosLig = True
    
   ' FrmTMKPrincipal.Enabled = False
    
    CompoeGridPed
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    'If Me.ActiveControl.Name = "grdligacao" Then
        
        'If KeyAscii = 13 Then
            'grdligacao_DblClick
        'End If
    
    'Else
        
        Call EventoEnter(KeyAscii)
        
    'End If
    
End Sub

Private Sub Form_Load()

    Me.Left = 0
    Me.Top = 0
    Me.Height = 8070
    Me.Width = 11055

    MskNroLigacao.Texto = 0
    ilCodCli = 0
    blLoad = True
    bgBloqPed = False
    bgConsultaPed = False

    GrdLigacao.TextMatrix(0, 0) = "Ligação"
    GrdLigacao.ColWidth(0) = 700
    GrdLigacao.TextMatrix(0, 1) = "Cod.Cli"
    GrdLigacao.ColWidth(1) = 600
    GrdLigacao.TextMatrix(0, 2) = "Cliente"
    GrdLigacao.ColWidth(2) = 3600
    GrdLigacao.TextMatrix(0, 3) = "Início"
    GrdLigacao.ColWidth(3) = 1600
    GrdLigacao.TextMatrix(0, 4) = "Final"
    GrdLigacao.ColWidth(4) = 1600
    GrdLigacao.TextMatrix(0, 5) = "Tipo"
    GrdLigacao.ColWidth(5) = 850
    GrdLigacao.TextMatrix(0, 6) = "Operador(a)"
    GrdLigacao.ColWidth(6) = 1800

End Sub

Private Sub grdligacao_DblClick()

    If GrdLigacao.RowSel = 0 Then
        Exit Sub
    End If
  
    lgSeqLig = GrdLigacao.TextMatrix(GrdLigacao.RowSel, 0)
    
    '.Enabled = True
    
    Unload Me
    
End Sub

Private Sub MskNroLigacao_GotFocus()
    
    Call SelecionaTudo
    
End Sub
