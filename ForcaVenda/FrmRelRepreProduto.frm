VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{368CC970-FF03-11D7-9B5A-000B6A03449D}#1.1#0"; "Combo_DB.ocx"
Object = "{F454059D-91FE-11D2-8865-AD1268A0A52F}#2.0#0"; "ActiveDate.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRelRepreProduto 
   BackColor       =   &H00FFECEC&
   Caption         =   "RELATÓRIO DE VENDAS NO PERÍODO POR REPRESENTANTE"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   2250
   ClientWidth     =   13815
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   13815
   Begin Crystal.CrystalReport rptcontprop 
      Left            =   360
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame FraFiltro 
      BackColor       =   &H00FFECEC&
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      TabIndex        =   13
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton CmdPrd 
         BackColor       =   &H00FFECEC&
         Caption         =   "PRODUTO"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdGrp 
         BackColor       =   &H00FFECEC&
         Caption         =   "GRUPO"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton BtoLimpaCTRC 
      BackColor       =   &H00C0FFFF&
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
      Height          =   885
      Left            =   11280
      Picture         =   "FrmRelRepreProduto.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1095
   End
   Begin Project_Combo_DB.Combo_DB CboRepre 
      Height          =   4695
      Left            =   3720
      TabIndex        =   2
      Top             =   360
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   8281
      Cols            =   0
      Cabecalho       =   -1  'True
   End
   Begin rdActiveDate.ActiveDate ActDtfim 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
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
      Left            =   120
      TabIndex        =   0
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
   Begin VB.CommandButton BtoGerar 
      BackColor       =   &H00FFC0C0&
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
      Height          =   885
      Left            =   9840
      MaskColor       =   &H000080FF&
      Picture         =   "FrmRelRepreProduto.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton CmdIncUtil 
      BackColor       =   &H00FFECEC&
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      MaskColor       =   &H00004040&
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Seleciona uma nota"
      Top             =   3000
      Width           =   645
   End
   Begin VB.CommandButton CmdIncDisp 
      BackColor       =   &H00FFECEC&
      Caption         =   "<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Retira uma nota selecionada"
      Top             =   4200
      Width           =   645
   End
   Begin VB.CommandButton CmdSair 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   12585
      Picture         =   "FrmRelRepreProduto.frx":1294
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1080
   End
   Begin MSFlexGridLib.MSFlexGrid GrdNfDisp 
      Height          =   5535
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9763
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16772332
      ForeColor       =   192
      BackColorFixed  =   16772332
      BackColorSel    =   65535
      ForeColorSel    =   16777215
      BackColorBkg    =   16772332
      GridColorFixed  =   192
      ScrollBars      =   2
      BorderStyle     =   0
      FormatString    =   "Código   |     Descrição                                                                                             ||"
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
   Begin MSFlexGridLib.MSFlexGrid GrdNfSel 
      Height          =   5550
      Left            =   7560
      TabIndex        =   7
      Top             =   840
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   9790
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16772332
      ForeColor       =   32768
      BackColorFixed  =   16772332
      BackColorBkg    =   16772332
      GridColorFixed  =   49152
      ScrollBars      =   2
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   $"FrmRelRepreProduto.frx":16D6
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFECEC&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFECEC&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFECEC&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmRelRepreProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim slProduto  As String
Dim slGrupo As String
Dim slrepre As String
Dim ilCodRep As Integer
Dim blPrd As Boolean
Dim blGrp As Boolean
Dim ilcont As Integer
Dim slNomeImpr  As String

Private Sub btoGerar_Click()

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
        
    If CboRepre.Criterio <> "" Then
        ilCodRep = CboRepre.Codigo
    End If

    slProduto = " "
    slGrupo = " "

    If blPrd = True Then
        
        For ilcont = 1 To GrdNfSel.Rows - 1
            
            If Trim(slProduto) = "" Then
                slProduto = Trim(GrdNfSel.TextMatrix(ilcont, 0))
            Else
                slProduto = slProduto & ", " & Trim(GrdNfSel.TextMatrix(ilcont, 0))
            End If
            
        Next ilcont
        
    End If

    If blGrp = True Then
        
        For ilcont = 1 To GrdNfSel.Rows - 1
            
            If Trim(slGrupo) = "" Then
                slGrupo = Trim(GrdNfSel.TextMatrix(ilcont, 0))
            Else
                slGrupo = slGrupo & ", " & Trim(GrdNfSel.TextMatrix(ilcont, 0))
            End If
            
        Next ilcont
        
    End If

    slNomeImpr = App.Path & "\Relatorios\RepreProduto.rpt"

    With rptcontprop
        .ReportFileName = slNomeImpr
        .StoredProcParam(0) = ilCodRep
        .StoredProcParam(1) = slProduto
        .StoredProcParam(2) = slGrupo
        .StoredProcParam(3) = Format(ActDtini.Text, "dd/mm/yyyy")
        .StoredProcParam(4) = Format(ActDtfim.Text, "dd/mm/yyyy")
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

Private Sub BtoLimpaCTRC_Click()
    
    slProduto = ""
    slGrupo = ""
    blPrd = False
    blGrp = False
    
    CmdGrp.BackColor = &HFFECEC
    CmdPrd.BackColor = &HFFECEC
    GrdNfDisp.Rows = 1
    GrdNfSel.Rows = 1
    CmdIncUtil.Enabled = False
    CmdIncDisp.Enabled = False
    GrdNfDisp.Enabled = False
    GrdNfSel.Enabled = False
    ActDtini.Text = ""
    ActDtfim.Text = ""
    CboRepre.Criterio = ""
    CboRepre.Codigo = 0
    
    ilCodRep = 0
    slrepre = ""
    
    CmdGrp.Enabled = True
    CmdPrd.Enabled = True
    FraFiltro.Enabled = True
    ActDtini.SetFocus
    
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

Private Sub CmdGrp_Click()
    
    slProduto = ""
    slGrupo = ""
    blPrd = False
    
    CmdGrp.BackColor = &HFF00&
    CmdPrd.BackColor = &HFFECEC
    
    blGrp = True

    FraFiltro.Enabled = False
    
    CarregaNotaDisp False
    
    GrdNfDisp.Enabled = True
    GrdNfSel.Enabled = True

End Sub

Private Sub CmdIncDisp_Click()
    
    CarregaNotaDisp True
    
End Sub

Private Sub CmdIncUtil_Click()
    
    CarregaNotaUtil
    
End Sub

Private Sub CmdPrd_Click()
    
    slProduto = ""
    slGrupo = ""
    blGrp = False
    
    CmdPrd.BackColor = &HFF00&
    CmdGrp.BackColor = &HFFECEC
    
    blPrd = True

    FraFiltro.Enabled = False
    
    CarregaNotaDisp False
    
    GrdNfDisp.Enabled = True
    GrdNfSel.Enabled = True

End Sub

Private Sub CmdSair_Click()
    
    Unload Me
    
End Sub

Private Sub CarregaNotaDisp(grd As Boolean)
    
    Dim ExisteNota As Boolean
    
    Select Case grd
    
        Case True
            
            If GrdNfSel.Rows = 1 Then
                Exit Sub
            End If
            
            If GrdNfSel.Row = 0 Then
                GrdNfSel.Row = 1
            End If
            
            'If Trim(GrdNfSel.TextMatrix(GrdNfSel.Row, 7)) = "" Then
                
                GrdNfDisp.AddItem ""
                GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 0) = GrdNfSel.TextMatrix(GrdNfSel.Row, 0)
                GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 1) = GrdNfSel.TextMatrix(GrdNfSel.Row, 1)
                GrdNfDisp.Refresh
                
                Select Case GrdNfSel.Rows
                    
                    Case 2
                        
                        GrdNfSel.Rows = 1
                    
                    Case Else
                        
                        GrdNfSel.RemoveItem (GrdNfSel.Row)
                
                End Select
            
            'End If
            
            If GrdNfSel.Rows = 1 Then
                CmdIncDisp.Enabled = False
                'CmdAtualizar.Enabled = False
            End If
         
            If GrdNfDisp.Rows = 1 Then
                CmdIncUtil.Enabled = False
            Else
                CmdIncUtil.Enabled = True
            End If
            
        Case False
            
            With GrdNfDisp
            
                If blPrd = True Then
                    sgQuery = "select codprd as codigo, dscprd as nome from produto"
                    sgQuery = sgQuery & "  Where FlgSitu = 'N'"
                    sgQuery = sgQuery & "    order by 1 "
                Else
                    sgQuery = "select distinct a.idegrp as codigo, a.nomgrp as nome from grupo_produto a, produto b"
                    sgQuery = sgQuery & "  Where a.idegrp = b.idegrp"
                    sgQuery = sgQuery & "    and b.FlgSitu = 'N'"
                    sgQuery = sgQuery & "    order by 1 "
                End If
                
                Consulta sgQuery
                
                .FormatString = "Código   |     Descrição                                                                                             ||"
                
                While Not Rs.EOF
                
                    ExisteNota = False
                    
                    If ExisteNota = False Then
                        GrdNfDisp.AddItem ""
                        GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 0) = Rs("codigo")
                        GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 1) = Trim(Rs("nome"))
                    End If
                    
                    Rs.MoveNext
                
                Wend
                
                If GrdNfDisp.Rows > 1 Then
                    CmdIncUtil.Enabled = True
                End If
                
            End With
            
    End Select
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{Tab}"
    End If
    
End Sub

Private Sub Form_Load()
    
    Top = 0
    Left = 0
    Width = 13935
    Height = 8010
    
    Me.Show

    slProduto = ""
    slGrupo = ""
    slrepre = ""
    ilCodRep = 0
    blPrd = False
    blGrp = False
    
    Set CboRepre.Conexao = Conexao
    
    CmdIncUtil.Enabled = False
    CmdIncDisp.Enabled = False
    GrdNfDisp.Enabled = False
    GrdNfSel.Enabled = False
    
End Sub

Private Sub CarregaNotaUtil()
    
    If GrdNfDisp.Rows = 1 Then
        Exit Sub
    End If
    
    If GrdNfDisp.Row = 0 Then
        GrdNfDisp.Row = 1
    End If
    
    With GrdNfSel
        .AddItem ""
        .TextMatrix(GrdNfSel.Rows - 1, 0) = GrdNfDisp.TextMatrix(GrdNfDisp.Row, 0)
        .TextMatrix(GrdNfSel.Rows - 1, 1) = GrdNfDisp.TextMatrix(GrdNfDisp.Row, 1)
        .Refresh
    End With
    
    Select Case GrdNfDisp.Rows
        
        Case 2
        
            GrdNfDisp.Rows = 1
        
        Case Else
        
            GrdNfDisp.RemoveItem (GrdNfDisp.Row)
            
    End Select
    
    If GrdNfDisp.Rows = 1 Then
        CmdIncUtil.Enabled = False
    End If
    
    CmdIncDisp.Enabled = True
    
End Sub
