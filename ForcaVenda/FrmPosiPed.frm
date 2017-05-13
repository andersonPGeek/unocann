VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{368CC970-FF03-11D7-9B5A-000B6A03449D}#1.1#0"; "Combo_DB.ocx"
Object = "{F454059D-91FE-11D2-8865-AD1268A0A52F}#2.0#0"; "ActiveDate.ocx"
Begin VB.Form FrmPosiPed 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Posição de Pedidos"
   ClientHeight    =   10065
   ClientLeft      =   -210
   ClientTop       =   1515
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin Project_Combo_DB.Combo_DB CboRemet 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   661
      Cols            =   0
      Cabecalho       =   -1  'True
   End
   Begin VB.CommandButton BtoGrava 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      MaskColor       =   &H00FF0000&
      Picture         =   "FrmPosiPed.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   975
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
      Height          =   975
      Left            =   13680
      Picture         =   "FrmPosiPed.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid GrdPedido 
      Height          =   6975
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   12303
      _Version        =   393216
      Rows            =   1
      Cols            =   12
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   192
      GridColorFixed  =   128
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   9375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   930
   End
End
Attribute VB_Name = "FrmPosiPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim slremet As String
Dim ilCodCli As Integer
Dim blLoad As Boolean


Function CompoeGridPed() As Boolean
Dim blI As Integer
On Error GoTo TratarErro

sgQuery = "select a.nroped, a.datped, a.codcli,c.NomCli, sum(b.vlrite) as Valor, d.DscCnd,"
sgQuery = sgQuery & "       a.NomTra , a.sitPed, a.datlib, a.datenv, TipNot, NroNot, dateminot"
sgQuery = sgQuery & "  from Pedido a, Item_pedido b, Cliente c, Condicao d"
sgQuery = sgQuery & "  Where a.nroped = b.nroped"
sgQuery = sgQuery & "    and a.codcli = c.codcli"
sgQuery = sgQuery & "    and a.codcnd = d.codcnd"
If APLICA = 1 Then
   sgQuery = sgQuery & "    and a.codrep = " & sgRepresentante
End If
If Trim(ActDtini.Text) <> "" Then
   sgQuery = sgQuery & "    and a.datped between convert(datetime,'" & Trim(ActDtini.Text) & "',103)"
   sgQuery = sgQuery & "                     and convert(datetime,'" & Trim(ActDtfim.Text) & "',103)"
End If
If ilCodCli <> 0 Then
   sgQuery = sgQuery & "    and a.codcli = " & ilCodCli
End If
sgQuery = sgQuery & "    group by a.nroped, a.datped, a.codcli, c.NomCli, d.DscCnd,"
sgQuery = sgQuery & "             a.NomTra, a.sitPed, a.datlib, a.datenv,"
sgQuery = sgQuery & "             a.TipNot , a.NroNot, a.dateminot"
sgQuery = sgQuery & "    order by 1 desc"

consulta sgQuery
If blLoad = False And Rs.RecordCount > 100 Then
   MsgBox "Resultado retornou mais de 100 linhas." & Chr(13) & "Favor refazer seu filtro de pesquisa.", vbExclamation + vbOKOnly, "Atenção!"
End If

GrdPedido.Rows = 1
blI = 1
Do While Not Rs.EOF
   If blI > 100 Then
      Exit Do
   End If
   
   GrdPedido.Rows = GrdPedido.Rows + 1
   GrdPedido.TextMatrix(blI, 0) = Format(Trim(Rs!NroPed), "000000")
   GrdPedido.TextMatrix(blI, 1) = Format(Trim(Rs!DatPed), "dd/mm/yyyy")
   GrdPedido.TextMatrix(blI, 2) = Format(Trim(Rs!DatPed), "00000")
   GrdPedido.TextMatrix(blI, 3) = Rs!NomCli
   GrdPedido.TextMatrix(blI, 4) = Format(Rs!Valor, "##,###,##0.00")
   GrdPedido.TextMatrix(blI, 5) = Rs!DscCnd
   GrdPedido.TextMatrix(blI, 6) = Rs!NomTra
   GrdPedido.TextMatrix(blI, 7) = Rs!SitPed
   GrdPedido.TextMatrix(blI, 8) = IIf(IsNull(Rs!Datlib), "", Format(Trim(Rs!Datlib), "dd/mm/yyyy"))
   GrdPedido.TextMatrix(blI, 9) = IIf(IsNull(Rs!DatEnv), "", Format(Trim(Rs!DatEnv), "dd/mm/yyyy"))
   GrdPedido.TextMatrix(blI, 10) = IIf(IsNull(Rs!NroNot), "", Rs!TipNot & Format(Trim(Rs!NroNot), "000000"))
   GrdPedido.TextMatrix(blI, 11) = IIf(IsNull(Rs!DatEmiNot), "", Format(Trim(Rs!DatEmiNot), "dd/mm/yyyy"))
   GrdPedido.Row = blI
'
   If IsNull(Rs!Datlib) And IsNull(Rs!Datlib) And IsNull(Rs!DatEmiNot) Then
      GrdPedido.Col = 0
      GrdPedido.CellBackColor = &HFF&
      GrdPedido.CellForeColor = &HFFFF&
'      GrdPedido.Col = 8
'      GrdPedido.CellBackColor = vbRed
'      GrdPedido.Col = 9
'      GrdPedido.CellBackColor = vbRed
'      GrdPedido.Col = 10
'      GrdPedido.CellBackColor = vbRed
'      GrdPedido.Col = 11
'      GrdPedido.CellBackColor = vbRed
   End If
   If Not IsNull(Rs!DatEnv) And IsNull(Rs!DatEmiNot) Then
      GrdPedido.Col = 0
      GrdPedido.CellBackColor = vbYellow
'      GrdPedido.Col = 8
'      GrdPedido.CellBackColor = vbYellow
'      GrdPedido.Col = 9
'      GrdPedido.CellBackColor = vbYellow
'      GrdPedido.Col = 10
'      GrdPedido.CellBackColor = vbYellow
'      GrdPedido.Col = 11
'      GrdPedido.CellBackColor = vbYellow
   End If
   If Not IsNull(Rs!DatEmiNot) Then
      GrdPedido.Col = 0
      GrdPedido.CellBackColor = &HFF00&
'      GrdPedido.Col = 8
'      GrdPedido.CellBackColor = &HC000&
'      GrdPedido.Col = 9
'      GrdPedido.CellBackColor = &HC000&
'      GrdPedido.Col = 10
'      GrdPedido.CellBackColor = &HC000&
'      GrdPedido.Col = 11
'      GrdPedido.CellBackColor = &HC000&
   End If
   If Trim(Rs!SitPed) = "C" Then
      GrdPedido.Col = 0
'      GrdPedido.CellBackColor = &HFFFFFF
      GrdPedido.CellBackColor = &H80000012
      GrdPedido.CellForeColor = &HFFFF&
   End If
   blI = blI + 1
   Rs.MoveNext
Loop

Rs.Close
blLoad = False

Set Rs = Nothing
DoEvents

Exit Function

TratarErro:
Rotina_Erro "CompoeGridPed"

End Function

Private Sub BtoGrava_Click()
 
If ActDtini.Text = "" And ActDtini.Text = "" And ilCodCli = 0 Then
   MsgBox "Informe dados para o filtro", vbInformation
   ActDtini.SetFocus
   Exit Sub
End If
   
If ActDtini.Text <> "" Or ActDtini.Text <> "" Then
   If CDate(ActDtini.Text) > CDate(ActDtfim.Text) Or _
      Year(CDate(ActDtini.Text)) < 1950 Or _
      Year(CDate(ActDtfim.Text)) < 1950 Then
      MsgBox "Intervalo de datas inválido", vbInformation
      ActDtini.SetFocus
      Exit Sub
   End If
End If

CompoeGridPed

End Sub

Private Sub BtoSair_Click()
  Unload Me
  
End Sub

Private Sub CboRemet_Consultar()
   slremet = ""
   CboRemet.query = "Select NomCli As Cliente, CodCli As Código, CgcCli as CNPJ From Cliente " & _
                                "Where " & IIf(IsNumeric(CboRemet.Criterio), "CodCli", "NomCli") & " Like '" & CboRemet.Criterio & "%' and CodRep = " & Trim(sgRepresentante) & " order by " & IIf(IsNumeric(CboRemet.Criterio), "CodCli", "NomCli")
End Sub

Private Sub CboRemet_GotFocus()
  Call SelecionaTudo

End Sub

Private Sub CboRemet_LostFocus()
    If CboRemet.Criterio <> "" Then
       slremet = CboRemet.Criterio
       ilCodCli = CboRemet.Codigo
    Else
       ilCodCli = 0
       slremet = ""
    End If
End Sub

Private Sub Form_Activate()
  blLoad = True
  CompoeGridPed
End Sub

Private Sub Form_Load()

Me.Left = 0
Me.Top = 0
Me.Height = 10750
Me.Width = 15360

Set CboRemet.Conexao = Conexao

ilCodCli = 0
slremet = ""
blLoad = True
bgBloqPed = False
bgConsultaPed = False

GrdPedido.TextMatrix(0, 0) = "Pedido"
GrdPedido.ColWidth(0) = 700
GrdPedido.TextMatrix(0, 1) = "Dt.Emissão"
GrdPedido.ColWidth(1) = 1000
GrdPedido.TextMatrix(0, 2) = "Cod.Cli"
GrdPedido.ColWidth(2) = 600
GrdPedido.TextMatrix(0, 3) = "Cliente"
GrdPedido.ColWidth(3) = 3700
GrdPedido.TextMatrix(0, 4) = "Val.Pedido"
GrdPedido.ColWidth(4) = 1000
GrdPedido.TextMatrix(0, 5) = "Cond.Pagto"
GrdPedido.ColWidth(5) = 1500
GrdPedido.TextMatrix(0, 6) = "Transp."
GrdPedido.ColWidth(6) = 2100
GrdPedido.TextMatrix(0, 7) = "Sit."
GrdPedido.ColWidth(7) = 300
GrdPedido.TextMatrix(0, 8) = "Dt.Liberação"
GrdPedido.ColWidth(8) = 1000
GrdPedido.TextMatrix(0, 9) = "Dt. Envio"
GrdPedido.ColWidth(9) = 1000
GrdPedido.TextMatrix(0, 10) = "N.Fiscal"
GrdPedido.ColWidth(10) = 1000
GrdPedido.TextMatrix(0, 11) = "Dt.Fatur."
GrdPedido.ColWidth(11) = 1000

'CompoeGridPed

End Sub

Private Sub GrdPedido_DblClick()
  bgBloqPed = False
  If GrdPedido.RowSel = 0 Then
     Exit Sub
  End If
  bgConsultaPed = True
  Me.Enabled = False
  igNroPed = GrdPedido.TextMatrix(GrdPedido.RowSel, 0)
  If Trim(GrdPedido.TextMatrix(GrdPedido.RowSel, 9)) <> "" Or GrdPedido.TextMatrix(GrdPedido.RowSel, 11) <> "" Then
     bgBloqPed = True
  End If
  
  FrmConhecimento.Show
End Sub
