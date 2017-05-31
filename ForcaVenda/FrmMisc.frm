VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{703944EE-9203-11D2-8865-AD1268A0A52F}#1.0#0"; "ActiveCal.OCX"
Begin VB.Form FrmMisc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indicadores"
   ClientHeight    =   9030
   ClientLeft      =   1935
   ClientTop       =   2580
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   14055
   Begin MSChart20Lib.MSChart GrfPedidos 
      Height          =   5055
      Left            =   120
      OleObjectBlob   =   "FrmMisc.frx":0000
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3000
      Width           =   13695
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
      Left            =   12840
      Picture         =   "FrmMisc.frx":2100
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8160
      Width           =   975
   End
   Begin rdActiveCal.ActiveCalendar ActiveCalendar1 
      DragIcon        =   "FrmMisc.frx":2542
      Height          =   2895
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5106
      Date            =   39289
      BorderStyle     =   0
   End
   Begin VB.Label LblDatEnv 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   2280
      TabIndex        =   31
      Top             =   8400
      Width           =   10335
   End
   Begin VB.Label LblQtdVencer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   11040
      TabIndex        =   30
      Top             =   600
      Width           =   615
   End
   Begin VB.Label LblQtdJuros 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   11040
      TabIndex        =   29
      Top             =   960
      Width           =   615
   End
   Begin VB.Label LblQtdVenc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   11040
      TabIndex        =   28
      Top             =   240
      Width           =   615
   End
   Begin VB.Label LblQtdFatu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   6960
      TabIndex        =   27
      Top             =   600
      Width           =   615
   End
   Begin VB.Label LblQtdPed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   6960
      TabIndex        =   26
      Top             =   240
      Width           =   615
   End
   Begin VB.Label LblQtdCart 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   6960
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LblPedaLib 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   11160
      TabIndex        =   24
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label LblPedaTrans 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   11160
      TabIndex        =   23
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label LblTotVenc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   9600
      TabIndex        =   20
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label LblPedfatu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5520
      TabIndex        =   19
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label LblQtdaTrans 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   12720
      TabIndex        =   18
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label LblQtdaLib 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   12720
      TabIndex        =   17
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label LblTotJuros 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   9600
      TabIndex        =   16
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label LblTotVencer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   9600
      TabIndex        =   15
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ped. a Transmitir "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Index           =   3
      Left            =   9600
      TabIndex        =   14
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pedidos a Liberar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   9600
      TabIndex        =   13
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Juros em Aberto "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Títulos a Vencer  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titutos Vencidos "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label LblCliNovo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label LblTotCli 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label LblPedCart 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LblTotPed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clientes Novos       "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Clientes          "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pedidos Faturados - Mês"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pedidos Carteira   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total de Pedidos     - Mês          "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "FrmMisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtoSair_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    BtoSair.SetFocus
    
End Sub

Private Sub Form_Load()
   
    Dim reg As Integer
    Dim i As Integer
    Dim Mes As String

    AjustaJanela Me, 14145, 9510, 700, 10
   
    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    
    Set ts = fso.OpenTextFile("c:\forca\log.txt", ForWriting, True)  'abre um arquivo para escrita , se não existir cria
    
    'sgQuery = "BACKUP LOG unocann WITH TRUNCATE_ONLY"
    'Conexao.Execute sgQuery

    'sgQuery = "DBCC SHRINKFILE (N'UNOCANN_log' , 0, TRUNCATEONLY)"
    'Conexao.Execute sgQuery
   
    'Total de Pedidos
    sgQuery = "select sum(vlrIte) - sum(distinct vlrsimples) as VlrIte, count(distinct b.NroPed) as QtdPed  from item_pedido a, Pedido b "
    sgQuery = sgQuery & "Where a.NroPed = b.NroPed "
    sgQuery = sgQuery & " and b.DatPed between convert(datetime, '" & data1 & " 00:00:00',103) "
    sgQuery = sgQuery & " and                 convert(datetime, '" & Format(datah, "dd/mm/yyyy") & " 23:59:59',103) "
    sgQuery = sgQuery & " and b.SitPed = 'N'"
    sgQuery = sgQuery & " and b.codrep = " & sgRepresentante
    
    'sgQuery = "select sum(vlrite) as vlrite, sum(qtdped) as qtdped from"
    'sgQuery = sgQuery & " (select sum(vlrIte) - sum(distinct vlrsimples) as VlrIte, count(distinct b.NroPed) as QtdPed  from item_pedido a, Pedido b "
    'sgQuery = sgQuery & "Where a.NroPed = b.NroPed "
    'sgQuery = sgQuery & "  and b.DatPed between convert(datetime, '" & data1 & "',103) "
    'sgQuery = sgQuery & "  and                 convert(datetime, '" & datah & "',103) "
    'sgQuery = sgQuery & "  and b.SitPed = 'N'"
    'sgQuery = sgQuery & "  and b.codrep = " & sgRepresentante
    'sgQuery = sgQuery & " Union All"
    'sgQuery = sgQuery & " select sum(vlrIte) - sum(distinct b.vlrsimples) as VlrIte, count(distinct b.NroPedsdo) as QtdPed  from item_pedido_saldo a, Pedido_saldo b, pedido c "
    'sgQuery = sgQuery & "Where a.NroPed = b.NroPed "
    'sgQuery = sgQuery & "  and a.NroPedsdo = b.NroPedsdo "
    'sgQuery = sgQuery & "  and b.DatPed between convert(datetime, '" & data1 & "',103) "
    'sgQuery = sgQuery & "  and                 convert(datetime, '" & datah & "',103) "
    'sgQuery = sgQuery & "  and b.SitPed = 'N'"
    'sgQuery = sgQuery & "  and b.NroPed = c.NroPed "
    'sgQuery = sgQuery & "  and c.codrep = " & sgRepresentante & ") a"
        
    Consulta sgQuery
    
    If Not Rs.EOF Then
        LblTotPed.Caption = Format(IIf(IsNull(Rs!VlrIte), 0, Trim(Rs!VlrIte)), "###,###,###,##0.00")
        LblQtdPed.Caption = Format(IIf(IsNull(Rs!QtdPed), 0, Trim(Rs!QtdPed)), "#,#00")
    End If
    
    Rs.Close
    
    Set Rs = Nothing
   
    'Total de Pedidos em Carteira
    sgQuery = "select sum(vlrIte) as VlrIte, count(distinct b.NroPed) as QtdPed  from item_pedido a, Pedido b "
    sgQuery = sgQuery & "Where a.NroPed = b.NroPed "
    sgQuery = sgQuery & " and b.SitPed = 'N'"
    sgQuery = sgQuery & " and b.codrep = " & sgRepresentante
    sgQuery = sgQuery & " and b.DatEnv is not null"
    sgQuery = sgQuery & " and b.DatEmiNot is null"
    
    Consulta sgQuery
    
    If Not Rs.EOF Then
        LblPedCart.Caption = Format(IIf(IsNull(Rs!VlrIte), 0, Trim(Rs!VlrIte)), "###,###,###,##0.00")
        LblQtdCart.Caption = Format(IIf(IsNull(Rs!QtdPed), 0, Trim(Rs!QtdPed)), "##,#00")
    End If
    
    Rs.Close
    
    Set Rs = Nothing
   
    'Total de Pedidos Faturados no mês
    sgQuery = "select sum(valnot) as valnot, count(*) as qtdped from"
    sgQuery = sgQuery & " (select min(valnot) as valnot, min(nronot) as nronot from"
    sgQuery = sgQuery & "             (select  a.nronot, a.valnot"
    sgQuery = sgQuery & "               from pedido a, cliente b, representante c"
    sgQuery = sgQuery & "     where a.dateminot between convert(datetime, '" & data1 & " 00:00:00',103)"
    sgQuery = sgQuery & "                           and convert(datetime, '" & Format(datah, "dd/mm/yyyy") & " 23:59:59',103) "
    sgQuery = sgQuery & "               and a.sitped = 'N'"
    sgQuery = sgQuery & "               and a.nronot is not null"
    sgQuery = sgQuery & "               and a.codrep = c.codrep"
    sgQuery = sgQuery & "               and a.codcli = b.codcli"
    sgQuery = sgQuery & "             Union All"
    sgQuery = sgQuery & "             select b.nronot, b.valnot"
    sgQuery = sgQuery & "               from pedido a, pedido_saldo b, representante c, cliente d"
    sgQuery = sgQuery & "             Where a.NroPed = b.NroPed"
    sgQuery = sgQuery & "               and b.dateminot between convert(datetime, '" & data1 & " 00:00:00',103)"
    sgQuery = sgQuery & "                                   and convert(datetime, '" & Format(datah, "dd/mm/yyyy") & " 23:59:59',103) "
    sgQuery = sgQuery & "               and b.sitped = 'N'"
    sgQuery = sgQuery & "               and b.nronot is not null"
    sgQuery = sgQuery & "               and a.codrep = c.codrep"
    sgQuery = sgQuery & "               and a.codcli = d.codcli"
    sgQuery = sgQuery & "               and a.codcli = d.codcli) a"
    sgQuery = sgQuery & "             group by nronot) b"
    
    Consulta sgQuery
       
    If Not Rs.EOF Then
        LblPedfatu.Caption = Format(IIf(IsNull(Rs!valnot), 0, Trim(Rs!valnot)), "###,###,###,##0.00")
        LblQtdFatu.Caption = Format(IIf(IsNull(Rs!QtdPed), 0, Trim(Rs!QtdPed)), "##,#00")
    End If
    
    Rs.Close
    
    Set Rs = Nothing
   
    'Total de clientes
    sgQuery = "select  count(*) as QtdCli  from cliente "
    sgQuery = sgQuery & "Where codrep = " & sgRepresentante
    
    Consulta sgQuery
    
    If Not Rs.EOF Then
        LblTotCli.Caption = Format(IIf(IsNull(Rs!QtdCli), 0, Trim(Rs!QtdCli)), "##,#00")
    End If
    
    Rs.Close
    
    Set Rs = Nothing
   
    'Total de clientes Novos
    sgQuery = "select  count(*) as QtdCli  from cliente "
    sgQuery = sgQuery & "Where DatPriComp between convert(datetime, '" & data1 & " 00:00:00',103) "
    sgQuery = sgQuery & "  and                    convert(datetime, '" & Format(datah, "dd/mm/yyyy") & " 23:59:59',103) "
    sgQuery = sgQuery & "  and codrep = " & sgRepresentante
    
    Consulta sgQuery
    
    If Not Rs.EOF Then
        LblCliNovo.Caption = Format(IIf(IsNull(Rs!QtdCli), 0, Trim(Rs!QtdCli)), "##,#00")
    End If
    
    Rs.Close
    
    Set Rs = Nothing
   
    'Total de Titulos Vencidos
    sgQuery = " select count(*) as QtdDup,  sum(vlrdup) as VlrDup from DUPLICATA a, cliente b"
    sgQuery = sgQuery & "    Where datpag Is Null And datven < (getdate() - 1)"
    sgQuery = sgQuery & "      AND a.CodCli = b.codcli"
    sgQuery = sgQuery & "      AND b.codrep = " & sgRepresentante
    
    Consulta sgQuery
    
    If Not Rs.EOF Then
        LblTotVenc.Caption = Format(IIf(IsNull(Rs!VlrDup), 0, Trim(Rs!VlrDup)), "###,###,###,##0.00")
        LblQtdVenc.Caption = Format(IIf(IsNull(Rs!QtdDup), 0, Trim(Rs!QtdDup)), "##,#00")
    End If
    
    Rs.Close
    
    Set Rs = Nothing
   
    'Total de Titulos a Vencer
    sgQuery = " select count(*) as QtdDup,  sum(vlrdup) as VlrDup from DUPLICATA a, cliente b"
    sgQuery = sgQuery & "    Where datpag is null and datven >= (getdate() - 1)"
    sgQuery = sgQuery & "      AND a.CodCli = b.codcli"
    sgQuery = sgQuery & "      AND b.codrep = " & sgRepresentante
    
    Consulta sgQuery
    
    If Not Rs.EOF Then
        LblTotVencer.Caption = Format(IIf(IsNull(Rs!VlrDup), 0, Trim(Rs!VlrDup)), "###,###,###,##0.00")
        LblQtdVencer.Caption = Format(IIf(IsNull(Rs!QtdDup), 0, Trim(Rs!QtdDup)), "##,#00")
    End If
    
    Rs.Close
    
    Set Rs = Nothing
   
    'Total de Juros em aberto
    sgQuery = " select count(*) as QtdJur,  sum(JurDev) as JurDev from DUPLICATA a, cliente b"
    sgQuery = sgQuery & "    Where datjur is null and JurDev > 0"
    sgQuery = sgQuery & "      AND a.CodCli = b.codcli"
    sgQuery = sgQuery & "      AND b.codrep = " & sgRepresentante
    
    Consulta sgQuery
    
    If Not Rs.EOF Then
        LblTotJuros.Caption = Format(IIf(IsNull(Rs!JurDev), 0, Trim(Rs!JurDev)), "###,###,###,##0.00")
        LblQtdJuros.Caption = Format(IIf(IsNull(Rs!QtdJur), 0, Trim(Rs!QtdJur)), "##,#00")
    End If
    
    Rs.Close
    
    Set Rs = Nothing
   
    'Total de Pedidos a Liberar
    sgQuery = "select sum(vlrIte) - sum(distinct vlrsimples) as VlrIte, count(distinct b.NroPed) as QtdPed  from item_pedido a, Pedido b "
    sgQuery = sgQuery & "Where a.NroPed = b.NroPed "
    sgQuery = sgQuery & " and b.SitPed = 'N'"
    sgQuery = sgQuery & " and b.codrep = " & sgRepresentante
    sgQuery = sgQuery & " and b.DatLib is null"
    
    Consulta sgQuery
    
    If Not Rs.EOF Then
        LblPedaLib.Caption = Format(IIf(IsNull(Rs!VlrIte), 0, Trim(Rs!VlrIte)), "###,###,###,##0.00")
        LblQtdaLib.Caption = Format(IIf(IsNull(Rs!QtdPed), 0, Trim(Rs!QtdPed)), "##,#00")
    End If
    
    Rs.Close
    
    Set Rs = Nothing
   
    'Total de Pedidos a Transmitir
    sgQuery = "select sum(vlrIte) - sum(distinct vlrsimples) as VlrIte, count(distinct b.NroPed) as QtdPed  from item_pedido a, Pedido b "
    sgQuery = sgQuery & "Where a.NroPed = b.NroPed "
    sgQuery = sgQuery & " and b.SitPed = 'N'"
    sgQuery = sgQuery & " and b.codrep = " & sgRepresentante
    sgQuery = sgQuery & " and b.DatEnv is null"
    
    Consulta sgQuery
    
    If Not Rs.EOF Then
        LblPedaTrans.Caption = Format(IIf(IsNull(Rs!VlrIte), 0, Trim(Rs!VlrIte)), "###,###,###,##0.00")
        LblQtdaTrans.Caption = Format(IIf(IsNull(Rs!QtdPed), 0, Trim(Rs!QtdPed)), "##,#00")
    End If
    
    Rs.Close
    
    Set Rs = Nothing
   
   
    'Gráfico Pedidos
    'Total de Pedidos
    'sgQuery = "select sum(vlrIte) - sum(distinct vlrsimples) as VlrIte, month(b.DatPed) as Mes  from item_pedido a, Pedido b "
    'sgQuery = sgQuery & "Where a.NroPed = b.NroPed "
    'sgQuery = sgQuery & " and b.SitPed = 'N'"
    'sgQuery = sgQuery & " and year(b.DatPed) = " & Year(datah) & " and  datped <= convert(datetime, '" & datah & "',103) "
    'sgQuery = sgQuery & " and b.codrep = " & sgRepresentante
    'sgQuery = sgQuery & " group by Month(DatPed) "
    'sgQuery = sgQuery & " order by 2 "

    sgQuery = "select sum(c.vlrite) as vlrite, c.mes, sum(c.valfat) as valfat from"
    sgQuery = sgQuery & " (select sum(vlrIte) - sum(distinct vlrsimples) as VlrIte,"
    sgQuery = sgQuery & "        month(b.DatPed) as Mes, 0 as valfat  from item_pedido a, Pedido b"
    sgQuery = sgQuery & "    Where a.NroPed = b.NroPed"
    sgQuery = sgQuery & "      and b.SitPed = 'N'"
    sgQuery = sgQuery & "      and year(b.DatPed) = " & Year(datah) & " and  datped <= convert(datetime, '" & Format(datah, "dd/mm/yyyy") & "',103)"
    sgQuery = sgQuery & "      and b.codrep = " & sgRepresentante
    sgQuery = sgQuery & "      group by Month(DatPed)"
    sgQuery = sgQuery & "   Union All"
    sgQuery = sgQuery & "   select 0 as vlrite, mes,sum(valnot) as valfat from"
    sgQuery = sgQuery & "      (select sum(valnot) as valnot, month(DatEmiNot) as Mes  from Pedido"
    sgQuery = sgQuery & "      Where year(DatEmiNot) = " & Year(datah) & " and DatEmiNot <= convert(datetime, '" & Format(datah, "dd/mm/yyyy") & "',103)"
    sgQuery = sgQuery & "       and SitPed = 'N'"
    sgQuery = sgQuery & "       and codrep = " & sgRepresentante
    sgQuery = sgQuery & "      group by month(DatEmiNot)"
    sgQuery = sgQuery & "      Union All"
    sgQuery = sgQuery & "      select sum(a.valnot) as valnot, month(a.DatEmiNot) as Mes  from pedido_saldo a, Pedido b"
    sgQuery = sgQuery & "      Where a.NroPed = b.NroPed"
    sgQuery = sgQuery & "        and year(a.DatEmiNot) = " & Year(datah) & " and a.DatEmiNot <= convert(datetime, '" & Format(datah, "dd/mm/yyyy") & "',103)"
    sgQuery = sgQuery & "        and a.SitPed = 'N'"
    sgQuery = sgQuery & "        and b.codrep = " & sgRepresentante
    sgQuery = sgQuery & "       group by month(a.DatEmiNot) ) a"
    sgQuery = sgQuery & "       group by mes) c"
    sgQuery = sgQuery & " group by c.mes"
    
    Consulta sgQuery

    ts.Write "Executei Gráfico Pedidos: " & sgQuery
    ts.Write "\n"
    'ts.Close
    
    reg = Rs.RecordCount
    
    'GrfPedidos.chartType = 2 'barra em duas dimensões
    GrfPedidos.ShowLegend = True 'não mostra legenda
    GrfPedidos.Title = "Indicador de Pedidos - Anual" 'titulo do gráfico
    GrfPedidos.ColumnCount = 2 'uma série
    GrfPedidos.RowCount = reg 'número sequencia de dados
    GrfPedidos.Visible = True
    
    While Not Rs.EOF()
    
        For i = 1 To reg
        
            GrfPedidos.row = i
            
            Select Case Rs("Mes")
                
                Case 1
                    
                    Mes = "Janeiro"
                
                Case 2
                    
                    Mes = "Fevereiro"
                    
                Case 3
                
                    Mes = "Março"
                    
                Case 4
                
                    Mes = "Abril"
                    
                Case 5
                    
                    Mes = "Maio"
                    
                Case 6
                
                    Mes = "Junho"
                
                Case 7
                
                    Mes = "Julho"
                
                Case 8
                
                    Mes = "Agosto"
                    
                Case 9
                    
                    Mes = "Setembro"
                Case 10
                
                    Mes = "Outubro"
                    
                Case 11
                
                    Mes = "Novembro"
                
                Case 12
                    
                    Mes = "Dezembro"
                    
            End Select
         
            GrfPedidos.RowLabel = Mes
            GrfPedidos.Column = 1
            GrfPedidos.ColumnLabel = "Pedidos Emitidos"
            GrfPedidos.Data = Format(IIf(IsNull(Rs!VlrIte), 0, Trim(Rs!VlrIte) / 1), "###,###,###,##0.00")
            GrfPedidos.Column = 2
            GrfPedidos.ColumnLabel = "Faturamento"
            GrfPedidos.Data = Format(IIf(IsNull(Rs!Valfat), 0, Trim(Rs!Valfat) / 1), "###,###,###,##0.00")
            
            Rs.MoveNext
            
        Next
        
    Wend
   
    'Data da ultima transmissão
    sgQuery = "select max(DatEnv) as dataenv from Pedido "
    
    Consulta sgQuery
    
    If Not Rs.EOF Then
    
        If Trim(Rs!dataenv) <> "" Then
            LblDatEnv.Caption = "Última transmissão de arquivos foi realizada no dia " & Format(Rs!dataenv, "dd/mm/yyyy hh:mm")
        End If
        
    End If
    
    Rs.Close
    
    Set Rs = Nothing
    
End Sub
