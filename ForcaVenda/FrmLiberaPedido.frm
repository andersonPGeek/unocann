VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLiberaPedido 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Liberação de Pedidos para Envio à UNOCANN"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   2250
   ClientWidth     =   14865
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
   ScaleWidth      =   14865
   Begin VB.CommandButton CmdIncDispT 
      BackColor       =   &H000000FF&
      Caption         =   "<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6135
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Retira todas as notas selecionadas"
      Top             =   5385
      Width           =   525
   End
   Begin VB.CommandButton CmdIncUtilT 
      BackColor       =   &H0000FF00&
      Caption         =   ">>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6135
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "seleciona todas as notas"
      Top             =   645
      Width           =   525
   End
   Begin VB.CommandButton CmdIncUtil 
      BackColor       =   &H0000FF00&
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6135
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Seleciona uma nota"
      Top             =   1065
      Width           =   525
   End
   Begin VB.CommandButton CmdIncDisp 
      BackColor       =   &H000000FF&
      Caption         =   "<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6135
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Retira uma nota selecionada"
      Top             =   5835
      Width           =   525
   End
   Begin VB.CommandButton CmdAtualizar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Atualizar"
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
      Left            =   12120
      Picture         =   "FrmLiberaPedido.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1095
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
      Left            =   13560
      Picture         =   "FrmLiberaPedido.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid GrdNfDisp 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   645
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   9763
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      BackColor       =   255
      ForeColor       =   -2147483639
      BackColorFixed  =   12640511
      BackColorSel    =   -2147483642
      BackColorBkg    =   255
      GridColorFixed  =   192
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   "Pedido   | Cliente                                                                      |Valor         | |"
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
      Left            =   6720
      TabIndex        =   7
      Top             =   645
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   9790
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      BackColor       =   49152
      ForeColor       =   -2147483639
      BackColorFixed  =   12648384
      BackColorBkg    =   49152
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   $"FrmLiberaPedido.frx":0884
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
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pedidos Liberados para Envio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   6870
      TabIndex        =   8
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " Pedidos Não Liberados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "FrmLiberaPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ilcont As Integer

Private Sub CmdAtualizar_Click()

    On Error GoTo TrataErro

    Conexao.BeginTrans
   
    'Pedidos não liberados
    
    For ilcont = 1 To GrdNfDisp.Rows - 1
        
        sgQuery = "update Pedido set datlib = null where NroPed = '" & Trim(GrdNfDisp.TextMatrix(ilcont, 0)) & "' and codrep = " & sgRepresentante
        
        Set Rs = Conexao.Execute(sgQuery)
        Set Rs = Nothing
        
    Next ilcont
      
    'Pedidos liberados
    
    For ilcont = 1 To GrdNfSel.Rows - 1
        
        sgQuery = "update Pedido set datlib = convert(datetime,'" & Trim(GrdNfSel.TextMatrix(ilcont, 3)) & "', 103) where NroPed = '" & Trim(GrdNfSel.TextMatrix(ilcont, 0)) & "' and codrep = " & sgRepresentante
        
        Set Rs = Conexao.Execute(sgQuery)
        Set Rs = Nothing
        
    Next ilcont
   
    Conexao.CommitTrans
   
    MsgBox "Liberação de pedidos para Envio terminada com sucesso", vbExclamation + vbOKOnly, "Atenção!"
   
    Unload Me
   
    Exit Sub

TrataErro:

    Rotina_Erro "CmdAtualizar_Click"
        
End Sub

Private Sub CmdIncDisp_Click()
    
    CarregaNotaDisp True
    
End Sub

Private Sub CmdIncDispT_Click()

    Dim ilDecresce
    
    MousePointer = 11
    ilDecresce = GrdNfSel.Rows - 1
    
    'If GrdNfDisp.Rows = 2 Then
        'GrdNfDisp.Rows = 1
    'End If
    
    While ilDecresce > 0
    
        'If Trim(GrdNfSel.TextMatrix(ilDecresce, 7)) = "" Then
            
            GrdNfDisp.AddItem ""
            GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 0) = GrdNfSel.TextMatrix(ilDecresce, 0)
            GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 1) = GrdNfSel.TextMatrix(ilDecresce, 1)
            GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 2) = GrdNfSel.TextMatrix(ilDecresce, 2)
            GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 3) = GrdNfSel.TextMatrix(ilDecresce, 3)
            
            Select Case GrdNfSel.Rows
                
                Case 2
                    
                    GrdNfSel.Rows = 1
                    
                Case Else
                    
                    GrdNfSel.RemoveItem (ilDecresce)
                    
            End Select
            
            GrdNfSel.Refresh
        
        'End If
        
        ilDecresce = ilDecresce - 1
        
    Wend
    
    MousePointer = 0
    
    If GrdNfDisp.Rows = 1 Then
        CmdIncUtilT.Enabled = False
        CmdIncUtil.Enabled = False
    Else
        CmdIncUtilT.Enabled = True
        CmdIncUtil.Enabled = True
    End If
    
    CmdIncDispT.Enabled = False
    CmdIncDisp.Enabled = False
    'CmdAtualizar.Enabled = False
    
End Sub

Private Sub CmdIncUtil_Click()
    
    CarregaNotaUtil False
    
End Sub

Private Sub CmdIncUtilT_Click()

    MousePointer = 11
    
    'If GrdNfSel.Rows = 2 Then
        'GrdNfSel.Rows = 1
    'End If
    
    While GrdNfDisp.Rows > 1
    
        GrdNfSel.AddItem ""
        GrdNfSel.TextMatrix(GrdNfSel.Rows - 1, 0) = GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 0)
        GrdNfSel.TextMatrix(GrdNfSel.Rows - 1, 1) = GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 1)
        GrdNfSel.TextMatrix(GrdNfSel.Rows - 1, 2) = GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 2)
        
        If Trim(GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 3)) = "" Then
            GrdNfSel.TextMatrix(GrdNfSel.Rows - 1, 3) = Format(CDate(Date) & " " & Time, "dd/mm/yyyy hh:mm:ss")
        Else
            GrdNfSel.TextMatrix(GrdNfSel.Rows - 1, 3) = GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 3)
        End If
        
        GrdNfDisp.Rows = GrdNfDisp.Rows - 1
        
    Wend
    
    GrdNfDisp.Rows = 1
    GrdNfDisp.Refresh
    
    MousePointer = 0
    
    CmdIncUtilT.Enabled = False
    CmdIncUtil.Enabled = False
    CmdIncDispT.Enabled = True
    CmdIncDisp.Enabled = True
    CmdAtualizar.Enabled = True
    
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
                GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 2) = GrdNfSel.TextMatrix(GrdNfSel.Row, 2)
                GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 3) = GrdNfSel.TextMatrix(GrdNfSel.Row, 3)
                GrdNfDisp.Refresh
                    
                Select Case GrdNfSel.Rows
                        
                    Case 2
                            
                        GrdNfSel.Rows = 1
                            
                    Case Else
                            
                        GrdNfSel.RemoveItem (GrdNfSel.Row)
                            
                End Select
                    
            'End If
                
            If GrdNfSel.Rows = 1 Then
                    
                CmdIncDispT.Enabled = False
                CmdIncDisp.Enabled = False
                'CmdAtualizar.Enabled = False
                    
            End If
         
            If GrdNfDisp.Rows = 1 Then
                CmdIncUtilT.Enabled = False
                CmdIncUtil.Enabled = False
            Else
                CmdIncUtilT.Enabled = True
                CmdIncUtil.Enabled = True
            End If
        
        Case False
                
            With GrdNfDisp
                    
                '.Rows = 1
                    
                sgQuery = "select a.nroped, c.NomCli, sum(b.vlrite) as Valor"
                sgQuery = sgQuery & "  from Pedido a, Item_pedido b, Cliente c"
                sgQuery = sgQuery & "  Where a.nroped = b.nroped"
                sgQuery = sgQuery & "    and a.codcli = c.codcli"
                sgQuery = sgQuery & "    and a.DatLib is null"
                sgQuery = sgQuery & "    and (a.FlgAlt is null or a.FlgAlt = 'A' or flgalt = 'F')"
                sgQuery = sgQuery & "    and a.SitPed = 'N'"
                sgQuery = sgQuery & "    and a.codrep = " & sgRepresentante
                sgQuery = sgQuery & "    group by a.nroped,  c.NomCli "
                sgQuery = sgQuery & "    order by 1 "
                    
                Consulta sgQuery
                    
                .FormatString = "Pedido       | Cliente                                                               |Valor                   | ||"
                    
                While Not Rs.EOF
                        
                    ExisteNota = False
                        
                    If ExisteNota = False Then
                            
                        GrdNfDisp.AddItem ""
                        GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 0) = Rs("nroped")
                        GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 1) = Trim(Rs("nomcli"))
                        GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 2) = Format(Rs("Valor"), "##,###,###,##0.00")
                        GrdNfDisp.TextMatrix(GrdNfDisp.Rows - 1, 3) = ""
                            
                    End If
                        
                    Rs.MoveNext
                        
                Wend
                    
                If GrdNfDisp.Rows > 1 Then
                    CmdIncUtilT.Enabled = True
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
    Width = 14985
    Height = 8010
    Me.Show

    CarregaNotaUtil True
    CarregaNotaDisp False

End Sub

Private Sub CarregaNotaUtil(blconsulta As Boolean)

    Select Case blconsulta
    
        Case False
        
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
                .TextMatrix(GrdNfSel.Rows - 1, 2) = GrdNfDisp.TextMatrix(GrdNfDisp.Row, 2)
                    
                If Trim(GrdNfDisp.TextMatrix(GrdNfDisp.Row, 3)) = "" Then
                    .TextMatrix(GrdNfSel.Rows - 1, 3) = Format(CDate(Date) & " " & Time, "dd/mm/yyyy hh:mm:ss")
                Else
                    .TextMatrix(GrdNfSel.Rows - 1, 3) = GrdNfDisp.TextMatrix(GrdNfDisp.Row, 3)
                End If
                
                .Refresh
                    
            End With
                
            Select Case GrdNfDisp.Rows
                    
                Case 2
                        
                    GrdNfDisp.Rows = 1
                    
                Case Else
                        
                    GrdNfDisp.RemoveItem (GrdNfDisp.Row)
                        
            End Select
            
            If GrdNfDisp.Rows = 1 Then
                CmdIncUtilT.Enabled = False
                CmdIncUtil.Enabled = False
            End If
                
            CmdIncDispT.Enabled = True
            CmdIncDisp.Enabled = True
            'CmdAtualizar.Enabled = False
                
        Case True
                
            sgQuery = "select a.nroped, c.NomCli, sum(b.vlrite) as Valor, a.DatLib"
            sgQuery = sgQuery & "  from Pedido a, Item_pedido b, Cliente c"
            sgQuery = sgQuery & "  Where a.nroped = b.nroped"
            sgQuery = sgQuery & "    and a.codcli = c.codcli"
            sgQuery = sgQuery & "    and a.DatLib is not null"
            sgQuery = sgQuery & "    and a.DatEnv is null"
            sgQuery = sgQuery & "    and a.SitPed = 'N'"
            sgQuery = sgQuery & "    and a.codrep = " & sgRepresentante
            sgQuery = sgQuery & "    group by a.nroped,  c.NomCli, a.DatLib"
            sgQuery = sgQuery & "    order by 1 "
                
            Call Consulta(sgQuery)
            
            If Rs.EOF Then
                Exit Sub
            End If
      
            GrdNfSel.Clear
            GrdNfSel.Rows = 1
            GrdNfSel.FormatString = "Pedido       | Cliente                                                                  |Valor                     |        DT. Liberação  |"
                
            While Not Rs.EOF
                    
                GrdNfSel.AddItem ""
                GrdNfSel.TextMatrix(Rs.Bookmark, 0) = Rs("nroped")
                GrdNfSel.TextMatrix(Rs.Bookmark, 1) = Trim(Rs("nomcli"))
                GrdNfSel.TextMatrix(Rs.Bookmark, 2) = Format(Rs("Valor"), "##,###,###,##0.00")
                GrdNfSel.TextMatrix(Rs.Bookmark, 3) = Format(Rs("datlib"), "dd/mm/yyyy hh:mm:ss")
                    
                Rs.MoveNext
                    
            Wend
                
            CmdIncDisp.Enabled = True
            CmdIncDispT.Enabled = True
            CmdAtualizar.Enabled = True
                
    End Select
            
End Sub

