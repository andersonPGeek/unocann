VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAviso 
   Caption         =   "Notificação de Pendências"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   13785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtoSair 
      BackColor       =   &H80000016&
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
      Left            =   12240
      Picture         =   "FrmAviso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8760
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid GrdPedido 
      Height          =   8175
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   14420
      _Version        =   393216
      Rows            =   1
      Cols            =   16
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim slremet As String
Dim ilCodCli As Integer
Dim blLoad As Boolean

Function GridAviso() As Boolean

    On Error GoTo TratarErro
    
    Dim blI As Integer
    
'    If CboRemet.Criterio <> "" Then
'        ilCodCli = CboRemet.Codigo
'    End If
'
'    If MskNroPedido.Texto > 0 Then
'
'        ilCodCli = 0
'
'        CboRemet.Criterio = ""
'
'    End If

    'Pedidos não Liberados
    sgQuery = "select a.nroped, a.datped, a.codcli,c.NomCli, sum(b.vlrite) - sum(distinct a.vlrsimples) as Valor,a.codcnd,a.codrep, d.DscCnd,"
    sgQuery = sgQuery & "       a.NomTra , a.codrep, a.sitPed, a.datlib, a.datenv, a.flgalt"
    sgQuery = sgQuery & "  from Pedido a, Item_pedido b, Cliente c, Condicao d"
    sgQuery = sgQuery & "  Where (a.datlib is null or a.flgalt = 'N')"
    sgQuery = sgQuery & "    and a.nroped = b.nroped"
    sgQuery = sgQuery & "    and a.codcli = c.codcli"
    sgQuery = sgQuery & "    and a.codcnd = d.codcnd"
    
'    If MskNroPedido.Texto > 0 Then
'        sgQuery = sgQuery & "    and a.nroped = " & Trim(MskNroPedido.Texto)
'    End If
'
    If APLICA = 1 Then
        sgQuery = sgQuery & "    and a.codrep = " & sgRepresentante
    End If
    
'    If Trim(ActDtini.Text) <> "" Then
'        sgQuery = sgQuery & "    and a.datped between convert(datetime,'" & Trim(ActDtini.Text) & "',103)"
'        sgQuery = sgQuery & "                     and convert(datetime,'" & Trim(ActDtfim.Text) & "',103)"
'    End If
    
    If ilCodCli <> 0 Then
        sgQuery = sgQuery & "    and a.codcli = " & ilCodCli
    End If
    
'    If ChkNotif.Value = 1 And MskNroPedido.Texto = 0 Then
        sgQuery = sgQuery & " and a.flgalt is not null and a.nronot is null and a.sitped = 'C'"
'    End If
'
    sgQuery = sgQuery & "    group by a.nroped, a.datped, a.codcli, c.NomCli, a.codcnd,a.codrep, d.DscCnd,"
    sgQuery = sgQuery & "             a.NomTra, a.codrep, a.sitPed, a.datlib, a.datenv,"
    sgQuery = sgQuery & "             a.TipNot , a.NroNot, a.dateminot, a.flgalt"
    sgQuery = sgQuery & "    order by 2 desc, 1 desc"
    
    Consulta sgQuery
    
    If blLoad = False And Rs.RecordCount > 50 Then
        MsgBox "Resultado retornou mais de 50 linhas." & Chr(13) & "Favor refazer seu filtro de pesquisa.", vbExclamation + vbOKOnly, "Atenção!"
    End If
    
    GrdPedido.Rows = 1
    GrdPedido.Visible = False

    blI = 0
    
    Do While Not Rs.EOF
        
        If blI > 100 Then
            Exit Do
        End If
        
        blI = blI + 1
        
        GrdPedido.Rows = GrdPedido.Rows + 1
        GrdPedido.TextMatrix(blI, 0) = Format(Trim(Rs!NroPed), "000000")
        GrdPedido.TextMatrix(blI, 1) = Format(Trim(Rs!Datped), "dd/mm/yyyy")
        GrdPedido.TextMatrix(blI, 2) = Format(Trim(Rs!Codcli), "00000")
        GrdPedido.TextMatrix(blI, 3) = Rs!NomCli
        GrdPedido.TextMatrix(blI, 4) = Format(Rs!Valor, "##,###,##0.00")
        GrdPedido.TextMatrix(blI, 5) = Rs!DscCnd
        GrdPedido.TextMatrix(blI, 6) = Rs!NomTra
        GrdPedido.TextMatrix(blI, 7) = IIf(Trim(Rs!SitPed) = "U", "C", Trim(Rs!SitPed))
        GrdPedido.TextMatrix(blI, 8) = IIf(IsNull(Rs!Datlib), "", Format(Trim(Rs!Datlib), "dd/mm/yyyy"))
        GrdPedido.TextMatrix(blI, 9) = IIf(IsNull(Rs!DatEnv), "", Format(Trim(Rs!DatEnv), "dd/mm/yyyy"))
        GrdPedido.TextMatrix(blI, 10) = IIf(IsNull(Rs!NroNot), "", Rs!TipNot & Format(Trim(Rs!NroNot), "000000"))
        GrdPedido.TextMatrix(blI, 11) = IIf(IsNull(Rs!DatEmiNot), "", Format(Trim(Rs!DatEmiNot), "dd/mm/yyyy"))
        GrdPedido.TextMatrix(blI, 12) = Rs!codrep
        GrdPedido.TextMatrix(blI, 13) = IIf(IsNull(Rs!FlgAlt), "", Rs!FlgAlt)
        GrdPedido.Row = blI
        GrdPedido.Col = 0
        GrdPedido.CellForeColor = &HFFFF&
        
        If Trim(Rs!FlgAlt) = "L" Then
            
            GrdPedido.CellBackColor = &H800080 'Roxo
            
        Else
        
            If Trim(Rs!FlgAlt) = "N" Then
                
                GrdPedido.CellBackColor = &H40C0& 'Marrom
                GrdPedido.CellForeColor = &HFFFFFF
                
            Else
                
                GrdPedido.CellBackColor = &HFF& 'Vermelho
                
            End If
            
        End If
      
        GrdPedido.Col = 15
        GrdPedido.CellBackColor = &HFFFFFF 'Branco
        
        If Not IsNull(Rs!DatEmiNot) And Trim(Rs!SitPed) <> "U" And Trim(Rs!SitPed) <> "C" Then
            
            sgQuery = "SELECT IsNull(Count(*), 0) As Conta FROM "
            sgQuery = sgQuery & "(SELECT A.CodPrd, A.QtdPrd, A.QtdPrdFat + IsNull(D.Sum_Saldo_Entregue,0) As TotEntreg "
            sgQuery = sgQuery & "FROM Pedido C "
            sgQuery = sgQuery & "INNER JOIN Item_Pedido A ON A.NroPed = C.NroPed "
            sgQuery = sgQuery & "LEFT OUTER JOIN (SELECT A.CodPrd, Sum_Saldo_Entregue = SUM(A.QtdPrdFat) FROM Item_Pedido_Saldo A, Pedido_Saldo B WHERE A.NroPed = " & Rs("NroPed") & " And A.NroPed = B.NroPed And A.NroPedSdo = B.NroPedSdo And B.SitPed = 'N' GROUP BY A.CodPrd) D ON A.CodPrd = D.CodPrd "
            sgQuery = sgQuery & "WHERE A.NroPed = " & Rs("NroPed") & ") A "
            sgQuery = sgQuery & "Where A.QtdPrd > A.TotEntreg"
            
            Consulta2 sgQuery
            
            If Not Rs2.EOF Then
                
                If Rs2!conta > 0 Then
                    GrdPedido.TextMatrix(blI, 15) = "S"
                    GrdPedido.CellForeColor = &HFFFF&
                    GrdPedido.CellBackColor = &HFF& 'Vermelho
                End If
                
            End If
            
            Rs2.Close
            
            Set Rs2 = Nothing
            
        End If
        
        Rs.MoveNext
        
    Loop

    Rs.Close
    
    Set Rs = Nothing

    'Demais Pedidos
    sgQuery = "select a.nroped, a.datped, a.codcli,c.NomCli, sum(b.vlrite) - sum(distinct a.vlrsimples) as Valor,a.codcnd,a.codrep, d.DscCnd,"
    sgQuery = sgQuery & "       a.NomTra , a.codrep, a.sitPed, a.datlib, a.datenv, TipNot, NroNot, dateminot, a.flgalt"
    sgQuery = sgQuery & "  from Pedido a, Item_pedido b, Cliente c, Condicao d"
    
'    If ChkNotif.Value = 1 And MskNroPedido.Texto = 0 Then
        sgQuery = sgQuery & " Where a.datlib is not null and a.FlgAlt is not null and a.FlgAlt <> 'N' "
        sgQuery = sgQuery & " and a.sitped = 'N' and a.nronot is null "
'    Else
'        sgQuery = sgQuery & "  Where a.datlib is not null and (a.FlgAlt <> 'N' or a.flgalt is null) "
'    End If
'
    sgQuery = sgQuery & "    and a.nroped = b.nroped"
    sgQuery = sgQuery & "    and a.codcli = c.codcli"
    sgQuery = sgQuery & "    and a.codcnd = d.codcnd"
    
'    If MskNroPedido.Texto > 0 Then
'        sgQuery = sgQuery & "    and a.nroped = " & Trim(MskNroPedido.Texto)
'    End If
    
    If APLICA = 1 Then
        sgQuery = sgQuery & "    and a.codrep = " & sgRepresentante
    End If
    
'    If Trim(ActDtini.Text) <> "" Then
'        sgQuery = sgQuery & "    and a.datped between convert(datetime,'" & Trim(ActDtini.Text) & "',103)"
'        sgQuery = sgQuery & "                     and convert(datetime,'" & Trim(ActDtfim.Text) & "',103)"
'    End If
'
'    If ilCodCli <> 0 Then
'        sgQuery = sgQuery & "    and a.codcli = " & ilCodCli
'    End If
    
    sgQuery = sgQuery & "    group by a.nroped, a.datped, a.codcli, c.NomCli, a.codcnd,a.codrep, d.DscCnd,"
    sgQuery = sgQuery & "             a.NomTra, a.codrep, a.sitPed, a.datlib, a.datenv,"
    sgQuery = sgQuery & "             a.TipNot , a.NroNot, a.dateminot, a.flgalt"
    sgQuery = sgQuery & "    order by 2 desc, 1 desc"
    
    Consulta sgQuery
    
    If Rs.EOF = True Then
        
        MsgBox "Não há histórico de pedidos emitidos para o cliente atual.", vbExclamation, "Força de Venda"
        
        GoTo SemPedidos
        
    End If
    
    Rs.MoveFirst
    
    Do While Not Rs.EOF
        
        If blI > 100 Then
            Exit Do
        End If
   
        blI = blI + 1
   
        GrdPedido.Rows = GrdPedido.Rows + 1
        GrdPedido.TextMatrix(blI, 0) = Format(Trim(Rs!NroPed), "000000")
        GrdPedido.TextMatrix(blI, 1) = Format(Trim(Rs!Datped), "dd/mm/yyyy")
        GrdPedido.TextMatrix(blI, 2) = Format(Trim(Rs!Codcli), "00000")
        GrdPedido.TextMatrix(blI, 3) = Rs!NomCli
        GrdPedido.TextMatrix(blI, 4) = Format(Rs!Valor, "##,###,##0.00")
        GrdPedido.TextMatrix(blI, 5) = Rs!DscCnd
        GrdPedido.TextMatrix(blI, 6) = Rs!NomTra
        GrdPedido.TextMatrix(blI, 7) = IIf(Trim(Rs!SitPed) = "U", "C", Trim(Rs!SitPed))
        GrdPedido.TextMatrix(blI, 8) = IIf(IsNull(Rs!DatEnv), "", Format(Trim(Rs!DatEnv), "dd/mm/yyyy"))
        GrdPedido.TextMatrix(blI, 9) = IIf(IsNull(Rs!texneg), "", Rs!texneg)
        GrdPedido.TextMatrix(blI, 10) = IIf(IsNull(Rs!FlgAlt), "", Rs!FlgAlt)
        GrdPedido.TextMatrix(blI, 11) = Rs!codrep
        GrdPedido.Row = blI
        
        If Not IsNull(Rs!DatEnv) And IsNull(Rs!DatEmiNot) Then
            
            GrdPedido.Col = 0
            GrdPedido.CellBackColor = vbYellow
                    
        End If
        
        If Not IsNull(Rs!DatEmiNot) Then
            
            GrdPedido.Col = 0
            GrdPedido.CellBackColor = &HFF00& 'Verde
        
        End If
        
        If Trim(Rs!FlgAlt) = "A" Then
            
            GrdPedido.TextMatrix(blI, 14) = "A"
            
            GrdPedido.Col = 0
            GrdPedido.CellBackColor = &H800080 'Roxo
            GrdPedido.CellForeColor = &HFFFFFF
            
            GrdPedido.Col = 14
            GrdPedido.CellBackColor = &H800080 'Roxo
            GrdPedido.CellForeColor = &HFFFFFF
            
        Else
            
            If Trim(Rs!FlgAlt) = "O" Then
                
                GrdPedido.TextMatrix(blI, 14) = "N"
                
                GrdPedido.Col = 0
                GrdPedido.CellBackColor = &H40C0& 'Marrom
                GrdPedido.CellForeColor = &HFFFFFF
                
                GrdPedido.Col = 14
                GrdPedido.CellBackColor = &H40C0& 'Marrom
                GrdPedido.CellForeColor = &HFFFFFF
                
            End If
            
        End If
        
        GrdPedido.Col = 15
        GrdPedido.CellBackColor = &HFFFFFF 'Branco
        
        If Not IsNull(Rs!DatEmiNot) And Trim(Rs!SitPed) <> "U" And Trim(Rs!SitPed) <> "C" Then
            
            sgQuery = "SELECT IsNull(Count(*), 0) As Conta FROM "
            sgQuery = sgQuery & "(SELECT A.CodPrd, A.QtdPrd, A.QtdPrdFat + IsNull(D.Sum_Saldo_Entregue,0) As TotEntreg "
            sgQuery = sgQuery & "FROM Pedido C "
            sgQuery = sgQuery & "INNER JOIN Item_Pedido A ON A.NroPed = C.NroPed "
            sgQuery = sgQuery & "LEFT OUTER JOIN (SELECT A.CodPrd, Sum_Saldo_Entregue = SUM(A.QtdPrdFat) FROM Item_Pedido_Saldo A, Pedido_Saldo B WHERE A.NroPed = " & Rs("NroPed") & " And A.NroPed = B.NroPed And A.NroPedSdo = B.NroPedSdo And B.SitPed = 'N' GROUP BY A.CodPrd) D ON A.CodPrd = D.CodPrd "
            sgQuery = sgQuery & "WHERE A.NroPed = " & Rs("NroPed") & ") A "
            sgQuery = sgQuery & "Where A.QtdPrd > A.TotEntreg"
            
            Consulta2 sgQuery
            
            If Not Rs2.EOF Then
                
                If Rs2!conta > 0 Then
                    GrdPedido.TextMatrix(blI, 15) = "S"
                    GrdPedido.CellForeColor = &HFFFF&
                    GrdPedido.CellBackColor = &HFF& 'Vermelho
                End If
                
            End If
            
            Rs2.Close
            
            Set Rs2 = Nothing
        
        End If
        
        If Trim(Rs!SitPed) = "C" Or Trim(Rs!SitPed) = "U" Then
            
            GrdPedido.Col = 0
            GrdPedido.CellBackColor = &H80000012 'Preto
            GrdPedido.CellForeColor = &HFFFF&
            
        End If
   
        Rs.MoveNext
        
    Loop

SemPedidos:

    Rs.Close
    
    blLoad = False
    
    Set Rs = Nothing
    
    GrdPedido.Visible = True

    DoEvents

'    If MskNroPedido.Texto > 0 And blI = 0 Then
'        MsgBox "Pedido inexistente", vbExclamation + vbOKOnly, "Atenção!"
'    End If

    Exit Function

TratarErro:

    Rotina_Erro "GridAviso"
    
End Function

Private Sub BtoGrava_Click()

    DoEvents

    If MskNroPedido.Texto = "" Then
        MskNroPedido.Texto = 0
    End If

    If MskNroPedido.Texto > 0 Then
        
        ActDtini.Text = ""
        ActDtfim.Text = ""
        ilCodCli = 0
        
        GoTo Compoe
        
    End If

    If ActDtini.Text = "" And ActDtini.Text = "" And ilCodCli = 0 Then
        GoTo Compoe
    End If
   
    If ActDtini.Text <> "" Or ActDtini.Text <> "" Then
        
        If CDate(ActDtini.Text) > CDate(ActDtfim.Text) Or Year(CDate(ActDtini.Text)) < 1950 Or Year(CDate(ActDtfim.Text)) < 1950 Then
            
            MsgBox "Intervalo de datas inválido", vbInformation
            
            ActDtini.SetFocus
            
            Exit Sub
            
        End If
        
    End If

Compoe:

    GridAviso

End Sub

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

Private Sub CboRemet_KeyPress(KeyAscii As Integer)
    
    'Call EventoEnter(KeyAscii)
    
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

Private Sub CmdLegenda_Click()
  
    FrmLegenda.Show vbModal
    
End Sub

Private Sub Form_Activate()

    FrmAviso.WindowState = 2
    
    blLoad = True
    
    GridAviso
    
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
    Me.Height = 10750
    Me.Width = 15360

   ' Set CboRemet.Conexao = Conexao
    
   ' MskNroPedido.Texto = 0
    
    ilCodCli = 0
    slremet = ""
    blLoad = True
    bgBloqPed = False
    bgConsultaPed = False

    GrdPedido.TextMatrix(0, 0) = "Pedido"
    GrdPedido.ColWidth(0) = 700
    GrdPedido.TextMatrix(0, 1) = "Dt.Emissão"
    GrdPedido.ColWidth(1) = 900
    GrdPedido.TextMatrix(0, 2) = "Cod.Cli"
    GrdPedido.ColWidth(2) = 600
    GrdPedido.TextMatrix(0, 3) = "Cliente"
    GrdPedido.ColWidth(3) = 3600
    GrdPedido.TextMatrix(0, 4) = "Val.Pedido"
    GrdPedido.ColWidth(4) = 1000
    GrdPedido.TextMatrix(0, 5) = "Cond.Pagto"
    GrdPedido.ColWidth(5) = 1750
    GrdPedido.TextMatrix(0, 6) = "Transp."
    GrdPedido.ColWidth(6) = 1850
    GrdPedido.TextMatrix(0, 7) = "Sit."
    GrdPedido.ColWidth(7) = 300
    GrdPedido.TextMatrix(0, 9) = "Dt. Envio"
    GrdPedido.ColWidth(9) = 1000
    GrdPedido.TextMatrix(0, 10) = "Notificação"
    GrdPedido.ColWidth(10) = 2000
    GrdPedido.TextMatrix(0, 15) = ""
    GrdPedido.ColWidth(15) = 180
    
    GridAviso

End Sub

Private Sub GrdPedido_DblClick()

'    bgBloqPed = False
'
'    If GrdPedido.RowSel = 0 Then
'        Exit Sub
'    End If
'
'    bgConsultaPed = True
'    Me.Enabled = False
'
'    igNroPed = GrdPedido.TextMatrix(GrdPedido.RowSel, 0)
'    sgRepresentante = Trim(GrdPedido.TextMatrix(GrdPedido.RowSel, 12))
'
'    If Trim(GrdPedido.TextMatrix(GrdPedido.RowSel, 9)) <> "" Then
'        bgBloqPed = True
'    End If
'
'    igTela = "PosPed"
'
'    FrmConhecimento.Show
    
End Sub

Private Sub MskNroPedido_GotFocus()
    
    Call SelecionaTudo
    
End Sub

