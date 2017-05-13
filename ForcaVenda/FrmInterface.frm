VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmInterface 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INTERFACE"
   ClientHeight    =   3690
   ClientLeft      =   2865
   ClientTop       =   3060
   ClientWidth     =   7155
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7155
   Begin VB.CommandButton btoSair 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FrmInterface.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton btoGerar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Transmitir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FrmInterface.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtResponse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   360
      Width           =   7095
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   600
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label LblProgParcial 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "FrmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************************************
'O código deste módulo foi estudado e comentado por André Corrêa, em 14/07/2008.
'*****************************************************************************************

Private vCont As Integer
Private vFileName As String
Private vOldCNPJ As String
Private vAtrelOld As String
Private vPlacaOld As String
Private vTipo() As String
Private ilFilialAnt As String
Private ilfilial As String
Private blEspera As Boolean
Private sss As Boolean
Private slExiste As Boolean
Private blErroInterf As Boolean


Public Function GravarLog(slErro As String, ilTipo As Integer)

    Dim slArq As String
    Dim slString As String
    Dim ilFile As Long
    Dim ilAcha As Integer

    ilAcha = InStr(1, slErro, "cannot find", 1)
    
    If ilAcha > 0 Then
        
        slExiste = False
        
        Exit Function
        
    End If

    ilAcha = InStr(1, slErro, "no such file", 1)
    
    If ilAcha > 0 Then
        
        slExiste = False
        
        Exit Function
        
    End If

    ilAcha = InStr(1, slErro, "file doesn't exist", 1)
    
    If ilAcha > 0 Then
        
        slExiste = False
        
        Exit Function
        
    End If

    blErroInterf = True
    
    txtResponse.Text = slErro
    
    MsgBox slErro
    
    sss = False
    slArq = "c:\INTERFACE\LOG" & Format(Date, "ddmmyyyy") & ".txt"
    ilFile = FreeFile
    
    If Dir(slArq) <> "" Then
        Open slArq For Append As #ilFile
    Else
        Open slArq For Output As #ilFile
    End If
    
    slString = slErro
    
    Print #ilFile, slString
    
    If ilTipo = 0 Then
        slString = "---------------> " & Format(CDate(Date) & " " & Time, "dd/mm/yyyy hh:mm:ss") & " <---------------"
    Else
        slString = "   "
    End If
    
    Print #ilFile, slString
    
    Close #ilFile

End Function

Public Function fPreencheEspacos(VSTRING As String, vTipo As Integer, vTamanho As Integer) As String

    '******************************************************************************
    '** Função para retornar os tamanhos corretos das strings, onde:             **
    '**   - vString é o texto ou o numero a retornar                             **
    '**   - vTipo: 0: String; 1:Inteiro; 2:Double c/ 2 casas dec; 3: Double c/ 4 **
    '**   - vTamanho: é o tamanho máximo da string retornada                     **
    '** Obs: as strings são sempre alinhadas a esquerda e os numeros a direita.  **
    '**      os numeros terão separacao de casas decimais, os ultimos            **
    '**         2 digitos são relativos às casas decimais                        **
    '**                                          Randerson Maurilio - 28/01/2004 **
    '******************************************************************************
    
    If (VSTRING = "" Or IsNull(VSTRING)) And vTipo <> 0 Then
        
        fPreencheEspacos = String(vTamanho, "0")
    
    Else
        
        Select Case vTipo
            
            Case 0:
                
                fPreencheEspacos = Mid(Trim$(VSTRING), 1, vTamanho) & String(IIf(vTamanho - Len(Trim$(VSTRING)) < 0, 0, vTamanho - Len(Trim$(VSTRING))), " ") 'STRINGS
            
            Case 1:
                
                fPreencheEspacos = Format(CDbl(VSTRING), String(vTamanho, "0")) 'Numeros Inteiros
            
            Case 2:
                
                fPreencheEspacos = Format(CDbl(VSTRING), "000.00")
                fPreencheEspacos = Right(fPreencheEspacos, vTamanho)
                fPreencheEspacos = Replace(fPreencheEspacos, ",", ".") 'Numeros reais c/ 2 casas
            
            Case 3:
                
                fPreencheEspacos = Format(CDbl(VSTRING), "0000.0000")
                fPreencheEspacos = Right(fPreencheEspacos, vTamanho)
                fPreencheEspacos = Replace(fPreencheEspacos, ",", ".") 'Numeros reais c/ 4 casas
        
        End Select
    
    End If
    
End Function

Function FTPFile(ByVal sFTPServer As String, ByVal sFTPCommand As String, ByVal sFTPUser As String, ByVal sFTPPwd As String) As Boolean

    On Error GoTo FTPFileExit

    Dim oFS As FileSystemObject
    Dim sURL As String

    On Error Resume Next

    If Inet1.StillExecuting Then
        Inet1.Cancel
    End If

    FTPFile = False

    Me.MousePointer = vbHourglass

    Set oFS = New FileSystemObject

    sURL = "ftp://" & sFTPServer

    Inet1.Protocol = icFTP
    Inet1.RequestTimeout = 100
    Inet1.RemotePort = 21
    Inet1.AccessType = icDirect
    Inet1.URL = sURL
    Inet1.UserName = sFTPUser
    Inet1.Password = sFTPPwd
    
    If sFTPCommand = "PUT" Then
        
        If oFS.FileExists(sFTPFileName) = False Then
            
            MsgBox "Arquivo " & sFTPTgtFileName & " não encontrado"
            
            GoTo FTPFileExit
        
        End If
        
        txtResponse.Text = txtResponse.Text & "Transferindo arquivo " & sFTPFileName & vbCrLf
        
        Me.Refresh
        
        Inet1.Execute , "PUT" & Space(1) & sFTPFileName & " " & sFTPTgtFileName
    
    Else
        
        If sFTPCommand = "GET" Then
            
            txtResponse.Text = txtResponse.Text & "Recebendo arquivo " & sFTPFileName & vbCrLf
            
            Me.Refresh
            
            If oFS.FileExists(sFTPFileName) = True Then
                
                oFS.DeleteFile sFTPFileName, True
                
            End If
            
            Inet1.Execute , "GET" & " " & sFTPTgtFileName & " " & sFTPFileName
            
        Else
            
            If sFTPCommand = "DIR" Then
                
                txtResponse.Text = txtResponse.Text & "Procurando arquivo " & sFTPFileName & vbCrLf
                
                Me.Refresh
                
                Inet1.Execute , "DIR" & " " & sFTPTgtFileName
            
            Else
                
                If sFTPCommand = "RENAME" Then
                    
                    txtResponse.Text = txtResponse.Text & "Renomeando arquivo " & sFTPFileName & vbCrLf
                    
                    Me.Refresh
                    
                    Inet1.Execute , "RENAME" & " " & sFTPTgtFileName & " " & sFTPFileName
                
                Else
                    
                    txtResponse.Text = txtResponse.Text & "Deletanto arquivo " & sFTPFileName & vbCrLf
                    
                    Me.Refresh
                    
                    Inet1.Execute , "DELETE" & " " & sFTPTgtFileName
                
                End If
                
            End If
            
        End If
    
    End If

    Do While Inet1.StillExecuting
        
        DoEvents
    
    Loop
    
    FTPFile = True
    
FTPFileExit:

    Set oFS = Nothing
    
    Me.MousePointer = vbDefault
    
End Function

Private Sub btoGerar_Click()

    On Error GoTo TrataErro
    
    Dim vSequencia As Integer
    Dim vCTRC As Long
    Dim vItemList As Integer
    Dim vHoraInicio As Date
    Dim vHoraFim As Date
    Dim VRecordCount As Long
    Dim slString As String
    Dim ss As Boolean
    Dim dlSeq As Double
    Dim slArqPed As String
    Dim slArqIte As String
    Dim slArqMax As String
    Dim slArqCli As String
    Dim Agora As String
    Dim slDatGer As String
    Dim slTipReg As String
    Dim slOper   As String * 1
    Dim ilSeqDsc As Integer
    Dim dlNroPed As String
    'Dim dlNroPed As Double
    Dim slMensa  As String
    Dim iloop    As Integer
    Dim slPedido As String
    Dim slItemPedido As String
    'Dim dlPedAnt As Double
    Dim dlPedAnt As String
    Dim ilCTPed As Integer
    Dim ilCTIte As Integer
    Dim ilCTPedConf As Integer
    Dim ilCTIteConf As Integer
    Dim iFile As Long
    Dim PedidoOK As Integer
    Dim slAux    As String
    Dim slNot    As String
        
    btoSair.Enabled = False
    btoGerar.Enabled = False
    blErroInterf = False
    
    Screen.MousePointer = vbHourglass
    
    Agora = Date & " " & Time
    slPedido = ""
    slItemPedido = ""
    ilCTPed = 0
    ilCTIte = 0
    ilCTPedConf = 0
    ilCTIteConf = 0
    PedidoOK = 0
        
    '*************************************************************************************
    'Avalia se existem novos pedidos para remessa à empresa.
    '*************************************************************************************
    
    '*************************************************************************************
    'O campo SeqEnv controla a quantidade de envios de pedidos enviados à fábrica via
    'interface pelo representante. A rotina a seguir levanta a quantidade atual e a
    'incrementa. Se não houver envio anterior, define seqüência = 1.
    '*************************************************************************************
    
    sgQuery = "Select max(SeqEnv) as ultimo From PEDIDO where codrep = " & sgRepresentante
    
    Consulta sgQuery
    
    If Rs.EOF Then
        dlSeq = 1
    Else
        dlSeq = IIf(IsNull(Rs!ultimo), 0, Rs!ultimo) + 1
    End If
    
    Rs.Close
    
    Set Rs = Nothing
    
    '*************************************************************************************
    'Atualiza o campo seqüência de envio de todos os pedidos liberados e não-enviados.
    '*************************************************************************************
    
    sgQuery = "Update PEDIDO set SeqEnv = " & dlSeq & " where datlib is not null and datEnv is null"
    
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
    
    '*************************************************************************************
    'Levanta todos os pedidos liberados e não-enviados. Se encontrar algum, monta o
    'arquivo com os cabeçalhos daqueles encontrados.
    '*************************************************************************************
    
    sgQuery = "Select * From PEDIDO where datlib is not null and datEnv is null"
    
    Consulta sgQuery
    
    igFileNumber = FreeFile
    
    If Rs.EOF Then
        
        MsgBox "Não existem pedidos novos LIBERADOS para remessa"
        
        GoTo Recebe
        
    Else
        
        slArqPed = "P" & Format(sgRepresentante, "0000") & Format(dlSeq, "0000") & ".TXT"
        vFileName = "c:\INTERFACE\" & Trim(slArqPed)
        slPedido = slArqPed
        
        If Dir(vFileName) <> "" Then
            Kill vFileName
        End If
        
        igFileNumber = FreeFile
        
        Open vFileName For Output As #igFileNumber
        
    End If
    
    LblProgParcial.Caption = "Gerando Arquivos . . ."
    LblProgParcial.Refresh
    
    Do While Not Rs.EOF
    
        slString = Rs!NroPed & "|" & _
        Format(Rs!Datped, "mm/dd/yyyy") & "|" & _
        Rs!Codcli & "|" & _
        Rs!codrep & "|" & _
        Rs!CodCnd & "|" & _
        Rs!CIFOB & "|" & _
        Rs!NomTra & "|" & _
        Rs!DscPdr & "|" & _
        Rs!DscPro & "|" & _
        Rs!DscCnd & "|"
        slString = slString & _
        Rs!DscFOB & "|" & _
        Rs!DscTot & "|" & _
        Rs!FlgContr & "|" & _
        Rs!UFCli & "|" & _
        Rs!AlqICM & "|" & _
        Rs!MgrMin & "|" & _
        Rs!MgrTot & "|" & _
        Rs!IdxFin & "|" & _
        Rs!IdxFrt & "|"
        slString = slString & _
        Rs!ComiNeg & "|" & _
        Rs!ComiOri & "|" & _
        Rs!ClasCor & "|" & _
        Rs!IdxPDD & "|" & _
        Rs!ChvDsc & "|" & _
        Rs!SitPed & "|" & _
        Rs!FlgKit & "|" & _
        Rs!VlrSimples & "|" & _
        Format(Rs!Datlib, "mm/dd/yyyy hh:mm:ss") & "|" & Format(CDate(Trim(Agora)), "mm/dd/yyyy hh:mm:ss") & "|" & dlSeq & "|"
        slString = Replace(slString, ",", ".")
        slString = slString & Trim(Replace(Rs!TexObs, "|", "-")) & "|" & Trim(Replace(Rs!texneg, "|", "-"))
        slString = slString & "|" & Trim(Rs!FlgAlt)
        
        Print #igFileNumber, slString
        
        ilCTPed = ilCTPed + 1
       
        Rs.MoveNext
        
    Loop
    
    Rs.Close
    
    Set Rs = Nothing
    
    Close #igFileNumber
    
    '*************************************************************************************
    'Gera os itens dos pedidos encontrados.
    '*************************************************************************************
   
    slString = ""
    sgQuery = "Select a.* From ITEM_PEDIDO a, PEDIDO B where a.nroped = b.nroped and b.datlib is not null and datEnv is null "
    
    Consulta sgQuery

    igFileNumber = FreeFile
    
    If Rs.EOF Then
        
        MsgBox "Não existem itens de pedidos novos para remessa"
        
        GoTo Recebe
        
    Else
    
        slArqIte = "I" & Format(sgRepresentante, "0000") & Format(dlSeq, "0000") & ".TXT"
        vFileName = "c:\INTERFACE\" & Trim(slArqIte)
        slItemPedido = slArqIte
        
        If Dir(vFileName) <> "" Then
            Kill vFileName
        End If
        
        igFileNumber = FreeFile
        
        Open vFileName For Output As #igFileNumber
        
    End If
    
    Do While Not Rs.EOF
        
        slString = Rs!NroPed & "|" & _
        Rs!CodPrd & "|" & _
        Rs!SeqIte & "|" & _
        Rs!qtdprd & "|" & _
        Rs!QtdEmb & "|" & _
        Rs!ValUnt & "|" & _
        Rs!IdxDsc & "|"
        slString = slString & _
        Rs!VlrIte & "|" & _
        Rs!FlgTab & "|" & _
        Rs!ValUntN & "|" & _
        Rs!MrgPrd & "|" & _
        Rs!ValCusUnt & "|" & _
        Rs!IdxFix
        slString = Replace(slString, ",", ".")
        
        Print #igFileNumber, slString
        
        ilCTIte = ilCTIte + 1
       
        Rs.MoveNext
        
    Loop
    
    Rs.Close
    
    Set Rs = Nothing

    Close #igFileNumber
    
    '*************************************************************************************
    'Inicia a configuração do controle Inet para a transmissão dos arquivos.
    '*************************************************************************************
        
    '*************************************************************************************
    'Transfere os cabeçalhos dos pedidos.
    '*************************************************************************************
    
    sFTPCommand = "PUT"
    sFTPFileName = "c:\INTERFACE\" & Trim(slArqPed)
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)
    
    Do While Inet1.StillExecuting
        DoEvents
    Loop
  
    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na transmissão dos arquivos - Transmite Pedido"
        
        Call GravarLog("E r r o  na transmissão dos arquivos - Transmite Pedido", 1)
        
        GoTo Recebe
        
    End If
    
    '*************************************************************************************
    'Transfere os itens de pedidos.
    '*************************************************************************************
    
    sFTPFileName = "c:\INTERFACE\" & Trim(slArqIte)
    sFTPTgtFileName = Trim(slArqIte)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)
   
    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na transmissão dos arquivos - Transmite Item Pedido"
        
        Call GravarLog("E r r o  na transmissão dos arquivos - Transmite Item Pedido", 1)
        
        GoTo Recebe
        
    End If

Recebe:

    '*************************************************************************************
    'Inicia o recebimento dos dados oriundos do escritório.
    '*************************************************************************************

    sFTPCommand = "GET"

    On Error Resume Next

    '*************************************************************************************
    'Recebe alterações de clientes.
    '*************************************************************************************
    
    '*************************************************************************************
    'Apaga arquivo com a última relação de clientes recebida e recebe arquivo com novas
    'alterações.
    '*************************************************************************************
    
    Kill "c:\INTERFACE\MOVICLI.TXT"
    
    slArqCli = "MOVRCLI" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPCommand = "GET"
    sFTPFileName = "c:\INTERFACE\MOVICLI.TXT"
    sFTPTgtFileName = Trim(slArqCli)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)
   
    Do While Inet1.StillExecuting
        DoEvents
    Loop
   
    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na recepção dos arquivos - Recebe Clientes"
        
        Call GravarLog("E r r o  na recepção dos arquivos - Recebe Clientes", 1)
        
        GoTo Erro
        
    End If
  
    On Error GoTo TrataErro
  
    '*************************************************************************************
    'Transfere clientes alterados do arquivo recebido para a tabela INTERF_CLIENTES na
    'base de dados.
    '*************************************************************************************

    Set Cmd = Nothing

    Conexao.BeginTrans

    sFTPFileName = "c:\INTERFACE\MOVICLI.TXT"
    
    If Dir(sFTPFileName) = "" Then
        
        Conexao.RollbackTrans
        
        GoTo continuaCli
        
    End If

    Set Cmd = New Command

    With Cmd
        .CommandText = "{call sp_InterfaceCliente}"
        .CommandType = adCmdText
        .ActiveConnection = Conexao
    End With
    
    Set Rs = Cmd.Execute
    Set Rs = Nothing
    Set Cmd = Nothing
   
    Conexao.CommitTrans
        
    '*************************************************************************************
    'Se o cliente já existe na base do representante e apenas foi alterado, executa-se um
    'UPDATE na tabela CLIENTE com dados de INTERF_CLIENTE.
    '*************************************************************************************
        
    Conexao.BeginTrans
    
    sgQuery = "update a set"
    sgQuery = sgQuery & " a.CodCli   = b.CodCli,"
    sgQuery = sgQuery & " a.NomCli   = b.NomCli,"
    sgQuery = sgQuery & " a.EndCli   = b.EndCli,"
    sgQuery = sgQuery & " a.BaiCli   = b.BaiCli,"
    sgQuery = sgQuery & " a.CidCli   = b.CidCli,"
    sgQuery = sgQuery & " a.CepCli   = b.CepCli,"
    sgQuery = sgQuery & " a.CgcCli   = b.CgcCli,"
    sgQuery = sgQuery & " a.InsCli   = b.InsCli,"
    sgQuery = sgQuery & " a.FonCli   = b.FonCli,"
    sgQuery = sgQuery & " a.CodRep   = b.CodRep,"
    sgQuery = sgQuery & " a.UFCli    = b.UFCli,"
    sgQuery = sgQuery & " a.EndCob   = b.EndCob,"
    sgQuery = sgQuery & " a.BaiCob   = b.BaiCob,"
    sgQuery = sgQuery & " a.CidCob   = b.CidCob,"
    sgQuery = sgQuery & " a.UFCob    = b.UFCob,"
    sgQuery = sgQuery & " a.CepCob   = b.CepCob,"
    sgQuery = sgQuery & " a.FlgContr = b.FlgContr,"
    sgQuery = sgQuery & " a.FlgSIMBa = b.FlgSIMBa,"
    sgQuery = sgQuery & " a.FlgSit   = b.FlgSit,"
    sgQuery = sgQuery & " a.DatPriComp = b.DatPriComp,"
    sgQuery = sgQuery & " a.DatAtu = convert(datetime, getdate(),103),"
    sgQuery = sgQuery & " a.SeqRec = b.SeqRec"
    sgQuery = sgQuery & "  from cliente a, interf_cliente b"
    sgQuery = sgQuery & "    Where a.CodCli = b.CodCli"
  
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
    
    '*************************************************************************************
    'Se o cliente é novo, então executa-se um INSERT na tabela CLIENTE com dados de
    'INTERF_CLIENTE.
    '*************************************************************************************
    
    sgQuery = "insert into cliente"
    sgQuery = sgQuery & " select a.CodCli,a.DigCli,a.NomCli,a.EndCli,a.BaiCli,a.CidCli, "
    sgQuery = sgQuery & " a.CepCli,a.CgcCli,a.InsCli,a.FonCli,a.CodRep, "
    sgQuery = sgQuery & " a.UFCli,a.EndCob,a.BaiCob,a.CidCob,a.UFCob,a.CepCob, "
    sgQuery = sgQuery & " a.FlgContr, FlgSIMBa, FlgSit, a.DatPriComp, 0, "
    sgQuery = sgQuery & " convert(DateTime, getdate(), 103), a.SeqRec"
    sgQuery = sgQuery & "  from interf_cliente a"
    sgQuery = sgQuery & "  Where not exists (select codcli from cliente"
    sgQuery = sgQuery & "                     Where CodCli = a.CodCli)"
  
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
  
    Conexao.CommitTrans
   
continuaCli:
   
    On Error Resume Next

    '*************************************************************************************
    'Apaga os arquivos de clientes alterados que acabaram de ser recebidos pelo
    'representante atual. A exclusão acontece na máquina local e no servidor.
    '*************************************************************************************

    Kill "c:\INTERFACE\MOVICLI.TXT"
    
    slArqCli = "MOVRCLI" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPCommand = "DELETE"
    sFTPFileName = "MOVRCLI" & Format(sgRepresentante, "0000") & ".OK"
    sFTPTgtFileName = Trim(slArqCli)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
        
        LblProgParcial.Caption = "E r r o no delete do arquivo no servidor -  Clientes"
        
        Call GravarLog("E r r o no delete dos arquivos no servidos -  Clientes", 1)
        
    End If
  
    '*************************************************************************************
    'Recebe cabeçalhos de pedidos alterados.
    '*************************************************************************************
    
    On Error Resume Next
    
    '*************************************************************************************
    'Apaga arquivo com transferências antigas na máquina local e inicia download de
    'possíveis arquivos disponíveis para o representante atual.
    '*************************************************************************************
    
    Kill "c:\INTERFACE\MOVIPED.TXT"
    
    sFTPCommand = "GET"
    slArqPed = "MOVRPED" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPFileName = "c:\INTERFACE\MOVIPED.TXT"
    sFTPTgtFileName = Trim(slArqPed)
    slExiste = True
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)
   
    Do While Inet1.StillExecuting
        DoEvents
    Loop
    
    If slExiste = False Then
        GoTo PedidoInc
    End If
    
    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na recepção dos arquivos - Recebe Pedidos"
        
        Call GravarLog("E r r o  na recepção dos arquivos - Recebe Pedidos", 1)
        
        GoTo Erro
        
    End If

    '*************************************************************************************
    'Recebe itens dos pedidos alterados.
    '*************************************************************************************
    
    slArqPed = "MOVRITE" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPFileName = "c:\INTERFACE\MOVITEM.TXT"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
     
        LblProgParcial.Caption = "E r r o  na recepção dos arquivos - Recebe Itens de Pedidos"
        
        Call GravarLog("E r r o  na recepção dos arquivos - Recebe Itens de Pedidos", 1)
        
        GoTo Erro
        
    End If

    On Error GoTo TrataErro

    '*************************************************************************************
    'Transfere cabeçalho dos pedidos alterados dos arquivos recebidos para a tabela
    'INTERF_PEDIDO na base de dados.
    '*************************************************************************************

    Set Cmd = Nothing

    Conexao.BeginTrans

    sFTPFileName = "c:\INTERFACE\MOVIPED.TXT"
    
    If Dir(sFTPFileName) = "" Then
        
        Conexao.RollbackTrans
        
        GoTo PedidoInc
        
    End If

    Set Cmd = New Command

    With Cmd
        .CommandText = "{call sp_InterfacePedido}"
        .CommandType = adCmdText
        .ActiveConnection = Conexao
    End With
    
    Set Rs = Cmd.Execute
    Set Rs = Nothing
    Set Cmd = Nothing

    '*************************************************************************************
    'Transfere itens dos pedidos alterados dos arquivos recebidos para a tabela
    'INTERF_ITEM_PEDIDO na base de dados.
    '*************************************************************************************

    sFTPFileName = "c:\INTERFACE\MOVITEM.TXT"
    
    If Dir(sFTPFileName) = "" Then
        GoTo ContinuaPed
    End If

    Set Cmd = New Command
    
    With Cmd
        .CommandText = "{call sp_InterfaceItemPedido}"
        .CommandType = adCmdText
        .ActiveConnection = Conexao
    End With
    
    Set Rs = Cmd.Execute
    Set Rs = Nothing
    Set Cmd = Nothing
   
ContinuaPed:
    
    Conexao.CommitTrans
    
    '*************************************************************************************
    'Se o cabeçalho do pedido já existe na base do representante e apenas foi alterado,
    'executa-se um UPDATE na tabela PEDIDO com dados de INTERF_PEDIDO (apenas para pedidos
    'não-cancelados).
    '*************************************************************************************
    
    '*************************************************************************************
    'Atualiza dados de faturamento.
    '*************************************************************************************
    
    Conexao.BeginTrans
    
    sgQuery = "Update a"
    sgQuery = sgQuery & " set a.SeqRet  = b.SeqRet,"
    sgQuery = sgQuery & "     a.TipNot  = case when b.nronot <> 0 then b.TipNot else null end,"
    sgQuery = sgQuery & "     a.NroNot  = case when b.nronot <> 0 then b.nroNot else null end,"
    sgQuery = sgQuery & "     a.DatEmiNot = case when b.nronot <> 0 then b.DatEmiNot else null end,"
    sgQuery = sgQuery & "     a.ValNot = case when b.nronot <> 0 then b.ValNot else null end,"
    sgQuery = sgQuery & "     a.DatAtu = convert(DateTime, getdate(), 103)"
    sgQuery = sgQuery & "   from pedido a, interf_NotaPedido b"
    sgQuery = sgQuery & " Where b.NroPed = a.NroPed"
    
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
    
    '*************************************************************************************
    'Atualiza situação do pedido.
    '*************************************************************************************
    
    sgQuery = "Update a"
    sgQuery = sgQuery & " set a.SitPed  = b.SitPed, a.SeqRet  = b.SeqRet,"
    sgQuery = sgQuery & "     a.DatAtu = convert(DateTime, getdate(), 103)"
    sgQuery = sgQuery & "   from pedido a, interf_NotaPedido b"
    sgQuery = sgQuery & " Where b.NroPed = a.NroPed"
    sgQuery = sgQuery & "   and b.SitPed <> 'N'"
    
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
    
    '*************************************************************************************
    'Se os itens dos pedidos já existem na base do representante e apenas foram alterados,
    'executa-se um UPDATE na tabela ITEM_PEDIDO com dados de INTERF_ITEM_PEDIDO (apenas
    'para pedidos não-cancelados).
    '*************************************************************************************
    
    sgQuery = "UPDATE A SET"
    sgQuery = sgQuery & "     a.QtdPrdFat = b.QtdPrdFat,"
    sgQuery = sgQuery & "     a.ValUntFat = b.ValUntFat,"
    sgQuery = sgQuery & "     a.IdxDscFat = b.IdxDscFat,"
    sgQuery = sgQuery & "     a.VlrIteFat = b.VlrIteFat,"
    sgQuery = sgQuery & "     a.DatAtu = convert(DateTime, getdate(), 103)"
    sgQuery = sgQuery & " from item_pedido a, interf_NotaItem_pedido b"
    sgQuery = sgQuery & "  Where b.nroped = a.nroped"
    sgQuery = sgQuery & "    and b.codprd = a.codprd"

    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing

    '*************************************************************************************
    'Se há itens de pedidos novos, então executa-se um INSERT na tabela ITEM_PEDIDO com
    'dados de INTERF_ITEM_PEDIDO.
    '*************************************************************************************
    
    sgQuery = "insert into item_pedido"
    sgQuery = sgQuery & " select distinct a.nroped, a.codprd, 99, "
    sgQuery = sgQuery & "       a.QtdPrdfat, b.QtdEmb, a.ValUntfat, a.IdxDscfat , a.VlrItefat,"
    sgQuery = sgQuery & "   ' ', 0, 0 , 0, 0 , a.QtdPrdfat, a.ValUntfat, a.IdxDscfat, a.VlrItefat, getdate()"
    sgQuery = sgQuery & " from interf_Notaitem_pedido a, produto b, pedido c"
    sgQuery = sgQuery & "  Where a.codprd = b.codprd"
    sgQuery = sgQuery & "    and a.nroped = c.nroped"
    sgQuery = sgQuery & "    and not exists (select nroped from item_pedido"
    sgQuery = sgQuery & "                     Where nroped = a.nroped"
    sgQuery = sgQuery & "                       and codprd = a.codprd)"

    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing

    Conexao.CommitTrans
   
    On Error Resume Next
    
    '*************************************************************************************
    'Apaga os arquivos de pedidos alterados que acabaram de ser recebidos pelo
    'representante atual. A exclusão acontece na máquina local e no servidor.
    '*************************************************************************************
    
    Kill "c:\INTERFACE\MOVIPED.TXT"
          
    slArqPed = "MOVRPED" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPCommand = "DELETE"
    sFTPFileName = "MOVRPED" & Format(sgRepresentante, "0000") & ".OK"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
        
        LblProgParcial.Caption = "E r r o no delete do arquivo no servidor -  Pedidos"
        
        Call GravarLog("E r r o no delete dos arquivos no servidos -  Pedidos", 1)
        
    End If
  
    '*************************************************************************************
    'Apaga os arquivos de itens de pedidos alterados que acabaram de ser recebidos pelo
    'representante atual. A exclusão acontece na máquina local e no servidor.
    '*************************************************************************************
  
    Kill "c:\INTERFACE\MOVITEM.TXT"
      
    slArqPed = "MOVRITE" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPCommand = "DELETE"
    sFTPFileName = "MOVRITE" & Format(sgRepresentante, "0000") & ".OK"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
     
        LblProgParcial.Caption = "E r r o no delete dos arquivos no servidos -  itens de Pedidos"
        
        Call GravarLog("E r r o no delete dos arquivos no servidos -  itens de Pedidos", 1)
  
    End If

PedidoInc:

    'Recebe pedidos incluidos no JP  (MINC)

    On Error Resume Next
    
    Kill "c:\INTERFACE\MINCPED.TXT"

    sFTPCommand = "GET"
    slArqPed = "MINCPED" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPFileName = "c:\INTERFACE\MINCPED.TXT"
    sFTPTgtFileName = Trim(slArqPed)
    slExiste = True
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)
   
    Do While Inet1.StillExecuting
        DoEvents
    Loop
  
    If slExiste = False Then
        GoTo PedidoSaldo
    End If
   
    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na recepção dos arquivos - Recebe Pedidos Inc."
        
        Call GravarLog("E r r o  na recepção dos arquivos - Recebe Pedidos Inc.", 1)
        
        GoTo Erro
        
    End If

    'Recebe Item de Pedidos
    slArqPed = "MINCITE" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPFileName = "c:\INTERFACE\MINCITE.TXT"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na recepção dos arquivos - Recebe Itens de Pedidos Inc."
        
        Call GravarLog("E r r o  na recepção dos arquivos - Recebe Itens de Pedidos Inc", 1)
        
        GoTo Erro
        
    End If

    On Error GoTo TrataErro

    'Importa MOVPED E MOVITEM PARA BANCO

    Set Cmd = Nothing
    
    Conexao.BeginTrans

    sFTPFileName = "c:\INTERFACE\MINCPED.TXT"
    
    If Dir(sFTPFileName) = "" Then
        Conexao.RollbackTrans
        GoTo PedidoSaldo
    End If

    Set Cmd = New Command

    With Cmd
        .CommandText = "{call sp_InterfPedido}"
        .CommandType = adCmdText
        .ActiveConnection = Conexao
    End With
    
    Set Rs = Cmd.Execute
    Set Rs = Nothing
    Set Cmd = Nothing
    
    sFTPFileName = "c:\INTERFACE\MINCITE.TXT"
    
    If Dir(sFTPFileName) = "" Then
        GoTo ContinuaPedInc
    End If

    Set Cmd = New Command
    
    With Cmd
        .CommandText = "{call sp_InterfItemPedido}"
        .CommandType = adCmdText
        .ActiveConnection = Conexao
    End With
    
    Set Rs = Cmd.Execute
    Set Rs = Nothing
    Set Cmd = Nothing
      
ContinuaPedInc:
    
    Conexao.CommitTrans
   
    'Atualiza tabelas de pedido e itens de pedido

    Conexao.BeginTrans
    
    'update de Pedidos
    sgQuery = "Update a"
    sgQuery = sgQuery & " set a.DatPed  = b.datped,"
    sgQuery = sgQuery & "     a.Codcli  = b.Codcli,"
    sgQuery = sgQuery & "     a.CodCnd  = b.CodCnd,"
    sgQuery = sgQuery & "     a.CIFOB   = b.CIFOB,"
    sgQuery = sgQuery & "     a.NomTra  = b.NomTra,"
    sgQuery = sgQuery & "     a.DscPdr  = b.DscPdr,"
    sgQuery = sgQuery & "     a.DscPro  = b.DscPro,"
    sgQuery = sgQuery & "     a.DscCnd  = b.DscCnd,"
    sgQuery = sgQuery & "     a.DscFOB  = b.DscFOB,"
    sgQuery = sgQuery & "     a.DscTot  = b.DscTot,"
    sgQuery = sgQuery & "     a.FlgContr = b.FlgContr,"
    sgQuery = sgQuery & "     a.UFCli   = b.UFCli,"
    sgQuery = sgQuery & "     a.AlqICM  = b.AlqICM,"
    sgQuery = sgQuery & "     a.MgrMin  = b.MgrMin,"
    sgQuery = sgQuery & "     a.MgrTot  = b.MgrTot,"
    sgQuery = sgQuery & "     a.IdxFin  = b.IdxFin,"
    sgQuery = sgQuery & "     a.IdxFrt  = b.IdxFrt,"
    sgQuery = sgQuery & "     a.IdxPDD  = b.IdxPDD,"
    sgQuery = sgQuery & "     a.ComiNeg = b.ComiNeg,"
    sgQuery = sgQuery & "     a.ComiOri = b.ComiOri,"
    sgQuery = sgQuery & "     a.TexNeg  = b.TexNeg,"
    sgQuery = sgQuery & "     a.TexObs  = b.TexObs,"
    sgQuery = sgQuery & "     a.ClasCor = b.ClasCor,"
    sgQuery = sgQuery & "     a.ChvDsc  = b.ChvDsc,"
    sgQuery = sgQuery & "     a.SitPed  = b.SitPed,"
    sgQuery = sgQuery & "     a.FlgKit  = b.FlgKit,"
    sgQuery = sgQuery & "     a.vlrsimples  = b.vlrsimples,"
    sgQuery = sgQuery & "     a.SeqRetInc  = b.SeqRet,"
    sgQuery = sgQuery & "     a.TipNot  = b.TipNot,"
    sgQuery = sgQuery & "     a.NroNot  = b.NroNot,"
    sgQuery = sgQuery & "     a.DatEmiNot = b.DatEmiNot,"
    sgQuery = sgQuery & "     a.ValNot = b.ValNot,"
    sgQuery = sgQuery & "     a.DatAtu = convert(DateTime, getdate(), 103)"
    sgQuery = sgQuery & "   from pedido a, interf_Pedido b"
    sgQuery = sgQuery & " Where b.NroPed = a.NroPed"
    sgQuery = sgQuery & "   and b.SitPed = 'N'"
    
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
    
    'Insert de Pedidos
    sgQuery = "insert into Pedido"
    sgQuery = sgQuery & " select a.nroped, a.datped, a.Codcli, a.CodRep, a.CodCnd, a.CIFOB,  a.NomTra,"
    sgQuery = sgQuery & "        a.DscPdr, a.DscPro, a.DscCnd, a.DscFOB, a.DscTot, a.FlgContr,"
    sgQuery = sgQuery & "        a.UFCli,  a.AlqICM, a.MgrMin, a.MgrTot, a.IdxFin, a.IdxFrt,"
    sgQuery = sgQuery & "        a.IdxPDD, a.ComiNeg,a.ComiOri,a.TexNeg, a.TexObs, a.ClasCor,"
    sgQuery = sgQuery & "        a.ChvDsc, a.datped, a.datped, 0, convert(datetime, getdate(),103),"
    sgQuery = sgQuery & "        0, a.SitPed, a.TipNot, a.NroNot, a.DatEmiNot, a.valNot, a.FlgKit, a.VlrSimples, ' ', SeqRet, a.datped"
    sgQuery = sgQuery & "   from interf_pedido a, cliente b, condicao c"
    sgQuery = sgQuery & " Where a.codcli = b.codcli"
    sgQuery = sgQuery & "   and a.codcnd = c.codcnd"
    sgQuery = sgQuery & "   and a.Nroped not in (select nroped from Pedido)"
    
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
    
    'update itens de pedidos
    sgQuery = "UPDATE A SET"
    sgQuery = sgQuery & "     a.SeqIte = b.SeqIte,"
    sgQuery = sgQuery & "     a.QtdPrd = b.QtdPrd,"
    sgQuery = sgQuery & "     a.QtdEmb = b.QtdEmb,"
    sgQuery = sgQuery & "     a.ValUnt = b.ValUnt,"
    sgQuery = sgQuery & "     a.IdxDsc = b.IdxDsc,"
    sgQuery = sgQuery & "     a.VlrIte = b.VlrIte,"
    sgQuery = sgQuery & "     a.FlgTab = b.FlgTab,"
    sgQuery = sgQuery & "     a.ValUntN = b.ValUntN,"
    sgQuery = sgQuery & "     a.MrgPrd = b.MrgPrd,"
    sgQuery = sgQuery & "     a.ValCusUnt = b.ValCusUnt,"
    sgQuery = sgQuery & "     a.IdxFix = b.IdxFix,"
    sgQuery = sgQuery & "     a.QtdPrdFat = b.QtdPrdFat,"
    sgQuery = sgQuery & "     a.ValUntFat = b.ValUntFat,"
    sgQuery = sgQuery & "     a.IdxDscFat = b.IdxDscFat,"
    sgQuery = sgQuery & "     a.VlrIteFat = b.VlrIteFat,"
    sgQuery = sgQuery & "     a.DatAtu = convert(DateTime, getdate(), 103)"
    sgQuery = sgQuery & " from item_pedido a, interf_item_pedido b, produto c"
    sgQuery = sgQuery & "  Where b.nroped = a.nroped"
    sgQuery = sgQuery & "    and b.codprd = a.codprd"
    sgQuery = sgQuery & "    and b.codprd = c.codprd"

    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
        
    'insert itens de pedidos
    sgQuery = "insert into item_pedido"
    sgQuery = sgQuery & " select distinct a.nroped, a.codprd, SeqIte, "
    sgQuery = sgQuery & "   a.QtdPrd, b.QtdEmb, a.ValUnt, a.IdxDsc , a.VlrIte,"
    sgQuery = sgQuery & "   a.FlgTab, a.ValUntN, a.MrgPrd , ValCusUnt, "
    sgQuery = sgQuery & "   Idxfix , a.QtdPrdfat, a.ValUntfat, a.IdxDscfat, a.VlrItefat, getdate()"
    sgQuery = sgQuery & " from interf_item_pedido a, produto b, pedido c"
    sgQuery = sgQuery & "  Where a.codprd = b.codprd"
    sgQuery = sgQuery & "    and a.nroped = c.nroped"
    sgQuery = sgQuery & "    and not exists (select nroped from item_pedido"
    sgQuery = sgQuery & "                     Where nroped = a.nroped"
    sgQuery = sgQuery & "                       and codprd = a.codprd)"

    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
    
    Conexao.CommitTrans
      
    On Error Resume Next
   
    'Deleta arquivos copiados (interface)
   
    Kill "c:\INTERFACE\MINCPED.TXT"
    Kill "c:\INTERFACE\MINCITE.TXT"
    
    'Renomeia arquivos copiados (FTP)

    'Pedidos
    slArqPed = "MINCPED" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPCommand = "DELETE"
    sFTPFileName = "MINCPED" & Format(sgRepresentante, "0000") & ".OK"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
        
        LblProgParcial.Caption = "E r r o no delete do arquivo no servidor -  Pedidos Inc."
        
        Call GravarLog("E r r o no delete dos arquivos no servidos -  Pedidos Inc.", 1)
        
    End If
  
    'Itens de Pedidos
    slArqPed = "MINCITE" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPCommand = "DELETE"
    sFTPFileName = "MINCITE" & Format(sgRepresentante, "0000") & ".OK"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
        
        LblProgParcial.Caption = "E r r o no delete dos arquivos no servidos -  itens de Pedidos Inc."
        
        Call GravarLog("E r r o no delete dos arquivos no servidos -  itens de Pedidos Inc.", 1)
        
    End If

PedidoSaldo:

    'Recebe Pedidos de Saldos

    On Error Resume Next
    
    Kill "c:\INTERFACE\MOVIPEDS.TXT"

    sFTPCommand = "GET"
    slArqPed = "MOVRPEDS" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPFileName = "c:\INTERFACE\MOVIPEDS.TXT"
    sFTPTgtFileName = Trim(slArqPed)
    slExiste = True
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)
   
    Do While Inet1.StillExecuting
        DoEvents
    Loop
  
    If slExiste = False Then
        GoTo Duplicatas
    End If
     
    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na recepção dos arquivos - Recebe Pedidos de Saldos"
     
        Call GravarLog("E r r o  na recepção dos arquivos - Recebe Pedidos de Saldos", 1)
     
        GoTo Erro
     
    End If

    'Recebe Item de Pedidos de saldos
    slArqPed = "MOVRITES" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPFileName = "c:\INTERFACE\MOVITEMS.TXT"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na recepção dos arquivos - Recebe Itens de Pedidos de Saldos"
        
        Call GravarLog("E r r o  na recepção dos arquivos - Recebe Itens de Pedidos de Saldos", 1)
        
        GoTo Erro
        
    End If

    On Error GoTo TrataErro

    'Importa MOVIPEDS E MOVITEMS PARA BANCO

    Set Cmd = Nothing

    sFTPFileName = "c:\INTERFACE\MOVIPEDS.TXT"
    
    If Dir(sFTPFileName) = "" Then
        
        Conexao.RollbackTrans
        
        GoTo Duplicatas
        
    End If

    Set Cmd = New Command

    With Cmd
        .CommandText = "{call sp_InterfacePedidoSaldo}"
        .CommandType = adCmdText
        .ActiveConnection = Conexao
    End With
    
    Set Rs = Cmd.Execute
    Set Rs = Nothing
    Set Cmd = Nothing

    sFTPFileName = "c:\INTERFACE\MOVITEMS.TXT"
    
    If Dir(sFTPFileName) = "" Then
        GoTo ContinuaPedSaldo
    End If

    Set Cmd = New Command
    
    With Cmd
        .CommandText = "{call sp_InterfaceItemPedidoSaldo}"
        .CommandType = adCmdText
        .ActiveConnection = Conexao
    End With
    
    Set Rs = Cmd.Execute
    Set Rs = Nothing
    Set Cmd = Nothing
      
ContinuaPedSaldo:
      
    'Atualiza tabelas de pedido e itens de pedido de saldos

    'update pedidos saldos de notas zeradas
    sgQuery = "UPDATE INTERF_PEDIDOSALDO set "
    sgQuery = sgQuery & "     nronot    = null,"
    sgQuery = sgQuery & "     tipnot    = null,"
    sgQuery = sgQuery & "     dateminot    = null,"
    sgQuery = sgQuery & "     valnot    = null"
    sgQuery = sgQuery & "  Where nronot = 0"

    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
   
    'update de Pedidos
    sgQuery = "Update a"
    sgQuery = sgQuery & " set a.DatPed  = b.datped,"
    sgQuery = sgQuery & "     a.CodCnd  = b.CodCnd,"
    sgQuery = sgQuery & "     a.CIFOB   = b.CIFOB,"
    sgQuery = sgQuery & "     a.NomTra  = b.NomTra,"
    sgQuery = sgQuery & "     a.TexObs  = b.TexObs,"
    sgQuery = sgQuery & "     a.TipNot  = b.TipNot,"
    sgQuery = sgQuery & "     a.NroNot  = b.NroNot,"
    sgQuery = sgQuery & "     a.DatEmiNot = b.DatEmiNot,"
    sgQuery = sgQuery & "     a.ValNot  = b.ValNot,"
    sgQuery = sgQuery & "     a.VlrSimples  = b.VlrSimples,"
    sgQuery = sgQuery & "     a.SitPed  = b.SitPed,"
    sgQuery = sgQuery & "     a.SeqRec  = b.SeqRec,"
    sgQuery = sgQuery & "     a.DatAtu = convert(DateTime, getdate(), 103)"
    sgQuery = sgQuery & "   from pedido_saldo a, interf_PedidoSaldo b"
    sgQuery = sgQuery & " Where b.NroPed    = a.NroPed"
    sgQuery = sgQuery & "   and b.NroPedSdo = a.NroPedSdo"
    
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
    
    'Insert de Pedidos de saldos
    sgQuery = "insert into Pedido_Saldo"
    sgQuery = sgQuery & " select a.nroped, a.nropedsdo, a.datped, a.CodCnd, a.CIFOB, a.NomTra,"
    sgQuery = sgQuery & "        a.TexObs, a.TipNot, a.NroNot, a.DatEmiNot, a.ValNot, a.VlrSimples, "
    sgQuery = sgQuery & "        a.sitped, a.seqrec, convert(DateTime, getdate(), 103)"
    sgQuery = sgQuery & "   from interf_pedidosaldo a"
    sgQuery = sgQuery & " Where not exists (select nroped from Pedido_saldo"
    sgQuery = sgQuery & "                     where nroped    = a.nroped   "
    sgQuery = sgQuery & "                       and nropedsdo = a.nropedsdo)"

    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing

    'insert itens de pedidos
    sgQuery = "insert into item_pedido_saldo"
    sgQuery = sgQuery & " select a.nroped, a.nropedsdo, a.codprd, "
    sgQuery = sgQuery & "        a.QtdPrd, a.ValUnt,"
    sgQuery = sgQuery & "        a.IdxDsc, a.VlrIte, "
    sgQuery = sgQuery & "        a.QtdPrdfat, a.ValUntfat,"
    sgQuery = sgQuery & "        a.IdxDscfat, a.VlrItefat, "
    sgQuery = sgQuery & "        convert(DateTime, getdate(), 103) "
    sgQuery = sgQuery & " from interf_itempedidoSaldo a"
    sgQuery = sgQuery & "  Where not exists (select nroped from item_pedido_saldo"
    sgQuery = sgQuery & "                     Where nroped = a.nroped"
    sgQuery = sgQuery & "                       and nropedsdo = a.nropedsdo"
    sgQuery = sgQuery & "                       and codprd = a.codprd)"

    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
    
    'update itens de pedidos saldo
    sgQuery = "UPDATE A SET"
    sgQuery = sgQuery & "     a.QtdPrd    = b.QtdPrd,"
    sgQuery = sgQuery & "     a.ValUnt    = b.ValUnt,"
    sgQuery = sgQuery & "     a.IdxDsc    = b.IdxDsc,"
    sgQuery = sgQuery & "     a.VlrIte    = b.VlrIte,"
    sgQuery = sgQuery & "     a.QtdPrdFat = b.QtdPrdfat,"
    sgQuery = sgQuery & "     a.ValUntFat = b.ValUntfat,"
    sgQuery = sgQuery & "     a.IdxDscFat = b.IdxDscfat,"
    sgQuery = sgQuery & "     a.VlrIteFat = b.VlrItefat,"
    sgQuery = sgQuery & "     a.DatAtu = convert(DateTime, getdate(), 103)"
    sgQuery = sgQuery & " from item_pedido_saldo a, interf_ItemPedidosaldo b"
    sgQuery = sgQuery & "  Where a.nroped = b.nroped"
    sgQuery = sgQuery & "    and a.nropedsdo = b.nropedsdo"
    sgQuery = sgQuery & "    and a.codprd = b.codprd"
    
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
      
    On Error Resume Next
   
    'Deleta arquivos copiados (interface)
   
    Kill "c:\INTERFACE\MOVIPEDS.TXT"
    Kill "c:\INTERFACE\MOVITEMS.TXT"
    
    'Renomeia arquivos copiados (FTP)

    'Pedidos Saldo
    slArqPed = "MOVRPEDS" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPCommand = "DELETE"
    sFTPFileName = "MOVRPEDS" & Format(sgRepresentante, "0000") & ".OK"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
        
        LblProgParcial.Caption = "E r r o no delete do arquivo no servidor -  Pedidos de saldos"
        
        Call GravarLog("E r r o no delete dos arquivos no servidos -  Pedidos", 1)
        
    End If
  
    'Itens de Pedidos Saldo
    slArqPed = "MOVRITES" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPCommand = "DELETE"
    sFTPFileName = "MOVRITES" & Format(sgRepresentante, "0000") & ".OK"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
        
        LblProgParcial.Caption = "E r r o no delete dos arquivos no servidos -  itens de Pedidos de saldos"
        
        Call GravarLog("E r r o no delete dos arquivos no servidos -  itens de Pedidos", 1)
        
    End If

Duplicatas:

    On Error Resume Next
    
    Kill "c:\INTERFACE\MOVIDUP.TXT"

    'Recebe duplicatas
    slArqPed = "MOVRDUP" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPCommand = "GET"
    sFTPFileName = "c:\INTERFACE\MOVIDUP.TXT"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)
   
    Do While Inet1.StillExecuting
        DoEvents
    Loop
   
    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na recepção dos arquivos - Recebe Duplicatas"
        
        Call GravarLog("E r r o  na recepção dos arquivos - Recebe Duplicatas", 1)
        
        GoTo Erro
        
    End If
  
    On Error GoTo TrataErro
  
    'Importa MOVDUP PARA BANCO

    Set Cmd = Nothing

    Conexao.BeginTrans

    sFTPFileName = "c:\INTERFACE\MOVIDUP.TXT"
    
    If Dir(sFTPFileName) = "" Then
        
        Conexao.RollbackTrans
        
        GoTo continuaDup
        
    End If

    Set Cmd = New Command

    With Cmd
        .CommandText = "{call sp_InterfaceDuplicata}"
        .CommandType = adCmdText
        .ActiveConnection = Conexao
    End With
    
    Set Rs = Cmd.Execute
    Set Rs = Nothing
    Set Cmd = Nothing

    Conexao.CommitTrans
      
    'Atualiza tabelas de DUPLICATAS
    Conexao.BeginTrans
  
    'Update duplicatas
    sgQuery = "update a set"
    sgQuery = sgQuery & "     a.DatEmi = b.DatEmi,"
    sgQuery = sgQuery & "     a.DatVen = b.DatVen,"
    sgQuery = sgQuery & "     a.DatPag = b.DatPag,"
    sgQuery = sgQuery & "     a.NroRec = b.NroRec,"
    sgQuery = sgQuery & "     a.VlrDup = b.VlrDup,"
    sgQuery = sgQuery & "     a.VlrDsc = b.VlrDsc,"
    sgQuery = sgQuery & "     a.VlrJur = b.VlrJur,"
    sgQuery = sgQuery & "     a.VlrPag = b.VlrPag,"
    sgQuery = sgQuery & "     a.CodBan = b.CodBan,"
    sgQuery = sgQuery & "     a.Prot   = b.Prot,"
    sgQuery = sgQuery & "     a.DatBax = b.DatBax,"
    sgQuery = sgQuery & "     a.JurDev = b.JurDev,"
    sgQuery = sgQuery & "     a.DatJur = b.DatJur,"
    sgQuery = sgQuery & "     a.DatAtu = convert(datetime, getdate(),103),"
    sgQuery = sgQuery & "     a.SeqRec = b.SeqRec"
    sgQuery = sgQuery & "  from duplicata a, interf_duplicata b"
    sgQuery = sgQuery & "    Where a.CodCli = b.CodCli"
    sgQuery = sgQuery & "      and a.TipDup = b.TipDup"
    sgQuery = sgQuery & "      and a.NroDup = b.NroDup"
    sgQuery = sgQuery & "      and a.Parc   = b.Parc"
  
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
    
    'insert duplicatas
    sgQuery = "insert into duplicata"
    sgQuery = sgQuery & " select a.CodCli, a.TipDup,a.NroDup,a.Parc,a.DatEmi,"
    sgQuery = sgQuery & "        a.DatVen,a.DatPag,a.NroRec,a.VlrDup,a.VlrDsc,"
    sgQuery = sgQuery & "        a.VlrJur,a.VlrPag,a.CodBan,a.Prot,a.DatBax,"
    sgQuery = sgQuery & "        a.JurDev , a.DatJur, convert(DateTime, getdate(), 103), a.SeqRec"
    sgQuery = sgQuery & "  from interf_duplicata a, cliente b"
    sgQuery = sgQuery & "  Where a.CodCli = b.CodCli"
    sgQuery = sgQuery & "    and not exists (select nrodup from duplicata"
    sgQuery = sgQuery & "                     Where CodCli = a.CodCli"
    sgQuery = sgQuery & "                       and TipDup = a.TipDup"
    sgQuery = sgQuery & "                       and NroDup = a.NroDup"
    sgQuery = sgQuery & "                       and Parc   = a.Parc)"
  
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
  
    Conexao.CommitTrans
   
continuaDup:
   
    On Error Resume Next

    Kill "c:\INTERFACE\MOVIDUP.TXT"

    'Rename MOVIDUP
    slArqPed = "MOVRDUP" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPCommand = "DELETE"
    sFTPFileName = "MOVRDUP" & Format(sgRepresentante, "0000") & ".OK"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
        
        LblProgParcial.Caption = "E r r o no delete do arquivo no servidor -  Duplicatas"
        
        Call GravarLog("E r r o no delete dos arquivos no servidos -  Duplicatas", 1)
        
    End If

    'Atualiza Preços e Descontos
    
    On Error GoTo TrataErro
    
    sgQuery = "select * from REPRESENTANTE where CodRep = " & Trim(sgRepresentante)
    
    Call Consulta(sgQuery)
    
    If Rs.EOF Then
        GoTo Finaliza
    Else
        slDatGer = IIf(IsNull(Rs!datger), "01/01/2000 00:00:00", Format(Rs!datger, "dd/mm/yyyy hh:mm:ss"))
    End If
    
    Rs.Close
    
    Set Rs = Nothing

    On Error Resume Next
    
    Kill "c:\INTERFACE\MOVIUTI.TXT"

    'Recebe Atulizações diversas (precos e descontos)
    slArqPed = "MOVRUTI" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPCommand = "GET"
    sFTPFileName = "c:\INTERFACE\MOVIUTI.TXT"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)
   
    Do While Inet1.StillExecuting
        DoEvents
    Loop
   
    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na recepção dos arquivos - Recebe Preço e Descontos"
        
        Call GravarLog("E r r o  na recepção dos arquivos - Recebe Preço e Descontos", 1)
        
        GoTo Erro
        
    End If
  
    sglinha = ""
    
    If Dir(sFTPFileName) <> "" Then
        
        Open sFTPFileName For Input As #1
     
        Do While Not EOF(1)
        
            Line Input #1, sglinha
            
            slTipReg = Left(sglinha, 2)
        
            On Error GoTo TrataErro
       
            Select Case slTipReg
        
                'Header
                Case "00"
            
                    slDatGer = Mid(sglinha, 3, 19)
             
                'Desconto promocional
                Case "01"
                    
                    sgQuery = "select SeqDsc from desconto_promocional"
                    sgQuery = sgQuery & " where codrep = " & Mid(sglinha, 3, 4)
                    sgQuery = sgQuery & "   and datativ = convert(datetime,'" & Mid(sglinha, 7, 10) & "',103)"
                    
                    If Trim(Mid(sglinha, 17, 5)) = "null" Then
                        sgQuery = sgQuery & " and idegrp is null "
                    Else
                        sgQuery = sgQuery & " and idegrp = " & Trim(Mid(sglinha, 17, 5))
                    End If
                    
                    If Trim(Mid(sglinha, 22, 5)) = "null" Then
                        sgQuery = sgQuery & " and CodPrd is null "
                    Else
                        sgQuery = sgQuery & " and CodPrd = " & Trim(Mid(sglinha, 22, 5))
                    End If
                    
                    Call Consulta(sgQuery)
                    
                    If Rs.EOF Then
                        slOper = "I"
                    Else
                        ilSeqDsc = Rs!SeqDsc
                        slOper = "A"
                    End If
                    
                    Rs.Close
                    
                    Set Rs = Nothing
                    
                    If slOper = "I" Then
                        sgQuery = "insert into desconto_promocional values "
                        sgQuery = sgQuery & "(" & Mid(sglinha, 3, 4) & ","
                        sgQuery = sgQuery & "convert(datetime,'" & Mid(sglinha, 7, 10) & "',103)" & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 17, 5)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 22, 5)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 27, 12)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 39, 12)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 51, 10)) & ")"
                    Else
                        sgQuery = "update desconto_promocional set "
                        sgQuery = sgQuery & " ValIni = " & Trim(Mid(sglinha, 27, 12)) & ","
                        sgQuery = sgQuery & " ValFim = " & Trim(Mid(sglinha, 39, 12)) & ","
                        sgQuery = sgQuery & " PerDsc = " & Trim(Mid(sglinha, 51, 10))
                        sgQuery = sgQuery & " where SeqDsc = " & Trim(ilSeqDsc)
                    End If
                    
                    Set Rs = Conexao.Execute(sgQuery)
                    Set Rs = Nothing
            
                'Desconto padrão
                Case "02"
                    
                    sgQuery = "select UfOri from desconto_padrao"
                    sgQuery = sgQuery & " where UfOri = '" & Mid(sglinha, 3, 2) & "'"
                    sgQuery = sgQuery & "   and UfDes = '" & Mid(sglinha, 5, 2) & "'"
                    sgQuery = sgQuery & "   and datativ = convert(datetime,'" & Mid(sglinha, 7, 10) & "',103)"
                    
                    Call Consulta(sgQuery)
                    
                    If Rs.EOF Then
                        slOper = "I"
                    Else
                        slOper = "A"
                    End If
                    
                    Rs.Close
                    
                    Set Rs = Nothing
                    
                    If slOper = "I" Then
                        sgQuery = "insert into desconto_padrao values "
                        sgQuery = sgQuery & "('" & Mid(sglinha, 3, 2) & "','"
                        sgQuery = sgQuery & Mid(sglinha, 5, 2) & "',"
                        sgQuery = sgQuery & "convert(datetime,'" & Mid(sglinha, 7, 10) & "',103)" & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 17, 10)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 27, 10)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 37, 10)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 47, 10)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 57, 10)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 67, 10)) & ")"
                    Else
                        sgQuery = "update desconto_padrao set "
                        sgQuery = sgQuery & " PerContr = " & Trim(Mid(sglinha, 17, 10)) & ","
                        sgQuery = sgQuery & " PerNContr = " & Trim(Mid(sglinha, 27, 10)) & ","
                        sgQuery = sgQuery & " PerContrSIMBa = " & Trim(Mid(sglinha, 37, 10)) & ","
                        sgQuery = sgQuery & " PerContrKit = " & Trim(Mid(sglinha, 47, 10)) & ","
                        sgQuery = sgQuery & " PerNContrKit = " & Trim(Mid(sglinha, 57, 10)) & ","
                        sgQuery = sgQuery & " PerContrSIMBaKit = " & Trim(Mid(sglinha, 67, 10))
                        sgQuery = sgQuery & " where UfOri = '" & Mid(sglinha, 3, 2) & "'"
                        sgQuery = sgQuery & "   and UfDes = '" & Mid(sglinha, 5, 2) & "'"
                        sgQuery = sgQuery & "   and datativ = convert(datetime,'" & Mid(sglinha, 7, 10) & "',103)"
                    End If
                    
                    Set Rs = Conexao.Execute(sgQuery)
                    Set Rs = Nothing
            
                'Tributação
                Case "03"
                    
                    sgQuery = "select UfOri from tributacao"
                    sgQuery = sgQuery & " where UfOri = '" & Mid(sglinha, 3, 2) & "'"
                    sgQuery = sgQuery & "   and UfDes = '" & Mid(sglinha, 5, 2) & "'"
                    sgQuery = sgQuery & "   and datativ = convert(datetime,'" & Mid(sglinha, 7, 10) & "',103)"
                    
                    Call Consulta(sgQuery)
                    
                    If Rs.EOF Then
                        slOper = "I"
                    Else
                        slOper = "A"
                    End If
                    
                    Rs.Close
                    
                    Set Rs = Nothing
                    
                    If slOper = "I" Then
                        sgQuery = "insert into tributacao values "
                        sgQuery = sgQuery & "('" & Mid(sglinha, 3, 2) & "','"
                        sgQuery = sgQuery & Mid(sglinha, 5, 2) & "',"
                        sgQuery = sgQuery & "convert(datetime,'" & Mid(sglinha, 7, 10) & "',103)" & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 17, 10)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 27, 10)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 37, 10)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 47, 10)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 57, 10)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 67, 10)) & ")"
                    Else
                        sgQuery = "update tributacao set "
                        sgQuery = sgQuery & " AlqICMContr = " & Trim(Mid(sglinha, 17, 10)) & ","
                        sgQuery = sgQuery & " AlqICMNContr = " & Trim(Mid(sglinha, 27, 10)) & ","
                        sgQuery = sgQuery & " AlqICMSimples = " & Trim(Mid(sglinha, 37, 10)) & ","
                        sgQuery = sgQuery & " AlqICMContrKit = " & Trim(Mid(sglinha, 47, 10)) & ","
                        sgQuery = sgQuery & " AlqICMNContrKit = " & Trim(Mid(sglinha, 57, 10)) & ","
                        sgQuery = sgQuery & " AlqICMSimplesKit = " & Trim(Mid(sglinha, 67, 10))
                        sgQuery = sgQuery & " where UfOri = '" & Mid(sglinha, 3, 2) & "'"
                        sgQuery = sgQuery & "   and UfDes = '" & Mid(sglinha, 5, 2) & "'"
                        sgQuery = sgQuery & "   and datativ = convert(datetime,'" & Mid(sglinha, 7, 10) & "',103)"
                    End If
                    
                    Set Rs = Conexao.Execute(sgQuery)
                    Set Rs = Nothing
            
                'Custo grupo produto
                Case "04"
                
                    sgQuery = "select IdeGrp from custo_grupo_produto"
                    sgQuery = sgQuery & " where IdeGrp = " & Mid(sglinha, 3, 5)
                    sgQuery = sgQuery & "   and datativ = convert(datetime,'" & Mid(sglinha, 8, 10) & "',103)"
                    
                    Call Consulta(sgQuery)
                    
                    If Rs.EOF Then
                        slOper = "I"
                    Else
                        slOper = "A"
                    End If
                    
                    Rs.Close
                    
                    Set Rs = Nothing
                    
                    If slOper = "I" Then
                        sgQuery = "insert into custo_grupo_produto values "
                        sgQuery = sgQuery & "(" & Mid(sglinha, 3, 5) & ","
                        sgQuery = sgQuery & "convert(datetime,'" & Mid(sglinha, 8, 10) & "',103)" & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 18, 12)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 30, 10)) & ")"
                    Else
                        sgQuery = "update custo_grupo_produto set "
                        sgQuery = sgQuery & " ValCusUnt = " & Trim(Mid(sglinha, 18, 12)) & ","
                        sgQuery = sgQuery & " IdxFix = " & Trim(Mid(sglinha, 30, 10))
                        sgQuery = sgQuery & " where IdeGrp = " & Mid(sglinha, 3, 5)
                        sgQuery = sgQuery & "   and datativ = convert(datetime,'" & Mid(sglinha, 8, 10) & "',103)"
                    End If
                    
                    Set Rs = Conexao.Execute(sgQuery)
                    Set Rs = Nothing
          
                'Preço Produto
                Case "05"
                    
                    sgQuery = "select CodPrd from preco_produto"
                    sgQuery = sgQuery & " where CodPrd = " & Mid(sglinha, 3, 5)
                    sgQuery = sgQuery & "   and datativ = convert(datetime,'" & Mid(sglinha, 8, 10) & "',103)"
                    
                    Call Consulta(sgQuery)
                    
                    If Rs.EOF Then
                        slOper = "I"
                    Else
                        slOper = "A"
                    End If
                    
                    Rs.Close
                    
                    Set Rs = Nothing
                    
                    If slOper = "I" Then
                        sgQuery = "insert into preco_produto values "
                        sgQuery = sgQuery & "(" & Mid(sglinha, 3, 5) & ","
                        sgQuery = sgQuery & "convert(datetime,'" & Mid(sglinha, 8, 10) & "',103)" & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 18, 12)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 30, 12)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 42, 12)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 54, 10)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 64, 12)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 76, 12)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 88, 10)) & ")"
                    Else
                        sgQuery = "update preco_produto set "
                        sgQuery = sgQuery & " ValUntN = " & Trim(Mid(sglinha, 18, 12)) & ","
                        sgQuery = sgQuery & " ValUntA = " & Trim(Mid(sglinha, 30, 12)) & ","
                        sgQuery = sgQuery & " ValUntB = " & Trim(Mid(sglinha, 42, 12)) & ","
                        sgQuery = sgQuery & " MrgPrd = " & Trim(Mid(sglinha, 54, 10)) & ","
                        sgQuery = sgQuery & " ValCusUntQtd = " & Trim(Mid(sglinha, 64, 12)) & ","
                        sgQuery = sgQuery & " ValCusAdicQtd = " & Trim(Mid(sglinha, 76, 12)) & ","
                        sgQuery = sgQuery & " AlqImpFed = " & Trim(Mid(sglinha, 88, 10))
                        sgQuery = sgQuery & " where CodPrd = " & Mid(sglinha, 3, 5)
                        sgQuery = sgQuery & "   and datativ = convert(datetime,'" & Mid(sglinha, 8, 10) & "',103)"
                    End If
                    
                    Set Rs = Conexao.Execute(sgQuery)
                    Set Rs = Nothing
            
                'Grupo Produto
                Case "06"
                    
                    sgQuery = "select IdeGrp from grupo_produto"
                    sgQuery = sgQuery & " where IdeGrp = " & Mid(sglinha, 3, 5)
                    
                    Call Consulta(sgQuery)
                    
                    If Rs.EOF Then
                        slOper = "I"
                    Else
                        slOper = "A"
                    End If
                    
                    Rs.Close
                    
                    Set Rs = Nothing
                    
                    slAux = Trim(Mid(sglinha, 23, 45))
                    slAux = Replace(slAux, ",", ".")
                    slAux = Replace(slAux, "'", "§")
                    slAux = Replace(slAux, """", "§")
                    slAux = Replace(slAux, "§§§§§§", """")
                    slAux = Replace(slAux, "§§§§§", """")
                    slAux = Replace(slAux, "§§§§", """")
                    slAux = Replace(slAux, "§§§", """")
                    slAux = Replace(slAux, "§§", """")
                    slAux = Replace(slAux, "§", """")
            
                    If slOper = "I" Then
                        sgQuery = "insert into grupo_produto values "
                        sgQuery = sgQuery & "(" & Mid(sglinha, 3, 5) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 8, 5)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 13, 5)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 18, 5)) & ","
                        sgQuery = sgQuery & "'" & Trim(slAux) & "'" & ")"
                    Else
                        sgQuery = "update grupo_produto set "
                        sgQuery = sgQuery & " IdeGrp1 = " & Trim(Mid(sglinha, 8, 5)) & ","
                        sgQuery = sgQuery & " IdeDep = " & Trim(Mid(sglinha, 13, 5)) & ","
                        sgQuery = sgQuery & " IdeSec = " & Trim(Mid(sglinha, 18, 5)) & ","
                        sgQuery = sgQuery & " NomGrp = '" & Trim(slAux) & "'"
                        sgQuery = sgQuery & " where IdeGrp = " & Mid(sglinha, 3, 5)
                    End If
                    
                    Set Rs = Conexao.Execute(sgQuery)
                    Set Rs = Nothing
            
                'Produto
                Case "07"
                    
                    sgQuery = "select CodPrd from produto"
                    sgQuery = sgQuery & " where CodPrd = " & Mid(sglinha, 3, 5)
                    
                    Call Consulta(sgQuery)
                    
                    If Rs.EOF Then
                        slOper = "I"
                    Else
                        slOper = "A"
                    End If
                    
                    Rs.Close
                    
                    Set Rs = Nothing
                    
                    slAux = Trim(Mid(sglinha, 29, 50))
                    slAux = Replace(slAux, ",", ".")
                    slAux = Replace(slAux, "'", "§")
                    slAux = Replace(slAux, """", "§")
                    slAux = Replace(slAux, "§§§§§§", """")
                    slAux = Replace(slAux, "§§§§§", """")
                    slAux = Replace(slAux, "§§§§", """")
                    slAux = Replace(slAux, "§§§", """")
                    slAux = Replace(slAux, "§§", """")
                    slAux = Replace(slAux, "§", """")
            
                    If slOper = "I" Then
                        sgQuery = "insert into produto values "
                        sgQuery = sgQuery & "(" & Mid(sglinha, 3, 5) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 8, 5)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 13, 5)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 18, 5)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 23, 5)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 28, 1)) & ","
                        sgQuery = sgQuery & "'" & Trim(slAux) & "',"
                        sgQuery = sgQuery & Trim(Mid(sglinha, 79, 11)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 90, 5)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 95, 5)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 100, 1)) & ","
                        sgQuery = sgQuery & "'" & Trim(Mid(sglinha, 101, 1)) & "',"
                        sgQuery = sgQuery & "'" & Trim(Mid(sglinha, 102, 1)) & "')"
                    Else
                        sgQuery = "update produto set "
                        sgQuery = sgQuery & " IdeGrp = " & Trim(Mid(sglinha, 8, 5)) & ","
                        sgQuery = sgQuery & " IdeGrp1 = " & Trim(Mid(sglinha, 13, 5)) & ","
                        sgQuery = sgQuery & " IdeDep = " & Trim(Mid(sglinha, 18, 5)) & ","
                        sgQuery = sgQuery & " IdeSec = " & Trim(Mid(sglinha, 23, 5)) & ","
                        sgQuery = sgQuery & " DigPrd = " & Trim(Mid(sglinha, 28, 1)) & ","
                        sgQuery = sgQuery & " DscPrd = '" & Trim(slAux) & "',"
                        sgQuery = sgQuery & " PesUnt = " & Trim(Mid(sglinha, 79, 11)) & ","
                        sgQuery = sgQuery & " QtdEmb = " & Trim(Mid(sglinha, 90, 5)) & ","
                        sgQuery = sgQuery & " SeqGrp = " & Trim(Mid(sglinha, 95, 5)) & ","
                        sgQuery = sgQuery & " FlgKit = " & Trim(Mid(sglinha, 100, 1)) & ","
                        sgQuery = sgQuery & " FlgInteg = '" & Trim(Mid(sglinha, 101, 1)) & "',"
                        sgQuery = sgQuery & " FlgSitu = '" & Trim(Mid(sglinha, 102, 1)) & "'"
                        sgQuery = sgQuery & " where Codprd = " & Mid(sglinha, 3, 5)
                    End If
                    
                    Set Rs = Conexao.Execute(sgQuery)
                    Set Rs = Nothing
                      
                'Condicao
                Case "08"
                    
                    sgQuery = "select CodCnd from Condicao"
                    sgQuery = sgQuery & " where CodCnd = " & Mid(sglinha, 3, 5)
                    
                    Call Consulta(sgQuery)
                    
                    If Rs.EOF Then
                        slOper = "I"
                    Else
                        slOper = "A"
                    End If
                    
                    Rs.Close
                    
                    Set Rs = Nothing
                    
                    slAux = Trim(Mid(sglinha, 8, 40))
                    slAux = Replace(slAux, ",", ".")
                    slAux = Replace(slAux, "'", "§")
                    slAux = Replace(slAux, """", "§")
                    slAux = Replace(slAux, "§§§§§§", """")
                    slAux = Replace(slAux, "§§§§§", """")
                    slAux = Replace(slAux, "§§§§", """")
                    slAux = Replace(slAux, "§§§", """")
                    slAux = Replace(slAux, "§§", """")
                    slAux = Replace(slAux, "§", """")
            
                    If slOper = "I" Then
                        sgQuery = "insert into condicao values "
                        sgQuery = sgQuery & "(" & Mid(sglinha, 3, 5) & ","
                        sgQuery = sgQuery & "'" & Trim(slAux) & "',"
                        sgQuery = sgQuery & Trim(Mid(sglinha, 48, 5)) & ","
                        sgQuery = sgQuery & Trim(Mid(sglinha, 53, 5)) & ","
                        sgQuery = sgQuery & "'" & Trim(Mid(sglinha, 58, 1)) & "')"
                    Else
                        sgQuery = "update condicao set "
                        sgQuery = sgQuery & " DscCnd = '" & Trim(slAux) & "',"
                        sgQuery = sgQuery & " QtdParCnd = " & Trim(Mid(sglinha, 48, 5)) & ","
                        sgQuery = sgQuery & " PrzMed = " & Trim(Mid(sglinha, 53, 5)) & ","
                        sgQuery = sgQuery & " BlqCnd = '" & Trim(Mid(sglinha, 58, 1)) & "'"
                        sgQuery = sgQuery & " where CodCnd = " & Mid(sglinha, 3, 5)
                    End If
                    
                    Set Rs = Conexao.Execute(sgQuery)
                    Set Rs = Nothing
            
            'Custo_Condicao
            Case "09"
            
                sgQuery = "select CodCnd from Custo_Condicao"
                sgQuery = sgQuery & " where CodCnd = " & Mid(sglinha, 3, 5)
                sgQuery = sgQuery & "   and datativ = convert(datetime,'" & Mid(sglinha, 8, 10) & "',103)"
                
                Call Consulta(sgQuery)
                
                If Rs.EOF Then
                    slOper = "I"
                Else
                    slOper = "A"
                End If
                
                Rs.Close
                
                Set Rs = Nothing
                
                If slOper = "I" Then
                    sgQuery = "insert into Custo_Condicao values "
                    sgQuery = sgQuery & "(" & Mid(sglinha, 3, 5) & ","
                    sgQuery = sgQuery & "convert(datetime,'" & Mid(sglinha, 8, 10) & "',103)" & ","
                    sgQuery = sgQuery & Trim(Mid(sglinha, 18, 10)) & ","
                    sgQuery = sgQuery & Trim(Mid(sglinha, 28, 10)) & ")"
                Else
                    sgQuery = "update Custo_Condicao set "
                    sgQuery = sgQuery & " PerCusFin = " & Trim(Mid(sglinha, 18, 10)) & ","
                    sgQuery = sgQuery & " PerDesCnd = " & Trim(Mid(sglinha, 28, 10))
                    sgQuery = sgQuery & " where CodCnd = " & Mid(sglinha, 3, 5)
                    sgQuery = sgQuery & "   and datativ = convert(datetime,'" & Mid(sglinha, 8, 10) & "',103)"
                End If
                
                Set Rs = Conexao.Execute(sgQuery)
                Set Rs = Nothing
          
            End Select
        
            On Error Resume Next
      
        Loop
         
    End If

    On Error GoTo TrataErro
    
    Conexao.BeginTrans
    
    sgQuery = "Update representante set DatGer = convert(datetime,'" & Trim(slDatGer) & "',103) "
    sgQuery = sgQuery & " where codrep = " & sgRepresentante
    
    Set Rs = Conexao.Execute(sgQuery)
    
    Set Rs = Nothing
    
    Conexao.CommitTrans

    On Error Resume Next
    
    Close #1

    Kill "c:\INTERFACE\MOVIUTI.TXT"

    'delete MOVRUTI
    slArqPed = "MOVRUTI" & Format(sgRepresentante, "0000") & ".TXT"
    sFTPCommand = "DELETE"
    sFTPFileName = "MOVRUTI" & Format(sgRepresentante, "0000") & ".OK"
    sFTPTgtFileName = Trim(slArqPed)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
    
        LblProgParcial.Caption = "E r r o no delete do arquivo no servidor - preços e descontos"
        
        Call GravarLog("E r r o no delete dos arquivos no servidos -  preços e descontos", 1)
        
    End If

    'Libera Pedidos para alteração

    On Error Resume Next
    Dim Num As Integer
    For Num = 1 To 5
    
        Kill "c:\INTERFACE\MOVRALT.TXT"
    
        'Recebe liberações de alterações de pedidos
        slArqPed = "MOVRALT" & Format(sgRepresentante, "0000_") & Num & ".TXT"
        sFTPCommand = "GET"
        sFTPFileName = "c:\INTERFACE\MOVRALT.TXT"
        sFTPTgtFileName = Trim(slArqPed)
        sss = True
        ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)
       
        Do While Inet1.StillExecuting
            DoEvents
        Loop
       
        If sss = False Then
            
            LblProgParcial.Caption = "E r r o  na recepção dos arquivos - Recebe liberação pedidos"
            
            Call GravarLog("E r r o  na recepção dos arquivos - Recebe liberação pedidos", 1)
            
            GoTo Erro
            
        End If
      
        sglinha = ""
        
        If Dir(sFTPFileName) <> "" Then
        
            Open sFTPFileName For Input As #1
         
            dlPedAnt = "" '0
         
            Do While Not EOF(1)
                
                Line Input #1, sglinha
                dlNroPed = Left(sglinha, 6)
                
                If dlNroPed <> dlPedAnt Then
                    
                    If dlPedAnt = "" Then
                        
                        If Mid(sglinha, 7, 1) = "1" Then
                            slNot = "A"
                        Else
                            slNot = "N"
                        End If
                        
                        slMensa = Mid(sglinha, 8, Len(Trim(sglinha))) & vbCrLf
                        dlPedAnt = dlNroPed
                    
                    Else
                        
                        On Error GoTo TrataErro
                        
                        sgQuery = "select * from Pedido"
                        
                        sgQuery = sgQuery & " where nroped = '" & Trim(dlPedAnt) & "'"
                        
                        Call Consulta(sgQuery)
                        
                        If Rs.EOF Then
                            
                            slOper = "I"
                            
                        Else
                            
                            slOper = "A"
                            
                            If Trim(Rs!texneg) <> "" Then
                                slMensa = Trim(Rs!texneg) & vbCrLf & Trim(slMensa)
                            End If
                        
                        End If
                        
                        Rs.Close
                        
                        Set Rs = Nothing
                        
                        If slOper = "A" Then
                            
                            If slNot = "A" Then
                                sgQuery = "update pedido set datlib = null, DatEnv = null, FlgAlt = '" & slNot & "',"
                                sgQuery = sgQuery & " TexNeg = '" & Trim(slMensa) & "'"
                                sgQuery = sgQuery & " where nroped = '" & Trim(dlPedAnt) & "'"
                            Else
                                sgQuery = "update pedido set FlgAlt = '" & slNot & "',"
                                sgQuery = sgQuery & " TexNeg = '" & Trim(slMensa) & "'"
                                sgQuery = sgQuery & " where nroped = '" & Trim(dlPedAnt) & "'"
                            End If
                            
                            Set Rs = Conexao.Execute(sgQuery)
                            Set Rs = Nothing
                        
                        End If
                        
                        On Error Resume Next
                        
                        If Mid(sglinha, 7, 1) = "1" Then
                            slNot = "A"
                        Else
                            slNot = "N"
                        End If
                        
                        slMensa = Mid(sglinha, 8, Len(Trim(sglinha))) & vbCrLf
                        dlPedAnt = dlNroPed
                    
                    End If
                
                Else
                
                    If Mid(sglinha, 7, 1) = "1" Then
                        slNot = "A"
                    Else
                        slNot = "N"
                    End If
                    
                    slMensa = slMensa & Mid(sglinha, 8, Len(Trim(sglinha))) & vbCrLf
                
                End If
                
            Loop
             
        End If
      
        Close #1
      
        On Error GoTo TrataErro
     
        sgQuery = "select * from Pedido"
        sgQuery = sgQuery & " where nroped = '" & Trim(dlNroPed) & "'"
        
        Call Consulta(sgQuery)
        
        If Rs.EOF Then
            
            slOper = "I"
            
        Else
            
            slOper = "A"
            
            If Trim(Rs!texneg) <> "" Then
                slMensa = Trim(Rs!texneg) & vbCrLf & Trim(slMensa)
            End If
        
        End If
        
        Rs.Close
        
        Set Rs = Nothing
        
        If slOper = "A" Then
        
            If slNot = "A" Then
                sgQuery = "update pedido set datlib = null, DatEnv = null, FlgAlt = 'L',"
                sgQuery = sgQuery & " TexNeg = '" & Trim(slMensa) & "'"
                sgQuery = sgQuery & " where nroped = '" & Trim(dlPedAnt) & "'"
            Else
                sgQuery = "update pedido set FlgAlt = 'N',"
                sgQuery = sgQuery & " TexNeg = '" & Trim(slMensa) & "'"
                sgQuery = sgQuery & " where nroped = '" & Trim(dlPedAnt) & "'"
            End If
            
            Set Rs = Conexao.Execute(sgQuery)
            Set Rs = Nothing
            
        End If
    
        On Error Resume Next
    
        Kill "c:\INTERFACE\MOVRALT.TXT"
    
        'delete MOVRALT
       ' slArqPed = "MOVRALT" & Format(sgRepresentante, "0000") & ".TXT"
        slArqPed = "MOVRALT" & Format(sgRepresentante, "0000_") & Num & ".TXT"
        sFTPCommand = "DELETE"
        sFTPFileName = "MOVRALT" & Format(sgRepresentante, "0000") & ".OK"
        sFTPTgtFileName = Trim(slArqPed)
        sss = True
        ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)
    
        Do While Inet1.StillExecuting
            DoEvents
        Loop
    
        If sss = False Then
            
            LblProgParcial.Caption = "E r r o no delete do arquivo no servidor - liberação pedidos"
            
            Call GravarLog("E r r o no delete dos arquivos no servidos - liberação pedidos", 1)
        
        End If
    Next Num
Finaliza:

    'CONFERE EXISTÊNCIA DOS PEDIDOS ENVIADOS NO SERVIDOR FTP (COPIA REVERSA)

    On Error Resume Next

    If blErroInterf = True Then
        GoTo Erro
    End If

    'Pedidos
    If slPedido = "" Then
        GoTo continua_final
    End If

    vFileName = "c:\INTERFACE\" & Trim(slPedido)
    
    If Dir(vFileName) <> "" Then
        Kill vFileName
    End If

    sFTPCommand = "GET"
    slArqPed = slPedido
    sFTPFileName = "c:\INTERFACE\" & Trim(slPedido)
    sFTPTgtFileName = Trim(slArqPed)
    slExiste = True
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)
 
    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na recepção dos arquivos - Recebe Pedidos (REVERSA)"
        
        Call GravarLog("E r r o  na recepção dos arquivos - Recebe Pedidos (REVERSA)", 1)
        
        GoTo Erro
        
    End If
    
    If slExiste = False Then
        GoTo Erro
    End If

    'reconta os pedidos enviados
    ilCTPedConf = 0
    
    On Error Resume Next
    iFile = FreeFile
    
    Open "c:\INTERFACE\" & Trim(slPedido) For Input As iFile

    Do While Not EOF(iFile)
    
        Input #iFile, sglinha
        
        ilCTPedConf = ilCTPedConf + 1
        
    Loop

    Close #iFile

    'If ilCTPed <> ilCTPedConf Then
        'GoTo erro
    'End If

    If ilCTPedConf = 0 Then
        GoTo Erro
    End If

    PedidoOK = 1

    'Item de Pedidos
    
    If slItemPedido = "" Then
        GoTo continua_final
    End If

    vFileName = "c:\INTERFACE\" & Trim(slItemPedido)
    
    If Dir(vFileName) <> "" Then
        Kill vFileName
    End If

    sFTPCommand = "GET"
    slArqPed = slItemPedido
    sFTPFileName = "c:\INTERFACE\" & Trim(slItemPedido)
    sFTPTgtFileName = Trim(slArqIte)
    slExiste = True
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)
 
    Do While Inet1.StillExecuting
        DoEvents
    Loop

    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na recepção dos arquivos - Recebe item de Pedidos (REVERSA)"
        
        Call GravarLog("E r r o  na recepção dos arquivos - Recebe Item de Pedidos (REVERSA)", 1)
        
        GoTo Erro
    
    End If

    If slExiste = False Then
        GoTo Erro
    End If

    ilCTIteConf = 0
    
    On Error Resume Next
    
    iFile = FreeFile
    
    Open "c:\INTERFACE\" & Trim(slItemPedido) For Input As iFile

    Do While Not EOF(iFile)
        
        Input #iFile, sglinha
        
        ilCTIteConf = ilCTIteConf + 1
    
    Loop

    Close #iFile
    
    'If ilCTIte <> ilCTIteConf Then
        'GoTo erro
    'End If

    If ilCTIteConf = 0 Then
        GoTo Erro
    End If

    PedidoOK = 2

continua_final:

    'Atualiza max

    If blErroInterf = True Then
        GoTo pula_max
    End If
  
    slString = ""
    sgQuery = "select (SELECT ISNULL(max(SEQREC),0) FROM  CLIENTE) AS CLI,"
    sgQuery = sgQuery & " (SELECT ISNULL(max(SEQREC),0) FROM DUPLICATA) AS DUP,"
    sgQuery = sgQuery & "  (SELECT ISNULL(max(SEQRET),0) FROM PEDIDO) AS PED,"
    sgQuery = sgQuery & "  (SELECT ISNULL(max(SEQRETINC),0) FROM PEDIDO) AS PEDINC,"
    sgQuery = sgQuery & "  (SELECT ISNULL(max(SEQREC),0) FROM PEDIDO_SALDO) AS SDO"
    
    Consulta sgQuery
    
    igFileNumber = FreeFile
    
    If Rs.EOF Then
    
        GoTo Erro
        
    Else
    
        slArqMax = "Max" & Format(sgRepresentante, "0000") & ".TXT"
        vFileName = "c:\INTERFACE\" & Trim(slArqMax)
        
        If Dir(vFileName) <> "" Then
            
            Kill vFileName
        
        End If
        
        igFileNumber = FreeFile
        
        Open vFileName For Output As #igFileNumber
        
    End If
    
    slString = Rs!CLI & "|" & Rs!DUP & "|" & Rs!PED & "|" & Rs!SDO & "|" & Trim(slDatGer) & Format(Rs!PEDINC, "0000")
    Print #igFileNumber, slString
    
    Rs.Close
    
    Set Rs = Nothing
    
    Close #igFileNumber
        
    'Transmissão do ARQUIVO MAX
    sFTPCommand = "PUT"
   
    'Transfere Max
    sFTPFileName = "c:\INTERFACE\" & Trim(slArqMax)
    sFTPTgtFileName = Trim(slArqMax)
    sss = True
    ss = FTPFile(sFTPServer, sFTPCommand, sFTPUser, sFTPPwd)

    Do While Inet1.StillExecuting
        iloop = 1
    Loop
  
    If sss = False Then
        
        LblProgParcial.Caption = "E r r o  na transmissão dos arquivos - Transmite Max"
        
        Call GravarLog("E r r o  na transmissão dos arquivos - Transmite Max", 1)
    
    End If
  
    Inet1.Execute , "quit"

pula_max:

    
    On Error GoTo TrataErro
    
    MDIProjUNO.Enabled = True
    
    If blErroInterf = True Then
        LblProgParcial.Caption = "T R A N S M I S S Ã O   C O M P L E T A! *****"
    Else
        LblProgParcial.Caption = "T R A N S M I S S Ã O   C O M P L E T A!"
    End If

    If PedidoOK = 2 Then
        
        Conexao.BeginTrans
        
        sgQuery = "Update PEDIDO set DatEnv = convert(datetime,'" & Trim(Agora) & "',103), SeqEnv = " & dlSeq
        sgQuery = sgQuery & " where datlib is not null and datEnv is null"
        
        Set Rs = Conexao.Execute(sgQuery)
        Set Rs = Nothing
        
        Conexao.CommitTrans
        
    End If

    MsgBox "Interface concluída !"
    
    Screen.MousePointer = vbDefault
    
    Unload Me
    
    Set FrmInterface = Nothing
    
    Exit Sub
   
Erro:

    MDIProjUNO.Enabled = True
    
    Screen.MousePointer = vbDefault
    
    LblProgParcial.Caption = "E R R O  NA INTERFACE!"
    
    MsgBox "E R R O  NA INTERFACE!, Se o problema persistir, favor comunicar ao administrador do sistema"
    
    LblProgParcial.Refresh
    btoGerar.Enabled = False
    btoSair.Enabled = True

    Exit Sub
   
TrataErro:

    MDIProjUNO.Enabled = True
    
    'Conexao.RollbackTrans
    
    Rotina_Erro "Interface arquivos"
    
End Sub
Private Sub BtoSair_Click()
    
    MDIProjUNO.Enabled = True
    
    Unload Me
    
    Set FrmInterface = Nothing
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Call EventoEnter(KeyAscii)
    
End Sub

Private Sub Form_Load()
    
    Me.Left = 0
    Me.Top = 0
    MDIProjUNO.Enabled = False
    btoGerar.Enabled = True
    sFTPServer = "201.65.158.22"
    sFTPUser = "unocann"
    sFTPPwd = "unodataac5621"
    'Senha alterada em 13/09/2010 - Afonso
    'sFTPPwd = "u2n4o5c3a1n4n4"
'
'    If strSenha = "" Then
'        sFTPPwd = "unodataac5621"
'    Else
'        sFTPPwd = strSenha
'    End If
    
    'sFTPServer = "ftp.unocann.com.br"
    'sFTPUser = "unocann"
    'sFTPPwd = "unodataac5621"
    'sFTPServer = "200.234.196.121"
    'sFTPUser = "capilefinanceira"
    'sFTPPwd = "awatdu8897"
 
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

    On Error Resume Next
    
    Select Case State
        
        Case icNone
        
        
        
        Case icResolvingHost
        
            txtResponse.Text = txtResponse.Text & "Resolvendo Host" & vbCrLf
        
        Case icHostResolved
        
            txtResponse.Text = txtResponse.Text & "Host Resolvido" & vbCrLf
        
        Case icConnecting
        
            txtResponse.Text = txtResponse.Text & "Conectando..." & vbCrLf
        
        Case icConnected
        
            txtResponse.Text = txtResponse.Text & "Conectado" & vbCrLf
        
        Case icResponseReceived
        
            If sFTPCommand = "GET" Then
                txtResponse.Text = txtResponse.Text & "Recebendo arquivo..." & sFTPTgtFileName & vbCrLf
            ElseIf sFTPCommand = "PUT" Then
                txtResponse.Text = txtResponse.Text & "Transferindo arquivo..." & sFTPTgtFileName & vbCrLf
            ElseIf sFTPCommand = "RENAME" Then
                txtResponse.Text = txtResponse.Text & "Renomeando arquivo..." & sFTPTgtFileName & vbCrLf
            ElseIf sFTPCommand = "DELETE" Then
                txtResponse.Text = txtResponse.Text & "Deletando arquivo..." & sFTPTgtFileName & vbCrLf
            End If
        
        Case icDisconnecting
        
            txtResponse.Text = txtResponse.Text & "Disconectando..." & vbCrLf
        
        Case icDisconnected
        
            txtResponse.Text = txtResponse.Text & "Disconectado" & vbCrLf
        
        Case icError:
            
            'txtResponse.Text = "Error:" & Inet1.ResponseCode & " " & Inet1.ResponseInfo
        
            Call GravarLog("Erro: " & Inet1.ResponseCode & " " & Inet1.ResponseInfo, 0)
            
        Case icResponseCompleted
    
            txtResponse.Text = txtResponse.Text & "Processo Completo." & vbCrLf
    
    End Select
    
    txtResponse.SelStart = Len(txtResponse.Text)
    txtResponse.Refresh

    Err.Clear
    
End Sub
