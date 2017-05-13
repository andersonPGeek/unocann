Attribute VB_Name = "MdlVariaveis"
Option Explicit

'DATAS
Public datah As Date    '' hoje
Public data1 As Date    '' Primeiro dia do m�s

Public sgQuery As String 'Utilizada para fazer as pesquisas SQL
Public sgQuery1 As String 'Utilizada para fazer as pesquisas SQL
Public sglinha As String
Public sgFlagOper As String * 1 'Informa o Flag Operaional (I=Inclus�o;A=Altera��o)
Public blI As Integer 'Variavel auxiliar - contador
Public sFTPServer As String
Public sFTPCommand As String
Public sFTPUser As String
Public sFTPPwd As String
Public sFTPFileName As String
Public sFTPTgtFileName As String
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'****** Variaveis utilizada no nivel de acesso *********
Public bgSenOK As Boolean ' Indica digita��o correta da senha
Public bgSenComi As Boolean ' Indica chamada rotina de senha para confirma��o comiss�o
Public LgCodUsuSis As Long   'Informa o Codigo do Usu�rio do sistema

Public igCodFil As Integer 'C�digo da Filial do Usu�rio
Public igCodFilCtrc As Integer 'C�digo da Filial do conhecimento
Public igCodCli As Integer
Public lgSeqLig As Long ' Sequencia da liga��o (telemarketing)
Public bgPosLig As Long ' Indica sele��o de liga��o (posi��o de liga��o)
Public bgPedMKT As Boolean  ' indica se pedido foi chamado do TeleMarketing
Public bgSimula As Boolean  ' indica se � uma simula��o de pedido

Public sgFlgUsu As String
Public sgNomUsuSis As String 'Informa o Nome do Usu�rio do sistema
Public sgservidor As String * 30 ' servidor que o sistema vai utilizar
Public sgsenhadb As String ' senha de acesso ao banco de dados
Public sgRepresentante As String
Public igNroPed As String
'Public igNroPed As Double
Public bgConsultaPed As Boolean
Public bgBloqPed As Boolean
Public bgabertura As Boolean
Public igTela As String

'***** Vari�veis utilizadas no Atrelamento ******
Public ogForm As Object 'Cria um form Auxiliar
Public igFileNumber As Long   'Informa qual o Arquivo livre (FREEFILE) para impress�o DOS - Randerson 10/11/2003
Public sgPortaImp As String 'Informa o nome da porta onde ser� feita a emiss�o no formato DOS - Randerson 13/11/2003
Public sgEstados As String 'Guarda os estados da tabela ESTADOS para fazer as consistencias - Randerson 12/11/2003


'**** Vari�veis Globais de Filiais ***************
'Public igCodFilAnt As Integer
'Public sgCodCCFil As String * 1
'Public sgRepomFil As String * 1
'Public sgDigAprFil As String * 1
'Public sgDigLibFil As String * 1
'Public sgBerUsiFil As String * 1
'Public igForConFil As Integer
'Public igForNFFil As Integer
'Public igForRFFil As Integer
'Public igForFatFil As Integer
'Public igForOCFil As Integer
'Public lgPerDesPneuFil As Long
'Public sgBolAutFil As String * 1
'Public sgImpISSQNFil As String * 1
'Public sgImpINSSFil As String * 1
'Public sgFreRetFrota As String * 1
'Public sgGuiCarFil As String * 1
'Public sgProGuiCarFil As String * 1
'Public dgDatProFil As Date
'Public igAlqINSSFil As Single
'Public igAlqISSFil As Single
'Public igPerAdiFro As Single
'Public igPerAdiTer As Single
'Public igEstMinBerUsi As Integer
'Public igEstMinBerInt As Integer
'Public sgFlgDigPed As String * 1
'Public sgPedIncFre As String * 1
'Public sgFlgPrgUsi As String * 5
'Public sgFlgNumBol As String * 1
'Public sgFlgNumPrg As String * 1
'Public sgFlgDimPrg As String * 1
'Public sgFlgItePrg As String * 1
'Public sgFlgRotMaxion As String * 1
'Public sgFlgCgaNfUsi As String * 1
'Public sgFlgAltPso As String * 1
'Public sgUfFil As String * 2
'Public sgCNPJFil As String * 14
'Public sgNomeEmp As String * 40
'Public dgPerSegMet As Double

'Guarda a senha de conexao com o servidor FTP
Public strSenha As String

'VARIAVEIS DE MANUTENCAO DO BANCO DE DADOS
Public Conexao As ADODB.Connection
Public Rs As ADODB.Recordset
Public Rs2 As ADODB.Recordset
Public Cmd As ADODB.Command

Dim constr As String

'***************************************************************************
'Constantes utilizadas nos programas para padronizar formata��es das strings
'Randerson Maurilio - 29/08/2003
'***************************************************************************
Public Const sgStrData = "__/__/____"
Public Const sgStrHora = "__:__"
Public Const sgStrDataHora = "__/__/____ __:__"
Public Const sgStrMesAno = "__/____"
Public Const sgStrF0 = "#,##0"
Public Const sgStrF2 = "#,##0.00"
Public Const sgStrF4 = "#,##0.0000"
Public Const LIMITE_PESQUISA = 60

Public APLICA As Long
Public blCarregou As Boolean

'Vari�veis do Conhecimento/Viagem
'Public sgPedeValorUnit As Boolean
'Public igNumViag As Integer
'Public sgCodLibPcary As String
'Public igCodAtrel As Integer
'Public sgDatAprVei As String
'Public sgDatIniCga As String
'Public sgDatFimCga As String
'Public igNumCTRC As Integer
'Public igTipDoc As Integer
'Public igAplic As Integer
'Public lgNumViag As Long
'Public sgTipViag As String
'Public igCodRota As Integer
'Public igSeqRota As Integer
'Public igCodRotaFre As Integer
'Public igSeqRotaFre As Integer
'Public igQtdCTRC As Integer
'Public igCodRotaFreBald As Integer
'Public igSeqRotaFreBald As Integer
'Public dgValAdcKM As Double
'Public dgValAdcKMFre As Double
'Public sgNomFrac As String
'Public sgTipFrac As String
'Public sgRecPedagio As String
'Public dgValTrfViag As String
'Public dgValTrfPedViag As Double
'Public sgLblRotaRec As String
'Public sgLblRotaPag As String
'Public sgCNPJDestFrac As String
'Public bgViagem As Boolean
