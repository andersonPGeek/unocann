VERSION 5.00
Object = "{BEACF734-D8AC-11D7-9B57-000B6A03449D}#2.0#0"; "Masked.ocx"
Begin VB.Form FrmGeraChave 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Chave de autorização para Desconto diferenciado"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5640
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5640
   Begin VB.CommandButton BtoSair 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4440
      Picture         =   "GeraChave.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Width           =   4335
      Begin VB.Label LblSenha 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4095
      End
   End
   Begin Project_Masked.Masked MskSenha 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      FormatoString   =   "000000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
      ValInteiro      =   7
      MaxLength       =   7
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decripta"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdGerar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gerar Chave"
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
      Left            =   3000
      MaskColor       =   &H00E0E0E0&
      Picture         =   "GeraChave.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin Project_Masked.Masked MskDesc 
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      FormatoString   =   "#0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12582912
      ValInteiro      =   2
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Desconto Concedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "FrmGeraChave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Chave de criptografia
Const KRYPTOKEY = "AJfH7CXuD1K5LE9TODCVNeUrfPw3pn0bG1piVhdfp4Ma8ZFIosR6udS2gf" 'esta chave pode ser qualquer coisa, letras numeros simbolos ai depende!!! =)
    
'-------------------------------------------------------------------------------------
' Procedure : Encripty
' DateTime  : 16/11/2006 16:03
' Author    : Raphael Leão Taveira
' Purpose   : Encriptar textos p/ seg de database
'---------------------------------------------------------------------------------------

Public Function Encripty(lpValue As String) As String
    
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'Declaração de variaveis
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    Dim lpTemp As Integer
    Dim lpLenghtOfKey As Integer
    Dim lpLenghtofValue As Integer
    Dim lpKeyAsc() As Integer
    Dim lpValueAsc() As Integer
    Dim i As Integer
    Dim J As Integer
    
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'Pegando o valor do lenghtkey
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    lpLenghtOfKey = Len(KRYPTOKEY)
            
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'pegando o valor do lenghtvalue
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    lpLenghtofValue = Len(lpValue)
        
    If lpLenghtofValue < 1 Then
        Exit Function
    End If
            
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'redimensiona o keyasc
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=
    ReDim lpKeyAsc(0 To lpLenghtOfKey)
                
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'redimensiona o valor
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    ReDim lpValueAsc(0 To lpLenghtofValue)
                    
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'grava o valor a ser encriptado
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    For i = 0 To lpLenghtofValue - 1
        lpValueAsc(i) = Asc(Mid(lpValue, i + 1, 1))
    Next i
                
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'Grava o valor da chave
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    For i = 0 To lpLenghtOfKey - 1
        lpKeyAsc(i) = Asc(Mid(KRYPTOKEY, i + 1, 1))
    Next i
                
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'Algoritmo para encriptar o texto
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
        
    For i = 0 To lpLenghtofValue - 1
                
        If J = lpLenghtOfKey - 1 Then
            J = 0
        Else
            J = J + 1
        End If
                  
        If lpValueAsc(i) + lpKeyAsc(J) > 255 Then
            lpTemp = (lpValueAsc(i) + lpKeyAsc(J)) - 255
        Else
            lpTemp = (lpValueAsc(i) + lpKeyAsc(J))
        End If
                    
        Encripty = Encripty & Chr(lpTemp)
            
    Next i
                
    Exit Function

    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'Tratamento de erros
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
   
ErrHand:

    MsgBox "Failed to encript text using the specified key!", vbCritical
        
End Function

'---------------------------------------------------------------------------------------
' Procedure : Decripty
' DateTime  : 16/11/2006 15:40
' Author    : Raphael Leão Taveira
' Purpose   : Desencriptar Textos usando a chave especificada
'---------------------------------------------------------------------------------------
'
Public Function Decripty(lpValue As String) As String

    On Error GoTo ErrHand

    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'declaração de variaveis
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    Dim lpTemp As Integer
    Dim lpLenghtOfKey As Integer
    Dim lpLenghtofValue As Integer
    Dim lpKeyAsc() As Integer
    Dim lpValueAsc() As Integer
    Dim i As Integer
    Dim J As Integer
    Dim EncryptedValue() As String
        
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'pegando o valor do lenghtkey
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    lpLenghtOfKey = Len(KRYPTOKEY)
        
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'pegando o valor do lenghtvalue
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    lpLenghtofValue = Len(lpValue)
        
    If lpLenghtofValue < 1 Then
        Exit Function
    End If
                
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'redimensiona o keyasc
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-==-=-=-=-=-=-=-=-=-=-=
    ReDim lpKeyAsc(0 To lpLenghtOfKey)
                
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'redimensiona o valor
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    ReDim lpValueAsc(0 To lpLenghtofValue)
            
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'grava o valor a ser encriptado
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    For i = 0 To lpLenghtofValue - 1
        lpValueAsc(i) = Asc(Mid(lpValue, i + 1, 1))
    Next i
           
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'grava o valor da chave
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    For i = 0 To lpLenghtOfKey - 1
        lpKeyAsc(i) = Asc(Mid(KRYPTOKEY, i + 1, 1))
    Next i
            
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'algoritmo para desencriptar o texto
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    For i = 0 To lpLenghtofValue - 1
                
        If J = lpLenghtOfKey - 1 Then
            J = 0
        Else
            J = J + 1
        End If
                    
        If lpValueAsc(i) - lpKeyAsc(J) < 0 Then
            lpTemp = (lpValueAsc(i) - lpKeyAsc(J)) + 255
        Else
            lpTemp = (lpValueAsc(i) - lpKeyAsc(J))
        End If
                    
        Decripty = Decripty & Chr(lpTemp)
            
    Next i
            
    Exit Function
        
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
    'Tratamento de erros
    '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

ErrHand:

    MsgBox "Failed to decript text using the specified key!", vbCritical
        
End Function

Private Sub BtoSair_Click()

    Unload Me
    
    Set FrmGeraChave = Nothing

End Sub

Private Sub CmdGerar_Click()
    
    Dim resultado As String
    Dim final As String
    Dim tamanho As Integer
    Dim Senha As String
    Dim Desconto As String
    Dim i As Integer
    
    '******************************************************************************************
    'Zera as variáveia a serem utilizadas.
    '******************************************************************************************
    
    resultado = ""
    final = ""
    Senha = ""
    Desconto = ""
    tamanho = 0
    i = 0

    '******************************************************************************************
    'Consiste o número do pedido.
    '******************************************************************************************

    If Trim(MskSenha.Texto) = "" Or Trim(MskSenha.Texto) = 0 Then
        
        MsgBox "Informe o Número do Pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        MskSenha.SetFocus
        
        Exit Sub
        
    End If
    
    '******************************************************************************************
    'Consiste o desconto concedido.
    '******************************************************************************************
    
    If Trim(MskDesc.Texto) = "" Or Trim(MskDesc.Texto) = 0 Then
        
        MsgBox "Informe o Desconto concedido", vbExclamation + vbOKOnly, "Atenção!"
        
        MskDesc.SetFocus
        
        Exit Sub
        
    End If

    '******************************************************************************************
    '
    '******************************************************************************************

    Desconto = Trim(MskDesc.Texto)
    
    If Len(Desconto) < 2 Then
        Desconto = "00" & Desconto
        Desconto = Right(Desconto, 2)
    End If

    '******************************************************************************************
    '
    '******************************************************************************************

    resultado = Trim(MskSenha.Texto)
    
    If Len(resultado) < 6 Then
        resultado = "00000" & resultado
        resultado = Right(resultado, 6)
    End If

    '******************************************************************************************
    'Levanta a quantidade de caracteres do número do pedido.
    '******************************************************************************************
    
    tamanho = Len(Trim(resultado))
    
    '******************************************************************************************
    'Inverte os dois últimos números do número do pedido.
    '******************************************************************************************
    
    final = Mid(Trim(resultado), tamanho, 1) & Mid(Trim(resultado), tamanho - 1, 1)
    
    '******************************************************************************************
    'Multiplica o número do pedido pelo resultado da soma do final do pedido invertido mais o
    'desconto concedido. O resultado desse cálculo é convertido em hexadecimal.
    '******************************************************************************************
    
    resultado = Hex(resultado * Val((Val(final) + Val(Desconto))))
    
    '******************************************************************************************
    'Armazena o novo tamanho da chave após a criação do hexadecimal.
    '******************************************************************************************
    
    tamanho = Len(Trim(resultado))
    
    '******************************************************************************************
    'Inverte o hexadecimal, reposicionando os caracteres de trás pra frente.
    '******************************************************************************************
    
    Senha = ""

    For i = tamanho To 1 Step -1
        Senha = Senha & Mid(resultado, i, 1)
    Next i
    
    '******************************************************************************************
    'Escreve o primeiro caracter do final do pedido invertido, aplica o hexadecimal invertido e
    'finaliza com o segundo caracter do final do pedido invertido e o desconto concedido.
    '******************************************************************************************
    
    Senha = Mid(final, 1, 1) & Senha & Mid(final, 2, 1) & Desconto
    
    '******************************************************************************************
    'Exibe a chave.
    '******************************************************************************************

    LblSenha.Caption = Senha

End Sub

Private Sub Command2_Click()
    
    Dim resultado As String
    Dim final As String
    Dim tamanho As Integer
    Dim Senha As String
    Dim Desconto As String
    Dim i As Integer

    resultado = ""
    final = ""
    Senha = ""
    Desconto = 0
    tamanho = 0
    i = 0

    tamanho = Len(Trim(LblSenha))
    Senha = Trim(LblSenha)
    Desconto = Val(Right(Senha, 2))
    Senha = Mid(Senha, 1, tamanho - 2)
    tamanho = Len(Senha)
    final = Mid(Senha, 1, 1) & Mid(Senha, tamanho, 1)
    Senha = Trim(Mid(Senha, 2, tamanho - 1))
    Senha = Mid(Senha, 1, tamanho - 2)
    tamanho = Len(Senha)

    For i = tamanho To 1 Step -1
        resultado = resultado & Mid(Senha, i, 1)
    Next i

    resultado = "&H" & resultado
    resultado = resultado / Val((Val(final) + Desconto))

    MskSenha.Texto = resultado

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Call EventoEnter(KeyAscii)
    
End Sub

Private Sub Form_Load()

    Me.Left = 0
    Me.Top = 0
    Me.Height = 4845
    Me.Width = 5760

    MskSenha.TipodeDados numero
    
End Sub

Private Sub MskDesc_GotFocus()
    
    Call SelecionaTudo
    
End Sub

Private Sub MskSenha_GotFocus()
    
    Call SelecionaTudo
    
End Sub
