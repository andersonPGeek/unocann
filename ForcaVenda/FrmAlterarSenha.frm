VERSION 5.00
Begin VB.Form FrmAlterarSenha 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Alterar Senha"
   ClientHeight    =   1770
   ClientLeft      =   3870
   ClientTop       =   3210
   ClientWidth     =   3420
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   3420
   Begin VB.TextBox TxtRedNewSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   810
      Width           =   1380
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   450
      TabIndex        =   3
      Top             =   1215
      Width           =   975
   End
   Begin VB.CommandButton Cmdsair 
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1688
      TabIndex        =   4
      Top             =   1215
      Width           =   975
   End
   Begin VB.TextBox TxtNewSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   442
      Width           =   1380
   End
   Begin VB.TextBox TxtSenhaAtu 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   75
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Redigite Nova Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   870
      Width           =   1830
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nova Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   495
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Senha Atual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   135
      Width           =   1050
   End
End
Attribute VB_Name = "FrmAlterarSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pwdcript As String

Private Sub cmdConfirmar_Click()

    On Error GoTo TratarErro
    
    Dim slDatCadUsuSis As String
    Dim slSenUsuSis As String
    
    If TxtSenhaAtu.Text = "" Then
        
        MsgBox "Digite a Senha Atual", vbInformation
        
        TxtSenhaAtu.SetFocus
        
        Exit Sub
        
    ElseIf Len(TxtNewSenha.Text) < 6 Then
        
        MsgBox "Senha Menor do que 6 Caracteres", vbInformation
        
        TxtNewSenha.SetFocus
        
        Exit Sub
        
    ElseIf Not IsNumeric(TxtNewSenha.Text) Then
        
        MsgBox "Senha deve conter somente números", vbInformation
        
        TxtNewSenha.SetFocus
        
        Exit Sub
        
    ElseIf TxtNewSenha.Text = "" Then
    
        MsgBox "Digite a Nova Senha", vbInformation
        
        TxtNewSenha.SetFocus
        
        Exit Sub
        
    ElseIf TxtRedNewSenha.Text = "" Then
        
        MsgBox "Redigite a Nova Senha", vbInformation
        
        TxtRedNewSenha.SetFocus
        
        Exit Sub
        
    ElseIf TxtSenhaAtu.Text = TxtNewSenha.Text Then
        
        MsgBox "Nova Senha é Igual a Senha Atual", vbInformation
        
        TxtNewSenha.Text = ""
        TxtRedNewSenha.Text = ""
        TxtNewSenha.SetFocus
        
        Exit Sub
        
    ElseIf TxtNewSenha.Text = TxtRedNewSenha.Text Then
        
        sgQuery = "SELECT PWDCOMPARE('" & Trim(TxtSenhaAtu.Text) & "',SenUsu, 0) AS Senha_OK from usuario where codusu = " & LgCodUsuSis
        
        consulta sgQuery
        
        If Not Rs.EOF Then
        
            If Rs("Senha_OK") = 0 Then
                
                MsgBox "Senha Atual Inválida", vbInformation
                
                TxtSenhaAtu.Text = ""
                TxtSenhaAtu.SetFocus
                
                Exit Sub
                
            Else
            
                sgQuery = "update usuario set SenUsu = convert(varbinary(100),PWDENCRYPT('" & Trim(TxtNewSenha.Text) & "')),DatUltAce = convert(datetime,'" & Date & "',103) where codusu = " & LgCodUsuSis
                
                Set Rs = Conexao.Execute(sgQuery)
                Set Rs = Nothing
                
                MsgBox "Senha Alterada Com Sucesso", vbInformation
                
                Unload FrmAcessar
                Unload Me
                
                MDIProjUNO.Show
                
            End If
            
        Else
            
            MsgBox "Usuário não Cadastrado", vbInformation
            
        End If
        
    Else
        
        MsgBox "Campos Nova Senha e Redigite Nova Senha São Diferentes", vbInformation
        
        TxtRedNewSenha.SetFocus
        TxtRedNewSenha.Text = ""
        
    End If
    
    Exit Sub

TratarErro:
   
    Rotina_Erro "cmdConfirmar_Click"
    
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub LimpaCampos()
    
    TxtSenhaAtu.Text = ""
    TxtNewSenha.Text = ""
    TxtRedNewSenha.Text = ""
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
    End If
    
End Sub

Private Sub Form_Load()
    
    ajustajanela Me, 3810, 2760, 3000, 3500
    
End Sub

