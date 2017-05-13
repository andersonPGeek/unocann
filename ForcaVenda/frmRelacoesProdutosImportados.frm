VERSION 5.00
Begin VB.Form frmRelacaoProdutosImportados 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relações entre produtos importados"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   8550
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame fraProdutosRelacionados 
      BackColor       =   &H80000018&
      Caption         =   "Produtos Relacionados"
      Height          =   2055
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   5895
      Begin VB.ListBox lstProdutosRelacionados 
         Height          =   1425
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdRelacionar 
      Caption         =   "Relacionar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtProdutoPolyvin 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   5055
   End
   Begin VB.TextBox txtCodigoProdutoPolyvin 
      Height          =   315
      Left            =   360
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtProdutoUnocann 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
   Begin VB.TextBox txtCodigoProdutoUnocann 
      Height          =   315
      Left            =   360
      MaxLength       =   3
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Line linSeparador 
      X1              =   6600
      X2              =   6600
      Y1              =   240
      Y2              =   3960
   End
   Begin VB.Label lblProdutoPolyvin 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Produto Polyvin"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   1110
   End
   Begin VB.Label lblProdutoUnocann 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Produto Unocann"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "frmRelacaoProdutosImportados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDados As New ADODB.Recordset
Private mSQL As String

Private Sub cmdLimpar_Click()

    LimparCampos

End Sub

Private Sub cmdRelacionar_Click()
    
    If txtCodigoProdutoUnocann.Text = "" Then
        
        MsgBox "Informe o código do produto Unocann.", vbCritical, "Força de Venda"
        
        txtCodigoProdutoUnocann.SetFocus
        
        Exit Sub
        
    ElseIf txtCodigoProdutoPolyvin.Text = "" Then
        
        MsgBox "Informe o código do produto Polyvin.", vbCritical, "Força de Venda"
        
        txtCodigoProdutoPolyvin.SetFocus
        
        Exit Sub
        
    End If
    
    mSQL = "SELECT * FROM ProdutosImportadosPolyvin WHERE CodProdutoUnocann = " & txtCodigoProdutoUnocann.Text
    
    mDados.Open mSQL, Conexao, adOpenDynamic, adLockPessimistic
    
    If mDados.EOF = False Then
        
        MsgBox "Já existe uma relação para o produto Unocann.", vbCritical, "Força de Venda"
        
        txtCodigoProdutoUnocann.SetFocus
        
        Exit Sub
        
    End If
    
    mDados.Close
    
    mSQL = "SELECT * FROM ProdutosImportadosPolyvin WHERE CodProdutoPolyvin = " & txtCodigoProdutoPolyvin.Text
    
    mDados.Open mSQL, Conexao, adOpenDynamic, adLockPessimistic
    
    If mDados.EOF = False Then
        
        MsgBox "Já existe uma relação para o produto Polyvin.", vbCritical, "Força de Venda"
        
        txtCodigoProdutoPolyvin.SetFocus
        
        Exit Sub
        
    End If
    
    mDados.Close
    
    mSQL = "INSERT INTO ProdutosImportadosPolyvin (CodProdutoUnocann, CodProdutoPolyvin, DataCadastro) VALUES (" & txtCodigoProdutoUnocann.Text & ", " & txtCodigoProdutoPolyvin.Text & ", '" & Now & "')"
    
    Conexao.Execute mSQL
    
    MsgBox "Relação estabelecida com sucesso!", vbInformation, "Força de Venda"
    
    LimparCampos
    CarregarProdutosRelacionados
    
End Sub

Private Sub CmdSair_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Me.Left = (MDIProjUNO.Width - Me.Width) / 2
    Me.Top = (MDIProjUNO.Height - Me.Height) / 4
    
    CarregarProdutosRelacionados
    
End Sub

Private Sub txtCodigoProdutoPolyvin_GotFocus()
    
    With txtCodigoProdutoPolyvin
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub txtCodigoProdutoPolyvin_KeyPress(KeyAscii As Integer)

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtCodigoProdutoPolyvin_LostFocus()
    
    If txtCodigoProdutoPolyvin.Text <> "" Then
        txtProdutoPolyvin.Text = ConsultarProduto(txtCodigoProdutoPolyvin)
    Else
        txtProdutoPolyvin.Text = ""
    End If
    
End Sub

Private Sub txtCodigoProdutoUnocann_GotFocus()
    
    With txtCodigoProdutoUnocann
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub txtCodigoProdutoUnocann_KeyPress(KeyAscii As Integer)

    If KeyAscii < 48 Or KeyAscii > 57 Then
        
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
        
    End If

End Sub

Private Sub txtCodigoProdutoUnocann_LostFocus()
    
    If txtCodigoProdutoUnocann.Text <> "" Then
        txtProdutoUnocann.Text = ConsultarProduto(txtCodigoProdutoUnocann)
    Else
        txtProdutoUnocann.Text = ""
    End If
    
End Sub

Private Sub CarregarProdutosRelacionados()

    mSQL = "SELECT PIP.*, PU.DscPrd As ProdutoUnocann, PP.DscPrd As ProdutoPolyvin FROM ProdutosImportadosPolyvin PIP INNER JOIN Produto PU ON PU.CodPrd = PIP.CodProdutoUnocann INNER JOIN Produto PP ON PP.CodPrd = PIP.CodProdutoPolyvin"
    
    mDados.Open mSQL, Conexao, adOpenDynamic, adLockPessimistic
        
    With lstProdutosRelacionados
    
        .Clear
        
        Do While mDados.EOF = False
            
            .AddItem Format(mDados("CodProdutoUnocann"), "000") & " - " & mDados("ProdutoUnocann")
            .AddItem Format(mDados("CodProdutoPolyvin"), "000") & " - " & mDados("ProdutoPolyvin")
            .AddItem " "
            
            mDados.MoveNext
            
        Loop
        
    End With
    
    mDados.Close

End Sub

Private Function ConsultarProduto(pControle As TextBox)
    
    mSQL = "SELECT DscPrd FROM Produto WHERE CodPrd = " & pControle.Text
    
    mDados.Open mSQL, Conexao, adOpenDynamic, adLockPessimistic
    
    If mDados.EOF = False Then
        
        ConsultarProduto = mDados("DscPrd")
        
    Else
        
        ConsultarProduto = ""
        
        MsgBox "Não há produto com o código informado.", vbCritical, "Força de Venda"
        
        pControle.SetFocus
        
    End If
    
    mDados.Close
    
End Function

Private Sub LimparCampos()

    txtCodigoProdutoUnocann.Text = ""
    txtCodigoProdutoPolyvin.Text = ""
    txtProdutoUnocann.Text = ""
    txtProdutoPolyvin.Text = ""
    txtCodigoProdutoUnocann.SetFocus

End Sub
