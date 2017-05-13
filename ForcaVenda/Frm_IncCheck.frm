VERSION 5.00
Begin VB.Form Frm_IncCheck 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Inclui opções na ficha."
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   152
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   6255
   End
   Begin VB.TextBox Txt_Descricao 
      Height          =   525
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Line Line4 
      BorderWidth     =   4
      X1              =   6555
      X2              =   6555
      Y1              =   3840
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderWidth     =   4
      X1              =   30
      X2              =   30
      Y1              =   3960
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   0
      X2              =   6600
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   0
      X2              =   6600
      Y1              =   30
      Y2              =   30
   End
End
Attribute VB_Name = "Frm_IncCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rst_Ficha         As rdoResultset

Private Sub bto_Incluir_Click()
    If Trim(Txt_Descricao) = "" Then
       MsgBox ("Descrição da ficha não foi informada.")
       Exit Sub
    End If
    
    If Sg_Flag = "I" Then
       Sl_Desc = " SELECT * FROM Tba_Ficha "
       Set Rst_Ficha = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
       Rst_Ficha.AddNew
           Rst_Ficha!Fix_CodigoFilho = Mid(sg_Proximo_codigoFilho, 2, Len(sg_Proximo_codigoFilho) - 2)
           Rst_Ficha!Fix_CodigoPai = Mid(Sg_CodigoPai, 2, Len(Sg_CodigoPai) - 2)
           Rst_Ficha!Fix_Descricao = Txt_Descricao
           Rst_Ficha!Fix_CodUsuSis = Sg_Usuario
           Sg_Descricao = Txt_Descricao
       Rst_Ficha.Update
       Rst_Ficha.Close
    End If
    If Sg_Flag = "A" Then
       Sl_Desc = " SELECT * FROM Tba_Ficha "
       Sl_Desc = Sl_Desc & " where Fix_CodigoFilho = '" & Mid(Sg_CodigoFilho, 2, Len(Sg_CodigoFilho) - 2) & "'"
       Set Rst_Ficha = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
       If Rst_Ficha.EOF = False Then
          If Rst_Ficha!Fix_CodUsuSis <> Val(Sg_Usuario) Then
             MsgBox ("Alteração não permitida.")
             Rst_Ficha.Close
             Sg_Flag = ""
             Unload Me
             Exit Sub
          Else
             Rst_Ficha.Edit
 
             Rst_Ficha!Fix_Descricao = Txt_Descricao
             Sg_Descricao = Txt_Descricao
             Rst_Ficha.Update
          End If
       End If
       Rst_Ficha.Close
    End If
    
    
    
    
    Frm_IncCliente.Enabled = True
    Unload Me
 
End Sub

Private Sub Bto_Sair_Click()
    Frm_IncCliente.Enabled = True
    Sg_Flag = ""
    Unload Me
End Sub

Private Sub Form_Load()
    Left = 3000
    Top = 2000
    Height = 3870
    Width = 6580
    Text2.Text = Sg_Texto
    Text2.Refresh
    If Sg_Flag = "A" Then
       Txt_Descricao = Sg_Descricao
       Txt_Descricao.Refresh
    End If
End Sub
