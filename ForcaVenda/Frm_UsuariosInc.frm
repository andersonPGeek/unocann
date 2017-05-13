VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frm_UsuariosInc 
   Caption         =   "Inclusão de usuarios."
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   MDIChild        =   -1  'True
   ScaleHeight     =   3075
   ScaleWidth      =   7185
   Begin VB.TextBox Txt_senususis 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5655
      Begin MSMask.MaskEdBox Msk_codususis 
         Height          =   400
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   714
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_nomususis 
         Height          =   400
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   5
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Lbl_codususis 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Lbl_nomususis 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Nome "
         Height          =   195
         Left            =   1440
         TabIndex        =   6
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status"
      Height          =   1215
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
      Begin VB.CheckBox chk_bloqueado 
         Caption         =   "Bloqueado"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Chk_habilitado 
         Caption         =   "Habilitado"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSMask.MaskEdBox msk_datultace 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Lbl_senususis 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Último Acesso"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
End
Attribute VB_Name = "Frm_UsuariosInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaração de variáveis de trabalho
Dim Rst_Grupo_Usuario             As rdoResultset
Dim Rst_Usuarios                  As rdoResultset
Dim grupos_de_usuarios()

'dim Sp_Usuarios                   As rdoPreparedStatement
Dim limite_array                  As Integer



Private Sub Bto_grupos_Click()
'   si_codususis2 = Lg_CodUsuSis
'   Frm_usuario_grupo.Show
End Sub

Private Sub Bto_ok_Click()
                                                                                                                                                                       
'    'On Error GoTo Trata_Erro
    
        If Msk_codususis.ClipText = "" Then
           MsgBox "Informe o código do usuário.", vbExclamation, "A T E N Ç Ã O"
           Msk_codususis.SetFocus
           Exit Sub
        End If
            
        If Trim$(txt_nomususis.Text) = "" Then
           MsgBox "Informe o nome do usuário.", vbExclamation, "A T E N Ç Ã O"
           txt_nomususis.SetFocus
           Exit Sub
        End If
   
        If Trim$(Txt_senususis.Text) = "" Then
           MsgBox "Informe a senha do usuário.", vbExclamation, "A T E N Ç Ã O"
           Txt_senususis.SetFocus
           Exit Sub
        End If
           
        If Len(LTrim(Mid(Txt_senususis.Text, 1, 5))) < 5 Then
           MsgBox "Senha deve conter, no mínimo, 5 caracteres.", vbExclamation, "A T E N Ç Ã O"
           Txt_senususis.SetFocus
           Exit Sub
        End If
               
        lg_codususis = Msk_codususis
        sg_nomususis = txt_nomususis.Text
        sg_senususis = Txt_senususis.Text
        
        Sl_Desc = "SELECT *"
        Sl_Desc = Sl_Desc & " FROM tba_usuarios"
        Sl_Desc = Sl_Desc & " where codususis = " & lg_codususis
        Set Rst_Usuarios = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
        If Rst_Usuarios.EOF = False Then
           Rst_Usuarios.Edit
        Else
           Rst_Usuarios.AddNew
        End If
        Rst_Usuarios("CodUsuSis") = lg_codususis
        Rst_Usuarios("NomUsuSis") = UCase(sg_nomususis)
        If Trim(sg_senususis) <> "" Then
           Rst_Usuarios("SenUsuSis") = UCase(sg_senususis)
        End If
        If Chk_habilitado.Value = 1 Then
           Rst_Usuarios("SitUsuSis") = "H"
        Else
           Rst_Usuarios("SitUsuSis") = "D"
        End If
        If chk_bloqueado.Value = 1 Then
           Rst_Usuarios("FlgBlqSis") = "S"
        Else
           Rst_Usuarios("FlgBlqSis") = "N"
        End If
        Rst_Usuarios.Update
        Rst_Usuarios.Close

        Unload Me

 
Exit Sub
Trata_Erro:
'    Cn.RollbackTrans
    Rotina_erro ("A") ' chama rotina de erro genérica
    Beep
End Sub



Private Sub Bto_Sair_Click()
   Unload Me
End Sub


Private Sub Form_Load()
   Left = 350
   Top = 330
   Height = 4710
   Width = 8535


 
 Frm_Usuarios.Enabled = False
 
 
   
' ReDim grupos_de_usuarios(Rst_Grupo_Usuario(0), 2)
' limite_array = Rst_Grupo_Usuario(0) - 1
  
 msk_datultace.Enabled = False
   
 If flag_Usuarios = "I" Then
    Msk_codususis.Enabled = True
    'Bto_grupos.Enabled = False
    Chk_habilitado.Value = 1
    Chk_habilitado.Enabled = False
    chk_bloqueado.Value = 0
    chk_bloqueado.Enabled = False
 Else
    Msk_codususis.Enabled = False
    'Bto_grupos.Enabled = True
    Chk_habilitado.Enabled = True
    If sg_sitususis = "H" Then
       Chk_habilitado.Value = 1
    Else
       Chk_habilitado.Value = 0
    End If
    If sg_flgblqsis = "S" Then
       chk_bloqueado.Value = 1
    Else
       chk_bloqueado.Value = 0
    End If
    If sg_datultace <> "" Then
       msk_datultace = Format(sg_datultace, "dd/mm/yyyy")
    End If
 End If

 If flag_Usuarios = "I" Then
    Frm_UsuariosInc.Caption = "Incluir Usuário"
 Else
     Frm_UsuariosInc.Caption = "Alterar Usuário"
     Msk_codususis.Text = lg_codususis
     txt_nomususis.Text = sg_nomususis
'     Txt_senususis.Text = sg_senususis
 End If
 
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Screen.MousePointer = vbHourglass
   Frm_Usuarios.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Msk_codususis_GotFocus()
   Call Posiciona_cursor(Msk_codususis)
End Sub
Private Sub txt_nomususis_GotFocus()
   Call Posiciona_cursor(txt_nomususis)
End Sub

Private Sub txt_senususis_GotFocus()
   Call Posiciona_cursor(Txt_senususis)
End Sub

Private Sub txt_senususis_LostFocus()
    Call Criptografa(Txt_senususis)
End Sub


