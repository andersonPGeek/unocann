VERSION 5.00
Begin VB.Form Frm_Parametros 
   Caption         =   "Parametros gerais do sistema."
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Frame Frame2 
      BackColor       =   &H80000018&
      Caption         =   "Classifica todos os clientes como Ativo exceto"
      Height          =   1935
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4575
      Begin VB.CheckBox Check4 
         BackColor       =   &H80000018&
         Caption         =   "Clientes Cancelados"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H80000018&
         Caption         =   "Clientes em desenvolvimento"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Classifica com Inativo os clientes que"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   4575
      Begin VB.PictureBox ActiveDate1 
         Height          =   315
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   1875
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Não Reclassifica os clientes em desenvolvimento"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   3975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Não Reclassifica os clientes Cancelados"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Não compram desde:"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Frm_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rst_Cliente        As rdoResultset

Private Sub Bto_Alterar_Click()
    Sl_Desc = "select * from Tba_Clientes "
    'where Cli_codigo = '" & Sg_CodCli & "'"
    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    While Rst_Cliente.EOF = False
           Rst_Cliente.Edit
           Rst_Cliente!Cli_CodigoStatus = 1
           Rst_Cliente!Cli_Status = Date + Time
           Rst_Cliente.Update
           Rst_Cliente.MoveNext
    Wend
    Rst_Cliente.Close
End Sub

