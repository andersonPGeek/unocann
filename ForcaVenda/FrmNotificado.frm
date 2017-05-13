VERSION 5.00
Object = "{368CC970-FF03-11D7-9B5A-000B6A03449D}#1.1#0"; "Combo_DB.ocx"
Object = "{F454059D-91FE-11D2-8865-AD1268A0A52F}#2.0#0"; "ActiveDate.ocx"
Begin VB.Form FrmNotificado 
   Caption         =   "Premiação"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   12780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPesquisa 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pesquisa de Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   12360
      Begin VB.CommandButton BtoSair 
         BackColor       =   &H00C0FFFF&
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
         Height          =   1005
         Left            =   10560
         Picture         =   "FrmNotificado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton CmdAtualizar 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Atualizar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   9000
         Picture         =   "FrmNotificado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin rdActiveDate.ActiveDate ActDtfim 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin rdActiveDate.ActiveDate ActDtini 
         Height          =   315
         Left            =   60
         TabIndex        =   0
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
      End
      Begin Project_Combo_DB.Combo_DB CboRemet 
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   600
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   661
         Cols            =   0
         Cabecalho       =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data Inicial"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data Final"
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Representante"
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Width           =   1410
      End
   End
   Begin VB.PictureBox LstPedidos 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   0
      ScaleHeight     =   6555
      ScaleWidth      =   12780
      TabIndex        =   5
      Top             =   1800
      Width           =   12840
   End
   Begin VB.Label lblTotalPed 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   13
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Total de Pedidos :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   12
      Top             =   8880
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total 10% :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   11
      Top             =   8400
      Width           =   1695
   End
   Begin VB.Label lblTotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11290
      TabIndex        =   10
      Top             =   8400
      Width           =   1215
   End
End
Attribute VB_Name = "FrmNotificado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
