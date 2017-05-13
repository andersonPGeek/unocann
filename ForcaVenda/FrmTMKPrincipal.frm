VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{368CC970-FF03-11D7-9B5A-000B6A03449D}#1.1#0"; "Combo_DB.ocx"
Object = "{BEACF734-D8AC-11D7-9B57-000B6A03449D}#2.0#0"; "Masked.ocx"
Begin VB.Form FrmTMKPrincipal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "U N O C A N N  -  Telemarketing"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12720
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   12720
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   7680
      TabIndex        =   73
      Top             =   7560
      Width           =   4935
      Begin VB.CommandButton BtoLimpaCTRC 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Limpar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2520
         Picture         =   "FrmTMKPrincipal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton BtoPesquisar 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Pesquisar Ligações"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         Picture         =   "FrmTMKPrincipal.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton BtoSimular 
         BackColor       =   &H0080C0FF&
         Caption         =   "Si&mular Pedido"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1320
         MaskColor       =   &H00E0E0E0&
         Picture         =   "FrmTMKPrincipal.frx":088C
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton BtoSair 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Sair"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3720
         Picture         =   "FrmTMKPrincipal.frx":16DE
         Style           =   1  'Graphical
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   4320
      TabIndex        =   70
      Top             =   7560
      Width           =   3015
      Begin VB.CommandButton BtoFimLigacao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Finalizar Ligação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1560
         MaskColor       =   &H00E0E0E0&
         Picture         =   "FrmTMKPrincipal.frx":1B20
         Style           =   1  'Graphical
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton BtoLigacao 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Iniciar Ligação"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         Picture         =   "FrmTMKPrincipal.frx":262A
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   120
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab SSTPrincipal 
      Height          =   9015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   15901
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   794
      BackColor       =   14737632
      ForeColor       =   4194368
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "PixelPoint"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Contato"
      TabPicture(0)   =   "FrmTMKPrincipal.frx":347C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label24"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblSaldoCli"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FramLigacao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Framtipo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TxtMensa"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboCli"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Pedidos"
      TabPicture(1)   =   "FrmTMKPrincipal.frx":3498
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdPedido"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cliente"
      TabPicture(2)   =   "FrmTMKPrincipal.frx":34B4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "LblCliPrimeira"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "LblCliNome"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblCliEndereco"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "LblCliCGC"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "LblCliBairro"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "LblCliUF"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "LblCliContr"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "LblCliCid"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "LblCliInscr"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "LblCliCep"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label31"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label29"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label32"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label33"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label34"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label35"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label36"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label37"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Label38"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Label40"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "LblCliUltimo"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "LblSimBahia"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Label45"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "LblFone"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).ControlCount=   25
      TabCaption(3)   =   "Histórico"
      TabPicture(3)   =   "FrmTMKPrincipal.frx":34D0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CmdDN"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdUP"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "TxtHistMensa"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label46"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin Project_Combo_DB.Combo_DB cboCli 
         Height          =   3135
         Left            =   480
         TabIndex        =   0
         Top             =   2160
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   5530
         Cols            =   0
         Cabecalho       =   -1  'True
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   600
         Width           =   855
         Begin VB.Label Label4 
            Caption         =   "Ligação"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   495
            Left            =   0
            TabIndex        =   82
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   9480
         TabIndex        =   79
         Top             =   480
         Width           =   375
         Begin VB.Label Label2 
            BackColor       =   &H00FFECEC&
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   0
            TabIndex        =   80
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   4815
         Left            =   -74880
         TabIndex        =   53
         Top             =   2520
         Width           =   12735
         Begin MSFlexGridLib.MSFlexGrid GrdCliVencer 
            Height          =   1575
            Left            =   6600
            TabIndex        =   54
            Top             =   600
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedCols       =   0
            BackColor       =   12648384
            ForeColor       =   192
            BackColorFixed  =   16711680
            ForeColorFixed  =   16777215
            ForeColorSel    =   255
            BackColorBkg    =   12648384
            GridColor       =   65280
            GridColorFixed  =   65280
            FocusRect       =   2
            ScrollBars      =   2
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Helvetica"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid GrdCliVencidos 
            Height          =   1575
            Left            =   240
            TabIndex        =   55
            Top             =   600
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            BackColor       =   12648384
            ForeColor       =   192
            BackColorFixed  =   16711680
            ForeColorFixed  =   16777215
            ForeColorSel    =   255
            BackColorBkg    =   12648384
            GridColor       =   65280
            GridColorFixed  =   65280
            AllowBigSelection=   0   'False
            FocusRect       =   2
            ScrollBars      =   2
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Helvetica"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid GrdCliJuros 
            Height          =   1575
            Left            =   240
            TabIndex        =   56
            Top             =   3000
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   1
            Cols            =   18
            FixedCols       =   0
            BackColor       =   12648384
            ForeColor       =   192
            BackColorFixed  =   16711680
            ForeColorFixed  =   16777215
            ForeColorSel    =   255
            BackColorBkg    =   12648384
            GridColor       =   65280
            GridColorFixed  =   65280
            FocusRect       =   2
            ScrollBars      =   2
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Helvetica"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label LblSumJur 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3480
            TabIndex        =   65
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label LblQtdJur 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2880
            TabIndex        =   64
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label LblSumAVen 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   9840
            TabIndex        =   63
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label LblQtdAVen 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   9240
            TabIndex        =   62
            Top             =   120
            Width           =   495
         End
         Begin VB.Label LblSumVen 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3600
            TabIndex        =   61
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label LblQtdVen 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3000
            TabIndex        =   60
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label27 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Juros em aberto"
            BeginProperty Font 
               Name            =   "Helvetica"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   375
            Left            =   240
            TabIndex        =   59
            Top             =   2520
            Width           =   2295
         End
         Begin VB.Label Label26 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Títulos a Vencer"
            BeginProperty Font 
               Name            =   "Helvetica"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   375
            Left            =   6600
            TabIndex        =   58
            Top             =   120
            Width           =   2295
         End
         Begin VB.Label Label25 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Títulos Vencidos"
            BeginProperty Font 
               Name            =   "Helvetica"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   375
            Left            =   240
            TabIndex        =   57
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.CommandButton CmdDN 
         BackColor       =   &H00FFC0C0&
         Height          =   720
         Left            =   -63000
         Picture         =   "FrmTMKPrincipal.frx":34EC
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   6000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdUP 
         BackColor       =   &H00FFC0C0&
         Height          =   720
         Left            =   -63000
         Picture         =   "FrmTMKPrincipal.frx":392E
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame3 
         Height          =   1815
         Left            =   480
         TabIndex        =   16
         Top             =   2880
         Width           =   10095
         Begin VB.Label LblUFPri 
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   9360
            TabIndex        =   68
            Top             =   480
            Width           =   470
         End
         Begin VB.Label Label12 
            Caption         =   "UF"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   255
            Left            =   9360
            TabIndex        =   69
            Top             =   240
            Width           =   375
         End
         Begin VB.Label LblCidCli 
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   5040
            TabIndex        =   23
            Top             =   480
            Width           =   4095
         End
         Begin VB.Label LblRepre 
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   5040
            TabIndex        =   21
            Top             =   1200
            Width           =   4790
         End
         Begin VB.Label Label13 
            Caption         =   "Cidade"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   255
            Left            =   5040
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Representante"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   255
            Left            =   5040
            TabIndex        =   22
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label LblContatoPri 
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   4575
         End
         Begin VB.Label LblTelcli 
            BackColor       =   &H80000013&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   4575
         End
         Begin VB.Label Label8 
            Caption         =   "Telefone"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Contato"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtMensa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFECEC&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   480
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   5400
         Width           =   10335
      End
      Begin RichTextLib.RichTextBox TxtHistMensa 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   960
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   11668
         _Version        =   393217
         BackColor       =   16772332
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"FrmTMKPrincipal.frx":3D70
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid GrdPedido 
         Height          =   6015
         Left            =   -75000
         TabIndex        =   67
         Top             =   1440
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   1
         Cols            =   17
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   16772332
         ForeColorFixed  =   4194304
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   192
         GridColorFixed  =   128
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Framtipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFECEC&
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   9360
         TabIndex        =   3
         Top             =   600
         Width           =   3015
         Begin VB.OptionButton OptReceptivo 
            BackColor       =   &H00FFECEC&
            Caption         =   "&Receptivo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   1320
            TabIndex        =   4
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptAtivo 
            BackColor       =   &H00FFECEC&
            Caption         =   "&Ativo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame FramLigacao 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   9015
         Begin Project_Masked.Masked MskLigacao 
            Height          =   495
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
            Texto           =   "9999999"
            FormatoString   =   "0000000"
            CampoDb         =   "9999999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   4895630
            ForeColor       =   12582912
            ValInteiro      =   7
            MaxLength       =   7
            Texto           =   "9999999"
         End
         Begin VB.Label LblDatEmi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Left            =   2280
            TabIndex        =   9
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label LblDatFim 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Left            =   4200
            TabIndex        =   11
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label LblOperador 
            Appearance      =   0  'Flat
            BackColor       =   &H00404000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Left            =   6120
            TabIndex        =   13
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Operador(a)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   6120
            TabIndex        =   14
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Final"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   4200
            TabIndex        =   12
            Top             =   380
            Width           =   615
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Início"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   2280
            TabIndex        =   10
            Top             =   380
            Width           =   615
         End
      End
      Begin VB.Label LblSaldoCli 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   9360
         TabIndex        =   77
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label24 
         Caption         =   "Mensagem"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label46 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mensagens"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   340
         Left            =   -74880
         TabIndex        =   52
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label LblCliPrimeira 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "99/99/9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -65640
         TabIndex        =   30
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label LblCliNome 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nome do cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -74880
         TabIndex        =   48
         Top             =   720
         Width           =   6855
      End
      Begin VB.Label lblCliEndereco 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Endereço do cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -74880
         TabIndex        =   47
         Top             =   1320
         Width           =   6855
      End
      Begin VB.Label LblCliCGC 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "99.999.999/0001-99"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -67920
         TabIndex        =   46
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label LblCliBairro 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bairro do cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -74880
         TabIndex        =   45
         Top             =   1920
         Width           =   3855
      End
      Begin VB.Label LblCliUF 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -66000
         TabIndex        =   44
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label LblCliContr 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Não Contribuinte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -65520
         TabIndex        =   43
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label LblCliCid 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cidade do cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -70920
         TabIndex        =   42
         Top             =   1920
         Width           =   3375
      End
      Begin VB.Label LblCliInscr 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "inscrição do cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -67920
         TabIndex        =   41
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label LblCliCep 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "99999-999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -67440
         TabIndex        =   40
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "Endereço"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   135
         Left            =   -74880
         TabIndex        =   39
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label29 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   135
         Left            =   -74880
         TabIndex        =   38
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label32 
         Caption         =   "Bairro"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   135
         Left            =   -74880
         TabIndex        =   37
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label33 
         Caption         =   "Cidade"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   135
         Left            =   -70920
         TabIndex        =   36
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label34 
         Caption         =   "C.E.P"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   135
         Left            =   -67440
         TabIndex        =   35
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label35 
         Caption         =   "CNPJ/CPF"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -67920
         TabIndex        =   34
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label36 
         Caption         =   "Inscr.Estadual"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   135
         Left            =   -67920
         TabIndex        =   33
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label37 
         Caption         =   "UF"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   135
         Left            =   -66000
         TabIndex        =   32
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label38 
         Caption         =   "Primeira Compra"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -65640
         TabIndex        =   31
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label40 
         Caption         =   " Último Faturamento"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -64200
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label LblCliUltimo 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "99/99/9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -64200
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label LblSimBahia 
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Simples"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -63840
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label45 
         Caption         =   "Fone"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   135
         Left            =   -65280
         TabIndex        =   26
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label LblFone 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -65280
         TabIndex        =   25
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   0
         Left            =   9360
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Width           =   1050
      End
   End
End
Attribute VB_Name = "FrmTMKPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim slremet As String
Dim ilCodCli As Integer
Dim blleitura As Boolean
Dim slUFCli As String
Dim slFlgContr As String
Dim slFlgSIMBa As String
Dim ilFlgSit As Integer
Dim ilind As Integer
Dim blI As Integer
Dim dlSeqLig As Double
Dim ilTipLig As Integer

Function Activate_Mkt()

    On Error GoTo TratarErro

    MskLigacao.Texto = lgSeqLig
    
    sgQuery = "select a.SeqLig, a.DatIniLig, a.DatFimLig, a.tipLig, a.codcli, a.Msglig, b.NomUsu, c.NomCli"
    sgQuery = sgQuery & "  from LIGACAO a, USUARIO b, CLIENTE c"
    sgQuery = sgQuery & "  Where a.codusu = b.codusu"
    sgQuery = sgQuery & "    and a.codcli *= c.codcli"
    sgQuery = sgQuery & "    and a.seqlig = " & Trim(MskLigacao.Texto)
    
    Consulta2 sgQuery
    
    If Rs2.EOF Then
        
        MsgBox "Erro na leitura da Ligação - " & Format(Trim(MskLigacao.Texto), "000000"), vbExclamation + vbOKOnly, "Atenção!"
        
        lgSeqLig = 0
        
        Exit Function
        
    Else
        
        LblOperador = Trim(Rs2!nomusu)
        LblDatEmi.Caption = Format(Trim(Rs2!Datinilig), "dd/mm/yyyy hh:mm:ss")
        LblDatFim.Caption = Format(Trim(Rs2!Datfimlig), "dd/mm/yyyy hh:mm:ss")
        
        If Rs2!tiplig = 1 Then
            OptAtivo.Value = True
        Else
            OptReceptivo.Value = True
        End If
        
        TxtMensa.Text = Trim(Rs2!MsgLig)
        cboCli.Criterio = IIf(IsNull(Rs2!Codcli), "", Trim(Rs2!Codcli))
        
        cboCli_LostFocus
        
        TxtMensa.Locked = True
        cboCli.Habilitado = False
        Framtipo.Enabled = False
        'OptAtivo.Enabled = False
        'OptReceptivo.Enabled = False
        'BtoLigacao.SetFocus
    
    End If
    
    BtoLimpaCTRC.Enabled = True
    BtoLigacao.Enabled = False
    
    Exit Function

TratarErro:

    Rotina_Erro "Activate_Mkt"

End Function

Function CarregaMensagens()

    SSTPrincipal.TabEnabled(3) = False
    
    Set Rs = Nothing
    
    sgQuery = " select a.seqlig, a.datinilig, a.tiplig, a.codusu, a.msglig, b.nomusu from LIGACAO a, USUARIO b"
    sgQuery = sgQuery & " where  a.Codusu = b.codusu"
    
    If ilCodCli > 0 Then
        sgQuery = sgQuery & " and a.CodCli = " & Trim(ilCodCli)
    Else
        sgQuery = sgQuery & " and a.CodCli is null "
    End If
    
    sgQuery = sgQuery & " order by a.DatiniLig desc "
    
    Consulta sgQuery
    
    TxtHistMensa.Text = ""
    
    blI = 0
    
    Do While Not Rs.EOF
        
        'TxtHistMensa.Text = TxtHistMensa.Text & "..............................................................................................................................................................................................." & vbCrLf
        TxtHistMensa.Text = TxtHistMensa.Text & "_____________________________________________________________________________________________________________" & vbCrLf
        TxtHistMensa.Text = TxtHistMensa.Text & "   ¤ LIGAÇÃO " & Format(Rs!seqlig, "0000000") & " - " & IIf(Trim(Rs!tiplig) = 1, "Ativa ", "Receptiva ") & _
                                              "      [Operador(a) - " & Rs!nomusu & " - " & Format(Rs!Datinilig, "dd/mm/yyyy hh:mm:ss") & "]" & vbCrLf
        TxtHistMensa.Text = TxtHistMensa.Text & "¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯" & vbCrLf
      
        TxtHistMensa.Text = TxtHistMensa.Text & Trim(Rs!MsgLig) & vbCrLf
        TxtHistMensa.Text = TxtHistMensa.Text & vbCrLf
        
        Rs.MoveNext
        
        blI = blI + 1
        
    Loop
    
    Rs.Close
    
    Set Rs = Nothing
    
    If blI > 0 Then
        SSTPrincipal.TabEnabled(3) = True
    End If

End Function

Function LimpaGeral()

    TxtMensa.Text = ""
    TxtHistMensa.Text = ""
    
    MskLigacao.Texto = 0

    LblDatEmi.Caption = ""
    LblDatFim.Caption = ""
    OptReceptivo.Value = 0
    OptAtivo.Value = 0
    GrdPedido.Rows = 1
    GrdCliVencidos.Rows = 1
    GrdCliVencer.Rows = 1
    GrdCliJuros.Rows = 1
    LblContatoPri = ""
    LblCidCli = ""
    LblUFPri = ""
    LblTelcli = ""
    LblRepre = ""
    LblCliNome = ""
    lblCliEndereco = ""
    LblCliBairro = ""
    LblCliCid = ""
    LblCliCep = ""
    LblCliUF = ""
    LblFone = ""
    LblCliPrimeira = ""
    LblCliCGC.Caption = ""
    LblCliInscr = ""
    LblCliContr = ""
    LblSumVen.Caption = ""
    LblQtdVen.Caption = ""
    LblSumAVen.Caption = ""
    LblQtdAVen.Caption = ""
    LblSumJur.Caption = ""
    LblQtdJur.Caption = ""
    LblCliUltimo = ""
    LblCliPrimeira = ""
    SSTPrincipal.TabEnabled(1) = False
    SSTPrincipal.TabEnabled(2) = False
    SSTPrincipal.TabEnabled(3) = False
    cboCli.Habilitado = True
    OptAtivo.Enabled = True
    OptReceptivo.Enabled = True
    TxtMensa.Locked = False
    BtoLigacao.Enabled = True
    BtoSair.Enabled = True
    BtoPesquisar.Enabled = True
    BtoFimLigacao.Enabled = False
    BtoLimpaCTRC.Enabled = False
    Framtipo.Enabled = True
    BtoSimular.Caption = "Si&mular Pedido"
    BtoSimular.Enabled = False
    cboCli.Criterio = ""
    
    slremet = ""
    ilCodCli = 0

    DoEvents

End Function

Function LeituraCliente() As Boolean
    
    On Error GoTo TrataErro

    Dim slCGCCli  As String
    Dim dlSumVen  As Double
    Dim dlSumAVen As Double
    Dim dlSumJur  As Double
    Dim dlQtdVen  As Integer
    Dim dlQtdAVen As Integer
    Dim dlQtdJur  As Integer

    LeituraCliente = True
  
    'If blLeitura = True Then
        'Exit Function
    'End If
  
    blleitura = True
  
    'Leitura Cliente
    slUFCli = ""
    slFlgContr = ""
    slFlgSIMBa = ""
    ilFlgSit = 0
    dlSumVen = 0
    dlSumAVen = 0
    dlSumJur = 0
    dlQtdVen = 0
    dlQtdAVen = 0
    dlQtdJur = 0
    sgQuery = "SELECT a.*, b.NomRep from CLIENTE a, REPRESENTANTE b  WHERE CodCli = " & Trim(ilCodCli)
    sgQuery = sgQuery + " and a.codrep *= b.codrep "
    
    Call Consulta(sgQuery)
    
    If Rs.EOF Then
    
        MsgBox "Erro na leitura do cliente", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        Exit Function
        
    Else
        
        slUFCli = IIf(IsNull(Rs!UFCli), "", Trim(Rs!UFCli))
        slFlgContr = IIf(IsNull(Rs!FlgContr), "", Trim(Rs!FlgContr))
        slFlgSIMBa = IIf(IsNull(Rs!FlgSIMBa), "", Trim(Rs!FlgSIMBa))
        ilFlgSit = IIf(IsNull(Rs!FlgSit), 0, Trim(Rs!FlgSit))
        LblCliNome = IIf(IsNull(Rs!NomCli), "", Trim(Rs!NomCli))
        lblCliEndereco = IIf(IsNull(Rs!EndCli), "", Trim(Rs!EndCli))
        LblCliBairro = IIf(IsNull(Rs!BaiCli), "", Trim(Rs!BaiCli))
        LblCliCid = IIf(IsNull(Rs!CidCli), "", Trim(Rs!CidCli))
        LblCidCli = IIf(IsNull(Rs!CidCli), "", Trim(Rs!CidCli))
        LblCliCep = IIf(IsNull(Rs!CepCli), "", Trim(Rs!CepCli))
        LblCliUF = IIf(IsNull(Rs!UFCli), "", Trim(Rs!UFCli))
        LblUFPri = IIf(IsNull(Rs!UFCli), "", Trim(Rs!UFCli))
        LblFone = IIf(IsNull(Rs!FonCli), "", Trim(Rs!FonCli))
        LblTelcli = IIf(IsNull(Rs!FonCli), "", Trim(Rs!FonCli))
        LblContatoPri = IIf(IsNull(Rs!NomCtt), "", Trim(Rs!NomCtt))
        LblRepre = Trim(Rs!codrep) & " - " & IIf(IsNull(Rs!NomRep), "", Trim(Rs!NomRep))
        LblCliPrimeira = Format(Rs!DatPriComp, "dd/mm/yyyy")
        slCGCCli = IIf(IsNull(Rs!CgcCli), "", Trim(Rs!CgcCli))
     
        If Len(slCGCCli) = 14 Then
            LblCliCGC.Caption = Format(Mid(slCGCCli, 1, 8), "00,000,000") & "/" & Format(Mid(slCGCCli, 9, 4), "0000") & "-" & Mid(slCGCCli, 13, 2)
        Else
            LblCliCGC.Caption = slCGCCli
        End If
        
        LblCliInscr = IIf(IsNull(Rs!InsCli), "", Trim(Rs!InsCli))
        
        If slFlgContr = "S" Then
            LblCliContr = "Contribuinte"
        Else
            LblCliContr = "Não Contribuinte"
        End If
        
        If slFlgSIMBa = "S" Then
            LblSimBahia.Visible = True
        Else
            LblSimBahia.Visible = False
        End If
        
    End If
    
    Rs.Close
    
    Set Rs = Nothing
  
    'If sgFlagOper <> "A" Then
        'slPedSimples = ""
    'End If

    'If Trim(slPedSimples) = "" Then
        
        'slPedSimples = "N"
        
        'If slUFRep = "BA" And slFlgSIMBa = "S" Then
            'slPedSimples = "S"
        'End If
        
    'End If
  
    'If Trim(slPedSimples) = "S" Then
        
        'slUFOri = "BA"
        
        'LblVlSimples.Visible = True
        'LblSimples.Visible = True
        
    'Else
        
        'LblVlSimples.Visible = False
        'LblSimples.Visible = False
        
    'End If
    
    SSTPrincipal.TabEnabled(2) = True
  
    'Títulos Vencidos
    sgQuery = "select nrodup, parc, datemi, datven, vlrdup, datediff(dd,datven, getdate()) as dias from DUPLICATA "
    sgQuery = sgQuery + " where datpag is null and datven < (getdate() - 1)"
    sgQuery = sgQuery + "   AND CodCli = " & Trim(ilCodCli)

    Call Consulta(sgQuery)
    
    GrdCliVencidos.Rows = 1
    
    ilind = GrdCliVencidos.Rows

    Do While Not Rs.EOF
        
        GrdCliVencidos.Rows = GrdCliVencidos.Rows + 1
        GrdCliVencidos.TextMatrix(ilind, 0) = Format(Trim(Rs!NroDup), "000,000")
        GrdCliVencidos.TextMatrix(ilind, 1) = Trim(Rs!Parc)
        GrdCliVencidos.TextMatrix(ilind, 2) = Format(Trim(Rs!datemi), "dd/mm/yyyy")
        GrdCliVencidos.TextMatrix(ilind, 3) = Format(Trim(Rs!datven), "dd/mm/yyyy")
        GrdCliVencidos.TextMatrix(ilind, 4) = Format(Trim(Rs!VlrDup), "##,###,###,##0.00")
        GrdCliVencidos.TextMatrix(ilind, 5) = Format(Trim(Rs!dias), "##,##0")
        
        dlSumVen = dlSumVen + Trim(Rs!VlrDup)
        dlQtdVen = dlQtdVen + 1
        ilind = ilind + 1
        'blVencidos = True
  
        Rs.MoveNext
        
    Loop

    Rs.Close
    
    Set Rs = Nothing
    
    LblSumVen.Caption = Format(dlSumVen, "##,###,###,##0.00")
    LblQtdVen.Caption = Format(dlQtdVen, "00")
  
    'Títulos a Vencer
    sgQuery = "select nrodup, parc, datemi, datven, vlrdup from DUPLICATA "
    sgQuery = sgQuery + " where datpag is null and datven >= (getdate() - 1)"
    sgQuery = sgQuery + "   AND CodCli = " & Trim(ilCodCli)

    Call Consulta(sgQuery)
    
    GrdCliVencer.Rows = 1
    ilind = GrdCliVencer.Rows

    Do While Not Rs.EOF
        
        GrdCliVencer.Rows = GrdCliVencer.Rows + 1
        GrdCliVencer.TextMatrix(ilind, 0) = Format(Trim(Rs!NroDup), "000,000")
        GrdCliVencer.TextMatrix(ilind, 1) = Trim(Rs!Parc)
        GrdCliVencer.TextMatrix(ilind, 2) = Format(Trim(Rs!datemi), "dd/mm/yyyy")
        GrdCliVencer.TextMatrix(ilind, 3) = Format(Trim(Rs!datven), "dd/mm/yyyy")
        GrdCliVencer.TextMatrix(ilind, 4) = Format(Trim(Rs!VlrDup), "##,###,###,##0.00")
        
        dlSumAVen = dlSumAVen + Trim(Rs!VlrDup)
        dlQtdAVen = dlQtdAVen + 1
        ilind = ilind + 1
  
        Rs.MoveNext
        
    Loop

    Rs.Close
    
    Set Rs = Nothing

    LblSumAVen.Caption = Format(dlSumAVen, "##,###,###,##0.00")
    LblQtdAVen.Caption = Format(dlQtdAVen, "00")
  
    'Juros em Aberto
    sgQuery = "select nrodup, parc, datpag, datven, vlrdup, JurDev from DUPLICATA "
    sgQuery = sgQuery + " where datjur is null and JurDev > 0 "
    sgQuery = sgQuery + "   AND CodCli = " & Trim(ilCodCli)

    Call Consulta(sgQuery)
    
    GrdCliJuros.Rows = 1
    
    ilind = GrdCliJuros.Rows

    Do While Not Rs.EOF
        
        GrdCliJuros.Rows = GrdCliJuros.Rows + 1
        GrdCliJuros.TextMatrix(ilind, 0) = Format(Trim(Rs!NroDup), "000,000")
        GrdCliJuros.TextMatrix(ilind, 1) = Trim(Rs!Parc)
        GrdCliJuros.TextMatrix(ilind, 2) = Format(Trim(Rs!datven), "dd/mm/yyyy")
        GrdCliJuros.TextMatrix(ilind, 3) = Format(Trim(Rs!datpag), "dd/mm/yyyy")
        GrdCliJuros.TextMatrix(ilind, 4) = Format(Trim(Rs!JurDev), "##,###,###,##0.00")
        
        dlSumJur = dlSumJur + Trim(Rs!JurDev)
        dlQtdJur = dlQtdJur + 1
        ilind = ilind + 1
  
        Rs.MoveNext
        
    Loop

    Rs.Close
    
    Set Rs = Nothing
  
    LblSumJur.Caption = Format(dlSumJur, "##,###,###,##0.00")
    LblQtdJur.Caption = Format(dlQtdJur, "00")
    
    'Último Faturamento
    sgQuery = "select max(Datemi) as ultimo from DUPLICATA "
    sgQuery = sgQuery + "  where CodCli = " & Trim(ilCodCli)

    Call Consulta(sgQuery)
 
    LblCliUltimo.Caption = ""
    
    If Not Rs.EOF Then
        LblCliUltimo.Caption = Format(Rs!ultimo, "dd/mm/yyyy")
    End If
  
    Rs.Close
    
    Set Rs = Nothing
    
    Exit Function

TrataErro:
    
    Rotina_Erro "LeituraCliente"
    
    LeituraCliente = False

End Function

Function CompoeGridPed() As Boolean

    Dim blI As Integer
    
    On Error GoTo TratarErro
    
    'Pedidos não Liberados
    sgQuery = "select a.nroped, a.datped, a.codcli,c.NomCli, sum(b.vlrite) - sum(distinct a.vlrsimples) as Valor,a.codcnd,a.codrep, d.DscCnd,"
    sgQuery = sgQuery & "       a.NomTra , a.codrep, a.sitPed, a.datlib, a.datenv, TipNot, NroNot, dateminot, a.flgalt, e.nomrep, a.SeqLig"
    sgQuery = sgQuery & "  from Pedido a, Item_pedido b, Cliente c, Condicao d, representante e"
    sgQuery = sgQuery & "  Where (a.datlib is null or a.flgalt = 'N')"
    sgQuery = sgQuery & "    and a.nroped = b.nroped"
    sgQuery = sgQuery & "    and a.codcli = c.codcli"
    sgQuery = sgQuery & "    and a.codcnd = d.codcnd"
    sgQuery = sgQuery & "    and a.codrep = e.codrep"

    'If MskNroPedido.Texto > 0 Then
        'sgQuery = sgQuery & "    and a.nroped = " & Trim(MskNroPedido.Texto)
    'End If
    
    If APLICA = 1 Then
        sgQuery = sgQuery & "    and a.codrep = " & sgRepresentante
    End If
    
    'If Trim(ActDtini.Text) <> "" Then
        sgQuery = sgQuery & "    and a.datped between (getdate() - 180) and Getdate()"
        'sgQuery = sgQuery & "    and a.datped between convert(datetime,'" & Trim(ActDtini.Text) & "',103)"
        'sgQuery = sgQuery & "                     and convert(datetime,'" & Trim(ActDtfim.Text) & "',103)"
    'End If
    
    If ilCodCli <> 0 Then
        sgQuery = sgQuery & "    and a.codcli = " & ilCodCli
    End If
    
    'If ChkNotif.Value = 1 And MskNroPedido.Texto = 0 Then
        'sgQuery = sgQuery & " and a.flgalt is not null "
    'End If

    sgQuery = sgQuery & "    group by a.nroped, a.datped, a.codcli, c.NomCli, a.codcnd,a.codrep, d.DscCnd,"
    sgQuery = sgQuery & "             a.NomTra, a.codrep, a.sitPed, a.datlib, a.datenv,"
    sgQuery = sgQuery & "             a.TipNot , a.NroNot, a.dateminot, a.flgalt, e.nomrep, a.SeqLig"
    sgQuery = sgQuery & "    order by 2 desc, 1 desc"

    Consulta sgQuery
    
    'If blLoad = False And Rs.RecordCount > 50 Then
        'MsgBox "Resultado retornou mais de 50 linhas." & Chr(13) & "Favor refazer seu filtro de pesquisa.", vbExclamation + vbOKOnly, "Atenção!"
    'End If

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
        'GrdPedido.TextMatrix(blI, 6) = Rs!NomTra
        GrdPedido.TextMatrix(blI, 7) = IIf(Trim(Rs!SitPed) = "U", "C", Trim(Rs!SitPed))
        GrdPedido.TextMatrix(blI, 8) = Rs!NomRep
        'GrdPedido.TextMatrix(blI, 9) = IIf(IsNull(Rs!DatEnv), "", Format(Trim(Rs!DatEnv), "dd/mm/yyyy"))
        GrdPedido.TextMatrix(blI, 10) = IIf(IsNull(Rs!NroNot), "", Rs!TipNot & Format(Trim(Rs!NroNot), "000000"))
        GrdPedido.TextMatrix(blI, 11) = IIf(IsNull(Rs!DatEmiNot), "", Format(Trim(Rs!DatEmiNot), "dd/mm/yyyy"))
        GrdPedido.TextMatrix(blI, 12) = Rs!codrep
        GrdPedido.TextMatrix(blI, 13) = IIf(IsNull(Rs!FlgAlt), "", Rs!FlgAlt)
        GrdPedido.TextMatrix(blI, 16) = IIf(IsNull(Rs!seqlig), "", Format(Trim(Rs!seqlig), "000000"))
        GrdPedido.Row = blI
        GrdPedido.Col = 0
        GrdPedido.CellForeColor = &HFFFF&
        
        If Trim(Rs!FlgAlt) = "L" Then
            
            GrdPedido.CellBackColor = &H800080
            
        Else
            
            If Trim(Rs!FlgAlt) = "N" Then
                GrdPedido.CellBackColor = &H40C0&
                GrdPedido.CellForeColor = &HFFFFFF
            Else
                GrdPedido.CellBackColor = &HFF&
            End If
            
        End If
   
        GrdPedido.Col = 15
        GrdPedido.CellBackColor = &HFFFFFF
        
        If Not IsNull(Rs!DatEmiNot) And Trim(Rs!SitPed) <> "U" And Trim(Rs!SitPed) <> "C" Then
            
            sgQuery = " select isnull(count(*),0) as conta from"
            sgQuery = sgQuery & "     (Select a.codprd, a.qtdprd, a.qtdprdfat + isnull(d.sum_saldo_entregue,0) as totentreg"
            sgQuery = sgQuery & "     From ITEM_PEDIDO a, pedido c,"
            sgQuery = sgQuery & "          (select a.codprd, sum_saldo_entregue = sum(a.qtdprdfat) from item_pedido_saldo a, pedido_saldo b"
            sgQuery = sgQuery & "             Where a.NroPed = " & Trim(Rs!NroPed)
            sgQuery = sgQuery & "               and a.NroPed = b.nroped"
            sgQuery = sgQuery & "               and a.NroPedsdo = b.nropedsdo"
            sgQuery = sgQuery & "               and b.SitPed = 'N'"
            sgQuery = sgQuery & "               group by a.codprd) d"
            sgQuery = sgQuery & "      Where a.NroPed = " & Trim(Rs!NroPed)
            sgQuery = sgQuery & "        and a.nroped = c.NroPed"
            sgQuery = sgQuery & "        and a.codprd *= d.codprd) a"
            sgQuery = sgQuery & "         where a.qtdprd > a.totentreg "
            
            Consulta2 sgQuery
            
            If Not Rs2.EOF Then
                
                If Rs2!conta > 0 Then
                    GrdPedido.TextMatrix(blI, 15) = "S"
                    GrdPedido.CellForeColor = &HFFFF&
                    GrdPedido.CellBackColor = &HFF&
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
    sgQuery = sgQuery & "       a.NomTra , a.codrep, a.sitPed, a.datlib, a.datenv, TipNot, NroNot, dateminot, a.flgalt, e.nomrep, a.SeqLig"
    sgQuery = sgQuery & "  from Pedido a, Item_pedido b, Cliente c, Condicao d, representante e"
    
    'If ChkNotif.Value = 1 And MskNroPedido.Texto = 0 Then
        'sgQuery = sgQuery & " Where a.datlib is not null and a.FlgAlt is not null and a.FlgAlt <> 'N' "
    'Else
        sgQuery = sgQuery & "  Where a.datlib is not null and (a.FlgAlt <> 'N' or a.flgalt is null) "
    'End If

    sgQuery = sgQuery & "    and a.nroped = b.nroped"
    sgQuery = sgQuery & "    and a.codcli = c.codcli"
    sgQuery = sgQuery & "    and a.codcnd = d.codcnd"
    sgQuery = sgQuery & "    and a.codrep = e.codrep"
    
    'If MskNroPedido.Texto > 0 Then
        'sgQuery = sgQuery & "    and a.nroped = " & Trim(MskNroPedido.Texto)
    'End If
    
    If APLICA = 1 Then
        sgQuery = sgQuery & "    and a.codrep = " & sgRepresentante
    End If
    
    'If Trim(ActDtini.Text) <> "" Then
        sgQuery = sgQuery & "    and a.datped between (getdate() - 180) and Getdate()"
        'sgQuery = sgQuery & "    and a.datped between convert(datetime,'" & Trim(ActDtini.Text) & "',103)"
        'sgQuery = sgQuery & "                     and convert(datetime,'" & Trim(ActDtfim.Text) & "',103)"
    'End If
    
    If ilCodCli <> 0 Then
        sgQuery = sgQuery & "    and a.codcli = " & ilCodCli
    End If
    
    sgQuery = sgQuery & "    group by a.nroped, a.datped, a.codcli, c.NomCli, a.codcnd,a.codrep, d.DscCnd,"
    sgQuery = sgQuery & "             a.NomTra, a.codrep, a.sitPed, a.datlib, a.datenv,"
    sgQuery = sgQuery & "             a.TipNot , a.NroNot, a.dateminot, a.flgalt, e.nomrep, a.SeqLig"
    sgQuery = sgQuery & "    order by 2 desc, 1 desc"

    Consulta sgQuery

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
        'GrdPedido.TextMatrix(blI, 6) = Rs!NomTra
        GrdPedido.TextMatrix(blI, 7) = IIf(Trim(Rs!SitPed) = "U", "C", Trim(Rs!SitPed))
        GrdPedido.TextMatrix(blI, 8) = Rs!NomRep
        'GrdPedido.TextMatrix(blI, 9) = IIf(IsNull(Rs!DatEnv), "", Format(Trim(Rs!DatEnv), "dd/mm/yyyy"))
        GrdPedido.TextMatrix(blI, 10) = IIf(IsNull(Rs!NroNot), "", Rs!TipNot & Format(Trim(Rs!NroNot), "000000"))
        GrdPedido.TextMatrix(blI, 11) = IIf(IsNull(Rs!DatEmiNot), "", Format(Trim(Rs!DatEmiNot), "dd/mm/yyyy"))
        GrdPedido.TextMatrix(blI, 12) = Rs!codrep
        GrdPedido.TextMatrix(blI, 13) = IIf(IsNull(Rs!FlgAlt), "", Rs!FlgAlt)
        GrdPedido.TextMatrix(blI, 16) = IIf(IsNull(Rs!seqlig), "", Format(Trim(Rs!seqlig), "000000"))
        GrdPedido.Row = blI

        'If IsNull(Rs!Datlib) And IsNull(Rs!Datlib) And IsNull(Rs!DatEmiNot) Then
            
            'GrdPedido.Col = 0
            'GrdPedido.CellBackColor = &HFF&
            'GrdPedido.CellForeColor = &HFFFF&
            'GrdPedido.Col = 8
            'GrdPedido.CellBackColor = vbRed
            'GrdPedido.Col = 9
            'GrdPedido.CellBackColor = vbRed
            'GrdPedido.Col = 10
            'GrdPedido.CellBackColor = vbRed
            'GrdPedido.Col = 11
            'GrdPedido.CellBackColor = vbRed
            
        'End If
        
        If Not IsNull(Rs!DatEnv) And IsNull(Rs!DatEmiNot) Then
            
            GrdPedido.Col = 0
            GrdPedido.CellBackColor = vbYellow
            'GrdPedido.Col = 8
            'GrdPedido.CellBackColor = vbYellow
            'GrdPedido.Col = 9
            'GrdPedido.CellBackColor = vbYellow
            'GrdPedido.Col = 10
            'GrdPedido.CellBackColor = vbYellow
            'GrdPedido.Col = 11
            'GrdPedido.CellBackColor = vbYellow
            
        End If
        
        If Not IsNull(Rs!DatEmiNot) Then
            
            GrdPedido.Col = 0
            GrdPedido.CellBackColor = &HFF00&
            'GrdPedido.Col = 8
            'GrdPedido.CellBackColor = &HC000&
            'GrdPedido.Col = 9
            'GrdPedido.CellBackColor = &HC000&
            'GrdPedido.Col = 10
            'GrdPedido.CellBackColor = &HC000&
            'GrdPedido.Col = 11
            'GrdPedido.CellBackColor = &HC000&
   
        End If
        
        If Trim(Rs!SitPed) = "C" Or Trim(Rs!SitPed) = "U" Then
            GrdPedido.Col = 0
            'GrdPedido.CellBackColor = &HFFFFFF
            GrdPedido.CellBackColor = &H80000012
            GrdPedido.CellForeColor = &HFFFF&
        End If
   
        If Trim(Rs!FlgAlt) = "A" Then
            
            GrdPedido.TextMatrix(blI, 14) = "A"
            GrdPedido.Col = 14
            GrdPedido.CellBackColor = &H800080
            GrdPedido.CellForeColor = &HFFFFFF
            
        Else
            
            If Trim(Rs!FlgAlt) = "O" Then
                GrdPedido.TextMatrix(blI, 14) = "N"
                GrdPedido.Col = 14
                GrdPedido.CellBackColor = &H40C0&
                GrdPedido.CellForeColor = &HFFFFFF
            End If
            
        End If
        
        GrdPedido.Col = 15
        GrdPedido.CellBackColor = &HFFFFFF
        
        If Not IsNull(Rs!DatEmiNot) And Trim(Rs!SitPed) <> "U" And Trim(Rs!SitPed) <> "C" Then
            
            sgQuery = " select isnull(count(*),0) as conta from"
            sgQuery = sgQuery & "     (Select a.codprd, a.qtdprd, a.qtdprdfat + isnull(d.sum_saldo_entregue,0) as totentreg"
            sgQuery = sgQuery & "     From ITEM_PEDIDO a, pedido c,"
            sgQuery = sgQuery & "          (select a.codprd, sum_saldo_entregue = sum(a.qtdprdfat) from item_pedido_saldo a, pedido_saldo b"
            sgQuery = sgQuery & "             Where a.NroPed = " & Trim(Rs!NroPed)
            sgQuery = sgQuery & "               and a.NroPed = b.nroped"
            sgQuery = sgQuery & "               and a.NroPedsdo = b.nropedsdo"
            sgQuery = sgQuery & "               and b.SitPed = 'N'"
            sgQuery = sgQuery & "               group by a.codprd) d"
            sgQuery = sgQuery & "      Where a.NroPed = " & Trim(Rs!NroPed)
            sgQuery = sgQuery & "        and a.nroped = c.NroPed"
            sgQuery = sgQuery & "        and a.codprd *= d.codprd) a"
            sgQuery = sgQuery & "         where a.qtdprd > a.totentreg "
            
            Consulta2 sgQuery
            
            If Not Rs2.EOF Then
                
                If Rs2!conta > 0 Then
                    GrdPedido.TextMatrix(blI, 15) = "S"
                    GrdPedido.CellForeColor = &HFFFF&
                    GrdPedido.CellBackColor = &HFF&
                End If
                
            End If
            
            Rs2.Close
            
            Set Rs2 = Nothing
            
        End If
   
        Rs.MoveNext
        
    Loop

    Rs.Close
    
    'blLoad = False
    
    Set Rs = Nothing
    
    GrdPedido.Visible = True

    DoEvents

    If GrdPedido.Rows > 1 Then
        SSTPrincipal.TabEnabled(1) = True
    End If

    Exit Function

TratarErro:

    Rotina_Erro "CompoeGridPed"

End Function

Private Sub BtoFimLigacao_Click()
    
    Dim slMensa As String
  
    slMensa = Trim(TxtMensa.Text)
    slMensa = Replace(slMensa, "'", "´")
    slMensa = Replace(slMensa, """", "§")
    slMensa = Replace(slMensa, "§", "´")
    sgQuery = "update LIGACAO set DatFimLig = getdate(), MsgLig = '" & Trim(slMensa) & "'"
    sgQuery = sgQuery & " where SeqLig = " & Trim(MskLigacao.Texto)
    
    Conexao.Execute sgQuery

    LimpaGeral
    
End Sub

Private Sub BtoLigacao_Click()
    
    Dim slCodCli As String
    
    If OptAtivo.Value = False And OptReceptivo = False Then
        
        MsgBox "Informe o tipo de ligação.", vbExclamation + vbOKOnly, "Atenção!"
        
        Exit Sub
        
    End If
    
    If OptAtivo.Value = True Then
        ilTipLig = 1
    Else
        ilTipLig = 2
    End If
    
    If Trim(cboCli.Codigo) = "" Then
        slCodCli = "null"
    Else
        slCodCli = Trim(cboCli.Codigo)
    End If
    
    If Trim(slremet) = "" Or Trim(cboCli.Codigo) = "" Then
    
        sgQuery = MsgBox("Deseja selecionar um cliente ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Atenção!")
        
        If sgQuery = vbYes Then
            
            slremet = ""
            ilCodCli = 0
            
            cboCli.Habilitado = True
            cboCli.SetFocus
            
            Exit Sub
            
        End If
        
    End If
    
    cboCli.Habilitado = False
    Framtipo.Enabled = False
    
    'OptAtivo.Enabled = False
    'OptReceptivo.Enabled = False
      
    sgQuery = "insert into ligacao"
    sgQuery = sgQuery & " select isnull(max(seqlig),0) + 1, getdate(), null, "
    sgQuery = sgQuery & LgCodUsuSis & "," & Trim(slCodCli) & "," & ilTipLig & ", '" & Trim(TxtMensa.Text) & "' from ligacao"
    
    Conexao.Execute sgQuery
    
    sgQuery = "select seqlig, datinilig from LIGACAO "
    sgQuery = sgQuery & " Where CodUsu = " & Trim(LgCodUsuSis)
    sgQuery = sgQuery & "   and SeqLig = (select max(seqlig) from LIGACAO Where CodUsu = " & Trim(LgCodUsuSis) & ")"
    
    Call Consulta(sgQuery)
    
    If Rs.EOF Then
        
        MsgBox "ERRO na geração do número da ligação.", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        Exit Sub
        
    End If
    
    dlSeqLig = Rs!seqlig
    
    LblDatEmi.Caption = Format(Rs!Datinilig, "dd/mm/yyyy hh:mm:ss")
    MskLigacao.Texto = dlSeqLig
    BtoPesquisar.Enabled = False
    BtoLigacao.Enabled = False
    BtoSair.Enabled = False
    BtoFimLigacao.Enabled = True
    BtoLimpaCTRC.Enabled = False
    BtoSimular.Enabled = True
    
End Sub

Private Sub BtoLimpaCTRC_Click()
    
    Dim slCliente As String
    
    slCliente = Trim(cboCli.Criterio)
   
    LimpaGeral
   
    'If lgSeqLig > 0 Then
        cboCli.Criterio = Trim(slCliente)
    'End If
   
End Sub

Private Sub BtoPesquisar_Click()

    igCodCli = 0
    
    If Trim(cboCli.Codigo) = "" Then
        igCodCli = 0
    Else
        igCodCli = Trim(cboCli.Codigo)
    End If

    FrmPosiLig.Show
    
End Sub

Private Sub BtoSair_Click()
    
    Unload Me
    
    Set FrmTMKPrincipal = Nothing
 
    DoEvents
 
    Unload Me
    
    Set FrmConhecimento = Nothing

End Sub

Private Sub BtoSimular_Click()

    bgSimula = False
    igCodCli = 0
    
    If Trim(cboCli.Codigo) = "" Then
        igCodCli = 0
    Else
        igCodCli = Trim(cboCli.Codigo)
    End If
  
    Me.Enabled = False
    bgPedMKT = True
    
    If igCodCli = 0 Then
        bgSimula = True
    End If
    
    sgRepresentante = 9999
    igTela = "PedidoMKT"
    lgSeqLig = Trim(MskLigacao.Texto)
  
    FrmConhecimento.Show

End Sub

Private Sub cboCli_Consultar()
    
    slremet = ""
    
    cboCli.query = "Select NomCli As Cliente, CodCli As Código, CgcCli as CNPJ, FlgContr As Contribuinte From Cliente Where " & IIf(IsNumeric(cboCli.Criterio), "CodCli", "NomCli") & " Like '" & cboCli.Criterio & "%' order by " & IIf(IsNumeric(cboCli.Criterio), "CodCli", "NomCli")
    
    slremet = Trim(cboCli.Criterio)
    
End Sub

Private Sub cboCli_GotFocus()
    
    Call SelecionaTudo
    
    SSTPrincipal.TabEnabled(1) = False
    SSTPrincipal.TabEnabled(2) = False
    
End Sub

Private Sub cboCli_LostFocus()
    
    If Me.ActiveControl.Name = "BtoSair" Then
        Exit Sub
    End If
    
    If Trim(cboCli.Codigo) > 0 And Trim(cboCli.Codigo) <> "" Then
        
        slremet = cboCli.Criterio
        ilCodCli = cboCli.Codigo
        
        LeituraCliente
        CompoeGridPed
        
        BtoSimular.Caption = "Inc&luir Pedido"
        
    Else
        
        BtoLimpaCTRC_Click
        
        SSTPrincipal.TabEnabled(1) = False
        SSTPrincipal.TabEnabled(2) = False
        
        BtoSimular.Caption = "Si&mular Pedido"
        
    End If
    
    CarregaMensagens
    
End Sub

Private Sub CmdDN_Click()
    
    Dim Num&
    Num& = ScrollText&(TxtHistMensa, 1)

End Sub

Private Sub cmdUP_Click()
    
    Dim Num&
    Num& = ScrollText&(TxtHistMensa, -1)

End Sub

Private Sub Form_Activate()
    
    SSTPrincipal.Tab = 0
    
    bgPedMKT = False
    bgSimula = False
    
    If igTela = "PedidoMKT" Then
        
        igTela = ""
        bgPosLig = False
        
        Exit Sub
        
    End If
    
    If bgPosLig = True Then
        
        If lgSeqLig > 0 Then
            Activate_Mkt
        Else
            BtoLimpaCTRC_Click
        End If
        
        bgPosLig = False
        
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Call EventoEnter(KeyAscii)
    
End Sub

Private Sub Form_Load()

    Me.Left = 0
    Me.Top = 0
    Me.Height = 9495
    Me.Width = 12810

    Set cboCli.Conexao = Conexao
   
    SSTPrincipal.TabEnabled(1) = False
    SSTPrincipal.TabEnabled(2) = False
    SSTPrincipal.TabEnabled(3) = False
    LblOperador.Caption = sgNomUsuSis
    GrdCliVencidos.TextMatrix(0, 0) = " Duplicata"
    GrdCliVencidos.ColWidth(0) = 900
    GrdCliVencidos.TextMatrix(0, 1) = " Parc."
    GrdCliVencidos.ColWidth(1) = 500
    GrdCliVencidos.TextMatrix(0, 2) = " Dt. Emissão"
    GrdCliVencidos.ColWidth(2) = 1100
    GrdCliVencidos.TextMatrix(0, 3) = "  Dt. Vencto."
    GrdCliVencidos.ColWidth(3) = 1100
    GrdCliVencidos.TextMatrix(0, 4) = "     Valor"
    GrdCliVencidos.ColWidth(4) = 1100
    GrdCliVencidos.TextMatrix(0, 5) = "Dias"
    GrdCliVencidos.ColWidth(5) = 600
    GrdCliVencidos.TextMatrix(0, 6) = ""
    GrdCliVencidos.ColWidth(6) = 600
    GrdCliVencer.TextMatrix(0, 0) = " Duplicata"
    GrdCliVencer.ColWidth(0) = 900
    GrdCliVencer.TextMatrix(0, 1) = " Parc."
    GrdCliVencer.ColWidth(1) = 500
    GrdCliVencer.TextMatrix(0, 2) = " Dt. Emissão"
    GrdCliVencer.ColWidth(2) = 1100
    GrdCliVencer.TextMatrix(0, 3) = "  Dt. Vencto."
    GrdCliVencer.ColWidth(3) = 1100
    GrdCliVencer.TextMatrix(0, 4) = "     Valor"
    GrdCliVencer.ColWidth(4) = 1100
    GrdCliVencer.TextMatrix(0, 5) = ""
    GrdCliVencer.ColWidth(5) = 600
    GrdCliJuros.TextMatrix(0, 0) = " Duplicata"
    GrdCliJuros.ColWidth(0) = 900
    GrdCliJuros.TextMatrix(0, 1) = " Parc."
    GrdCliJuros.ColWidth(1) = 500
    GrdCliJuros.TextMatrix(0, 2) = "  Dt. Vencto."
    GrdCliJuros.ColWidth(2) = 1100
    GrdCliJuros.TextMatrix(0, 3) = "  Dt. Pagto."
    GrdCliJuros.ColWidth(3) = 1100
    GrdCliJuros.TextMatrix(0, 4) = "Valor Juros"
    GrdCliJuros.ColWidth(4) = 1100
    GrdCliJuros.TextMatrix(0, 5) = ""
    GrdCliJuros.ColWidth(5) = 600
    GrdPedido.TextMatrix(0, 0) = "Pedido"
    GrdPedido.ColWidth(0) = 650
    GrdPedido.TextMatrix(0, 1) = "Dt.Emissão"
    GrdPedido.ColWidth(1) = 800
    GrdPedido.TextMatrix(0, 2) = "Cod.Cli"
    GrdPedido.ColWidth(2) = 500
    GrdPedido.TextMatrix(0, 3) = "Cliente"
    GrdPedido.ColWidth(3) = 3000
    GrdPedido.TextMatrix(0, 4) = "Val.Pedido"
    GrdPedido.ColWidth(4) = 800
    GrdPedido.TextMatrix(0, 5) = "Cond.Pagto"
    GrdPedido.ColWidth(5) = 1650
    GrdPedido.TextMatrix(0, 6) = ""
    GrdPedido.ColWidth(6) = 0
    GrdPedido.TextMatrix(0, 7) = "Sit."
    GrdPedido.ColWidth(7) = 0
    GrdPedido.TextMatrix(0, 8) = "Representante"
    GrdPedido.ColWidth(8) = 1900
    GrdPedido.TextMatrix(0, 9) = "Pr.Entrega"
    GrdPedido.ColWidth(9) = 800
    GrdPedido.TextMatrix(0, 10) = "N.Fiscal"
    GrdPedido.ColWidth(10) = 700
    GrdPedido.TextMatrix(0, 11) = "Dt.Fatur."
    GrdPedido.ColWidth(11) = 800
    GrdPedido.TextMatrix(0, 12) = ""
    GrdPedido.ColWidth(12) = 0
    GrdPedido.TextMatrix(0, 13) = ""
    GrdPedido.ColWidth(13) = 0
    GrdPedido.TextMatrix(0, 14) = ""
    GrdPedido.ColWidth(14) = 190
    GrdPedido.TextMatrix(0, 15) = ""
    GrdPedido.ColWidth(15) = 180
    GrdPedido.TextMatrix(0, 16) = "Ligação"
    GrdPedido.ColWidth(16) = 600
   
    LimpaGeral
    
End Sub

Private Sub GrdPedido_DblClick()
    
    bgBloqPed = False
    
    If GrdPedido.RowSel = 0 Then
        Exit Sub
    End If
    
    bgConsultaPed = True
    Me.Enabled = False
    igNroPed = GrdPedido.TextMatrix(GrdPedido.RowSel, 0)
    sgRepresentante = Trim(GrdPedido.TextMatrix(GrdPedido.RowSel, 12))
    bgBloqPed = True
    igTela = "PedidoMKT"
  
    FrmConhecimento.Show
    
End Sub

Private Sub SSTPrincipal_Click(PreviousTab As Integer)
    
    DoEvents
    
    'If SSTPrincipal.TabEnabled(1) = False Then
        'SSTPrincipal.Tab = 0
    'End If
  
End Sub

Private Sub SSTPrincipal_KeyPress(KeyAscii As Integer)
    
    Call EventoEnter(KeyAscii)
    
End Sub
