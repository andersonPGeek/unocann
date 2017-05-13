VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{368CC970-FF03-11D7-9B5A-000B6A03449D}#1.1#0"; "Combo_DB.ocx"
Object = "{BEACF734-D8AC-11D7-9B57-000B6A03449D}#2.0#0"; "Masked.ocx"
Begin VB.Form frmCadastro_Geral 
   Caption         =   "Cadastro Geral"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   11130
   Begin VB.CommandButton btoLimpar 
      Caption         =   "&Limpar"
      Height          =   495
      Left            =   4185
      TabIndex        =   64
      Top             =   6615
      Width           =   975
   End
   Begin VB.CommandButton btoSair 
      Caption         =   "&Sair"
      Height          =   495
      Left            =   5505
      TabIndex        =   65
      Top             =   6615
      Width           =   975
   End
   Begin VB.CommandButton btoExcluir 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2865
      TabIndex        =   63
      Top             =   6615
      Width           =   975
   End
   Begin VB.CommandButton btoAlterar 
      Caption         =   "&Alterar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1545
      TabIndex        =   62
      Top             =   6615
      Width           =   975
   End
   Begin VB.CommandButton btoIncluir 
      Caption         =   "&Incluir"
      Height          =   495
      Left            =   225
      TabIndex        =   61
      Top             =   6615
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7185
      Left            =   45
      TabIndex        =   71
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   12674
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "D&ados Gerais"
      TabPicture(0)   =   "frmCadastro_Geral.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCNPJ"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label12"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Atividade"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtRazSoc"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraTipCli"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtNomRed"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtInscrMun"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkRetISS"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkINSS"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtHomePage"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtEmail"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtCNPJ"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtInscrEst"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "CboDesAtvGer"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "&Endereço/Cobrança"
      TabPicture(1)   =   "frmCadastro_Geral.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboCidade(1)"
      Tab(1).Control(1)=   "cboUF(1)"
      Tab(1).Control(2)=   "btoLimpar2(0)"
      Tab(1).Control(3)=   "btoModificar(0)"
      Tab(1).Control(4)=   "btoRemover(0)"
      Tab(1).Control(5)=   "btoAdicionar(0)"
      Tab(1).Control(6)=   "txtEndereçoEndCob"
      Tab(1).Control(7)=   "txtBaiEndCob"
      Tab(1).Control(8)=   "txtComplEndCob"
      Tab(1).Control(9)=   "txtTelEndCob"
      Tab(1).Control(10)=   "txtFaxEndCob"
      Tab(1).Control(11)=   "grdEndCob"
      Tab(1).Control(12)=   "txtCepEndCob"
      Tab(1).Control(13)=   "Label49"
      Tab(1).Control(14)=   "Label48"
      Tab(1).Control(15)=   "Label40"
      Tab(1).Control(16)=   "Label41"
      Tab(1).Control(17)=   "Label43"
      Tab(1).Control(18)=   "Label44"
      Tab(1).Control(19)=   "Label45"
      Tab(1).Control(20)=   "Label46"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "&Contato"
      TabPicture(2)   =   "frmCadastro_Geral.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label31"
      Tab(2).Control(1)=   "Label33"
      Tab(2).Control(2)=   "Label34"
      Tab(2).Control(3)=   "Label35"
      Tab(2).Control(4)=   "Label36"
      Tab(2).Control(5)=   "Label37"
      Tab(2).Control(6)=   "Label38"
      Tab(2).Control(7)=   "Label39"
      Tab(2).Control(8)=   "grdContato"
      Tab(2).Control(9)=   "txtNomeContato"
      Tab(2).Control(10)=   "txtDeptoContato"
      Tab(2).Control(11)=   "txtEmailContato"
      Tab(2).Control(12)=   "txtTelContato"
      Tab(2).Control(13)=   "txtCelContato"
      Tab(2).Control(14)=   "txtFaxContato"
      Tab(2).Control(15)=   "txtCargoContato"
      Tab(2).Control(16)=   "btoAdicionar(1)"
      Tab(2).Control(17)=   "btoRemover(1)"
      Tab(2).Control(18)=   "btoModificar(1)"
      Tab(2).Control(19)=   "btoLimpar2(1)"
      Tab(2).Control(20)=   "mskAniversarioContato"
      Tab(2).ControlCount=   21
      TabCaption(3)   =   "L&ocal Coleta/Entrega"
      TabPicture(3)   =   "frmCadastro_Geral.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label25"
      Tab(3).Control(1)=   "Label30"
      Tab(3).Control(2)=   "Label29"
      Tab(3).Control(3)=   "Label28"
      Tab(3).Control(4)=   "Label27"
      Tab(3).Control(5)=   "Label26"
      Tab(3).Control(6)=   "Label24"
      Tab(3).Control(7)=   "Label22"
      Tab(3).Control(8)=   "Label21"
      Tab(3).Control(9)=   "Label20"
      Tab(3).Control(10)=   "Label23"
      Tab(3).Control(11)=   "Label42"
      Tab(3).Control(12)=   "Label32"
      Tab(3).Control(13)=   "grdLocalColEnt"
      Tab(3).Control(14)=   "txtComplLocColEnt"
      Tab(3).Control(15)=   "txtObsLocColEnt"
      Tab(3).Control(16)=   "txtEmailLocColEnt"
      Tab(3).Control(17)=   "txtNomeContatoLocColEnt"
      Tab(3).Control(18)=   "txtTelLocColEnt"
      Tab(3).Control(19)=   "txtBaiLocColEnt"
      Tab(3).Control(20)=   "txtEndLocColEnt"
      Tab(3).Control(21)=   "txtNomeLocColEnt"
      Tab(3).Control(22)=   "btoAdicionar(2)"
      Tab(3).Control(23)=   "btoRemover(2)"
      Tab(3).Control(24)=   "btoModificar(2)"
      Tab(3).Control(25)=   "btoLimpar2(2)"
      Tab(3).Control(26)=   "cboUF(2)"
      Tab(3).Control(27)=   "txtCepLocColEnt"
      Tab(3).Control(28)=   "cboCidade(2)"
      Tab(3).Control(29)=   "txtFaxLocColEnt"
      Tab(3).ControlCount=   30
      TabCaption(4)   =   "Co&nsultar"
      TabPicture(4)   =   "frmCadastro_Geral.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label6"
      Tab(4).Control(1)=   "grdCadGer"
      Tab(4).Control(2)=   "txtConCadGeral"
      Tab(4).ControlCount=   3
      Begin Project_Combo_DB.Combo_DB cboCidade 
         Height          =   405
         Index           =   1
         Left            =   -73305
         TabIndex        =   24
         Top             =   2340
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   714
         Cols            =   0
         Rows            =   0
      End
      Begin VB.TextBox txtFaxLocColEnt 
         Height          =   315
         Left            =   -68280
         MaxLength       =   15
         TabIndex        =   53
         Top             =   1200
         Width           =   1575
      End
      Begin Project_Combo_DB.Combo_DB cboCidade 
         Height          =   405
         Index           =   2
         Left            =   -73905
         TabIndex        =   50
         Top             =   2475
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   714
         Cols            =   0
         Rows            =   0
      End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   5820
         TabIndex        =   119
         Top             =   4470
         Width           =   5160
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Usuário Liberação:"
            Height          =   195
            Left            =   135
            TabIndex        =   129
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Data Liberação:"
            Height          =   195
            Left            =   135
            TabIndex        =   128
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Usuário Bloqueio:"
            Height          =   195
            Left            =   135
            TabIndex        =   127
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Data Bloqueio:"
            Height          =   195
            Left            =   135
            TabIndex        =   126
            Top             =   480
            Width           =   1050
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Motivo Bloqueio:"
            Height          =   195
            Left            =   135
            TabIndex        =   125
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label LblcodMotBlq 
            Caption         =   "Label50"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1515
            TabIndex        =   124
            Top             =   240
            Width           =   3180
         End
         Begin VB.Label lblDatBlqGer 
            Caption         =   "Label51"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1515
            TabIndex        =   123
            Top             =   480
            Width           =   3180
         End
         Begin VB.Label LblCodUsuBlq 
            Caption         =   "Label52"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1515
            TabIndex        =   122
            Top             =   720
            Width           =   3180
         End
         Begin VB.Label LblDatLibBlq 
            Caption         =   "Label53"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1515
            TabIndex        =   121
            Top             =   960
            Width           =   3180
         End
         Begin VB.Label lblcodUsuLib 
            Caption         =   "Label54"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1515
            TabIndex        =   120
            Top             =   1200
            Width           =   3180
         End
      End
      Begin Project_Combo_DB.Combo_DB CboDesAtvGer 
         Height          =   390
         Left            =   7605
         TabIndex        =   16
         Top             =   1980
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   688
         Cols            =   0
         Rows            =   0
      End
      Begin Project_Masked.Masked txtInscrEst 
         Height          =   315
         Left            =   8280
         TabIndex        =   14
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
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
         MaxLength       =   16
      End
      Begin Project_Masked.Masked txtCNPJ 
         Height          =   330
         Left            =   1800
         TabIndex        =   3
         Top             =   1635
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project_Masked.Masked txtCepLocColEnt 
         Height          =   315
         Left            =   -70980
         TabIndex        =   48
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
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
         MaxLength       =   9
      End
      Begin Project_Masked.Masked mskAniversarioContato 
         Height          =   315
         Left            =   -67590
         TabIndex        =   40
         Top             =   2370
         Width           =   1125
         _ExtentX        =   1984
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
      End
      Begin VB.ComboBox cboUF 
         Height          =   315
         Index           =   2
         Left            =   -73800
         TabIndex        =   49
         Top             =   2070
         Width           =   630
      End
      Begin VB.ComboBox cboUF 
         Height          =   315
         Index           =   1
         Left            =   -73305
         TabIndex        =   23
         Top             =   1935
         Width           =   630
      End
      Begin VB.CommandButton btoLimpar2 
         Caption         =   "Lim&par"
         Height          =   270
         Index           =   2
         Left            =   -65250
         TabIndex        =   60
         Top             =   3270
         Width           =   945
      End
      Begin VB.CommandButton btoModificar 
         Caption         =   "&Modificar"
         Height          =   270
         Index           =   2
         Left            =   -66360
         TabIndex        =   59
         Top             =   3270
         Width           =   945
      End
      Begin VB.CommandButton btoRemover 
         Caption         =   "&Remover"
         Height          =   270
         Index           =   2
         Left            =   -67455
         TabIndex        =   58
         Top             =   3270
         Width           =   945
      End
      Begin VB.CommandButton btoAdicionar 
         Caption         =   "A&dicionar"
         Height          =   270
         Index           =   2
         Left            =   -68565
         TabIndex        =   57
         Top             =   3270
         Width           =   945
      End
      Begin VB.CommandButton btoLimpar2 
         Caption         =   "Lim&par"
         Height          =   270
         Index           =   1
         Left            =   -65250
         TabIndex        =   44
         Top             =   3270
         Width           =   945
      End
      Begin VB.CommandButton btoModificar 
         Caption         =   "&Modificar"
         Height          =   270
         Index           =   1
         Left            =   -66360
         TabIndex        =   43
         Top             =   3270
         Width           =   945
      End
      Begin VB.CommandButton btoRemover 
         Caption         =   "&Remover"
         Height          =   270
         Index           =   1
         Left            =   -67455
         TabIndex        =   42
         Top             =   3270
         Width           =   945
      End
      Begin VB.CommandButton btoAdicionar 
         Caption         =   "A&dicionar"
         Height          =   270
         Index           =   1
         Left            =   -68565
         TabIndex        =   41
         Top             =   3270
         Width           =   945
      End
      Begin VB.CommandButton btoLimpar2 
         Caption         =   "Lim&par"
         Height          =   270
         Index           =   0
         Left            =   -65250
         TabIndex        =   32
         Top             =   3270
         Width           =   945
      End
      Begin VB.CommandButton btoModificar 
         Caption         =   "&Modificar"
         Height          =   270
         Index           =   0
         Left            =   -66360
         TabIndex        =   31
         Top             =   3270
         Width           =   945
      End
      Begin VB.CommandButton btoRemover 
         Caption         =   "&Remover"
         Height          =   270
         Index           =   0
         Left            =   -67455
         TabIndex        =   30
         Top             =   3270
         Width           =   945
      End
      Begin VB.CommandButton btoAdicionar 
         Caption         =   "A&dicionar"
         Height          =   270
         Index           =   0
         Left            =   -68565
         TabIndex        =   29
         Top             =   3270
         Width           =   945
      End
      Begin VB.TextBox txtNomeLocColEnt 
         Height          =   315
         Left            =   -73800
         MaxLength       =   25
         TabIndex        =   45
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtEndLocColEnt 
         Height          =   315
         Left            =   -73800
         MaxLength       =   40
         TabIndex        =   46
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtBaiLocColEnt 
         Height          =   315
         Left            =   -73800
         MaxLength       =   20
         TabIndex        =   47
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtTelLocColEnt 
         Height          =   315
         Left            =   -68280
         MaxLength       =   15
         TabIndex        =   52
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtNomeContatoLocColEnt 
         Height          =   315
         Left            =   -68280
         MaxLength       =   30
         TabIndex        =   54
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txtEmailLocColEnt 
         Height          =   315
         Left            =   -68280
         MaxLength       =   40
         TabIndex        =   55
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txtObsLocColEnt 
         Height          =   495
         Left            =   -68280
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         Top             =   2640
         Width           =   4215
      End
      Begin VB.TextBox txtEndereçoEndCob 
         Height          =   315
         Left            =   -73305
         MaxLength       =   40
         TabIndex        =   21
         Top             =   990
         Width           =   3855
      End
      Begin VB.TextBox txtBaiEndCob 
         Height          =   315
         Left            =   -73305
         MaxLength       =   20
         TabIndex        =   22
         Top             =   1485
         Width           =   2055
      End
      Begin VB.TextBox txtComplEndCob 
         Height          =   315
         Left            =   -67545
         MaxLength       =   10
         TabIndex        =   26
         Top             =   1470
         Width           =   1095
      End
      Begin VB.TextBox txtTelEndCob 
         Height          =   315
         Left            =   -67545
         MaxLength       =   15
         TabIndex        =   27
         Top             =   1950
         Width           =   1575
      End
      Begin VB.TextBox txtFaxEndCob 
         Height          =   315
         Left            =   -67545
         MaxLength       =   15
         TabIndex        =   28
         Top             =   2430
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados comerciais"
         Height          =   3780
         Left            =   120
         TabIndex        =   89
         Top             =   2775
         Width           =   5655
         Begin Project_Masked.Masked txtCep 
            Height          =   315
            Left            =   1680
            TabIndex        =   10
            Top             =   2010
            Width           =   1020
            _ExtentX        =   1799
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
         End
         Begin Project_Combo_DB.Combo_DB cboCidade 
            Height          =   315
            Index           =   0
            Left            =   1650
            TabIndex        =   8
            Top             =   990
            Width           =   3960
            _ExtentX        =   6985
            _ExtentY        =   556
            Cols            =   0
            Rows            =   0
         End
         Begin VB.TextBox txtEndCom 
            Height          =   315
            Left            =   1680
            MaxLength       =   40
            TabIndex        =   6
            Top             =   255
            Width           =   3855
         End
         Begin VB.ComboBox cboUF 
            Height          =   315
            Index           =   0
            ItemData        =   "frmCadastro_Geral.frx":008C
            Left            =   1680
            List            =   "frmCadastro_Geral.frx":008E
            TabIndex        =   7
            Top             =   615
            Width           =   630
         End
         Begin VB.TextBox txtBaiCml 
            Height          =   315
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   9
            Top             =   1650
            Width           =   2055
         End
         Begin VB.TextBox txtComplEndCml 
            Height          =   315
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   11
            Top             =   2370
            Width           =   1590
         End
         Begin VB.TextBox txtTelCml 
            Height          =   315
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   12
            Top             =   2730
            Width           =   1245
         End
         Begin VB.TextBox txtFaxCml 
            Height          =   315
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   13
            Top             =   3090
            Width           =   1245
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            Height          =   195
            Left            =   105
            TabIndex        =   97
            Top             =   615
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            Height          =   195
            Left            =   120
            TabIndex        =   96
            Top             =   255
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   1650
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   1050
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   2010
            Width           =   360
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Compl. do Endereço:"
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   2370
            Width           =   1485
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Telefone:"
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   2730
            Width           =   675
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   3090
            Width           =   300
         End
      End
      Begin VB.TextBox txtConCadGeral 
         Height          =   315
         Left            =   -73020
         MaxLength       =   50
         TabIndex        =   66
         Top             =   705
         Width           =   4680
      End
      Begin VB.TextBox txtCargoContato 
         Height          =   315
         Left            =   -73590
         MaxLength       =   30
         TabIndex        =   35
         Top             =   1935
         Width           =   2895
      End
      Begin VB.TextBox txtFaxContato 
         Height          =   315
         Left            =   -67590
         MaxLength       =   15
         TabIndex        =   39
         Top             =   1935
         Width           =   1575
      End
      Begin VB.TextBox txtCelContato 
         Height          =   315
         Left            =   -67590
         MaxLength       =   15
         TabIndex        =   38
         Top             =   1455
         Width           =   1575
      End
      Begin VB.TextBox txtTelContato 
         Height          =   315
         Left            =   -67590
         MaxLength       =   15
         TabIndex        =   37
         Top             =   975
         Width           =   1575
      End
      Begin VB.TextBox txtEmailContato 
         Height          =   315
         Left            =   -73590
         MaxLength       =   40
         TabIndex        =   36
         Top             =   2415
         Width           =   3855
      End
      Begin VB.TextBox txtDeptoContato 
         Height          =   315
         Left            =   -73590
         MaxLength       =   30
         TabIndex        =   34
         Top             =   1455
         Width           =   2895
      End
      Begin VB.TextBox txtNomeContato 
         Height          =   315
         Left            =   -73590
         MaxLength       =   40
         TabIndex        =   33
         Top             =   975
         Width           =   3855
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   6600
         MaxLength       =   50
         TabIndex        =   20
         Top             =   3540
         Width           =   4335
      End
      Begin VB.TextBox txtHomePage 
         Height          =   315
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   19
         Top             =   3060
         Width           =   3975
      End
      Begin VB.CheckBox chkINSS 
         Caption         =   "Retêm INSS"
         Height          =   255
         Left            =   8520
         TabIndex        =   18
         Top             =   2580
         Width           =   1935
      End
      Begin VB.CheckBox chkRetISS 
         Caption         =   "Retêm ISS"
         Height          =   255
         Left            =   6840
         TabIndex        =   17
         Top             =   2580
         Width           =   1935
      End
      Begin VB.TextBox txtInscrMun 
         Height          =   315
         Left            =   8280
         MaxLength       =   16
         TabIndex        =   15
         Top             =   1500
         Width           =   2055
      End
      Begin VB.TextBox txtNomRed 
         Height          =   315
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2385
         Width           =   1575
      End
      Begin VB.Frame fraTipCli 
         Caption         =   "Pessoa"
         Height          =   615
         Left            =   240
         TabIndex        =   74
         Top             =   840
         Width           =   3840
         Begin VB.OptionButton optPessoa 
            Caption         =   "Exterior"
            Height          =   255
            Index           =   2
            Left            =   2880
            TabIndex        =   2
            Top             =   240
            Width           =   900
         End
         Begin VB.OptionButton optPessoa 
            Caption         =   "Jurídica"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optPessoa 
            Caption         =   "Física"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtRazSoc 
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2017
         Width           =   4980
      End
      Begin MSDataGridLib.DataGrid grdCadGer 
         Height          =   5310
         Left            =   -74640
         TabIndex        =   70
         Top             =   1200
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   9366
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdContato 
         Height          =   2895
         Left            =   -74640
         TabIndex        =   68
         Top             =   3600
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5106
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdEndCob 
         Height          =   2895
         Left            =   -74640
         TabIndex        =   67
         Top             =   3600
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5106
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin VB.TextBox txtComplLocColEnt 
         Height          =   315
         Left            =   -73560
         MaxLength       =   10
         TabIndex        =   51
         Top             =   3120
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid grdLocalColEnt 
         Height          =   2895
         Left            =   -74640
         TabIndex        =   69
         Top             =   3675
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5106
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
      End
      Begin Project_Masked.Masked txtCepEndCob 
         Height          =   315
         Left            =   -67545
         TabIndex        =   25
         Top             =   1035
         Width           =   1080
         _ExtentX        =   1905
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
      End
      Begin VB.Label Label32 
         Caption         =   "Label32"
         Height          =   420
         Left            =   -71310
         TabIndex        =   130
         Top             =   2445
         Width           =   1230
      End
      Begin VB.Label Atividade 
         AutoSize        =   -1  'True
         Caption         =   "Atividade"
         Height          =   195
         Left            =   6840
         TabIndex        =   118
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   117
         Top             =   2505
         Width           =   540
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   116
         Top             =   2085
         Width           =   255
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   195
         Left            =   -74625
         TabIndex        =   115
         Top             =   2370
         Width           =   540
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "UF:"
         Height          =   195
         Left            =   -74625
         TabIndex        =   114
         Top             =   1950
         Width           =   255
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   113
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   112
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   111
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         Height          =   195
         Left            =   -71445
         TabIndex        =   110
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   -69600
         TabIndex        =   108
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
         Height          =   195
         Left            =   -69600
         TabIndex        =   107
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Contato:"
         Height          =   195
         Left            =   -69600
         TabIndex        =   106
         Top             =   1680
         Width           =   1290
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "E-mail do Contato:"
         Height          =   195
         Left            =   -69600
         TabIndex        =   105
         Top             =   2160
         Width           =   1290
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
         Height          =   195
         Left            =   -69600
         TabIndex        =   104
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   -74625
         TabIndex        =   103
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   -74625
         TabIndex        =   102
         Top             =   1470
         Width           =   450
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         Height          =   195
         Left            =   -68625
         TabIndex        =   101
         Top             =   1035
         Width           =   360
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         Height          =   195
         Left            =   -68625
         TabIndex        =   100
         Top             =   1470
         Width           =   1005
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   -68625
         TabIndex        =   99
         Top             =   1950
         Width           =   675
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
         Height          =   195
         Left            =   -68625
         TabIndex        =   98
         Top             =   2430
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ / Razão Social:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   88
         Top             =   705
         Width           =   1560
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Cargo:"
         Height          =   195
         Left            =   -74670
         TabIndex        =   87
         Top             =   1935
         Width           =   465
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Data de Aniversário:"
         Height          =   195
         Left            =   -69150
         TabIndex        =   86
         Top             =   2415
         Width           =   1440
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
         Height          =   195
         Left            =   -69150
         TabIndex        =   85
         Top             =   1935
         Width           =   300
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Celular:"
         Height          =   195
         Left            =   -69150
         TabIndex        =   84
         Top             =   1455
         Width           =   525
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Telefone:"
         Height          =   195
         Left            =   -69150
         TabIndex        =   83
         Top             =   975
         Width           =   675
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "E-mail:"
         Height          =   195
         Left            =   -74670
         TabIndex        =   82
         Top             =   2415
         Width           =   465
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Departamento:"
         Height          =   195
         Left            =   -74670
         TabIndex        =   81
         Top             =   1455
         Width           =   1050
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Left            =   -74670
         TabIndex        =   80
         Top             =   975
         Width           =   465
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail:"
         Height          =   195
         Left            =   6000
         TabIndex        =   79
         Top             =   3660
         Width           =   480
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Home Page:"
         Height          =   195
         Left            =   6000
         TabIndex        =   78
         Top             =   3180
         Width           =   885
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Municipal:"
         Height          =   195
         Left            =   6840
         TabIndex        =   77
         Top             =   1620
         Width           =   1410
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Estadual:"
         Height          =   195
         Left            =   6840
         TabIndex        =   76
         Top             =   1140
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome Reduzido:"
         Height          =   195
         Left            =   240
         TabIndex        =   75
         Top             =   2385
         Width           =   1185
      End
      Begin VB.Label lblCNPJ 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ:"
         Height          =   195
         Left            =   240
         TabIndex        =   73
         Top             =   1740
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Razão Social:"
         Height          =   195
         Left            =   240
         TabIndex        =   72
         Top             =   2010
         Width           =   990
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Complemento:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   109
         Top             =   3120
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmCadastro_Geral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'* Desenvolvido por César Augusto de Oliveira Barros em 17/07/2003 *
'* Alterado por Flávia Aguiar em 06/11/2003
'*******************************************************************
Option Explicit
Public blI As Byte
Dim slps As String * 1

Private Sub btoAdicionar_Click(Index As Integer)
    'carrega o grid correspondente ao index com os dados da tela
On Error GoTo trataerro
    Dim ilColuna As Integer, ilLinha As Integer, JaExiste As Boolean
    Select Case Index
        Case 0 'Endereço/Cobrança
            
            'Tab 1
            JaExiste = False
            For ilLinha = 2 To grdEndCob.Rows
                If txtEndereçoEndCob.Text = grdEndCob.TextMatrix(ilLinha - 1, 2) And txtBaiEndCob.Text = grdEndCob.TextMatrix(ilLinha - 1, 3) Then
                    JaExiste = True
                    MsgBox "Esse endereço de cobrança já foi cadastrado.", vbInformation, "Aviso!"
                    Exit Sub
                End If
            Next ilLinha
            If cboCidade(1).Codigo = "" Then
               MsgBox "Favor escolher a cidade", vbInformation, "Atenção"
               Exit Sub
            End If
            If Not VerificaCampo(txtEndereçoEndCob, "Endereço de Cobrança", 1) Then Exit Sub
            If Not VerificaCampo(txtBaiEndCob, "Bairro de Cobrança", 1) Then Exit Sub
            If Not VerificaCampo(cboUF(1), "Estado de Cobrança", 1) Then Exit Sub
            If Not VerificaCampo(cboCidade(1), "Cidade de Cobrança", 1) Then Exit Sub
            If Not VerificaCampo(txtCepEndCob, "Cep de Cobrança", 1) Then Exit Sub
            'If Not VerificaCampo(txtTelEndCob, "Telefone de Cobrança", 1) Then Exit Sub
            
            sgQuery = "Select Sequência = SeqCob, Endereço = EndCob, Bairro = BaiCob, Cidade = CodCidCob, Cep = CepCob, Complemento = ComCob, Telefone = TelCob, Fax = FaxCob "
            sgQuery = sgQuery & " from EndCob_Cadastro_Geral"
            consulta sgQuery
            
            grdEndCob.Rows = grdEndCob.Rows + 1
            grdEndCob.TextMatrix(grdEndCob.Rows - 1, 0) = cboCidade(1).Codigo
            
            'Sequência
            If grdEndCob.Rows > 2 Then
                grdEndCob.TextMatrix(grdEndCob.Rows - 1, 1) = Format(grdEndCob.TextMatrix(grdEndCob.Rows - 2, 1) + 1, "00000")
            Else
                grdEndCob.TextMatrix(grdEndCob.Rows - 1, 1) = Format(1, "00000")
            End If
            
            grdEndCob.TextMatrix(grdEndCob.Rows - 1, 2) = txtEndereçoEndCob.Text
            grdEndCob.TextMatrix(grdEndCob.Rows - 1, 3) = txtBaiEndCob.Text
            grdEndCob.TextMatrix(grdEndCob.Rows - 1, 4) = Trim(cboCidade(1).Criterio)
            grdEndCob.TextMatrix(grdEndCob.Rows - 1, 5) = txtCepEndCob.Texto
            grdEndCob.TextMatrix(grdEndCob.Rows - 1, 6) = Trim(txtComplEndCob.Text)
            grdEndCob.TextMatrix(grdEndCob.Rows - 1, 7) = Trim(txtTelEndCob.Text)
            grdEndCob.TextMatrix(grdEndCob.Rows - 1, 8) = Trim(txtFaxEndCob.Text)
            
            'Colocando o tamanho da coluna de acordo com o texto
            AjustaColWidth grdEndCob
            If Not BtoIncluir.Enabled Then
                Grava_End_Cob
            End If
        Case 1 'Contato
            
            'Tab 2
            JaExiste = False
            For ilLinha = 2 To grdContato.Rows
                If txtNomeContato.Text = grdContato.TextMatrix(ilLinha - 1, 1) And txtDeptoContato.Text = grdContato.TextMatrix(ilLinha - 1, 2) Then
                    JaExiste = True
                    MsgBox "Esse contato já foi cadastrado.", vbInformation, "Aviso!"
                    Exit Sub
                End If
            Next ilLinha
            
            If Not VerificaCampo(txtNomeContato, "Nome do Contato", 2) Then Exit Sub
            'If Not VerificaCampo(txtDeptoContato, "Departamento do Contato", 2) Then Exit Sub
            'If Not VerificaCampo(txtCargoContato, "Cargo do Contato", 2) Then Exit Sub
            'If Not VerificaCampo(txtTelContato, "Telefone do Contato", 2) Then Exit Sub
            
            sgQuery = "Select Sequência = SeqCto, Nome = NomCto, Departamento = DepCto, Cargo = CgoCto, [E-mail] = MailCto, Telefone = TelCto, Celular = CelCto, Fax = FaxCto, Aniversário = DatAnvCto "
            sgQuery = sgQuery & " From Contato_Cadastro_Geral"
            consulta sgQuery
            
            grdContato.Rows = grdContato.Rows + 1
                        
            'Sequência
            If grdContato.Rows > 2 Then
                grdContato.TextMatrix(grdContato.Rows - 1, 0) = Format(grdContato.TextMatrix(grdContato.Rows - 2, 0) + 1, "00000")
            Else
                grdContato.TextMatrix(grdContato.Rows - 1, 0) = Format(1, "00000")
            End If
            
            grdContato.TextMatrix(grdContato.Rows - 1, 1) = txtNomeContato.Text
            grdContato.TextMatrix(grdContato.Rows - 1, 2) = txtDeptoContato.Text
            grdContato.TextMatrix(grdContato.Rows - 1, 3) = txtCargoContato.Text
            grdContato.TextMatrix(grdContato.Rows - 1, 4) = txtEmailContato.Text
            grdContato.TextMatrix(grdContato.Rows - 1, 5) = txtTelContato.Text
            grdContato.TextMatrix(grdContato.Rows - 1, 6) = txtCelContato.Text
            grdContato.TextMatrix(grdContato.Rows - 1, 7) = txtFaxContato.Text
            grdContato.TextMatrix(grdContato.Rows - 1, 8) = mskAniversarioContato.Texto
            
            'Colocando o tamanho da coluna de acordo com o texto
            AjustaColWidth grdContato
            If Not BtoIncluir.Enabled Then
                Grava_Contato
            End If
        Case 2 'Local Coleta/Entrega
            
            'Tab 3
            
            JaExiste = False
            For ilLinha = 2 To grdLocalColEnt.Rows
                If txtNomeLocColEnt.Text = grdLocalColEnt.TextMatrix(ilLinha - 1, 2) And txtEndLocColEnt.Text = grdLocalColEnt.TextMatrix(ilLinha - 1, 3) Then
                    JaExiste = True
                    MsgBox "Esse local de coleta/entrega já foi cadastrado.", vbInformation, "Aviso!"
                    Exit Sub
                End If
            Next ilLinha
            
            If Not VerificaCampo(txtNomeLocColEnt, "Nome do Local de Coleta/Entrega", 3) Then Exit Sub
            If Not VerificaCampo(txtEndLocColEnt, "Endereço do Local de Coleta/Entrega", 3) Then Exit Sub
            If Not VerificaCampo(txtBaiLocColEnt, "Bairro do Local de Coleta/Entrega", 3) Then Exit Sub
            If Not VerificaCampo(txtCepLocColEnt, "Cep do Local de Coleta/Entrega", 3) Then Exit Sub
            If Not VerificaCampo(cboUF(2), "Estado do Local de Coleta/Entrega", 3) Then Exit Sub
            If Not VerificaCampo(cboCidade(2), "Cidade do Local de Coleta/Entrega", 3) Then Exit Sub
            'If Not VerificaCampo(txtTelLocColEnt, "Telefone do Local de Coleta/Entrega", 3) Then Exit Sub
            'If Not VerificaCampo(txtNomeContatoLocColEnt, "Nome do Contato do Local de Coleta/Entrega", 3) Then Exit Sub
            
            sgQuery = "Select Sequência = SeqLoc, Nome = NomLoc, Endereço = EndLoc, Bairro = BaiLoc, Cep = CepLoc, Cidade = CodCid, Complemento = ComEndLoc, Telefone = TelLoc, Fax = FaxLoc, [Nome do Contato] = NomCto, [E-mail do Contato] = MailCto, Observação = ObsLoc "
            sgQuery = sgQuery & " from Local_Coleta_Entrega"
            consulta sgQuery
            
            grdLocalColEnt.Rows = grdLocalColEnt.Rows + 1
            
            grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 0) = cboCidade(2).Codigo
            
            'Sequência
            If grdLocalColEnt.Rows > 2 Then
                grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 1) = Format(grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 2, 1) + 1, "00000")
            Else
                grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 1) = Format(1, "00000")
            End If
            
            grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 2) = Trim(txtNomeLocColEnt.Text)
            grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 3) = Trim(txtEndLocColEnt.Text)
            grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 4) = Trim(txtBaiLocColEnt.Text)
            grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 5) = Trim(txtCepLocColEnt.Texto)
            grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 6) = Trim(cboCidade(2).Criterio)
            grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 7) = Trim(txtComplLocColEnt.Text)
            grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 8) = Trim(txtTelLocColEnt.Text)
            grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 9) = Trim(txtFaxLocColEnt.Text)
            grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 10) = Trim(txtNomeContatoLocColEnt.Text)
            grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 11) = Trim(txtEmailLocColEnt.Text)
            grdLocalColEnt.TextMatrix(grdLocalColEnt.Rows - 1, 12) = Trim(txtObsLocColEnt.Text)
            
            'Colocando o tamanho da coluna de acordo com o texto
            AjustaColWidth grdLocalColEnt
            If Not BtoIncluir.Enabled Then
                Grava_Ent_Col
            End If
    End Select
    Call btoLimpar2_Click(Index)
    Exit Sub
trataerro:
    Rotina_Erro "btoAdicionar_Click"
End Sub

Private Sub btoAlterar_Click()
Dim slps As String
On Error GoTo trataerro
    sgFlagOper = "A"
    slps = "O"
    slps = IIf(optPessoa(0).Value, "P", "J")
    If VerificaCampos(slps) = False Then
       Exit Sub
    End If
    If grdEndCob.Rows = 1 Then
      MsgBox "Favor informar o endereço de cobrança", vbInformation, "Atenção"
      SSTab1.Tab = 1
      Exit Sub
    End If
    Call Gravar
    Call LimpaCampos
    Call LimpaGrids
    Call SetBotoes(True)
    txtCNPJ.SetFocus
    Set grdCadGer.DataSource = Nothing
    Exit Sub
trataerro:
    Rotina_Erro "btoAlterar_Click"
End Sub

Private Sub btoExcluir_Click()
On Error GoTo trataerro
    If MsgBox("Tem certeza que você deseja EXCLUIR " & txtRazSoc.Text & " " & lblCNPJ.Caption & " " & txtCNPJ.Texto & " ?", vbCritical + vbYesNo + vbDefaultButton2, "Confirmação:") = vbYes Then
        Conexao.BeginTrans
        Conexao.Execute "Delete From ENDCOB_CADASTRO_GERAL Where CNPJGer='" & txtCNPJ.Texto & "'"
        Set Rs = Nothing
        Conexao.Execute "Delete From LOCAL_COLETA_ENTREGA Where CNPJGer='" & txtCNPJ.Texto & "'"
        Set Rs = Nothing
        Conexao.Execute "Delete From CONTATO_CADASTRO_GERAL Where CNPJGer='" & txtCNPJ.Texto & "'"
        Set Rs = Nothing
        Conexao.Execute "Delete From CADASTRO_GERAL Where CNPJGer='" & txtCNPJ.Texto & "'"
        Set Rs = Nothing
        Conexao.CommitTrans
        LimpaCampos
        LimpaGrids
        Set grdCadGer.DataSource = Nothing
        Call SetBotoes(True)
    End If
    Exit Sub
    Resume
trataerro:
    Rotina_Erro "btoExcluir_Click"
End Sub

Private Sub btoIncluir_Click()
On Error GoTo trataerro
    sgFlagOper = "I"
'    slps = "E"
'    slps = IIf(optPessoa(0).Value, "P", "J")
'    If optPessoa(0).Value Then
'        slps = "P"
'    ElseIf optPessoa(1).Value Then
'        slps = "J"
'    ElseIf optPessoa(2).Value Then
'        slps = "E"
'    End If
    If VerificaCampos(slps) = False Then
       Exit Sub
    End If
   If grdEndCob.Rows = 1 Then
      MsgBox "Favor informar o endereço de cobrança", vbInformation, "Atenção"
      Exit Sub
    End If
    Call Gravar
    Call LimpaCampos
    Call LimpaGrids
    Call SetBotoes(True)
    txtCNPJ.SetFocus
    Set grdCadGer.DataSource = Nothing
    Exit Sub
trataerro:
    Rotina_Erro "btoIncluir_Click"
End Sub

Private Sub btoLimpar_Click()
On Error GoTo trataerro
    LimpaCampos
    LimpaGrids
    Call SetBotoes(True)
    Set grdCadGer.DataSource = Nothing
    SSTab1.TabEnabled(4) = True
    Exit Sub
trataerro:
    Rotina_Erro "btoLimpar_Click"
End Sub

Private Sub btoLimpar2_Click(Index As Integer)
On Error GoTo trataerro
    Select Case Index
        Case 0 'Endereço/Cobrança
            txtEndereçoEndCob.Text = ""
            txtBaiEndCob.Text = ""
            txtCepEndCob.Texto = ""
            txtComplEndCob.Text = ""
            txtTelEndCob.Text = ""
            txtFaxEndCob.Text = ""
            cboUF(1).Text = ""
            cboCidade(1).Criterio = ""
            txtEndereçoEndCob.SetFocus
        Case 1 'Contato
            txtNomeContato.Text = ""
            txtDeptoContato.Text = ""
            txtCargoContato.Text = ""
            txtEmailContato.Text = ""
            txtTelContato.Text = ""
            txtCelContato.Text = ""
            txtFaxContato.Text = ""
            mskAniversarioContato.Texto = ""
            txtNomeContato.SetFocus
        Case 2 'Local Coleta/Entrega
            txtNomeLocColEnt.Text = ""
            txtEndLocColEnt.Text = ""
            txtBaiLocColEnt.Text = ""
            txtCepLocColEnt.Texto = ""
            cboUF(2).Text = ""
            cboCidade(2).Criterio = ""
            txtComplLocColEnt.Text = ""
            txtTelLocColEnt.Text = ""
            txtFaxLocColEnt.Text = ""
            txtNomeContatoLocColEnt.Text = ""
            txtEmailLocColEnt.Text = ""
            txtObsLocColEnt.Text = ""
            txtNomeLocColEnt.SetFocus
    End Select
    Exit Sub
trataerro:
    Rotina_Erro "btoLimpar2_Click"
End Sub

Private Sub btoLimpar2_LostFocus(Index As Integer)
    Select Case Index
        Case 0
            SSTab1.Tab = 2
            txtNomeContato.SetFocus
        Case 1
            SSTab1.Tab = 3
            txtNomeLocColEnt.SetFocus
        Case 2
        Case 3
    End Select
    
End Sub

Private Sub btoModificar_Click(Index As Integer)
On Error GoTo trataerro
    Select Case Index
        Case 0 'Endereço/Cobrança
            If grdEndCob.Rows > 1 Then
                'Tab 1
                If Not VerificaCampo(txtEndereçoEndCob, "Endereço de Cobrança", 1) Then Exit Sub
                If Not VerificaCampo(txtBaiEndCob, "Bairro de Cobrança", 1) Then Exit Sub
                If Not VerificaCampo(cboUF(1), "Estado de Cobrança", 1) Then Exit Sub
                If Not VerificaCampo(cboCidade(1), "Cidade de Cobrança", 1) Then Exit Sub
                If Not VerificaCampo(txtCepEndCob, "Cep de Cobrança", 1) Then Exit Sub
                'If Not VerificaCampo(txtTelEndCob, "Telefone de Cobrança", 1) Then Exit Sub
                
                grdEndCob.TextMatrix(grdEndCob.RowSel, 0) = grdEndCob.TextMatrix(grdEndCob.RowSel, 0)
                grdEndCob.TextMatrix(grdEndCob.RowSel, 2) = Trim(txtEndereçoEndCob.Text)
                grdEndCob.TextMatrix(grdEndCob.RowSel, 3) = Trim(txtBaiEndCob.Text)
                grdEndCob.TextMatrix(grdEndCob.RowSel, 4) = cboCidade(1).Criterio
                grdEndCob.TextMatrix(grdEndCob.RowSel, 5) = RetiraFormatacao(txtCepEndCob.Texto)
                grdEndCob.TextMatrix(grdEndCob.RowSel, 6) = Trim(txtComplEndCob.Text)
                grdEndCob.TextMatrix(grdEndCob.RowSel, 7) = Trim(txtTelEndCob.Text)
                grdEndCob.TextMatrix(grdEndCob.RowSel, 8) = Trim(txtFaxEndCob.Text)
            End If
        Case 1 'Contato
            If grdContato.Rows > 1 Then
                'Tab 2
                If Not VerificaCampo(txtNomeContato, "Nome do Contato", 2) Then Exit Sub
                'If Not VerificaCampo(txtDeptoContato, "Departamento do Contato", 2) Then Exit Sub
                'If Not VerificaCampo(txtCargoContato, "Cargo do Contato", 2) Then Exit Sub
                'If Not VerificaCampo(txtTelContato, "Telefone do Contato", 2) Then Exit Sub
                
                grdContato.TextMatrix(grdContato.Rows - 1, 1) = txtNomeContato.Text
                grdContato.TextMatrix(grdContato.Rows - 1, 2) = txtDeptoContato.Text
                grdContato.TextMatrix(grdContato.Rows - 1, 3) = txtCargoContato.Text
                grdContato.TextMatrix(grdContato.Rows - 1, 4) = txtEmailContato.Text
                grdContato.TextMatrix(grdContato.Rows - 1, 5) = txtTelContato.Text
                grdContato.TextMatrix(grdContato.Rows - 1, 6) = txtCelContato.Text
                grdContato.TextMatrix(grdContato.Rows - 1, 7) = txtFaxContato.Text
                grdContato.TextMatrix(grdContato.Rows - 1, 8) = mskAniversarioContato.Texto
            End If
        Case 2 'Local Coleta/Entrega
            If grdLocalColEnt.Rows > 1 Then
                'Tab 3
                If Not VerificaCampo(txtNomeLocColEnt, "Nome do Local de Coleta/Entrega", 3) Then Exit Sub
                If Not VerificaCampo(txtEndLocColEnt, "Endereço do Local de Coleta/Entrega", 3) Then Exit Sub
                If Not VerificaCampo(txtBaiLocColEnt, "Bairro do Local de Coleta/Entrega", 3) Then Exit Sub
                If Not VerificaCampo(txtCepLocColEnt, "Cep do Local de Coleta/Entrega", 3) Then Exit Sub
                If Not VerificaCampo(cboUF(2), "Estado do Local de Coleta/Entrega", 3) Then Exit Sub
                If Not VerificaCampo(cboCidade(2), "Cidade do Local de Coleta/Entrega", 3) Then Exit Sub
                'If Not VerificaCampo(txtTelLocColEnt, "Telefone do Local de Coleta/Entrega", 3) Then Exit Sub
                'If Not VerificaCampo(txtNomeContatoLocColEnt, "Nome do Contato do Local de Coleta/Entrega", 3) Then Exit Sub
                
                grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 0) = cboCidade(2).Codigo
                grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 2) = Trim(txtNomeLocColEnt.Text)
                grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 3) = Trim(txtEndLocColEnt.Text)
                grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 4) = Trim(txtBaiLocColEnt.Text)
                grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 5) = RetiraFormatacao(txtCepLocColEnt.Texto)
                grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 6) = cboCidade(2).Criterio
                grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 7) = Trim(txtComplLocColEnt.Text)
                grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 8) = Trim(txtTelLocColEnt.Text)
                grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 9) = Trim(txtFaxLocColEnt.Text)
                grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 10) = Trim(txtNomeContatoLocColEnt.Text)
                grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 11) = Trim(txtEmailLocColEnt.Text)
                grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 12) = Trim(txtObsLocColEnt.Text)
            End If
    End Select
    Exit Sub
trataerro:
    Rotina_Erro "btoModificar_Click"
End Sub

Private Sub btoRemover_Click(Index As Integer)
On Error GoTo trataerro
    Select Case Index
        Case 0 'Endereço/Cobrança
            If grdEndCob.Rows = 1 Then Exit Sub
            consulta "Select CNPJGer From ENDCOB_CADASTRO_GERAL Where CNPJGer='" & txtCNPJ.Texto & "' And SeqCob='" & grdEndCob.TextMatrix(grdEndCob.RowSel, 1) & "'"
            If Rs.RecordCount > 0 Then
                Conexao.Execute "Delete From ENDCOB_CADASTRO_GERAL Where CNPJGer='" & txtCNPJ.Texto & "' And SeqCob='" & grdEndCob.TextMatrix(grdEndCob.RowSel, 1) & "'"
            End If
            If grdEndCob.RowSel > 1 Or grdEndCob.Rows > 2 Then
                grdEndCob.RemoveItem grdEndCob.RowSel
            Else
                grdEndCob.Rows = 1
            End If
            
        Case 1 'Contato
            If grdContato.Rows = 1 Then Exit Sub
            consulta "Select CNPJGer From CONTATO_CADASTRO_GERAL Where CNPJGer='" & txtCNPJ.Texto & "' And SeqCto='" & grdContato.TextMatrix(grdContato.RowSel, 0) & "'"
            If Rs.RecordCount > 0 Then
                Conexao.Execute "Delete From CONTATO_CADASTRO_GERAL Where CNPJGer='" & txtCNPJ.Texto & "' And SeqCto='" & grdContato.TextMatrix(grdContato.RowSel, 0) & "'"
            End If
            If grdContato.RowSel > 1 Or grdContato.Rows > 2 Then
                grdContato.RemoveItem grdContato.RowSel
            Else
                grdContato.Rows = 1
            End If
            
        Case 2 'Local Coleta/Entrega
            If grdLocalColEnt.Rows = 1 Then Exit Sub
            consulta "Select CNPJGer From LOCAL_COLETA_ENTREGA Where CNPJGer='" & txtCNPJ.Texto & "' And SeqLoc='" & grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 1) & "'"
            If Rs.RecordCount > 0 Then
                Conexao.Execute "Delete From LOCAL_COLETA_ENTREGA Where CNPJGer='" & txtCNPJ.Texto & "' And SeqLoc='" & grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 1) & "'"
            End If
            If grdLocalColEnt.RowSel > 1 Or grdLocalColEnt.Rows > 2 Then
                grdLocalColEnt.RemoveItem grdLocalColEnt.RowSel
            Else
                grdLocalColEnt.Rows = 1
            End If
            
    End Select
    Exit Sub
trataerro:
    Rotina_Erro "btoRemover_Click"
End Sub

Private Sub BtoSair_Click()
    Unload Me
End Sub

Private Sub cboCidade_Consultar(Index As Integer)
    cboCidade(Index).query = "Select NomCid, CodCid From Cidade Where CodUF='" & Trim(cboUF(Index).Text) & "' And NomCid like '" & Trim(cboCidade(Index).Criterio) & "%'"
End Sub

Private Sub cboCidade_GotFocus(Index As Integer)
    SelecionaTudo
End Sub

Private Sub cboCidade_LostFocus(Index As Integer)
    If cboCidade(0).Codigo <> "" And Trim(cboCidade(1).Criterio) = "" Then
       cboCidade(1).Codigo = cboCidade(0).Codigo
       cboCidade(1).Criterio = cboCidade(0).Criterio
    End If
End Sub

Private Sub CboDesAtvGer_Consultar()
    CboDesAtvGer.query = "Select 'Nome Atividade' = DesAtvGer ,'Código Atividade' = CodAtvGer From Atividade where " & IIf(IsNumeric(CboDesAtvGer.Criterio), "CodAtvGer =" & CboDesAtvGer.Criterio, "Desatvger like '" & CboDesAtvGer.Criterio & "%'") & " order by " & IIf(IsNumeric(CboDesAtvGer.Criterio), "CodAtvGer", "Desatvger")
End Sub

Private Sub cboUF_GotFocus(Index As Integer)
    SelecionaTudo
    cboCidade(Index).Criterio = ""
End Sub

Private Sub cboUF_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Len(cboUF(Index).Text) = 2 And KeyCode <> vbKeyShift Then
'        SendKeys "{Tab}"
'    End If
End Sub

Private Sub cboUF_LostFocus(Index As Integer)
On Error GoTo trataerro
'    cboUF(Index).Text = UCase(cboUF(Index).Text)
'    If Not ConsisteUF(cboUF(Index).Text) Then
'        cboUF(Index).Text = ""
'        MsgBox "Estado inexistente !", vbExclamation, "Atenção!"
'        cboUF(Index).SetFocus
'        Exit Sub
'    End If
'
    If cboUF(0).Text <> "" And Trim(cboUF(1).Text) = "" Then
       cboUF(1).Text = cboUF(0).Text
       cboUF(1).Text = cboUF(0).Text
    End If
    
    Exit Sub
trataerro:
    Rotina_Erro "cboUF_LostFocus"
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo trataerro
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
    End If
    Exit Sub
trataerro:
    Rotina_Erro "Form_KeyPress"
End Sub

Private Sub Form_Load()
On Error GoTo trataerro
    Top = 0
    Left = 0
    Height = 7590
    Width = 11250
    SSTab1.Tab = 0
    
    grdEndCob.FormatString = "CodCid|Sequência|Endereço|Bairro|Cidade|Cep|Complemento|Telefone|Fax"
    grdEndCob.ColWidth(0) = 0
    grdEndCob.Rows = 1
    
    grdContato.FormatString = "Sequência|Nome|Departamento|Cargo|E-mail|Telefone|Celular|Fax|Aniversário"
    grdContato.Rows = 1
    
    grdLocalColEnt.FormatString = "CodCid|Sequência|Nome|Endereço|Bairro|Cep|Cidade|Complemento|Telefone|Fax|Nome do Contato|E-mail do Contato|Observação"
    grdLocalColEnt.ColWidth(0) = 0
    grdLocalColEnt.Rows = 1
    
    Set CboDesAtvGer.Conexao = Conexao
    CboDesAtvGer.Criterio = 2 ' Indústria
    
    consulta "Select CodUF From Estado"
    For blI = 0 To cboUF.UBound
        Set cboCidade(blI).Conexao = Conexao
        cboUF(blI).Clear
        Rs.MoveFirst
        Do While Not Rs.EOF
            cboUF(blI).AddItem Rs!CodUf
            Rs.MoveNext
        Loop
    Next
    Rs.Close
    Call SetBotoes(True)
    Call OptPessoa_Click(1)
    txtInscrEst.TipodeDados Literal
    txtCep.TipodeDados CEP
    txtCepEndCob.TipodeDados CEP
    txtCepLocColEnt.TipodeDados CEP
    mskAniversarioContato.TipodeDados Data
    txtCep.Texto = ""
    LblCodUsuBlq.Caption = ""
    lblDatBlqGer.Caption = ""
    LblcodMotBlq.Caption = ""
    LblDatLibBlq.Caption = ""
    lblcodUsuLib.Caption = ""
    Exit Sub
trataerro:
    'Rotina_Erro "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next 'Tratamento genérico previsto
    Rs.Close: Set Rs = Nothing
End Sub

Private Sub grdCadGer_DblClick()
On Error GoTo trataerro
    Dim sldoc As String
    sldoc = txtConCadGeral.Text
    If grdCadGer.Row = -1 Then Exit Sub
    'cboCidade(0).query = "Select NomCid, CodCid From Cidade Where CodUF='" & cboUF(0).Text & "' And NomCid like '" & cboCidade(0).Criterio & "%'"
    LimpaCampos
    LimpaGrids
    If Not CarregaTelaCadGeral(grdCadGer.Columns(0).Text) Then Exit Sub
    Verifica_bloqueio grdCadGer.Columns(0).Text
    Exit Sub
trataerro:
    Rotina_Erro "grdCadGer_DblClick"
End Sub

Private Sub grdCadGer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call grdCadGer_DblClick
    End If
End Sub

Private Sub grdContato_DblClick()
    MostraDadosContato
End Sub

Private Sub grdEndCob_DblClick()
    MostraDadosEndCob
End Sub

Private Sub grdLocalColEnt_DblClick()
    MostraDadosLocColEnt
End Sub



Private Sub mskAniversarioContato_GotFocus()
    SelecionaTudo
End Sub

Private Sub OptPessoa_Click(Index As Integer)
    cboUF(0).Enabled = True
    txtCNPJ.Enabled = True
    If Index = 0 Then 'Pessoa Física
        lblCNPJ.Caption = "CPF:"
        txtCNPJ.TipodeDados cpf
        txtInscrEst.Limpar
        txtInscrMun.Text = ""
        txtInscrEst.Enabled = False
        txtInscrMun.Enabled = False
        slps = "F"
    ElseIf Index = 1 Then 'Pessoa Jurídica
        lblCNPJ.Caption = "CNPJ:"
        txtCNPJ.TipodeDados cnpj
        txtInscrEst.Enabled = True
        txtInscrMun.Enabled = True
        txtInscrEst.Limpar
        txtInscrMun.Text = ""
        slps = "J"
    Else 'EXTERIOR
        lblCNPJ.Caption = "Documento:"
        txtCNPJ.TipodeDados Literal
        txtCNPJ.Limpar
        txtCNPJ.Texto = calcula_num_exterior
        txtCNPJ.Enabled = False
        txtInscrEst.Texto = "ISENTO"
        txtInscrEst.Enabled = False
        txtInscrMun.Enabled = False
        slps = "E"
        procuraCombo "EX", cboUF(0)
        cboUF(0).Enabled = False
    End If
End Sub

Private Sub txtCodUserLibBlq_LostFocus()
    MudarTab
End Sub

Private Sub MudarTab()
    If SSTab1.Tab + 1 <= SSTab1.Tabs Then
        SSTab1.Tab = SSTab1.Tab + 1
    End If
End Sub

Private Sub LimpaCampos()
On Error GoTo trataerro
    optPessoa(0).Value = True
    Dim clControles As Control
    For Each clControles In frmCadastro_Geral
        If TypeOf clControles Is TextBox Then
            clControles.Text = ""
        ElseIf TypeOf clControles Is Masked Then
            clControles.Texto = ""
        ElseIf TypeOf clControles Is CheckBox Then
            clControles.Value = 0
        ElseIf TypeOf clControles Is ComboBox Then
            If clControles.Style = 2 Then
                clControles.ListIndex = -1
            Else
                clControles.Text = ""
            End If
        End If
    Next
    mskAniversarioContato.Texto = ""
    For blI = 0 To cboCidade.UBound
        cboCidade(blI).Criterio = ""
    Next blI
    CboDesAtvGer.Criterio = ""
    CboDesAtvGer.Codigo = ""
    SSTab1.TabEnabled(4) = True
    LblCodUsuBlq.Caption = ""
    lblDatBlqGer.Caption = ""
    LblcodMotBlq.Caption = ""
    LblDatLibBlq.Caption = ""
    lblcodUsuLib.Caption = ""
    optPessoa(1).Value = True
    SSTab1.Tab = 0
    Exit Sub
    Resume
trataerro:
    Rotina_Erro "LimpaCampos"
End Sub

Private Sub LimpaGrids()
    grdEndCob.Rows = 1
    grdContato.Rows = 1
    grdLocalColEnt.Rows = 1
End Sub

Private Sub txtConDesCfo_LostFocus()
    CarregaCadGeral
End Sub

Function CarregaCadGeral()
    'Carrega o grid de consulta com os dados cadastrados de acordo com o CPF/CNPJ ou Razão Social informada
On Error GoTo trataerro
    sgQuery = "Select [CNPJ/CPF] = CNPJGer,[Razão Social] = RazSoc,[Nome Reduzido] = NomRed," & _
    "[Endereço] = EndCml,[Bairro] = BaiCml,[Cidade] = Cidade.NomCid,Cep = CepCml,Complemento = ComCml," & _
    "Telefone = TelCml,Fax = FaxCml,[Inscrição Estadual] = InsEstGer,[Inscrição Municipal] = InsMunGer,[Retêm Iss] = FlgRetISS," & _
    "[Retêm INSS] = FlgRetINSS,[Home Page] = HomePGer,[E-mail] = MailGer, [Usuário] = NomUsuSis, [Ultima Alteração] = cadastro_geral.DatUltAlt " & _
    "From CADASTRO_GERAL, CIDADE, Usuario " & _
    "Where cadastro_geral.CodCidCml = cidade.codcid and cadastro_geral.CodUsuSis = Usuario.CodUsuSis"
    If Trim(txtConCadGeral.Text) <> "" Then
       sgQuery = sgQuery & " And " & IIf(IsNumeric(txtConCadGeral.Text), "CNPJGer", "RazSoc") & " like '" & Trim(txtConCadGeral.Text) & "%' Order By " & IIf(IsNumeric(txtConCadGeral.Text), "CNPJGer", "RazSoc")
    End If
    consulta sgQuery
    
    If Rs.EOF = True Then
       MsgBox "Nenhum registro a apresentar.", vbExclamation, "Atenção!"
       Rs.Close
       Set grdCadGer.DataSource = Nothing
    Else
       Set grdCadGer.DataSource = Rs
    End If
    Exit Function
    Resume
trataerro:
    Rotina_Erro "CarregaCadGeral"
End Function

Function CarregaTelaCadGeral(cnpj As String) As Boolean
'Pega as informações do banco de dados e as disponibiliza na tela
On Error GoTo trataerro
    CarregaTelaCadGeral = False
    sgQuery = "Select a.*, b.NomCid, b.CodUF,c.DesAtvGer From CADASTRO_GERAL a, Cidade b ,Atividade c"
    sgQuery = sgQuery & " Where a.CNPJGer = '" & cnpj & "' And a.CodCidCml = b.CodCid and a.codatvger = c.codatvger"
    consulta sgQuery
    blI = 1
    If Not Rs.EOF Then
        With txtCNPJ
            Select Case Rs!CodTipGer
                Case "F"
                    optPessoa(0).Value = True
                    '.TipodeDados cpf
                Case "J"
                    optPessoa(1).Value = True
                    '.TipodeDados cnpj
                Case "E"
                    optPessoa(2).Value = True
                    '.TipodeDados Literal
            End Select
        End With
        'txtCNPJ.TipodeDados IIf(Rs!CodTipGer = "F", cpf, IIf(Rs!CodTipGer = "J", CNPJ, Literal))
        txtCNPJ.Texto = Trim(Rs!cnpjger)
        txtRazSoc.Text = Rs!razsoc
        txtNomRed.Text = Rs!NomRed
        txtEndCom.Text = Rs!EndCml
        txtBaiCml.Text = Rs!BaiCml
        CboDesAtvGer.Criterio = Rs!Desatvger
        cboCidade(0).Codigo = Rs!CodCidCml
        cboUF(0).Text = Rs!CodUf
        cboCidade(0).Criterio = Trim(Rs!NomCid)
        txtCep.Texto = Rs!CepCml
        txtComplEndCml.Text = Rs!ComCml
        txtTelCml.Text = Rs!TelCml
        txtFaxCml.Text = Rs!FaxCml
        txtInscrEst.Texto = Rs!InsEstGer
        txtInscrMun.Text = Rs!InsMunGer
        chkRetISS.Value = IIf(Rs!FlgRetISS = "S", 1, 0)
        chkINSS.Value = IIf(Rs!FlgRetINSS = "S", 1, 0)
        txtHomePage.Text = Rs!HomePGer
        txtEmail.Text = Rs!MailGer
        blI = blI + 1
        CarregaTelaCadGeral = True
        Call CarregaDadosEndCob
        Call CarregaDadosContato
        Call CarregaDadosLocColEnt
        SSTab1.Tab = 0
        Call SetBotoes(False)
    End If
    If Rs.State <> adStateClosed Then Rs.Close
    'grdCadGer.SetFocus
    Exit Function
trataerro:
    Rotina_Erro "CarregaTelaCadGeral"
End Function

Function VerificaCampos(slpessoa As String) As Boolean
    VerificaCampos = True
    If txtCNPJ.CampoDb = "" Then
        MsgBox "CNPJ/CPF em Branco", vbInformation
        txtCNPJ.SetFocus
        VerificaCampos = False
        Exit Function
    ElseIf txtRazSoc.Text = "" Then
        MsgBox "Razão Social em Branco", vbInformation
        txtRazSoc.SetFocus
        VerificaCampos = False
        Exit Function
    ElseIf txtNomRed.Text = "" Then
        MsgBox "Nome Reduzido em Branco", vbInformation
        txtNomRed.SetFocus
        VerificaCampos = False
        Exit Function
    ElseIf txtEndCom.Text = "" Then
        MsgBox "Endereço em Branco", vbInformation
        txtEndCom.SetFocus
        VerificaCampos = False
        Exit Function
    ElseIf cboUF(0).Text = "" Then
        MsgBox "UF em Branco", vbInformation
        cboUF(0).SetFocus
        VerificaCampos = False
        Exit Function
    ElseIf cboCidade(0).Criterio = "" Then
        MsgBox "Cidade em Branco", vbInformation
        cboCidade(0).SetFocus
        VerificaCampos = False
        Exit Function
    ElseIf txtCep.CampoDb = "" Or txtCep.CampoDb = "0" Then
        MsgBox "CEP em Branco", vbInformation
        txtCep.SetFocus
        VerificaCampos = False
        Exit Function
'    ElseIf txtTelCml.Text = "" Then
'        MsgBox "Telefone Comercial em Branco", vbInformation
'        txtTelCml.SetFocus
'        VerificaCampos = False
'        Exit Function
'    ElseIf txtFaxCml.Text = "" Then
'        MsgBox "Telefone Fax em Branco", vbInformation
'        txtFaxCml.SetFocus
'        VerificaCampos = False
'        Exit Function
    ElseIf CboDesAtvGer.Codigo = "" Then
        MsgBox "Atividade  em Branco", vbInformation
        CboDesAtvGer.SetFocus
        VerificaCampos = False
        Exit Function
    End If
    If slpessoa = "J" Then
        If txtInscrEst.CampoDb = "" Then
            MsgBox "Inscrição Estadual em Branco.", vbInformation
            txtInscrEst.SetFocus
            VerificaCampos = False
            Exit Function
        'ElseIf txtInscrMun.Text = "" Then
        '    MsgBox "Incrição Municipal em Branco", vbInformation
        '    txtInscrMun.SetFocus
        '    VerificaCampos = False
        '    Exit Function
        End If
    End If
End Function
#If 0 Then
'Function VerificaCampos_Vei() As Boolean
'    'Verifica preenchimento dos campos na tela
'On Error GoTo Trataerro
'    VerificaCampos = False
'
'    'Tab 0
'
'    If optPessoa(0).Value Or optPessoa(1).Value Then
'        'Testando CNPJ/CPF
'        If RetiraFormatacao(txtCNPJ.Texto) <> "" Then
'            If Not ConfereCPFCGC(IIf(optPessoa(0).Value, True, False), RetiraFormatacao(txtCNPJ.Texto)) Then
'                MsgBox Left(lblCNPJ.Caption, Len(lblCNPJ.Caption) - 1) & " inválido.", vbExclamation, "Aviso!"
'                txtCNPJ.SetFocus
'                Exit Function
'            End If
'        End If
'    Else
'        Exit Function
'    End If
'
'    If RetiraFormatacao(txtInscrEst.Text) <> "" Then
'        If Not ChecaInscrE(cboUF(0).Text, UCase(RetiraFormatacao(txtInscrEst.Text))) Then
'            'Testando Inscrição Estadual
'            MsgBox "Inscrição Estadual inválida.", vbExclamation, "Aviso!"
'            txtInscrEst.SetFocus
'            Exit Function
'        End If
'    Else
'        Exit Function
'    End If
'
'    If Not VerificaCampo(txtCNPJ, lblCNPJ.Caption, 0) Then Exit Function
'    If Not VerificaCampo(txtRazSoc, "Razão Social", 0) Then Exit Function
'    If Not VerificaCampo(txtNomRed, "Nome Reduzido", 0) Then Exit Function
'    If Not VerificaCampo(txtEndCom, "Endereço", 0) Then Exit Function
'    If Not VerificaCampo(txtBaiCml, "Bairro", 0) Then Exit Function
'    If Not VerificaCampo(cboUF(0), "Estado", 0) Then Exit Function
'    If Not VerificaCampo(cboCidade(0), "Cidade", 0) Then Exit Function
'    If Not VerificaCampo(txtCep, "Cep", 0) Then Exit Function
'    If Not VerificaCampo(txtTelCml, "Telefone", 0) Then Exit Function
'    If Not VerificaCampo(txtInscrEst, "Inscrição Estadual", 0) Then Exit Function
'    If Not VerificaCampo(txtInscrMun, "Inscrição Municipal", 0) Then Exit Function
'
'    VerificaCampos = True
'    Exit Function
'    Resume
'trataerro:
'    Rotina_Erro "VerificaCampos"
'End Function
#End If

Function Gravar()
    On Error GoTo trataerro
    
    'Cadastro Geral
    Conexao.BeginTrans
    
    Set Cmd = New Command
    With Cmd
        .CommandText = "{call MNCADASTRO_GERAL (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
        .CommandType = adCmdText
        .ActiveConnection = Conexao
        .Parameters.Refresh
        .Parameters(0).Value = txtCNPJ.CampoDb  'Trim(RetiraFormatacao(txtCNPJ.Texto))
        .Parameters(1).Value = Trim(txtRazSoc.Text)
        .Parameters(2).Value = Trim(txtNomRed.Text)
        .Parameters(3).Value = Trim(txtEndCom.Text)
        .Parameters(4).Value = Trim(txtBaiCml.Text)
        'consulta "Select CodCid From Cidade Where CodUF='" & cboUF(0) & "' And NomCid='" & cboCidade(0).Criterio & "'"
        .Parameters(5).Value = cboCidade(0).Codigo 'Rs!codcid
        .Parameters(6).Value = RetiraFormatacao(txtCep.Texto)
        .Parameters(7).Value = Trim(txtComplEndCml.Text)
        .Parameters(8).Value = Trim(txtTelCml.Text)
        .Parameters(9).Value = Trim(txtFaxCml.Text)
        .Parameters(10).Value = txtInscrEst.CampoDb
        .Parameters(11).Value = Trim(txtInscrMun.Text)
        .Parameters(12).Value = IIf(chkRetISS.Value, "S", "N")
        .Parameters(13).Value = IIf(chkINSS.Value, "S", "N")
        .Parameters(14).Value = Trim(txtHomePage.Text)
        .Parameters(15).Value = Trim(txtEmail.Text)
        .Parameters(16).Value = slps
        .Parameters(17).Value = "" 'Flag usado para saber a origem dos dados
        .Parameters(18).Value = LgCodUsuSis
        .Parameters(19).Value = CboDesAtvGer.Codigo
        .Parameters(20).Value = sgFlagOper
'        .Parameters(16).Value = IIf(lblMotBloq.Caption = "", Null, lblMotBloq.Caption)
'        .Parameters(17).Value = IIf(lblDataMotBlq.Caption = "", Null, lblDataMotBlq.Caption)
'        .Parameters(18).Value = IIf(lblCodUserEfetBlq.Caption = "", Null, lblCodUserEfetBlq.Caption)
'        .Parameters(20).Value = IIf(lblDataLibBlq.Caption = "", Null, lblDataLibBlq.Caption)
'        .Parameters(21).Value = IIf(lblCodUserLibBlq.Caption = "", Null, lblCodUserLibBlq.Caption)

     End With
 
    Set Rs = Cmd.Execute
    Set Rs = Nothing
    Set Cmd = Nothing
    
    'Endereço/Cobrança
    Grava_End_Cob
    
    'Contato
    Grava_Contato
    
    'Local Coleta/Entrega
    Grava_Ent_Col
    
    Conexao.CommitTrans
    Exit Function
    Resume
trataerro:
    Rotina_Erro "Gravar"
    'Conexao.RollbackTrans
    Set Rs = Nothing
    Set Cmd = Nothing
    Set grdCadGer.DataSource = Nothing
End Function

Private Sub CarregaDadosEndCob()
    'Manda os dados do Grid para serem editados
On Error GoTo trataerro
    consulta "Select ENDCOB_CADASTRO_GERAL.*, Cidade.NomCid,Cidade.CodCid From ENDCOB_CADASTRO_GERAL, Cidade Where ENDCOB_CADASTRO_GERAL.CNPJGer ='" & txtCNPJ.Texto & "' And ENDCOB_CADASTRO_GERAL.CodCidCob = Cidade.CodCid"
    grdEndCob.Rows = 1
    blI = 1
    Do While Not Rs.EOF
        grdEndCob.Rows = grdEndCob.Rows + 1
        grdEndCob.TextMatrix(blI, 0) = Rs!CodCidCob
        grdEndCob.TextMatrix(blI, 1) = Rs!SeqCob
        grdEndCob.TextMatrix(blI, 2) = Trim(Rs!EndCob) 'txtEndereçoEndCob.Text
        grdEndCob.TextMatrix(blI, 3) = Trim(Rs!BaiCob) 'txtBaiEndCob.Text
        grdEndCob.TextMatrix(blI, 4) = Trim(Rs!NomCid) 'cboCidade(1).Criterio
        grdEndCob.TextMatrix(blI, 5) = Rs!CepCob 'txtCepEndCob.Texto
        grdEndCob.TextMatrix(blI, 6) = Trim(Rs!ComCob) 'txtComplEndCob.Text
        grdEndCob.TextMatrix(blI, 7) = Trim(Rs!TelCob) 'txtTelEndCob.Text
        grdEndCob.TextMatrix(blI, 8) = Trim(Rs!FaxCob) 'txtFaxEndCob.Text
        
        blI = blI + 1
        Rs.MoveNext
    Loop
    AjustaColWidth grdEndCob
    Rs.Close
    
    'envia primeiro registro para campos da tela
    If grdEndCob.Rows > 1 Then
        txtEndereçoEndCob.Text = grdEndCob.TextMatrix(1, 2)
        txtBaiEndCob.Text = grdEndCob.TextMatrix(1, 3)
        cboCidade(1).Criterio = grdEndCob.TextMatrix(1, 4)
        txtCepEndCob.Texto = Trim(grdEndCob.TextMatrix(1, 5))
        txtComplEndCob.Text = grdEndCob.TextMatrix(1, 6)
        txtTelEndCob.Text = grdEndCob.TextMatrix(1, 7)
        txtFaxEndCob.Text = grdEndCob.TextMatrix(1, 8)
        
        consulta "Select CodUF From Cidade Where CodCid = '" & grdEndCob.TextMatrix(1, 0) & "'"
        cboUF(1).Text = Rs!CodUf
        Rs.Close
    End If
    Exit Sub
trataerro:
    Rotina_Erro "CarregaDadosEndCob"
End Sub

Private Sub MostraDadosEndCob()
On Error GoTo trataerro
    txtEndereçoEndCob.Text = grdEndCob.TextMatrix(grdEndCob.RowSel, 2)
    txtBaiEndCob.Text = grdEndCob.TextMatrix(grdEndCob.RowSel, 3)
    cboCidade(1).Criterio = grdEndCob.TextMatrix(grdEndCob.RowSel, 4)
    txtCepEndCob.Texto = Trim(grdEndCob.TextMatrix(grdEndCob.RowSel, 5))
    txtComplEndCob.Text = grdEndCob.TextMatrix(grdEndCob.RowSel, 6)
    txtTelEndCob.Text = grdEndCob.TextMatrix(grdEndCob.RowSel, 7)
    txtFaxEndCob.Text = grdEndCob.TextMatrix(grdEndCob.RowSel, 8)
    
    consulta "Select CodUF From Cidade Where CodCid='" & grdEndCob.TextMatrix(grdEndCob.RowSel, 0) & "'"
    cboUF(1).Text = Rs!CodUf
    Rs.Close
    Exit Sub
    Resume
trataerro:
    Rotina_Erro "MostraDadosEndCob"
End Sub

Private Sub CarregaDadosContato()
    'Manda os dados do Grid para serem editados
On Error GoTo trataerro
    consulta "Select * From CONTATO_CADASTRO_GERAL Where CNPJGer='" & txtCNPJ.Texto & "'"
    grdContato.Rows = 1
    blI = 1
    Do While Not Rs.EOF
        grdContato.Rows = grdContato.Rows + 1
        grdContato.TextMatrix(blI, 0) = Rs!SeqCto
        grdContato.TextMatrix(blI, 1) = Rs!NomCto 'txtNomeContato.Text
        grdContato.TextMatrix(blI, 2) = Rs!DepCto 'txtDeptoContato.Text
        grdContato.TextMatrix(blI, 3) = Rs!CgoCto 'txtCargoContato.Text
        grdContato.TextMatrix(blI, 4) = Rs!MailCto 'txtEmailContato.Text
        grdContato.TextMatrix(blI, 5) = Rs!TelCto 'txtTelContato.Text
        grdContato.TextMatrix(blI, 6) = Rs!CelCto 'txtCelContato.Text
        grdContato.TextMatrix(blI, 7) = Rs!FaxCto 'txtFaxContato.Text
        grdContato.TextMatrix(blI, 8) = IIf(IsNull(Rs!DatAnvCto), "__/__/____", Rs!DatAnvCto) 'mskAniversarioContato.Texto
        Rs.MoveNext
        blI = blI + 1
    Loop
    AjustaColWidth grdContato
    Rs.Close
    
    'carrega primeiro registro para os campos da tela
    If grdContato.Rows > 1 Then
        txtNomeContato.Text = grdContato.TextMatrix(1, 1)
        txtDeptoContato.Text = grdContato.TextMatrix(1, 2)
        txtCargoContato.Text = grdContato.TextMatrix(1, 3)
        txtEmailContato.Text = grdContato.TextMatrix(1, 4)
        txtTelContato.Text = grdContato.TextMatrix(1, 5)
        txtCelContato.Text = grdContato.TextMatrix(1, 6)
        txtFaxContato.Text = grdContato.TextMatrix(1, 7)
        mskAniversarioContato.Texto = grdContato.TextMatrix(1, 8)
    End If
    
    Exit Sub
    Resume
trataerro:
    Rotina_Erro "CarregaDadosContato"
End Sub

Private Sub MostraDadosContato()
On Error GoTo trataerro
    'Manda os dados do Grid para serem editados
    txtNomeContato.Text = grdContato.TextMatrix(grdContato.RowSel, 1)
    txtDeptoContato.Text = grdContato.TextMatrix(grdContato.RowSel, 2)
    txtCargoContato.Text = grdContato.TextMatrix(grdContato.RowSel, 3)
    txtEmailContato.Text = grdContato.TextMatrix(grdContato.RowSel, 4)
    txtTelContato.Text = grdContato.TextMatrix(grdContato.RowSel, 5)
    txtCelContato.Text = grdContato.TextMatrix(grdContato.RowSel, 6)
    txtFaxContato.Text = grdContato.TextMatrix(grdContato.RowSel, 7)
    mskAniversarioContato.Texto = grdContato.TextMatrix(grdContato.RowSel, 8)
    Exit Sub
trataerro:
    Rotina_Erro "MostraDadosContato"
End Sub

Private Sub CarregaDadosLocColEnt()
    'Manda os dados do Grid para serem editados
On Error GoTo trataerro
    consulta "Select LOCAL_COLETA_ENTREGA.*, Cidade.NomCid From LOCAL_COLETA_ENTREGA, Cidade Where Local_Coleta_Entrega.CNPJGer = '" & Trim(txtCNPJ.Texto) & "' And LOCAL_COLETA_ENTREGA.CodCid = Cidade.CodCid"
    grdLocalColEnt.Rows = 1
    blI = 1
    Do While Not Rs.EOF
        grdLocalColEnt.Rows = grdLocalColEnt.Rows + 1
        grdLocalColEnt.TextMatrix(blI, 0) = Rs!codcid
        grdLocalColEnt.TextMatrix(blI, 1) = Rs!SeqLoc
        grdLocalColEnt.TextMatrix(blI, 2) = Rs!NomLoc 'txtNomeLocColEnt.Text
        grdLocalColEnt.TextMatrix(blI, 3) = Rs!EndLoc 'txtEndLocColEnt.Text
        grdLocalColEnt.TextMatrix(blI, 4) = Rs!BaiLoc 'txtBaiLocColEnt.Text
        grdLocalColEnt.TextMatrix(blI, 5) = Rs!CepLoc 'txtCepLocColEnt.Texto
        grdLocalColEnt.TextMatrix(blI, 6) = Rs!NomCid 'cboCidade(2).Criterio
        grdLocalColEnt.TextMatrix(blI, 7) = Rs!ComEndLoc 'txtComplLocColEnt.Text
        grdLocalColEnt.TextMatrix(blI, 8) = Rs!TelLoc 'txtTelLocColEnt.Text
        grdLocalColEnt.TextMatrix(blI, 9) = Rs!FaxLoc 'txtFaxLocColEnt.Text
        grdLocalColEnt.TextMatrix(blI, 10) = IIf(Not IsNull(Rs!NomCto), Rs!NomCto, "") 'txtNomeContatoLocColEnt.Text)
        grdLocalColEnt.TextMatrix(blI, 11) = IIf(Not IsNull(Rs!MailCto), Rs!MailCto, "") 'txtEmailLocColEnt.Text
        grdLocalColEnt.TextMatrix(blI, 12) = IIf(Not IsNull(Rs!ObsLoc), Rs!ObsLoc, "") 'txtObsLocColEnt.Text
        Rs.MoveNext
        blI = blI + 1
    Loop
    AjustaColWidth grdLocalColEnt
    Rs.Close
    
    'carrega primeiro registro para campos da tela
    If grdLocalColEnt.Rows > 1 Then
        txtNomeLocColEnt.Text = grdLocalColEnt.TextMatrix(1, 2)
        txtEndLocColEnt.Text = grdLocalColEnt.TextMatrix(1, 3)
        txtBaiLocColEnt.Text = grdLocalColEnt.TextMatrix(1, 4)
        txtCepLocColEnt.Texto = grdLocalColEnt.TextMatrix(1, 5)
        txtComplLocColEnt.Text = grdLocalColEnt.TextMatrix(1, 7)
        txtTelLocColEnt.Text = grdLocalColEnt.TextMatrix(1, 8)
        txtFaxLocColEnt.Text = grdLocalColEnt.TextMatrix(1, 9)
        txtNomeContatoLocColEnt.Text = grdLocalColEnt.TextMatrix(1, 10)
        txtEmailLocColEnt.Text = grdLocalColEnt.TextMatrix(1, 11)
        txtObsLocColEnt.Text = grdLocalColEnt.TextMatrix(1, 12)
        consulta "Select CodUF From Cidade Where CodCid='" & grdLocalColEnt.TextMatrix(1, 0) & "'"
        cboUF(2).Text = Rs!CodUf
        cboCidade(2).Criterio = Trim(grdLocalColEnt.TextMatrix(1, 6))
        Rs.Close
        btoModificar(2).Enabled = trava_coleta_entrega(Trim(txtCNPJ.Texto), grdLocalColEnt.TextMatrix(grdLocalColEnt.Row, 1))
    End If
    Exit Sub
trataerro:
    Rotina_Erro "CarregaDadosLocColEnt"
End Sub

Private Sub MostraDadosLocColEnt()
    'Manda os dados do Grid para serem editados
On Error GoTo trataerro
    txtNomeLocColEnt.Text = grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 2)
    txtEndLocColEnt.Text = grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 3)
    txtBaiLocColEnt.Text = grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 4)
    txtCepLocColEnt.Texto = grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 5)
    txtComplLocColEnt.Text = grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 7)
    txtTelLocColEnt.Text = grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 8)
    txtFaxLocColEnt.Text = grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 9)
    txtNomeContatoLocColEnt.Text = grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 10)
    txtEmailLocColEnt.Text = grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 11)
    txtObsLocColEnt.Text = grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 12)
    consulta "Select CodUF From Cidade Where CodCid='" & Val(grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 0)) & "'"
    cboUF(2).Text = Rs!CodUf
    cboCidade(2).Criterio = Trim(grdLocalColEnt.TextMatrix(grdLocalColEnt.RowSel, 6))
    Rs.Close
    btoModificar(2).Enabled = trava_coleta_entrega(Trim(txtCNPJ.Texto), grdLocalColEnt.TextMatrix(grdLocalColEnt.Row, 1))
    Exit Sub
trataerro:
    Rotina_Erro "MostraDadosLocColEnt"
End Sub
Private Sub SetBotoes(vOpc As Boolean)
On Error Resume Next
'Se o valor passado como parâmetro for TRUE é uma inclusão
'Se o valor passado for FALSE é uma alteração
   SSTab1.TabEnabled(4) = vOpc
   'Habilita campos chaves
   txtCNPJ.Enabled = vOpc
   'Habilita os botões
   BtoIncluir.Enabled = vOpc
   BtoAlterar.Enabled = Not vOpc
   Btoexcluir.Enabled = Not vOpc
End Sub
Private Sub AjustaColWidth(grid As MSFlexGrid)
    'Colocando o tamanho da coluna do grid de acordo com o tamanho do texto
On Error GoTo trataerro
    Dim Indice As Integer, Linhas As Integer
    For Indice = 1 To grid.Cols - 1
        For Linhas = 1 To grid.Rows - 1
            If TextWidth(grid.TextMatrix(grid.Rows - 1, Indice)) + 200 >= grid.ColWidth(Indice) Then
                grid.ColWidth(Indice) = TextWidth(grid.TextMatrix(grid.Rows - 1, Indice)) + 200
            End If
        Next Linhas
    Next Indice
    Exit Sub
trataerro:
    Rotina_Erro "AjustaColWidth"
End Sub

Public Function RetiraFormatacao(Texto As String)
'Retira caracteres especiais deixando somente números
On Error GoTo trataerro
    If Texto = "" Then Exit Function
    If UCase(Trim(Texto)) = "ISENTO" Then RetiraFormatacao = Trim(Texto): Exit Function
    Dim blIndex As Byte
    RetiraFormatacao = vbNullString
    For blIndex = 1 To Len(Texto)
        If Asc(Mid(Texto, blIndex, 1)) >= 48 And Asc(Mid(Texto, blIndex, 1)) <= 57 Then
            RetiraFormatacao = RetiraFormatacao & Mid(Texto, blIndex, 1)
        End If
    Next blIndex
    Exit Function
trataerro:
    Rotina_Erro "RetiraFormatacao"
End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo trataerro
    Btoexcluir.Enabled = IIf(SSTab1.Tab = 0 And txtCNPJ.Texto <> "", True, False)
            
    'If PreviousTab = 0 And txtCNPJ.Texto <> "" And btoIncluir.Enabled = False Then
    If PreviousTab = 0 And txtCNPJ.CampoDb <> "" Then
        If txtEndereçoEndCob.Text = "" Then
            cboUF(1).Text = cboUF(0).Text
            txtEndereçoEndCob.Text = txtEndCom.Text
            txtBaiEndCob.Text = txtBaiCml.Text
            cboCidade(1).Codigo = cboCidade(0).Codigo
            cboCidade(1).Criterio = cboCidade(0).Criterio
            txtCepEndCob.Texto = txtCep.Texto
            txtComplEndCob.Text = txtComplEndCml.Text
            txtTelEndCob.Text = txtTelCml.Text
            txtFaxEndCob.Text = txtFaxCml.Text
        End If
            
        If txtEndLocColEnt.Text = "" Then
            cboUF(1).Text = cboUF(0).Text
            cboUF(2).Text = cboUF(0).Text
            cboCidade(1).Criterio = cboCidade(0).Criterio
            cboCidade(2).Criterio = cboCidade(0).Criterio
            txtNomeLocColEnt.Text = txtNomRed.Text
            txtEndLocColEnt.Text = txtEndCom.Text
            txtBaiLocColEnt.Text = txtBaiCml.Text
            txtCepLocColEnt.Texto = txtCep.Texto
            txtComplLocColEnt.Text = txtComplEndCml.Text
            txtTelLocColEnt.Text = txtTelCml.Text
            txtFaxLocColEnt.Text = txtFaxCml.Text
            txtEmailLocColEnt.Text = txtEmail.Text
        End If
        If BtoIncluir.Enabled = True Then
            If grdEndCob.Rows = 1 Then Call btoAdicionar_Click(0)
            If grdLocalColEnt.Rows = 1 Then Call btoAdicionar_Click(2)
        End If
    End If
    If SSTab1.Tab = 4 Then txtConCadGeral.SetFocus
    Exit Sub
trataerro:
    Rotina_Erro "SSTab1_Click"
End Sub


Private Sub txtBaiCml_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtBaiCml_LostFocus()
    If Trim(txtBaiCml.Text) <> "" Then
       txtBaiEndCob.Text = Trim(txtBaiCml.Text)
    End If
End Sub

Private Sub txtBaiEndCob_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtBaiLocColEnt_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtCargoContato_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtCelContato_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtCep_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtCep_LostFocus()
    If RetiraFormatacao(txtCep.Texto) <> "" Then
       txtCepEndCob.Texto = txtCep.Texto
    End If
End Sub

Private Sub txtCepEndCob_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtCepLocColEnt_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtCNPJ_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtCNPJ_LostFocus()
    If txtCNPJ.CampoDb <> "" Then
        Verifica_bloqueio txtCNPJ.CampoDb
        If Not CarregaTelaCadGeral(txtCNPJ.CampoDb) Then Exit Sub
    End If
End Sub

Private Sub txtComplEndCml_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtComplEndCml_LostFocus()
    If Trim(txtComplEndCml.Text) <> "" Then
       txtComplEndCob.Text = Trim(txtComplEndCml.Text)
    End If
End Sub

Private Sub txtComplEndCob_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtComplLocColEnt_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtConCadGeral_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtConCadGeral_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtConCadGeral_LostFocus
    End If
End Sub

Private Sub txtConCadGeral_LostFocus()
    If Screen.ActiveControl.Name = "SSTab1" Or txtConCadGeral.Text = "" Then Exit Sub
    Call CarregaCadGeral
End Sub

Private Function VerificaCampo(Campo As Control, Msg As String, bTab As Byte) As Boolean
'Verifica se o campo requerido foi devidamente preenchido
    VerificaCampo = False
    If TypeOf Campo Is TextBox Or TypeOf Campo Is ComboBox Then
        If Trim(Campo.Text) = "" Then
            MsgBox "Informe o(a) " & Msg & ".", vbExclamation, "Atenção!"
            Campo.SetFocus
            SSTab1.Tab = bTab
            Exit Function
        End If
    ElseIf TypeOf Campo Is Combo_DB Then
        If Trim(Campo.Criterio) = "" Then
            MsgBox "Informe o(a) " & Msg & ".", vbExclamation, "Atenção!"
            Campo.SetFocus
            SSTab1.Tab = bTab
            Exit Function
        End If
    ElseIf TypeOf Campo Is Masked Then
        If Trim(Campo.Texto) = "" Then
            MsgBox "Informe o(a) " & Msg & ".", vbExclamation, "Atenção!"
            Campo.SetFocus
            SSTab1.Tab = bTab
            Exit Function
        End If
    End If
    VerificaCampo = True
End Function

Private Sub txtDeptoContato_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtEmail_GotFocus()
    SelecionaTudo
End Sub


Private Sub txtEmail_LostFocus()
    SSTab1.Tab = 1
    txtEndereçoEndCob.SetFocus
End Sub

Private Sub txtEmailContato_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtEmailLocColEnt_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtEndCom_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtEndCom_LostFocus()
    If Trim(txtEndCom.Text) <> "" Then
       txtEndereçoEndCob.Text = Trim(txtEndCom.Text)
    End If
End Sub

Private Sub txtEndereçoEndCob_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtEndLocColEnt_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtFaxCml_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtFaxCml_LostFocus()
    If Trim(txtFaxCml.Text) <> "" Then
       txtFaxEndCob.Text = Trim(txtFaxCml.Text)
    End If
End Sub

Private Sub txtFaxContato_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtFaxEndCob_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtFaxLocColEnt_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtHomePage_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtInscrEst_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtInscrEst_LostFocus()
    If txtInscrEst.Texto = "" Then
        txtInscrEst.Texto = "ISENTO"
    ElseIf Not ChecaInscrE(cboUF(0), txtInscrEst.Texto) Then
        MsgBox "Incrição Estadual Inválida", vbInformation
        txtInscrEst.SetFocus
    End If
End Sub

Private Sub txtInscrMun_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtNomeContato_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtNomeContatoLocColEnt_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtNomeLocColEnt_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtNomRed_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtRazSoc_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtTelCml_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtTelCml_LostFocus()
    If Trim(txtTelCml.Text) <> "" Then
       txtTelEndCob.Text = Trim(txtTelCml.Text)
    End If
End Sub

Private Sub txtTelContato_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtTelEndCob_GotFocus()
    SelecionaTudo
End Sub

Private Sub txtTelLocColEnt_GotFocus()
    SelecionaTudo
End Sub
Private Sub Verifica_bloqueio(sldoc As String)
    
    sgQuery = "select 'Usuário Bloq'= Blq.NomUsuSis,'Motivo Bloqueio' = DesMotBlq,'Data Bloqueio'= DatBlqGer,'Data Liberacao' = a.DatLibBlq,'Usuario Desbloq' = Dbq.NomUsuSis"
    sgQuery = sgQuery & " From bloqueio_cadastro_geral a,motivo_bloqueio b,Cadastro_geral c,Usuario blq,Usuario dbq Where a.CodMotBlq = b.CodMotBlq and c.CNPJGer = a.CNPJGer and"
    sgQuery = sgQuery & " a.codusublq = blq.codususis and a.codusulib *= dbq.codususis and a.DatBlqGer =  (select Max(DatBlqGer) from bloqueio_cadastro_geral d where d.cnpjger = a.cnpjger)   and"
    sgQuery = sgQuery & " a.cnpjger like '" & sldoc & "%' order by razsoc"
    consulta sgQuery
    If Not Rs.EOF Then
        LblCodUsuBlq.Caption = Trim(Rs("Usuário Bloq"))
        lblDatBlqGer.Caption = Trim(Rs("Data Bloqueio"))
        LblcodMotBlq.Caption = Trim(Rs("Motivo Bloqueio"))
        LblDatLibBlq.Caption = IIf(Not IsNull(Rs("Data Liberacao")), Rs("Data Liberacao"), "")
        lblcodUsuLib.Caption = IIf(Not IsNull(Rs("Usuario Desbloq")), Rs("Usuario Desbloq"), "")
    Else
        LblCodUsuBlq.Caption = ""
        lblDatBlqGer.Caption = ""
        LblcodMotBlq.Caption = ""
        LblDatLibBlq.Caption = ""
        lblcodUsuLib.Caption = ""
    End If
    
End Sub
Private Sub Grava_End_Cob()
    For blI = 1 To grdEndCob.Rows - 1
        consulta "Select CNPJGer From ENDCOB_CADASTRO_GERAL Where CNPJGer='" & txtCNPJ.Texto & "' And SeqCob='" & Val(grdEndCob.TextMatrix(blI, 1)) & "'"
        sgFlagOper = IIf(Rs.RecordCount > 0, "A", "I")
        Set Cmd = New Command
        With Cmd
            .CommandText = "{call MNENDCOB_CADASTRO_GERAL (?,?,?,?,?,?,?,?,?,?,?)}"
            .CommandType = adCmdText
            .ActiveConnection = Conexao
            .Parameters.Refresh
            'CodCid|Sequência|Endereço|Bairro|Cidade|Cep|Complemento|Telefone|Fax
            '(CNPJGer, SeqCob, EndCob, BaiCob, CodCidCob, CepCob, ComCob, TelCob, FaxCob, CodUsuSis, DatUltAlt)
            .Parameters(0).Value = txtCNPJ.CampoDb
            .Parameters(1).Value = Format(grdEndCob.TextMatrix(blI, 1), "00000") 'Sequência
            .Parameters(2).Value = grdEndCob.TextMatrix(blI, 2) 'txtEndereçoEndCob.Text
            .Parameters(3).Value = grdEndCob.TextMatrix(blI, 3) 'txtBaiEndCob.Text
            .Parameters(4).Value = grdEndCob.TextMatrix(blI, 0)
            .Parameters(5).Value = RetiraFormatacao(grdEndCob.TextMatrix(blI, 5)) 'txtCepEndCob.Texto
            .Parameters(6).Value = grdEndCob.TextMatrix(blI, 6) 'txtComplEndCob.Text
            .Parameters(7).Value = grdEndCob.TextMatrix(blI, 7) 'txtTelEndCob.Text
            .Parameters(8).Value = grdEndCob.TextMatrix(blI, 8) 'txtFaxEndCob.Text
            .Parameters(9).Value = LgCodUsuSis
            .Parameters(10).Value = sgFlagOper
        End With
        Set Rs = Cmd.Execute
        Set Rs = Nothing
        Set Cmd = Nothing
            
    Next blI
End Sub

Private Sub Grava_Contato()
    For blI = 1 To grdContato.Rows - 1
    
        consulta "Select CNPJGer From CONTATO_CADASTRO_GERAL Where CNPJGer='" & txtCNPJ.Texto & "' And SeqCto='" & grdContato.TextMatrix(blI, 0) & "'"
        sgFlagOper = IIf(Rs.RecordCount > 0, "A", "I")
        
        Set Cmd = New Command
        With Cmd
            .CommandText = "{call MNCONTATO_CADASTRO_GERAL (?,?,?,?,?,?,?,?,?,?,?,?)}"
            .CommandType = adCmdText
            .ActiveConnection = Conexao
            .Parameters.Refresh
            'Sequência|Nome|Departamento|Cargo|E-mail|Telefone|Celular|Fax|Aniversário
            '(CNPJGer,SeqCto,NomCto,DepCto,CgoCto,MailCto,TelCto,CelCto,FaxCto,DatAnvCto,CodUsuSis,DatUltAlt)
            .Parameters(0).Value = txtCNPJ.CampoDb
            .Parameters(1).Value = Format(grdContato.TextMatrix(blI, 0), "00000")  'Sequência
            .Parameters(2).Value = Trim(grdContato.TextMatrix(blI, 1)) 'txtNomeContato.Text
            .Parameters(3).Value = Trim(grdContato.TextMatrix(blI, 2)) 'txtDeptoContato.Text
            .Parameters(4).Value = Trim(grdContato.TextMatrix(blI, 3)) 'txtCargoContato.Text
            .Parameters(5).Value = Trim(grdContato.TextMatrix(blI, 4)) 'txtEmailContato.Text
            .Parameters(6).Value = Trim(grdContato.TextMatrix(blI, 5)) 'txtTelContato.Text
            .Parameters(7).Value = Trim(grdContato.TextMatrix(blI, 6)) 'txtCelContato.Text
            .Parameters(8).Value = Trim(grdContato.TextMatrix(blI, 7)) 'txtFaxContato.Text
            .Parameters(9).Value = IIf(RetiraFormatacao(grdContato.TextMatrix(blI, 8)) = "", Null, grdContato.TextMatrix(blI, 8)) 'mskAniversarioContato.Texto
            .Parameters(10).Value = LgCodUsuSis
            .Parameters(11).Value = sgFlagOper
         End With
        
         Set Rs = Cmd.Execute
         Set Rs = Nothing
         Set Cmd = Nothing
             
    Next blI
End Sub

Private Sub Grava_Ent_Col()
    For blI = 1 To grdLocalColEnt.Rows - 1
    
        consulta "Select CNPJGer From LOCAL_COLETA_ENTREGA Where CNPJGer='" & txtCNPJ.Texto & "' And SeqLoc='" & grdLocalColEnt.TextMatrix(blI, 1) & "'"
        sgFlagOper = IIf(Rs.RecordCount > 0, "A", "I")
    
        Set Cmd = New Command
        With Cmd
            .CommandText = "{call MNLOCAL_COLETA_ENTREGA (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
            .CommandType = adCmdText
            .ActiveConnection = Conexao
            .Parameters.Refresh
            'CodCid|Sequência|Nome|Endereço|Bairro|Cep|Cidade|Complemento|Telefone|Fax|Nome do Contato|E-mail do Contato|Observação
            '(CNPJGer,SeqLoc,NomLoc,EndLoc,BaiLoc,CodCid,CepLoc,ComEndLoc,TelLoc,FaxLoc,NomCto,MailCto,ObsLoc,CodUsuSis,DatUltAlt)
            .Parameters(0).Value = txtCNPJ.CampoDb
            .Parameters(1).Value = Format(grdLocalColEnt.TextMatrix(blI, 1), "00000")  'Sequência
            .Parameters(2).Value = grdLocalColEnt.TextMatrix(blI, 2) 'txtNomeLocColEnt.Text
            .Parameters(3).Value = grdLocalColEnt.TextMatrix(blI, 3) 'txtEndLocColEnt.Text
            .Parameters(4).Value = grdLocalColEnt.TextMatrix(blI, 4) 'txtBaiLocColEnt.Text
            .Parameters(5).Value = grdLocalColEnt.TextMatrix(blI, 0)
            .Parameters(6).Value = RetiraFormatacao(grdLocalColEnt.TextMatrix(blI, 5))  'txtCepLocColEnt.Texto
            .Parameters(7).Value = grdLocalColEnt.TextMatrix(blI, 7) 'txtComplLocColEnt.Text
            .Parameters(8).Value = grdLocalColEnt.TextMatrix(blI, 8) 'txtTelLocColEnt.Text
            .Parameters(9).Value = grdLocalColEnt.TextMatrix(blI, 9) 'txtFaxLocColEnt.Text
            .Parameters(10).Value = grdLocalColEnt.TextMatrix(blI, 10) 'txtNomeContatoLocColEnt.Text
            .Parameters(11).Value = grdLocalColEnt.TextMatrix(blI, 11) 'txtEmailLocColEnt.Text
            .Parameters(12).Value = grdLocalColEnt.TextMatrix(blI, 12) 'txtObsLocColEnt.Text
            .Parameters(13).Value = LgCodUsuSis
            .Parameters(14).Value = sgFlagOper
        End With
        
        Set Rs = Cmd.Execute
        Set Rs = Nothing
        Set Cmd = Nothing
    
    Next blI

End Sub
Private Function trava_coleta_entrega(cnpj As String, sequencia As Integer) As Boolean
    'FUNCAO PARA TRAVAR (FALSE) ALTERAÇÕES NOS DADOS DE COLETA/ENTREGA DOS CLIENTES QUE POSSUEM
    '-ITEM_PROGRAMACAO_CARGA
    '-NOTA_CLIENTE
    '-NOTA_CTRC
    'ROTINA CRIADA POR CARLOS FABIANO EM 23/12/2003
    trava_coleta_entrega = False
    '--programaçao de carga
    sgQuery = "select CNPJRem,SeqLocCol From programacao_carga  Where CNPJRem ='" & cnpj & "' And SeqLocCol =" & sequencia
    consulta sgQuery
    If Not Rs.EOF Then
        Exit Function
    End If
    '--item programacao de carga
    sgQuery = "select CNPJCli,SeqLocEnt from item_programacao_carga  Where CNPJCli ='" & cnpj & "' and SeqLocEnt =" & sequencia
    consulta sgQuery
    If Not Rs.EOF Then
        Exit Function
    End If
    '--nota cliente (rementente)
    sgQuery = " select CNPJDest,SeqLocEnt from nota_cliente  Where CNPJDest ='" & cnpj & "' And SeqLocEnt =" & sequencia
    consulta sgQuery
    If Not Rs.EOF Then
        Exit Function
    End If
    '--nota cliente (destinatario)
    sgQuery = "select CNPJRem,SeqLocCol from nota_cliente Where CNPJRem ='" & cnpj & "' And SeqLocCol =" & sequencia
    consulta sgQuery
    If Not Rs.EOF Then
        Exit Function
    End If
    '--nota CTRC (remetente)
    sgQuery = "Select CNPJRem,SeqLocCol from nota_ctrc Where CNPJRem ='" & cnpj & "' AND SeqLocCol =" & sequencia
    consulta sgQuery
    If Not Rs.EOF Then
        Exit Function
    End If
    '--nota CTRC (Destinatario)
    sgQuery = "Select CNPJDest,SeqLocEnt from nota_ctrc Where CNPJDest ='" & cnpj & "' And SeqLocEnt =" & sequencia
    consulta sgQuery
    If Not Rs.EOF Then
        Exit Function
    End If
    trava_coleta_entrega = True
End Function
Private Function calcula_num_exterior() As String
    Dim rstemp As ADODB.Recordset
    sgQuery = "select cadger =  max(cnpjger)+1 from cadastro_geral where CodTipGer = 'E'"
    Set rstemp = Conexao.Execute(sgQuery)
    If Not rstemp.EOF Then
        calcula_num_exterior = Format(rstemp("CadGer"), "00000")
    Else
        calcula_num_exterior = "00001"
    End If
End Function

