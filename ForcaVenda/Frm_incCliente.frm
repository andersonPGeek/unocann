VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.ocx"
Begin VB.Form Frm_IncCliente 
   Caption         =   "Cadastro de cliente"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3240
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   10455
      Begin VB.Frame Frame2 
         BackColor       =   &H80000018&
         Caption         =   "Status do Cliente"
         Height          =   735
         Left            =   4680
         TabIndex        =   56
         Top             =   240
         Width           =   4095
         Begin VB.ComboBox Cbo_NomeStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Frm_incCliente.frx":0000
            Left            =   120
            List            =   "Frm_incCliente.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   240
            Width           =   3855
         End
         Begin VB.ComboBox Cbo_CodigoStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "Frm_incCliente.frx":0004
            Left            =   120
            List            =   "Frm_incCliente.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   240
            Visible         =   0   'False
            Width           =   2295
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         ForeColor       =   &H00000080&
         Height          =   1305
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   1695
         Begin VB.OptionButton Opt_Juridica 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pessoa &Jurídica"
            Height          =   225
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1500
         End
         Begin VB.OptionButton Opt_Fisica 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pessoa &Física"
            Height          =   225
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   1380
         End
      End
      Begin VB.TextBox txt_NomCli 
         Height          =   330
         Left            =   1860
         MaxLength       =   50
         TabIndex        =   24
         Top             =   1040
         Width           =   6945
      End
      Begin MSMask.MaskEdBox msk_CNPJCli 
         Height          =   330
         Left            =   1860
         TabIndex        =   23
         Top             =   360
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         HideSelection   =   0   'False
         MaxLength       =   18
         Mask            =   "##.###.###/####-##"
         PromptChar      =   " "
      End
      Begin VB.Label Label22 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   1860
         TabIndex        =   29
         Top             =   795
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Identificação"
         Height          =   225
         Left            =   1860
         TabIndex        =   28
         Top             =   165
         Width           =   1800
      End
   End
   Begin TabDlg.SSTab Sst_Cliente 
      Height          =   5535
      Left            =   120
      TabIndex        =   30
      Top             =   1680
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   12582912
      TabCaption(0)   =   "Correspondência"
      TabPicture(0)   =   "Frm_incCliente.frx":0008
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label35"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label30"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label32"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Msk_CepCor"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Msk_FaxCor"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "MsK_TelCor"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cbo_UfeCor"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Txt_EndCor"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Txt_BaiCor"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Txt_CidCor"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Txt_PaiCor"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "MsK_TelCorRamal"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Txt_HomePage"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Msk_FaxCorRamal"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Frame3"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Geral"
      TabPicture(1)   =   "Frm_incCliente.frx":0024
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Comercial"
      TabPicture(2)   =   "Frm_incCliente.frx":0040
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Chk_Linha"
      Tab(2).Control(1)=   "Grafico"
      Tab(2).Control(2)=   "Grd_Familia"
      Tab(2).Control(3)=   "Grd_Vendas"
      Tab(2).Control(4)=   "Label33"
      Tab(2).Control(5)=   "Label31"
      Tab(2).ControlCount=   6
      Begin VB.Frame Frame3 
         Caption         =   "Como localizou este cliente"
         Height          =   1935
         Left            =   3600
         TabIndex        =   98
         Top             =   3480
         Width           =   3735
         Begin VB.TextBox Nom_Veiculo 
            Height          =   735
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   3375
         End
         Begin VB.ComboBox Cbo_Veiculo 
            Height          =   315
            ItemData        =   "Frm_incCliente.frx":005C
            Left            =   120
            List            =   "Frm_incCliente.frx":0075
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label38 
            Caption         =   "Nome do veículo."
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   720
            Width           =   2655
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   61
         Top             =   480
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   8705
         _Version        =   393216
         TabHeight       =   520
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Contatos"
         TabPicture(0)   =   "Frm_incCliente.frx":00BA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label29"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label28"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label27"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label26"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label25"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label23"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label16"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label12"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "TreeView1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Cbo_CodCargo"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Txt_FaxRamal"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Txt_Email"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Txt_Celular"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Txt_TelefoneRamal"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Txt_Telefone"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Cbo_NomCargo"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Txt_Fax"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Txt_NomeContato"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).ControlCount=   18
         TabCaption(1)   =   "Ficha de Dados"
         TabPicture(1)   =   "Frm_incCliente.frx":00D6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TreeView2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Agenda"
         TabPicture(2)   =   "Frm_incCliente.frx":00F2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Cbo_Quem"
         Tab(2).Control(1)=   "Cbo_QuemCodigo"
         Tab(2).Control(2)=   "Txt_Conclusao"
         Tab(2).Control(3)=   "Dta_Realizado"
         Tab(2).Control(4)=   "Txt_Oque"
         Tab(2).Control(5)=   "Grd_Agenda"
         Tab(2).Control(6)=   "Dta_Quando"
         Tab(2).Control(7)=   "Label37"
         Tab(2).Control(8)=   "Label36"
         Tab(2).Control(9)=   "Label34"
         Tab(2).Control(10)=   "Label24"
         Tab(2).Control(11)=   "Label21"
         Tab(2).ControlCount=   12
         Begin VB.ComboBox Cbo_Quem 
            Height          =   315
            Left            =   -74040
            TabIndex        =   82
            Top             =   1200
            Width           =   4335
         End
         Begin VB.ComboBox Cbo_QuemCodigo 
            Height          =   315
            ItemData        =   "Frm_incCliente.frx":010E
            Left            =   -74040
            List            =   "Frm_incCliente.frx":0110
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   1200
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.TextBox Txt_Conclusao 
            Height          =   1335
            Left            =   -69600
            MultiLine       =   -1  'True
            TabIndex        =   85
            Top             =   600
            Width           =   3615
         End
         Begin VB.PictureBox Dta_Realizado 
            Height          =   315
            Left            =   -71520
            ScaleHeight     =   255
            ScaleWidth      =   1755
            TabIndex        =   84
            Top             =   1600
            Width           =   1815
         End
         Begin VB.TextBox Txt_Oque 
            Height          =   645
            Left            =   -74040
            MultiLine       =   -1  'True
            TabIndex        =   81
            Top             =   480
            Width           =   4335
         End
         Begin VB.PictureBox Grd_Agenda 
            BackColor       =   &H00FFFFFF&
            Height          =   2775
            Left            =   -74880
            ScaleHeight     =   2715
            ScaleWidth      =   8835
            TabIndex        =   76
            Top             =   2040
            Width           =   8895
         End
         Begin MSComctlLib.TreeView TreeView2 
            Height          =   4335
            Left            =   -74880
            TabIndex        =   75
            Top             =   480
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   7646
            _Version        =   393217
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
         End
         Begin VB.TextBox Txt_NomeContato 
            Height          =   330
            Left            =   3480
            TabIndex        =   14
            Top             =   720
            Width           =   3495
         End
         Begin VB.TextBox Txt_Fax 
            Height          =   330
            Left            =   3480
            TabIndex        =   18
            Top             =   1920
            Width           =   1575
         End
         Begin VB.ComboBox Cbo_NomCargo 
            Height          =   315
            ItemData        =   "Frm_incCliente.frx":0112
            Left            =   3480
            List            =   "Frm_incCliente.frx":0114
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   3240
            Width           =   3495
         End
         Begin VB.TextBox Txt_Telefone 
            Height          =   330
            Left            =   3480
            TabIndex        =   15
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox Txt_TelefoneRamal 
            Height          =   330
            Left            =   5040
            TabIndex        =   16
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox Txt_Celular 
            Height          =   330
            Left            =   5760
            TabIndex        =   17
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox Txt_Email 
            Height          =   330
            Left            =   3480
            TabIndex        =   20
            Top             =   2520
            Width           =   3495
         End
         Begin VB.TextBox Txt_FaxRamal 
            Height          =   330
            Left            =   5040
            TabIndex        =   19
            Top             =   1920
            Width           =   615
         End
         Begin VB.ComboBox Cbo_CodCargo 
            Height          =   315
            ItemData        =   "Frm_incCliente.frx":0116
            Left            =   3480
            List            =   "Frm_incCliente.frx":0118
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   3240
            Visible         =   0   'False
            Width           =   2175
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   4335
            Left            =   120
            TabIndex        =   74
            Top             =   480
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   7646
            _Version        =   393217
            Style           =   7
            Appearance      =   1
         End
         Begin VB.PictureBox Dta_Quando 
            Height          =   315
            Left            =   -74040
            ScaleHeight     =   255
            ScaleWidth      =   1515
            TabIndex        =   83
            Top             =   1600
            Width           =   1575
         End
         Begin VB.Label Label37 
            Caption         =   "Relatório:"
            Height          =   255
            Left            =   -69600
            TabIndex        =   89
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label36 
            Caption         =   "Realizado"
            Height          =   255
            Left            =   -72360
            TabIndex        =   88
            Top             =   1650
            Width           =   1935
         End
         Begin VB.Label Label34 
            Caption         =   "Quando ?"
            Height          =   255
            Left            =   -74880
            TabIndex        =   87
            Top             =   1650
            Width           =   735
         End
         Begin VB.Label Label24 
            Caption         =   "Quem ?"
            Height          =   255
            Left            =   -74880
            TabIndex        =   86
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "O que ?"
            Height          =   375
            Left            =   -74880
            TabIndex        =   80
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label12 
            Caption         =   "Nome"
            Height          =   255
            Left            =   3480
            TabIndex        =   70
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label16 
            Caption         =   "E-Mail"
            Height          =   255
            Left            =   3480
            TabIndex        =   69
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label Label23 
            Caption         =   "Cargo ou Função"
            Height          =   255
            Left            =   3480
            TabIndex        =   68
            Top             =   3000
            Width           =   2055
         End
         Begin VB.Label Label25 
            Caption         =   "Telefone"
            Height          =   255
            Left            =   3480
            TabIndex        =   67
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "Ramal"
            Height          =   255
            Left            =   5040
            TabIndex        =   66
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label27 
            Caption         =   "Celular"
            Height          =   255
            Left            =   5760
            TabIndex        =   65
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label28 
            Caption         =   "Fax"
            Height          =   255
            Left            =   3480
            TabIndex        =   64
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label29 
            Caption         =   "Ramal"
            Height          =   255
            Left            =   5040
            TabIndex        =   63
            Top             =   1680
            Width           =   975
         End
      End
      Begin VB.CheckBox Chk_Linha 
         BackColor       =   &H80000014&
         Caption         =   "Gráfico Linha"
         Height          =   255
         Left            =   -72360
         TabIndex        =   60
         Top             =   4560
         Width           =   1335
      End
      Begin MSChart20Lib.MSChart Grafico 
         Height          =   4575
         Left            =   -72480
         OleObjectBlob   =   "Frm_incCliente.frx":011A
         TabIndex        =   59
         Top             =   360
         Width           =   6255
      End
      Begin VB.TextBox Msk_FaxCorRamal 
         Height          =   330
         Left            =   2280
         TabIndex        =   10
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox Txt_HomePage 
         Height          =   330
         Left            =   3600
         TabIndex        =   11
         Top             =   3000
         Width           =   3735
      End
      Begin VB.TextBox MsK_TelCorRamal 
         Height          =   330
         Left            =   2280
         TabIndex        =   8
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Txt_PaiCor 
         Height          =   330
         Left            =   3600
         TabIndex        =   6
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox Txt_CidCor 
         Height          =   330
         Left            =   3600
         TabIndex        =   3
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox Txt_BaiCor 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox Txt_EndCor 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   7215
      End
      Begin VB.ComboBox cbo_UfeCor 
         Height          =   315
         ItemData        =   "Frm_incCliente.frx":247F
         Left            =   120
         List            =   "Frm_incCliente.frx":24D4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1920
         Width           =   645
      End
      Begin MSMask.MaskEdBox MsK_TelCor 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Msk_FaxCor 
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Msk_CepCor 
         Height          =   330
         Left            =   840
         TabIndex        =   5
         Top             =   1920
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   10
         Mask            =   "##.###-###"
         PromptChar      =   " "
      End
      Begin VB.PictureBox Grd_Familia 
         BackColor       =   &H00FFFFFF&
         Height          =   1560
         Left            =   -74880
         ScaleHeight     =   1500
         ScaleWidth      =   2220
         TabIndex        =   51
         Top             =   600
         Width           =   2280
      End
      Begin VB.PictureBox Grd_Vendas 
         BackColor       =   &H00FFFFFF&
         Height          =   2400
         Left            =   -74880
         ScaleHeight     =   2340
         ScaleWidth      =   2220
         TabIndex        =   53
         Top             =   2520
         Width           =   2280
      End
      Begin VB.Label Label33 
         Caption         =   "Volume compra em KG?"
         Height          =   255
         Left            =   -74880
         TabIndex        =   54
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label31 
         Caption         =   "O que compra desde Jan/2001 ?"
         Height          =   255
         Left            =   -74880
         TabIndex        =   52
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Ramal"
         Height          =   255
         Left            =   2280
         TabIndex        =   50
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label32 
         Caption         =   "Home page"
         Height          =   255
         Left            =   3600
         TabIndex        =   49
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label30 
         Caption         =   "Ramal"
         Height          =   255
         Left            =   2280
         TabIndex        =   48
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Rua/Av/Número"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label35 
         Caption         =   "Telefone"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Estado"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   3600
         TabIndex        =   44
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Bairro"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Pais"
         Height          =   255
         Left            =   3600
         TabIndex        =   42
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "FAX"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000016&
         Caption         =   "Ramal"
         Height          =   255
         Left            =   -73320
         TabIndex        =   40
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000016&
         Caption         =   "FAX"
         Height          =   255
         Left            =   -74400
         TabIndex        =   39
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Cep"
         Height          =   255
         Left            =   -73680
         TabIndex        =   38
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000016&
         Caption         =   "Endereço"
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000016&
         Caption         =   "Bairro"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000016&
         Caption         =   "Cidade"
         Height          =   255
         Left            =   -72480
         TabIndex        =   35
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000016&
         Caption         =   "Estado"
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         Caption         =   "Ramal"
         Height          =   255
         Left            =   -73320
         TabIndex        =   33
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         Caption         =   "Telefone"
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Cep"
         Height          =   225
         Left            =   840
         TabIndex        =   31
         Top             =   1680
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2520
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_IncCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nodTreeView        As Node
Dim Il_Num_Gridrow     As Long
Dim il_item            As Integer
Dim Rst_Cliente        As rdoResultset
Dim Rst_Cargos         As rdoResultset
Dim Rst_Contatos       As rdoResultset
Dim Rst_Agenda         As rdoResultset
Dim Rst_Vendas         As rdoResultset
Dim Rst_Ficha As rdoResultset
Dim Dg_Indice          As Double
Dim Dg_valor1(100)     As Double
Dim Dg_valor2(100)     As Double
Dim Sg_Nome(100)       As String
Dim Tot_Nodes As Double
Dim Il_Check(1000) As Integer
Dim Il_Indice As Integer

Dim Dl_NumCar As Double
Dim Dl_NCol As Double
Dim il_Valor_Linha              As Integer
Dim il_Valor_loop               As Integer
 Dim il_Num_LinhaMaior           As Integer
Dim il_Num_LinhaMenor           As Integer
Dim il_Num_LinhaAtual           As Integer
Dim nl_num_gridrow              As Integer

Dim Sl_CodIns          As String
Dim ll_Contabil        As Long
Dim sl_TipCli          As String
 Dim Dg_VlrMinimo       As Double
Dim Sl_Pai             As String
Dim Sl_Filho           As String
Dim sl_Endcor          As String
Dim sl_Baicor          As String
Dim sl_Cidcor          As String
Dim sl_Paicor          As String
Dim sl_Ufecor          As String
Dim ll_Cepcor          As Long
Dim ll_DDDcor          As Long
Dim ll_TelCor          As Long
Dim ll_TelRamCor       As Long
Dim ll_FaxCor          As Long
Dim ll_FaxRamCor       As Long
Dim sl_ConCor          As String
Dim sl_EmailCor        As String
Dim sl_Cod_Delecao As String

Dim imgX As ListImage
Dim ll_DDDcob          As Long
Dim ll_TelCob          As Long
Dim ll_TelRamCob       As Long
Dim ll_FaxCob          As Long
Dim ll_FaxRamCob       As Long
Dim sl_ConCob          As String
Dim sl_EmailCob        As String
Dim sl_EndCob          As String
Dim sl_BaiCob          As String
Dim sl_CidCob          As String
Dim sl_PaiCob          As String
Dim sl_UfeCob          As String
Dim ll_CepCob          As Long

Dim ll_DDDEnt          As Long
Dim ll_TelEnt          As Long
Dim ll_TelRamEnt       As Long
Dim ll_FaxEnt          As Long
Dim ll_FaxRamEnt       As Long
Dim sl_ConEnt          As String
Dim sl_EmailEnt        As String
Dim sl_EndEnt          As String
Dim sl_BaiEnt          As String
Dim sl_CidEnt          As String
Dim sl_PaiEnt          As String
Dim sl_UfeEnt          As String
Dim ll_CepEnt          As Long
Dim Ll_Nlinha          As Long

Private Sub Bto_Agenda_Click()
    Frm_Agenda.Show
End Sub

Private Sub Bto_Atualiza_Click()
    If Trim(txt_NomCli) = "" Then
       MsgBox ("Nome do cliente não informado..")
       txt_NomCli.SetFocus
       Exit Sub
    End If

    If Trim(Txt_EndCor) = "" Then
       MsgBox ("Endereço não informado..")
       Txt_EndCor.SetFocus
       Exit Sub
    End If
    
    If Trim(Txt_BaiCor) = "" Then
       MsgBox ("Bairro não informado..")
       Txt_BaiCor.SetFocus
       Exit Sub
    End If
    
    If Trim(Txt_CidCor) = "" Then
       MsgBox ("Municipio não informado..")
       Txt_CidCor.SetFocus
       Exit Sub
    End If
    
    If Trim(Txt_PaiCor) = "" Then
       MsgBox ("Pais não informado..")
       Txt_PaiCor.SetFocus
       Exit Sub
    End If
    
    
    If Trim(MsK_TelCor) = "" Then
       MsgBox ("Telefone não informado..")
       MsK_TelCor.SetFocus
       Exit Sub
    End If
    
    
    If Trim(Msk_FaxCor) = "" Then
       MsgBox ("Fax não informado..")
       Msk_FaxCor.SetFocus
       Exit Sub
    End If
    If Trim(cbo_UfeCor) = "" Then
       MsgBox ("Estado não informado..")
       cbo_UfeCor.SetFocus
       Exit Sub
    End If
    If Trim(Nom_Veiculo) = "" And Cbo_Veiculo.ListIndex >= 0 Then
       MsgBox ("Não Informado o nome do veículo de localização do cliente.")
       Exit Sub
    End If
    If Trim(Nom_Veiculo) <> "" And Cbo_Veiculo.ListIndex = 0 Then
       MsgBox ("Tipo do veículo de localização do cliente não informado.")
       Exit Sub
    End If
    
    
    If Trim(Msk_CepCor) = "" Then
       MsgBox ("Cep. Não Informado")
       Msk_CepCor.SetFocus
       Exit Sub
    End If
    
    
    
    Cbo_CodCargo.ListIndex = Cbo_NomCargo.ListIndex
'    Cbo_CodigoEncarregado.ListIndex = Cbo_Encarregado.ListIndex
    
    Sl_Desc = "select * from Tba_Clientes where Cli_codigo = '" & Sg_CodCli & "'"
    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    If flag_Cliente = "I" Then
       Rst_Cliente.AddNew
    Else
       Rst_Cliente.Edit
    End If
    If Opt_Fisica.Value = True Then
       Rst_Cliente!cli_tipo = 1
    Else
       Rst_Cliente!cli_tipo = 2
    End If
    Rst_Cliente!Cli_Codigo = Sg_CodCli
    
    Rst_Cliente!Cli_Nome = txt_NomCli
    Rst_Cliente!Cli_Endereco = Txt_EndCor
    Rst_Cliente!Cli_Bairro = Txt_BaiCor
    Rst_Cliente!Cli_Municipio = Txt_CidCor
    Rst_Cliente!Cli_UF = cbo_UfeCor
    Rst_Cliente!Cli_CEP = Msk_CepCor
    
    
    If Not IsNull(Cbo_Veiculo) Then
       Rst_Cliente!Cli_Veiculo = Cbo_Veiculo
    End If
    If Not IsNull(Nom_Veiculo) Then
       Rst_Cliente!cli_nomeVeiculo = Nom_Veiculo
    End If
    
    
    
    If Txt_PaiCor <> "" Then
       Rst_Cliente!Cli_Pais = Txt_PaiCor
    End If
    Rst_Cliente!Cli_Telcor = MsK_TelCor
    If Trim(MsK_TelCorRamal) = "" Then
       Rst_Cliente!Cli_TelcorRamal = Null
    Else
       Rst_Cliente!Cli_TelcorRamal = MsK_TelCorRamal
    End If
    Rst_Cliente!Cli_Fax = Msk_FaxCor
    Rst_Cliente!Cli_FaxRamal = Msk_FaxCorRamal
    Rst_Cliente!Cli_homepage = Txt_HomePage
    
    Cbo_CodigoStatus.ListIndex = Cbo_NomeStatus.ListIndex
    Rst_Cliente!Cli_CodigoStatus = Cbo_CodigoStatus.Text
    
    
    Rst_Cliente.Update
    Rst_Cliente.Close
End Sub

Private Sub Bto_AtualizaContato_Click()
    Sl_Desc = "select * "
    Sl_Desc = Sl_Desc & " from Tba_Contatos "
    Sl_Desc = Sl_Desc & " where Con_Cliente = '" & Sg_CodCli & "'"
    Sl_Desc = Sl_Desc & "   and Con_Codigo  =  " & Val(Mid(TreeView1.SelectedItem, 1, 3))
    Set Rst_Contatos = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    Rst_Contatos.Edit
 
    Rst_Contatos!Con_Nome = Txt_NomeContato
    Rst_Contatos!Con_Telefone = Txt_Telefone
    Rst_Contatos!Con_TelefoneRamal = Txt_TelefoneRamal
    Rst_Contatos!Con_Fax = Txt_Fax
    Rst_Contatos!Con_FaxRamal = Txt_FaxRamal
    Rst_Contatos!Con_Celular = Txt_Celular
    Rst_Contatos!Con_Email = Txt_Email
    Cbo_CodCargo.ListIndex = Cbo_NomCargo.ListIndex
    Rst_Contatos!Con_Cargo = Cbo_CodCargo.Text
    Rst_Contatos.Update
    
    
    Busca_Contatos
    
    Txt_NomeContato = ""
    Txt_Telefone = ""
    Txt_TelefoneRamal = ""
    Txt_Fax = ""
    Txt_FaxRamal = ""
    Txt_Celular = ""
    Txt_Email = ""
    Cbo_CodCargo.ListIndex = -1
    Cbo_CodCargo.ListIndex = -1
    Cbo_NomCargo.ListIndex = -1
 
End Sub

Private Sub Bto_AtualizaFicha_Click()
    Sl_Desc = "select * from Tba_FichaCliente where Cli_codigo = '" & Sg_CodCli & "'"
    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    While Rst_Cliente.EOF = False
          Rst_Cliente.Edit
          Rst_Cliente.Delete
          Rst_Cliente.MoveNext
    Wend
    Sg_FlagCliente = ""
    Rst_Cliente.Close
    Sl_Desc = "select * from Tba_FichaCliente where Cli_codigo = '" & Sg_CodCli & "'"
    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    For Tot_Nodes = 1 To TreeView2.Nodes.Count
           Rst_Cliente.AddNew
           Rst_Cliente!Cli_Codigo = Sg_CodCli
           Rst_Cliente!Fix_CodigoFilho = Mid(TreeView2.Nodes.Item(Tot_Nodes).Key, 2, Len(TreeView2.Nodes.Item(Tot_Nodes).Key) - 2)
           Rst_Cliente!Fix_CodigoPai = 0
           If TreeView2.Nodes.Item(Tot_Nodes).Checked = True Then
              Rst_Cliente!Fix_Selecionado = "S"
              Sg_FlagCliente = "*"
           End If
           Rst_Cliente!CodUsuSis = Sg_Usuario
           Rst_Cliente!Fix_Data = Date + Time
           Rst_Cliente.Update
 
   Next Tot_Nodes
   Rst_Cliente.Close
   Sl_Desc = "select * from Tba_Clientes where Cli_codigo = '" & Sg_CodCli & "'"
   Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
   If Rst_Cliente.EOF = False Then
         Rst_Cliente.Edit
         If Sg_FlagCliente = "*" Then
            Rst_Cliente!Cli_Status = "R"
         Else
            Rst_Cliente!Cli_Status = Null
         End If
         Rst_Cliente.Update
   End If
   Rst_Cliente.Close
   Busca_Ficha


End Sub

Private Sub Bto_Check_Click()
    Frm_IncCliente.Enabled = False
    Sg_Flag = "I"
    Frm_IncCheck.Show
End Sub

Private Sub Bto_CheckAlt_Click()
    Frm_IncCliente.Enabled = False
    Sg_Flag = "A"
    Frm_IncCheck.Show
End Sub

Private Sub Bto_Elimina_Click()
    Dim erro As Integer
    
 
    If Trim(Grd_Agenda.Text) = "" Then
        MsgBox "Pelo menos um documento deve ser selecionada para exclusão.", vbOK + vbExclamation, "A T E N Ç Ã O"
        Exit Sub
    End If
    flag_Cliente = "E"
    sl_Num_LinhasSel = ""
    sl_Cod_Delecao = ""
    
    If Grd_Agenda.SelStartRow = 0 Then
       Grd_Agenda.SelStartRow = 1
    End If
    
    il_Valor_Linha = Grd_Agenda.SelStartRow
    
    If Grd_Agenda.SelStartRow <> Grd_Agenda.SelEndRow Then
        sl_Num_LinhasSel = "*"
    End If

     
    For il_Valor_loop = Grd_Agenda.SelStartRow To Grd_Agenda.SelEndRow
        
        Grd_Agenda.Row = il_Valor_Linha

        Grd_Agenda.Col = 0
        Sl_Desc_Mensagem1 = "da Tarefa '" & Trim(Grd_Agenda.Text) & "' ?"
        Sl_Desc_Mensagem2 = "de todas as tarefas selecionadas ?"
        sl_Texto_Atributo = Grd_Agenda.Text
        Call ConfirmarExclusao(Sl_Desc_Mensagem1, Sl_Desc_Mensagem2, sl_Texto_Atributo, sl_Cod_Delecao, sl_Num_LinhasSel, il_Valor_Linha, sl_Cod_Retorno)
         
        Grd_Agenda.Col = 0
        
        If sl_Cod_Retorno = "S" Then
          Sl_Desc = "select * from Tba_Agenda where Cli_codigo = '" & Sg_CodCli & "'"
          Sl_Desc = Sl_Desc & " and Age_Codigo = " & Dg_Age_Codigo
          Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
          While Rst_Cliente.EOF = False
             Rst_Cliente.Edit
             Rst_Cliente.Delete
             Rst_Cliente.MoveNext
          Wend
          Rst_Cliente.Close
        End If
'        CN.CommitTrans
    Next
    Busca_Agenda
End Sub

Private Sub Bto_ExcluiNodes_Click()
 
    Dim erro As Integer
    
 
'    If Trim(Grd_Cliente.Text) = "" Then
'        MsgBox "Pelo menos um documento deve ser selecionada para exclusão.", vbOK + vbExclamation, "A T E N Ç Ã O"
'        Exit Sub
'    End If
 
    sl_Num_LinhasSel = ""
    sl_Cod_Delecao = ""
    
 
    Sl_Desc_Mensagem1 = "da questao  '" & Sg_Descricao & "' ?"
    Sl_Desc_Mensagem2 = "  ?"
  
    Call ConfirmarExclusao(Sl_Desc_Mensagem1, Sl_Desc_Mensagem2, sl_Texto_Atributo, sl_Cod_Delecao, sl_Num_LinhasSel, il_Valor_Linha, sl_Cod_Retorno)
    
    If sl_Cod_Retorno = "S" Then
       Sl_Desc = " SELECT * FROM Tba_Ficha "
       Sl_Desc = Sl_Desc & " where Fix_CodigoFilho = '" & Mid(Sg_CodigoFilho, 2, Len(Sg_CodigoFilho) - 2) & "'"
       Set Rst_Ficha = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
       If Rst_Ficha.EOF = False Then
          Rst_Ficha.Edit
          Rst_Ficha.Delete
       End If
       Rst_Ficha.Close
       Busca_Ficha
   End If


''
''
''           If Grd_Cliente.Rows = 1 Then
''              Grd_Cliente.Row = 0
''              Grd_Cliente.Col = 0
''              Grd_Cliente.Text = ""
''              Grd_Cliente.Col = 1
''              Grd_Cliente.Text = ""
''           Else
''              If Grd_Cliente.SelStartRow = Grd_Cliente.SelEndRow Then
''                 Grd_Cliente.Col = 0
''                 Grd_Cliente.Text = ""
''                 Grd_Cliente.Col = 1
''                 Grd_Cliente.Text = ""
''              Else
''                 Grd_Cliente.RemoveItem Grd_Cliente.Row
''              End If
''           End If
''        End If
'''        CN.CommitTrans
''    Next
''
''   'Atualização do Grid
''    Sl_Desc = "SELECT * "
''    Sl_Desc = Sl_Desc & " FROM Tba_ClientesStatus b,"
''    Sl_Desc = Sl_Desc & "      Tba_CLIENTES a"
''    Sl_Desc = Sl_Desc & " where b.Cli_CodigoStatus  = a.Cli_CodigoStatus"
''    Sl_Desc = Sl_Desc & " ORDER BY CLI_NOME "
''    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenKeyset, rdConcurRowVer)
''    Il_Num_Gridrow = 1
''    Do While Rst_Cliente.EOF = False
''       Grd_Cliente.Rows = Il_Num_Gridrow + 1
''       Grd_Cliente.Row = Il_Num_Gridrow
''       Grd_Cliente.Col = 0
''       Grd_Cliente.Text = Rst_Cliente("Cli_Codigo")
''       Grd_Cliente.Col = 1
''       Grd_Cliente.Text = Rst_Cliente("Cli_Nome")
''       Grd_Cliente.Col = 2
''       Grd_Cliente.Text = Rst_Cliente("Cli_NomeStatus")
''
''       Rst_Cliente.MoveNext
''       Il_Num_Gridrow = Il_Num_Gridrow + 1
''    Loop
''    Grd_Cliente.Row = Grd_Cliente.SelStartRow
''    Rst_Cliente.Close
''  Exit Sub
''
''Trata_Erro:
''  Rotina_erro ("A")
''  Beep
End Sub
 

Private Sub Bto_Excluir_Click()
     Dim erro As Integer
 
     If Trim(TreeView1.SelectedItem) = "Contatos" Then
         MsgBox "Pelo menos um contato deve ser selecionado para exclusão.", vbOK + vbExclamation, "A T E N Ç Ã O"
         Exit Sub
     End If
     
     Sl_Desc = "select * "
     Sl_Desc = Sl_Desc & " from Tba_Contatos "
     Sl_Desc = Sl_Desc & " where Con_Cliente = '" & Sg_CodCli & "'"
     Sl_Desc = Sl_Desc & "   and Con_CodigoEncarregado  =  " & Val(Mid(TreeView1.SelectedItem, 1, 3))
     Set Rst_Contatos = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
     If Rst_Contatos.EOF = False Then
        MsgBox ("Não é possível eliminar um encarregado sem eliminar primeiro os subordinados..")
        Rst_Contatos.Close
        Exit Sub
     End If
     Rst_Contatos.Close
     
     
     
     sl_Cod_Delecao = ""

 
     Sl_Desc_Mensagem1 = "do Contato '" & Trim(Txt_NomeContato) & "' ?"
     Sl_Desc_Mensagem2 = ""
     Call ConfirmarExclusao(Sl_Desc_Mensagem1, Sl_Desc_Mensagem2, sl_Texto_Atributo, sl_Cod_Delecao, sl_Num_LinhasSel, il_Valor_Linha, sl_Cod_Retorno)
    
    
     If sl_Cod_Retorno = "S" Then
        Sl_Desc = "select * "
        Sl_Desc = Sl_Desc & " from Tba_Contatos "
        Sl_Desc = Sl_Desc & " where Con_Cliente = '" & Sg_CodCli & "'"
        Sl_Desc = Sl_Desc & "   and Con_Codigo  =  " & Val(Mid(TreeView1.SelectedItem, 1, 3))
        Set Rst_Contatos = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
           Rst_Contatos.Edit
           Rst_Contatos.Delete
           Rst_Contatos.Close
     End If
     Busca_Contatos
     Bto_Excluir.Enabled = False
     Exit Sub
    
Trata_Erro:
  Rotina_erro ("A")
  Beep

End Sub

Private Sub Bto_Gravar_Click()
    If Trim(Txt_Oque) = "" Then
       MsgBox ("Não informado O que deve ser feito.")
       Txt_Oque.SetFocus
       Exit Sub
    End If
    If Trim(Dta_Quando) = "" Then
       MsgBox ("Quando deve ser feito não informado.")
       Dta_Quando.SetFocus
       Exit Sub
    End If
    If Trim(Dta_Realizado) <> "" Then
       If Trim(Cbo_QuemCodigo.Text) = "" Then
          MsgBox ("Quem realizou a tarefa não informado.")
          Exit Sub
       End If
    End If
    If Trim(Dta_Realizado) <> "" Then
       If Trim(Txt_Conclusao) = "" Then
          MsgBox ("Histórico não informado.")
          Exit Sub
       End If
    End If
    
    
    If Trim(Dta_Realizado) <> "" Then
       If CDate(Dta_Realizado) < CDate(Dta_Quando) Then
          MsgBox ("Data da realização não pode ser menor que a data prevista para realização.")
          Dta_Realizado = ""
          Exit Sub
       End If
    End If
    Grd_Agenda.Col = 5
    Dg_Age_Codigo = Grd_Agenda.Text
    Cbo_QuemCodigo.ListIndex = Cbo_Quem.ListIndex
    Sl_Desc = "select * from Tba_Agenda where Cli_codigo = '" & Sg_CodCli & "'"
    Sl_Desc = Sl_Desc & " and Age_Codigo = " & Dg_Age_Codigo
    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    Rst_Cliente.Edit
 
        Rst_Cliente!Age_Oque = Txt_Oque
        If Trim(Cbo_QuemCodigo.Text) = "" Then
           Rst_Cliente!CodUsuSis = 0
        Else
           Rst_Cliente!CodUsuSis = Cbo_QuemCodigo.Text
           
        End If
          
        If Trim(Dta_Quando) <> "" Then
           Rst_Cliente!Age_Quando = CDate(Dta_Quando)
        Else
           Rst_Cliente!Age_Quando = Null
        End If
        
        
        If Trim(Dta_Realizado) <> "" Then
           Rst_Cliente!Age_Realizado = CDate(Dta_Realizado)
        Else
           Rst_Cliente!Age_Realizado = Null
        End If
        If Trim(Txt_Conclusao) <> "" Then
           Rst_Cliente!Age_Relatorio = Txt_Conclusao
        Else
           Rst_Cliente!Age_Relatorio = Null
        End If
    Rst_Cliente.Update
    Rst_Cliente.Close
    Busca_Agenda
End Sub

Private Sub Bto_GravarAgenda_Click()
    If Trim(Txt_Oque) = "" Then
       MsgBox ("Não informado O que deve ser feito.")
       Txt_Oque.SetFocus
       Exit Sub
    End If
    If Trim(Dta_Quando) = "" Then
       MsgBox ("Quando deve ser feito não informado.")
       Dta_Quando.SetFocus
       Exit Sub
    End If
    Cbo_QuemCodigo.ListIndex = Cbo_Quem.ListIndex
    Sl_Desc = "select * from Tba_Agenda where Cli_codigo = '" & Sg_CodCli & "'"
    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    Rst_Cliente.AddNew
        Rst_Cliente!Cli_Codigo = Sg_CodCli
        Rst_Cliente!Age_Oque = Txt_Oque
        If Cbo_QuemCodigo.ListIndex > 0 Then
           Rst_Cliente!CodUsuSis = Cbo_QuemCodigo.Text
        End If
        Rst_Cliente!Age_Quando = Dta_Quando
        If Trim(Dta_Realizado) <> "" Then
           Rst_Cliente!Age_Realizado = Dta_Realizado
        End If
        If Trim(Txt_Conclusao) <> "" Then
           Rst_Cliente!Age_Relatorio = Txt_Conclusao
        End If
    Rst_Cliente.Update
    Rst_Cliente.Close
    Busca_Agenda

End Sub

Private Sub bto_Incluir_Click()
    If Trim(Txt_NomeContato) = "" Then
       MsgBox ("Não informado os dados do contato.")
       Exit Sub
    End If
    
    Cbo_CodCargo.ListIndex = Cbo_NomCargo.ListIndex
'    Cbo_CodigoEncarregado.ListIndex = Cbo_Encarregado.ListIndex



    
    Sl_Desc = "select max(Con_Codigo) from Tba_Contatos where Con_Cliente = '" & Sg_CodCli & "'"
    Set Rst_Contatos = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    If IsNull(Rst_Contatos(0)) Then
       Dg_Codigo = 1
    Else
       Dg_Codigo = Rst_Contatos(0) + 1
    End If
    Rst_Contatos.Close
    
    Sl_Desc = "select * from Tba_Contatos where Con_Cliente = '" & Sg_CodCli & "'"
    Set Rst_Contatos = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    Rst_Contatos.AddNew
    Rst_Contatos!Con_Cliente = Sg_CodCli
    Rst_Contatos!Con_Codigo = Dg_Codigo
    Rst_Contatos!Con_Nome = Txt_NomeContato
    Rst_Contatos!Con_Telefone = Txt_Telefone
    Rst_Contatos!Con_TelefoneRamal = Txt_TelefoneRamal
    Rst_Contatos!Con_Fax = Txt_Fax
    Rst_Contatos!Con_FaxRamal = Txt_FaxRamal
    Rst_Contatos!Con_Celular = Txt_Celular
    If Not IsNull(Txt_Email) Then
       Rst_Contatos!Con_Email = Txt_Email
    End If
    Rst_Contatos!Con_Cargo = Cbo_CodCargo.Text
 
      
    
    Rst_Contatos.Update
    Rst_Contatos.Close
    Busca_Contatos

End Sub


Private Sub Bto_Limpa_Click()
    Txt_Conclusao = ""
    Txt_Oque = ""
    Cbo_Quem.ListIndex = 0
    Cbo_QuemCodigo.ListIndex = 0
    Dta_Quando = ""
    Dta_Realizado = ""
    Bto_GravarAgenda.Enabled = True
    Bto_Elimina.Enabled = False
    Bto_Gravar.Enabled = False
End Sub

Private Sub Bto_LimpaEndereco_Click()
    Txt_EndCor = ""
    Txt_BaiCor = ""
    Txt_CidCor = ""
    Txt_PaiCor = ""
    MsK_TelCor = ""
    MsK_TelCorRamal = ""
    Msk_FaxCor = ""
    Msk_FaxCorRamal = ""
    Txt_PaiCor = ""
    Txt_HomePage = ""
End Sub

Private Sub Bto_LimpaFicha_Click()
    Bto_Incluir.Enabled = True
    Bto_Excluir.Enabled = False
    Bto_AtualizaContato.Enabled = False
     
    Txt_NomeContato = ""
    Txt_Telefone = ""
    Txt_TelefoneRamal = ""
    Txt_Celular = ""
    Txt_Fax = ""
    Txt_FaxRamal = ""
    Txt_Email = ""
    Cbo_NomCargo.ListIndex = -1
End Sub

Private Sub Bto_RelFicha_Click()
    If Trim(txt_NomCli.Text) = "" Then
       MsgBox ("Nenhum contato selelcionado para emissão da ficha.")
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    Sl_Desc = " SELECT"
    Sl_Desc = Sl_Desc & " Tba_Ficha.`Fix_CodigoFilho`, Tba_Ficha.`Fix_CodigoPai`, Tba_Ficha.`Fix_Descricao`,"
    Sl_Desc = Sl_Desc & "     Tba_FichaCliente.`Fix_Selecionado`,"
    Sl_Desc = Sl_Desc & "     Tba_Clientes.`Cli_Codigo`, Tba_Clientes.`Cli_Nome`, Tba_Clientes.`Cli_Endereco`, Tba_Clientes.`Cli_Bairro`, Tba_Clientes.`Cli_Municipio`, Tba_Clientes.`Cli_UF`, Tba_Clientes.`Cli_CEP`, Tba_Clientes.`Cli_Telcor`, Tba_Clientes.`Cli_TelcorRamal`, Tba_Clientes.`Cli_Fax`, Tba_Clientes.`Cli_FaxRamal`, Tba_Clientes.`Cli_homepage`"
    Sl_Desc = Sl_Desc & " From"
    Sl_Desc = Sl_Desc & "     (`Tba_Ficha` Tba_Ficha INNER JOIN `Tba_FichaCliente` Tba_FichaCliente ON"
    Sl_Desc = Sl_Desc & "         Tba_Ficha.`Fix_CodigoFilho` = Tba_FichaCliente.`Fix_CodigoFilho`)"
    Sl_Desc = Sl_Desc & "      INNER JOIN `Tba_Clientes` Tba_Clientes ON"
    Sl_Desc = Sl_Desc & "         Tba_FichaCliente.`Cli_Codigo` = Tba_Clientes.`Cli_Codigo`"
    Sl_Desc = Sl_Desc & "  where Tba_FichaCliente.`Cli_Codigo` = '" & Sg_CodCli & "'"
     
    
    Sl_Desc = Sl_Desc & " Order By"
    Sl_Desc = Sl_Desc & "     Tba_Ficha.`Fix_CodigoFilho` ASC,"
    Sl_Desc = Sl_Desc & "     Tba_Ficha.`Fix_CodigoPai` ASC "
    
    CrystalReport1.ReportFileName = App.Path & "\Relatorios\RelFicha.rpt"
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.SQLQuery = Sl_Desc
    CrystalReport1.PrintReport
    Screen.MousePointer = vbDefault
  
    
    
   
    
End Sub

Private Sub Bto_Sair_Click()
    Unload Me
End Sub
 
Private Sub Command1_Click()
Busca_Vendas
End Sub

Private Sub Cbo_Quem_Click()
    Cbo_QuemCodigo.ListIndex = Cbo_Quem.ListIndex
End Sub

Private Sub Chk_Linha_Click()
  If Chk_Linha.Value = 1 Then
     Grafico.chartType = 3
  Else
     Grafico.chartType = 1
  End If
     
  For Dl_NCol = 1 To Dg_Indice
      Grafico.Row = Dl_NCol
      Grafico.RowLabel = Sg_Nome(Dl_NCol)
        
      Grafico.Column = 1
      Grafico.Data = Format(Dg_valor1(Dl_NCol), "###,##0.00")
  Next Dl_NCol
 
  Exit Sub

  
End Sub

Private Sub Form_Activate()
    Bto_Check.Enabled = False
    Bto_CheckAlt.Enabled = False
    If Sg_Flag = "I" Then
       Set nodTreeView = TreeView2.Nodes.Add(Sg_CodigoPai, tvwChild, sg_Proximo_codigoFilho, Sg_Descricao, 0)
    End If
    If flag_Cliente = "I" Then
       Opt_Fisica.SetFocus
    End If
    
    
    If Sg_Flag = "A" Then
       Busca_Ficha
    End If
      
    Sg_Flag = ""
End Sub

Private Sub Form_Load()

    Left = 0
    Top = 0
    Height = 7880
    Width = 11850

 
    Sst_Cliente.Tab = 0

 
    Bto_AtualizaContato.Enabled = False
    Bto_Check.Enabled = False
    Bto_CheckAlt.Enabled = False

    Bto_GravarAgenda.Enabled = True
    Bto_Elimina.Enabled = False
    Bto_Gravar.Enabled = False
    
    Bto_Incluir.Enabled = False
    Bto_Excluir.Enabled = False
    Bto_AtualizaContato.Enabled = False

    If flag_Cliente = "I" Then
       msk_CNPJCli.Enabled = True
       Opt_Fisica.Enabled = True
       Opt_Juridica.Enabled = True
       Opt_Juridica.Value = True
       Sst_Cliente.TabEnabled(0) = True
       Sst_Cliente.TabEnabled(1) = True
       
 
    Else
       msk_CNPJCli.Enabled = False
       Opt_Fisica.Enabled = False
       Opt_Juridica.Enabled = False
       Sst_Cliente.TabEnabled(0) = True
       Sst_Cliente.TabEnabled(1) = True
   
   
    End If
    
    Grafico.RowCount = 10
    Grafico.ColumnCount = 1
        
    For Dl_NCol = 1 To 10
        Grafico.Row = Dl_NCol
        Grafico.RowLabel = 0
        
        Grafico.Column = 1
        Grafico.Data = 0
    Next Dl_NCol
    Grafico.Refresh
    
    Cbo_CodCargo.Clear
    Cbo_NomCargo.Clear
    Sl_Desc = "select * from Tba_Cargos"
    Set Rst_Cargos = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    Cbo_CodCargo.AddItem "999"
     Cbo_NomCargo.AddItem "NÃO ESPECIFICADO"
    
    While Rst_Cargos.EOF = False
          Cbo_CodCargo.AddItem Rst_Cargos("Car_Codigo")
          Cbo_NomCargo.AddItem Rst_Cargos("Car_Nome")
          Rst_Cargos.MoveNext
    Wend
    Rst_Cargos.Close
   
   
   
    Cbo_Quem.Clear
    Cbo_QuemCodigo.Clear
    Sl_Desc = "select * from tba_usuarios order by nomususis"
    Set Rst_Cargos = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    Cbo_Quem.AddItem " "
    Cbo_QuemCodigo.AddItem " "
    While Rst_Cargos.EOF = False
          Cbo_Quem.AddItem Rst_Cargos("NomUsuSis")
          Cbo_QuemCodigo.AddItem Rst_Cargos("CodUsuSis")
          Rst_Cargos.MoveNext
    Wend
    Rst_Cargos.Close
    
    
 
    
    Cbo_NomeStatus.Clear
    Cbo_CodigoStatus.Clear
    Sl_Desc = "select * from Tba_ClientesStatus"
    Set Rst_Contatos = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    While Rst_Contatos.EOF = False
          Cbo_CodigoStatus.AddItem Rst_Contatos!Cli_CodigoStatus
          Cbo_NomeStatus.AddItem Rst_Contatos!Cli_NomeStatus
          Rst_Contatos.MoveNext
    Wend
    Rst_Contatos.Close
   
    If flag_Cliente = "I" Then
       Frm_IncCliente.Caption = "Incluir Cliente"
       Exit Sub
    Else
       Frm_IncCliente.Caption = "Alterar Cliente"
    End If

    
    
    
     
    Sl_Desc = "select * from Tba_Clientes where Cli_codigo = '" & Sg_CodCli & "'"
    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    If Rst_Cliente.EOF = True Then
       MsgBox "Erro inesperado - CNPJCli = '" & Sg_CNPJCli & "' não encontrado em Tba_Cliente."
       Rst_Cliente.Close
       Exit Sub
    End If
    
    If Rst_Cliente("Cli_Tipo") = "1" Then
       Opt_Juridica.Value = True
    Else
       Opt_Fisica.Value = True
    End If
      
    If Opt_Juridica.Value = True Then
       msk_CNPJCli.Mask = "##.###.###/####-##"
    ElseIf Opt_Fisica.Value = True Then
       msk_CNPJCli.Mask = "###.###.###-##"
    End If
    msk_CNPJCli.Text = Sg_CodCli
    txt_NomCli.Text = Rst_Cliente!Cli_Nome
    Txt_EndCor = Rst_Cliente!Cli_Endereco
    Txt_BaiCor = Rst_Cliente!Cli_Bairro
    Txt_CidCor = Rst_Cliente!Cli_Municipio
    cbo_UfeCor = Rst_Cliente!Cli_UF
    Msk_CepCor = Rst_Cliente!Cli_CEP
    If Not IsNull(Rst_Cliente!Cli_Pais) Then
       Txt_PaiCor = Rst_Cliente!Cli_Pais
    End If
    If Not IsNull(Rst_Cliente!Cli_Telcor) Then
       MsK_TelCor = Rst_Cliente!Cli_Telcor
    End If
    If Not IsNull(Rst_Cliente!Cli_TelcorRamal) Then
       MsK_TelCorRamal = Rst_Cliente!Cli_TelcorRamal
    End If
    If Not IsNull(Rst_Cliente!Cli_Fax) Then
       Msk_FaxCor = Rst_Cliente!Cli_Fax
    End If
    Msk_FaxCorRamal = Rst_Cliente!Cli_FaxRamal
    If Not IsNull(Rst_Cliente!Cli_homepage) Then
       Txt_HomePage = Rst_Cliente!Cli_homepage
    End If
    Cbo_CodigoStatus = Rst_Cliente!Cli_CodigoStatus
    Cbo_NomeStatus.ListIndex = Cbo_CodigoStatus.ListIndex
    
    If Not IsNull(Rst_Cliente!Cli_Veiculo) Then
       If Trim(Rst_Cliente!Cli_Veiculo) = "" Then
          Cbo_Veiculo.ListIndex = -1
        Else
           Cbo_Veiculo = Rst_Cliente!Cli_Veiculo
        End If
    End If
    If Not IsNull(Rst_Cliente!cli_nomeVeiculo) Then
       Nom_Veiculo = Rst_Cliente!cli_nomeVeiculo
    End If
    
    Rst_Cliente.Close
    
    Busca_Contatos
    Busca_Ficha
    Busca_Vendas
    Busca_Agenda
    Exit Sub

Trata_Erro:

    Rotina_erro ("CAD008")
    Beep

End Sub
Private Sub Busca_Agenda()
    Grd_Agenda.Rows = 2
    Grd_Agenda.Row = 0
    
    
    Grd_Agenda.Col = 0
    Grd_Agenda.ColWidth(0) = 2000
    Grd_Agenda.Text = "O que ?"
    
    Grd_Agenda.Col = 1
    Grd_Agenda.ColWidth(1) = 1000
    Grd_Agenda.Text = "Quem ?"
    
    Grd_Agenda.Col = 2
    Grd_Agenda.ColWidth(2) = 1000
    Grd_Agenda.Text = " Quando ?"
    
    Grd_Agenda.Col = 3
    Grd_Agenda.ColWidth(3) = 1000
    Grd_Agenda.Text = " Realizado ?"
     
    Grd_Agenda.Col = 4
    Grd_Agenda.ColWidth(4) = 3500
    Grd_Agenda.Text = " Conclusão ?"
    
    Grd_Agenda.Col = 5
    Grd_Agenda.ColWidth(5) = 1
    Grd_Agenda.Text = ""
    
    Grd_Agenda.Col = 6
    Grd_Agenda.ColWidth(6) = 1
    Grd_Agenda.Text = ""

    
    Grd_Agenda.Row = 1
    Grd_Agenda.Col = 0
    Grd_Agenda.Text = " "
    Grd_Agenda.Col = 1
    Grd_Agenda.Text = " "
    Grd_Agenda.Col = 2
    Grd_Agenda.Text = " "
    Grd_Agenda.Col = 3
    Grd_Agenda.Text = " "
    Grd_Agenda.Col = 4
    Grd_Agenda.Text = " "
    Grd_Agenda.Col = 5
    Grd_Agenda.Text = " "
    Grd_Agenda.Col = 6
    Grd_Agenda.Text = " "
    
    Sl_Desc = "SELECT *"
    Sl_Desc = Sl_Desc & " FROM Tba_Usuarios RIGHT JOIN Tba_Agenda "
    Sl_Desc = Sl_Desc & " ON [Tba_Usuarios].[CodUsuSis] = [Tba_Agenda].[CodUsuSis]"
    Sl_Desc = Sl_Desc & " where Cli_Codigo = '" & Sg_CodCli & "'"
    Sl_Desc = Sl_Desc & " order by Age_Quando desc;"
    
    
    
    Set Rst_Agenda = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    Il_Num_Gridrow = 1
    Do While Rst_Agenda.EOF = False
       Grd_Agenda.Rows = Il_Num_Gridrow + 1
       Grd_Agenda.Row = Il_Num_Gridrow
       Grd_Agenda.Col = 0
       Grd_Agenda.Text = Rst_Agenda("Age_Oque")
       
       Grd_Agenda.Col = 1
       If Not IsNull(Rst_Agenda("NomUsuSis")) Then
          Grd_Agenda.Text = Rst_Agenda("NomUsuSis")
       End If
       
       Grd_Agenda.Col = 2
       If Not IsNull(Rst_Agenda("Age_Quando")) Then
          Grd_Agenda.Text = Rst_Agenda("Age_Quando")
       Else
          Grd_Agenda.Text = ""
       End If
       Grd_Agenda.Col = 3
       If Not IsNull(Rst_Agenda("Age_Realizado")) Then
          Grd_Agenda.Text = Rst_Agenda("Age_Realizado")
       End If
       Grd_Agenda.Col = 4
       If Not IsNull(Rst_Agenda("Age_Relatorio")) Then
          Grd_Agenda.Text = Rst_Agenda("Age_Relatorio")
       End If
       Grd_Agenda.Col = 5
       Grd_Agenda.Text = Rst_Agenda("Age_Codigo")
       
       Grd_Agenda.Col = 6
       If Not IsNull(Rst_Agenda("CodUsuSis")) Then
          Grd_Agenda.Text = Rst_Agenda("CodUsuSis")
       End If
       Rst_Agenda.MoveNext
       Il_Num_Gridrow = Il_Num_Gridrow + 1
    Loop
    Grd_Agenda.Row = Grd_Agenda.SelStartRow
    Rst_Agenda.Close
    Exit Sub

    
    
    
End Sub
Private Sub Busca_Vendas()
    Grd_Vendas.Rows = 2
    Grd_Vendas.Col = 0
    Grd_Vendas.Row = 0
    Grd_Vendas.ColWidth(0) = 800
    Grd_Vendas.Text = "Data"
    Grd_Vendas.Col = 1
    Grd_Vendas.ColAlignment(1) = 1
    Grd_Vendas.ColWidth(1) = 1000
    Grd_Vendas.Text = " Peso"
    
    Grd_Vendas.Row = 1
    Grd_Vendas.Col = 0
    Grd_Vendas.Text = " "
    Grd_Vendas.Col = 1
    Grd_Vendas.Text = " "
    
    
    Grd_Familia.Rows = 2
    Grd_Familia.Col = 0
    Grd_Familia.Row = 0
    Grd_Familia.ColWidth(0) = 800
    Grd_Familia.Text = "Produto"
    Grd_Familia.Col = 1
    Grd_Familia.ColAlignment(1) = 1
    Grd_Familia.ColWidth(1) = 1000
    Grd_Familia.Text = " Peso"
    
    Grd_Familia.Row = 1
    Grd_Familia.Col = 0
    Grd_Familia.Text = " "
    Grd_Familia.Col = 1
    Grd_Familia.Text = " "
    
    
    
    Sl_Desc = "select Ven_Produto, sum(Ven_Peso)   "
    Sl_Desc = Sl_Desc & " from Tba_Vendas "
    Sl_Desc = Sl_Desc & " where Ven_Cliente = '" & Sg_CodCli & "'"
    Sl_Desc = Sl_Desc & " group by Ven_Produto"
    Set Rst_Vendas = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    Il_Num_Gridrow = 1
    Do While Rst_Vendas.EOF = False
       Grd_Familia.Rows = Il_Num_Gridrow + 1
       Grd_Familia.Row = Il_Num_Gridrow
       Grd_Familia.Col = 0
       Grd_Familia.Text = Rst_Vendas("Ven_Produto")
       Grd_Familia.Col = 1
       Grd_Familia.Text = Format(Rst_Vendas(1), "###,##0.00")
       Rst_Vendas.MoveNext
       Il_Num_Gridrow = Il_Num_Gridrow + 1
    Loop
    Grd_Familia.Row = Grd_Familia.SelStartRow
    Rst_Vendas.Close
    Exit Sub

     
     
     
     
     
     
     
     
End Sub
Private Sub Busca_Ficha()
    ImageList2.ListImages.Clear
   
    
    Set imgX = ImageList2.ListImages.Add(, "to", LoadPicture(App.Path & "\imagens\Arquivo.ICO"))
    Set imgX = ImageList2.ListImages.Add(, "po", LoadPicture(App.Path & "\imagens\Ficha.ICO"))
    Set imgX = ImageList2.ListImages.Add(, "ro", LoadPicture(App.Path & "\imagens\ON.ICO"))
    Set imgX = ImageList2.ListImages.Add(, "bo", LoadPicture(App.Path & "\imagens\OFF.ICO"))
    Set imgX = ImageList2.ListImages.Add(, "jo", LoadPicture(App.Path & "\imagens\Folha1.ICO"))
    Set imgX = ImageList2.ListImages.Add(, "oo", LoadPicture(App.Path & "\imagens\Folha2.ICO"))
    

    
   ' Set TreeView2.ImageList = ImageList2
    
   TreeView2.LabelEdit = tvwManual ' Set property to manual.
   TreeView2.Nodes.Clear
   TreeView2.HotTracking = True
    ''TreeView1.Style = tvwTreelinesPlusMinusText
   Sl_Pai = "'1'"
    
   Set nodTreeView = TreeView2.Nodes.Add(, tvwNext, Sl_Pai, "Ficha de dados", 0)
    
    
    Sl_Desc = " SELECT * FROM Tba_Ficha "
    Set Rst_Contatos = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    While Rst_Contatos.EOF = False
          Sl_Pai = "'" & Rst_Contatos!Fix_CodigoPai & "'"
          Sl_Filho = "'" & Rst_Contatos!Fix_CodigoFilho & "'"
  
          Set nodTreeView = TreeView2.Nodes.Add(Sl_Pai, tvwChild, Sl_Filho, Rst_Contatos!Fix_Descricao, 0)
           
          Rst_Contatos.MoveNext
   Wend
   Rst_Contatos.Close
   TreeView2.Nodes.Item(1).Expanded = True
   
   
   Sl_Desc = "select * from Tba_FichaCliente where Cli_codigo = '" & Sg_CodCli & "'"
   Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
   Tot_Nodes = 0
   While Rst_Cliente.EOF = False
         Tot_Nodes = Tot_Nodes + 1
         If Trim(Rst_Cliente!Fix_Selecionado) = "S" Then
            TreeView2.Nodes.Item(Tot_Nodes).Checked = True
         Else
            TreeView2.Nodes.Item(Tot_Nodes).Checked = False
         End If
         Rst_Cliente.MoveNext
   Wend
   Rst_Cliente.Close
 
    


End Sub
Private Sub Busca_Contatos()
   
    ImageList1.ListImages.Clear
'    Set imgX = ImageList1.ListImages.Add(, "to", LoadPicture(App.Path & "\imagens\ON.ICO"))
'    Set imgX = ImageList1.ListImages.Add(, "ro", LoadPicture(App.Path & "\imagens\OFF.ICO"))
'    Set TreeView1.ImageList = ImageList1
    
   TreeView1.LabelEdit = tvwManual ' Set property to manual.
   TreeView1.Nodes.Clear
   TreeView1.HotTracking = True
    ''TreeView1.Style = tvwTreelinesPlusMinusText
   Sl_Pai = "'0'"
    
   Set nodTreeView = TreeView1.Nodes.Add(, , Sl_Pai, "Contatos", 0)
     
   Sl_Desc = "select * "
   Sl_Desc = Sl_Desc & " from Tba_Contatos "
   Sl_Desc = Sl_Desc & " where Con_Cliente = '" & Sg_CodCli & "'"
   Sl_Desc = Sl_Desc & " order by con_codigoEncarregado, con_codigo"
   Set Rst_Contatos = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
   While Rst_Contatos.EOF = False
         Sl_Pai = "'" & Rst_Contatos!Con_CodigoEncarregado & "'"
         If IsNull(Rst_Contatos!Con_CodigoEncarregado) Then
            Sl_Pai = "'0'"
         End If
         Sl_Filho = "'" & Rst_Contatos!Con_Codigo & "'"
   
         Set nodTreeView = TreeView1.Nodes.Add(Sl_Pai, tvwChild, Sl_Filho, Format(Rst_Contatos!Con_Codigo, "000") & "  -  " & Rst_Contatos!Con_Nome, 0)
          
         
         Rst_Contatos.MoveNext
   Wend
   Rst_Contatos.Close
   TreeView1.Nodes.Item(1).Expanded = True

    
    
    
    



End Sub

Private Sub Form_Unload(Cancel As Integer)

    Screen.MousePointer = vbHourglass
'    Cliente_objt.Enabled = True
    Screen.MousePointer = vbDefault
    
End Sub

 
 

 
Private Sub Grd_Agenda_Click()
    If Trim(Grd_Agenda.Text) = "" Then
       Exit Sub
    End If
    Grd_Agenda.Col = 0
    Txt_Oque = Grd_Agenda.Text
    Grd_Agenda.Col = 1
    If Trim(Grd_Agenda.Text) <> "" Then
       Cbo_Quem.Text = Trim(Grd_Agenda.Text)
       Grd_Agenda.Col = 6
       Cbo_QuemCodigo = Grd_Agenda.Text
       Cbo_Quem.ListIndex = Cbo_QuemCodigo.ListIndex
    Else
       Cbo_Quem.ListIndex = 0
       Cbo_QuemCodigo.ListIndex = 0
       Cbo_Quem.Refresh
       Cbo_QuemCodigo.Refresh
       
    End If
 
    Grd_Agenda.Col = 2
    Dta_Quando = Grd_Agenda.Text
    Grd_Agenda.Col = 3
    Dta_Realizado = Grd_Agenda.Text
    Grd_Agenda.Col = 4
    Txt_Conclusao = Grd_Agenda.Text
    Bto_GravarAgenda.Enabled = False
    Bto_Elimina.Enabled = True
    Bto_Gravar.Enabled = True
    Grd_Agenda.Col = 5
    Dg_Age_Codigo = Grd_Agenda.Text

End Sub

Private Sub Grd_Familia_Click()
    Grd_Vendas.Rows = 2
    Grd_Vendas.Col = 0
    Grd_Vendas.Row = 0
    Grd_Vendas.ColWidth(0) = 800
    Grd_Vendas.Text = "Data"
    Grd_Vendas.Col = 1
    Grd_Vendas.ColAlignment(1) = 1
    Grd_Vendas.ColWidth(1) = 1000
    Grd_Vendas.Text = " Peso"
    
    Grd_Vendas.Row = 1
    Grd_Vendas.Col = 0
    Grd_Vendas.Text = " "
    Grd_Vendas.Col = 1
    Grd_Vendas.Text = " "
    
    
    Sl_Desc = "select Ven_data, sum(Ven_Peso) "
    Sl_Desc = Sl_Desc & " from Tba_Vendas "
    Sl_Desc = Sl_Desc & " where Ven_Cliente = '" & Sg_CodCli & "'"
    Grd_Familia.Col = 0
    Sl_Desc = Sl_Desc & "   and Ven_produto = '" & Grd_Familia.Text & "'"
    Sl_Desc = Sl_Desc & " group by Ven_Data"
    Set Rst_Vendas = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    Il_Num_Gridrow = 1
    Do While Rst_Vendas.EOF = False
       Grd_Vendas.Rows = Il_Num_Gridrow + 1
       Grd_Vendas.Row = Il_Num_Gridrow
       Grd_Vendas.Col = 0
       Grd_Vendas.Text = Rst_Vendas("Ven_Data")
       Grd_Vendas.Col = 1
       Grd_Vendas.Text = Format(Rst_Vendas(1), "###,##0.00")
       Rst_Vendas.MoveNext
       Il_Num_Gridrow = Il_Num_Gridrow + 1
    Loop
    Grd_Vendas.Row = Grd_Vendas.SelStartRow
    Rst_Vendas.Close
    
    
    Grafico.ColumnCount = 1
    
    i = 1
    Grafico.Plot.SeriesCollection(i).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeOutside
    Grafico.Plot.SeriesCollection(i).DataPoints(-1).DataPointLabel.VtFont.Size = 6
    Grafico.Column = 1
    Grafico.ColumnLabel = "Divisor"
     
    Sl_Desc = "select mid(Ven_data,1,4), sum(Ven_Peso), count(Ven_data) "
    Sl_Desc = Sl_Desc & " from Tba_Vendas "
    Sl_Desc = Sl_Desc & " where Ven_Cliente = '" & Sg_CodCli & "'"
    Grd_Familia.Col = 0
    Sl_Desc = Sl_Desc & "   and Ven_produto = '" & Grd_Familia.Text & "'"
    Sl_Desc = Sl_Desc & " group by mid(Ven_data,1,4)"
    
    Set Rst_Vendas = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    Dg_Indice = 0
    Do While Rst_Vendas.EOF = False
       Dg_Indice = Dg_Indice + 1
       Sg_Nome(Dg_Indice) = Rst_Vendas(0)
       Dg_valor1(Dg_Indice) = Rst_Vendas(1) / Rst_Vendas(2)
       Rst_Vendas.MoveNext
    Loop
    Rst_Vendas.Close
     
    Grafico.RowCount = Dg_Indice
    Grafico.ColumnCount = 1
        
    For Dl_NCol = 1 To Dg_Indice
        Grafico.Row = Dl_NCol
        Grafico.RowLabel = Sg_Nome(Dl_NCol)
        
        Grafico.Column = 1
        Grafico.Data = Format(Dg_valor1(Dl_NCol), "###,##0.00")
    Next Dl_NCol
 
    Exit Sub

End Sub

Private Sub Msk_CepCor_GotFocus()
    Call Posiciona_cursor(Msk_CepCor)
End Sub

 
Private Sub msk_CNPJCli_GotFocus()
   msk_CNPJCli.ToolTipText = "                  "
   
   Call Posiciona_cursor(msk_CNPJCli)
   If Opt_Fisica.Value = False And Opt_Juridica.Value = False Then
      MsgBox "Escolha o tipo do cadastrado", vbExclamation, "A T E N Ç Ã O"
      Opt_Juridica.SetFocus
   End If
   
   If Opt_Juridica.Value = True Then
      msk_CNPJCli.Mask = "##.###.###/####-##"
   End If
   If Opt_Fisica.Value = True Then
      msk_CNPJCli.Mask = "   ###.###.###-##"
   End If
   Call Posiciona_cursor(msk_CNPJCli)
End Sub

Private Sub msk_CNPJCli_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
     msk_CNPJCli = Sg_CNPJCli
  End If
End Sub


Private Sub msk_CNPJCli_LostFocus()
   If Trim(msk_CNPJCli) = "" Then
      Exit Sub
   End If
   
   
   Sg_CodCli = Trim(msk_CNPJCli.ClipText)
   If flag_Cliente = "I" Then
    Sl_Desc = "SELECT * FROM Tba_Clientes "
    Sl_Desc = Sl_Desc & "WHERE Cli_Codigo = '" & Sg_CodCli & "' "
    Set Rst_Cliente = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
        If Rst_Cliente.EOF = False Then
           txt_NomCli = Rst_Cliente("Cli_Nome")
           MsgBox "Cliente já cadastrado!", vbInformation, "A T E N Ç Ã O"
           msk_CNPJCli = ""
           msk_CNPJCli.SetFocus
           Exit Sub
        End If
   End If
  
     If flag_Cliente = "I" Then
      If Opt_Juridica.Value = True Then
       If Not ConfereCPFCNPJ(1, msk_CNPJCli.ClipText) Then
          MsgBox "CNPJ do Cliente inválido", vbExclamation, "A T E N Ç Ã O"
''          Sst_Cliente.Tab = 0
''          Sst_Cliente.TabEnabled(0) = False
''          Sst_Cliente.TabEnabled(1) = False
''          Sst_Cliente.TabEnabled(2) = False
''          Sst_Cliente.TabEnabled(3) = False
           msk_CNPJCli = ""
           msk_CNPJCli.SetFocus
          Exit Sub
       End If
      ElseIf Opt_Fisica.Value = True Then
            If Not ConfereCPFCNPJ(-1, msk_CNPJCli.ClipText) Then
               MsgBox "CPF do Cliente inválido", vbExclamation, "A T E N Ç Ã O"
''               Sst_Cliente.Tab = 0
''               Sst_Cliente.TabEnabled(0) = False
''               Sst_Cliente.TabEnabled(1) = False
''               Sst_Cliente.TabEnabled(2) = False
''               Sst_Cliente.TabEnabled(3) = False
           msk_CNPJCli = ""
           msk_CNPJCli.SetFocus
               Exit Sub
            End If
      End If
    End If
   
    Cbo_NomeStatus.ListIndex = 2
    Cbo_CodigoStatus.ListIndex = Cbo_NomeStatus.ListIndex

End Sub

 

Private Sub opt_Fisica_Click()
If flag_Cliente = "I" Then
  msk_CNPJCli = ""
  msk_CNPJCli.SetFocus
End If
End Sub

Private Sub opt_Juridica_Click()
If flag_Cliente = "I" Then
   msk_CNPJCli = ""
End If

End Sub

Private Sub opt_outros_Click()
If flag_Cliente = "I" Then
  msk_CNPJCli = ""
  msk_CNPJCli.SetFocus
End If
End Sub

 

Private Sub Text1_GotFocus()
   Sst_Cliente = 0
   If flag_Cliente = "I" Then
      msk_CNPJCli.SetFocus
   Else
      txt_NomCli.SetFocus
   End If
   
End Sub

Private Sub Text2_GotFocus()
   Sst_Cliente = 2
'   cbo_Paises.SetFocus

End Sub

Private Sub Text3_GotFocus()
   Sst_Cliente = 1
   Txt_EndCor.SetFocus
End Sub

Private Sub Text4_GotFocus()
  Sst_Cliente.Tab = 0
  Txt_EndCor.SetFocus
End Sub

Private Sub Text5_GotFocus()
   Sst_Cliente.Tab = 2
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub SSCommand4_Click()

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Bto_Incluir.Enabled = False
    Bto_Excluir.Enabled = True
    Bto_AtualizaContato.Enabled = True
    Txt_NomeContato = ""
    Txt_Telefone = ""
    Txt_TelefoneRamal = ""
    Txt_Fax = ""
    Txt_FaxRamal = ""
    Txt_Celular = ""
    Txt_Email = ""
    Cbo_CodCargo.ListIndex = -1
    Cbo_CodCargo.ListIndex = -1
    Cbo_NomCargo.ListIndex = -1


    If Trim(TreeView1.SelectedItem) = "Contatos" Then
       Bto_AtualizaContato.Enabled = False
       Exit Sub
    End If
    Bto_AtualizaContato.Enabled = True
    Sl_Desc = "select * "
    Sl_Desc = Sl_Desc & " from Tba_Contatos "
    Sl_Desc = Sl_Desc & " where Con_Cliente = '" & Sg_CodCli & "'"
    Sl_Desc = Sl_Desc & "   and Con_Codigo  =  " & Val(Mid(TreeView1.SelectedItem, 1, 3))
    Set Rst_Contatos = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    If Rst_Contatos.EOF = False Then
       Txt_NomeContato = Rst_Contatos!Con_Nome
       Txt_Telefone = Rst_Contatos!Con_Telefone
       Txt_TelefoneRamal = Rst_Contatos!Con_TelefoneRamal
       Txt_Fax = Rst_Contatos!Con_Fax
       Txt_FaxRamal = Rst_Contatos!Con_FaxRamal
       Txt_Celular = Rst_Contatos!Con_Celular
       Txt_Email = Rst_Contatos!Con_Email
       If Trim(Rst_Contatos!Con_Cargo) <> "" Then
          Cbo_CodCargo = Rst_Contatos!Con_Cargo
         Cbo_NomCargo.ListIndex = Cbo_CodCargo.ListIndex
       End If
    End If

End Sub

Private Sub TreeView2_Click()
 

   If TreeView2.Nodes.Item(TreeView2.SelectedItem.Index).Expanded = True Then
       TreeView2.Nodes.Item(TreeView2.SelectedItem.Index).ForeColor = &HFF&
   Else
      TreeView2.Nodes.Item(TreeView2.SelectedItem.Index).ForeColor = &H80000012
   End If

 

End Sub

 

Private Sub TreeView2_NodeClick(ByVal Node As MSComctlLib.Node)
    Bto_Check.Enabled = True
    Bto_CheckAlt.Enabled = True
    Sg_Texto = TreeView2.Nodes.Item(TreeView2.SelectedItem.Index).FullPath
    Sg_CodigoFilho = TreeView2.Nodes.Item(TreeView2.SelectedItem.Index).Key
    Sg_CodigoPai = TreeView2.Nodes.Item(TreeView2.SelectedItem.Index).Parent.Key
    Sg_Descricao = TreeView2.Nodes.Item(TreeView2.SelectedItem.Index).Text
    sg_Proximo_codigoFilho = TreeView2.Nodes.Item(TreeView2.SelectedItem.Index).LastSibling.Key
    Dl_NumCar = Len(TreeView2.Nodes.Item(TreeView2.SelectedItem.Index).LastSibling.Key)
    sg_Proximo_codigoFilho = "'" & Val(Mid(sg_Proximo_codigoFilho, 2, Dl_NumCar - 2)) + 1 & "'"
    
End Sub

Private Sub Txt_BaiCor_GotFocus()
    Call Posiciona_cursor(Txt_BaiCor)
End Sub
 
 
 

