VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{368CC970-FF03-11D7-9B5A-000B6A03449D}#1.1#0"; "Combo_DB.OCX"
Object = "{BEACF734-D8AC-11D7-9B57-000B6A03449D}#2.0#0"; "Masked.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmConhecimento 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manutenção de Pedidos"
   ClientHeight    =   10320
   ClientLeft      =   -435
   ClientTop       =   435
   ClientWidth     =   15240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmConhecimento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   15240
   Begin Crystal.CrystalReport rptcontprop 
      Left            =   240
      Top             =   9000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1695
      Left            =   12360
      TabIndex        =   104
      Top             =   8640
      Visible         =   0   'False
      Width           =   2895
      Begin Project_Masked.Masked T100Rep 
         Height          =   375
         Left            =   360
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         FormatoString   =   "####0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         AutoTab         =   -1  'True
         ForeColor       =   64
         ValInteiro      =   8
      End
      Begin Project_Masked.Masked T100Cli 
         Height          =   375
         Left            =   1320
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         FormatoString   =   "####0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         AutoTab         =   -1  'True
         ForeColor       =   64
         ValInteiro      =   8
      End
      Begin Project_Masked.Masked T100Ped 
         Height          =   375
         Left            =   1320
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         FormatoString   =   "####0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         AutoTab         =   -1  'True
         ForeColor       =   64
         ValInteiro      =   8
      End
      Begin VB.Label Label44 
         BackColor       =   &H000080FF&
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
         Height          =   375
         Left            =   360
         TabIndex        =   111
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label41 
         BackColor       =   &H000080FF&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1320
         TabIndex        =   109
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label43 
         BackColor       =   &H0080C0FF&
         Caption         =   "  Indicador Tubo 100 (%)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   107
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label42 
         BackColor       =   &H000080FF&
         Caption         =   "Repres."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   106
         Top             =   360
         Width           =   975
      End
   End
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
      Height          =   975
      Left            =   11160
      Picture         =   "FrmConhecimento.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton BtoGrava 
      BackColor       =   &H80000013&
      Caption         =   "&Gravar Pedido"
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
      Left            =   6600
      Picture         =   "FrmConhecimento.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton BtoLimpaCTRC 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Limpar"
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
      Left            =   8520
      Picture         =   "FrmConhecimento.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   6375
      Left            =   12360
      TabIndex        =   47
      Top             =   3960
      Width           =   2895
      Begin VB.Frame Frame7 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1335
         Left            =   0
         TabIndex        =   65
         Top             =   3120
         Width           =   2895
         Begin Project_Masked.Masked MskMargem 
            Height          =   495
            Left            =   960
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            FormatoString   =   "####0.00"
            ValDecimal      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   12648447
            AutoTab         =   -1  'True
            ForeColor       =   65280
            ValInteiro      =   8
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000000FF&
            FillColor       =   &H000000FF&
            Height          =   1335
            Left            =   120
            Shape           =   2  'Oval
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label Label23 
            BackColor       =   &H0080C0FF&
            Caption         =   " %"
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   1680
            TabIndex        =   67
            Top             =   600
            Width           =   375
         End
      End
      Begin Project_Masked.Masked vl1 
         Height          =   375
         Left            =   120
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         FormatoString   =   "##,##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   65280
         AutoTab         =   -1  'True
         ForeColor       =   16384
         ValInteiro      =   8
      End
      Begin Project_Masked.Masked vl2 
         Height          =   375
         Left            =   120
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         FormatoString   =   "##,##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         AutoTab         =   -1  'True
         ForeColor       =   -2147483647
         ValInteiro      =   8
      End
      Begin Project_Masked.Masked vl3 
         Height          =   375
         Left            =   120
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         FormatoString   =   "##,##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         AutoTab         =   -1  'True
         ForeColor       =   12582912
         ValInteiro      =   8
      End
      Begin VB.Label LblIN 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1320
         TabIndex        =   61
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label LblPerIN 
         BackColor       =   &H000000FF&
         Caption         =   " %"
         Height          =   375
         Left            =   2160
         TabIndex        =   60
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label LblTextoideal 
         BackColor       =   &H000000FF&
         Caption         =   "Valor Tabela (Normal)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label LblIdeal 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1320
         TabIndex        =   58
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblpercideal 
         BackColor       =   &H000000FF&
         Caption         =   " %"
         Height          =   375
         Left            =   2160
         TabIndex        =   57
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label11 
         BackColor       =   &H000000FF&
         Caption         =   " %"
         Height          =   375
         Left            =   2160
         TabIndex        =   55
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label LblI 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   1320
         TabIndex        =   54
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label20 
         BackColor       =   &H000000FF&
         Caption         =   "Valor Realizado (Pedido)"
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
         Left            =   120
         TabIndex        =   51
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label19 
         BackColor       =   &H000000FF&
         Caption         =   "Mínimo p/ Venda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Caption         =   "Cliente"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   13800
      TabIndex        =   30
      Top             =   0
      Width           =   1455
      Begin VB.TextBox d9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   120
         TabIndex        =   113
         Text            =   "SimBahia"
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin Project_Masked.Masked d6 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoTab         =   -1  'True
         ForeColor       =   16711680
      End
      Begin Project_Masked.Masked d7 
         Height          =   375
         Left            =   120
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1320
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoTab         =   -1  'True
         ForeColor       =   16711680
      End
      Begin Project_Masked.Masked d8 
         Height          =   375
         Left            =   120
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         FormatoString   =   "##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoTab         =   -1  'True
         ForeColor       =   16711680
         ValInteiro      =   3
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFF80&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FF0000&
         Caption         =   "Contribuinte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF0000&
         Caption         =   "UF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FF0000&
         Caption         =   "ICMS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Caption         =   "Descontos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3975
      Left            =   12360
      TabIndex        =   19
      Top             =   0
      Width           =   1455
      Begin Project_Masked.Masked d1 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         FormatoString   =   "##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoTab         =   -1  'True
         ForeColor       =   128
         ValInteiro      =   3
      End
      Begin Project_Masked.Masked d2 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         FormatoString   =   "##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoTab         =   -1  'True
         ForeColor       =   128
         ValInteiro      =   3
      End
      Begin Project_Masked.Masked d3 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2040
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         FormatoString   =   "##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoTab         =   -1  'True
         ForeColor       =   128
         ValInteiro      =   3
      End
      Begin Project_Masked.Masked d4 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2760
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         FormatoString   =   "##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoTab         =   -1  'True
         ForeColor       =   128
         ValInteiro      =   3
      End
      Begin Project_Masked.Masked d5 
         Height          =   375
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         FormatoString   =   "##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoTab         =   -1  'True
         ForeColor       =   128
         ValInteiro      =   3
      End
      Begin VB.Label LblC 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   116
         Top             =   3480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Descontos"
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
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   975
      End
      Begin VB.Label label10 
         BackColor       =   &H000000C0&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H000000C0&
         Caption         =   "Frete (FOB)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H000000C0&
         Caption         =   "Cond.Pagto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000C0&
         Caption         =   "Promocional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H000000C0&
         Caption         =   "Padrão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTConhec 
      DragIcon        =   "FrmConhecimento.frx":1108
      Height          =   8655
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   15266
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12648447
      ForeColor       =   8388736
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pedido"
      TabPicture(0)   =   "FrmConhecimento.frx":1252
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label28"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblRotaRec"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblRotaPag"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LblBloqPed"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LblSimples"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LblTot"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LblDesc"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "LblSub"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "LblVlSimples"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "tab_simulacao_pedido"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Novacor"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ChkKit"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Status"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "CboCondPag"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "FraSenha"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "CmdImpr"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "CmdLibera"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cboCli"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "FraParametro"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "FraGrupo"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Comunicação"
      TabPicture(1)   =   "FrmConhecimento.frx":126E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdUP"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "CmdDN"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TxtNegocio"
      Tab(1).Control(3)=   "CmdCancelar"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame10"
      Tab(1).Control(5)=   "TxtObserva"
      Tab(1).Control(6)=   "Label46"
      Tab(1).Control(7)=   "Label24"
      Tab(1).Control(8)=   "LblResultNegocio"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Posição do Cliente"
      TabPicture(2)   =   "FrmConhecimento.frx":128A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(1)=   "LblFone"
      Tab(2).Control(2)=   "Label45"
      Tab(2).Control(3)=   "LblSimBahia"
      Tab(2).Control(4)=   "LblCliUltimo"
      Tab(2).Control(5)=   "Label40"
      Tab(2).Control(6)=   "LblCliPrimeira"
      Tab(2).Control(7)=   "Label38"
      Tab(2).Control(8)=   "Label37"
      Tab(2).Control(9)=   "Label36"
      Tab(2).Control(10)=   "Label35"
      Tab(2).Control(11)=   "Label34"
      Tab(2).Control(12)=   "Label33"
      Tab(2).Control(13)=   "Label32"
      Tab(2).Control(14)=   "Label29"
      Tab(2).Control(15)=   "Label31"
      Tab(2).Control(16)=   "LblCliCep"
      Tab(2).Control(17)=   "LblCliInscr"
      Tab(2).Control(18)=   "LblCliCid"
      Tab(2).Control(19)=   "LblCliContr"
      Tab(2).Control(20)=   "LblCliUF"
      Tab(2).Control(21)=   "LblCliBairro"
      Tab(2).Control(22)=   "LblCliCGC"
      Tab(2).Control(23)=   "lblCliEndereco"
      Tab(2).Control(24)=   "LblCliNome"
      Tab(2).ControlCount=   25
      TabCaption(3)   =   "Entregas"
      TabPicture(3)   =   "FrmConhecimento.frx":12A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text1"
      Tab(3).Control(1)=   "GrdEntrega"
      Tab(3).ControlCount=   2
      Begin VB.Frame FraGrupo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7095
         Left            =   0
         TabIndex        =   145
         Top             =   3000
         Visible         =   0   'False
         Width           =   12375
         Begin VB.CommandButton CmdRetorna 
            BackColor       =   &H00C0FFC0&
            Caption         =   "&Retornar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   11040
            Picture         =   "FrmConhecimento.frx":12C2
            Style           =   1  'Graphical
            TabIndex        =   146
            Top             =   5880
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid GrdGrupo 
            Height          =   6855
            Left            =   480
            TabIndex        =   147
            Top             =   4800
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   12091
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            BackColor       =   14737632
            ForeColor       =   16384
            BackColorFixed  =   16384
            ForeColorFixed  =   16777215
            BackColorSel    =   65535
            ForeColorSel    =   255
            BackColorBkg    =   14737632
            GridColor       =   16777215
            GridColorFixed  =   12632256
            ScrollBars      =   2
            SelectionMode   =   1
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid GrdProduto 
            Height          =   5535
            Left            =   4320
            TabIndex        =   148
            Top             =   4200
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   9763
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            BackColor       =   16777215
            ForeColor       =   16384
            BackColorFixed  =   32768
            ForeColorFixed  =   16777215
            BackColorSel    =   65535
            ForeColorSel    =   192
            BackColorBkg    =   16777215
            GridColor       =   49152
            ScrollBars      =   2
            SelectionMode   =   1
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FraParametro 
         BackColor       =   &H00FFECEC&
         Caption         =   "Parâmetros para Simulação (Cliente)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   975
         Left            =   3360
         TabIndex        =   138
         Top             =   0
         Visible         =   0   'False
         Width           =   8895
         Begin VB.ComboBox CboSimBA 
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   345
            ItemData        =   "FrmConhecimento.frx":1704
            Left            =   7200
            List            =   "FrmConhecimento.frx":1711
            TabIndex        =   143
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox CboContribuinte 
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   345
            ItemData        =   "FrmConhecimento.frx":1726
            Left            =   3840
            List            =   "FrmConhecimento.frx":1733
            TabIndex        =   140
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox CboUFCli 
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   345
            ItemData        =   "FrmConhecimento.frx":1748
            Left            =   720
            List            =   "FrmConhecimento.frx":175B
            TabIndex        =   139
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label LblSim 
            BackColor       =   &H00FFECEC&
            Caption         =   "Simples(BA)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   5760
            TabIndex        =   144
            Top             =   360
            Width           =   1290
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFECEC&
            Caption         =   "Contribuinte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   142
            Top             =   360
            Width           =   1290
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFECEC&
            Caption         =   "UF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   141
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   -67080
         TabIndex        =   137
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdUP 
         BackColor       =   &H000000FF&
         Height          =   720
         Left            =   -63240
         Picture         =   "FrmConhecimento.frx":1773
         Style           =   1  'Graphical
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton CmdDN 
         BackColor       =   &H000000FF&
         Height          =   720
         Left            =   -63240
         Picture         =   "FrmConhecimento.frx":1BB5
         Style           =   1  'Graphical
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   3600
         Width           =   495
      End
      Begin RichTextLib.RichTextBox TxtNegocio 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   840
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7011
         _Version        =   393217
         BackColor       =   8454143
         Enabled         =   0   'False
         Appearance      =   0
         TextRTF         =   $"FrmConhecimento.frx":1FF7
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
      Begin Project_Combo_DB.Combo_DB cboCli 
         Height          =   375
         Left            =   4560
         TabIndex        =   1
         Top             =   600
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   661
         Cols            =   0
         Cabecalho       =   -1  'True
      End
      Begin VB.CommandButton CmdLibera 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Li&bera"
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
         Left            =   9960
         Picture         =   "FrmConhecimento.frx":2070
         Style           =   1  'Graphical
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancelar 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cancelar Pedido"
         Enabled         =   0   'False
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
         Left            =   -65160
         Picture         =   "FrmConhecimento.frx":24B2
         Style           =   1  'Graphical
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   7560
         Width           =   1215
      End
      Begin VB.CommandButton CmdImpr 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Imprimir"
         Enabled         =   0   'False
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
         Left            =   9960
         Picture         =   "FrmConhecimento.frx":28F4
         Style           =   1  'Graphical
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   7560
         Width           =   1095
      End
      Begin VB.Frame FraSenha 
         Caption         =   "Digite a Senha"
         ForeColor       =   &H00008000&
         Height          =   1575
         Left            =   6240
         TabIndex        =   117
         Top             =   4200
         Visible         =   0   'False
         Width           =   2895
         Begin VB.TextBox TxtPwdoper 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   8
            PasswordChar    =   "*"
            TabIndex        =   118
            Top             =   435
            Width           =   2220
         End
         Begin VB.Label lblmens 
            BackColor       =   &H00000000&
            Caption         =   "Senha Inválida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Left            =   240
            TabIndex        =   119
            Top             =   1080
            Visible         =   0   'False
            Width           =   2250
         End
      End
      Begin VB.Frame Frame10 
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
         Height          =   1215
         Left            =   -67320
         TabIndex        =   115
         Top             =   5760
         Width           =   4335
         Begin VB.TextBox TxtChave 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
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
            Left            =   240
            TabIndex        =   53
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   -75000
         TabIndex        =   87
         Top             =   2400
         Width           =   12375
         Begin MSFlexGridLib.MSFlexGrid GrdCliVencer 
            Height          =   1575
            Left            =   6600
            TabIndex        =   88
            Top             =   1080
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
            TabIndex        =   89
            Top             =   1080
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
            TabIndex        =   90
            Top             =   3480
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
            TabIndex        =   99
            Top             =   600
            Width           =   2415
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
            TabIndex        =   98
            Top             =   600
            Width           =   2295
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
            TabIndex        =   97
            Top             =   3000
            Width           =   2295
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
            TabIndex        =   96
            Top             =   600
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
            TabIndex        =   95
            Top             =   600
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
            TabIndex        =   94
            Top             =   600
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
            TabIndex        =   93
            Top             =   600
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
            TabIndex        =   92
            Top             =   3000
            Width           =   495
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
            TabIndex        =   91
            Top             =   3000
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtObserva 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   52
         Top             =   5280
         Width           =   6855
      End
      Begin Project_Combo_DB.Combo_DB CboCondPag 
         Height          =   375
         Left            =   4560
         TabIndex        =   2
         Top             =   1200
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   661
         Cols            =   0
         Cabecalho       =   -1  'True
         ModoAtualizacao =   0
      End
      Begin VB.Frame Frame2 
         Caption         =   "Transporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1035
         Left            =   8640
         TabIndex        =   14
         Top             =   1080
         Width           =   3615
         Begin VB.TextBox TxtTransp 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   600
            Width           =   3375
         End
         Begin VB.OptionButton Opt_CIF 
            Caption         =   "CIF"
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
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton Opt_FOB 
            Caption         =   "FOB"
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
            Left            =   960
            TabIndex        =   15
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Número do Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1335
         Left            =   120
         TabIndex        =   13
         Top             =   420
         Width           =   2775
         Begin MSMask.MaskEdBox MskNroPedido 
            Height          =   555
            Left            =   240
            TabIndex        =   0
            Top             =   480
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   979
            _Version        =   393216
            BackColor       =   -2147483624
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
      End
      Begin VB.CommandButton Status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox ChkKit 
         Caption         =   "Kit Irrigação"
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
         Left            =   3360
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid GrdEntrega 
         Height          =   6855
         Left            =   -75000
         TabIndex        =   136
         Top             =   360
         Width           =   12315
         _ExtentX        =   21722
         _ExtentY        =   12091
         _Version        =   393216
         Rows            =   1
         Cols            =   13
         FixedCols       =   7
         BackColor       =   14737632
         ForeColor       =   0
         BackColorFixed  =   4210688
         ForeColorFixed  =   16777215
         BackColorSel    =   65535
         ForeColorSel    =   0
         BackColorBkg    =   -2147483633
         GridColor       =   16777215
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Novacor 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   1800
         Width           =   2775
      End
      Begin TabDlg.SSTab tab_simulacao_pedido 
         Height          =   5295
         Left            =   120
         TabIndex        =   150
         Top             =   2160
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   9340
         _Version        =   393216
         Tabs            =   8
         TabsPerRow      =   4
         TabHeight       =   520
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "PREDIAL ESGOTO BRANCO"
         TabPicture(0)   =   "FrmConhecimento.frx":2A3E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label39"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label30"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "tubos_conexoes_predial"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Bto_aplica__tubos_conexoes_predial"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "desconto_tubos_conexoes_predial"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "PREDIAL SOLDÁVEL MARROM"
         TabPicture(1)   =   "FrmConhecimento.frx":2A5A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label8"
         Tab(1).Control(1)=   "Label2"
         Tab(1).Control(2)=   "tubos_conexoes_agua"
         Tab(1).Control(3)=   "Bto_aplica_tubos_conexoes_agua"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "desconto_tubos_conexoes_agua"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "PREDIAL ROSCÁVEL BRANCO"
         TabPicture(2)   =   "FrmConhecimento.frx":2A76
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label48"
         Tab(2).Control(1)=   "Label47"
         Tab(2).Control(2)=   "tubos_conexoes_roscaveis"
         Tab(2).Control(3)=   "Bto_aplica__tubos_conexoes_roscaveis"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "desconto_tubos_conexoes_roscaveis"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "IRRIGA AGROPECUÁRIO AZUL"
         TabPicture(3)   =   "FrmConhecimento.frx":2A92
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label68"
         Tab(3).Control(1)=   "Label69"
         Tab(3).Control(2)=   "tubos_conexoes_irri_azuis"
         Tab(3).Control(3)=   "desconto_tubos_conexoes_irri_azuis"
         Tab(3).Control(4)=   "Bto_Aplica_tubos_conexoes_irri_azuis"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).ControlCount=   5
         TabCaption(4)   =   "INFRA ESGOTO OCRE"
         TabPicture(4)   =   "FrmConhecimento.frx":2AAE
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "desconto_tubos_conexoes_coletor_esgoto_ocre"
         Tab(4).Control(1)=   "Bto_aplica_tubos_conexoes_coletor_esgoto_ocre"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "tubo_tubos_conexoes_coletor_esgoto"
         Tab(4).Control(3)=   "Label64"
         Tab(4).Control(4)=   "Label65"
         Tab(4).ControlCount=   5
         TabCaption(5)   =   "INFRA PBA JUNTA ELÁSTICA MARROM"
         TabPicture(5)   =   "FrmConhecimento.frx":2ACA
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Label63"
         Tab(5).Control(1)=   "Label62"
         Tab(5).Control(2)=   "tubo_conexoes_pba"
         Tab(5).Control(3)=   "Bto_aplica_tubos_conexoes_pba"
         Tab(5).Control(3).Enabled=   0   'False
         Tab(5).Control(4)=   "desconto_tubos_conexoes_pba"
         Tab(5).ControlCount=   5
         TabCaption(6)   =   "INFRA DEFOFO AZUL"
         TabPicture(6)   =   "FrmConhecimento.frx":2AE6
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Label61"
         Tab(6).Control(1)=   "Label60"
         Tab(6).Control(2)=   "tubo_conexoes_defofo"
         Tab(6).Control(3)=   "desconto_tubos_conexoes_defofo"
         Tab(6).Control(4)=   "Bto_aplica_tubos_conexoes_defofo"
         Tab(6).Control(4).Enabled=   0   'False
         Tab(6).ControlCount=   5
         TabCaption(7)   =   "RESUMO"
         TabPicture(7)   =   "FrmConhecimento.frx":2B02
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Label66"
         Tab(7).Control(1)=   "Label67"
         Tab(7).Control(2)=   "GrdNotaCliente"
         Tab(7).Control(3)=   "Bto_Aplica_resumo"
         Tab(7).Control(3).Enabled=   0   'False
         Tab(7).Control(4)=   "desconto_resumo"
         Tab(7).ControlCount=   5
         Begin VB.CommandButton Bto_Aplica_tubos_conexoes_irri_azuis 
            BackColor       =   &H00FF0000&
            Height          =   480
            Left            =   -73560
            Picture         =   "FrmConhecimento.frx":2B1E
            Style           =   1  'Graphical
            TabIndex        =   199
            TabStop         =   0   'False
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox desconto_tubos_conexoes_irri_azuis 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -74760
            TabIndex        =   198
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox desconto_tubos_conexoes_roscaveis 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -74760
            TabIndex        =   164
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CommandButton Bto_aplica__tubos_conexoes_roscaveis 
            BackColor       =   &H00FF0000&
            Height          =   480
            Left            =   -73560
            Picture         =   "FrmConhecimento.frx":2F60
            Style           =   1  'Graphical
            TabIndex        =   163
            TabStop         =   0   'False
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox desconto_tubos_conexoes_agua 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -74760
            TabIndex        =   162
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton Bto_aplica_tubos_conexoes_agua 
            BackColor       =   &H00FF0000&
            Height          =   480
            Left            =   -73560
            Picture         =   "FrmConhecimento.frx":33A2
            Style           =   1  'Graphical
            TabIndex        =   161
            TabStop         =   0   'False
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox desconto_tubos_conexoes_predial 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   240
            TabIndex        =   160
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton Bto_aplica__tubos_conexoes_predial 
            BackColor       =   &H00FF0000&
            Height          =   480
            Left            =   1440
            Picture         =   "FrmConhecimento.frx":37E4
            Style           =   1  'Graphical
            TabIndex        =   159
            TabStop         =   0   'False
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox desconto_tubos_conexoes_coletor_esgoto_ocre 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -74760
            TabIndex        =   158
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton Bto_aplica_tubos_conexoes_coletor_esgoto_ocre 
            BackColor       =   &H00FF0000&
            Height          =   480
            Left            =   -73560
            Picture         =   "FrmConhecimento.frx":3C26
            Style           =   1  'Graphical
            TabIndex        =   157
            TabStop         =   0   'False
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox desconto_tubos_conexoes_pba 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -74760
            TabIndex        =   156
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton Bto_aplica_tubos_conexoes_pba 
            BackColor       =   &H00FF0000&
            Height          =   480
            Left            =   -73560
            Picture         =   "FrmConhecimento.frx":4068
            Style           =   1  'Graphical
            TabIndex        =   155
            TabStop         =   0   'False
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton Bto_aplica_tubos_conexoes_defofo 
            BackColor       =   &H00FF0000&
            Height          =   480
            Left            =   -73560
            Picture         =   "FrmConhecimento.frx":44AA
            Style           =   1  'Graphical
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox desconto_tubos_conexoes_defofo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -74760
            TabIndex        =   153
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox desconto_resumo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   -74760
            TabIndex        =   152
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton Bto_Aplica_resumo 
            BackColor       =   &H00FF0000&
            Height          =   480
            Left            =   -73560
            Picture         =   "FrmConhecimento.frx":48EC
            Style           =   1  'Graphical
            TabIndex        =   151
            TabStop         =   0   'False
            Top             =   1080
            Width           =   375
         End
         Begin MSFlexGridLib.MSFlexGrid tubos_conexoes_agua 
            Height          =   3495
            Left            =   -74880
            TabIndex        =   165
            Top             =   1680
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   6165
            _Version        =   393216
            Rows            =   6
            Cols            =   18
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
         Begin MSFlexGridLib.MSFlexGrid tubos_conexoes_predial 
            Height          =   3495
            Left            =   120
            TabIndex        =   166
            Top             =   1680
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   6165
            _Version        =   393216
            Rows            =   6
            Cols            =   18
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
         Begin MSFlexGridLib.MSFlexGrid tubos_conexoes_roscaveis 
            Height          =   3375
            Left            =   -74880
            TabIndex        =   167
            Top             =   1800
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   5953
            _Version        =   393216
            Rows            =   6
            Cols            =   18
            BackColor       =   16777215
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
         Begin MSFlexGridLib.MSFlexGrid tubo_tubos_conexoes_coletor_esgoto 
            Height          =   3495
            Left            =   -74880
            TabIndex        =   168
            Top             =   1680
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   6165
            _Version        =   393216
            Rows            =   6
            Cols            =   18
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
         Begin MSFlexGridLib.MSFlexGrid tubo_conexoes_pba 
            Height          =   3495
            Left            =   -74880
            TabIndex        =   169
            Top             =   1680
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   6165
            _Version        =   393216
            Rows            =   6
            Cols            =   18
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
         Begin MSFlexGridLib.MSFlexGrid tubo_conexoes_defofo 
            Height          =   3495
            Left            =   -74880
            TabIndex        =   170
            Top             =   1680
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   6165
            _Version        =   393216
            Rows            =   6
            Cols            =   18
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
         Begin MSFlexGridLib.MSFlexGrid GrdNotaCliente 
            Height          =   3375
            Left            =   -74880
            TabIndex        =   171
            Top             =   1680
            Width           =   11715
            _ExtentX        =   20664
            _ExtentY        =   5953
            _Version        =   393216
            Rows            =   1
            Cols            =   20
            FixedCols       =   0
            BackColor       =   12648447
            ForeColor       =   16711680
            BackColorFixed  =   12582912
            ForeColorFixed  =   16777215
            BackColorSel    =   65535
            ForeColorSel    =   16711680
            BackColorBkg    =   16777215
            GridColor       =   8438015
            SelectionMode   =   1
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid tubos_conexoes_irri_azuis 
            Height          =   3375
            Left            =   -74880
            TabIndex        =   200
            Top             =   1800
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   5953
            _Version        =   393216
            Rows            =   6
            Cols            =   18
            BackColor       =   16777215
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
         Begin VB.Label Label69 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   202
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label68 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   -74520
            TabIndex        =   201
            Top             =   900
            Width           =   825
         End
         Begin VB.Label Label67 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   197
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label66 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   -74520
            TabIndex        =   196
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label47 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   -74520
            TabIndex        =   195
            Top             =   900
            Width           =   825
         End
         Begin VB.Label Label48 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   194
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label49 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   -74520
            TabIndex        =   193
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label50 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   192
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label51 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   -74520
            TabIndex        =   191
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label52 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   190
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label53 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   -74520
            TabIndex        =   189
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label54 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   188
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label55 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   -74520
            TabIndex        =   187
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label56 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   186
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label57 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   185
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   4
            Left            =   -74520
            TabIndex        =   184
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label58 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   -74520
            TabIndex        =   183
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label59 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   182
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   181
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label8 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   -74520
            TabIndex        =   180
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label30 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   179
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label39 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   480
            TabIndex        =   178
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label60 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   177
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label61 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   -74520
            TabIndex        =   176
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label62 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   175
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label63 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   -74520
            TabIndex        =   174
            Top             =   780
            Width           =   825
         End
         Begin VB.Label Label64 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   173
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label65 
            Caption         =   "Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   -74520
            TabIndex        =   172
            Top             =   780
            Width           =   825
         End
      End
      Begin VB.Label Label46 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Notificações"
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   -74880
         TabIndex        =   133
         Top             =   480
         Width           =   1695
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
         TabIndex        =   131
         Top             =   1920
         Width           =   2055
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
         TabIndex        =   130
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label LblVlSimples 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   3000
         TabIndex        =   129
         Top             =   8040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LblSub 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   120
         TabIndex        =   128
         Top             =   8040
         Width           =   1455
      End
      Begin VB.Label LblDesc 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   1680
         TabIndex        =   127
         Top             =   8040
         Width           =   1215
      End
      Begin VB.Label LblTot 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   4680
         TabIndex        =   126
         Top             =   8040
         Width           =   1455
      End
      Begin VB.Label LblSimples 
         Caption         =   "Desc.Simples"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   3000
         TabIndex        =   125
         Top             =   7800
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   18
         Top             =   600
         Width           =   930
      End
      Begin VB.Label LblBloqPed 
         Caption         =   "Este pedido não pode ser alterado"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4920
         TabIndex        =   120
         Top             =   3120
         Visible         =   0   'False
         Width           =   6255
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
         TabIndex        =   112
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
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
         TabIndex        =   103
         Top             =   720
         Width           =   1335
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
         TabIndex        =   102
         Top             =   480
         Width           =   1335
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
         TabIndex        =   101
         Top             =   720
         Width           =   1335
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
         TabIndex        =   100
         Top             =   480
         Width           =   1215
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
         TabIndex        =   86
         Top             =   1680
         Width           =   615
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
         TabIndex        =   85
         Top             =   1080
         Width           =   1815
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
         TabIndex        =   84
         Top             =   480
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
         TabIndex        =   83
         Top             =   1680
         Width           =   615
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
         TabIndex        =   82
         Top             =   1680
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
         TabIndex        =   81
         Top             =   1680
         Width           =   855
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
         TabIndex        =   80
         Top             =   480
         Width           =   1815
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
         TabIndex        =   79
         Top             =   1080
         Width           =   1815
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
         TabIndex        =   78
         Top             =   1920
         Width           =   1215
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
         TabIndex        =   77
         Top             =   1320
         Width           =   2175
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
         TabIndex        =   76
         Top             =   1920
         Width           =   3375
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
         TabIndex        =   75
         Top             =   1320
         Width           =   1575
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
         TabIndex        =   74
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label LblCliCidade 
         Caption         =   "Cidade do cliente"
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
         Height          =   375
         Left            =   3120
         TabIndex        =   73
         Top             =   -3480
         Width           =   4095
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
         TabIndex        =   72
         Top             =   1920
         Width           =   3855
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
         TabIndex        =   71
         Top             =   720
         Width           =   1935
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
         TabIndex        =   70
         Top             =   1320
         Width           =   6855
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
         TabIndex        =   69
         Top             =   720
         Width           =   6855
      End
      Begin VB.Label Label24 
         Caption         =   "Observações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   68
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label LblResultNegocio 
         Height          =   495
         Left            =   -74880
         TabIndex        =   62
         Top             =   6960
         Width           =   8175
      End
      Begin VB.Label Label1 
         Caption         =   "Cond. Pagto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   17
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label Label4 
         Caption         =   "Sub-Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   7800
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Desconto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1680
         TabIndex        =   10
         Top             =   7800
         Width           =   1005
      End
      Begin VB.Label LblRotaPag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5655
         TabIndex        =   9
         Top             =   4035
         Width           =   5895
      End
      Begin VB.Label LblRotaRec 
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   2640
         Width           =   6615
      End
      Begin VB.Label Label28 
         Caption         =   "Valor Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   4680
         TabIndex        =   7
         Top             =   7800
         Width           =   1305
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GrdIndice 
      Height          =   1575
      Left            =   0
      TabIndex        =   37
      Top             =   8640
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   2778
      _Version        =   393216
      Rows            =   1
      Cols            =   18
      FixedCols       =   0
      BackColor       =   12648384
      ForeColor       =   192
      BackColorFixed  =   255
      ForeColorFixed  =   16777215
      ForeColorSel    =   255
      BackColorBkg    =   16777215
      GridColor       =   255
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1815
      Left            =   8880
      TabIndex        =   38
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
      Begin Project_Masked.Masked idx1 
         Height          =   375
         Left            =   1200
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         FormatoString   =   "##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   65280
         AutoTab         =   -1  'True
         ForeColor       =   -2147483647
         ValInteiro      =   8
      End
      Begin Project_Masked.Masked idx2 
         Height          =   375
         Left            =   240
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         FormatoString   =   "##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoTab         =   -1  'True
         ForeColor       =   -2147483647
         ValInteiro      =   8
      End
      Begin Project_Masked.Masked idx3 
         Height          =   375
         Left            =   240
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         FormatoString   =   "##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoTab         =   -1  'True
         ForeColor       =   -2147483647
         ValInteiro      =   8
      End
      Begin Project_Masked.Masked idx4 
         Height          =   375
         Left            =   1200
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         FormatoString   =   "##0.00"
         ValDecimal      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   65535
         AutoTab         =   -1  'True
         ForeColor       =   0
         ValInteiro      =   8
      End
      Begin VB.Label Label18 
         BackColor       =   &H00800080&
         Caption         =   "IDX Verde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1200
         TabIndex        =   46
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label17 
         BackColor       =   &H00800080&
         Caption         =   "Tot.Liq."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label16 
         BackColor       =   &H00800080&
         Caption         =   "Tot.Peso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00800080&
         Caption         =   "IDX Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   43
         Top             =   960
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmConhecimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lCol As Integer
Dim ilGrupo As Integer
Dim ilCelula As Integer
Dim slChave As String
Dim ilDescChave As Double
Dim blVencidos As Boolean
Dim blImpr As Boolean
Dim blDescZero As Boolean
Dim blEbahia As Boolean
Dim slPedSimples As String * 1
Dim blRetornoDupls As Boolean
Dim Num As Integer
Dim iDscRegiao As Integer
Dim bAVista As Boolean
Dim Datped As Date

Dim vlPercs As Variant
Dim dlCem As Double
Dim blLG As Boolean
Dim slIrriga As String * 1
Dim ilFlgKit As Integer
Dim blFechaComi As Boolean
Dim SlTabela As String * 1
Dim slClasCor As String * 1

Dim slUFOri As String

Dim slremet As String
Dim ilCodCli As Double ' TROQUEI EM 08/02/2011 Integer
Dim slUFCli As String
Dim slFlgContr As String
Dim slFlgSIMBa As String
Dim ilFlgSit As Integer

Dim ilCodCnd As Integer
Dim dlPerCusFin As Double
Dim dlPerDesCnd As Double
Dim ilQtdParCnd As Integer
Dim ilPrzMed As Integer

Dim ilIdeGrp As Integer
Dim ilQtdPar As Integer
Dim dlValUntN As Double
Dim dlValUntA As Double
Dim dlValUntB As Double
Dim dlMrgPrd As Double
Dim ilQtdEmb As Integer
Dim dlValCusUntQtd As Double
Dim dlValCusAdicQtd As Double
Dim dlAlqImpFed As Double

Dim dlValItem As Double
Dim ilNumTab As Integer
Dim ilind As Integer
Dim Linhas As Integer
Dim VlIdealItem As Double

Dim dlPerDesPrd As Double
Dim dlSumDscItem As Double
Dim dlSumDscItemORIG As Double

Dim dlPerContr As Double
Dim dlPerNContr As Double
Dim dlPerContrSIMBa As Double
Dim dlPerContrKit As Double
Dim dlPerNContrKit As Double
Dim dlPerContrSIMBaKit As Double

Dim dlPerTubo100Cli As Double
Dim dlPerTubo100Rep As Double

Dim dlPerDesPadrao As Double

Dim blimpa As Boolean
Dim blleitura As Boolean

Dim ilCodRep As Integer
Dim slUFRep As String
Dim dlPerDesRep As Double
Dim slNomRep As String
Dim dlPerCusFrt As Double
Dim dlPerComiN As Double
Dim dlPerComiA As Double
Dim dlPerComiB As Double
Dim dlPerDesFOB As Double
Dim dlPerDesFOBReal As Double
Dim slFlgSugComi As String
Dim dlIdxPDD As Double
Dim dlIdxAzul As Double
Dim dlValPedMin As Double
Dim dlValLimPrz1 As Double
Dim dlValParMin As Double
Dim ilPrzMed1 As Integer
Dim ilPrzMed2 As Integer

Dim dlAlqICMContr As Double
Dim dlAlqICMNContr As Double

Dim PerCusIcm As Double

Dim dlPesUnt As Double

'Testa validade de pedido para kit irrigação
Dim QtdTubo As Double
Dim QtdTuboRosc As Double
Dim QtdAspe As Double
Dim QtdConx As Double

'Para cálculo dos índices
Dim Linhas1 As Integer
Dim ilI As Integer
Dim ilInd1 As Integer
Dim blAchou As Boolean
Dim vlPercs1 As Variant
Dim dlCem1 As Double
Dim dlSumDscTot As Double
Dim dlSumGrd As Double
Dim dlSumSubGrd As Double
Dim dlSumPes As Double
Dim dlIDX As Double
Dim dlIDXA As Double
Dim dlIDXB As Double
Dim dlTotBru As Double
Dim dlTotIdeal As Double
Dim dlTotLiq As Double
Dim dlSimples As Double
Dim dlTotPes As Double
Dim dlaux As Double
Dim dlMediaIDX As Double
Dim dlIlb As Double
Dim dlIdeallb As Double
Dim dlSumIdealGrd As Double
Dim dlINlb As Double
Dim dlMargemGeral As Double
Dim slFlgAlt As String

'Para sugestão de comissão
Dim PerSugIni As Double
Dim PerSugFimA As Double
Dim PerSugFimB As Double
Dim PerSug1Ini As Double
Dim PerSug1FimA As Double
Dim PerSug1FimB As Double
Dim PerSug2Ini As Double
Dim PerSug2FimA As Double
Dim PerSug2FimB As Double
Dim ilNroSit As Integer
Dim dlIABS As Double
Dim dlComiSug As Double
Dim slAceita As Boolean
Dim dlPerComiCalc As Double
Dim dlPerComiNeg As Double
Dim Seqini As String
'Dim SeqIni As Double
'Dim SeqFim As Double
Dim SeqFim As String
Dim dlAlqICMContrKIT As Double
Dim dlAlqICMNContrKIT As Double
Dim dlAlqICMSimplesKIT As Double
Dim blModificar As Boolean
Dim blI As Integer

Dim dDscRegiao As Double
'Dim bAVista As Boolean
'Dim iDscRegiao As Integer
Dim iDscForaRegiao As Integer
Dim bEKit As Boolean
Dim bSo100 As Boolean
Dim dTotb100 As Double

'Para o controle do tab
Dim auxSelChange As Boolean
Dim auxChange As Boolean
Dim rowAux As Integer
Dim auxSelChangeSaldaveisAgua As Boolean
Dim auxChangeSaldaveisAgua As Boolean
Dim rowAuxSaldaveisAgua As Integer
Dim auxSelChangeRoscaveis As Boolean
Dim auxChangeRoscaveis  As Boolean
Dim rowAuxRoscaveis As Integer
Dim auxSelChangeTubosConexoesIrriAzuis As Boolean
Dim auxChangeTubosConexoesIrriAzuis  As Boolean
Dim rowAuxTubosConexoesIrriAzuis As Integer
Dim auxSelChangeTubosConexoesRoscaveis As Boolean
Dim auxChangeTubosConexoesRoscaveis  As Boolean
Dim rowAuxTubosConexoesRoscaveis As Integer
Dim auxSelChangeTuboConexoesPredial As Boolean
Dim auxChangeTuboConexoesPredial  As Boolean
Dim rowAuxTuboConexoesPredial As Integer
Dim auxSelChangeTuboConexoesAgua As Boolean
Dim auxChangeTuboConexoesAgua  As Boolean
Dim rowAuxTuboConexoesAgua As Integer

Dim auxSelChangeGrdNotaCliente As Boolean
Dim auxChangeGrdNotaCliente  As Boolean
Dim rowAuxGrdNotaCliente As Integer

Dim titleTab As String
Dim houveDigitacaoSoldaveisAgua As Boolean
Dim houveDigitacaoEsgoto As Boolean
Dim houveDigitacaoTubosRoscaveis As Boolean
Dim houveDigitacaoTubosConexoesIrriAzuis As Boolean
Dim houveDigitacaoTubosConexoesRoscaveis As Boolean
Dim houveDigitacaoTuboConexoesPredial As Boolean
Dim houveDigitacaoTuboConexoesAgua As Boolean
Dim houveDigitacaoGrdNotaCliente As Boolean
Dim MskSerie As Integer     'Quantidade
Dim MskVlrUnit As Double    'Valor Unitário
Dim MskDatEmiNf As Double   'Desconto
Dim MskNumNf As String      'Código do produto
Dim ControleLostFocus As Boolean
Dim ControleAtualizaGrid As Boolean

Private Sub CboCondPag_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        CboCondPag_LostFocus
    End If
    
End Sub


Private Sub desconto_tubos_conexoes_agua_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(desconto_tubos_conexoes_agua.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        Bto_aplica_tubos_conexoes_agua_Click
    End If
    
End Sub

Private Sub desconto_tubos_conexoes_defofo_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(desconto_tubos_conexoes_defofo.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        Bto_aplica_tubos_conexoes_defofo_Click
    End If
    
End Sub

Private Sub desconto_tubos_conexoes_coletor_esgoto_ocre_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(desconto_tubos_conexoes_coletor_esgoto_ocre.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        Bto_aplica_tubos_conexoes_coletor_esgoto_ocre_Click
    End If
    
End Sub

Private Sub desconto_tubos_conexoes_pba_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(desconto_tubos_conexoes_pba.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        Bto_aplica_tubos_conexoes_pba_Click
    End If
    
End Sub

Private Sub desconto_resumo_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(desconto_resumo.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        Bto_Aplica_resumo_Click
    End If
    
End Sub

Private Sub desconto_tubos_conexoes_roscaveis_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(desconto_tubos_conexoes_roscaveis.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        Bto_aplica__tubos_conexoes_roscaveis_Click
    End If
    
End Sub

Private Sub desconto_tubos_conexoes_irri_azuis_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(desconto_tubos_conexoes_irri_azuis.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        Bto_Aplica_tubos_conexoes_irri_azuis_Click
    End If
    
End Sub

Private Sub desconto_tubos_conexoes_predial_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(desconto_tubos_conexoes_predial.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        Bto_aplica__tubos_conexoes_predial_Click
    End If
    
End Sub

Private Sub tubo_conexoes_defofo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If tubo_conexoes_defofo.col = 4 Or tubo_conexoes_defofo.col = 6 Then
            With tubo_conexoes_defofo
            'Zera o campo
               If Len(.Text) Then
                  .Text = 0
               End If
            End With
            If tubo_conexoes_defofo.col = 4 Then
                houveDigitacaoEsgoto = True
                auxSelChange = True
                rowAux = tubo_conexoes_defofo.row
                tubo_conexoes_defofo_SelChange
            End If
            auxChange = True
        End If
    End If
End Sub

Private Sub tubo_conexoes_pba_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If tubo_conexoes_pba.col = 4 Or tubo_conexoes_pba.col = 6 Then
            With tubo_conexoes_pba
            'Zera o campo
               If Len(.Text) Then
                  .Text = 0
               End If
            End With
            If tubo_conexoes_pba.col = 4 Then
                houveDigitacaoTubosRoscaveis = True
                auxSelChangeRoscaveis = True
                rowAuxRoscaveis = tubo_conexoes_pba.row
                tubo_conexoes_pba_SelChange
            End If
            auxChangeRoscaveis = True
        End If
    End If
End Sub

Private Sub tubo_tubos_conexoes_coletor_esgoto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If tubo_tubos_conexoes_coletor_esgoto.col = 4 Or tubo_tubos_conexoes_coletor_esgoto.col = 6 Then
            With tubo_tubos_conexoes_coletor_esgoto
            'Zera o campo
               If Len(.Text) Then
                  .Text = 0
               End If
            End With
            If tubo_tubos_conexoes_coletor_esgoto.col = 4 Then
                houveDigitacaoSoldaveisAgua = True
                auxSelChangeSaldaveisAgua = True
                rowAuxSaldaveisAgua = tubo_tubos_conexoes_coletor_esgoto.row
                tubo_tubos_conexoes_coletor_esgoto_SelChange
            End If
            auxChangeSaldaveisAgua = True
        End If
    End If
End Sub

Private Sub tubos_conexoes_agua_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If tubos_conexoes_agua.col = 4 Or tubos_conexoes_agua.col = 6 Then
            With tubos_conexoes_agua
            'Zera o campo
               If Len(.Text) Then
                  .Text = 0
               End If
            End With
            If tubos_conexoes_agua.col = 4 Then
                houveDigitacaoTuboConexoesAgua = True
                auxSelChangeTuboConexoesAgua = True
                rowAuxTuboConexoesAgua = tubos_conexoes_agua.row
                tubos_conexoes_agua_SelChange
            End If
            auxChangeTuboConexoesAgua = True
        End If
    End If
End Sub

Private Sub tubos_conexoes_irri_azuis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If tubos_conexoes_irri_azuis.col = 4 Or tubos_conexoes_irri_azuis.col = 6 Then
            With tubos_conexoes_irri_azuis
            'Zera o campo
               If Len(.Text) Then
                  .Text = 0
               End If
            End With
            If tubos_conexoes_irri_azuis.col = 4 Then
                houveDigitacaoTubosConexoesIrriAzuis = True
                auxSelChangeTubosConexoesIrriAzuis = True
                rowAuxTubosConexoesIrriAzuis = tubos_conexoes_irri_azuis.row
                tubos_conexoes_irri_azuis_SelChange
            End If
            auxChangeTubosConexoesIrriAzuis = True
        End If
    End If
End Sub

Private Sub tubos_conexoes_predial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If tubos_conexoes_predial.col = 4 Or tubos_conexoes_predial.col = 6 Then
            With tubos_conexoes_predial
            'Zera o campo
               If Len(.Text) Then
                  .Text = 0
               End If
            End With
            If tubos_conexoes_predial.col = 4 Then
                houveDigitacaoTuboConexoesPredial = True
                auxSelChangeTuboConexoesPredial = True
                rowAuxTuboConexoesPredial = tubos_conexoes_predial.row
                tubos_conexoes_predial_SelChange
            End If
            auxChangeTuboConexoesPredial = True
        End If
    End If
End Sub

Private Sub tubos_conexoes_roscaveis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If tubos_conexoes_roscaveis.col = 4 Or tubos_conexoes_roscaveis.col = 6 Then
            With tubos_conexoes_roscaveis
            'Zera o campo
               If Len(.Text) Then
                  .Text = 0
               End If
            End With
            If tubos_conexoes_roscaveis.col = 4 Then
                houveDigitacaoTubosConexoesRoscaveis = True
                auxSelChangeTubosConexoesRoscaveis = True
                rowAuxTubosConexoesRoscaveis = tubos_conexoes_roscaveis.row
                tubos_conexoes_roscaveis_SelChange
            End If
            auxChangeTubosConexoesRoscaveis = True
        End If
    End If
End Sub

Private Sub GrdNotaCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        Dim col As Integer
        Dim row As Integer
        Dim sgQuery As String
        col = GrdNotaCliente.col
        row = GrdNotaCliente.row
        
        sgQuery = MsgBox("Deseja deletar o produto Código: " & Trim(GrdNotaCliente.TextMatrix(row, col)) & " deste pedido ?", vbQuestion + vbYesNo + vbDefaultButton2, "Atenção!")
        
        If sgQuery = vbYes Or sgQuery = vbOK Then
            LimpaRegistroGridAuxiliar GrdNotaCliente.TextMatrix(row, 0)
            If GrdNotaCliente.rows = 2 Then
                GrdNotaCliente.rows = GrdNotaCliente.rows - 1
            Else
                GrdNotaCliente.RemoveItem row
            End If
            
            CalculaIndice
            Exit Sub
        End If
        
        If GrdNotaCliente.col = 3 Or GrdNotaCliente.col = 5 Then
            With GrdNotaCliente
            'Zera o campo
               If Len(.Text) Then
                  .Text = 0
               End If
            End With
            If GrdNotaCliente.col = 4 Then
                houveDigitacaoGrdNotaCliente = True
                auxSelChangeGrdNotaCliente = True
                rowAuxGrdNotaCliente = GrdNotaCliente.row
                GrdNotaCliente_SelChange
            End If
            auxChangeGrdNotaCliente = True
        End If
    End If
End Sub

Private Sub tubos_conexoes_agua_Click()
    
    If tubos_conexoes_agua.SelectionMode = flexSelectionFree Then
        If tubos_conexoes_agua.col = 1 Then
            tubos_conexoes_agua.row = tubos_conexoes_agua.row
            tubos_conexoes_agua.col = 4
        Else
            tubos_conexoes_agua.row = tubos_conexoes_agua.row
            tubos_conexoes_agua.col = tubos_conexoes_agua.col
        End If
    End If
    ControleLostFocus = False
End Sub

Private Sub tubos_conexoes_agua_DblClick()
    
    If tubos_conexoes_agua.SelectionMode = flexSelectionByRow Then
        tubos_conexoes_agua.SelectionMode = flexSelectionFree
        tubos_conexoes_agua.Refresh
        tubos_conexoes_agua_Click
        'tubos_conexoes_irri_azuis_SelChange
        Exit Sub
    End If
    
End Sub

Private Sub tubos_conexoes_agua_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim col As Integer
    Dim row As Integer
    
    Select Case KeyAscii
    
    Case vbKeyReturn, vbKeyTab
    'move para a proxima celula.
    
    With tubos_conexoes_agua
    
      If .col + 1 <= .Cols - 1 Then
         .col = .col + 1
      Else
         If .row + 1 <= .rows - 1 Then
             .row = .row + 1
             .col = 0
         Else
             .row = 1
             .col = 0
         End If
      End If
    End With
    
    Case Is < 32
        
    Case Else
        
        If tubos_conexoes_agua.col = 6 Then
            If auxChangeTuboConexoesAgua = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_agua
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "," Then
                    With tubos_conexoes_agua
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "." Then
                    With tubos_conexoes_agua
                       .Text = .Text & ","
                    End With
                End If
            End If
            If auxChangeTuboConexoesAgua = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_agua
                       .Text = Chr(KeyAscii)
                    End With
                End If
                auxChangeTuboConexoesAgua = False
                auxSelChangeTuboConexoesAgua = True
                rowAuxTuboConexoesAgua = tubos_conexoes_agua.row
            End If
        End If
        If tubos_conexoes_agua.col = 4 Then
            If auxChangeTuboConexoesAgua = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_agua
                       .Text = .Text & Chr(KeyAscii)
                    End With
                    houveDigitacaoTuboConexoesAgua = True
                    auxSelChangeTuboConexoesAgua = True
                    rowAuxTuboConexoesAgua = tubos_conexoes_agua.row
                End If
            End If
            If auxChangeTuboConexoesAgua = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_agua
                       .Text = Chr(KeyAscii)
                    End With
                    houveDigitacaoTuboConexoesAgua = True
                    auxSelChangeTuboConexoesAgua = True
                    rowAuxTuboConexoesAgua = tubos_conexoes_agua.row
                End If
                auxChangeTuboConexoesAgua = False
            End If
        End If
    End Select
    
    If KeyAscii = 8 Then
        If tubos_conexoes_agua.col = 4 Or tubos_conexoes_agua.col = 6 Then
             With tubos_conexoes_agua
            'remove o ultimo caractere
               If Len(.Text) Then
                  .Text = Left(.Text, Len(.Text) - 1)
               End If
            End With
            row = tubos_conexoes_agua.row
            col = tubos_conexoes_agua.col
            If tubos_conexoes_agua.TextMatrix(row, col) = "" Then
                tubos_conexoes_agua.TextMatrix(row, col) = 0
                If col = 4 Then
                    houveDigitacaoTuboConexoesAgua = True
                    auxSelChangeTuboConexoesAgua = True
                    rowAuxTuboConexoesAgua = tubos_conexoes_agua.row
                    tubos_conexoes_agua_SelChange
                End If
                auxChangeTuboConexoesAgua = True
            End If
        End If
    End If
    
    If KeyAscii = 13 Then
        titleTab = tab_simulacao_pedido.Caption
        auxChangeTuboConexoesAgua = True
        AtualizaValorComDesconto tubos_conexoes_agua.row, tubos_conexoes_agua
        If houveDigitacaoTuboConexoesAgua = True Or tubos_conexoes_agua.TextMatrix(rowAuxTuboConexoesAgua, 4) > 0 Then
            PopularVariaveisCalculo tubos_conexoes_agua.row, tubos_conexoes_agua
            CalculaTotal tubos_conexoes_agua.row, tubos_conexoes_agua
            carregaResumo tubos_conexoes_agua.row, tubos_conexoes_agua
            houveDigitacaoTuboConexoesAgua = False
        End If
    End If
    
End Sub

Private Sub GrdNotaCliente_Click()
    
    If GrdNotaCliente.SelectionMode = flexSelectionFree Then
        If GrdNotaCliente.col = 1 Then
            GrdNotaCliente.row = GrdNotaCliente.row
            GrdNotaCliente.col = 3
        Else
            GrdNotaCliente.row = GrdNotaCliente.row
            GrdNotaCliente.col = GrdNotaCliente.col
        End If
    End If
    ControleLostFocus = False
    
End Sub

Private Sub tubos_conexoes_predial_Click()
    
    If tubos_conexoes_predial.SelectionMode = flexSelectionFree Then
        If tubos_conexoes_predial.col = 1 Then
            tubos_conexoes_predial.row = tubos_conexoes_predial.row
            tubos_conexoes_predial.col = 4
        Else
            tubos_conexoes_predial.row = tubos_conexoes_predial.row
            tubos_conexoes_predial.col = tubos_conexoes_predial.col
        End If
    End If
    ControleLostFocus = False
End Sub

Private Sub tubos_conexoes_predial_DblClick()
    
    If tubos_conexoes_predial.SelectionMode = flexSelectionByRow Then
        tubos_conexoes_predial.SelectionMode = flexSelectionFree
        tubos_conexoes_predial.Refresh
        tubos_conexoes_predial_Click
        'tubos_conexoes_irri_azuis_SelChange
        Exit Sub
    End If
    
End Sub

Private Sub tubos_conexoes_predial_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim col As Integer
    Dim row As Integer
    
    Select Case KeyAscii
    
    Case vbKeyReturn, vbKeyTab
    'move para a proxima celula.
    
    With tubos_conexoes_predial
    
      If .col + 1 <= .Cols - 1 Then
         .col = .col + 1
      Else
         If .row + 1 <= .rows - 1 Then
             .row = .row + 1
             .col = 0
         Else
             .row = 1
             .col = 0
         End If
      End If
    End With
    
    Case Is < 32
        
    Case Else
        
        If tubos_conexoes_predial.col = 6 Then
            If auxChangeTuboConexoesPredial = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_predial
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "," Then
                    With tubos_conexoes_predial
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "." Then
                    With tubos_conexoes_predial
                       .Text = .Text & ","
                    End With
                End If
            End If
            If auxChangeTuboConexoesPredial = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_predial
                       .Text = Chr(KeyAscii)
                    End With
                End If
                auxChangeTuboConexoesPredial = False
                auxSelChangeTuboConexoesPredial = True
                rowAuxTuboConexoesPredial = tubos_conexoes_predial.row
            End If
        End If
        If tubos_conexoes_predial.col = 4 Then
            If auxChangeTuboConexoesPredial = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_predial
                       .Text = .Text & Chr(KeyAscii)
                    End With
                    houveDigitacaoTuboConexoesPredial = True
                    auxSelChangeTuboConexoesPredial = True
                    rowAuxTuboConexoesPredial = tubos_conexoes_predial.row
                End If
            End If
            If auxChangeTuboConexoesPredial = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_predial
                       .Text = Chr(KeyAscii)
                    End With
                    houveDigitacaoTuboConexoesPredial = True
                    auxSelChangeTuboConexoesPredial = True
                    rowAuxTuboConexoesPredial = tubos_conexoes_predial.row
                End If
                auxChangeTuboConexoesPredial = False
            End If
        End If
    End Select
    
    If KeyAscii = 8 Then
        If tubos_conexoes_predial.col = 4 Or tubos_conexoes_predial.col = 6 Then
            With tubos_conexoes_predial
            'remove o ultimo caractere
               If Len(.Text) Then
                  .Text = Left(.Text, Len(.Text) - 1)
               End If
            End With
            row = tubos_conexoes_predial.row
            col = tubos_conexoes_predial.col
            If tubos_conexoes_predial.TextMatrix(row, col) = "" Then
                tubos_conexoes_predial.TextMatrix(row, col) = 0
                If col = 4 Then
                    houveDigitacaoTuboConexoesPredial = True
                    auxSelChangeTuboConexoesPredial = True
                    rowAuxTuboConexoesPredial = tubos_conexoes_predial.row
                    tubos_conexoes_predial_SelChange
                End If
                auxChangeTuboConexoesPredial = True
            End If
        End If
    End If
    
    If KeyAscii = 13 Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto tubos_conexoes_predial.row, tubos_conexoes_predial
        auxChangeTuboConexoesPredial = True
        If houveDigitacaoTuboConexoesPredial = True Or tubos_conexoes_predial.TextMatrix(rowAuxTuboConexoesPredial, 4) > 0 Then
            PopularVariaveisCalculo tubos_conexoes_predial.row, tubos_conexoes_predial
            CalculaTotal tubos_conexoes_predial.row, tubos_conexoes_predial
            carregaResumo tubos_conexoes_predial.row, tubos_conexoes_predial
            houveDigitacaoTuboConexoesPredial = False
        End If
    End If
    
End Sub

Private Sub tubos_conexoes_roscaveis_Click()
    
    If tubos_conexoes_roscaveis.SelectionMode = flexSelectionFree Then
        If tubos_conexoes_roscaveis.col = 1 Then
            tubos_conexoes_roscaveis.row = tubos_conexoes_roscaveis.row
            tubos_conexoes_roscaveis.col = 4
        Else
            tubos_conexoes_roscaveis.row = tubos_conexoes_roscaveis.row
            tubos_conexoes_roscaveis.col = tubos_conexoes_roscaveis.col
        End If
    End If
    ControleLostFocus = False
End Sub

Private Sub tubos_conexoes_roscaveis_DblClick()
    
    If tubos_conexoes_roscaveis.SelectionMode = flexSelectionByRow Then
        tubos_conexoes_roscaveis.SelectionMode = flexSelectionFree
        tubos_conexoes_roscaveis.Refresh
        tubos_conexoes_roscaveis_Click
        'tubos_conexoes_irri_azuis_SelChange
        Exit Sub
    End If
    
End Sub

Private Sub tubos_conexoes_roscaveis_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim col As Integer
    Dim row As Integer
    
    Select Case KeyAscii
    
    Case vbKeyReturn, vbKeyTab
    'move para a proxima celula.
    
    With tubos_conexoes_roscaveis
    
      If .col + 1 <= .Cols - 1 Then
         .col = .col + 1
      Else
         If .row + 1 <= .rows - 1 Then
             .row = .row + 1
             .col = 0
         Else
             .row = 1
             .col = 0
         End If
      End If
    End With
    
    Case Is < 32
        
    Case Else
        
        If tubos_conexoes_roscaveis.col = 6 Then
            If auxChangeTubosConexoesRoscaveis = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_roscaveis
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "," Then
                    With tubos_conexoes_roscaveis
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "." Then
                    With tubos_conexoes_roscaveis
                       .Text = .Text & ","
                    End With
                End If
            End If
            If auxChangeTubosConexoesRoscaveis = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_roscaveis
                       .Text = Chr(KeyAscii)
                    End With
                End If
                auxChangeTubosConexoesRoscaveis = False
                auxSelChangeTubosConexoesRoscaveis = True
                rowAuxTubosConexoesRoscaveis = tubos_conexoes_roscaveis.row
            End If
        End If
        If tubos_conexoes_roscaveis.col = 4 Then
            If auxChangeTubosConexoesRoscaveis = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_roscaveis
                       .Text = .Text & Chr(KeyAscii)
                    End With
                    houveDigitacaoTubosConexoesRoscaveis = True
                    auxSelChangeTubosConexoesRoscaveis = True
                    rowAuxTubosConexoesRoscaveis = tubos_conexoes_roscaveis.row
                End If
            End If
            If auxChangeTubosConexoesRoscaveis = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_roscaveis
                       .Text = Chr(KeyAscii)
                    End With
                    houveDigitacaoTubosConexoesRoscaveis = True
                    auxSelChangeTubosConexoesRoscaveis = True
                    rowAuxTubosConexoesRoscaveis = tubos_conexoes_roscaveis.row
                End If
                auxChangeTubosConexoesRoscaveis = False
            End If
        End If
    End Select
    
    If KeyAscii = 8 Then
        If tubos_conexoes_roscaveis.col = 4 Or tubos_conexoes_roscaveis.col = 6 Then
            With tubos_conexoes_roscaveis
            'remove o ultimo caractere
               If Len(.Text) Then
                  .Text = Left(.Text, Len(.Text) - 1)
               End If
            End With
            row = tubos_conexoes_roscaveis.row
            col = tubos_conexoes_roscaveis.col
            If tubos_conexoes_roscaveis.TextMatrix(row, col) = "" Then
                tubos_conexoes_roscaveis.TextMatrix(row, col) = 0
                If col = 4 Then
                    houveDigitacaoTubosConexoesRoscaveis = True
                    auxSelChangeTubosConexoesRoscaveis = True
                    rowAuxTubosConexoesRoscaveis = tubos_conexoes_roscaveis.row
                    tubos_conexoes_roscaveis_SelChange
                End If
                auxChangeTubosConexoesRoscaveis = True
            End If
        End If
    End If
    
    If KeyAscii = 13 Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto tubos_conexoes_roscaveis.row, tubos_conexoes_roscaveis
        auxChangeTubosConexoesRoscaveis = True
        If houveDigitacaoTubosConexoesRoscaveis = True Or tubos_conexoes_roscaveis.TextMatrix(rowAuxTubosConexoesRoscaveis, 4) > 0 Then
            PopularVariaveisCalculo tubos_conexoes_roscaveis.row, tubos_conexoes_roscaveis
            CalculaTotal tubos_conexoes_roscaveis.row, tubos_conexoes_roscaveis
            carregaResumo tubos_conexoes_roscaveis.row, tubos_conexoes_roscaveis
            houveDigitacaoTubosConexoesRoscaveis = False
        End If
    End If
    
End Sub

Private Sub tab_simulacao_pedido_Click(PreviousTab As Integer)
    'If titleTab <> "" Then
        'If tab_simulacao_pedido.Caption <> titleTab Then
            'tab_simulacao_pedido.Tab = PreviousTab
            'titleTab = ""
        'End If
    'End If
    
    Select Case PreviousTab
        Case 0
            tubos_conexoes_predial_SelChange
        Case 1
            tubos_conexoes_agua_SelChange
        Case 2
            tubos_conexoes_roscaveis_SelChange
        Case 3
            tubos_conexoes_irri_azuis_SelChange
        Case 4
            tubo_tubos_conexoes_coletor_esgoto_SelChange
        Case 5
            tubo_conexoes_pba_SelChange
        Case 6
            tubo_conexoes_defofo_SelChange
        Case 7
            GrdNotaCliente_SelChange
    End Select
    
End Sub

Private Sub tubos_conexoes_irri_azuis_Click()
    
    If tubos_conexoes_irri_azuis.SelectionMode = flexSelectionFree Then
        If tubos_conexoes_irri_azuis.col = 1 Then
            tubos_conexoes_irri_azuis.row = tubos_conexoes_irri_azuis.row
            tubos_conexoes_irri_azuis.col = 4
        Else
            tubos_conexoes_irri_azuis.row = tubos_conexoes_irri_azuis.row
            tubos_conexoes_irri_azuis.col = tubos_conexoes_irri_azuis.col
        End If
    End If
    ControleLostFocus = False
End Sub

Private Sub tubos_conexoes_irri_azuis_DblClick()
    
    If tubos_conexoes_irri_azuis.SelectionMode = flexSelectionByRow Then
        tubos_conexoes_irri_azuis.SelectionMode = flexSelectionFree
        tubos_conexoes_irri_azuis.Refresh
        tubos_conexoes_irri_azuis_Click
        'tubos_conexoes_irri_azuis_SelChange
        Exit Sub
    End If
    
End Sub

Private Sub GrdNotaCliente_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim col As Integer
    Dim row As Integer
    Dim dlDesc   As Double
    
    
    Select Case KeyAscii
    
    Case vbKeyReturn, vbKeyTab
    'move para a proxima celula.
    
    With GrdNotaCliente
    
      If .col + 1 <= .Cols - 1 Then
         .col = .col + 1
      Else
         If .row + 1 <= .rows - 1 Then
             .row = .row + 1
             .col = 0
         Else
             .row = 1
             .col = 0
         End If
      End If
    End With
    
    Case Is < 32
        
    Case Else
        
        If GrdNotaCliente.col = 6 Then
            'GrdNotaCliente.SelectionMode = flexSelectionFree
            If auxChangeGrdNotaCliente = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With GrdNotaCliente
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "," Then
                    With GrdNotaCliente
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "." Then
                    With GrdNotaCliente
                       .Text = .Text & ","
                    End With
                End If
            End If
            If auxChangeGrdNotaCliente = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With GrdNotaCliente
                       .Text = Chr(KeyAscii)
                    End With
                End If
                auxChangeGrdNotaCliente = False
                auxSelChangeGrdNotaCliente = True
                rowAuxGrdNotaCliente = GrdNotaCliente.row
            End If
        End If
        If GrdNotaCliente.col = 3 Then
            'GrdNotaCliente.SelectionMode = flexSelectionFree
            If auxChangeGrdNotaCliente = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With GrdNotaCliente
                       .Text = .Text & Chr(KeyAscii)
                    End With
                    houveDigitacaoGrdNotaCliente = True
                    auxSelChangeGrdNotaCliente = True
                    rowAuxGrdNotaCliente = GrdNotaCliente.row
                End If
            End If
            If auxChangeGrdNotaCliente = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With GrdNotaCliente
                       .Text = Chr(KeyAscii)
                    End With
                    houveDigitacaoGrdNotaCliente = True
                    auxSelChangeGrdNotaCliente = True
                    rowAuxGrdNotaCliente = GrdNotaCliente.row
                End If
                auxChangeGrdNotaCliente = False
            End If
        End If
    End Select
    
    If KeyAscii = 8 Then
        If GrdNotaCliente.col = 3 Or GrdNotaCliente.col = 6 Then
             With GrdNotaCliente
            'remove o ultimo caractere
               If Len(.Text) Then
                  .Text = Left(.Text, Len(.Text) - 1)
               End If
            End With
            row = GrdNotaCliente.row
            col = GrdNotaCliente.col
            If GrdNotaCliente.TextMatrix(row, col) = "" Then
                GrdNotaCliente.TextMatrix(row, col) = 0
                auxChangeGrdNotaCliente = True
            End If
        End If
    End If
    
    If KeyAscii = 13 Then
        titleTab = tab_simulacao_pedido.Caption
        'AtualizaValorComDescontoResumo GrdNotaCliente.row, GrdNotaCliente
        auxChangeGrdNotaCliente = True
        If houveDigitacaoGrdNotaCliente = True Or GrdNotaCliente.TextMatrix(rowAuxGrdNotaCliente, 3) > 0 Then
            PopularVariaveisCalculoResumo GrdNotaCliente.row, GrdNotaCliente
            'CalculaTotalResumo GrdNotaCliente.row, GrdNotaCliente
            'CalculaIndice
            AtualizaGridAuxiliar rowAuxGrdNotaCliente
            DefineCorResumo GrdNotaCliente.row, GrdNotaCliente
            
            dlDesc = Format(Trim(MskDatEmiNf), "##0.00")
    
            If dlDesc > dlSumDscItem Then
            
                
                MsgBox "Desconto informado maior que o permitido para esta venda.", vbExclamation + vbOKOnly, "Atenção!"
                
                GrdNotaCliente.TextMatrix(rowAuxGrdNotaCliente, 6) = dlSumDscItem
                
                MskDatEmiNf = dlSumDscItem
                
                AtualizaGridAuxiliar rowAuxGrdNotaCliente
                'AtualizaValorComDescontoResumo rowAuxGrdNotaCliente, GrdNotaCliente
                
                'CalculaTotalResumo rowAuxGrdNotaCliente, GrdNotaCliente
                
                tab_simulacao_pedido.SetFocus
                
            End If
            
            'Bto_Aplica_resumo_Click
            'carregaResumo GrdNotaCliente.row, GrdNotaCliente
            houveDigitacaoGrdNotaCliente = False
        End If
    End If
    
End Sub

Private Sub tubos_conexoes_irri_azuis_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim col As Integer
    Dim row As Integer
    
    
    Select Case KeyAscii
    
    Case vbKeyReturn, vbKeyTab
    'move para a proxima celula.
    
    With tubos_conexoes_irri_azuis
    
      If .col + 1 <= .Cols - 1 Then
         .col = .col + 1
      Else
         If .row + 1 <= .rows - 1 Then
             .row = .row + 1
             .col = 0
         Else
             .row = 1
             .col = 0
         End If
      End If
    End With
    
    Case Is < 32
        
    Case Else
        
        If tubos_conexoes_irri_azuis.col = 6 Then
            tubos_conexoes_irri_azuis.SelectionMode = flexSelectionFree
            If auxChangeTubosConexoesIrriAzuis = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_irri_azuis
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "," Then
                    With tubos_conexoes_irri_azuis
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "." Then
                    With tubos_conexoes_irri_azuis
                       .Text = .Text & ","
                    End With
                End If
            End If
            If auxChangeTubosConexoesIrriAzuis = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_irri_azuis
                       .Text = Chr(KeyAscii)
                    End With
                End If
                auxChangeTubosConexoesIrriAzuis = False
                auxSelChangeTubosConexoesIrriAzuis = True
                rowAuxTubosConexoesIrriAzuis = tubos_conexoes_irri_azuis.row
            End If
        End If
        If tubos_conexoes_irri_azuis.col = 4 Then
            tubos_conexoes_irri_azuis.SelectionMode = flexSelectionFree
            If auxChangeTubosConexoesIrriAzuis = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_irri_azuis
                       .Text = .Text & Chr(KeyAscii)
                    End With
                    houveDigitacaoTubosConexoesIrriAzuis = True
                    auxSelChangeTubosConexoesIrriAzuis = True
                    rowAuxTubosConexoesIrriAzuis = tubos_conexoes_irri_azuis.row
                End If
            End If
            If auxChangeTubosConexoesIrriAzuis = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubos_conexoes_irri_azuis
                       .Text = Chr(KeyAscii)
                    End With
                    houveDigitacaoTubosConexoesIrriAzuis = True
                    auxSelChangeTubosConexoesIrriAzuis = True
                    rowAuxTubosConexoesIrriAzuis = tubos_conexoes_irri_azuis.row
                End If
                auxChangeTubosConexoesIrriAzuis = False
            End If
        End If
    End Select
    
    If KeyAscii = 8 Then
        If tubos_conexoes_irri_azuis.col = 4 Or tubos_conexoes_irri_azuis.col = 6 Then
            With tubos_conexoes_irri_azuis
            'remove o ultimo caractere
               If Len(.Text) Then
                  .Text = Left(.Text, Len(.Text) - 1)
               End If
            End With
            row = tubos_conexoes_irri_azuis.row
            col = tubos_conexoes_irri_azuis.col
            If tubos_conexoes_irri_azuis.TextMatrix(row, col) = "" Then
                tubos_conexoes_irri_azuis.TextMatrix(row, col) = 0
                If col = 4 Then
                    houveDigitacaoTubosConexoesIrriAzuis = True
                    auxSelChangeTubosConexoesIrriAzuis = True
                    rowAuxTubosConexoesIrriAzuis = tubos_conexoes_irri_azuis.row
                    tubos_conexoes_irri_azuis_SelChange
                End If
                auxChangeTubosConexoesIrriAzuis = True
            End If
        End If
    End If
    
    If KeyAscii = 13 Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto tubos_conexoes_irri_azuis.row, tubos_conexoes_irri_azuis
        auxChangeTubosConexoesIrriAzuis = True
        If houveDigitacaoTubosConexoesIrriAzuis = True Or tubos_conexoes_irri_azuis.TextMatrix(rowAuxTubosConexoesIrriAzuis, 4) > 0 Then
            PopularVariaveisCalculo tubos_conexoes_irri_azuis.row, tubos_conexoes_irri_azuis
            CalculaTotal tubos_conexoes_irri_azuis.row, tubos_conexoes_irri_azuis
            carregaResumo tubos_conexoes_irri_azuis.row, tubos_conexoes_irri_azuis
            houveDigitacaoTubosConexoesIrriAzuis = False
        End If
    End If
    
End Sub

Private Sub tubo_conexoes_pba_Click()
    
    If tubo_conexoes_pba.SelectionMode = flexSelectionFree Then
        If tubo_conexoes_pba.col = 1 Then
            tubo_conexoes_pba.row = tubo_conexoes_pba.row
            tubo_conexoes_pba.col = 4
        Else
            tubo_conexoes_pba.row = tubo_conexoes_pba.row
            tubo_conexoes_pba.col = tubo_conexoes_pba.col
        End If
    End If
    ControleLostFocus = False
End Sub

Private Sub tubo_conexoes_pba_DblClick()
    
    If tubo_conexoes_pba.SelectionMode = flexSelectionByRow Then
        tubo_conexoes_pba.SelectionMode = flexSelectionFree
        tubo_conexoes_pba.Refresh
        tubo_conexoes_pba_Click
        'tubos_conexoes_irri_azuis_SelChange
        Exit Sub
    End If
    
End Sub

Private Sub tubo_conexoes_pba_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim col As Integer
    Dim row As Integer
    
    Select Case KeyAscii
    
    Case vbKeyReturn, vbKeyTab
    'move para a proxima celula.
    
    With tubo_conexoes_pba
    
      If .col + 1 <= .Cols - 1 Then
         .col = .col + 1
      Else
         If .row + 1 <= .rows - 1 Then
             .row = .row + 1
             .col = 0
         Else
             .row = 1
             .col = 0
         End If
      End If
    End With
    
    Case Is < 32
        
    Case Else
        
        If tubo_conexoes_pba.col = 6 Then
            If auxChangeRoscaveis = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubo_conexoes_pba
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "," Then
                    With tubo_conexoes_pba
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "." Then
                    With tubo_conexoes_pba
                       .Text = .Text & ","
                    End With
                End If
            End If
            If auxChangeRoscaveis = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubo_conexoes_pba
                       .Text = Chr(KeyAscii)
                    End With
                End If
                auxChangeRoscaveis = False
                auxSelChangeRoscaveis = True
                rowAuxRoscaveis = tubo_conexoes_pba.row
            End If
        End If
        If tubo_conexoes_pba.col = 4 Then
            If auxChangeRoscaveis = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubo_conexoes_pba
                       .Text = .Text & Chr(KeyAscii)
                    End With
                    houveDigitacaoTubosRoscaveis = True
                    auxSelChangeRoscaveis = True
                    rowAuxRoscaveis = tubo_conexoes_pba.row
                End If
            End If
            If auxChangeRoscaveis = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubo_conexoes_pba
                       .Text = Chr(KeyAscii)
                    End With
                    houveDigitacaoTubosRoscaveis = True
                    auxSelChangeRoscaveis = True
                    rowAuxRoscaveis = tubo_conexoes_pba.row
                End If
                auxChangeRoscaveis = False
            End If
        End If
    End Select
    
    If KeyAscii = 8 Then
        If tubo_conexoes_pba.col = 4 Or tubo_conexoes_pba.col = 6 Then
            With tubo_conexoes_pba
            'remove o ultimo caractere
               If Len(.Text) Then
                  .Text = Left(.Text, Len(.Text) - 1)
               End If
            End With
            row = tubo_conexoes_pba.row
            col = tubo_conexoes_pba.col
            If tubo_conexoes_pba.TextMatrix(row, col) = "" Then
                tubo_conexoes_pba.TextMatrix(row, col) = 0
                If col = 4 Then
                    houveDigitacaoTubosRoscaveis = True
                    auxSelChangeRoscaveis = True
                    rowAuxRoscaveis = tubo_conexoes_pba.row
                    tubo_conexoes_pba_SelChange
                End If
                auxChangeRoscaveis = True
            End If
        End If
    End If
    
    If KeyAscii = 13 Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto tubo_conexoes_pba.row, tubo_conexoes_pba
        auxChangeRoscaveis = True
        If houveDigitacaoTubosRoscaveis = True Or tubo_conexoes_pba.TextMatrix(rowAuxRoscaveis, 4) > 0 Then
            PopularVariaveisCalculo tubo_conexoes_pba.row, tubo_conexoes_pba
            CalculaTotal tubo_conexoes_pba.row, tubo_conexoes_pba
            carregaResumo tubo_conexoes_pba.row, tubo_conexoes_pba
            houveDigitacaoTubosRoscaveis = False
        End If
    End If
    
End Sub

Private Sub tubo_tubos_conexoes_coletor_esgoto_Click()
    
    If tubo_tubos_conexoes_coletor_esgoto.SelectionMode = flexSelectionFree Then
        If tubo_tubos_conexoes_coletor_esgoto.col = 1 Then
            tubo_tubos_conexoes_coletor_esgoto.row = tubo_tubos_conexoes_coletor_esgoto.row
            tubo_tubos_conexoes_coletor_esgoto.col = 4
        Else
            tubo_tubos_conexoes_coletor_esgoto.row = tubo_tubos_conexoes_coletor_esgoto.row
            tubo_tubos_conexoes_coletor_esgoto.col = tubo_tubos_conexoes_coletor_esgoto.col
        End If
    End If
    ControleLostFocus = False
End Sub

Private Sub tubo_tubos_conexoes_coletor_esgoto_DblClick()
    
    If tubo_tubos_conexoes_coletor_esgoto.SelectionMode = flexSelectionByRow Then
        tubo_tubos_conexoes_coletor_esgoto.SelectionMode = flexSelectionFree
        tubo_tubos_conexoes_coletor_esgoto.Refresh
        tubo_tubos_conexoes_coletor_esgoto_Click
        'tubos_conexoes_irri_azuis_SelChange
        Exit Sub
    End If
    
End Sub

Private Sub tubo_tubos_conexoes_coletor_esgoto_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim col As Integer
    Dim row As Integer
    
    Select Case KeyAscii
    
    Case vbKeyReturn, vbKeyTab
    'move para a proxima celula.
    
    With tubo_tubos_conexoes_coletor_esgoto
    
      If .col + 1 <= .Cols - 1 Then
         .col = .col + 1
      Else
         If .row + 1 <= .rows - 1 Then
             .row = .row + 1
             .col = 0
         Else
             .row = 1
             .col = 0
         End If
      End If
    End With
    
    Case Is < 32
        
    Case Else
        
        If tubo_tubos_conexoes_coletor_esgoto.col = 6 Then
            If auxChangeSaldaveisAgua = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubo_tubos_conexoes_coletor_esgoto
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "," Then
                    With tubo_tubos_conexoes_coletor_esgoto
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "." Then
                    With tubo_tubos_conexoes_coletor_esgoto
                       .Text = .Text & ","
                    End With
                End If
            End If
            If auxChangeSaldaveisAgua = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubo_tubos_conexoes_coletor_esgoto
                       .Text = Chr(KeyAscii)
                    End With
                End If
                auxChangeSaldaveisAgua = False
                auxSelChangeSaldaveisAgua = True
                rowAuxSaldaveisAgua = tubo_tubos_conexoes_coletor_esgoto.row
            End If
        End If
        If tubo_tubos_conexoes_coletor_esgoto.col = 4 Then
            If auxChangeSaldaveisAgua = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubo_tubos_conexoes_coletor_esgoto
                       .Text = .Text & Chr(KeyAscii)
                    End With
                    houveDigitacaoSoldaveisAgua = True
                    auxSelChangeSaldaveisAgua = True
                    rowAuxSaldaveisAgua = tubo_tubos_conexoes_coletor_esgoto.row
                End If
            End If
            If auxChangeSaldaveisAgua = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubo_tubos_conexoes_coletor_esgoto
                       .Text = Chr(KeyAscii)
                    End With
                    houveDigitacaoSoldaveisAgua = True
                    auxSelChangeSaldaveisAgua = True
                    rowAuxSaldaveisAgua = tubo_tubos_conexoes_coletor_esgoto.row
                End If
                auxChangeSaldaveisAgua = False
            End If
        End If
    End Select
    
    If KeyAscii = 8 Then
        If tubo_tubos_conexoes_coletor_esgoto.col = 4 Or tubo_tubos_conexoes_coletor_esgoto.col = 6 Then
            With tubo_tubos_conexoes_coletor_esgoto
            'remove o ultimo caractere
               If Len(.Text) Then
                  .Text = Left(.Text, Len(.Text) - 1)
               End If
            End With
            row = tubo_tubos_conexoes_coletor_esgoto.row
            col = tubo_tubos_conexoes_coletor_esgoto.col
            If tubo_tubos_conexoes_coletor_esgoto.TextMatrix(row, col) = "" Then
                tubo_tubos_conexoes_coletor_esgoto.TextMatrix(row, col) = 0
                If col = 4 Then
                    houveDigitacaoSoldaveisAgua = True
                    auxSelChangeSaldaveisAgua = True
                    rowAuxSaldaveisAgua = tubo_tubos_conexoes_coletor_esgoto.row
                    tubo_tubos_conexoes_coletor_esgoto_SelChange
                End If
                auxChangeSaldaveisAgua = True
            End If
        End If
    End If
    
    If KeyAscii = 13 Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto tubo_tubos_conexoes_coletor_esgoto.row, tubo_tubos_conexoes_coletor_esgoto
        auxChangeSaldaveisAgua = True
        If houveDigitacaoSoldaveisAgua = True Or tubo_tubos_conexoes_coletor_esgoto.TextMatrix(rowAuxSaldaveisAgua, 4) > 0 Then
            PopularVariaveisCalculo tubo_tubos_conexoes_coletor_esgoto.row, tubo_tubos_conexoes_coletor_esgoto
            CalculaTotal tubo_tubos_conexoes_coletor_esgoto.row, tubo_tubos_conexoes_coletor_esgoto
            carregaResumo tubo_tubos_conexoes_coletor_esgoto.row, tubo_tubos_conexoes_coletor_esgoto
            houveDigitacaoSoldaveisAgua = False
        End If
    End If
    
End Sub

Private Sub tubo_conexoes_defofo_Click()
    
    If tubo_conexoes_defofo.SelectionMode = flexSelectionFree Then
        If tubo_conexoes_defofo.col = 1 Then
            tubo_conexoes_defofo.row = tubo_conexoes_defofo.row
            tubo_conexoes_defofo.col = 4
        Else
            tubo_conexoes_defofo.row = tubo_conexoes_defofo.row
            tubo_conexoes_defofo.col = tubo_conexoes_defofo.col
        End If
    End If
    ControleLostFocus = False
End Sub

Private Sub tubo_conexoes_defofo_DblClick()
    
    If tubo_conexoes_defofo.SelectionMode = flexSelectionByRow Then
        tubo_conexoes_defofo.SelectionMode = flexSelectionFree
        tubo_conexoes_defofo.Refresh
        tubo_conexoes_defofo_Click
        'tubos_conexoes_irri_azuis_SelChange
        Exit Sub
    End If
    
End Sub

Private Sub tubo_conexoes_defofo_KeyPress(KeyAscii As Integer)
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim col As Integer
    Dim row As Integer
    
    Select Case KeyAscii
    
    Case vbKeyReturn, vbKeyTab
    'move para a proxima celula.
    
    With tubo_conexoes_defofo
    
      If .col + 1 <= .Cols - 1 Then
         .col = .col + 1
      Else
         If .row + 1 <= .rows - 1 Then
             .row = .row + 1
             .col = 0
         Else
             .row = 1
             .col = 0
         End If
      End If
    End With
    
    Case Is < 32
        
    Case Else
        
        If tubo_conexoes_defofo.col = 6 Then
            If auxChange = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubo_conexoes_defofo
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "," Then
                    With tubo_conexoes_defofo
                       .Text = .Text & Chr(KeyAscii)
                    End With
                End If
                If Chr(KeyAscii) = "." Then
                    With tubo_conexoes_defofo
                       .Text = .Text & ","
                    End With
                End If
            End If
            If auxChange = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubo_conexoes_defofo
                       .Text = Chr(KeyAscii)
                    End With
                End If
                auxChange = False
                auxSelChange = True
                rowAux = tubo_conexoes_defofo.row
            End If
        End If
        If tubo_conexoes_defofo.col = 4 Then
            If auxChange = False Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubo_conexoes_defofo
                       .Text = .Text & Chr(KeyAscii)
                    End With
                    houveDigitacaoEsgoto = True
                    auxSelChange = True
                    rowAux = tubo_conexoes_defofo.row
                End If
            End If
            If auxChange = True Then
                If IsNumeric(Chr(KeyAscii)) Then
                    With tubo_conexoes_defofo
                       .Text = Chr(KeyAscii)
                    End With
                    houveDigitacaoEsgoto = True
                    auxSelChange = True
                    rowAux = tubo_conexoes_defofo.row
                End If
                auxChange = False
            End If
        End If
    End Select
    
    If KeyAscii = 8 Then
        If tubo_conexoes_defofo.col = 4 Or tubo_conexoes_defofo.col = 6 Then
            With tubo_conexoes_defofo
            'remove o ultimo caractere
               If Len(.Text) Then
                  .Text = Left(.Text, Len(.Text) - 1)
               End If
            End With
            row = tubo_conexoes_defofo.row
            col = tubo_conexoes_defofo.col
            If tubo_conexoes_defofo.TextMatrix(row, col) = "" Then
                tubo_conexoes_defofo.TextMatrix(row, col) = 0
                If col = 4 Then
                    houveDigitacaoEsgoto = True
                    auxSelChange = True
                    rowAux = tubo_conexoes_defofo.row
                    tubo_conexoes_defofo_SelChange
                End If
                auxChange = True
            End If
        End If
    End If
    If KeyAscii = 13 Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto tubo_conexoes_defofo.row, tubo_conexoes_defofo
        auxChange = True
        If houveDigitacaoEsgoto = True Or tubo_conexoes_defofo.TextMatrix(rowAux, 4) > 0 Then
            PopularVariaveisCalculo tubo_conexoes_defofo.row, tubo_conexoes_defofo
            CalculaTotal tubo_conexoes_defofo.row, tubo_conexoes_defofo
            carregaResumo tubo_conexoes_defofo.row, tubo_conexoes_defofo
            houveDigitacaoEsgoto = False
        End If
    End If
    
End Sub

Private Sub Bto_aplica_tubos_conexoes_pba_Click()
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim row As Integer
    Dim i As Integer
    
    If desconto_tubos_conexoes_pba.Text = "" Then
        MsgBox "Favor informar o desconto!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    If desconto_tubos_conexoes_pba.Text <> "" Then
        If desconto_tubos_conexoes_pba.Text > 36 Then
            MsgBox "Desconto maior do que o permitido", vbOKOnly, "Atenção"
            'desconto_tubos_conexoes_pba.Text = 34.2
        End If
    End If
        
    If tubo_conexoes_pba.rows = 1 Then
        MsgBox "Não existe produto para aplicar ajuste!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    For i = 1 To tubo_conexoes_pba.rows - 1
        row = i
        'If tubo_conexoes_pba.TextMatrix(i, 6) <> 0 Then
            'Atribui valor ao Preço Unitário
            tubo_conexoes_pba.TextMatrix(i, 6) = desconto_tubos_conexoes_pba
            PopularVariaveisCalculo row, tubo_conexoes_pba
            CalculaDesconto
            
            If desconto_tubos_conexoes_pba > dlSumDscItem Then
                tubo_conexoes_pba.TextMatrix(i, 6) = dlSumDscItem
                tubo_conexoes_pba.TextMatrix(i, 7) = Format$(tubo_conexoes_pba.TextMatrix(i, 5) - ((tubo_conexoes_pba.TextMatrix(i, 5) * dlSumDscItem) / 100), "Currency")
                CalculaTotal row, tubo_conexoes_pba
            Else
                tubo_conexoes_pba.TextMatrix(i, 6) = desconto_tubos_conexoes_pba
                tubo_conexoes_pba.TextMatrix(i, 7) = Format$(tubo_conexoes_pba.TextMatrix(i, 5) - ((tubo_conexoes_pba.TextMatrix(i, 5) * desconto_tubos_conexoes_pba) / 100), "Currency")
                CalculaTotal row, tubo_conexoes_pba
            End If
        'End If
        
        If tubo_conexoes_pba.TextMatrix(i, 4) <> 0 Then
            PopularVariaveisCalculo row, tubo_conexoes_pba
            carregaResumo row, tubo_conexoes_pba
        End If
    Next
    
End Sub

Private Sub CalculaTotalResumo(row As Integer, grid As MSFlexGrid)
    
    grid.TextMatrix(row, 18) = grid.TextMatrix(row, 3) * grid.TextMatrix(row, 7)
    grid.TextMatrix(row, 9) = Format(Trim(MskVlrUnit) * Trim(MskSerie), "##,###,##0.00")
    'grid.TextMatrix(row, 18) = Format$(grid.TextMatrix(row, 18), "Currency")
    grid.Refresh
    
End Sub

Private Sub CalculaTotal(row As Integer, grid As MSFlexGrid)
    
    grid.TextMatrix(row, 8) = grid.TextMatrix(row, 4) * grid.TextMatrix(row, 7)
    grid.TextMatrix(row, 8) = Format$(grid.TextMatrix(row, 8), "Currency")
    grid.Refresh
    
End Sub

Private Sub Bto_aplica_tubos_conexoes_agua_Click()
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim row As Integer
    Dim i As Integer
    
    If desconto_tubos_conexoes_agua.Text = "" Then
        MsgBox "Favor informar o desconto!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    If tubos_conexoes_agua.rows = 1 Then
        MsgBox "Não existe produto para aplicar ajuste!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    If desconto_tubos_conexoes_agua.Text <> "" Then
        If desconto_tubos_conexoes_agua.Text > 36 Then
            MsgBox "Desconto maior do que o permitido", vbOKOnly, "Atenção"
            'desconto_tubos_conexoes_agua.Text = 34.2
        End If
    End If
    
    For i = 1 To tubos_conexoes_agua.rows - 1
        row = i
        'If tubo_conexoes_defofo.TextMatrix(i, 4) <> 0 Then
            'Atribui valor ao Preço Unitário
            tubos_conexoes_agua.TextMatrix(i, 6) = desconto_tubos_conexoes_agua
            PopularVariaveisCalculo row, tubos_conexoes_agua
            CalculaDesconto
            
            If desconto_tubos_conexoes_agua > dlSumDscItem Then
                tubos_conexoes_agua.TextMatrix(i, 6) = dlSumDscItem
                tubos_conexoes_agua.TextMatrix(i, 7) = Format$(tubos_conexoes_agua.TextMatrix(i, 5) - ((tubos_conexoes_agua.TextMatrix(i, 5) * dlSumDscItem) / 100), "Currency")
                CalculaTotal row, tubos_conexoes_agua
            Else
                tubos_conexoes_agua.TextMatrix(i, 6) = desconto_tubos_conexoes_agua
                tubos_conexoes_agua.TextMatrix(i, 7) = Format$(tubos_conexoes_agua.TextMatrix(i, 5) - ((tubos_conexoes_agua.TextMatrix(i, 5) * desconto_tubos_conexoes_agua) / 100), "Currency")
                CalculaTotal row, tubos_conexoes_agua
            End If
            
        'End If
        
        If tubos_conexoes_agua.TextMatrix(i, 4) <> 0 Then
            PopularVariaveisCalculo row, tubos_conexoes_agua
            carregaResumo row, tubos_conexoes_agua
        End If
    Next
    
End Sub

Private Sub Bto_aplica_tubos_conexoes_defofo_Click()
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim row As Integer
    Dim i As Integer
    
    If desconto_tubos_conexoes_defofo.Text = "" Then
        MsgBox "Favor informar o desconto!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    If desconto_tubos_conexoes_defofo.Text <> "" Then
        If desconto_tubos_conexoes_defofo.Text > 36 Then
            MsgBox "Desconto maior do que o permitido", vbOKOnly, "Atenção"
            'desconto_tubos_conexoes_defofo.Text = 34.2
        End If
    End If
    
    If tubo_conexoes_defofo.rows = 1 Then
        MsgBox "Não existe produto para aplicar ajuste!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    For i = 1 To tubo_conexoes_defofo.rows - 1
        row = i
        'If tubo_conexoes_defofo.TextMatrix(i, 4) <> 0 Then
            'Atribui valor ao Preço Unitário
            tubo_conexoes_defofo.TextMatrix(i, 6) = desconto_tubos_conexoes_defofo
            PopularVariaveisCalculo row, tubo_conexoes_defofo
            CalculaDesconto
            
            If desconto_tubos_conexoes_defofo > dlSumDscItem Then
                tubo_conexoes_defofo.TextMatrix(i, 6) = dlSumDscItem
                tubo_conexoes_defofo.TextMatrix(i, 7) = Format$(tubo_conexoes_defofo.TextMatrix(i, 5) - ((tubo_conexoes_defofo.TextMatrix(i, 5) * dlSumDscItem) / 100), "Currency")
                CalculaTotal row, tubo_conexoes_defofo
            Else
                tubo_conexoes_defofo.TextMatrix(i, 6) = desconto_tubos_conexoes_defofo
                tubo_conexoes_defofo.TextMatrix(i, 7) = Format$(tubo_conexoes_defofo.TextMatrix(i, 5) - ((tubo_conexoes_defofo.TextMatrix(i, 5) * desconto_tubos_conexoes_defofo) / 100), "Currency")
                CalculaTotal row, tubo_conexoes_defofo
            End If
        'End If
        
        If tubo_conexoes_defofo.TextMatrix(i, 4) <> 0 Then
            PopularVariaveisCalculo row, tubo_conexoes_defofo
            carregaResumo row, tubo_conexoes_defofo
        End If
    Next
    
End Sub

Private Sub Bto_aplica_tubos_conexoes_coletor_esgoto_ocre_Click()
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim row As Integer
    Dim i As Integer
    
    If desconto_tubos_conexoes_coletor_esgoto_ocre.Text = "" Then
        MsgBox "Favor informar o desconto!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    If desconto_tubos_conexoes_coletor_esgoto_ocre.Text <> "" Then
        If desconto_tubos_conexoes_coletor_esgoto_ocre.Text > 36 Then
            MsgBox "Desconto maior do que o permitido", vbOKOnly, "Atenção"
            'desconto_tubos_conexoes_coletor_esgoto_ocre.Text = 34.2
        End If
    End If
    
    If tubo_tubos_conexoes_coletor_esgoto.rows = 1 Then
        MsgBox "Não existe produto para aplicar ajuste!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    For i = 1 To tubo_tubos_conexoes_coletor_esgoto.rows - 1
        row = i
        'If tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 6) <> 0 Then
            'Atribui valor ao Preço Unitário
            tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 6) = desconto_tubos_conexoes_coletor_esgoto_ocre
            PopularVariaveisCalculo row, tubo_tubos_conexoes_coletor_esgoto
            CalculaDesconto
            
            If desconto_tubos_conexoes_coletor_esgoto_ocre > dlSumDscItem Then
                tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 6) = dlSumDscItem
                tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 7) = Format$(tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 5) - ((tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 5) * dlSumDscItem) / 100), "Currency")
                CalculaTotal row, tubo_tubos_conexoes_coletor_esgoto
            Else
                tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 6) = desconto_tubos_conexoes_coletor_esgoto_ocre
                tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 7) = Format$(tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 5) - ((tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 5) * desconto_tubos_conexoes_coletor_esgoto_ocre) / 100), "Currency")
                CalculaTotal row, tubo_tubos_conexoes_coletor_esgoto
            End If
        'End If
        
        If tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 4) <> 0 Then
            PopularVariaveisCalculo row, tubo_tubos_conexoes_coletor_esgoto
            carregaResumo row, tubo_tubos_conexoes_coletor_esgoto
        End If
    Next
    
End Sub

Private Sub Bto_aplica__tubos_conexoes_predial_Click()
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim row As Integer
    Dim i As Integer
    
    If desconto_tubos_conexoes_predial.Text = "" Then
        MsgBox "Favor informar o desconto!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    If desconto_tubos_conexoes_predial.Text <> "" Then
        If desconto_tubos_conexoes_predial.Text > 36 Then
            MsgBox "Desconto maior do que o permitido", vbOKOnly, "Atenção"
            'desconto_tubos_conexoes_predial.Text = 34.2
        End If
    End If
    
    If tubos_conexoes_predial.rows = 1 Then
        MsgBox "Não existe produto para aplicar ajuste!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    For i = 1 To tubos_conexoes_predial.rows - 1
        row = i
        'If tubos_conexoes_predial.TextMatrix(i, 4) <> 0 Then
            'Atribui valor ao Preço Unitário
            tubos_conexoes_predial.TextMatrix(i, 6) = desconto_tubos_conexoes_predial
            PopularVariaveisCalculo row, tubos_conexoes_predial
            CalculaDesconto
            
            If desconto_tubos_conexoes_predial > dlSumDscItem Then
                tubos_conexoes_predial.TextMatrix(i, 6) = dlSumDscItem
                tubos_conexoes_predial.TextMatrix(i, 7) = Format$(tubos_conexoes_predial.TextMatrix(i, 5) - ((tubos_conexoes_predial.TextMatrix(i, 5) * dlSumDscItem) / 100), "Currency")
                CalculaTotal row, tubos_conexoes_predial
            Else
                tubos_conexoes_predial.TextMatrix(i, 6) = desconto_tubos_conexoes_predial
                tubos_conexoes_predial.TextMatrix(i, 7) = Format$(tubos_conexoes_predial.TextMatrix(i, 5) - ((tubos_conexoes_predial.TextMatrix(i, 5) * desconto_tubos_conexoes_predial) / 100), "Currency")
                CalculaTotal row, tubos_conexoes_predial
            End If
        'End If
        
        If tubos_conexoes_predial.TextMatrix(i, 4) <> 0 Then
            PopularVariaveisCalculo row, tubos_conexoes_predial
            carregaResumo row, tubos_conexoes_predial
        End If
    Next
    
End Sub

Private Sub PopulaVariaveisQuantidade()
    
    On Error Resume Next
    
    If Trim(MskNumNf) = "" Or Trim(MskNumNf) = 0 Then
        Exit Sub
    End If
    
    sgQuery = "SELECT a.*, b.* from PRODUTO a, PRECO_PRODUTO b"
    sgQuery = sgQuery + "   WHERE a.flgsitu = 'N'"
    sgQuery = sgQuery + "     and a.Codprd = " & Trim(MskNumNf)
    sgQuery = sgQuery + "     and a.codprd = b.codprd"
    sgQuery = sgQuery + "     and b.datativ = (select max(datativ) from preco_produto"
    sgQuery = sgQuery + "                       Where Codprd = " & Trim(MskNumNf)
    sgQuery = sgQuery + "                         and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"
  
    Call Consulta(sgQuery)
    
    If Rs.EOF Then
    
        MsgBox "Produto inexistente ou fora de linha", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        'MskNumNf.SetFocus
        
        Exit Sub
        
    End If
  
    LblRotaRec.Caption = IIf(IsNull(Rs!DSCPRD), "", Rs!DSCPRD)
    
    ilIdeGrp = IIf(Trim(Rs!IdeGrp) = "", 0, Trim(Rs!IdeGrp))
    dlPesUnt = IIf(Trim(Rs!PesUnt) = "", 0, Trim(Rs!PesUnt))
    dlValUntN = IIf(Trim(Rs!ValUntN) = "", 0, Trim(Rs!ValUntN))
    dlValUntA = IIf(Trim(Rs!ValUntA) = "", 0, Trim(Rs!ValUntA))
    dlValUntB = IIf(Trim(Rs!ValUntB) = "", 0, Trim(Rs!ValUntB))
    dlMrgPrd = IIf(Trim(Rs!MrgPrd) = "", 0, Trim(Rs!MrgPrd))
    dlValCusUntQtd = IIf(Trim(Rs!valcusuntqtd) = "", 0, Trim(Rs!valcusuntqtd))
    dlValCusAdicQtd = IIf(Trim(Rs!valcusadicqtd) = "", 0, Trim(Rs!valcusadicqtd))
    dlAlqImpFed = IIf(Trim(Rs!AlqImpFed) = "", 0, Trim(Rs!AlqImpFed))
    ilQtdEmb = IIf(Trim(Rs!QtdEmb) = "", 1, Trim(Rs!QtdEmb))
    ilFlgKit = Trim(Rs!FlgKit)
    
    Rs.Close
    
    Set Rs = Nothing
    
    If LeituraCliente = False Then
        LimpaGeral
    End If
    
    Dim Linhas As Integer
    
    DoEvents
        
    'If Me.ActiveControl.Name = "BtoSair" Or Me.ActiveControl.Name = "BtoLimpaCTRC" Or Me.ActiveControl.Name = "BtoLimpaNF" Or blLG = True Or bgBloqPed = True Or MskNroPedido.Text = 0 Then
        'Exit Sub
    'End If

    If Trim(MskNumNf) = "" Or Trim(MskNumNf) = 0 Then
        Exit Sub
    End If
    
    sgQuery = "SELECT a.*, b.* from PRODUTO a, PRECO_PRODUTO b"
    sgQuery = sgQuery + "   WHERE a.flgsitu = 'N'"
    sgQuery = sgQuery + "     and a.Codprd = " & Trim(MskNumNf)
    sgQuery = sgQuery + "     and a.codprd = b.codprd"
    sgQuery = sgQuery + "     and b.datativ = (select max(datativ) from preco_produto"
    sgQuery = sgQuery + "                       Where Codprd = " & Trim(MskNumNf)
    sgQuery = sgQuery + "                         and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"
  
    Call Consulta(sgQuery)
    
    If Rs.EOF Then
    
        MsgBox "Produto inexistente ou fora de linha", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        'MskNumNf.SetFocus
        
        Exit Sub
        
    End If
  
    LblRotaRec.Caption = IIf(IsNull(Rs!DSCPRD), "", Rs!DSCPRD)
    
    ilIdeGrp = IIf(Trim(Rs!IdeGrp) = "", 0, Trim(Rs!IdeGrp))
    dlPesUnt = IIf(Trim(Rs!PesUnt) = "", 0, Trim(Rs!PesUnt))
    dlValUntN = IIf(Trim(Rs!ValUntN) = "", 0, Trim(Rs!ValUntN))
    dlValUntA = IIf(Trim(Rs!ValUntA) = "", 0, Trim(Rs!ValUntA))
    dlValUntB = IIf(Trim(Rs!ValUntB) = "", 0, Trim(Rs!ValUntB))
    dlMrgPrd = IIf(Trim(Rs!MrgPrd) = "", 0, Trim(Rs!MrgPrd))
    dlValCusUntQtd = IIf(Trim(Rs!valcusuntqtd) = "", 0, Trim(Rs!valcusuntqtd))
    dlValCusAdicQtd = IIf(Trim(Rs!valcusadicqtd) = "", 0, Trim(Rs!valcusadicqtd))
    dlAlqImpFed = IIf(Trim(Rs!AlqImpFed) = "", 0, Trim(Rs!AlqImpFed))
    ilQtdEmb = IIf(Trim(Rs!QtdEmb) = "", 1, Trim(Rs!QtdEmb))
    ilFlgKit = Trim(Rs!FlgKit)
    
    Rs.Close
    
    Set Rs = Nothing

    If blModificar = True Then
    
        ''VSValUnit.Value = ilNumTab
        
        If ilNumTab = 0 Then
            
            MskVlrUnit = dlValUntN
            ''LblUnit.Caption = "Valor Unitário"
            
        Else
        
            If ilNumTab = 1 Then
        
                If dlValUntA = 0 Then
                    MskVlrUnit = dlValUntN
                    ''VSValUnit.Value = 0
                    ''LblUnit.Caption = "Valor Unitário"
                Else
                    MskVlrUnit = dlValUntA
                    ''LblUnit.Caption = "Valor Unitário - A"
                End If
                
            Else
        
                If dlValUntB = 0 Then
                    MskVlrUnit = dlValUntN
                    ''VSValUnit.Value = 0
                    ''LblUnit.Caption = "Valor Unitário"
                Else
                    MskVlrUnit = dlValUntB
                    ''LblUnit.Caption = "Valor Unitário - B"
                End If
        
            End If
        
        End If
    
        DoEvents
    
    Else
     
        If ilNumTab = 0 Then
            MskVlrUnit = Format(dlValUntN, sgStrF2)
            ''VSValUnit.Value = 0
        End If
        
        For Linhas = 1 To GrdNotaCliente.rows - 1
        
            If Trim(GrdNotaCliente.TextMatrix(Linhas, 0)) = Format(Trim(MskNumNf), "0000") Then
                
                If GrdNotaCliente.rows = 2 Then
                    GrdNotaCliente.rows = GrdNotaCliente.rows - 1
                Else
                    GrdNotaCliente.RemoveItem (Linhas)
                    GrdNotaCliente.Refresh
                End If
                Exit For
            
            End If
        
        Next Linhas
    
    End If
  
    'Acha desconto promocional destacado para o produto ou grupo (por representante)
    
    'dlPerDesPrd = 0
    
    'sgQuery = "select PerDsc from Desconto_promocional where CodRep = " & Trim(ilCodRep)
    'sgQuery = sgQuery + " and IdeGrp = " & ilIdeGrp
    'sgQuery = sgQuery + " and Codprd = " & Trim(MskNumNf)
    
    'Call consulta(sgQuery)
    
    'If Not Rs.EOF Then
        
        'dlPerDesPrd = IIf(Trim(Rs!PerDsc) = "", 0, Trim(Rs!PerDsc))
    
    'Else
        
        'Rs.Close
        
        'Set Rs = Nothing
        
        'sgQuery = "select PerDsc from Desconto_promocional where CodRep = " & Trim(ilCodRep)
        'sgQuery = sgQuery + " and IdeGrp = " & ilIdeGrp
        'sgQuery = sgQuery + " and Codprd is null "
        
        'Call consulta(sgQuery)
        
        'If Not Rs.EOF Then
            'dlPerDesPrd = IIf(Trim(Rs!PerDsc) = "", 0, Trim(Rs!PerDsc))
        'End If
    
    'End If
    
    'If dlPerDesPrd = 0 Then
        dlPerDesPrd = dlPerDesRep
    'End If
    
    'Rs.Close
    
    'Set Rs = Nothing
  
    CalculaDesconto
    Call SelecionaTudo
    
End Sub

Private Sub Bto_aplica__tubos_conexoes_roscaveis_Click()
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim row As Integer
    Dim i As Integer
    
    If desconto_tubos_conexoes_roscaveis.Text = "" Then
        MsgBox "Favor informar o desconto!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    If desconto_tubos_conexoes_roscaveis.Text <> "" Then
        If desconto_tubos_conexoes_roscaveis.Text > 36 Then
            MsgBox "Desconto maior do que o permitido", vbOKOnly, "Atenção"
            'desconto_tubos_conexoes_roscaveis.Text = 34.2
        End If
    End If
    
    If tubos_conexoes_roscaveis.rows = 1 Then
        MsgBox "Não existe produto para aplicar ajuste!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    For i = 1 To tubos_conexoes_roscaveis.rows - 1
        row = i
        'If tubos_conexoes_roscaveis.TextMatrix(i, 4) <> 0 Then
            'Atribui valor ao Preço Unitário
            tubos_conexoes_roscaveis.TextMatrix(i, 6) = desconto_tubos_conexoes_roscaveis
            PopularVariaveisCalculo row, tubos_conexoes_roscaveis
            CalculaDesconto
            
            If desconto_tubos_conexoes_roscaveis > dlSumDscItem Then
                tubos_conexoes_roscaveis.TextMatrix(i, 6) = dlSumDscItem
                tubos_conexoes_roscaveis.TextMatrix(i, 7) = Format$(tubos_conexoes_roscaveis.TextMatrix(i, 5) - ((tubos_conexoes_roscaveis.TextMatrix(i, 5) * desconto_tubos_conexoes_roscaveis) / 100), "Currency")
                CalculaTotal row, tubos_conexoes_roscaveis
            Else
                tubos_conexoes_roscaveis.TextMatrix(i, 6) = desconto_tubos_conexoes_roscaveis
                tubos_conexoes_roscaveis.TextMatrix(i, 7) = Format$(tubos_conexoes_roscaveis.TextMatrix(i, 5) - ((tubos_conexoes_roscaveis.TextMatrix(i, 5) * desconto_tubos_conexoes_roscaveis) / 100), "Currency")
                CalculaTotal row, tubos_conexoes_roscaveis
            End If
        'End If
        
        If tubos_conexoes_roscaveis.TextMatrix(i, 4) <> 0 Then
            PopularVariaveisCalculo row, tubos_conexoes_roscaveis
            carregaResumo row, tubos_conexoes_roscaveis
        End If
    Next
    
End Sub

Private Sub Bto_Aplica_tubos_conexoes_irri_azuis_Click()
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim row As Integer
    Dim i As Integer
    
    If desconto_tubos_conexoes_irri_azuis.Text = "" Then
        MsgBox "Favor informar o desconto!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    If desconto_tubos_conexoes_irri_azuis.Text <> "" Then
        If desconto_tubos_conexoes_irri_azuis.Text > 36 Then
            MsgBox "Desconto maior do que o permitido", vbOKOnly, "Atenção"
        End If
    End If
    
    If tubos_conexoes_irri_azuis.rows = 1 Then
        MsgBox "Não existe produto para aplicar ajuste!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    For i = 1 To tubos_conexoes_irri_azuis.rows - 1
        row = i
        'If tubos_conexoes_irri_azuis.TextMatrix(i, 4) <> 0 Then
            'Atribui valor ao Preço Unitário
            tubos_conexoes_irri_azuis.TextMatrix(i, 6) = desconto_tubos_conexoes_irri_azuis
            PopularVariaveisCalculo row, tubos_conexoes_irri_azuis
            CalculaDesconto
            
            If desconto_tubos_conexoes_irri_azuis > dlSumDscItem Then
                tubos_conexoes_irri_azuis.TextMatrix(i, 6) = dlSumDscItem
                tubos_conexoes_irri_azuis.TextMatrix(i, 7) = Format$(tubos_conexoes_irri_azuis.TextMatrix(i, 5) - ((tubos_conexoes_irri_azuis.TextMatrix(i, 5) * dlSumDscItem) / 100), "Currency")
                CalculaTotal row, tubos_conexoes_irri_azuis
            Else
                tubos_conexoes_irri_azuis.TextMatrix(i, 6) = desconto_tubos_conexoes_irri_azuis
                tubos_conexoes_irri_azuis.TextMatrix(i, 7) = Format$(tubos_conexoes_irri_azuis.TextMatrix(i, 5) - ((tubos_conexoes_irri_azuis.TextMatrix(i, 5) * desconto_tubos_conexoes_irri_azuis) / 100), "Currency")
                CalculaTotal row, tubos_conexoes_irri_azuis
            End If
        'End If
        
        If tubos_conexoes_irri_azuis.TextMatrix(i, 4) <> 0 Then
            PopularVariaveisCalculo row, tubos_conexoes_irri_azuis
            carregaResumo row, tubos_conexoes_irri_azuis
        End If
    Next
    
End Sub

Private Sub AtualizaGridAuxiliar(row As Integer)
    
    Dim i As Integer
    
    For i = 1 To tubo_conexoes_defofo.rows - 1
        If Format(tubo_conexoes_defofo.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(row, 0) And (Format(tubo_conexoes_defofo.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(row, 6) Or tubo_conexoes_defofo.TextMatrix(i, 4) <> GrdNotaCliente.TextMatrix(row, 3)) Then
            tubo_conexoes_defofo.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(row, 6)
            tubo_conexoes_defofo.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(row, 3)
            tubo_conexoes_defofo.TextMatrix(i, 7) = (tubo_conexoes_defofo.TextMatrix(i, 5) - ((tubo_conexoes_defofo.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(row, 6)) / 100))
            tubo_conexoes_defofo.TextMatrix(i, 7) = Format$(tubo_conexoes_defofo.TextMatrix(i, 7), "Currency")
            CalculaTotal i, tubo_conexoes_defofo
            ControleAtualizaGrid = True
            carregaResumo i, tubo_conexoes_defofo
        End If
    Next
    
    For i = 1 To tubos_conexoes_irri_azuis.rows - 1
        If Format(tubos_conexoes_irri_azuis.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(row, 0) And (Format(tubos_conexoes_irri_azuis.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(row, 6) Or tubos_conexoes_irri_azuis.TextMatrix(i, 4) <> GrdNotaCliente.TextMatrix(row, 3)) Then
            tubos_conexoes_irri_azuis.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(row, 6)
            tubos_conexoes_irri_azuis.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(row, 3)
            tubos_conexoes_irri_azuis.TextMatrix(i, 7) = (tubos_conexoes_irri_azuis.TextMatrix(i, 5) - ((tubos_conexoes_irri_azuis.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(row, 6)) / 100))
            tubos_conexoes_irri_azuis.TextMatrix(i, 7) = Format$(tubos_conexoes_irri_azuis.TextMatrix(i, 7), "Currency")
            CalculaTotal i, tubos_conexoes_irri_azuis
            ControleAtualizaGrid = True
            carregaResumo i, tubos_conexoes_irri_azuis
        End If
    Next
    
    For i = 1 To tubos_conexoes_roscaveis.rows - 1
        If Format(tubos_conexoes_roscaveis.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(row, 0) And (Format(tubos_conexoes_roscaveis.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(row, 6) Or tubos_conexoes_roscaveis.TextMatrix(i, 4) <> GrdNotaCliente.TextMatrix(row, 3)) Then
            tubos_conexoes_roscaveis.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(row, 6)
            tubos_conexoes_roscaveis.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(row, 3)
            tubos_conexoes_roscaveis.TextMatrix(i, 7) = (tubos_conexoes_roscaveis.TextMatrix(i, 5) - ((tubos_conexoes_roscaveis.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(row, 6)) / 100))
            tubos_conexoes_roscaveis.TextMatrix(i, 7) = Format$(tubos_conexoes_roscaveis.TextMatrix(i, 7), "Currency")
            CalculaTotal i, tubos_conexoes_roscaveis
            ControleAtualizaGrid = True
            carregaResumo i, tubos_conexoes_roscaveis
        End If
    Next
    
    For i = 1 To tubos_conexoes_predial.rows - 1
        If Format(tubos_conexoes_predial.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(row, 0) And (Format(tubos_conexoes_predial.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(row, 6) Or tubos_conexoes_predial.TextMatrix(i, 4) <> GrdNotaCliente.TextMatrix(row, 3)) Then
            tubos_conexoes_predial.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(row, 6)
            tubos_conexoes_predial.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(row, 3)
            tubos_conexoes_predial.TextMatrix(i, 7) = (tubos_conexoes_predial.TextMatrix(i, 5) - ((tubos_conexoes_predial.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(row, 6)) / 100))
            tubos_conexoes_predial.TextMatrix(i, 7) = Format$(tubos_conexoes_predial.TextMatrix(i, 7), "Currency")
            CalculaTotal i, tubos_conexoes_predial
            ControleAtualizaGrid = True
            carregaResumo i, tubos_conexoes_predial
        End If
    Next
    
    For i = 1 To tubo_tubos_conexoes_coletor_esgoto.rows - 1
        If Format(tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(row, 0) And (Format(tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(row, 6) Or tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 4) <> GrdNotaCliente.TextMatrix(row, 3)) Then
            tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(row, 6)
            tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(row, 3)
            tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 7) = (tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 5) - ((tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(row, 6)) / 100))
            tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 7) = Format$(tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 7), "Currency")
            CalculaTotal i, tubo_tubos_conexoes_coletor_esgoto
            ControleAtualizaGrid = True
            carregaResumo i, tubo_tubos_conexoes_coletor_esgoto
        End If
    Next
    
    For i = 1 To tubo_conexoes_pba.rows - 1
        If Format(tubo_conexoes_pba.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(row, 0) And (Format(tubo_conexoes_pba.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(row, 6) Or tubo_conexoes_pba.TextMatrix(i, 4) <> GrdNotaCliente.TextMatrix(row, 3)) Then
            tubo_conexoes_pba.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(row, 6)
            tubo_conexoes_pba.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(row, 3)
            tubo_conexoes_pba.TextMatrix(i, 7) = (tubo_conexoes_pba.TextMatrix(i, 5) - ((tubo_conexoes_pba.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(row, 6)) / 100))
            tubo_conexoes_pba.TextMatrix(i, 7) = Format$(tubo_conexoes_pba.TextMatrix(i, 7), "Currency")
            CalculaTotal i, tubo_conexoes_pba
            ControleAtualizaGrid = True
            carregaResumo i, tubo_conexoes_pba
        End If
    Next
    
    For i = 1 To tubos_conexoes_agua.rows - 1
        If Format(tubos_conexoes_agua.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(row, 0) And (Format(tubos_conexoes_agua.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(row, 6) Or tubos_conexoes_agua.TextMatrix(i, 4) <> GrdNotaCliente.TextMatrix(row, 3)) Then
            tubos_conexoes_agua.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(row, 6)
            tubos_conexoes_agua.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(row, 3)
            tubos_conexoes_agua.TextMatrix(i, 7) = (tubos_conexoes_agua.TextMatrix(i, 5) - ((tubos_conexoes_agua.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(row, 6)) / 100))
            tubos_conexoes_agua.TextMatrix(i, 7) = Format$(tubos_conexoes_agua.TextMatrix(i, 7), "Currency")
            CalculaTotal i, tubos_conexoes_agua
            ControleAtualizaGrid = True
            carregaResumo i, tubos_conexoes_agua
        End If
    Next
    
    ControleAtualizaGrid = False
End Sub

Private Sub LimpaGridAuxiliar()
    
    Dim i As Integer
    Dim J As Integer
    
    For i = 1 To tubo_conexoes_defofo.rows - 1
        tubo_conexoes_defofo.TextMatrix(i, 6) = 0
        tubo_conexoes_defofo.TextMatrix(i, 4) = 0
        tubo_conexoes_defofo.TextMatrix(i, 7) = Format$(0, "Currency")
        CalculaTotal i, tubo_conexoes_defofo
    Next
    
    For i = 1 To tubos_conexoes_irri_azuis.rows - 1
        tubos_conexoes_irri_azuis.TextMatrix(i, 6) = 0
        tubos_conexoes_irri_azuis.TextMatrix(i, 4) = 0
        tubos_conexoes_irri_azuis.TextMatrix(i, 7) = Format$(0, "Currency")
        CalculaTotal i, tubos_conexoes_irri_azuis
    Next
    
    For i = 1 To tubos_conexoes_roscaveis.rows - 1
        tubos_conexoes_roscaveis.TextMatrix(i, 6) = 0
        tubos_conexoes_roscaveis.TextMatrix(i, 4) = 0
        tubos_conexoes_roscaveis.TextMatrix(i, 7) = Format$(0, "Currency")
        CalculaTotal i, tubos_conexoes_roscaveis
    Next
    
    For i = 1 To tubos_conexoes_predial.rows - 1
        tubos_conexoes_predial.TextMatrix(i, 6) = 0
        tubos_conexoes_predial.TextMatrix(i, 4) = 0
        tubos_conexoes_predial.TextMatrix(i, 7) = Format$(0, "Currency")
        CalculaTotal i, tubos_conexoes_predial
    Next
    
    For i = 1 To tubo_tubos_conexoes_coletor_esgoto.rows - 1
        tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 6) = 0
        tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 4) = 0
        tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 7) = Format$(0, "Currency")
        CalculaTotal i, tubo_tubos_conexoes_coletor_esgoto
    Next
    
    For i = 1 To tubo_conexoes_pba.rows - 1
        tubo_conexoes_pba.TextMatrix(i, 6) = 0
        tubo_conexoes_pba.TextMatrix(i, 4) = 0
        tubo_conexoes_pba.TextMatrix(i, 7) = Format$(0, "Currency")
        CalculaTotal i, tubo_conexoes_pba
    Next
    
    For i = 1 To tubos_conexoes_agua.rows - 1
        tubos_conexoes_agua.TextMatrix(i, 6) = 0
        tubos_conexoes_agua.TextMatrix(i, 4) = 0
        tubos_conexoes_agua.TextMatrix(i, 7) = Format$(0, "Currency")
        CalculaTotal i, tubos_conexoes_agua
    Next
    
End Sub

Private Sub LimpaRegistroGridAuxiliar(codigo As Integer)
    
    Dim i As Integer
    Dim J As Integer
    
    For i = 1 To tubo_conexoes_defofo.rows - 1
        If tubo_conexoes_defofo.TextMatrix(i, 1) = codigo Then
            tubo_conexoes_defofo.TextMatrix(i, 6) = 0
            tubo_conexoes_defofo.TextMatrix(i, 4) = 0
            tubo_conexoes_defofo.TextMatrix(i, 7) = tubo_conexoes_defofo.TextMatrix(i, 5)
            CalculaTotal i, tubo_conexoes_defofo
            Exit Sub
        End If
    Next
    
    For i = 1 To tubos_conexoes_irri_azuis.rows - 1
        If tubos_conexoes_irri_azuis.TextMatrix(i, 1) = codigo Then
            tubos_conexoes_irri_azuis.TextMatrix(i, 6) = 0
            tubos_conexoes_irri_azuis.TextMatrix(i, 4) = 0
            tubos_conexoes_irri_azuis.TextMatrix(i, 7) = tubos_conexoes_irri_azuis.TextMatrix(i, 5)
            CalculaTotal i, tubos_conexoes_irri_azuis
            Exit Sub
        End If
    Next
    
    For i = 1 To tubos_conexoes_roscaveis.rows - 1
        If tubos_conexoes_roscaveis.TextMatrix(i, 1) = codigo Then
            tubos_conexoes_roscaveis.TextMatrix(i, 6) = 0
            tubos_conexoes_roscaveis.TextMatrix(i, 4) = 0
            tubos_conexoes_roscaveis.TextMatrix(i, 7) = tubos_conexoes_roscaveis.TextMatrix(i, 5)
            CalculaTotal i, tubos_conexoes_roscaveis
            Exit Sub
        End If
    Next
    
    For i = 1 To tubos_conexoes_predial.rows - 1
        If tubos_conexoes_predial.TextMatrix(i, 1) = codigo Then
            tubos_conexoes_predial.TextMatrix(i, 6) = 0
            tubos_conexoes_predial.TextMatrix(i, 4) = 0
            tubos_conexoes_predial.TextMatrix(i, 7) = tubos_conexoes_predial.TextMatrix(i, 5)
            CalculaTotal i, tubos_conexoes_predial
            Exit Sub
        End If
    Next
    
    For i = 1 To tubo_tubos_conexoes_coletor_esgoto.rows - 1
        If tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 1) = codigo Then
            tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 6) = 0
            tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 4) = 0
            tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 7) = tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 5)
            CalculaTotal i, tubo_tubos_conexoes_coletor_esgoto
            Exit Sub
        End If
    Next
    
    For i = 1 To tubo_conexoes_pba.rows - 1
        If tubo_conexoes_pba.TextMatrix(i, 1) = codigo Then
            tubo_conexoes_pba.TextMatrix(i, 6) = 0
            tubo_conexoes_pba.TextMatrix(i, 4) = 0
            tubo_conexoes_pba.TextMatrix(i, 7) = tubo_conexoes_pba.TextMatrix(i, 5)
            CalculaTotal i, tubo_conexoes_pba
            Exit Sub
        End If
    Next
    
    For i = 1 To tubos_conexoes_agua.rows - 1
        If tubos_conexoes_agua.TextMatrix(i, 1) = codigo Then
            tubos_conexoes_agua.TextMatrix(i, 6) = 0
            tubos_conexoes_agua.TextMatrix(i, 4) = 0
            tubos_conexoes_agua.TextMatrix(i, 7) = tubos_conexoes_agua.TextMatrix(i, 5)
            CalculaTotal i, tubos_conexoes_agua
            Exit Sub
        End If
    Next
    
End Sub

Private Sub Bto_Aplica_resumo_Click()
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    Dim i As Integer
    Dim J As Integer
    
    If desconto_resumo.Text = "" Then
        MsgBox "Favor informar o desconto!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    If desconto_resumo.Text <> "" Then
        If desconto_resumo.Text > 36 Then
            MsgBox "Desconto maior do que o permitido", vbOKOnly, "Atenção"
            'desconto_resumo.Text = 34.2
        End If
    End If
    
    If GrdNotaCliente.rows = 1 Then
        MsgBox "Não existe produto para aplicar ajuste!", vbOKOnly, "Atenção"
        Exit Sub
    End If
    
    For i = 1 To GrdNotaCliente.rows - 1
        GrdNotaCliente.TextMatrix(i, 6) = desconto_resumo
        PopularVariaveisCalculoResumo i, GrdNotaCliente
        CalculaDesconto
        
        If desconto_resumo > dlSumDscItem Then
            GrdNotaCliente.TextMatrix(i, 6) = Format(dlSumDscItem, "##0.00")
            GrdNotaCliente.TextMatrix(i, 7) = Format((GrdNotaCliente.TextMatrix(i, 4) - ((GrdNotaCliente.TextMatrix(i, 4) * MskDatEmiNf) / 100)), "##,###,##0.00")
            CalculaTotalResumo i, GrdNotaCliente
        Else
            GrdNotaCliente.TextMatrix(i, 6) = Format(desconto_resumo.Text, "##0.00")
            GrdNotaCliente.TextMatrix(i, 7) = Format((GrdNotaCliente.TextMatrix(i, 4) - ((GrdNotaCliente.TextMatrix(i, 4) * MskDatEmiNf) / 100)), "##,###,##0.00")
            CalculaTotalResumo i, GrdNotaCliente
        End If
    Next
    
    CalculaIndice
    DefineCorResumo GrdNotaCliente.row, GrdNotaCliente
    
    For i = 1 To tubo_conexoes_defofo.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubo_conexoes_defofo.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubo_conexoes_defofo.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubo_conexoes_defofo.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubo_conexoes_defofo.TextMatrix(i, 7) = Format$((tubo_conexoes_defofo.TextMatrix(i, 5) - ((tubo_conexoes_defofo.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                CalculaTotal i, tubo_conexoes_defofo
            End If
        Next
    Next
    
    For i = 1 To tubos_conexoes_irri_azuis.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubos_conexoes_irri_azuis.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubos_conexoes_irri_azuis.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubos_conexoes_irri_azuis.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubos_conexoes_irri_azuis.TextMatrix(i, 7) = Format$((tubos_conexoes_irri_azuis.TextMatrix(i, 5) - ((tubos_conexoes_irri_azuis.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                CalculaTotal i, tubos_conexoes_irri_azuis
            End If
        Next
    Next
    
    For i = 1 To tubos_conexoes_roscaveis.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubos_conexoes_roscaveis.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubos_conexoes_roscaveis.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubos_conexoes_roscaveis.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubos_conexoes_roscaveis.TextMatrix(i, 7) = Format$((tubos_conexoes_roscaveis.TextMatrix(i, 5) - ((tubos_conexoes_roscaveis.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                CalculaTotal i, tubos_conexoes_roscaveis
            End If
        Next
    Next
    
    For i = 1 To tubos_conexoes_predial.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubos_conexoes_predial.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubos_conexoes_predial.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubos_conexoes_predial.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubos_conexoes_predial.TextMatrix(i, 7) = Format$((tubos_conexoes_predial.TextMatrix(i, 5) - ((tubos_conexoes_predial.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                CalculaTotal i, tubos_conexoes_predial
            End If
        Next
    Next
    
    For i = 1 To tubo_tubos_conexoes_coletor_esgoto.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 7) = Format$((tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 5) - ((tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                CalculaTotal i, tubo_tubos_conexoes_coletor_esgoto
            End If
        Next
    Next
    
    For i = 1 To tubo_conexoes_pba.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubo_conexoes_pba.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubo_conexoes_pba.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubo_conexoes_pba.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubo_conexoes_pba.TextMatrix(i, 7) = Format$((tubo_conexoes_pba.TextMatrix(i, 5) - ((tubo_conexoes_pba.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                CalculaTotal i, tubo_conexoes_pba
            End If
        Next
    Next
    
    For i = 1 To tubos_conexoes_agua.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubos_conexoes_agua.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubos_conexoes_agua.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubos_conexoes_agua.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubos_conexoes_agua.TextMatrix(i, 7) = Format$((tubos_conexoes_agua.TextMatrix(i, 5) - ((tubos_conexoes_agua.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                CalculaTotal i, tubos_conexoes_agua
            End If
        Next
    Next
    
End Sub

Private Sub CarregaGridConsulta()
    
    Dim i As Integer
    Dim J As Integer
    
    For i = 1 To tubo_conexoes_defofo.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubo_conexoes_defofo.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubo_conexoes_defofo.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubo_conexoes_defofo.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubo_conexoes_defofo.TextMatrix(i, 7) = Format$((tubo_conexoes_defofo.TextMatrix(i, 5) - ((tubo_conexoes_defofo.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                tubo_conexoes_defofo.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(J, 3)
                tubo_conexoes_defofo.TextMatrix(i, 8) = Format$(GrdNotaCliente.TextMatrix(J, 9), "Currency")
            End If
        Next
    Next
    
    For i = 1 To tubos_conexoes_irri_azuis.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubos_conexoes_irri_azuis.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubos_conexoes_irri_azuis.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubos_conexoes_irri_azuis.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubos_conexoes_irri_azuis.TextMatrix(i, 7) = Format$((tubos_conexoes_irri_azuis.TextMatrix(i, 5) - ((tubos_conexoes_irri_azuis.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                tubos_conexoes_irri_azuis.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(J, 3)
                tubos_conexoes_irri_azuis.TextMatrix(i, 8) = Format$(GrdNotaCliente.TextMatrix(J, 9), "Currency")
            End If
        Next
    Next
    
    For i = 1 To tubos_conexoes_roscaveis.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubos_conexoes_roscaveis.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubos_conexoes_roscaveis.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubos_conexoes_roscaveis.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubos_conexoes_roscaveis.TextMatrix(i, 7) = Format$((tubos_conexoes_roscaveis.TextMatrix(i, 5) - ((tubos_conexoes_roscaveis.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                tubos_conexoes_roscaveis.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(J, 3)
                tubos_conexoes_roscaveis.TextMatrix(i, 8) = Format$(GrdNotaCliente.TextMatrix(J, 9), "Currency")
            End If
        Next
    Next
    
    For i = 1 To tubos_conexoes_predial.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubos_conexoes_predial.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubos_conexoes_predial.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubos_conexoes_predial.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubos_conexoes_predial.TextMatrix(i, 7) = Format$((tubos_conexoes_predial.TextMatrix(i, 5) - ((tubos_conexoes_predial.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                tubos_conexoes_predial.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(J, 3)
                tubos_conexoes_predial.TextMatrix(i, 8) = Format$(GrdNotaCliente.TextMatrix(J, 9), "Currency")
            End If
        Next
    Next
    
    For i = 1 To tubo_tubos_conexoes_coletor_esgoto.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 7) = Format$((tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 5) - ((tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(J, 3)
                tubo_tubos_conexoes_coletor_esgoto.TextMatrix(i, 8) = Format$(GrdNotaCliente.TextMatrix(J, 9), "Currency")
            End If
        Next
    Next
    
    For i = 1 To tubo_conexoes_pba.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubo_conexoes_pba.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubo_conexoes_pba.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubo_conexoes_pba.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubo_conexoes_pba.TextMatrix(i, 7) = Format$((tubo_conexoes_pba.TextMatrix(i, 5) - ((tubo_conexoes_pba.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                tubo_conexoes_pba.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(J, 3)
                tubo_conexoes_pba.TextMatrix(i, 8) = Format$(GrdNotaCliente.TextMatrix(J, 9), "Currency")
            End If
        Next
    Next
    
    For i = 1 To tubos_conexoes_agua.rows - 1
        For J = 1 To GrdNotaCliente.rows - 1
            If Format(tubos_conexoes_agua.TextMatrix(i, 1), "0000") = GrdNotaCliente.TextMatrix(J, 0) And Format(tubos_conexoes_agua.TextMatrix(i, 6), "##0.00") <> GrdNotaCliente.TextMatrix(J, 6) Then
                tubos_conexoes_agua.TextMatrix(i, 6) = GrdNotaCliente.TextMatrix(J, 6)
                tubos_conexoes_agua.TextMatrix(i, 7) = Format$((tubos_conexoes_agua.TextMatrix(i, 5) - ((tubos_conexoes_agua.TextMatrix(i, 5) * GrdNotaCliente.TextMatrix(J, 6)) / 100)), "Currency")
                tubos_conexoes_agua.TextMatrix(i, 4) = GrdNotaCliente.TextMatrix(J, 3)
                tubos_conexoes_agua.TextMatrix(i, 8) = Format$(GrdNotaCliente.TextMatrix(J, 9), "Currency")
            End If
        Next
    Next
    
End Sub

Private Sub carregaResumo(row As Integer, grid As MSFlexGrid)
    
    Dim dlPesBru As Double
    Dim dlDesc   As Double
    Dim ilindAux As Integer
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    If ilCodCli = 0 Then
        MsgBox "Favor selecionar o cliente antes de iniciar a simulação", vbExclamation + vbOKOnly, "Atenção!"
        grid.TextMatrix(row, 6) = 0
        grid.TextMatrix(row, 4) = 0
        grid.TextMatrix(row, 7) = grid.TextMatrix(row, 5)
        grid.TextMatrix(row, 8) = Format$(0, "Currency")
        Exit Sub
    End If
    
    If CboCondPag.codigo = "" Then
        MsgBox "Favor selecionar a forma de pagamento antes de iniciar a simulação", vbExclamation + vbOKOnly, "Atenção!"
        grid.TextMatrix(row, 6) = 0
        grid.TextMatrix(row, 4) = 0
        grid.TextMatrix(row, 7) = grid.TextMatrix(row, 5)
        grid.TextMatrix(row, 8) = Format$(0, "Currency")
        Exit Sub
    Else
        ilCodCnd = CboCondPag.codigo
    End If
    
    Dim i As Integer
    Dim rowsResumo As Integer
    
    rowsResumo = GrdNotaCliente.rows
    
    If rowsResumo > 1 Then
        For i = 1 To GrdNotaCliente.rows - 1
            If GrdNotaCliente.TextMatrix(i, 0) = Format(grid.TextMatrix(row, 1), "0000") And GrdNotaCliente.TextMatrix(i, 6) = Format(grid.TextMatrix(row, 6), "##0.00") And GrdNotaCliente.TextMatrix(i, 3) = grid.TextMatrix(row, 4) And ControleAtualizaGrid = False Then
                Exit Sub
            End If
            
            If GrdNotaCliente.TextMatrix(i, 0) = Format(grid.TextMatrix(row, 1), "0000") And grid.TextMatrix(row, 4) = 0 Then
                If GrdNotaCliente.rows = 2 Then
                    GrdNotaCliente.rows = GrdNotaCliente.rows - 1
                    LblSub.Caption = Format(0, "##,###,##0.00")
                    LblVlSimples.Caption = Format(0, "##,###,##0.00")
                    LblTot.Caption = Format(0, "##,###,##0.00")
                    LblDesc.Caption = Format(0, "##,###,##0.00")
                Else
                    GrdNotaCliente.RemoveItem (i)
                    CalculaIndice
                    DefineCorResumo GrdNotaCliente.row, GrdNotaCliente
                    LblSub.Caption = Format(dlTotBru, "##,###,##0.00")
                    LblVlSimples.Caption = Format(dlSimples, "##,###,##0.00")
                    LblTot.Caption = Format(dlTotLiq, "##,###,##0.00")
                    LblDesc.Caption = Format(dlTotBru - (dlTotLiq + dlSimples), "##,###,##0.00")
                End If
                'grid_resumo.Rows = grid_resumo.Rows - 1
                Exit Sub
            End If
            
            If GrdNotaCliente.TextMatrix(i, 0) = Format(grid.TextMatrix(row, 1), "0000") And GrdNotaCliente.TextMatrix(i, 6) = Format(grid.TextMatrix(row, 6), "##0.00") Then
                If GrdNotaCliente.rows = 2 Then
                    GrdNotaCliente.rows = GrdNotaCliente.rows - 1
                Else
                    GrdNotaCliente.RemoveItem (i)
                    GrdNotaCliente.Refresh
                End If
                Exit For
            End If
            
            If GrdNotaCliente.TextMatrix(i, 0) = Format(grid.TextMatrix(row, 1), "0000") And GrdNotaCliente.TextMatrix(i, 6) <> Format(grid.TextMatrix(row, 6), "##0.00") Then
                If GrdNotaCliente.rows = 2 Then
                    GrdNotaCliente.rows = GrdNotaCliente.rows - 1
                Else
                    GrdNotaCliente.RemoveItem (i)
                    GrdNotaCliente.Refresh
                End If
                Exit For
            End If
            
        Next
    End If
    
    PopulaVariaveisQuantidade
    
    '*****************************************************************************************
    'Zera as variáveis de cálculo e definição das cores. Define a comissão do representante.
    '*****************************************************************************************
    
    dlPerComiNeg = 0
    slClasCor = ""
    blFechaComi = False
    dlPerComiCalc = dlPerComiN
    
    '*****************************************************************************************
    'A lista de pedido não pode conter mais de 65 itens.
    '*****************************************************************************************
    
    If GrdNotaCliente.rows > 65 Then
    
        MsgBox "É pertimido a inclusão de até 65 itens por pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        grid.TextMatrix(row, 6) = 0
        grid.TextMatrix(row, 4) = 0
        grid.TextMatrix(row, 7) = grid.TextMatrix(row, 5)
        grid.TextMatrix(row, 8) = Format$(0, "Currency")
        
        LimpaLinhaNF
        
        ''MskNumNf.SetFocus
        
        Exit Sub
        
    End If
    
    '*****************************************************************************************
    '
    '*****************************************************************************************
    
    If Me.ActiveControl.Name = "BtoSair" Or Me.ActiveControl.Name = "BtoLimpaCTRC" Then
        Exit Sub
    End If
    
    '*****************************************************************************************
    'Faz consistência do campo "Desconto". Se estiver vazio, recebe um valor numérico (zero).
    '*****************************************************************************************
    
    If Trim(MskDatEmiNf) = "" Then
        MskDatEmiNf = 0
    End If
    '-----------------------------------------------------------------
    If ilCodCnd = 1 Or ilCodCnd = 12 Or ilCodCnd = 24 Then 'A vista ou 14 dias
       bAVista = True
    Else
       bAVista = False
    End If

    If bAVista = True Then 'estou trabalhando a vista
            
        If Trim(MskDatEmiNf) <= dDscRegiao Then
            iDscRegiao = iDscRegiao + 1
        Else
            iDscForaRegiao = iDscForaRegiao + 1
        End If
        
     Else
    
        If Trim(MskDatEmiNf) <= dDscRegiao Then
            iDscRegiao = iDscRegiao + 1
        Else
            iDscForaRegiao = iDscForaRegiao + 1
        End If
    
    
    End If
    '*****************************************************************************************
    'Faz a consistência do campo "Código do Produto".
    '*****************************************************************************************
    
    If Trim(MskNumNf) = "" Or Trim(MskNumNf) = 0 Then
    
        MsgBox "Informe o Código do produto.", vbExclamation + vbOKOnly, "Atenção!"
        
        ''MskNumNf.SetFocus
        
        Exit Sub
        
    End If
    
    '*****************************************************************************************
    'Se o usuário indicou que o pedido é de um Kit Irrigação, o sistema verifica se o produto
    'a ser inserido na lista pode fazer parte de um kit. Se não puder, não entra na lista.
    '*****************************************************************************************
    
    If ChkKit.Value = 1 And ilFlgKit = 0 Then
    
        MsgBox "Este Produto não compõe a linha de irrigação !", vbExclamation + vbOKOnly, "Atenção!"
        grid.TextMatrix(row, 6) = 0
        grid.TextMatrix(row, 4) = 0
        grid.TextMatrix(row, 7) = grid.TextMatrix(row, 5)
        grid.TextMatrix(row, 8) = Format$(0, "Currency")
        
        Exit Sub
        
    End If
    
    '*****************************************************************************************
    'Faz a consistência do campo "Quantidade".
    '*****************************************************************************************
    
    If Trim(MskSerie) = "" Or Trim(MskSerie) = 0 Then
    
        MsgBox "Informe a Quantidade do produto.", vbExclamation + vbOKOnly, "Atenção!"
        
        'MskSerie.SetFocus
        
        Exit Sub
        
    End If
    
    '*****************************************************************************************
    'Aplica formato de moeda ao desconto e avalia se seu valor respeita o limite permitido
    'para a venda.
    '*****************************************************************************************
    
    dlDesc = Format(Trim(MskDatEmiNf), "##0.00")
    
    If dlDesc > dlSumDscItem Then
    
        
        MsgBox "Desconto informado maior que o permitido para esta venda.", vbExclamation + vbOKOnly, "Atenção!"
        
        grid.TextMatrix(row, 6) = dlSumDscItem
        
        MskDatEmiNf = dlSumDscItem
        
        AtualizaValorComDesconto row, grid
        
        CalculaTotal row, grid
        
        tab_simulacao_pedido.SetFocus
        
    End If
    
    '*****************************************************************************************
    'Verifica se o produto atual já existe na lista. Pode até haver inclusão duplicada de um
    'mesmo produto, porém a quantidade informada deve ser diferente daquela que já foi
    'inserida.
    '*****************************************************************************************
    
    If blModificar = False Then
    
        For Linhas = 1 To GrdNotaCliente.rows - 1
        
            If GrdNotaCliente.TextMatrix(Linhas, 0) = Format(Trim(MskNumNf), "0000") And GrdNotaCliente.TextMatrix(Linhas, 6) = Format(Trim(MskSerie), "##0.00") Then
            
                If GrdNotaCliente.rows = 2 Then
                    GrdNotaCliente.rows = GrdNotaCliente.rows - 1
                Else
                    GrdNotaCliente.RemoveItem (Linhas)
                    GrdNotaCliente.Refresh
                End If
                Exit For
                
            End If
            
        Next Linhas
        
    End If
    
    '*****************************************************************************************
    'A rotina a seguir insere os itens do pedido no grid.
    '*****************************************************************************************
    
    '*****************************************************************************************
    'Segue abaixo uma lista com a posição de cada informação no grid:
    '0: Código do produto
    '1: Descrição do produto
    '2: Quantidade vendida por embalagem
    '3: Quantidade vendida
    '4: Preço unitário (bruto ou com desconto?)
    '5: Identificação da tabela aplicada (vazio para tabela normal, A e B)
    '6: Percentual de desconto aplicado para o produto
    '7: Valor total líqüido do produto
    '8: Código do grupo de produto
    '9: Valor total bruto do produto
    '10: Peso bruto do produto
    '11: Preço ideal da venda
    '12: Preço unitário bruto
    '13: Margem de lucro
    '14: Kit Irrigação?
    '15: Custo unitário de aquisição do produto
    '16: Custo unitário adicional de aquisição do produto
    '17: Alíquota de imposto federal
    '18: Demonstração gráfica da margem do produto. Margem individual.
    '*****************************************************************************************
    
    '*************************************************************************************
    'Inclui o produto na lista.
    '*************************************************************************************
    
    ilind = GrdNotaCliente.rows
    
    GrdNotaCliente.rows = GrdNotaCliente.rows + 1
    GrdNotaCliente.TextMatrix(ilind, 0) = Format(Trim(MskNumNf), "0000")
    GrdNotaCliente.TextMatrix(ilind, 1) = Trim(LblRotaRec.Caption)
    GrdNotaCliente.TextMatrix(ilind, 2) = ilQtdEmb
    GrdNotaCliente.TextMatrix(ilind, 3) = IIf(Trim(MskSerie) = "", "0", MskSerie)
    GrdNotaCliente.TextMatrix(ilind, 4) = IIf(Trim(MskVlrUnit) = "", "0", Format(MskVlrUnit, "##,###,##0.00"))
    
    '*************************************************************************************
    'Identifica a tabela aplicada: vazio para tabela normal ou tabelas A e B.
    '*************************************************************************************
    
    If ilNumTab = 0 Then
        
        GrdNotaCliente.TextMatrix(ilind, 5) = ""
        
    Else
    
        If ilNumTab = 1 Then
            GrdNotaCliente.TextMatrix(ilind, 5) = "A"
        Else
            GrdNotaCliente.TextMatrix(ilind, 5) = "B"
        End If
        
    End If
    
    GrdNotaCliente.TextMatrix(ilind, 6) = IIf(Trim(MskDatEmiNf) = "", "0", Format(MskDatEmiNf, "##0.00"))
    
    '*************************************************************************************
    'Se houver desconto aplicado ao produto, o valor total do item será o preço bruto,
    'menos o desconto aplicado, vezes a quantidade vendida. Se não houver desconto, o
    'valor total do item será o preço bruto vezes a quantidade vendida.
    '*************************************************************************************
    
    If Trim(MskDatEmiNf) = 0 Then
        dlValItem = Trim(MskVlrUnit) * (Trim(MskSerie))
    Else
        dlValItem = (MskVlrUnit - ((MskVlrUnit * MskDatEmiNf) / 100)) * Trim(MskSerie)
    End If
    
    GrdNotaCliente.TextMatrix(ilind, 7) = Format((GrdNotaCliente.TextMatrix(ilind, 4) - ((GrdNotaCliente.TextMatrix(ilind, 4) * MskDatEmiNf) / 100)), "##,###,##0.00")
    GrdNotaCliente.TextMatrix(ilind, 18) = Format(dlValItem, "##,###,##0.00")
    GrdNotaCliente.TextMatrix(ilind, 8) = ilIdeGrp
    GrdNotaCliente.TextMatrix(ilind, 9) = Format(Trim(MskVlrUnit) * (Trim(MskSerie)), "##,###,##0.00")
    
    '*************************************************************************************
    'O cálculo do peso bruto do produto é simples: multiplica-se o peso unitário da peça
    'pelo resultado da quantidade de embalagens vendidas vezes a quantidade de produtos
    'que compõe cada embalagem.
    '*************************************************************************************
    
    dlPesBru = dlPesUnt * (Trim(MskSerie))
    
    GrdNotaCliente.TextMatrix(ilind, 10) = Format(dlPesBru, "###,##0.0000")
    
    '*************************************************************************************
    'A linha a seguir calcula aquilo que foi chamado de "valor ideal" do produto: trata-se
    'do preço total líqüido. Importante frisar que o desconto é a soma de todos aqueles
    'encontrados para o perfil da venda, independente de haver chave para o pedido ou não.
    '*************************************************************************************
    
    VlIdealItem = (dlValUntN - ((dlValUntN * dlSumDscItemORIG) / 100)) * (Trim(MskSerie))
    
    GrdNotaCliente.TextMatrix(ilind, 11) = Format(VlIdealItem, "###,##0.00")
    GrdNotaCliente.TextMatrix(ilind, 12) = Format(dlValUntN, "###,##0.00")
    GrdNotaCliente.TextMatrix(ilind, 13) = Format(dlMrgPrd, "##0.000")
    GrdNotaCliente.TextMatrix(ilind, 14) = ilFlgKit
    GrdNotaCliente.TextMatrix(ilind, 15) = Format(dlValCusUntQtd, "###,##0.00")
    GrdNotaCliente.TextMatrix(ilind, 16) = Format(dlValCusAdicQtd, "###,##0.00")
    GrdNotaCliente.TextMatrix(ilind, 17) = Format(dlAlqImpFed, "###,##0.000")
    
    '*****************************************************************************************
    'Limpra os campos por onde o produto foi inserido.
    '*****************************************************************************************
    
    LimpaLinhaNF
    
    '*****************************************************************************************
    'Habilita os campos para uma possível inserção de novo produto. Também desabilita o campo
    'que permite a marcação de Kit Irrigação; como o primeiro item já foi aceito, o pedido
    'atual já foi definido como Kit e não pode mais deixar esse status.
    '*****************************************************************************************
    
    '''MskNumNf.Enabled = True
    ''BtoProduto.Enabled = True
    'MskSerie.Enabled = True
    ''Bto_Aplica.Enabled = True
    ''BtoExcNF.Enabled = False
    ''BtoAdiNF.Enabled = True
    'MskNumNf.SetFocus
    
    blModificar = False
    
    If bgSimula = False Then
        BtoGrava.Enabled = True
    End If
    
    GrdNotaCliente.Enabled = True
    
    ilNumTab = 0
    
    LblResultNegocio.Caption = ""
    ChkKit.Enabled = False
    
    ilindAux = ilind
    '*****************************************************************************************
    'Calcula os índices do pedido. Se houver qualquer problema durante o cálculo, todo o
    'formulário será limpo e o pedido anulado.
    '*****************************************************************************************
    
    If CalculaIndice = False Then
        
        LimpaGeral
        
    Else
    
    'If MskDatEmiNf > GrdNotaCliente.TextMatrix(ilindAux, 6) Then
        'grid.TextMatrix(row, 6) = GrdNotaCliente.TextMatrix(ilindAux, 6)
        'grid.TextMatrix(row, 7) = (grid.TextMatrix(row, 5) - ((grid.TextMatrix(row, 5) * grid.TextMatrix(row, 6)) / 100))
        'grid.TextMatrix(row, 7) = Format$(grid.TextMatrix(row, 7), "Currency")
        'CalculaTotal row, grid
        
        'MsgBox "Desconto aplicado de " & MskDatEmiNf & "% é maior que o permitido para o perfil da venda", vbExclamation + vbOKOnly, "Atenção!"
        
        'tab_simulacao_pedido.SetFocus
        
        'ControleLostFocus = False

    'End If
    
        ilind = GrdNotaCliente.rows - 1
        
        GrdNotaCliente.col = 19
        GrdNotaCliente.row = ilind
    
        If GrdIndice.TextMatrix(ilind, 15) < 6 Then
            GrdNotaCliente.CellBackColor = &HFF&
        ElseIf GrdIndice.TextMatrix(ilind, 15) >= 6 And GrdIndice.TextMatrix(ilind, 15) < 8.6 Then
            GrdNotaCliente.CellBackColor = &H80FFFF 'Amarelo
        ElseIf GrdIndice.TextMatrix(ilind, 15) >= 8.6 And GrdIndice.TextMatrix(ilind, 15) <= 12 Then
            GrdNotaCliente.CellBackColor = &HFF00&
        ElseIf GrdIndice.TextMatrix(ilind, 15) > 12 Then
            GrdNotaCliente.CellBackColor = &HFF0000
        End If
        
    End If
    
    GrdNotaCliente.col = 1
    GrdNotaCliente.Sort = flexSortGenericAscending
    GrdNotaCliente.Refresh
End Sub

Private Sub DefineCorResumo(row As Integer, grid As MSFlexGrid)
    
    Dim ilind As Integer
    
    ilind = row
        
    grid.col = 19
    grid.row = ilind
        
    If GrdIndice.TextMatrix(ilind, 15) < 6 Then
        grid.CellBackColor = &HFF&
    ElseIf GrdIndice.TextMatrix(ilind, 15) >= 6 And GrdIndice.TextMatrix(ilind, 15) < 8.6 Then
        grid.CellBackColor = &H80FFFF 'Amarelo
    ElseIf GrdIndice.TextMatrix(ilind, 15) >= 8.6 And GrdIndice.TextMatrix(ilind, 15) <= 12 Then
        grid.CellBackColor = &HFF00&
    ElseIf GrdIndice.TextMatrix(ilind, 15) > 12 Then
        grid.CellBackColor = &HFF0000
    End If
    
    grid.Refresh
    
End Sub

Private Sub AtualizaValorComDescontoResumo(row As Integer, grid As MSFlexGrid)

    Dim valueCel As Double
    'Recupera desconto digitado
    valueCel = grid.TextMatrix(row, 6)
    
    'If valueCel > 34.2 Then
        'MsgBox "Desconto não pode ser maior que 34.2%"
        'grid.TextMatrix(row, 6) = 34.2
        'Exit Sub
    'End If
    
    'Atribui valor ao Preço Unitário
    grid.TextMatrix(row, 7) = Format((grid.TextMatrix(row, 4) - ((grid.TextMatrix(row, 4) * valueCel) / 100)), "##,###,##0.00")
    'grid.TextMatrix(row, 7) = Format$(grid.TextMatrix(row, 7), "Currency")
    
End Sub

Private Sub AtualizaValorComDesconto(row As Integer, grid As MSFlexGrid)
    Dim valueCel As Double
    'Recupera desconto digitado
    valueCel = grid.TextMatrix(row, 6)
    
    'If valueCel > 34.2 Then
        'MsgBox "Desconto não pode ser maior que 34.2%"
        'grid.TextMatrix(row, 6) = 34.2
        'Exit Sub
    'End If
    
    'Atribui valor ao Preço Unitário
    grid.TextMatrix(row, 7) = (grid.TextMatrix(row, 5) - ((grid.TextMatrix(row, 5) * valueCel) / 100))
    grid.TextMatrix(row, 7) = Format$(grid.TextMatrix(row, 7), "Currency")
    
End Sub

Private Sub tubos_conexoes_predial_SelChange()
    
    If tubos_conexoes_predial.col = 1 Then
        If tubos_conexoes_predial.SelectionMode = flexSelectionFree Then
            tubos_conexoes_predial.SelectionMode = flexSelectionByRow
            tubos_conexoes_predial.Refresh
        End If
    End If
    
    auxChangeTuboConexoesPredial = True
    If auxSelChangeTuboConexoesPredial = True Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto rowAuxTuboConexoesPredial, tubos_conexoes_predial
        If houveDigitacaoTuboConexoesPredial = True Or tubos_conexoes_predial.TextMatrix(rowAuxTuboConexoesPredial, 4) > 0 Then
            PopularVariaveisCalculo rowAuxTuboConexoesPredial, tubos_conexoes_predial
            CalculaTotal rowAuxTuboConexoesPredial, tubos_conexoes_predial
            carregaResumo rowAuxTuboConexoesPredial, tubos_conexoes_predial
            houveDigitacaoTuboConexoesPredial = False
            auxSelChangeTuboConexoesPredial = False
        End If
    End If
End Sub

Private Sub tubos_conexoes_roscaveis_SelChange()
    
    If tubos_conexoes_roscaveis.col = 1 Then
        If tubos_conexoes_roscaveis.SelectionMode = flexSelectionFree Then
            tubos_conexoes_roscaveis.SelectionMode = flexSelectionByRow
            tubos_conexoes_roscaveis.Refresh
        End If
    End If
    
    auxChangeTubosConexoesRoscaveis = True
    If auxSelChangeTubosConexoesRoscaveis = True Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto rowAuxTubosConexoesRoscaveis, tubos_conexoes_roscaveis
        If houveDigitacaoTubosConexoesRoscaveis = True Or tubos_conexoes_roscaveis.TextMatrix(rowAuxTubosConexoesRoscaveis, 4) > 0 Then
            PopularVariaveisCalculo rowAuxTubosConexoesRoscaveis, tubos_conexoes_roscaveis
            CalculaTotal rowAuxTubosConexoesRoscaveis, tubos_conexoes_roscaveis
            carregaResumo rowAuxTubosConexoesRoscaveis, tubos_conexoes_roscaveis
            houveDigitacaoTubosConexoesRoscaveis = False
            auxSelChangeTubosConexoesRoscaveis = False
        End If
    End If
End Sub

Private Sub GrdNotaCliente_SelChange()

    If GrdNotaCliente.col = 1 Then
        If GrdNotaCliente.SelectionMode = flexSelectionFree Then
            GrdNotaCliente.SelectionMode = flexSelectionByRow
            GrdNotaCliente.Refresh
        End If
    End If
    
    auxChangeGrdNotaCliente = True
    If auxSelChangeGrdNotaCliente = True Then
        titleTab = tab_simulacao_pedido.Caption
        'AtualizaValorComDescontoResumo rowAuxGrdNotaCliente, GrdNotaCliente
        If houveDigitacaoGrdNotaCliente = True Or GrdNotaCliente.TextMatrix(rowAuxGrdNotaCliente, 3) > 0 Then
            PopularVariaveisCalculoResumo rowAuxGrdNotaCliente, GrdNotaCliente
            'CalculaTotalResumo rowAuxGrdNotaCliente, GrdNotaCliente
            'carregaResumo rowAuxTubosConexoesIrriAzuis, tubos_conexoes_irri_azuis
            'CalculaIndice
            AtualizaGridAuxiliar rowAuxGrdNotaCliente
            DefineCorResumo rowAuxGrdNotaCliente, GrdNotaCliente
            houveDigitacaoGrdNotaCliente = False
            auxSelChangeGrdNotaCliente = False
        End If
    End If
    
End Sub

Private Sub tubos_conexoes_irri_azuis_SelChange()
    
    If tubos_conexoes_irri_azuis.col = 1 Then
        If tubos_conexoes_irri_azuis.SelectionMode = flexSelectionFree Then
            tubos_conexoes_irri_azuis.SelectionMode = flexSelectionByRow
            tubos_conexoes_irri_azuis.Refresh
        End If
    End If
    
    auxChangeTubosConexoesIrriAzuis = True
    If auxSelChangeTubosConexoesIrriAzuis = True Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto rowAuxTubosConexoesIrriAzuis, tubos_conexoes_irri_azuis
        If houveDigitacaoTubosConexoesIrriAzuis = True Or tubos_conexoes_irri_azuis.TextMatrix(rowAuxTubosConexoesIrriAzuis, 4) > 0 Then
            PopularVariaveisCalculo rowAuxTubosConexoesIrriAzuis, tubos_conexoes_irri_azuis
            CalculaTotal rowAuxTubosConexoesIrriAzuis, tubos_conexoes_irri_azuis
            carregaResumo rowAuxTubosConexoesIrriAzuis, tubos_conexoes_irri_azuis
            houveDigitacaoTubosConexoesIrriAzuis = False
            auxSelChangeTubosConexoesIrriAzuis = False
        End If
    End If
End Sub

Private Sub tubo_conexoes_defofo_SelChange()
    
    If tubo_conexoes_defofo.col = 1 Then
        If tubo_conexoes_defofo.SelectionMode = flexSelectionFree Then
            tubo_conexoes_defofo.SelectionMode = flexSelectionByRow
            tubo_conexoes_defofo.Refresh
        End If
    End If
    
    auxChange = True
    If auxSelChange = True Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto rowAux, tubo_conexoes_defofo
        If houveDigitacaoEsgoto = True Or tubo_conexoes_defofo.TextMatrix(rowAux, 4) > 0 Then
            PopularVariaveisCalculo rowAux, tubo_conexoes_defofo
            CalculaTotal rowAux, tubo_conexoes_defofo
            carregaResumo rowAux, tubo_conexoes_defofo
            auxSelChange = False
            houveDigitacaoEsgoto = False
        End If
    End If
End Sub

Private Sub tubo_tubos_conexoes_coletor_esgoto_SelChange()
    
    If tubo_tubos_conexoes_coletor_esgoto.col = 1 Then
        If tubo_tubos_conexoes_coletor_esgoto.SelectionMode = flexSelectionFree Then
            tubo_tubos_conexoes_coletor_esgoto.SelectionMode = flexSelectionByRow
            tubo_tubos_conexoes_coletor_esgoto.Refresh
        End If
    End If
    
    auxChangeSaldaveisAgua = True
    If auxSelChangeSaldaveisAgua = True Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto rowAuxSaldaveisAgua, tubo_tubos_conexoes_coletor_esgoto
        If houveDigitacaoSoldaveisAgua = True Or tubo_tubos_conexoes_coletor_esgoto.TextMatrix(rowAuxSaldaveisAgua, 4) > 0 Then
            PopularVariaveisCalculo rowAuxSaldaveisAgua, tubo_tubos_conexoes_coletor_esgoto
            CalculaTotal rowAuxSaldaveisAgua, tubo_tubos_conexoes_coletor_esgoto
            carregaResumo rowAuxSaldaveisAgua, tubo_tubos_conexoes_coletor_esgoto
            auxSelChangeSaldaveisAgua = False
            houveDigitacaoSoldaveisAgua = False
        End If
    End If
End Sub

Private Sub tubo_conexoes_pba_SelChange()
    
    If tubo_conexoes_pba.col = 1 Then
        If tubo_conexoes_pba.SelectionMode = flexSelectionFree Then
            tubo_conexoes_pba.SelectionMode = flexSelectionByRow
            tubo_conexoes_pba.Refresh
        End If
    End If
    
    auxChangeRoscaveis = True
    If auxSelChangeRoscaveis = True Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto rowAuxRoscaveis, tubo_conexoes_pba
        If houveDigitacaoTubosRoscaveis = True Or tubo_conexoes_pba.TextMatrix(rowAuxRoscaveis, 4) > 0 Then
            PopularVariaveisCalculo rowAuxRoscaveis, tubo_conexoes_pba
            CalculaTotal rowAuxRoscaveis, tubo_conexoes_pba
            carregaResumo rowAuxRoscaveis, tubo_conexoes_pba
            auxSelChangeRoscaveis = False
            houveDigitacaoTubosRoscaveis = False
        End If
    End If
End Sub

Private Sub tubos_conexoes_agua_SelChange()
    
    If tubos_conexoes_agua.col = 1 Then
        If tubos_conexoes_agua.SelectionMode = flexSelectionFree Then
            tubos_conexoes_agua.SelectionMode = flexSelectionByRow
            tubos_conexoes_agua.Refresh
        End If
    End If
    
    auxChangeTuboConexoesAgua = True
    If auxSelChangeTuboConexoesAgua = True Then
        titleTab = tab_simulacao_pedido.Caption
        AtualizaValorComDesconto rowAuxTuboConexoesAgua, tubos_conexoes_agua
        If houveDigitacaoTuboConexoesAgua = True Or tubos_conexoes_agua.TextMatrix(rowAuxTuboConexoesAgua, 4) > 0 Then
            PopularVariaveisCalculo rowAuxTuboConexoesAgua, tubos_conexoes_agua
            CalculaTotal rowAuxTuboConexoesAgua, tubos_conexoes_agua
            carregaResumo rowAuxTuboConexoesAgua, tubos_conexoes_agua
            auxSelChangeTuboConexoesAgua = False
            houveDigitacaoTuboConexoesAgua = False
        End If
    End If
End Sub

Private Sub CarregaTemporaria()

    Dim sql_tmp As String
    
    'CRIANDO TABELA TEMPORÁRIA DE PRODUTOS
    sql_tmp = "INSERT INTO tmp_produto ("
    sql_tmp = sql_tmp & " CODIGO,"
    sql_tmp = sql_tmp & " PRODUTO,"
    sql_tmp = sql_tmp & " EMBALAGEM,"
    sql_tmp = sql_tmp & " QUANTIDADE,"
    sql_tmp = sql_tmp & " TABELA,"
    sql_tmp = sql_tmp & " DESCONTO,"
    sql_tmp = sql_tmp & " PRECO_UNITARIO,"
    sql_tmp = sql_tmp & " VALOR_TOTAL,"
    sql_tmp = sql_tmp & " SITUACAO,"
    sql_tmp = sql_tmp & " CUSTO,"
    sql_tmp = sql_tmp & " CUSTO_TOTAL,"
    sql_tmp = sql_tmp & " IdeGrp,"
    sql_tmp = sql_tmp & " PesUnt,"
    sql_tmp = sql_tmp & " MrgPrd,"
    sql_tmp = sql_tmp & " valcusuntqtd,"
    sql_tmp = sql_tmp & " valcusadicqtd,"
    sql_tmp = sql_tmp & " AlqImpFed"
    sql_tmp = sql_tmp & " )"
    sql_tmp = sql_tmp & " SELECT a.Codprd, a.Dscprd, a.QtdEmb, 0, b.ValUntN, 0 , b.ValUntN, 0.00, 0, 0.0, 0.0,a.IdeGrp,a.PesUnt,b.MrgPrd,b.valcusuntqtd, b.valcusadicqtd,b.AlqImpFed  from PRODUTO a, PRECO_PRODUTO b"
    sql_tmp = sql_tmp & " WHERE a.flgsitu = 'N'"
    sql_tmp = sql_tmp & " and a.codprd = b.codprd"
    sql_tmp = sql_tmp & " and b.DatATiv = (SELECT MAX(DatAtiv) FROM PRECO_PRODUTO p WHERE p.CodPrd = a.CodPrd)"
    sql_tmp = sql_tmp & " and NOT EXISTS (SELECT 1 FROM tmp_produto p WHERE p.CODIGO = a.codprd)"
    'sql_tmp = sql_tmp & " and (a.DSCPRD LIKE 'TB%' OR a.DSCPRD LIKE 'TU%')"
    sql_tmp = sql_tmp & " order by a.Dscprd"

    Conexao.Execute sql_tmp
    
    slUFOri = "MG"
    slUFRep = ""
    slNomRep = ""
    dlPerCusFrt = 0
    dlPerComiN = 0
    dlPerComiA = 0
    dlPerComiB = 0
    dlPerDesFOB = 0
    dlIdxPDD = 0
    dlIdxAzul = 0
    slFlgSugComi = ""
    dlValPedMin = 0
    dlValLimPrz1 = 0
    dlValParMin = 0
    ilPrzMed1 = 0
    ilPrzMed2 = 0
   
    sgQuery = "select * from REPRESENTANTE where CodRep = " & Trim(sgRepresentante)
    
    Call Consulta2(sgQuery)
    
    If Not Rs2.EOF Then
    
        slNomRep = IIf(IsNull(Rs2!NomRep), "", Trim(Rs2!NomRep))
        slUFRep = IIf(IsNull(Rs2!UFRep), "", Trim(Rs2!UFRep))
        dlPerCusFrt = IIf(IsNull(Rs2!PerCusFrt), 0, Trim(Rs2!PerCusFrt))
        dlPerComiN = IIf(IsNull(Rs2!PerComiN), 0, Trim(Rs2!PerComiN))
        dlPerComiA = IIf(IsNull(Rs2!PerComiA), 0, Trim(Rs2!PerComiA))
        dlPerComiB = IIf(IsNull(Rs2!PerComiB), 0, Trim(Rs2!PerComiB))
        dlPerDesFOB = IIf(IsNull(Rs2!PerDesFOB), 0, Trim(Rs2!PerDesFOB))
        slFlgSugComi = IIf(IsNull(Rs2!FlgSugComi), "", Trim(Rs2!FlgSugComi))
        dlIdxPDD = IIf(IsNull(Rs2!IdxPDD), 0, Trim(Rs2!IdxPDD))
        dlIdxAzul = IIf(IsNull(Rs2!IdxAzul), 0, Trim(Rs2!IdxAzul))
        dlPerComiCalc = dlPerComiN
        dlPerTubo100Rep = IIf(IsNull(Rs2!PerTubo100), 0, Trim(Rs2!PerTubo100))
        dlValPedMin = IIf(IsNull(Rs2!ValPedMin), 0, Trim(Rs2!ValPedMin))
        dlValLimPrz1 = IIf(IsNull(Rs2!ValLimPrz1), 0, Trim(Rs2!ValLimPrz1))
        dlValParMin = IIf(IsNull(Rs2!ValParMin), 0, Trim(Rs2!ValParMin))
        ilPrzMed1 = IIf(IsNull(Rs2!PrzMed1), 0, Trim(Rs2!PrzMed1))
        ilPrzMed2 = IIf(IsNull(Rs2!PrzMed2), 0, Trim(Rs2!PrzMed2))
    
    Else
    
        MsgBox "Registro do Representante não encontrado, informe ao administrador do sistema", vbExclamation + vbOKOnly, "Atenção!"
    
    End If
  
    Rs2.Close
    
    Set Rs2 = Nothing
   
    'leitura BALIZA_SUGESTAO (situação [1] Abaixo do limite)
    PerSug1Ini = 0
    PerSug1FimA = 0
    PerSug1FimB = 0

    sgQuery = "select * from BALIZA_SUGESTAO where CodRep = " & Trim(ilCodRep) & " and NroSit = 1 "
    
    Call Consulta2(sgQuery)
    
    If Not Rs2.EOF Then
        
        PerSug1Ini = IIf(IsNull(Rs2!PerSugIni), 0, Trim(Rs2!PerSugIni))
        PerSug1FimA = IIf(IsNull(Rs2!PerSugFimA), 0, Trim(Rs2!PerSugFimA))
        PerSug1FimB = IIf(IsNull(Rs2!PerSugFimB), 0, Trim(Rs2!PerSugFimB))
        
    Else
    
        Rs2.Close
        
        Set Rs2 = Nothing
        
        sgQuery = "select * from BALIZA_SUGESTAO where CodRep is null and NroSit = 1"
        
        Call Consulta2(sgQuery)
        
        If Not Rs2.EOF Then
            PerSug1Ini = IIf(IsNull(Rs2!PerSugIni), 0, Trim(Rs2!PerSugIni))
            PerSug1FimA = IIf(IsNull(Rs2!PerSugFimA), 0, Trim(Rs2!PerSugFimA))
            PerSug1FimB = IIf(IsNull(Rs2!PerSugFimB), 0, Trim(Rs2!PerSugFimB))
        End If
        
    End If
    
    Rs2.Close
    
    Set Rs2 = Nothing
   
    'leitura BALIZA_SUGESTAO (situação [2] Acima do limite)
    PerSug2Ini = 0
    PerSug2FimA = 0
    PerSug2FimB = 0
   
    sgQuery = "select * from BALIZA_SUGESTAO where CodRep = " & Trim(ilCodRep) & " and NroSit = 2 "
    
    Call Consulta2(sgQuery)
    
    If Not Rs2.EOF Then
        
        PerSug2Ini = IIf(IsNull(Rs2!PerSugIni), 0, Trim(Rs2!PerSugIni))
        PerSug2FimA = IIf(IsNull(Rs2!PerSugFimA), 0, Trim(Rs2!PerSugFimA))
        PerSug2FimB = IIf(IsNull(Rs2!PerSugFimB), 0, Trim(Rs2!PerSugFimB))
        
    Else
    
        Rs2.Close
        
        Set Rs2 = Nothing
        
        sgQuery = "select * from BALIZA_SUGESTAO where CodRep is null and NroSit = 2"
        
        Call Consulta2(sgQuery)
        
        If Not Rs2.EOF Then
            PerSug2Ini = IIf(IsNull(Rs2!PerSugIni), 0, Trim(Rs2!PerSugIni))
            PerSug2FimA = IIf(IsNull(Rs2!PerSugFimA), 0, Trim(Rs2!PerSugFimA))
            PerSug2FimB = IIf(IsNull(Rs2!PerSugFimB), 0, Trim(Rs2!PerSugFimB))
        End If
    
    End If
    
    Rs2.Close
    
    Set Rs2 = Nothing
    
End Sub

Private Sub CarregaGrid(sqlAux As String, grid As MSFlexGrid)
    
    Dim sql As String

    sql = " SELECT "
    sql = sql & " CODIGO,"
    sql = sql & " PRODUTO,"
    sql = sql & " EMBALAGEM,"
    sql = sql & " QUANTIDADE,"
    sql = sql & " TABELA,"
    sql = sql & " DESCONTO,"
    sql = sql & " PRECO_UNITARIO,"
    sql = sql & " VALOR_TOTAL,"
    sql = sql & " SITUACAO,"
    sql = sql & " CUSTO,"
    sql = sql & " CUSTO_TOTAL,"
    sql = sql & " IdeGrp,"
    sql = sql & " PesUnt,"
    sql = sql & " MrgPrd,"
    sql = sql & " valcusuntqtd,"
    sql = sql & " valcusadicqtd,"
    sql = sql & " AlqImpFed"
    sql = sql & " FROM tmp_produto WHERE CODIGO IN ( " & sqlAux & ")"
    sql = sql & " AND (PRODUTO LIKE 'TB%' OR PRODUTO LIKE 'TU%') ORDER BY SUBSTRING(PRODUTO,CASE WHEN CODIGO = 314 THEN 2 ELSE 1 END,CASE WHEN PATINDEX('%[0-9]%',PRODUTO) = 0 THEN LEN(PRODUTO) ELSE PATINDEX('%[0-9]%',PRODUTO)-1 END),SUBSTRING(PRODUTO,PATINDEX('%º%',PRODUTO)-2, 3), CAST(dbo.udf_GetNumeric(produto) AS BIGINT) asc"
    
    Call Consulta(sql)

    'define o numero de linhas e colunas e configura o grid
    grid.rows = Rs.RecordCount + 1
    grid.Cols = Rs.Fields.Count + 1
    grid.row = 1
    grid.col = 1
    grid.RowSel = grid.rows - 1
    grid.ColSel = grid.Cols - 1
    
    'estamos usando a propriedade Clip e o método GetString para selecionar uma região do grid
    grid.Clip = Rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    grid.row = 1
    grid.Visible = True
    
    sql = " SELECT "
    sql = sql & " CODIGO,"
    sql = sql & " PRODUTO,"
    sql = sql & " EMBALAGEM,"
    sql = sql & " QUANTIDADE,"
    sql = sql & " TABELA,"
    sql = sql & " DESCONTO,"
    sql = sql & " PRECO_UNITARIO,"
    sql = sql & " VALOR_TOTAL,"
    sql = sql & " SITUACAO,"
    sql = sql & " CUSTO,"
    sql = sql & " CUSTO_TOTAL,"
    sql = sql & " IdeGrp,"
    sql = sql & " PesUnt,"
    sql = sql & " MrgPrd,"
    sql = sql & " valcusuntqtd,"
    sql = sql & " valcusadicqtd,"
    sql = sql & " AlqImpFed"
    sql = sql & " FROM tmp_produto a WHERE a.CODIGO IN ( " & sqlAux & ")"
    sql = sql & " and NOT EXISTS (SELECT * FROM tmp_produto p"
    sql = sql & " WHERE (p.PRODUTO LIKE 'TB%' or p.PRODUTO LIKE 'TU%') AND p.CODIGO = a.CODIGO)"
    sql = sql & " order by SUBSTRING(PRODUTO,CASE WHEN CODIGO = 314 THEN 2 ELSE 1 END,CASE WHEN PATINDEX('%[0-9]%',PRODUTO) = 0 THEN LEN(PRODUTO) ELSE PATINDEX('%[0-9]%',PRODUTO)-1 END), SUBSTRING(PRODUTO,PATINDEX('%º%',PRODUTO)-2, 3), CAST(dbo.udf_GetNumeric(a.produto) AS BIGINT) asc"
    
    Call Consulta(sql)

    'define o numero de linhas e colunas e configura o grid
    Dim rows As Integer
    rows = grid.rows
    
    grid.rows = Rs.RecordCount + rows
    grid.Cols = Rs.Fields.Count + 1
    grid.row = rows
    grid.col = 1
    grid.RowSel = grid.rows - 1
    grid.ColSel = grid.Cols - 1
    
    'estamos usando a propriedade Clip e o método GetString para selecionar uma região do grid
    grid.Clip = Rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    'grid.row = 1
    grid.Visible = True
    
    grid.Refresh
    
End Sub

Private Sub DefineClassificacao(grid As MSFlexGrid)
    Dim valorTotal As Double
    Dim x As Integer
    Dim sgQuery As String
    
    sgQuery = "SELECT b.PerCusFin, b.PerDesCnd, a.QtdParCnd, a.PrzMed "
    sgQuery = sgQuery + " from CONDICAO a, CUSTO_CONDICAO b "
    sgQuery = sgQuery + "  Where a.CodCnd = 4"
    sgQuery = sgQuery + "    and a.codcnd = b.codcnd"
    sgQuery = sgQuery + "    and b.datativ = (select max(datativ) from CUSTO_CONDICAO"
    sgQuery = sgQuery + "                      Where Codcnd = b.codcnd"
    sgQuery = sgQuery + "                        and datativ <= convert(datetime,GETDATE(),103))"
    
    Call Consulta(sgQuery)
    
    If Rs.EOF Then
    
        MsgBox "Erro na leitura da Condição de Pagamento", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        LimpaGeral
        
    Else
    
        dlPerCusFin = IIf(IsNull(Rs!PerCusFin), 0, Trim(Rs!PerCusFin))
        dlPerDesCnd = IIf(IsNull(Rs!PerDesCnd), 0, Trim(Rs!PerDesCnd))
        ilQtdParCnd = IIf(IsNull(Rs!QtdParCnd), 0, Trim(Rs!QtdParCnd))
        ilPrzMed = IIf(IsNull(Rs!PrzMed), 0, Trim(Rs!PrzMed))
    
    End If
    
    Rs.Close
    
    Set Rs = Nothing
    
    CalculaMargem grid
    
    For x = 1 To grid.rows - 1
        
        grid.TextMatrix(x, 5) = Format$(grid.TextMatrix(x, 5), "Currency")
        grid.TextMatrix(x, 8) = Format$(grid.TextMatrix(x, 8), "Currency")
        
        grid.row = x
        grid.TextMatrix(x, 4) = 1
        grid.TextMatrix(x, 7) = (grid.TextMatrix(x, 5) - (grid.TextMatrix(x, 5) * (34.2 / 100) * 1))
        
        
        If grid.TextMatrix(x, 15) < 6 Then
        
            grid.col = 2
            grid.row = x
            grid.CellForeColor = &HFF&
            
        ElseIf grid.TextMatrix(x, 15) >= 6 And grid.TextMatrix(x, 15) < 8.5 Then
            
            grid.col = 2
            grid.row = x
            grid.CellForeColor = &HC0C0&
            
        ElseIf grid.TextMatrix(x, 15) >= 8.6 And grid.TextMatrix(x, 15) <= 12 Then
            
            grid.col = 2
            grid.row = x
            grid.CellForeColor = &HC000&
            
        ElseIf grid.TextMatrix(x, 15) > 12 Then
            
            grid.col = 2
            grid.row = x
            grid.CellForeColor = &HFF0000
            
        End If
        
        grid.TextMatrix(x, 10) = 0
        grid.TextMatrix(x, 11) = 0
        grid.TextMatrix(x, 4) = 0
        grid.TextMatrix(x, 7) = grid.TextMatrix(x, 5)
    Next
End Sub

Private Sub PopularVariaveisCalculoResumo(row As Integer, grid As MSFlexGrid)

    MskSerie = grid.TextMatrix(row, 3)       'Quantidade
    MskVlrUnit = grid.TextMatrix(row, 4)     'Valor Unitário
    MskDatEmiNf = grid.TextMatrix(row, 6)    'Desconto
    MskNumNf = grid.TextMatrix(row, 0)       'Código do produto
    
End Sub

Private Sub PopularVariaveisCalculo(row As Integer, grid As MSFlexGrid)
    On Error Resume Next
    
    If grid.TextMatrix(row, 4) > 32000 Then
       MsgBox "Quantidade excede a capacidade do sistema !", vbExclamation + vbOKOnly, "Atenção!"
       grid.TextMatrix(row, 4) = 0
       grid.TextMatrix(row, 6) = 0
    Else
        MskSerie = grid.TextMatrix(row, 4)       'Quantidade
        MskVlrUnit = grid.TextMatrix(row, 5)     'Valor Unitário
        MskDatEmiNf = grid.TextMatrix(row, 6)    'Desconto
        MskNumNf = grid.TextMatrix(row, 1)       'Código do produto
    End If
End Sub

Private Sub ConfiguraFlexGrid(grid As MSFlexGrid)

    With grid
        .ColWidth(0) = 0
        .ColWidth(1) = 500
        .ColWidth(2) = 4500
        .ColWidth(3) = 500
        .ColWidth(4) = 900
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1500
        .ColWidth(8) = 1250
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        
        .row = 0
        .col = 1
        .Text = "COD"
        .col = 2
        .Text = "PRODUTO"
        .col = 3
        .Text = "EMB"
        .col = 4
        .Text = "QUANT"
        .col = 5
        .Text = "TABELA"
        .col = 6
        .Text = "DESC%"
        .col = 7
        .Text = "PRECO LIQUIDO"
        .col = 8
        .Text = "VALOR TOTAL"
        .col = 10
        .Text = "CUSTO"
        .col = 11
        .Text = "CUSTO TOTAL"
        
        .ColAlignment(1) = 4
        .ColAlignment(2) = 1
        .ColAlignment(3) = 4
        .ColAlignment(4) = 4
        .ColAlignment(5) = 4
        .ColAlignment(6) = 4
        .ColAlignment(7) = 4
        .ColAlignment(8) = 4
        
    End With
End Sub

Function GravaCTRCTMK()

    On Error GoTo TrataErro

    Set Cmd = Nothing
   
    Conexao.BeginTrans
   
    Set Cmd = New Command

    With Cmd
    
        .CommandText = "{call MNPEDIDOTMK (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
        .CommandType = adCmdText
        .ActiveConnection = Conexao
        .Parameters.Refresh
        
        '@NroPed int ,
        .Parameters(0).Value = IIf(Trim(MskNroPedido.Text) = "", "", Trim(MskNroPedido.Text))
        '@Codcli int ,
        .Parameters(1).Value = ilCodCli
        '@CodRep int ,
        .Parameters(2).Value = ilCodRep
        '@CodCnd int ,
        .Parameters(3).Value = ilCodCnd
        
        '@CIFOB char(1) ,
        
        If Opt_CIF.Value = True Then
            .Parameters(4).Value = "C"
        Else
            .Parameters(4).Value = "F"
        End If
    
        '@DscPdr decimal(6, 3) ,
        .Parameters(5).Value = IIf(Trim(d1.Texto) = "", 0, Trim(d1.Texto))
        '@DscPro decimal(6, 3) ,
        .Parameters(6).Value = IIf(Trim(d2.Texto) = "", 0, Trim(d2.Texto))
        '@DscCnd decimal(6, 3) ,
        .Parameters(7).Value = IIf(Trim(d3.Texto) = "", 0, Trim(d3.Texto))
        '@DscFOB decimal(6, 3) ,
        .Parameters(8).Value = IIf(Trim(d4.Texto) = "", 0, Trim(d4.Texto))
        '@DscTot decimal(6, 3) ,
        .Parameters(9).Value = IIf(Trim(d5.Texto) = "", 0, Trim(d5.Texto))
        '@FlgContr char(1) ,
        .Parameters(10).Value = Trim(d6.Texto)
        
        'If slPedSimples = "S" Then
            '.Parameters(10).Value = "B"
        'End If
        
        '@UFCli char(2) ,
        .Parameters(11).Value = Trim(d7.Texto)
        '@AlqICM decimal(6, 3) ,
        .Parameters(12).Value = IIf(Trim(d8.Texto) = "", 0, Trim(d8.Texto))
        '@MgrMin decimal(6, 3) ,
        .Parameters(13).Value = Format(Trim(LblI.Caption), "########.###")
        '@MgrTot decimal(6, 3) ,
        .Parameters(14).Value = Format(Trim(dlMargemGeral), "####.##")
        '@IdxFin decimal(6, 3) ,
        .Parameters(15).Value = dlPerCusFin
        '@IdxFrt decimal(6, 3) ,
        .Parameters(16).Value = dlPerCusFrt
        '@ComiNeg decimal(6, 3) ,
        .Parameters(17).Value = dlPerComiNeg
        '@ComiOri decimal(6, 3) ,
        .Parameters(18).Value = dlPerComiCalc
        '@TexNeg ntext ,
        .Parameters(19).Value = Trim(TxtNegocio.Text)
        '@TexObs ntext ,
        .Parameters(20).Value = Trim(TxtObserva.Text)
        
        '@ClasCor char(1),
        
        If Status.BackColor = &HFF00& Then
            
            .Parameters(21).Value = "G"
            
        Else
            
            If Status.BackColor = &HFF& Then
                .Parameters(21).Value = "R"
            Else
                .Parameters(21).Value = "B"
            End If
            
        End If
         
        '@IdxPDD decimal(6, 3),
        .Parameters(22).Value = dlIdxPDD
        '@NomTra varchar(40) ,
        .Parameters(23).Value = Trim(TxtTransp.Text)
        '@ChvDsc  Varchar(20)
        .Parameters(24).Value = Trim(slChave)
        '@SitPed char(1)
        .Parameters(25).Value = "N"
        '@DatPed char(1)
        .Parameters(26).Value = Datped
        
        '@FlgKit char(1)
        
        If ChkKit.Value = 1 Then
            .Parameters(27).Value = "S"
        Else
            .Parameters(27).Value = "N"
        End If
    
        '@VlrSimples decimal(10, 2),
        .Parameters(28).Value = dlSimples
        
        '@FlgAlt char(1),
        
        If Trim(slFlgAlt) = "" Then
            .Parameters(29).Value = Null
        Else
            .Parameters(29).Value = "A"
        End If
        
        '@SeqLig int ,
        .Parameters(30).Value = lgSeqLig
        '@FlagOper char(1)
        .Parameters(31).Value = Trim(sgFlagOper)
        
    End With
    
    Set Rs = Cmd.Execute
    Set Rs = Nothing
    Set Cmd = Nothing
    
    If Trim(MskNroPedido.Text) = "" Then
    
        sgQuery = "SELECT isnull(max(nroped),"") as nroped "
        sgQuery = sgQuery + " from PEDIDO "
        sgQuery = sgQuery + "  Where SeqLig = " & Trim(lgSeqLig)
       
        Call Consulta2(sgQuery)
        
        If Rs2.EOF Then
        
            MsgBox "Erro na leitura do pedido - Ligação-> " & Trim(lgSeqLig), vbExclamation + vbOKOnly, "Atenção!"
            
            Rs2.Close
            
            Set Rs2 = Nothing
            
            Conexao.RollbackTrans
            
            LimpaGeral
            
            Exit Function
            
        Else
        
            MskNroPedido.Text = Rs2!NroPed
            
            Rs2.Close
            
            Set Rs2 = Nothing
            
            sgQuery = "insert into PEDIDO_LIGACAO  values ('" & Trim(MskNroPedido.Text) & "'," & lgSeqLig & ")"
            
            Conexao.Execute sgQuery
        
        End If
        
    End If
    
    'DELETA TODOS OS ITENS
    
    sgQuery = "Delete ITEM_PEDIDO where NroPed = '" & Trim(MskNroPedido.Text) & "'"
    
    Conexao.Execute sgQuery
   
    'GRAVA ITENS DO PEDIDO
    
    If GrdNotaCliente.rows > 1 Then
    
        For Linhas = 1 To GrdNotaCliente.rows - 1
        
            If Trim(GrdNotaCliente.TextMatrix(Linhas, 0)) > 0 Then
            
                Set Cmd = New Command
                
                With Cmd
                
                    .CommandText = "{call MNITEMPEDIDOTMK (?,?,?,?,?,?,?,?,?,?,?,?,?)}"
                    .CommandType = adCmdText
                    .ActiveConnection = Conexao
                    .Parameters.Refresh
                    
                    '@NroPed int,
                    .Parameters(0).Value = Trim(MskNroPedido.Text)
                    '@CodPrd int,
                    .Parameters(1).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 0))
                    '@SeqIte int,
                    .Parameters(2).Value = Linhas
                    '@QtdPrd int,
                    .Parameters(3).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 3))
                    '@QtdEmb int,
                    .Parameters(4).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 2))
                    '@ValUnt decimal(10, 2),
                    .Parameters(5).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 4))
                    '@IdxDsc decimal(6, 3) ,
                    .Parameters(6).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 6))
                    '@VlrIte decimal(10, 2),
                    .Parameters(7).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 18))
                    '@FlgTab char(1)       ,
                    .Parameters(8).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 5))
                    '@ValUntN decimal(10, 2),
                    .Parameters(9).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 12))
                    '@MrgPrd decimal(6, 3) ,
                    .Parameters(10).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 13))
                    '@ValCusUnt decimal(10, 4),
                    .Parameters(11).Value = Trim(GrdIndice.TextMatrix(Linhas, 9))
                    '@IdxFix decimal(6, 3),
                    .Parameters(12).Value = Trim(GrdIndice.TextMatrix(Linhas, 10))
                
                End With
                
                Set Rs = Cmd.Execute
                Set Rs = Nothing
                Set Cmd = Nothing
                
            End If
            
        Next Linhas
        
    End If
    
    Conexao.CommitTrans
    
    Exit Function
   
TrataErro:

    Rotina_Erro "GravaCTRCTMK"

End Function

Function Numero_Ped() As Boolean

    Numero_Ped = True
     
    If APLICA = 0 Then
    
        MsgBox "Pedido não cadastrado !", vbExclamation + vbOKOnly, "Atenção!"
        
        Numero_Ped = False
        
        Exit Function
        
    End If
    
    sgQuery = "select * from SEQUENCIA_PEDIDO where CodRep = " & Trim(ilCodRep)
    
    Call Consulta(sgQuery)
    
    If Rs.EOF Then
        
        MsgBox "Registro do Controle de numeração de pedidos inexistente, contate o administrador do sistema", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        Numero_Ped = False
        
        Exit Function
        
    End If
     
    Seqini = IIf(IsNull(Rs!Seqini), "", Trim(Rs!Seqini))
    SeqFim = IIf(IsNull(Rs!SeqFim), "", Trim(Rs!SeqFim))
    
    Rs.Close
    
    Set Rs = Nothing
    
    If Seqini = "" Or SeqFim = "" Then
    
        MsgBox "Registro do Controle de numeração de pedidos inexistente, contate o administrador do sistema", vbExclamation + vbOKOnly, "Atenção!"
        
        Numero_Ped = False
        
        Exit Function
        
    End If
    
    '****************************************************************************
    'A consulta SQL que está comentada a seguir eu fiz no dia 26/04/2010, como
    'forma de corrigir um erro na seqüência de pedidos do representante Luciano.
    'Voltei a usar a consulta normal no memso dia e nenhum outro representante
    'teve acesso a ela. (André Corrêa)
    '****************************************************************************
    
    'sgQuery = "select max(nroped) as pedido from PEDIDO where CodRep = " & Trim(ilCodRep) & "And NroPed < 897000"
    sgQuery = "select max(nroped) as pedido from PEDIDO where CodRep = " & Trim(ilCodRep)
    
    Call Consulta(sgQuery)
    
    If IsNull(Rs("pedido")) = True Then
    
        Dim x As Double
        x = Val(Mid(Rs("pedido"), 2, 5))
        MskNroPedido.Text = Mid(Rs("pedido"), 1, 1) + (Format(x + 1, "00000"))
    
       ' MskNroPedido.Text = SeqIni
        
        Exit Function
        
    Else
    
        ' Dim x As Double
        x = Val(Mid(Rs("Pedido"), 2, 5))
        MskNroPedido.Text = Mid(Rs("Pedido"), 1, 1) + (Format(x + 1, "00000"))
    
       ' MskNroPedido.Text = Rs("Pedido") + 1
        
    End If
    
    Rs.Close
        
    Set Rs = Nothing
    
    If x < Val(Mid(Seqini, 2, 5)) Or x > Val(Mid(SeqFim, 2, 5)) And APLICA = 1 Then
    
    'If (Trim(MskNroPedido.Text) < Seqini Or Trim(MskNroPedido.Text) > SeqFim) And APLICA = 1 Then
        
        MsgBox "Número do pedido fora do intervalo permitido, contate o administrador do sistema", vbExclamation + vbOKOnly, "Atenção!"
        
        Numero_Ped = False
        
    End If

End Function

Function Activate_Ped()

    If bgBloqPed = True Then
    
        If APLICA = 1 Then
            LblBloqPed.Visible = True
        End If
    
    Else
        
        LblBloqPed.Visible = False
        
    End If
   
    slUFOri = "MG"
    slUFRep = ""
    slNomRep = ""
    dlPerCusFrt = 0
    dlPerComiN = 0
    dlPerComiA = 0
    dlPerComiB = 0
    dlPerDesFOB = 0
    dlIdxPDD = 0
    dlIdxAzul = 0
    slFlgSugComi = ""
    dlValPedMin = 0
    dlValLimPrz1 = 0
    dlValParMin = 0
    ilPrzMed1 = 0
    ilPrzMed2 = 0
   
    sgQuery = "select * from REPRESENTANTE where CodRep = " & Trim(ilCodRep)
    
    Call Consulta2(sgQuery)
    
    If Not Rs2.EOF Then
    
        slNomRep = IIf(IsNull(Rs2!NomRep), "", Trim(Rs2!NomRep))
        slUFRep = IIf(IsNull(Rs2!UFRep), "", Trim(Rs2!UFRep))
        dlPerCusFrt = IIf(IsNull(Rs2!PerCusFrt), 0, Trim(Rs2!PerCusFrt))
        dlPerComiN = IIf(IsNull(Rs2!PerComiN), 0, Trim(Rs2!PerComiN))
        dlPerComiA = IIf(IsNull(Rs2!PerComiA), 0, Trim(Rs2!PerComiA))
        dlPerComiB = IIf(IsNull(Rs2!PerComiB), 0, Trim(Rs2!PerComiB))
        dlPerDesFOB = IIf(IsNull(Rs2!PerDesFOB), 0, Trim(Rs2!PerDesFOB))
        slFlgSugComi = IIf(IsNull(Rs2!FlgSugComi), "", Trim(Rs2!FlgSugComi))
        dlIdxPDD = IIf(IsNull(Rs2!IdxPDD), 0, Trim(Rs2!IdxPDD))
        dlIdxAzul = IIf(IsNull(Rs2!IdxAzul), 0, Trim(Rs2!IdxAzul))
        dlPerComiCalc = dlPerComiN
        dlPerTubo100Rep = IIf(IsNull(Rs2!PerTubo100), 0, Trim(Rs2!PerTubo100))
        dlValPedMin = IIf(IsNull(Rs2!ValPedMin), 0, Trim(Rs2!ValPedMin))
        dlValLimPrz1 = IIf(IsNull(Rs2!ValLimPrz1), 0, Trim(Rs2!ValLimPrz1))
        dlValParMin = IIf(IsNull(Rs2!ValParMin), 0, Trim(Rs2!ValParMin))
        ilPrzMed1 = IIf(IsNull(Rs2!PrzMed1), 0, Trim(Rs2!PrzMed1))
        ilPrzMed2 = IIf(IsNull(Rs2!PrzMed2), 0, Trim(Rs2!PrzMed2))
    
    Else
    
        MsgBox "Registro do Representante não encontrado, informe ao administrador do sistema", vbExclamation + vbOKOnly, "Atenção!"
    
    End If
  
    Rs2.Close
    
    Set Rs2 = Nothing
   
    'leitura BALIZA_SUGESTAO (situação [1] Abaixo do limite)
    PerSug1Ini = 0
    PerSug1FimA = 0
    PerSug1FimB = 0

    sgQuery = "select * from BALIZA_SUGESTAO where CodRep = " & Trim(ilCodRep) & " and NroSit = 1 "
    
    Call Consulta2(sgQuery)
    
    If Not Rs2.EOF Then
        
        PerSug1Ini = IIf(IsNull(Rs2!PerSugIni), 0, Trim(Rs2!PerSugIni))
        PerSug1FimA = IIf(IsNull(Rs2!PerSugFimA), 0, Trim(Rs2!PerSugFimA))
        PerSug1FimB = IIf(IsNull(Rs2!PerSugFimB), 0, Trim(Rs2!PerSugFimB))
        
    Else
    
        Rs2.Close
        
        Set Rs2 = Nothing
        
        sgQuery = "select * from BALIZA_SUGESTAO where CodRep is null and NroSit = 1"
        
        Call Consulta2(sgQuery)
        
        If Not Rs2.EOF Then
            PerSug1Ini = IIf(IsNull(Rs2!PerSugIni), 0, Trim(Rs2!PerSugIni))
            PerSug1FimA = IIf(IsNull(Rs2!PerSugFimA), 0, Trim(Rs2!PerSugFimA))
            PerSug1FimB = IIf(IsNull(Rs2!PerSugFimB), 0, Trim(Rs2!PerSugFimB))
        End If
        
    End If
    
    Rs2.Close
    
    Set Rs2 = Nothing
   
    'leitura BALIZA_SUGESTAO (situação [2] Acima do limite)
    PerSug2Ini = 0
    PerSug2FimA = 0
    PerSug2FimB = 0
   
    sgQuery = "select * from BALIZA_SUGESTAO where CodRep = " & Trim(ilCodRep) & " and NroSit = 2 "
    
    Call Consulta2(sgQuery)
    
    If Not Rs2.EOF Then
        
        PerSug2Ini = IIf(IsNull(Rs2!PerSugIni), 0, Trim(Rs2!PerSugIni))
        PerSug2FimA = IIf(IsNull(Rs2!PerSugFimA), 0, Trim(Rs2!PerSugFimA))
        PerSug2FimB = IIf(IsNull(Rs2!PerSugFimB), 0, Trim(Rs2!PerSugFimB))
        
    Else
    
        Rs2.Close
        
        Set Rs2 = Nothing
        
        sgQuery = "select * from BALIZA_SUGESTAO where CodRep is null and NroSit = 2"
        
        Call Consulta2(sgQuery)
        
        If Not Rs2.EOF Then
            PerSug2Ini = IIf(IsNull(Rs2!PerSugIni), 0, Trim(Rs2!PerSugIni))
            PerSug2FimA = IIf(IsNull(Rs2!PerSugFimA), 0, Trim(Rs2!PerSugFimA))
            PerSug2FimB = IIf(IsNull(Rs2!PerSugFimB), 0, Trim(Rs2!PerSugFimB))
        End If
    
    End If
    
    Rs2.Close
    
    Set Rs2 = Nothing
    
    Me.Caption = "UNOCANN Tubos e Conexões - Simulação de Pedidos  [" & Trim(slNomRep) & "]"
    'Me.Caption = "UNOCANN Tubos e Conexões - Manutenção de Pedidos  [" & Trim(slNomRep) & "]"

    If bgConsultaPed = True Then
        
        MskNroPedido.Text = igNroPed
        MskNroPedido_LostFocus
        'CarregaTela
        'Exit Function
        
    End If
   
    If bgPedMKT = True Then
        
        'MskNroPedido.Text = 999999
        
        MskNroPedido_LostFocus
        MskNroPedido.Enabled = False
        
        Exit Function
        
    End If
    
End Function

Function DecrypChave() As Boolean

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

    DecrypChave = False

    TxtChave.Text = UCase(TxtChave.Text)

    tamanho = Len(Trim(TxtChave.Text))
    Senha = Trim(TxtChave.Text)
    Desconto = Right(Senha, 2)
    
    If Not IsNumeric(Desconto) Or Trim(Desconto) = 0 Then
        
        MsgBox "Chave inválida", vbExclamation + vbOKOnly, "Atenção!"
        
        Exit Function
        
    End If

    Senha = Mid(Senha, 1, tamanho - 2)
    tamanho = Len(Senha)
    final = Mid(Senha, 1, 1) & Mid(Senha, tamanho, 1)

    If Not IsNumeric(final) Then
        
        MsgBox "Chave inválida", vbExclamation + vbOKOnly, "Atenção!"
        
        Exit Function
        
    End If

    Senha = Trim(Mid(Senha, 2, tamanho - 1))
    Senha = Mid(Senha, 1, tamanho - 2)
    tamanho = Len(Senha)

    For i = tamanho To 1 Step -1
        resultado = resultado & Mid(Senha, i, 1)
    Next i

    resultado = "&H" & resultado
    resultado = resultado / Val((Val(final) + Desconto))

    If Trim(MskNroPedido.Text) <> Trim(resultado) Then
        
        MsgBox "Chave inválida", vbExclamation + vbOKOnly, "Atenção!"
        
        Exit Function
        
    End If

    DecrypChave = True
    slChave = Trim(UCase(TxtChave.Text))
    ilDescChave = Val(Desconto)

End Function

Function CarregaGridGrupo()

    GrdProduto.rows = 1
    
    Set Rs = Nothing
    
    sgQuery = " select idegrp, nomgrp from grupo_produto "
    sgQuery = sgQuery & " where idegrp in (select idegrp from produto where flgsitu = 'N') "
    sgQuery = sgQuery & " order by nomgrp "
    
    Consulta sgQuery
    
    GrdGrupo.rows = 1
    
    blI = 1
    
    Do While Not Rs.EOF
    
        GrdGrupo.rows = GrdGrupo.rows + 1
        GrdGrupo.TextMatrix(blI, 0) = Trim(Rs!NomGrp)
        GrdGrupo.TextMatrix(blI, 1) = Trim(Rs!IdeGrp)
        
        blI = blI + 1
        
        Rs.MoveNext
        
    Loop
    
    Rs.Close
    
    Set Rs = Nothing

End Function

Function CarregaGridProduto() As Boolean

    Set Rs = Nothing
    
    'sgQuery = " select Codprd, Dscprd, QtdEmb, ValUntN, ValUntA, ValUntB, MrgPrd from PRODUTO "
    'sgQuery = sgQuery & " where idegrp = " & ilGrupo & " and flgsitu = 'N'"
    'sgQuery = sgQuery & " order by DscPrd "
        
    sgQuery = "SELECT a.Codprd, a.Dscprd, a.QtdEmb, b.ValUntN, b.ValUntA, b.ValUntB, b.MrgPrd from PRODUTO a, PRECO_PRODUTO b"
    sgQuery = sgQuery + "   WHERE a.flgsitu = 'N'"
    sgQuery = sgQuery + "     and a.idegrp = " & ilGrupo
    sgQuery = sgQuery + "     and a.codprd = b.codprd"
    sgQuery = sgQuery + "     and b.datativ = (select max(datativ) from preco_produto"
    sgQuery = sgQuery + "                       Where Codprd = b.codprd"
    sgQuery = sgQuery + "                         and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"
    sgQuery = sgQuery & " order by a.DscPrd "
    
    Consulta sgQuery
    
    GrdProduto.rows = 1
    
    blI = 1
    
    Do While Not Rs.EOF
    
        GrdProduto.rows = GrdProduto.rows + 1
        GrdProduto.TextMatrix(blI, 1) = Trim(Rs!CodPrd)
        GrdProduto.TextMatrix(blI, 2) = Trim(Rs!DSCPRD)
        GrdProduto.TextMatrix(blI, 3) = Trim(Rs!QtdEmb)
        GrdProduto.TextMatrix(blI, 4) = Format(Rs!ValUntN, "##,###,##0.00")
        'GrdProduto.TextMatrix(blI, 5) = Format(Rs!ValUntA, "##,###,##0.00")
        'GrdProduto.TextMatrix(blI, 6) = Format(Rs!ValUntB, "##,###,##0.00")
        GrdProduto.TextMatrix(blI, 5) = Format(0, "##,###,##0.00")
        GrdProduto.TextMatrix(blI, 6) = Format(0, "##,###,##0.00")
        GrdProduto.row = blI
        
        With GrdProduto
        
            .col = 0
        
            'alterna a cor da col0 conforme margem do produto
            
            If Rs!MrgPrd = 0 Then
                
                .CellBackColor = vbWhite
                
            Else
            
                If Rs!MrgPrd < 8 Then
                    
                    .CellBackColor = vbRed
                    
                Else
                
                    If Rs!MrgPrd <= 15 Then
                        .CellBackColor = vbGreen
                    Else
                        .CellBackColor = vbBlue
                    End If
                    
                End If
                
            End If
            
        End With
       
        blI = blI + 1
        
        Rs.MoveNext
        
    Loop
    
    Rs.Close
    
    Set Rs = Nothing
    
End Function

Function CarregaTela() As Boolean

    '*************************************************************************************
    'Função responsável por levantar todas as informações referentes ao pedido e exibir os
    'resultados em tela.
    '*************************************************************************************
    
    On Error GoTo TrataErro
    
    Dim NroPed As String
    'Dim NroPed As Double
    Dim PrdAnt As Double
    Dim blI As Integer
    Dim dlPesBru As Double
    
    Dim lDados1 As New ADODB.Recordset
    Dim lDados2 As New ADODB.Recordset
    Dim lQuantidadeProdutoPedido As Integer
    Dim lQuantidadeProdutoFaturado As Integer
    Dim lQuantidadeProdutoEntregue As Integer
    
    CarregaTela = True
    
    blI = 1
    
    '*************************************************************************************
    'Avalia se o pedido informado realmente existe.
    '*************************************************************************************
    
    sgQuery = "Select nroped From PEDIDO Where NroPed = '" & Trim(MskNroPedido.Text) & "'"
    
    Consulta sgQuery
    
    If Rs.EOF Then
        blI = 0
    Else
        NroPed = IIf(IsNull(Rs!NroPed), "", Trim(Rs!NroPed))
    End If
    
    If blI = 0 Or NroPed = "" Then
    
        MsgBox "Pedido não cadastrado !", vbExclamation + vbOKOnly, "Atenção!"
        
        CarregaTela = False
        
        Exit Function
        
    End If
    
    '*************************************************************************************
    'Levanta e exibe os dados do cabeçalho do pedido. Preenche também a guia COMUNICAÇÃO.
    '*************************************************************************************
    
    sgQuery = "Select a.*, b.NomCli, c.DscCnd From PEDIDO a, CLIENTE b, CONDICAO c"
    sgQuery = sgQuery & " Where a.NroPed = '" & Trim(MskNroPedido.Text) & "' And a.CodCli = b.CodCli and a.codcnd = c.codcnd"
    
    Consulta sgQuery
    
    blI = 1
    
    If Not Rs.EOF Then
        
        cboCli.Criterio = Trim(Rs!NomCli)
        slremet = Trim(Rs!NomCli)
        Datped = Trim(Rs!Datped)
        cboCli.codigo = Trim(Rs!Codcli)
        CboCondPag.Criterio = Trim(Rs!DscCnd)
        CboCondPag.codigo = Trim(Rs!CodCnd)
        ilCodCnd = Trim(Rs!CodCnd)

'-------------------------------------------------------------------------
    If ilCodCnd = 1 Or ilCodCnd = 12 Or ilCodCnd = 24 Then 'A vista ou 14 dias
        bAVista = True
        If bEKit = True Then 'Venda a vista Kit irrigação
        
            If ilCodRep = 2 Or ilCodRep = 7 Or ilCodRep = 8 Then
                dDscRegiao = 26.87
          ElseIf ilCodRep = 600 Or ilCodRep = 800 Or ilCodRep = 1001 Or ilCodRep = 1900 Or ilCodRep = 2100 Or ilCodRep = 6000 Or ilCodRep = 7050 Or ilCodRep = 7060 Or ilCodRep = 7075 Then
                dDscRegiao = 24.15
            ElseIf ilCodRep = 905 Or ilCodRep = 5000 Or ilCodRep = 5001 Then
                dDscRegiao = 22.37
'            Else
'                dDscRegiao = 24.15
            End If
        
        Else
           
            If ilCodRep = 2 Or ilCodRep = 7 Or ilCodRep = 8 Then
                dDscRegiao = 22.37
            ElseIf ilCodRep = 600 Or ilCodRep = 800 Or ilCodRep = 1001 Or ilCodRep = 1900 Or ilCodRep = 2100 Or ilCodRep = 6000 Or ilCodRep = 7050 Or ilCodRep = 7060 Or ilCodRep = 7075 Then
                dDscRegiao = 19.69
            ElseIf ilCodRep = 905 Or ilCodRep = 5000 Or ilCodRep = 5001 Then
                dDscRegiao = 20.36
'            Else
'                dDscRegiao = 22.37
            End If
                
        End If

    Else 'A Prazo
     bAVista = False
        If bEKit = True Then 'Venda a Normal Kit irrigação
        
            If ilCodRep = 2 Or ilCodRep = 7 Or ilCodRep = 8 Then
                dDscRegiao = 24.56
            ElseIf ilCodRep = 600 Or ilCodRep = 800 Or ilCodRep = 1001 Or ilCodRep = 1900 Or ilCodRep = 2100 Or ilCodRep = 6000 Or ilCodRep = 7050 Or ilCodRep = 7060 Or ilCodRep = 7075 Then
             dDscRegiao = 21.8
            ElseIf ilCodRep = 905 Or ilCodRep = 5000 Or ilCodRep = 5001 Then
                dDscRegiao = 19.96
            '            Else
            '                dDscRegiao = 21.8
            End If
        
        Else
            
            If ilCodRep = 2 Or ilCodRep = 7 Or ilCodRep = 8 Then
                dDscRegiao = 19.96
            ElseIf ilCodRep = 600 Or ilCodRep = 800 Or ilCodRep = 1001 Or ilCodRep = 1900 Or ilCodRep = 2100 Or ilCodRep = 6000 Or ilCodRep = 7050 Or ilCodRep = 7060 Or ilCodRep = 7075 Then
                dDscRegiao = 17.2
            ElseIf ilCodRep = 905 Or ilCodRep = 5000 Or ilCodRep = 5001 Then
                dDscRegiao = 17.89
'            Else
'                dDscRegiao = 19.96
            End If
                
        End If

    End If


'------------------------------------------------------------
'    If ilCodCnd = 1 Or ilCodCnd = 12 Then 'A vista ou 14 dias
' '       dDscRegiao = IIf(Trim(d5.Texto) = "", 0, Trim(d5.Texto))
'        bAVista = True
'    Else 'A Prazo
' '       dDscRegiao = IIf(Trim(d1.Texto) = "", 0, Trim(d1.Texto))
'        bAVista = False
'    End If
'''------------------------------------------------------------
        
        
        dlPerComiNeg = Rs!ComiNeg
        slFlgAlt = IIf(IsNull(Rs!flgalt), "", Trim(Rs!flgalt))
        slClasCor = Rs!ClasCor
        
        If Rs!VlrSimples > 0 Then
            slPedSimples = "S"
        Else
            slPedSimples = "N"
        End If
        
        If Trim(Rs!CIFOB) = "F" Then
            Opt_FOB.Value = True
        End If
        
        If Trim(Rs!FlgKit) = "S" Then
            ChkKit.Value = 1
            bEKit = True
        Else
            ChkKit.Value = 0
            bEKit = False
        End If
        
        TxtTransp.Text = Trim(Rs!NomTra)
        TxtNegocio.Text = Rs!texneg
        TxtObserva.Text = Rs!TexObs
        slChave = Trim(Rs!ChvDsc)
        
        If Trim(slChave) <> "" Then
            ilDescChave = Right(slChave, 2)
            dlSumDscItem = ilDescChave
        End If
        
        If IsNull(Rs!DatEnv) Then
            
            CmdCancelar.Enabled = True
        
        Else
            
            bgBloqPed = True
            
            If APLICA = 1 Then
                'LblBloqPed.Visible = True
                MsgBox "Este pedido não pode ser alterado", vbExclamation + vbOKOnly, "Atenção!"
            End If
            
            CmdCancelar.Enabled = False
            
        End If
        
        If Trim(Rs!SitPed) = "C" Or Trim(Rs!SitPed) = "U" Then
            bgBloqPed = True
            'LblBloqPed.Caption = "PEDIDO CANCELADO em " & Format(Rs!DatAtu, "dd/mm/yyyy")
            MsgBox "PEDIDO CANCELADO em " & Format(Rs!DatAtu, "dd/mm/yyy"), vbExclamation + vbOKOnly, "Atenção!"
            LblBloqPed.Visible = True
            CmdCancelar.Enabled = False
        End If
                        
        '*********************************************************************************
        'Determina se o botão LIBERA estará habilitado ou não.
        '*********************************************************************************
        
        blImpr = True
        
        If igTela = "Monit" Then
        
            If Rs!ClasCor = "R" And IsNull(Rs!CodUsuLib) And (Trim(Rs!SitPed) <> "C" And Trim(Rs!SitPed) <> "U") Then
                
                If sgFlgUsu = "L" Then
                    CmdLibera.Visible = True
                    CmdImpr.Visible = False
                Else
                    blImpr = False
                End If
            
            End If
            
        End If
       
        SSTConhec.Tab = 0
        
        Rs.Close
        
        Set Rs = Nothing
        
        '*********************************************************************************
        'Monta o grid com dados referentes aos itens do pedido.
        '*********************************************************************************

        sgQuery = "Select a.*, b.DscPrd, b.IdeGrp, b.PesUnt,  b.FlgKit,"
        sgQuery = sgQuery + "        c.MrgPrd as MrgPrdPro, c.ValUntN as ValUntNPro,"
        sgQuery = sgQuery + "        c.valcusuntqtd , c.valcusadicqtd, c.alqimpfed"
        sgQuery = sgQuery + "  From ITEM_PEDIDO a, PRODUTO b, PRECO_PRODUTO c"
        sgQuery = sgQuery + "  Where a.NroPed = '" & Trim(MskNroPedido.Text) & "'"
        sgQuery = sgQuery + "    and a.CodPrd = b.CodPrd"
        sgQuery = sgQuery + "    and b.codprd = c.codprd"
        sgQuery = sgQuery + "    and c.datativ = (select max(datativ) from preco_produto"
        sgQuery = sgQuery + "                     Where codprd = c.codprd"
        sgQuery = sgQuery + "                       and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"
        sgQuery = sgQuery + "    order by a.SeqIte"
       
        Consulta sgQuery
        
        GrdNotaCliente.rows = 1
        
        blI = 1
        
        Do While Not Rs.EOF
            
            GrdNotaCliente.rows = GrdNotaCliente.rows + 1
            GrdNotaCliente.TextMatrix(blI, 0) = Format(Trim(Rs!CodPrd), "0000")
            GrdNotaCliente.TextMatrix(blI, 1) = Trim(Rs!DSCPRD)
            GrdNotaCliente.TextMatrix(blI, 2) = Rs!QtdEmb
            GrdNotaCliente.TextMatrix(blI, 3) = Rs!qtdprd
            GrdNotaCliente.TextMatrix(blI, 4) = Format(Rs!ValUnt, "##,###,##0.00")
            GrdNotaCliente.TextMatrix(blI, 5) = Rs!FlgTab
            
            If Trim(Rs!FlgTab) = "" Then
                
                ilNumTab = 0
                
            Else
                
                If Trim(Rs!FlgTab) = "A" Then
                    ilNumTab = 1
                Else
                    ilNumTab = 2
                End If
                
            End If
            
            GrdNotaCliente.TextMatrix(blI, 6) = Format(Rs!IdxDsc, "##0.00")
'--------------------------------------------------------------------------
            If bAVista = True Then 'estou trabalhando a vista
            
                If GrdNotaCliente.TextMatrix(blI, 6) <= dDscRegiao Then
                    iDscRegiao = iDscRegiao + 1 'estou trabalhando com desconto da região
                Else
                    iDscForaRegiao = iDscForaRegiao + 1 'o desconto está fora da região
                End If
                
             Else
            
                If GrdNotaCliente.TextMatrix(blI, 6) <= dDscRegiao Then
                    iDscRegiao = iDscRegiao + 1
                Else
                    iDscForaRegiao = iDscForaRegiao + 1
                End If
            
            
            End If
'-------------------------------------------------------------------------
            GrdNotaCliente.TextMatrix(blI, 7) = Format(GrdNotaCliente.TextMatrix(blI, 4) - (GrdNotaCliente.TextMatrix(blI, 4) * (GrdNotaCliente.TextMatrix(blI, 6) / 100)), "##,###,##0.00")
            GrdNotaCliente.TextMatrix(blI, 18) = Format(Rs!VlrIte, "##,###,##0.00")
            GrdNotaCliente.TextMatrix(blI, 8) = Rs!IdeGrp
            GrdNotaCliente.TextMatrix(blI, 9) = Format(Rs!ValUnt * Rs!qtdprd, "##,###,##0.00")
            
            dlPesBru = Rs!PesUnt * Rs!qtdprd
            
            GrdNotaCliente.TextMatrix(blI, 10) = Format(dlPesBru, "###,##0.0000")
            
            VlIdealItem = (Rs!ValUntNPro - ((Rs!ValUntNPro * dlSumDscItemORIG) / 100)) * Rs!qtdprd
            
            GrdNotaCliente.TextMatrix(blI, 11) = Format(VlIdealItem, "###,##0.00")
            GrdNotaCliente.TextMatrix(blI, 12) = Format(Rs!ValUntNPro, "###,##0.00")
            GrdNotaCliente.TextMatrix(blI, 13) = Format(Rs!MrgPrdPro, "##0.000")
            GrdNotaCliente.TextMatrix(blI, 14) = Rs!FlgKit
            GrdNotaCliente.TextMatrix(blI, 15) = Format(Rs!valcusuntqtd, "###,##0.00")
            GrdNotaCliente.TextMatrix(blI, 16) = Format(Rs!valcusadicqtd, "###,##0.00")
            GrdNotaCliente.TextMatrix(blI, 17) = Format(Rs!AlqImpFed, "##,##0.000")
            
            blI = blI + 1
            
            Rs.MoveNext
            
        Loop
       
        Rs.Close
        
        Set Rs = Nothing
        
        sgFlagOper = "A"
    
        '*********************************************************************************
        'Monta o grid com dados referentes a entregas realizadas e possíveis gerações de
        'saldos. Esses dados estão disponíveis na guia ENTREGAS. Se o pedido ainda não
        'tiver sido faturado, apernas os seus itens serão informados, com os campos de
        'quantidades entregues zerados.
        '*********************************************************************************
        
        'sgQuery = "Select  a.codprd, a.qtdprd, a.qtdprdfat, b.DscPrd, c.nronot, c.dateminot, a.qtdprdfat + isnull(d.sum_saldo_entregue,0) as totentreg, c.nroped"
        'sgQuery = sgQuery + "  From ITEM_PEDIDO a, PRODUTO b, pedido c,"
        'sgQuery = sgQuery + "       (select a.codprd, sum_saldo_entregue = sum(a.qtdprdfat) from item_pedido_saldo a, pedido_saldo b"
        'sgQuery = sgQuery + "          Where a.NroPed = " & Trim(MskNroPedido.Text)
        'sgQuery = sgQuery + "            and a.NroPed = b.nroped "
        'sgQuery = sgQuery + "            and a.NroPedsdo = b.nropedsdo "
        'sgQuery = sgQuery + "            and b.SitPed = 'N'"
        'sgQuery = sgQuery + "            group by a.codprd) d"
        'sgQuery = sgQuery + "   Where a.NroPed = " & Trim(MskNroPedido.Text)
        'sgQuery = sgQuery + "     and a.CodPrd = b.CodPrd"
        'sgQuery = sgQuery + "     and a.nroped = c.NroPed"
        'sgQuery = sgQuery + "     and a.codprd *= d.codprd"
        'sgQuery = sgQuery + "     order by a.SeqIte"
        
        sgQuery = "SELECT A.CodPrd, A.QtdPrd, A.QtdPrdFat, B.DscPrd, C.NroNot, C.DatEmiNot, A.QtdPrdFat As TotEntReg, C.NroPed "
        sgQuery = sgQuery & "FROM Item_Pedido A "
        sgQuery = sgQuery & "INNER JOIN Produto B ON A.CodPrd = B.CodPrd "
        sgQuery = sgQuery & "INNER JOIN Pedido C ON A.NroPed = C.NroPed "
        sgQuery = sgQuery & "LEFT OUTER JOIN (SELECT A.CodPrd, Sum_Saldo_Entregue = SUM(A.QtdPrdFat) FROM Pedido_Saldo B INNER JOIN Item_Pedido_Saldo A ON A.NroPedSdo = B.NroPedSdo WHERE A.NroPed = '" & MskNroPedido.Text & "' And B.SitPed = 'N' GROUP BY A.CodPrd) D ON A.CodPrd = D.CodPrd "
        sgQuery = sgQuery & "WHERE A.NroPed = '" & MskNroPedido.Text & "'"
        sgQuery = sgQuery & "ORDER BY A.SeqIte"
        
        Consulta sgQuery
        
        GrdEntrega.rows = 1
        
        blI = 1
        
        Do While Not Rs.EOF
        
            '*****************************************************************
            'Alguns tubos vendidos pela Unocann estão sendo produzidos pela
            'Polyvin. No entanto, os tubos vindos de Uberaba não podem ficar
            'disponíveis aos representantes, que nem devem saber desta
            'operação. Assim, criei uma tabela chamada
            'ProdutosImportadosPolyvin e nela relacionei os tubos fabricados
            'pela Unocann com seus similares oriundos do triângulo.
            
            'Na hora de exibir um pedido ao vendedor, se a quantidade pedida
            'de um determinado produto for diferente da faturada, verifica-se
            'se o referido produto possui similar cedido pela Polyvin. Se sim,
            'o programa checa se há registro deste similar para o pedido atual
            'e, se houver, informa os seus dados de faturamento em lugar do
            'produto Unocann. Caso o produto atual não tenha similar Polyvin
            'cadastrado, os dados de faturamento são informados normalmente.
            
            'Por fim, os dados do produto cedido pela empresa de Uberaba não
            'podem aparecer para o representante de vendas. Se o registro
            'atual for de um tubo do triângulo mineiro, o loop dá um salto e
            'passa para o próximo.
            
            'Determinação passada por Jacson Nogueira em fins de agosto. Foi
            'executada por André Corrêa, em 04/09/2009.
            '*****************************************************************
            
            If Rs("QtdPrd") <> Rs("QtdPrdFat") Then
                
                sgQuery = "SELECT * FROM ProdutosImportadosPolyvin WHERE CodProdutoUnocann = " & Rs("CodPrd")
                
                lDados1.Open sgQuery, Conexao, adOpenDynamic, adLockOptimistic
                
                If lDados1.EOF = False Then
                    
                    sgQuery = "SELECT P.NROPED, SUM(IP.QTDPRDFAT + U.QTDE) AS 'Entregas' "
                    sgQuery = sgQuery & "FROM PEDIDO P "
                    sgQuery = sgQuery & "INNER JOIN ITEM_PEDIDO IP ON P.NROPED = IP.NROPED "
                    sgQuery = sgQuery & "LEFT OUTER JOIN (SELECT P.NROPED, ISNULL(IPS.QTDPRDFAT, 0) AS QTDE FROM PEDIDO P LEFT OUTER JOIN ITEM_PEDIDO_SALDO IPS ON P.NROPED = IPS.NROPED WHERE P.NROPED = '" & Rs("NroPed") & "') AS U ON P.NROPED = U.NROPED "
                    sgQuery = sgQuery & "WHERE P.NroPed = '" & Rs("NroPed") & "' And IP.CODPRD = " & lDados1("CodProdutoPolyvin") & " "
                    sgQuery = sgQuery & "GROUP BY P.NROPED"
                            
                    Consulta2 sgQuery
                    
                    sgQuery = "SELECT * FROM ITEM_PEDIDO WHERE NroPed = '" & Rs("NroPed") & "' And CodPrd = " & lDados1("CodProdutoPolyvin")
                    
                    lDados2.Open sgQuery, Conexao, adOpenDynamic, adLockOptimistic
                    
                    If lDados2.EOF = False Then
                        
                        lQuantidadeProdutoPedido = lDados2("QtdPrd")
                        lQuantidadeProdutoFaturado = lDados2("QtdPrdFat")
                        lQuantidadeProdutoEntregue = Rs2("Entregas")
                        
                    Else
                    
                        lQuantidadeProdutoPedido = Rs("QtdPrd")
                        lQuantidadeProdutoFaturado = Rs("QtdPrdFat")
                        lQuantidadeProdutoEntregue = Rs("TotEntReg")
                        
                    End If
                    
                    Rs2.Close
                    lDados2.Close
                    
                Else
                
                    lQuantidadeProdutoPedido = Rs("QtdPrd")
                    lQuantidadeProdutoFaturado = Rs("QtdPrdFat")
                    lQuantidadeProdutoEntregue = Rs("TotEntReg")
                    
                End If
                
                lDados1.Close
                                
            Else
                
                sgQuery = "SELECT * FROM ProdutosImportadosPolyvin WHERE CodProdutoPolyvin = " & Rs("CodPrd")
                
                lDados1.Open sgQuery, Conexao, adOpenDynamic, adLockOptimistic
                
                If lDados1.EOF = False Then
                    
                    lDados1.Close
                    
                    GoTo Continua_Item
                    
                End If
                
                lQuantidadeProdutoPedido = Rs("QtdPrd")
                lQuantidadeProdutoFaturado = Rs("QtdPrdFat")
                lQuantidadeProdutoEntregue = Rs("TotEntReg")
                
                lDados1.Close
                
            End If
            
            '*****************************************************************
            'Exibe os dados do produto atual.
            '*****************************************************************
            
            GrdEntrega.rows = GrdEntrega.rows + 1
            GrdEntrega.TextMatrix(blI, 0) = Format(Trim(Rs!CodPrd), "0000")
            GrdEntrega.TextMatrix(blI, 1) = Trim(Rs!DSCPRD)
            GrdEntrega.TextMatrix(blI, 2) = lQuantidadeProdutoPedido
            GrdEntrega.TextMatrix(blI, 3) = lQuantidadeProdutoFaturado
            GrdEntrega.TextMatrix(blI, 4) = lQuantidadeProdutoEntregue
            GrdEntrega.row = blI
            GrdEntrega.col = 4
            GrdEntrega.CellBackColor = &HC0FFFF
            GrdEntrega.CellForeColor = &H80000012
            GrdEntrega.CellAlignment = vbAlignRight
            
            If lQuantidadeProdutoFaturado > 0 Then
                GrdEntrega.TextMatrix(blI, 5) = Format(Rs!NroNot, "000000")
                GrdEntrega.TextMatrix(blI, 6) = Format(Rs!DatEmiNot, "dd/mm/yyyy")
            End If
             
            GrdEntrega.col = 1
            GrdEntrega.CellForeColor = &H80000012
            
            '*****************************************************************************
            'Os itens com que tiverem sido completamente entregues ao cliente aparecem no
            'grid com fundo verde; aqueles que geraram saldo aparecem com fundo vermelho.
            '*****************************************************************************
            
            If lQuantidadeProdutoEntregue = lQuantidadeProdutoPedido Then
                
                GrdEntrega.CellBackColor = &HFF00&
                
            Else
                
                If lQuantidadeProdutoEntregue < lQuantidadeProdutoPedido Then
                    GrdEntrega.CellForeColor = &HFFFF&
                    GrdEntrega.CellBackColor = vbRed
                Else
                    GrdEntrega.CellForeColor = &H80000004
                    GrdEntrega.CellBackColor = &HFF0000
                End If
                
            End If
                       
            '*****************************************************************************
            'Verifica se há saldo de pedido gerado para o item em questão. Se não houver,
            'salta a inserção de registros no grid de saldos e segue com o loop.
            '*****************************************************************************

            sgQuery = "select a.nropedsdo, b.datped, a.qtdprd,"
            sgQuery = sgQuery + " a.qtdprdfat , b.nronot, b.dateminot"
            sgQuery = sgQuery + " From item_pedido_saldo a, pedido_saldo b"
            sgQuery = sgQuery + " Where a.NroPed = '" & Trim(MskNroPedido.Text) & "'"
            sgQuery = sgQuery + " and a.nroped = b.nroped"
            sgQuery = sgQuery + " and a.nropedsdo = b.nropedsdo"
            sgQuery = sgQuery + " and b.sitped = 'N'"
            sgQuery = sgQuery + " and a.codprd = " & Trim(Rs!CodPrd)
            
            Consulta2 sgQuery
          
            If Rs2.EOF = True Then
                
                Rs2.Close
                
                Set Rs2 = Nothing
                
                GoTo Continua_Item
                
            End If
          
            GrdEntrega.col = 8
            GrdEntrega.CellBackColor = &H400000
            GrdEntrega.CellForeColor = &HFFFF&
            GrdEntrega.TextMatrix(blI, 8) = Rs2!nropedsdo
            GrdEntrega.TextMatrix(blI, 9) = Rs2!qtdprd
            GrdEntrega.TextMatrix(blI, 10) = Rs2!qtdprdfat
            
            If Rs2!qtdprdfat > 0 Then
                GrdEntrega.TextMatrix(blI, 11) = Format(Rs2!NroNot, "000000")
                GrdEntrega.TextMatrix(blI, 12) = Format(Rs2!DatEmiNot, "dd/mm/yyyy")
            End If
                        
            blI = blI + 1
            
            Rs2.MoveNext
             
            Do While Not Rs2.EOF
                
                GrdEntrega.rows = GrdEntrega.rows + 1
                GrdEntrega.row = blI
                GrdEntrega.col = 8
                GrdEntrega.CellBackColor = &H400000
                GrdEntrega.CellForeColor = &HFFFF&
                GrdEntrega.TextMatrix(blI, 8) = Rs2!nropedsdo
                GrdEntrega.TextMatrix(blI, 9) = Rs2!qtdprd
                GrdEntrega.TextMatrix(blI, 10) = Rs2!qtdprdfat
                
                If Rs2!qtdprdfat > 0 Then
                    GrdEntrega.TextMatrix(blI, 11) = Format(Rs2!NroNot, "000000")
                    GrdEntrega.TextMatrix(blI, 12) = Format(Rs2!DatEmiNot, "dd/mm/yyyy")
                End If
             
                blI = blI + 1
                
                Rs2.MoveNext
                
            Loop
          
            Rs2.Close
            
            Set Rs2 = Nothing
            
            blI = blI - 1
          
Continua_Item:
          
            blI = blI + 1
            
            Rs.MoveNext
            
        Loop
                      
        '*********************************************************************************
        'A consulta a seguir levanta possíveis alterações no saldo de pedido em relação ao
        'pedido original. Podem ocorrer alterações nos pedidos após a sua digitação no JP;
        'como o força bloqueia tentativas de alteração nessa fase, os representantes
        'enviam tais mudanças por outros meios, como fax ou e-mail. Essas modificações são
        'processadas diretamente no JP e não ficam visíveis ao Força. Assim, possíveis
        'inclusões de itens de pedidos podem aparecer como saldo sem existirem no pedido
        'original. Daí a razão de a consulta a seguir procurar itens de saldo sem registro
        'correspondente nos itens de pedido.
        '*********************************************************************************

        sgQuery = "select a.codprd, a.nropedsdo, b.datped, a.qtdprd,"
        sgQuery = sgQuery + " a.qtdprdfat , b.nronot, b.dateminot, c.dscprd"
        sgQuery = sgQuery + " From item_pedido_saldo a, pedido_saldo b, produto c"
        sgQuery = sgQuery + " Where a.NroPed = '" & Trim(MskNroPedido.Text) & "'"
        sgQuery = sgQuery + " and a.nroped = b.nroped"
        sgQuery = sgQuery + " and a.nropedsdo = b.nropedsdo"
        sgQuery = sgQuery + " and b.sitped = 'N'"
        sgQuery = sgQuery + " and a.codprd = c.codprd"
        sgQuery = sgQuery + " and not exists (select codprd from item_pedido "
        sgQuery = sgQuery + "                   where NroPed = '" & Trim(MskNroPedido.Text) & "'"
        sgQuery = sgQuery + "                     and codprd = a.codprd)"
        
        Consulta2 sgQuery
                        
        Do While Not Rs2.EOF
             
            GrdEntrega.rows = GrdEntrega.rows + 1
            GrdEntrega.row = blI
            GrdEntrega.TextMatrix(blI, 0) = Format(Trim(Rs2!CodPrd), "0000")
            GrdEntrega.TextMatrix(blI, 1) = Trim(Rs2!DSCPRD)
            GrdEntrega.col = 8
            GrdEntrega.CellBackColor = &H400000
            GrdEntrega.CellForeColor = &HFFFF&
            GrdEntrega.TextMatrix(blI, 8) = Rs2!nropedsdo
            GrdEntrega.TextMatrix(blI, 9) = Rs2!qtdprd
            GrdEntrega.TextMatrix(blI, 10) = Rs2!qtdprdfat
            
            If Rs2!qtdprdfat > 0 Then
                GrdEntrega.TextMatrix(blI, 11) = Format(Rs2!NroNot, "000000")
                GrdEntrega.TextMatrix(blI, 12) = Format(Rs2!DatEmiNot, "dd/mm/yyyy")
            End If
             
            GrdEntrega.col = 1
          
            If Rs2!qtdprdfat > 0 Then
                GrdEntrega.CellForeColor = &H80000004
                GrdEntrega.CellBackColor = &HFF0000
            End If
             
            blI = blI + 1
            
            Rs2.MoveNext
             
        Loop
       
        blI = blI - 1
                       
        Rs2.Close
        
        Set Rs = Nothing
        
    End If
    
    CarregaGridConsulta
    
    tab_simulacao_pedido.Tab = 7
    Exit Function

TrataErro:
    
    Rotina_Erro "CarregaTela"
    
    CarregaTela = False
    
End Function

Function GravaCTRC()

    On Error GoTo TrataErro

    Set Cmd = Nothing
   
    Conexao.BeginTrans
   
    Set Cmd = New Command

    With Cmd
    
        .CommandText = "{call MNPEDIDO (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
        .CommandType = adCmdText
        
        .ActiveConnection = Conexao
        
        .Parameters.Refresh
        
        '@NroPed varchar ,
        .Parameters(0).Value = Trim(MskNroPedido.Text)
        '@Codcli int ,
        .Parameters(1).Value = ilCodCli
        '@CodRep int ,
        .Parameters(2).Value = ilCodRep
        '@CodCnd int ,
        .Parameters(3).Value = ilCodCnd
        '@CIFOB char(1) ,
        
        If Opt_CIF.Value = True Then
            .Parameters(4).Value = "C"
        Else
            .Parameters(4).Value = "F"
        End If
        
        '@DscPdr decimal(6, 3) ,
        .Parameters(5).Value = IIf(Trim(d1.Texto) = "", 0, Trim(d1.Texto))
        '@DscPro decimal(6, 3) ,
        .Parameters(6).Value = IIf(Trim(d2.Texto) = "", 0, Trim(d2.Texto))
        '@DscCnd decimal(6, 3) ,
        .Parameters(7).Value = IIf(Trim(d3.Texto) = "", 0, Trim(d3.Texto))
        '@DscFOB decimal(6, 3) ,
        .Parameters(8).Value = IIf(Trim(d4.Texto) = "", 0, Trim(d4.Texto))
        '@DscTot decimal(6, 3) ,
        .Parameters(9).Value = IIf(Trim(d5.Texto) = "", 0, Trim(d5.Texto))
        '@FlgContr char(1) ,
        .Parameters(10).Value = Trim(d6.Texto)
        
        'If slPedSimples = "S" Then
            '.Parameters(10).Value = "B"
        'End If
        
        '@UFCli char(2) ,
        .Parameters(11).Value = Trim(d7.Texto)
        '@AlqICM decimal(6, 3) ,
        .Parameters(12).Value = IIf(Trim(d8.Texto) = "", 0, Trim(d8.Texto))
        '@MgrMin decimal(6, 3) ,
        .Parameters(13).Value = 0 'Format(Trim(LblI.Caption), "########.###")
        '@MgrTot decimal(6, 3) ,
        .Parameters(14).Value = Format(Trim(dlMargemGeral), "####.##")
        '@IdxFin decimal(6, 3) ,
        .Parameters(15).Value = dlPerCusFin
        '@IdxFrt decimal(6, 3) ,
        .Parameters(16).Value = dlPerCusFrt
        '@ComiNeg decimal(6, 3) ,
        .Parameters(17).Value = dlPerComiNeg
        '@ComiOri decimal(6, 3) ,
        .Parameters(18).Value = dlPerComiCalc
        '@TexNeg ntext ,
        .Parameters(19).Value = Trim(TxtNegocio.Text)
        '@TexObs ntext ,
        .Parameters(20).Value = Trim(TxtObserva.Text)
        '@ClasCor char(1),
        
        If Status.BackColor = &HFF00& Then
            
            .Parameters(21).Value = "G"
            
        Else
        
            If Status.BackColor = &HFF& Then
                .Parameters(21).Value = "R"
            Else
                .Parameters(21).Value = "B"
            End If
            
        End If
        
        '@IdxPDD decimal(6, 3),
        .Parameters(22).Value = dlIdxPDD
        '@NomTra varchar(40) ,
        .Parameters(23).Value = Trim(TxtTransp.Text)
        '@ChvDsc  Varchar(20)
        .Parameters(24).Value = Trim(slChave)
        '@SitPed char(1)
        .Parameters(25).Value = "N"
        '@DatPed char(1)
        .Parameters(26).Value = Datped
        '@FlgKit char(1)
        
        If ChkKit.Value = 1 Then
            .Parameters(27).Value = "S"
        Else
            .Parameters(27).Value = "N"
        End If
        
        '@VlrSimples decimal(10, 2),
        .Parameters(28).Value = dlSimples
        
        '@FlgAlt char(1),
        If Trim(slFlgAlt) = "" Then
            
            If iDscForaRegiao >= 1 Then
               If dlMargemGeral < 9 Then
                  .Parameters(29).Value = "F"
               Else
                  .Parameters(29).Value = Null
               End If
            Else
               .Parameters(29).Value = Null
            End If
        
    '        .Parameters(29).Value = Null
        Else
            .Parameters(29).Value = "A"
        End If
            
        '@FlagOper char(1)
        .Parameters(30).Value = Trim(sgFlagOper)
        
    End With
    
    Set Rs = Cmd.Execute
    Set Rs = Nothing
    Set Cmd = Nothing
    
    'DELETA TODOS OS ITENS
    sgQuery = "Delete ITEM_PEDIDO where NroPed = '" & Trim(MskNroPedido.Text) & "'"
    
    Conexao.Execute sgQuery
   
    'GRAVA ITENS DO PEDIDO
    
    If GrdNotaCliente.rows > 1 Then
    
        For Linhas = 1 To GrdNotaCliente.rows - 1
        
            If Trim(GrdNotaCliente.TextMatrix(Linhas, 0)) > 0 Then
            
                Set Cmd = New Command
                
                With Cmd
                
                    .CommandText = "{call MNITEMPEDIDO (?,?,?,?,?,?,?,?,?,?,?,?,?)}"
                    .CommandType = adCmdText
                    
                    .ActiveConnection = Conexao
                    
                    .Parameters.Refresh
                    
                    '@NroPed varchar(7),
                    .Parameters(0).Value = Trim(MskNroPedido.Text)
                    '@CodPrd int,
                    .Parameters(1).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 0))
                    '@SeqIte int,
                    .Parameters(2).Value = Linhas
                    '@QtdPrd int,
                    .Parameters(3).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 3))
                    '@QtdEmb int,
                    .Parameters(4).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 2))
                    '@ValUnt decimal(10, 2),
                    .Parameters(5).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 4))
                    '@IdxDsc decimal(6, 3) ,
                    .Parameters(6).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 6))
                    '@VlrIte decimal(10, 2),
                    .Parameters(7).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 18))
                    '@FlgTab char(1)       ,
                    .Parameters(8).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 5))
                    '@ValUntN decimal(10, 2),
                    .Parameters(9).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 12))
                    '@MrgPrd decimal(6, 3) ,
                    .Parameters(10).Value = Trim(GrdNotaCliente.TextMatrix(Linhas, 13))
                    '@ValCusUnt decimal(10, 4),
                    .Parameters(11).Value = Trim(GrdIndice.TextMatrix(Linhas, 9))
                    '@IdxFix decimal(6, 3),
                    .Parameters(12).Value = Trim(GrdIndice.TextMatrix(Linhas, 10))
                    
                End With
                
                Set Rs = Cmd.Execute
                Set Rs = Nothing
                Set Cmd = Nothing
                
            End If
            
        Next Linhas
        
    End If
    
   Conexao.CommitTrans
   
    Exit Function
   
TrataErro:

    Rotina_Erro "GravaCTRC"

End Function

Function LimpaGeral()

    'BtoLimpaNF_Click
    LimpaGridAuxiliar
    ControleLostFocus = True
    blRetornoDupls = False
    sgFlagOper = "I"
    slFlgAlt = ""
    Datped = CDate(Date)

    CboCondPag.Habilitado = True
    Opt_CIF.Enabled = True
    Opt_FOB.Enabled = True
    TxtTransp.Text = ""
    blVencidos = False
    'ChkKit.Enabled = True
    LblVlSimples.Visible = False
    LblSimples.Visible = False
    '''MskNumNf.Enabled = False
    ''BtoProduto.Enabled = False
    'MskSerie.Enabled = False
    ''Bto_Aplica.Enabled = False
    'MskVlrUnit.Enabled = False
    'MskDatEmiNf.Enabled = False
    'VSValUnit.Enabled = False
    ''BtoAdiNF.Enabled = False
    ''BtoExcNF.Enabled = False
    ''BtoLimpaNF.Enabled = False
    BtoGrava.Enabled = False
    GrdNotaCliente.Enabled = False
    
    If APLICA = 1 Then
        SSTConhec.TabVisible(3) = False
    Else
        SSTConhec.TabVisible(3) = False
    End If

    FraGrupo.Visible = False
    LblC.Visible = False
    CmdImpr.Enabled = False
    ChkKit.Value = 0

    blEbahia = False
    slPedSimples = ""
    bgBloqPed = False
    
    LblBloqPed.Caption = "Este pedido não pode ser alterado"
    LblBloqPed.Visible = False
    CboCondPag.Criterio = ""
    
    LimpaLinhaNF

    GrdNotaCliente.rows = 1
    GrdIndice.rows = 1
    LblRotaRec.Caption = ""
    TxtNegocio.Text = ""
    'lblNegocio.Visible = False
    'TxtNegocio.Visible = False
    Opt_FOB.Enabled = False
    Opt_CIF.Enabled = False
    
    ilCodCnd = 0
    ilQtdPar = 0
    dlValUntN = 0
    dlValUntA = 0
    dlValUntB = 0
    dlValItem = 0
    dlPerComiNeg = 0
    QtdTubo = 0
    QtdTuboRosc = 0
    QtdAspe = 0
    QtdConx = 0
    slClasCor = ""
    dlPerDesPadrao = 0
    ilNumTab = 0
    blModificar = False
    blleitura = False
    dlPerDesFOBReal = 0
    slUFOri = "MG"

    d1.Texto = 0
    d2.Texto = 0
    d3.Texto = 0
    d4.Texto = 0
    d5.Texto = 0
    d6.Texto = ""
    d7.Texto = ""
    d8.Texto = 0
    d9.Visible = False
    vl1.Texto = 0
    vl2.Texto = 0
    vl3.Texto = 0
    MskMargem.Texto = 0
    T100Ped.Texto = 0
    T100Cli.Texto = 0
    T100Ped.Texto = 0

    ilDescChave = 0
    slChave = ""
    TxtChave.Text = ""
    dlPerCusFin = 0
    dlPerDesCnd = 0
    ilQtdParCnd = 0
    ilPrzMed = 0
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

    TxtNegocio.Text = ""
    TxtObserva.Text = ""
    
    slIrriga = ""
    dlPerComiCalc = dlPerComiN
    
    LblResultNegocio.Caption = ""
    Opt_CIF.Value = True
    TxtTransp.Text = "UNOCANN TRANSPORTES LTDA"
    LblSimBahia.Visible = False
    LblTextoideal.Visible = False
    vl3.Visible = False
    LblIdeal.Visible = False
    lblpercideal.Visible = False
    LblSub.Caption = ""
    LblTot.Caption = ""
    LblVlSimples.Caption = ""
    LblDesc.Caption = ""
    LblI.Caption = ""
    LblIN.Caption = ""
    LblIdeal.Caption = ""
    ''LblUnit.Caption = "Valor Unitário"
    MskMargem.Limpar
    Status.BackColor = &HC0C0C0
    Novacor.BackColor = &HC0C0C0
    SSTConhec.Tab = 0
    
    

    If bgPedMKT = True Then
        
        CboCondPag.SetFocus
        
    Else
        
        blLG = True
        slremet = ""
        
        GrdCliVencidos.rows = 1
        GrdCliVencer.rows = 1
        GrdCliJuros.rows = 1
        MskNroPedido.Text = ""
        cboCli.Criterio = ""
        cboCli.Habilitado = True
        MskNroPedido.Enabled = True
        SSTConhec.TabEnabled(1) = False
        SSTConhec.TabEnabled(2) = False
        SSTConhec.TabVisible(3) = False
        MskNroPedido.SetFocus
    
    End If

    'Dim sql As String
    
    'sql = "DROP TABLE ##tmp_produto_aux"
    
    'Conexao.Execute sql
    
    DoEvents
    
End Function

Function ConferePrazo() As Boolean

    Dim dlValorPedido    As Double
    Dim dlValorParcela   As Double
    LeituraCondicao

    ConferePrazo = True
    dlValorPedido = IIf(IsNull(LblTot.Caption), 0, LblTot.Caption)

    If dlValorPedido < dlValPedMin Then
        
        MsgBox "Valor do pedido abaixo do mínimo, [" & Format(dlValPedMin, "R$ ##,###,##0.00") & "]", vbExclamation + vbOKOnly, "Atenção!"
        
        'ConferePrazo = False
        
        Exit Function
        
    End If

    If ilQtdParCnd > 0 Then
        dlValorParcela = Format(dlValorPedido / ilQtdParCnd, "##,###,##0.00")
    End If
    
    If dlValorParcela < dlValParMin Then
        
        MsgBox "Valor de duplicata abaixo do mínimo, [" & Format(dlValParMin, "R$ ##,###,##0.00") & "], número de parcelas [" & ilQtdParCnd & "]", vbExclamation + vbOKOnly, "Atenção!"
        
        'ConferePrazo = False
        
        Exit Function
        
    End If

    If dlValorPedido < dlValLimPrz1 And ilPrzMed > ilPrzMed1 Then
        
        MsgBox "Prazo médio acima do permitido para valor do pedido, [" & ilPrzMed1 & "] dias", vbExclamation + vbOKOnly, "Atenção!"
        
        'ConferePrazo = False
        
        Exit Function
        
    End If

End Function

Function ValidaKit() As Boolean
    
    Dim ilResult As Double
    Dim ilResto   As Integer
 
    ValidaKit = True

    If QtdTubo < 500 Then
        
        If QtdTubo > 0 Then
            ilResult = 1
        Else
            ilResult = 0
        End If
        
    Else
        
        ilResto = QtdTubo Mod 500
        
        If ilResto = 0 Then
            ilResult = (QtdTubo / 500)
        Else
            ilResult = (QtdTubo / 500) + 1
        End If
    
    End If

    If QtdAspe = 0 Or QtdConx = 0 Or QtdTubo = 0 Then
        
        MsgBox "Kit de irrigação deve conter os TRÊS produtos, Tubo, Aspersor e Conexão", vbExclamation + vbOKOnly, "Atenção!"
        
        ValidaKit = False
        
        Exit Function
        
    End If

    If QtdAspe = 0 Or QtdConx < 2 Then
        
        MsgBox "Kit de irrigação deve conter aos menos 1 aspersor e 2 duas conexões", vbExclamation + vbOKOnly, "Atenção!"
        
        ValidaKit = False
        
        Exit Function
        
    End If
    
    ilResult = Int(ilResult)
    
    If QtdAspe < ilResult Then
        
        MsgBox "Kit de irrigação deve conter aos menos 1 aspersor para cada 500 tubos", vbExclamation + vbOKOnly, "Atenção!"
        
        ValidaKit = False
        
        Exit Function
        
    End If

    If QtdTuboRosc > 3 Then
        
        MsgBox "Kit de irrigação pode conter até 3 tipos de tubo Roscável", vbExclamation + vbOKOnly, "Atenção!"
        
        ValidaKit = False
        
        Exit Function
        
    End If

End Function

Function LeituraCondicao() As Boolean

    On Error GoTo TrataErro
    
    LeituraCondicao = True
    ilCodCnd = CboCondPag.codigo
    
    'MskSerie.Enabled = True
    ''Bto_Aplica.Enabled = True
    'MskDatEmiNf.Enabled = True
    
    If blModificar = False Then
        ''MskNumNf.Enabled = True
        ''BtoProduto.Enabled = True
        'MskNumNf.SetFocus
    End If
    
    'Leitura Condição de pagamento
    dlPerCusFin = 0
    dlPerDesCnd = 0
    ilQtdParCnd = 0
    ilPrzMed = 0
    
    sgQuery = "SELECT b.PerCusFin, b.PerDesCnd, a.QtdParCnd, a.PrzMed "
    sgQuery = sgQuery + " from CONDICAO a, CUSTO_CONDICAO b "
    sgQuery = sgQuery + "  Where a.CodCnd = " & Trim(ilCodCnd)
    sgQuery = sgQuery + "    and a.codcnd = b.codcnd"
    sgQuery = sgQuery + "    and b.datativ = (select max(datativ) from CUSTO_CONDICAO"
    sgQuery = sgQuery + "                      Where Codcnd = b.codcnd"
    sgQuery = sgQuery + "                        and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"
    
    Call Consulta(sgQuery)
    
    If Rs.EOF Then
    
        MsgBox "Erro na leitura da Condição de Pagamento", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        LimpaGeral
        
        Exit Function
        
    Else
    
        dlPerCusFin = IIf(IsNull(Rs!PerCusFin), 0, Trim(Rs!PerCusFin))
        dlPerDesCnd = IIf(IsNull(Rs!PerDesCnd), 0, Trim(Rs!PerDesCnd))
        ilQtdParCnd = IIf(IsNull(Rs!QtdParCnd), 0, Trim(Rs!QtdParCnd))
        ilPrzMed = IIf(IsNull(Rs!PrzMed), 0, Trim(Rs!PrzMed))
    
    End If
    
    Rs.Close
    
    Set Rs = Nothing
    
    CalculaDesconto
    
    If CalculaIndice = False Then
        LimpaGeral
    End If
      
    Exit Function

TrataErro:

    Rotina_Erro "LeituraCondicao"
    
    LeituraCondicao = False

End Function

Function LeituraPadrao() As Boolean
    
    '*****************************************************************************************
    'Esta função faz o levantamento de todos os descontos e tributações aplicáveis ao
    'representante logado ao Força no momento.
    '*****************************************************************************************
    
    On Error GoTo TrataErro

    LeituraPadrao = True

    '*****************************************************************************************
    'Apura o desconto promocional mais recente para o representante logado no momento.
    '*****************************************************************************************
    
    sgQuery = "select PerDsc from Desconto_promocional"
    sgQuery = sgQuery & "  Where CodRep = " & Trim(ilCodRep)
    sgQuery = sgQuery & "    and datativ = (select max(datativ) from Desconto_promocional"
    sgQuery = sgQuery & "                    Where CodRep = " & Trim(ilCodRep)
    sgQuery = sgQuery & "                      and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"
    
    Call Consulta(sgQuery)
    
    '*****************************************************************************************
    'Se encontrar algum desconto, guarda na variável dlPerDesRep. Se não encontrar desconto
    'específico para o representante logado, pesquisa para saber qual é o desconto mais
    'recente válido para todos os representantes. Se houver algum, guarda na mesma variável
    'citada acima.
    '*****************************************************************************************

    If Rs.EOF = False Then
        
        dlPerDesRep = IIf(Trim(Rs!PerDsc) = "", 0, Trim(Rs!PerDsc))
        
    Else
                
        Rs.Close
        
        Set Rs = Nothing
        
        sgQuery = "select PerDsc from Desconto_promocional"
        sgQuery = sgQuery & "  Where CodRep is null"
        sgQuery = sgQuery & "    and datativ = (select max(datativ) from Desconto_promocional"
        sgQuery = sgQuery & "                    Where CodRep is null"
        sgQuery = sgQuery & "                      and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"
        
        Call Consulta(sgQuery)
        
        If Rs.EOF = False Then
            dlPerDesRep = IIf(Trim(Rs!PerDsc) = "", 0, Trim(Rs!PerDsc))
        End If
        
    End If
    
    Rs.Close
    
    Set Rs = Nothing
   
    '*****************************************************************************************
    'Faz a leitura do desconto padrão. Levanta o desconto padrão mais recente de acordo com os
    'estados de origem e destino do pedido. A variável dlPerDesPrd, que armazena desconto
    'aplicável ao produto, recebe o valor do desconto promocional concedido ao representante
    'logado.
    '*****************************************************************************************
   
    dlPerDesPrd = dlPerDesRep
    dlPerContr = 0
    dlPerNContr = 0
    dlPerContrSIMBa = 0
    dlPerContrKit = 0
    dlPerNContrKit = 0
    dlPerContrSIMBaKit = 0
    dlPerDesPadrao = 0
    
    sgQuery = "SELECT * from DESCONTO_PADRAO"
    sgQuery = sgQuery & "   WHERE UFOri = '" & Trim(slUFOri) & "' "
    sgQuery = sgQuery & "     and UFDes = '" & slUFRep & "' "
    sgQuery = sgQuery & "     and datativ = (select max(datativ) from desconto_padrao"
    sgQuery = sgQuery & "                     WHERE UFOri = '" & Trim(slUFOri) & "' "
    sgQuery = sgQuery & "                       and UFDes = '" & slUFRep & "' "
    sgQuery = sgQuery & "                       and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"

    Call Consulta(sgQuery)
    
    '*****************************************************************************************
    'Deve sempre haver um desconto padrão para cada origem e destino onde a Unocann atua. Se
    'nenhum registro for encontrado é porque há algum problema. Caso contrário, se houver
    'desconto cadastrado para o perfil da venda em questão, os dados do desconto encontrado
    'são armazenados em variáveis.
    '*****************************************************************************************
    
    If Rs.EOF = True Then
        
        MsgBox "Erro na leitura do Desconto Padrão", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        LimpaGeral
        
        Exit Function
        
    Else
        
        dlPerContr = IIf(IsNull(Rs!PerContr), 0, Trim(Rs!PerContr))
        dlPerNContr = IIf(IsNull(Rs!PerNContr), 0, Trim(Rs!PerNContr))
        dlPerContrSIMBa = IIf(IsNull(Rs!PerContrSIMBa), 0, Trim(Rs!PerContrSIMBa))
        dlPerContrKit = IIf(IsNull(Rs!PerContrKit), 0, Trim(Rs!PerContrKit))
        dlPerNContrKit = IIf(IsNull(Rs!PerNContrKit), 0, Trim(Rs!PerNContrKit))
        dlPerContrSIMBaKit = IIf(IsNull(Rs!PerContrSIMBaKit), 0, Trim(Rs!PerContrSIMBaKit))
    
    End If
    
    Rs.Close
    
    Set Rs = Nothing

    If slFlgContr = "S" Then
        dlPerDesPadrao = dlPerContr
    Else
        dlPerDesPadrao = dlPerNContr
    End If

    '*****************************************************************************************
    'Faz a leitura das alíquotas de impostos. Levanta os tributos mais recentes de acordo com
    'os estados de origem e destino dos produtos.
    '*****************************************************************************************
    
    dlAlqICMContr = 0
    dlAlqICMNContr = 0
    dlAlqICMContrKIT = 0
    dlAlqICMNContrKIT = 0
    dlAlqICMSimplesKIT = 0
    PerCusIcm = 0

    sgQuery = "SELECT * from TRIBUTACAO"
    sgQuery = sgQuery & "   WHERE UFOri = '" & Trim(slUFOri) & "' "
    sgQuery = sgQuery & "     and UFDes = '" & slUFCli & "' "
    sgQuery = sgQuery & "     and datativ = (select max(datativ) from TRIBUTACAO"
    sgQuery = sgQuery & "                     WHERE UFOri = '" & Trim(slUFOri) & "' "
    sgQuery = sgQuery & "                       and UFDes = '" & slUFCli & "' "
    sgQuery = sgQuery & "                       and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"

    Call Consulta(sgQuery)
    
    '*****************************************************************************************
    'Deve sempre haver uma definição de tributos para cada origem e destino onde a Unocann
    'atua. Se nenhum registro for encontrado é porque há algum problema. Caso contrário, se
    'houverem tributos cadastrados para o perfil da venda em questão, os dados desses tributos
    'encontrados são armazenados em variáveis.
    '*****************************************************************************************
    
    If Rs.EOF = True Then
        
        MsgBox "Erro na leitura da Tabela de Tributação", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        LimpaGeral
        
        Exit Function
    
    Else
        
        If slPedSimples = "S" Then
            dlAlqICMContr = IIf(IsNull(Rs!AlqICMSimples), 0, Trim(Rs!AlqICMSimples))
        Else
            dlAlqICMContr = IIf(IsNull(Rs!AlqICMContr), 0, Trim(Rs!AlqICMContr))
        End If
        
        dlAlqICMNContr = IIf(IsNull(Rs!AlqICMNContr), 0, Trim(Rs!AlqICMNContr))
        dlAlqICMContrKIT = IIf(IsNull(Rs!AlqICMContrKit), 0, Trim(Rs!AlqICMContrKit))
        dlAlqICMNContrKIT = IIf(IsNull(Rs!AlqICMNContrKit), 0, Trim(Rs!AlqICMNContrKit))
        dlAlqICMSimplesKIT = IIf(IsNull(Rs!AlqICMSimplesKit), 0, Trim(Rs!AlqICMSimplesKit))
    
    End If
    
    Rs.Close
    
    Set Rs = Nothing

    If slFlgContr = "S" Then
        PerCusIcm = dlAlqICMContr
    Else
        PerCusIcm = dlAlqICMNContr
    End If

    Exit Function
    
    '*****************************************************************************************
    'Tratamento de erros.
    '*****************************************************************************************
    
TrataErro:
    
    Rotina_Erro "LeituraPadrao"
    
    LeituraPadrao = False
    
End Function

Function LeituraCliente() As Boolean

    '*************************************************************************************
    'Função que faz a leitura de informações referentes ao cliente em questão. Os dados
    'são exibidos na guia POSIÇÃO DO CLIENTE.
    '*************************************************************************************

    On Error GoTo TrataErro

    Dim slCGCCli  As String
    Dim dlSumVen  As Double
    Dim dlSumAVen As Double
    Dim dlSumJur  As Double
    Dim dlQtdVen  As Integer
    Dim dlQtdAVen As Integer
    Dim dlQtdJur  As Integer

    LeituraCliente = True
        
    blleitura = True
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
  
    '*************************************************************************************
    'Verifica se está ocorrendo apenas uma simulação. Se for o caso, desvia a execução
    'para a rotina apropriada.
    '*************************************************************************************
  
    If bgSimula = True Then
        GoTo Simulacao
    End If
  
    '*************************************************************************************
    'Levanta informações acerca do cliente em questão.
    '*************************************************************************************
  
    sgQuery = "SELECT * from CLIENTE WHERE CodCli = " & Trim(ilCodCli)
    
    Call Consulta(sgQuery)
    
    If Rs.EOF Then
     
        MsgBox "Erro na leitura do cliente", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        LimpaGeral
        
        LeituraCliente = False
        
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
        LblCliCep = IIf(IsNull(Rs!CepCli), "", Trim(Rs!CepCli))
        LblCliUF = IIf(IsNull(Rs!UFCli), "", Trim(Rs!UFCli))
        LblFone = IIf(IsNull(Rs!FonCli), "", Trim(Rs!FonCli))
        LblCliPrimeira = Format(Rs!DatPriComp, "dd/mm/yyyy")
        dlPerTubo100Cli = IIf(IsNull(Rs!PerTubo100), 0, Trim(Rs!PerTubo100))
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
        
        T100Rep.Texto = dlPerTubo100Rep
        T100Cli.Texto = dlPerTubo100Cli
    
    End If
    
    Rs.Close
    
    Set Rs = Nothing
  
    If sgFlagOper <> "A" Then
        slPedSimples = ""
    End If
  
    If Trim(slPedSimples) = "" Then
        
        slPedSimples = "N"
        
        If slUFRep = "BA" And slFlgSIMBa = "S" Then
            slPedSimples = "S"
        End If
  
    End If
  
    If bgBloqPed = False And slFlgSIMBa <> "S" Then
        slPedSimples = "N"
    End If
  
    If bgBloqPed = False Then
        dlPerComiNeg = 0
    End If
  
    If Trim(slPedSimples) = "S" Then
        slUFOri = "BA"
        LblVlSimples.Visible = True
        LblSimples.Visible = True
    Else
        LblVlSimples.Visible = False
        LblSimples.Visible = False
    End If
    
    '*****************************************************************************************
    'Realiza levantamento de descontos e alíquotas de tributos aplicáveis ao pedido.
    '*****************************************************************************************
  
    If LeituraPadrao = False Then
        
        LeituraCliente = False
        
        Exit Function
        
    End If
    
    '*****************************************************************************************
    'Calcula todos as informações que apontam se o pedido é viável ou não.
    'UPDATE 03/07/2017 Não tem necessidade de calcular o indice neste momento
    '*****************************************************************************************
    
    'If CalculaIndice = False Then
        
        'LimpaGeral
        
        'Exit Function
        
    'End If
  
    '*****************************************************************************************
    'Consulta e exibe na guia "Posição do Cliente" informações sobre títulos vencidos do
    'cliente selecionado.
    '*****************************************************************************************
  
    sgQuery = "select nrodup, parc, datemi, datven, vlrdup, datediff(dd,datven, getdate()) as dias from DUPLICATA "
    sgQuery = sgQuery + " where datpag is null and datven < (getdate() - 1)"
    sgQuery = sgQuery + "   AND CodCli = " & Trim(ilCodCli)

    Call Consulta(sgQuery)
    
    GrdCliVencidos.rows = 1
    
    ilind = GrdCliVencidos.rows

    Do While Not Rs.EOF
        
        GrdCliVencidos.rows = GrdCliVencidos.rows + 1
        GrdCliVencidos.TextMatrix(ilind, 0) = Format(Trim(Rs!NroDup), "000,000")
        GrdCliVencidos.TextMatrix(ilind, 1) = Trim(Rs!Parc)
        GrdCliVencidos.TextMatrix(ilind, 2) = Format(Trim(Rs!datemi), "dd/mm/yyyy")
        GrdCliVencidos.TextMatrix(ilind, 3) = Format(Trim(Rs!datven), "dd/mm/yyyy")
        GrdCliVencidos.TextMatrix(ilind, 4) = Format(Trim(Rs!VlrDup), "##,###,###,##0.00")
        GrdCliVencidos.TextMatrix(ilind, 5) = Format(Trim(Rs!dias), "##,##0")
        
        dlSumVen = dlSumVen + Trim(Rs!VlrDup)
        dlQtdVen = dlQtdVen + 1
        ilind = ilind + 1
        blVencidos = True
  
        Rs.MoveNext
    
    Loop

    Rs.Close
    
    Set Rs = Nothing

    LblSumVen.Caption = Format(dlSumVen, "##,###,###,##0.00")
    LblQtdVen.Caption = Format(dlQtdVen, "00")
  
    '*****************************************************************************************
    'Consulta e exibe na guia "Posição do Cliente" informações sobre títulos a vencer do
    'cliente selecionado.
    '*****************************************************************************************
    
    sgQuery = "select nrodup, parc, datemi, datven, vlrdup from DUPLICATA "
    sgQuery = sgQuery + " where datpag is null and datven >= (getdate() - 1)"
    sgQuery = sgQuery + "   AND CodCli = " & Trim(ilCodCli)

    Call Consulta(sgQuery)
    
    GrdCliVencer.rows = 1
    
    ilind = GrdCliVencer.rows

    Do While Not Rs.EOF
        
        GrdCliVencer.rows = GrdCliVencer.rows + 1
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
    
    '*****************************************************************************************
    'Consulta e exibe na guia "Posição do Cliente" informações sobre juros em aberto para o
    'cliente selecionado.
    '*****************************************************************************************
    
    sgQuery = "select nrodup, parc, datpag, datven, vlrdup, JurDev from DUPLICATA "
    sgQuery = sgQuery + " where datjur is null and JurDev > 0 "
    sgQuery = sgQuery + "   AND CodCli = " & Trim(ilCodCli)

    Call Consulta(sgQuery)
    
    GrdCliJuros.rows = 1
    
    ilind = GrdCliJuros.rows

    Do While Not Rs.EOF
        
        GrdCliJuros.rows = GrdCliJuros.rows + 1
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
    
    '*****************************************************************************************
    '
    '*****************************************************************************************
    
    'Ultimo Faturamento
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

Simulacao:

    '*****************************************************************************************
    '
    '*****************************************************************************************
  
    If Trim(CboUFCli.Text) = "" Then
    
        MsgBox "Informe a UF do cliente", vbExclamation + vbOKOnly, "Atenção!"
        
        LeituraCliente = False
        
        Exit Function
        
    End If

    If Trim(CboContribuinte.Text) = "" Then
        
        MsgBox "Informe <Sim/Não> para cliente contribuinte", vbExclamation + vbOKOnly, "Atenção!"
        
        LeituraCliente = False
        
        Exit Function
        
    End If

    If (Trim(CboUFCli.Text) = "BA" And Trim(CboContribuinte.Text) = "Sim") And Trim(CboSimBA.Text) = "" Then
        
        MsgBox "Informe <Sim/Não> para cliente Simples(BA)", vbExclamation + vbOKOnly, "Atenção!"
        
        LeituraCliente = False
        
        Exit Function
        
    End If

    slUFCli = Trim(CboUFCli.Text)
  
    If Trim(CboContribuinte.Text) = "Sim" Then
        slFlgContr = "S"
    Else
        slFlgContr = "N"
    End If
  
    If Trim(CboUFCli.Text) <> "" Then
        
        If Trim(CboSimBA.Text) = "Sim" Then
            slFlgSIMBa = "S"
            slPedSimples = "S"
        Else
            slFlgSIMBa = "N"
            slPedSimples = "N"
        End If
        
    End If

    dlPerComiNeg = 0
  
    If slFlgContr = "N" Then
        slPedSimples = "N"
        CboSimBA.Text = "Não"
    End If
      
    If slUFCli = "ES" Then
        
        ilCodRep = 7000
        
    Else
        
        If slUFCli = "MG" Then
            
            ilCodRep = 2100
            
        Else
        
            If slUFCli = "BA" Then
                ilCodRep = 5000
            Else
                ilCodRep = 2
                slUFCli = "MG"
            End If
        
        End If
    
    End If
        
    If Trim(CboUFCli.Text) <> "BA" Then
        slPedSimples = "N"
        CboSimBA.Text = "Não"
    End If
  
    sgQuery = "select * from REPRESENTANTE where CodRep = " & Trim(ilCodRep)
    
    Call Consulta2(sgQuery)
    
    If Not Rs2.EOF Then
        
        slUFRep = IIf(IsNull(Rs2!UFRep), "", Trim(Rs2!UFRep))
        dlPerCusFrt = IIf(IsNull(Rs2!PerCusFrt), 0, Trim(Rs2!PerCusFrt))
        dlPerComiN = 0
        dlPerComiA = 0
        dlPerComiB = 0
        dlPerDesFOB = IIf(IsNull(Rs2!PerDesFOB), 0, Trim(Rs2!PerDesFOB))
        dlPerComiCalc = dlPerComiN
        
    Else
    
        MsgBox "Registro do Representante não encontrado, informe ao administrador do sistema", vbExclamation + vbOKOnly, "Atenção!"
        
    End If
  
    Rs2.Close
    
    Set Rs2 = Nothing
  
    If LeituraPadrao = False Then
        
        LeituraCliente = False
        
        Exit Function
        
    End If
  
    '***********************************************************************************
    'UPDATE: 03/07/2017 Não tem necessidade de calcular indice neste momento
    '***********************************************************************************
    
    'If CalculaIndice = False Then
        
        'LimpaGeral
        'Exit Function
        
    'End If
    
    Exit Function

TrataErro:

    Rotina_Erro "LeituraCliente"
    
    LeituraCliente = False

End Function

Function PegaSenha() As Boolean

    Do While bgSenComi = False
    
        TxtPwdoper.SetFocus

        DoEvents
    
    Loop

End Function

Function EquilibraComissao() As Boolean

    On Error GoTo TrataErro

    EquilibraComissao = True
    PerSugIni = 0
    PerSugFimA = 0
    PerSugFimB = 0
    blFechaComi = False

    'If dlINlb = 0 Then
        'Exit Function
    'End If

    If dlINlb > 0 And dlIdeallb > 0 And dlIlb > 0 Then
        
        'Status.BackColor = &HC00000
        
        Exit Function
    
    End If
   
    If dlIlb < 0 Then
        
        ilNroSit = 1
        PerSugIni = PerSug1Ini
        PerSugFimA = PerSug1FimA
        PerSugFimB = PerSug1FimB
        'dlIABS = Abs(dlIlb)
        
    Else
        
        ilNroSit = 2
        PerSugIni = PerSug2Ini
        PerSugFimA = PerSug2FimA
        PerSugFimB = PerSug2FimB
        'dlIABS = Abs(dlINlb)
        
    End If
    
    If Trim(SlTabela) = "" Then
        
        dlPerComiNeg = 0
        
        Exit Function
    
    End If

    If dlPerComiNeg > 0 Then
        
        dlComiSug = dlPerComiNeg
    
    Else
        
        If SlTabela = "M" Then
            
            If dlIABS < PerSugIni Then
                
                'Status.BackColor = &HFF00&
                
                Exit Function
                
            Else
                
                If dlIABS >= PerSugIni And dlIABS <= PerSugFimA Then
                    
                    dlComiSug = dlPerComiA
                    
                Else
                    
                    If dlIABS > PerSugFimA And dlIABS <= PerSugFimB Then
                        dlComiSug = dlPerComiB
                    Else
                        Exit Function
                    End If
                    
                End If
            
            End If
            
        Else
        
            If SlTabela = "A" Then
                dlComiSug = dlPerComiA
            Else
                dlComiSug = dlPerComiB
            End If
        
        End If
    
    End If

    slAceita = True
    
    If SlTabela = "M" Then
        sgQuery = MsgBox("Aceita comissão de " & Trim(dlComiSug) & "% visando o equilíbrio deste pedido ?", vbQuestion + vbYesNo + vbDefaultButton2, "Atenção!")
    Else
        sgQuery = MsgBox("A Comissão para este pedido será de " & Trim(dlComiSug) & "% Tabela - " & SlTabela, vbExclamation + vbOKOnly, "Atenção!")
    End If
    
    If sgQuery = vbYes Or sgQuery = vbOK Then
        
        'lblmens.Visible = False
        
        bgSenComi = False
        
        'TxtPwdoper.Text = ""
        'FraSenha.Visible = True
        'TxtPwdoper.SetFocus
        
        PegaSenha
        
        'If Me.ActiveControl.Name = "BtoSair" Or _

            'Me.ActiveControl.Name = "BtoLimpaCTRC" Then
            
            'EquilibraComissao = False
            
            'Exit Function
        
        'End If
        
        'FraSenha.Visible = False
        
        If Not bgSenOK Then
            
            dlPerComiNeg = 0
            slClasCor = ""
            
            If CalculaIndice = False Then
                LimpaGeral
            End If
            
            If SlTabela <> "M" Then
                blFechaComi = False
                EquilibraComissao = False
            End If
            
            LblResultNegocio.Caption = ""
        
        Else
        
            dlPerComiNeg = Trim(dlComiSug)
            
            If CalculaIndice = False Then
                LimpaGeral
            End If
            
            If SlTabela = "M" Then
                
                'Status.BackColor = &HFF00&
                
            Else
                
                blFechaComi = True
                
                If dlIlb > 0 Then
                    'Status.BackColor = &HFF00&
                End If
            
            End If
            
            LblResultNegocio.Caption = "Comissão do pedido negociada com o representante em " & Trim(dlComiSug) & "%"
        
        End If
        
    Else
        
        dlPerComiNeg = 0
        slClasCor = ""
        
        If CalculaIndice = False Then
            LimpaGeral
        End If
        
        If SlTabela <> "M" Then
            
            blFechaComi = False
            
            EquilibraComissao = False
            
        End If
        
        LblResultNegocio.Caption = ""
        
    End If

    Me.Refresh

    Exit Function

TrataErro:

    Rotina_Erro "EquilibraComissao"
    
End Function

Function CalculaDesconto() As Boolean

    Dim dlValUntN As Double
    Dim dlValItem As Double
    Dim ilQtde As Double
    Dim ilQtdEmb As Integer
    
    CalculaDesconto = True
    
    '*****************************************************************************************
    'Se o cliente for contribuinte, o sistema avalia se ele é optante pelo SIMPLES da Bahia.
    'Se for optante, o desconto padrão será aquele permitido pelo SIMPLES; se não, tal
    'desconto será aquele destinado aos contribuintes normais. Se o cliente não for
    'contribuinte, o desconto padrão será aquele reservado aos compradores nessa situação.
    '*****************************************************************************************
    
    If slFlgContr = "S" Then
        
        If slFlgSIMBa = "S" Then
            dlPerDesPadrao = dlPerContrSIMBa
        Else
            dlPerDesPadrao = dlPerContr
        End If
    
    Else
         
         dlPerDesPadrao = dlPerNContr
         
    End If
    
    '*****************************************************************************************
    'Avalia se o pedido é de Kit Irrigação. Se for, o sistema define a origem do pedido em MG,
    'define o pedido como não-optante pelo SIMPLES e refaz a leitura padrão com esses
    'parâmetros, a fim de levantar descontos compatíveis com a compra do cliente. A partir
    'disso são definidos os percentuais de desconto padrão e custo de ICMS. Os campos que
    'informam desconto de SIMPLES na totalização do pedido são omitidos.
    '*****************************************************************************************
    
    '*****************************************************************************************
    'Se o pedido não for de Kit Irrigação, o programa avalia se o representante logado atende
    'o mercado da Bahia e se o cliente é optante pelo SIMPLES baiano. Se a condição for aceita
    'o cliente é definido como optante pelo SIMPLES e a origem do pedido é definida no estado
    'da Bahia. Os campos que informam desconto de SIMPLES na totalização do pedido são
    'exibidos.
    '*****************************************************************************************
    
    If ChkKit.Value = 1 Then
            
        slUFOri = "MG"
        slPedSimples = "N"
            
        LblVlSimples.Visible = False
        LblSimples.Visible = False
        
        '*****************************************************************************************
        'LeituraPadrao faz o levantamento de todos os descontos e tributações aplicáveis ao
        'representante logado ao Força no momento.
        '*****************************************************************************************
        
        LeituraPadrao
            
        If slFlgContr = "S" Then
                
            If slPedSimples = "S" Then
                dlPerDesPadrao = dlPerContrSIMBaKit
            Else
                dlPerDesPadrao = dlPerContrKit
            End If
            
        Else
                
            dlPerDesPadrao = dlPerNContrKit
            
        End If
            
        If slFlgContr = "S" Then
            PerCusIcm = dlAlqICMContrKIT
        Else
            PerCusIcm = dlAlqICMNContrKIT
        End If
        
    Else
        
        If slUFRep = "BA" And slFlgSIMBa = "S" Then
                
            slPedSimples = "S"
            slUFOri = "BA"
                
            LblVlSimples.Visible = True
            LblSimples.Visible = True
                
        End If
        
    End If
    
    '*****************************************************************************************
    'Armazena todos os descontos encontrados num array e soma todos eles. O primeiro desconto
    'é aplicado a 100% do valor do produto; os demais são aplicados ao valor do produto após a
    'aplicação dos outros descontos encontrados.
    '*****************************************************************************************
    
    dlCem = 100
    dlSumDscItem = 0
    dlSumDscItemORIG = 0
    
    '*****************************************************************************************
    'Se o cliente for pessoa jurídica, aplica-se o desconto promocional. Se for pessoa física,
    'o desconto promocional é suprimido da lista dos descontos. A mudança ainda precisa de
    'autorização do Jacson para vigorar.
    '*****************************************************************************************
    
    'If Len(LblCliCGC.Caption) = 14 Then
    
        vlPercs = Array(dlPerDesPadrao, dlPerDesPrd, dlPerDesCnd, dlPerDesFOBReal)
            
        For ilind = 0 To 3
        
            dlCem = 100 - dlSumDscItem
        
            If vlPercs(ilind) > 0 Then
                dlSumDscItem = dlSumDscItem + ((vlPercs(ilind) * dlCem) / 100)
            End If
        
        Next
        
    'Else
    
        'vlPercs = Array(dlPerDesPadrao, dlPerDesCnd, dlPerDesFOBReal)
            
        'For ilind = 0 To 2
        
            'dlCem = 100 - dlSumDscItem
        
            'If vlPercs(ilind) > 0 Then
                'dlSumDscItem = dlSumDscItem + ((vlPercs(ilind) * dlCem) / 100)
            'End If
        
        'Next
        
    'End If
    
    '*****************************************************************************************
    'A variável dlSumDscItemORIG guarda o valor original do desconto encontrado, já que aquele
    'armazenado na variável dlSumDscItem pode sofrer modificação imposta por chave.
    '*****************************************************************************************
    
    dlSumDscItemORIG = dlSumDscItem
    
    '*****************************************************************************************
    'Verifica se foi cedida chave de desconto para o pedido. Se tal chave existir, ela
    'despreza quaisquer descontos concedidos anteriormente. O desconto da chave é informado
    'em um label que é exibido nos parâmetros ocultos de descontos, ao lado do desconto total.
    '*****************************************************************************************
    
    '*****************************************************************************************
    'Os valores das variáveis slChave e ilDescChave são definidos na decriptação das chaves.
    '*****************************************************************************************
    
    If Trim(slChave) <> "" Then
            
        dlSumDscItem = ilDescChave
            
        LblC.Caption = dlSumDscItem
        LblC.Visible = True
            
    End If
  
    dlSumDscItem = Format(dlSumDscItem, "##0.00")
    
    '*****************************************************************************************
    'O loop a seguir define os preços dos itens do pedido com os descontos encontrados. Seria
    'o preço líqüido, que fica armazenado na coluna 11 do grid.
    '*****************************************************************************************
    
    '*****************************************************************************************
    'A matemática para chegar ao preço líqüido consiste em:
    '1) Encontrar o valor monetário do desconto, que até então existe apenas em valor
    '   percentual;
    '2) Subtrair o preço bruto do produto pelo desconto para obter seu preço unitário líqüido;
    '3) Multiplicar a quantidade de itens vendidos pela quantidade comportada por cada
    '   embalagem;
    '4) Multiplicar o preço líqüido do produto pela quantidade de unidades vendidas.
    '*****************************************************************************************
    
   For Linhas = 1 To GrdNotaCliente.rows - 1
            
        ilQtde = Trim(GrdNotaCliente.TextMatrix(Linhas, 3))
        ilQtdEmb = Trim(GrdNotaCliente.TextMatrix(Linhas, 2))
        dlValUntN = Trim(GrdNotaCliente.TextMatrix(Linhas, 12))
        'VlIdealItem = (dlValUntN - ((dlValUntN * dlSumDscItemORIG) / 100)) * (ilQtde * ilQtdEmb)
        VlIdealItem = (dlValUntN - ((dlValUntN * dlSumDscItemORIG) / 100)) * ilQtde 'Alterado a pedido do Manoel 17/04/2017
            
        GrdNotaCliente.TextMatrix(Linhas, 11) = Format(VlIdealItem, "###,##0.00")
        
    Next Linhas

    Me.Refresh
    
    '*****************************************************************************************
    'Exibe nos parâmetros ocultos de descontos as informações relativas ao assunto encontradas
    'na base.
    '*****************************************************************************************
  
    d1.Texto = dlPerDesPadrao
    d2.Texto = dlPerDesPrd
    d3.Texto = dlPerDesCnd
    d4.Texto = dlPerDesFOBReal
    d5.Texto = dlSumDscItemORIG
    d6.Texto = slFlgContr
    d7.Texto = slUFCli
    d8.Texto = PerCusIcm
'-----------------------------------------------------------------
'    If ilCodCnd = 1 Or ilCodCnd = 12 Then 'A vista ou 14 dias
'        dDscRegiao = IIf(Trim(d5.Texto) = "", 0, Trim(d5.Texto))
'        bAVista = True
'    Else 'A Prazo
'        dDscRegiao = IIf(Trim(d1.Texto) = "", 0, Trim(d1.Texto))
'        bAVista = False
'    End If
'-----------------------------------------------------------------
        
    If slFlgSIMBa = "S" Then
        d9.Visible = True
    Else
        d9.Visible = False
    End If
  
    Me.Refresh
  
End Function

Function AplicaDescontoGrid() As Boolean

    Dim dlValItem As Double
    Dim dlValUnit As Double
    Dim dlValUntN As Double
    Dim ilQtde    As Double
    Dim ilQtdEmb  As Integer
    Dim dlValDesc As Double
  
    AplicaDescontoGrid = True
  
    dlValDesc = IIf(Trim(MskDatEmiNf) = "", 0, Trim(MskDatEmiNf))
  
    If dlValDesc > dlSumDscItem Then
        
        MsgBox "Desconto informado maior que o permitido para esta venda.", vbExclamation + vbOKOnly, "Atenção!"
        
        AplicaDescontoGrid = False
        
        Exit Function
        
    End If

    For Linhas = 1 To GrdNotaCliente.rows - 1
    
        dlValUnit = Trim(GrdNotaCliente.TextMatrix(Linhas, 4))
        ilQtde = Trim(GrdNotaCliente.TextMatrix(Linhas, 3))
        ilQtdEmb = Trim(GrdNotaCliente.TextMatrix(Linhas, 2))
        dlValUntN = Trim(GrdNotaCliente.TextMatrix(Linhas, 12))
       ' dlValItem = (dlValUnit - ((dlValUnit * dlValDesc) / 100)) * (ilQtde * ilQtdEmb)
        dlValItem = (dlValUnit * (1 - (dlValDesc / 100))) * ilQtde
        
        GrdNotaCliente.TextMatrix(Linhas, 18) = Format(dlValItem, "##,###,##0.00")
        GrdNotaCliente.TextMatrix(Linhas, 6) = Format(dlValDesc, "##,###,##0.00")
  '-----------------------------------------------------------------------------
        If bAVista = True Then 'estou trabalhando a vista
            
                If GrdNotaCliente.TextMatrix(Linhas, 6) <= dDscRegiao Then
                    iDscRegiao = iDscRegiao + 1
                Else
                    iDscForaRegiao = iDscForaRegiao + 1
                End If
                
             Else
            
                If GrdNotaCliente.TextMatrix(Linhas, 6) <= dDscRegiao Then
                    iDscRegiao = iDscRegiao + 1
                Else
                    iDscForaRegiao = iDscForaRegiao + 1
                End If
            
            
            End If
        
  '-----------------------------------------------------------------------------
        
        
        VlIdealItem = (dlValUntN - ((dlValUntN * dlSumDscItemORIG) / 100)) * ilQtde
        
        GrdNotaCliente.TextMatrix(Linhas, 11) = Format(VlIdealItem, "###,##0.00")
        
    Next Linhas

    Me.Refresh

End Function

Function RecalculaDescontoGrid() As Boolean

    Dim dlValItem As Double
    Dim dlValUnit As Double
    Dim dlValUntN As Double
    Dim ilQtde    As Double
    Dim ilQtdEmb  As Integer
    Dim dlValDesc As Double
    
    '*****************************************************************************************
    'Se o pedido não puder sofrer alterações, esta rotina não pode ser executada. O programa
    'irá abandoná-la.
    '*****************************************************************************************
    
    If bgBloqPed = True Then
        Exit Function
    End If
    
    '*****************************************************************************************
    'Executa o recálculo para cada item de produto.
    '*****************************************************************************************
    
    For Linhas = 1 To GrdNotaCliente.rows - 1
    
        '*************************************************************************************
        'Armazena em variável o desconto aplicado ao produto.
        '*************************************************************************************
    
        dlValDesc = Trim(GrdNotaCliente.TextMatrix(Linhas, 6))
        
        '*************************************************************************************
        'Avalia se o desconto concedido ao produto em questão é maior que a soma dos descontos
        'encontrados para o perfil da venda.
        '*************************************************************************************
        
        '*************************************************************************************
        'Em 25/09/2008 eu ainda não entendi o porque da necessidade de se fazer esse
        'recálculo; um exemplo com desconto de produto pode me ajudar nisso.
        '*************************************************************************************
        
'        If dlValDesc > dlSumDscItem Then
'
'            If GrdNotaCliente.RowSel > 0 Then
'
'                If GrdNotaCliente.Rows > 2 Then
'
'                    GrdNotaCliente.RemoveItem Linhas 'GrdNotaCliente.RowSel
'
'                Else
'
'                    GrdNotaCliente.Rows = 1
'                    'ChkKit.Enabled = True
'                    GrdNotaCliente.Enabled = False
'
'                End If
'
'            End If
'
'            TrocaCIF_FOB
'
'        Else
'
            
            dlValUnit = Trim(GrdNotaCliente.TextMatrix(Linhas, 4))
            ilQtde = Trim(GrdNotaCliente.TextMatrix(Linhas, 3))
            ilQtdEmb = Trim(GrdNotaCliente.TextMatrix(Linhas, 2))
            dlValUntN = Trim(GrdNotaCliente.TextMatrix(Linhas, 12))
            'dlValItem = (dlValUnit - ((dlValUnit * dlSumDscItem) / 100)) * (ilQtde * ilQtdEmb)
           '-> dlValItem = (dlValUnit - ((dlValUnit * dlValDesc) / 100)) * (ilQtde * ilQtdEmb)
           'dlValItem = (dlValUnit * (1 - (dlValDesc / 100))) * (ilQtde * ilQtdEmb)
           dlValItem = (dlValUnit * (1 - (dlValDesc / 100))) * ilQtde 'Alterado a pedido do Manoel 17/04/2017
            
            GrdNotaCliente.TextMatrix(Linhas, 18) = Format(dlValItem, "##,###,##0.00")
'            GrdNotaCliente.TextMatrix(Linhas, 6) = Format(dlSumDscItem, "##,###,##0.00")
            If dlValDesc > dlSumDscItem Then
                GrdNotaCliente.TextMatrix(Linhas, 6) = 0
            Else
                GrdNotaCliente.TextMatrix(Linhas, 6) = Format(dlValDesc, "##,###,##0.00")
            End If
                       
            VlIdealItem = (dlValUntN - ((dlValUntN * dlSumDscItemORIG) / 100)) * ilQtde
            
            GrdNotaCliente.TextMatrix(Linhas, 11) = Format(VlIdealItem, "###,##0.00")
            
'        End If
        
    Next Linhas

End Function

Function CalculaIndice() As Boolean

    '*************************************************************************************
    'Rotina responsável por calcular os índices de custos e margem de lucro que ficam além
    'das margens do formulário visíveis aos representantes.
    '*************************************************************************************
    
    On Error GoTo TrataErro
    
    iDscForaRegiao = 0
    
    Dim dlPerCusFrtCalc As Double
    Dim dlValBruto As Double
    Dim dlValMargem As Double
    Dim dlMargemReal As Double
    Dim dlValCusUnt As Double
    Dim dlCusUntQtd As Double
    Dim dlCusAdiQtd As Double
    Dim dlAlqImpFed As Double
    Dim dlMrgPrd As Double
    Dim dlOutrosCus As Double
    Dim blAlertaDesconto As Boolean
    Dim blVlUnt As Boolean
    Dim dlSoma9 As Double
    Dim ilSumQtd As Double
    Dim dlCusCompo As Double
    Dim Margem As Double
    
    Dim dVlrMin As Double

    CalculaIndice = True
    
    '*****************************************************************************
    '"Status" é o controle que utiliza cores para informar ao representante se o
    'pedido está em condições de ser aceito ou não. Sua cor inicial (padrão) é
    'verde.
    '*****************************************************************************
    
    Status.BackColor = &HFF00&
    Novacor.BackColor = &HFF00&
    GrdIndice.rows = 1
    
    '*************************************************************************************
    'Zera as variáveis de cálculos.
    '*************************************************************************************
    
    dlCusUntQtd = 0
    dlCusAdiQtd = 0
    dlAlqImpFed = 0
    dlSumPes = 0
    dlSumGrd = 0
    dlTotBru = 0
    dlTotPes = 0
    dlMediaIDX = 0
    dlTotLiq = 0
    dlSimples = 0
    dlSumDscTot = 0
    dlTotBru = 0
    dlIlb = 0
    dlTotIdeal = 0
    dlSumIdealGrd = 0
    dlINlb = 0
    blAlertaDesconto = False
    blDescZero = False
    slIrriga = " "
    QtdTubo = 0
    QtdTuboRosc = 0
    QtdAspe = 0
    QtdConx = 0
    ilSumQtd = 0
    dlOutrosCus = 0
    dlCusCompo = 0
    
    '*****************************************************************************************
    'Se o pedido estiver liberado para modificação e o cliente não for optante pelo SIMPLES
    'baiano, o sistema define que os cálculos não devem envolver nada relacionado ao SIMPLES.
    '*****************************************************************************************
            
    If bgBloqPed = False And slFlgSIMBa <> "S" Then
        slPedSimples = "N"
    End If
    
    '*****************************************************************************************
    'Se o sistema tiver negociado comissão com o representante, a comissão oficial será a
    'negociada e haverá uma notificação na janela do pedido. Se não houver negociação, segue o
    'valor padrão definido.
    '*****************************************************************************************
    
    If dlPerComiNeg > 0 Then
        dlPerComiCalc = dlPerComiNeg
        LblResultNegocio.Caption = "Comissão do pedido negociada com o representante em " & Trim(dlPerComiNeg) & "%"
    Else
        dlPerComiCalc = dlPerComiN
        LblResultNegocio.Caption = ""
    End If

    '*****************************************************************************************
    'Se não houver nenhum item de pedido inserido no grid, limpa e trata todos os campos dos
    'parâmetros ocultos de preços de venda e totalização dos pedidos. Abandona a rotina a
    'seguir.
    '*****************************************************************************************

    If GrdNotaCliente.rows < 2 Then
        
        vl1.Texto = 0
        vl2.Texto = 0
        vl3.Texto = 0
        LblTextoideal.Visible = False
        vl3.Visible = False
        LblIdeal.Visible = False
        lblpercideal.Visible = False
        LblSub.Caption = ""
        LblTot.Caption = ""
        LblVlSimples.Caption = ""
        LblDesc.Caption = ""
        LblI.Caption = ""
        LblIdeal.Caption = ""
        ''LblUnit.Caption = "Valor Unitário"
        Status.BackColor = &HC0C0C0
        Novacor.BackColor = &HC0C0C0
        MskMargem.Texto = ""
        
        Exit Function
        
    End If
    
    '*****************************************************************************************
    'A função CalculaDesconto procura os descontos relacionados ao perfil da venda, lista os
    'dados nos parâmetros ocultos do sistema e calcula o preço líqüido de cada item do pedido.
    '*****************************************************************************************
    
    CalculaDesconto
    
    '*****************************************************************************************
    'Explicar aqui para que serve a função RecalculaDescontoGrid.
    '*****************************************************************************************
    
    RecalculaDescontoGrid
    
    sgQuery = "SELECT b.PerCusFin "
    sgQuery = sgQuery + " from CONDICAO a, CUSTO_CONDICAO b "
    sgQuery = sgQuery + "  Where a.CodCnd = " & Trim(ilCodCnd)
    sgQuery = sgQuery + "    and a.codcnd = b.codcnd"
    sgQuery = sgQuery + "    and b.datativ = (select max(datativ) from CUSTO_CONDICAO"
    sgQuery = sgQuery + "                      Where Codcnd = b.codcnd"
    sgQuery = sgQuery + "                        and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"
    
    Call Consulta(sgQuery)
    
    If Rs.EOF Then
    
        MsgBox "Erro na leitura da Condição de Pagamento", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        LimpaGeral
        
        Exit Function
        
    Else
    
        dlPerCusFin = IIf(IsNull(Rs!PerCusFin), 0, Trim(Rs!PerCusFin))
    
    End If

    Rs.Close

    Set Rs = Nothing
    '*****************************************************************************************
    'Começa a calcular e compor o grid com os índices do pedido.
    '*****************************************************************************************
    
    SlTabela = " "
    
    For Linhas = 1 To GrdNotaCliente.rows - 1
    
        '**************************************************************************************
        '
        '**************************************************************************************
    
        If Linhas > 1 And (Trim(GrdNotaCliente.TextMatrix(Linhas, 5)) <> Trim(SlTabela)) Then
            SlTabela = "M"
        End If
        
        '**************************************************************************************
        'Os grupos de produtos cujos códigos estão definidos entre 100 e 199 são tubos. A
        'rotina a seguir soma a quantidade de tubos existente no pedido. A somatória despreza
        'os tubos roscáveis, que serão tratados na seqüência.
        '**************************************************************************************
        
        If Trim(GrdNotaCliente.TextMatrix(Linhas, 8)) >= 100 And Trim(GrdNotaCliente.TextMatrix(Linhas, 8)) < 200 And Trim(GrdNotaCliente.TextMatrix(Linhas, 8)) <> 130 Then
            QtdTubo = QtdTubo + Trim(GrdNotaCliente.TextMatrix(Linhas, 3))
        End If
        
        '**************************************************************************************
        'Calcula a quantidade de tubos roscáveis existente no pedido.
        '**************************************************************************************
        
        If Trim(GrdNotaCliente.TextMatrix(Linhas, 8)) = 130 Then
            QtdTuboRosc = QtdTuboRosc + 1
        End If
        
        '**************************************************************************************
        'Calcula a quantidade de aspersores para irrigação existente no pedido.
        '**************************************************************************************
        
        If Trim(GrdNotaCliente.TextMatrix(Linhas, 8)) = 812 Then
            QtdAspe = QtdAspe + Trim(GrdNotaCliente.TextMatrix(Linhas, 3))
        End If
        
        '**************************************************************************************
        'Os grupos de produtos cujos códigos estão definidos entre 300 e 499 são conexões. A
        'rotina a seguir soma a quantidade de conexões que existente no pedido.
        '**************************************************************************************
        
        If Left(Trim(GrdNotaCliente.TextMatrix(Linhas, 8)), 3) >= 300 And Left(Trim(GrdNotaCliente.TextMatrix(Linhas, 8)), 3) < 500 Then
            QtdConx = QtdConx + 1
        End If
        
        '**************************************************************************************
        'Se a tabela aplicada não é a normal, então define qual foi a utilizada pelo
        'representante.
        '**************************************************************************************
        
        If Trim(SlTabela) <> "M" Then
            SlTabela = GrdNotaCliente.TextMatrix(Linhas, 5)
        End If
        
        '**************************************************************************************
        'Calcula o valor total líqüido da venda de tubos de esgoto 100 milímetros. Item a item.
        '**************************************************************************************
        
        If Trim(GrdNotaCliente.TextMatrix(Linhas, 0)) = 9 Then
            dlSoma9 = dlSoma9 + Format(GrdNotaCliente.TextMatrix(Linhas, 18), "##,###,##0.00")
            dTotb100 = Format(GrdNotaCliente.TextMatrix(Linhas, 18), "##,###,##0.00")
        End If
        
        '*******************************************************
        'Calculo se o pedido contem apenas tubo de 100
        '*******************************************************
        
        If Trim(GrdNotaCliente.TextMatrix(Linhas, 0)) = 9 And GrdNotaCliente.rows = 2 Then
           bSo100 = True
        Else
           bSo100 = False
        End If
              
        '**************************************************************************************
        'Define se o produto em questão não recebeu desconto individual.
        '**************************************************************************************
        
        If Val(GrdNotaCliente.TextMatrix(Linhas, 6)) = 0 Then
            blDescZero = True
        End If
 '------------------------------------------------------------------------
        If bAVista = True Then 'estou trabalhando a vista
            
                If GrdNotaCliente.TextMatrix(Linhas, 6) <= dDscRegiao Then
                    iDscRegiao = iDscRegiao + 1
                Else
                    iDscForaRegiao = iDscForaRegiao + 1
                End If
                
         Else
        
            If GrdNotaCliente.TextMatrix(Linhas, 6) <= dDscRegiao Then
                iDscRegiao = iDscRegiao + 1
            Else
                iDscForaRegiao = iDscForaRegiao + 1
            End If
        
        
        End If
        
'------------------------------------------------------------------------
        
        '**************************************************************************************
        'Calcula o preço total de tabela (bruto) do pedido.
        '**************************************************************************************
        dlTotBru = dlTotBru + Trim(GrdNotaCliente.TextMatrix(Linhas, 9))
        
        '**************************************************************************************
        'Calcula o peso total de todos os produtos do pedido.
        '**************************************************************************************
        dlTotPes = dlTotPes + Trim(GrdNotaCliente.TextMatrix(Linhas, 10))
        
        '**************************************************************************************
        'Calcula o preço total líqüido do pedido (preço líqüido é o preço concedido com o
        'desconto aplicado pelo representante.
        '**************************************************************************************
        dlTotLiq = dlTotLiq + Trim(GrdNotaCliente.TextMatrix(Linhas, 18))
        
        '**************************************************************************************
        'Calcula o preço total ideal para o pedido. Preço ideal é o mínimo a que o pedido pode
        'chegar, levando-se em conta todos os descontos possíveis. Descontos concedidos por
        'chave não fazem parte da lista dos possíveis.
        '**************************************************************************************
        dlTotIdeal = dlTotIdeal + Trim(GrdNotaCliente.TextMatrix(Linhas, 11))
        
        blAchou = False
        
        '**************************************************************************************
        'Procura na base de dados o custo do grupo de produto mais recente. Isso é o mesmo que
        'procurar pelo preço do composto do produto em questão, com o objetivo de envolver tal
        'preço nos cálculos dos índices.
        '**************************************************************************************
        
        '**************************************************************************************
        'Em 05/07/2010 os tubos de 100mm passaram a contar com um grupo de produto exclusivo.
        'Assim, todos os pedidos emitidos até 05/07/2010 são consultados informando o código
        '110 como sendo do grupo de produto dos tubos de 100mm. Do dia 5 em diante informa-se o
        'código 113. Como não havia grupo 113 em datas anteriores à informada, a consulta na
        'sua forma original estava gerando um erro.
        '**************************************************************************************
        
        If GrdNotaCliente.TextMatrix(Linhas, 8) = 113 Then
        
            If Datped < #7/5/2010# Then
            
                sgQuery = "select a.*, b.* from GRUPO_PRODUTO a, CUSTO_GRUPO_PRODUTO b"
                sgQuery = sgQuery + "    Where a.IdeGrp = " & Trim(GrdNotaCliente.TextMatrix(Linhas, 8))
                sgQuery = sgQuery + "      and a.IdeGrp = b.IdeGrp"
                sgQuery = sgQuery + "      and b.datativ = (select max(datativ) from CUSTO_GRUPO_PRODUTO"
                sgQuery = sgQuery + "                        Where IdeGrp = " & Trim(GrdNotaCliente.TextMatrix(Linhas, 8))
                sgQuery = sgQuery + "                          and datativ <= convert(datetime,'" & "05/07/2010" & "',103))"
                
            Else
        
                sgQuery = "select a.*, b.* from GRUPO_PRODUTO a, CUSTO_GRUPO_PRODUTO b"
                sgQuery = sgQuery + "    Where a.IdeGrp = " & Trim(GrdNotaCliente.TextMatrix(Linhas, 8))
                sgQuery = sgQuery + "      and a.IdeGrp = b.IdeGrp"
                sgQuery = sgQuery + "      and b.datativ = (select max(datativ) from CUSTO_GRUPO_PRODUTO"
                sgQuery = sgQuery + "                        Where IdeGrp = " & Trim(GrdNotaCliente.TextMatrix(Linhas, 8))
                sgQuery = sgQuery + "                          and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"
                
            End If
            
        Else
        
            sgQuery = "select a.*, b.* from GRUPO_PRODUTO a, CUSTO_GRUPO_PRODUTO b"
            sgQuery = sgQuery + "    Where a.IdeGrp = " & Trim(GrdNotaCliente.TextMatrix(Linhas, 8))
            sgQuery = sgQuery + "      and a.IdeGrp = b.IdeGrp"
            sgQuery = sgQuery + "      and b.datativ = (select max(datativ) from CUSTO_GRUPO_PRODUTO"
            sgQuery = sgQuery + "                        Where IdeGrp = " & Trim(GrdNotaCliente.TextMatrix(Linhas, 8))
            sgQuery = sgQuery + "                          and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"
            
        End If
            
        Call Consulta(sgQuery)
        
        '*************************************************************************************
        'Retorna um erro se não encontrar preço de composto para o produto em questão. Encerra
        'a rotina depois disso.
        '*************************************************************************************
        
        If Rs.EOF = True Then
                
            sgQuery = "Grupo não encontrado [" & Trim(GrdNotaCliente.TextMatrix(Linhas, 8)) & "], informe ao administrador do sistema, CalculaIndice"
                
            MsgBox sgQuery, vbExclamation + vbOKOnly, "Atenção!"
                
            CalculaIndice = False
                
            Exit Function
            
        End If
        
        ilI = GrdIndice.rows
        
        '*************************************************************************************
        'Define custo de frete. Se for FOB, custo zero; se não, informa-se o preço cobrado.
        '*************************************************************************************
        
        If Opt_FOB = True Then
            dlPerCusFrtCalc = 0
        Else
            dlPerCusFrtCalc = dlPerCusFrt
        End If
        
        '*************************************************************************************
        '
        '*************************************************************************************
        
        '*************************************************************************************
        'Segue abaixo uma lista com a posição de cada informação no grid dos índices:
        '0: Código do produto
        '1: Descrição do produto
        '2: Índice de custo fixo para o produto
        '3: Percentual de custo financeiro
        '4: Percentual de comissão do representante
        '5: Custo de frete
        '6: Provisão de Devedores Duvidosos (PDD)
        '7: Margem de lucro
        '8: Alíquota de ICMS
        '9: Custo unitário do produto (quando adquirido de terceiros, para revenda)
        '10: Índice de custo do composto
        '11: Preço mínimo aceitável para o pedido
        '12: Valor total líqüido do produto
        '13: Preço ideal de venda
        '14: Peso bruto do produto
        '15: Margem real de lucro (percentual)
        '16: Valor real de lucro (em moeda)
        '17: Soma dos fatores que compõe o preço (sem margem)
        '*************************************************************************************
        
        If GrdIndice.rows = GrdNotaCliente.rows Then
           Exit Function
        End If

        GrdIndice.rows = GrdIndice.rows + 1
        GrdIndice.TextMatrix(ilI, 0) = Format(Trim(GrdNotaCliente.TextMatrix(Linhas, 0)), "0000")
        GrdIndice.TextMatrix(ilI, 1) = Trim(GrdNotaCliente.TextMatrix(Linhas, 1))
        GrdIndice.TextMatrix(ilI, 2) = Format(Rs!IdxFix, "##0.00")
        GrdIndice.TextMatrix(ilI, 3) = Format(dlPerCusFin, "##0.00")
        GrdIndice.TextMatrix(ilI, 4) = Format(dlPerComiCalc, "##0.00")
        GrdIndice.TextMatrix(ilI, 5) = Format(dlPerCusFrtCalc, "##0.00")
        GrdIndice.TextMatrix(ilI, 6) = Format(dlIdxPDD, "##0.00")
        GrdIndice.TextMatrix(ilI, 7) = Format(Trim(GrdNotaCliente.TextMatrix(Linhas, 13)), "##0.00")
        GrdIndice.TextMatrix(ilI, 8) = Format(PerCusIcm, "##0.00")
        
        '*************************************************************************************
        'Armazena em variáveis, respectivamente: custo unitário de aquisição do produto, custo
        'unitário adicional de aquisição do produto e alíquota de imposto federal.
        '*************************************************************************************
        
        dlCusUntQtd = Trim(GrdNotaCliente.TextMatrix(Linhas, 15))
        dlCusAdiQtd = Trim(GrdNotaCliente.TextMatrix(Linhas, 16))
        dlAlqImpFed = Trim(GrdNotaCliente.TextMatrix(Linhas, 17))
        
        '*************************************************************************************
        'Se o produto que está sendo vendido foi adquirido de terceiros para revenda, seu
        'preço de custo é informado no grid dos índices. Se tal produto é produzido pela
        'Unocann, informa-se o custo do seu composto.
        '*************************************************************************************
        
        If dlCusUntQtd > 0 Then
            GrdIndice.TextMatrix(ilI, 9) = Format(dlCusUntQtd, "##,###,##0.00")
            blVlUnt = True
        Else
            GrdIndice.TextMatrix(ilI, 9) = Format(Rs!ValCusUnt, "##0.00")
            blVlUnt = False
        End If
        
        '*************************************************************************************
        'Armazena a margem de lucro obtida para o produto.
        '*************************************************************************************
        
        dlMrgPrd = Format(Trim(GrdNotaCliente.TextMatrix(Linhas, 13)), "##0.00")
        
        '*************************************************************************************
        'Soma todos os fatores envolvidos na composição de preço do produto, ou seja, custos e
        'margem de lucro. Pela ordem, são os seguintes: índice de custo fixo médio (pessoal,
        'máquinário, etc.), custo financeiro, comissão de representante, frete, margem de
        'lucro, ICMS, provisão de devedores duvidosos e impostos federais.
        '*************************************************************************************
        
        dlSumGrd = Rs!IdxFix + dlPerCusFin + dlPerComiCalc + dlPerCusFrtCalc + dlMrgPrd + PerCusIcm + dlIdxPDD + dlAlqImpFed
        
        '*************************************************************************************
        'Soma todos os custos envolvidos na composição de preço do produto. É o mesmo cálculo
        'da linha anterior, porém sem margem de lucro.
        '*************************************************************************************
        
        dlSumSubGrd = Rs!IdxFix + dlPerCusFin + dlPerComiCalc + dlPerCusFrtCalc + PerCusIcm + dlIdxPDD + dlAlqImpFed
        
        '*************************************************************************************
        'Armazena em variável a quantidade vendida do produto atual.
        '*************************************************************************************
                
        ilSumQtd = Trim(GrdNotaCliente.TextMatrix(Linhas, 3))
        
        '*************************************************************************************
        'Obtém o percentual que sobra do preço total de venda após o desconto dos custos e da
        'margem de lucro. Este percentual vai servir para calcular o preço mínimo da venda
        'atual.
        '*************************************************************************************
        
        dlIDX = (100 - dlSumGrd) / 100
        
        '*************************************************************************************
        'Se o produto foi comprado para revenda, o fator de cálculo será igual ao custo de
        'aquisição dividido pelo valor antigo do fator; o custo do composto será igual ao
        'custo de aquisição multiplicado pela quantidade de peças vendidas. Se o produto foi
        'produzido pela Unocann, o fator de cálculo será igual ao custo do composto dividido
        'pelo antigo valor do fator. O novo valor do fator será armazenado na coluna 10 do
        'grid de índices.
        '*************************************************************************************
        
        '*************************************************************************************
        'Calcula o percentual de participação do custo do composto sobre o preço de venda,
        'descontados os outros custos e a margem de lucro. O resultado será um preço de
        'composto mais alto que o cadastrado, já que tal preço terá sofrido os impactos dos
        'outros custos.
        '*************************************************************************************
        
        If blVlUnt = True Then
            dlIDX = Format(dlCusUntQtd / dlIDX, "###,###,##0.00")
            dlCusCompo = dlCusUntQtd * ilSumQtd
        Else
            dlIDX = Format(Rs!ValCusUnt / dlIDX, "##0.00")
        End If
        
        GrdIndice.TextMatrix(ilI, 10) = Format(dlIDX, "##0.00")
        
        '*************************************************************************************
        'Armazena em variável o peso bruto do produto.
        '*************************************************************************************
        
        dlSumPes = Trim(GrdNotaCliente.TextMatrix(Linhas, 10))
        
        '*************************************************************************************
        'A seguir, calcula-se o preço mínimo para o pedido. Se o produto foi adquirido para
        'revenda, multiplica-se o fator de cálculo pela quantidade vendida e soma-se esse
        'resultado ao resultado da multiplicação da quantidade vendida pelo custo unitário
        'adicional de aquisição do produto. Caso o produto tenha sido produzido pela Unocann,
        'multiplica-se o fator de cálculo pelo peso total do produto vendido e soma-se esse
        'resultado ao resultado da multiplicação da quantidade vendida pelo custo unitário
        'adicional de aquisição do produto.
        '*************************************************************************************
        
        '*************************************************************************************
        'Multiplica o custo do composto para o produto (após os outros custos e a margem de
        'lucro) pelo peso total do produto. O resultado é o preço mínimo aceitável para o
        'produto em questão.
        '*************************************************************************************
        
        If blVlUnt = True Then
            dVlrMin = 0
            GrdIndice.TextMatrix(ilI, 11) = Format((dlIDX * ilSumQtd) + (ilSumQtd * dlCusAdiQtd), "##,###,##0.00")
            dVlrMin = GrdIndice.TextMatrix(ilI, 11) '= Format((dlIDX * ilSumQtd) + (ilSumQtd * dlCusAdiQtd), "##,###,##0.00")
        Else
            GrdIndice.TextMatrix(ilI, 11) = Format((dlIDX * dlSumPes) + (ilSumQtd * dlCusAdiQtd), "##,###,##0.00")
        End If
        
        '*************************************************************************************
        'Armazena no grid dos índices, pela ordem: valor total líqüido do produto, preço ideal
        'da venda e peso bruto do produto.
        '*************************************************************************************
        
        GrdIndice.TextMatrix(ilI, 12) = Format(GrdNotaCliente.TextMatrix(Linhas, 18), "##,##0.00")
        GrdIndice.TextMatrix(ilI, 13) = Format(GrdNotaCliente.TextMatrix(Linhas, 11), "##,##0.00")
        GrdIndice.TextMatrix(ilI, 14) = Format(GrdNotaCliente.TextMatrix(Linhas, 10), "##0.000")
        
        '*************************************************************************************
        'Calcula a margem de lucro.
        '*************************************************************************************
        
        '*************************************************************************************
        'Armazena em variável o preço total líqüido do produto.
        '*************************************************************************************
        
        dlSumGrd = Format(GrdNotaCliente.TextMatrix(Linhas, 18), "##,##0.00")
        
        '*************************************************************************************
        'O fator de cálculo passa a ser a soma de todos os elementos de composição de preço do
        'produto, sem margem de lucro.
        '*************************************************************************************
        
        dlIDX = dlSumSubGrd
        
        '*************************************************************************************
        'Se o cliente for SIMPLES, tem direito a 10.752% de desconto. Daí calcula-se o que
        'representa este percentual sobre o custo total, incluindo margem de lucro. O
        'resultado é subtraído do custo total, de modo a melhorar os índices e facilitar a
        'concessão de maiores descontos ao cliente.
        '*************************************************************************************
        
        If slPedSimples = "S" Then
            dlSimples = (dlSumGrd * 10.752) / 100
            dlSumGrd = dlSumGrd - dlSimples
        End If
        
        '*************************************************************************************
        'Calcula custo de composto, outros custos de produção, valor da margem (em moeda) e
        'margem real de lucro (em percentual). As variáveis de cálculo mudam levando-se em
        'consideração se o produto foi adquirido para revenda ou produzido pela Unocann.
        '*************************************************************************************
        
        '*************************************************************************************
        'Avalia se o produto foi adquirido para revenda.
        '*************************************************************************************
        
        If blVlUnt = True Then
        
            '*********************************************************************************
            'Calcula outros custos multiplicando o preço líqüido do produto pelo fator de
            'cálculo e dividinho o resultado por 100. Essa conta faz sentido, já que o fator
            'de cálculo foi obtido a partir da composição de preço do produto sem a margem de
            'lucro.
            '*********************************************************************************
            dlOutrosCus = Format((dlSumGrd * dlIDX) / 100, "####0.00")
            
            '*********************************************************************************
            'Calcula o custo do composto multiplicando o preço cobrado pelo fornecedor externo
            'pela quantidade vendida; multiplica-se também a quantidade vendida pelo custo
            'adicional de aquisição do produto. O custo do composto é a soma dos dois
            'resultados.
            '*********************************************************************************
            dlCusCompo = (dlCusUntQtd * ilSumQtd) + (ilSumQtd * dlCusAdiQtd)
            
            '*********************************************************************************
            'Calcula a margem de lucro em moeda subtraindo o preço líqüido do produto pelo
            'resultado da soma dos custos (composto e outros).
            '*********************************************************************************
            dlValMargem = dlSumGrd - (dlOutrosCus + dlCusCompo)
            dlValMargem = Format(dlValMargem, "0.00")
            
            '*********************************************************************************
            'Calcula a margem real de lucro (em percentual) dividindo o valor do lucro em
            'moeda pelo preço total líqüido e dividindo seu resultado por 100.
            '*********************************************************************************
            dlMargemReal = (dlValMargem / dlSumGrd) * 100
            
        Else
            
            '*********************************************************************************
            'Calcula outros custos multiplicando o preço líqüido do produto pelo fator de
            'cálculo e dividindo o resultado por 100. Essa conta faz sentido, já que o fator
            'de cálculo foi obtido a partir da composição de preço do produto sem a margem de
            'lucro.
            '*********************************************************************************
            dlOutrosCus = Format((dlSumGrd * dlIDX) / 100, "####0.00")
            
            '*********************************************************************************
            'Calcula o custo do composto utilizado multiplicando seu preço (kg) pelo peso dos
            'itens vendidos; multiplica-se também a quantidade vendida pelo custo adicional de
            'aquisição do produto. O custo do composto é a soma dos dois resultados.
            '*********************************************************************************
            dlCusCompo = Rs!ValCusUnt * dlSumPes + (ilSumQtd * dlCusAdiQtd)
            
            '*********************************************************************************
            'Calcula a margem de lucro em moeda subtraindo o preço líqüido do produto pelo
            'resultado da soma dos custos (composto e outros).
            '*********************************************************************************
            dlValMargem = dlSumGrd - (dlOutrosCus + dlCusCompo)
            
            '*********************************************************************************
            'Calcula a margem real de lucro (em percentual) dividindo o valor do lucro em
            'moeda pelo preço total líqüido e dividindo seu resultado por 100.
            '*********************************************************************************
            dlMargemReal = (dlValMargem / dlSumGrd) * 100
            
        End If
        
        '*************************************************************************************
        'Armazena no grid dos índices, pela ordem: margem real de lucro (em percentual), valor
        'da margem de lucro (em moeda) e a soma dos fatores que compõe o preço (sem margem).
        '*************************************************************************************
        GrdIndice.TextMatrix(ilI, 15) = Format(dlMargemReal, "##0.00")
        GrdIndice.TextMatrix(ilI, 16) = Format(dlValMargem, "##,###,##0.00")
        GrdIndice.TextMatrix(ilI, 17) = Format(dlIDX, "##,###,##0.00")
        
        '*************************************************************************************
        'Se o produto tiver sido adquirido para revenda, a sua célula que armazena o custo do
        'composto terá fundo amarelo e fonte vermelha.
        '*************************************************************************************
        
        If blVlUnt = True Then
            GrdIndice.row = ilI
            GrdIndice.col = 10
            GrdIndice.CellForeColor = &HFFFF&
            GrdIndice.CellBackColor = &HFF&
        End If
        
        Rs.Close
            
        Set Rs = Nothing
        
    Next Linhas

    '*****************************************************************************************
    'Começa a calcular a margem do pedido, oculta do lado direito do formulário.
    '*****************************************************************************************
            
    dlaux = 0
    dlValMargem = 0
    
    For Linhas1 = 1 To GrdIndice.rows - 1
    
        '*************************************************************************************
        'Calcula o valor total mínimo permitido para o pedido.
        '*************************************************************************************
        dlaux = dlaux + GrdIndice.TextMatrix(Linhas1, 11)
        
        '*************************************************************************************
        'Calcula o lucro (em moeda) obtido com o pedido.
        '*************************************************************************************
        dlValMargem = dlValMargem + GrdIndice.TextMatrix(Linhas1, 16)
    
    Next Linhas1
    
    '*****************************************************************************************
    'Se o cliente for optante pelo SIMPLES, concede desconto de 10.752 no valor líqüido total
    'do pedido. Tal desconto refere-se a incentivo concedido pelo governo da Bahia.
    '*****************************************************************************************
    
    If slPedSimples = "S" Then
        dlSimples = (dlTotLiq * 10.752) / 100
        dlTotLiq = dlTotLiq - dlSimples
    End If
    
    '*****************************************************************************************
    'Retira o cálculo do SIMPLES do total ideal e calcula a margem geral do pedido.
    '*****************************************************************************************
    
    dlTotIdeal = dlTotIdeal - dlSimples
    dlMargemGeral = (dlValMargem / dlTotLiq) * 100
    
    MskMargem.Texto = dlMargemGeral
    
    
    '*****************************************************************************************
    'Se a margem geral for maior que zero, seu valor será exibido com fonte verde. Se a margem
    'for negativa, a fonte será vermelha.
    '*****************************************************************************************
    
    If dlMargemGeral > 0 Then
        MskMargem.ForeColor = &HC000&
    Else
        MskMargem.ForeColor = &HFF&
    End If
'
    '*****************************************************************************************
    'Exibe nos campos ocultos o preço mínimo aceitável para o pedido e o lucro obtido.
    '*****************************************************************************************
    
    vl1.Texto = Format(dlaux, "##,##0.00")
    vl2.Texto = Format(dlTotLiq, "##,##0.00")

    '*****************************************************************************************
    'Exibe as totalizações do pedido abaixo do grid dos produtos.
    '*****************************************************************************************

    LblSub.Caption = Format(dlTotBru, "##,###,##0.00")
    LblVlSimples.Caption = Format(dlSimples, "##,###,##0.00")
    LblTot.Caption = Format(dlTotLiq, "##,###,##0.00")
    LblDesc.Caption = Format(dlTotBru - (dlTotLiq + dlSimples), "##,###,##0.00")
    
    '*****************************************************************************************
    'Calcula a participação de tubos 100 milímetros no pedido. Para isso, divide o total
    'líqüido dos tubos de 100 pelo total do pedido e multiplica o resultado por 100.
    '*****************************************************************************************
    
    T100Ped.Texto = (dlSoma9 / dlTotLiq) * 100
    
    '*****************************************************************************************
    'Calcula o índice do pedido. Tal índice resulta da divisão do total líqüido do pedido pelo
    'valor mínimo aceitável para a venda.
    '*****************************************************************************************
            
    If dlaux > 0 Then
        dlIlb = ((dlTotLiq / dlaux) * 100) - 100
    End If

    LblI.Caption = Format(dlIlb, "##0.00")
    
    '*****************************************************************************************
    'Calcula índice ideal. Índice ideal é aquele que leva em consideração o valor do pedido
    'com os descontos disponíveis e o valor mínimo aceitável.
    '*****************************************************************************************
    
    dlIdeallb = 0
    
    '*****************************************************************************************
    'Oculta todos os campos que exibem informações sobre o índice ideal.
    '*****************************************************************************************
    
    LblTextoideal.Visible = False
    vl3.Visible = False
    LblIdeal.Visible = False
    lblpercideal.Visible = False
    
    If dlaux > 0 Then
        
        '*************************************************************************************
        'Calcula o índice e exibe seus valores.
        '*************************************************************************************
        
        dlIdeallb = ((dlTotIdeal / dlaux) * 100) - 100
        
        LblTextoideal.Visible = True
        vl3.Visible = True
        LblIdeal.Visible = True
        lblpercideal.Visible = True
        LblIdeal.Caption = Format(dlIdeallb, "##,##0.00")
        vl3.Texto = Format(dlTotIdeal, "##,##0.00")
        
    End If
    
    '*****************************************************************************************
    'Se o índice ideal for negativo, a fonte do campo que exibe esse valor será vermelha. Caso
    'contrário, a fonte será verde.
    '*****************************************************************************************
    
    If dlIdeallb < 0 Then
        LblIdeal.ForeColor = &HFF&
    Else
        LblIdeal.ForeColor = &HC00000
    End If
    
    '*****************************************************************************************
    'Se o total ideal for menor que o mínimo aceitável, a fonte do campo que exibe esse valor
    'será vermelha. Caso contrário, será verde.
    '*****************************************************************************************
    
    If dlTotIdeal < dlaux Then
        vl3.ForeColor = &HFF&
    Else
        vl3.ForeColor = &HC00000
    End If
    
    '*****************************************************************************************
    'Se o índice do pedido for maior ou igual a zero, avalia se esse índice índice é maior que
    'o ideal, se o total ideal é maior que o mínimo aceitável e se o mínimo aceitável é maior
    'que zero. Se todas as condições forem atendidas, a fonte do campo que exibe o índice do
    'pedido será azul. Caso alguma dessas condições não seja atendida, tal fonte será verde.
    'Se o índice do pedido for negativo, a fonte será vermelha.
    '*****************************************************************************************

    If dlIlb >= 0 Then
   
        If dlIlb > dlIdeallb And (dlTotIdeal > dlaux And dlaux > 0) Then
            LblI.ForeColor = &HC00000
        Else
            LblI.ForeColor = &H8000&
        End If
    
    Else
        
        LblI.ForeColor = &HFF&
    
    End If

    '*****************************************************************************************
    'Calcula a diferença entre o índice do pedido e o índice ideal. Exibe os campos que
    'informam esses valores.
    '*****************************************************************************************

    LblPerIN.Visible = True
    LblIN.Visible = True

    dlINlb = dlIlb - dlIdeallb

    LblIN.Caption = Format(dlINlb, "##0.00")
    
    '*****************************************************************************************
    'Se a diferença entre os índices for igual a zero ou positiva, a fonte do campo que exibe
    'tal diferença será azul. Caso contrário, será vermelha.
    '*****************************************************************************************
    
    If dlINlb >= 0 Then
        LblIN.ForeColor = &HC00000
    Else
        LblIN.ForeColor = &HFF&
    End If
    
    '*****************************************************************************************
    'Determina a cor do controle que informa a qualidade do pedido ao representante. Se a
    'margem geral for menor que 10, o controle de status ficará vermelho. Se a margem for
    'maior que 10 e menor que 15 o status ficará verde. Se a margem for maior que 15, o status
    'ficará azul.
    '*****************************************************************************************
    
    'verde = &HFF00&
    'azul = &H00FF0000&
    'vermelho = &HFF& &H000000FF&
    'amarelo = &H0080FFFF&
    
'    If Format(dlMargemGeral, "##0,00") < 10 Then
'        Status.BackColor = &HFF&
'    ElseIf Format(dlMargemGeral, "##0,00") >= 10 And Format(dlMargemGeral, "##0,00") < 15 Then
'        Status.BackColor = &HFF00&
'    ElseIf Format(dlMargemGeral, "##0,00") >= 15 Then
'        Status.BackColor = &HFF0000 'azul
'    End If
    
    If dlMargemGeral < 6 Then
        Novacor.BackColor = &HFF&       'vermelha
    ElseIf dlMargemGeral >= 6 And dlMargemGeral < 8.6 Then
        Novacor.BackColor = &H80FFFF          ' amarela &HFF&       'vermelha '
    ElseIf dlMargemGeral >= 8.6 And dlMargemGeral <= 12 Then
        Novacor.BackColor = &HFF00&    ' verde
    ElseIf Format(dlMargemGeral, "##0,00") > 12 Then
        Novacor.BackColor = &HFF0000 'azul
    End If
    
'    Select Case dlMargemGeral
'        Case Is < 6
'            Novacor.BackColor = &HFF&       'vermelha
'        Case 6.01 To 8.99
'              Novacor.BackColor = &H80FFFF          ' amarela
'        Case Is > 9  'To 12
'            Novacor.BackColor = &HFF00&    ' verde
' '       Case Is > 12
' '           Novacor.BackColor = &HFF0000 'azul
'     End Select
    
      
    
    
    '*****************************************************************************************
    'Se a diferença entre os índices e o índice geral forem maiores que zero e o índice ideal
    'for menor que zero, a rotina é abandonada.
    '*****************************************************************************************
    
    If dlINlb > 0 And dlIlb > 0 And dlIdeallb < 0 Then
        Exit Function
    End If
    
    '*****************************************************************************************
    'Escolhe um índice para sugestão.
    '*****************************************************************************************
                
    If dlIlb < 0 And dlINlb < 0 Then
    
        '*************************************************************************************
        'Define o percentual a ser reduzido da comissão do representante, para que o pedido
        'possa ser aceito. Em seguida, se o índice geral for menor que a diferença entre eles,
        'cria-se um novo índice com o valor absoluto do índice geral. Caso a diferença seja
        'maior ou igual ao geral, o novo índice será igual ao valor absoluto da diferença.
        '*************************************************************************************
        
        PerSugIni = PerSug1Ini
        
        If dlIlb < dlINlb Then
            dlIABS = Abs(dlIlb)
        Else
            dlIABS = Abs(dlINlb)
        End If
    
    Else
                    
        '*************************************************************************************
        'Se o índice do pedido for negativo, o percentual de redução da comissão a ser
        'sugerido é o padrão; o novo índice será igual ao valor absoluto do índice geral. Caso
        'contrário, o percentual de redução será igual à segunda opção encontrada e o novo
        'índice será igual ao valor absoluto da diferença entre os índices.
        '*************************************************************************************
        
        If dlIlb < 0 Then
            PerSugIni = PerSug1Ini
            dlIABS = Abs(dlIlb)
        Else
            PerSugIni = PerSug2Ini
            dlIABS = Abs(dlINlb)
        End If
        
    End If
    
    '*****************************************************************************************
    'Se o grid com os itens do produto estiver vazio, esvazia também o grid dos índices.
    '*****************************************************************************************
    
    If GrdNotaCliente.rows = 1 Then
        GrdIndice.rows = 1
    End If

    Me.Refresh

    Exit Function

TrataErro:

    Rotina_Erro "CalculaIndice"

End Function

Private Sub CalculaMargem(grid As MSFlexGrid)
    
    '*************************************************************************************
    'Rotina responsável por calcular os índices de custos e margem de lucro que ficam além
    'das margens do formulário visíveis aos representantes.
    '*************************************************************************************
    
    On Error GoTo TrataErro
    
    Dim dlPerCusFrtCalc As Double
    Dim dlValMargem As Double
    Dim dlMargemReal As Double
    Dim dlCusUntQtd As Double
    Dim dlCusAdiQtd As Double
    Dim dlAlqImpFed As Double
    Dim dlMrgPrd As Double
    Dim dlOutrosCus As Double
    Dim blVlUnt As Boolean
    Dim ilSumQtd As Double
    Dim dlCusCompo As Double
    
    '*************************************************************************************
    'Zera as variáveis de cálculos.
    '*************************************************************************************
    
    dlCusUntQtd = 0
    dlCusAdiQtd = 0
    dlAlqImpFed = 0
    dlSumPes = 0
    dlSumGrd = 0
    dlSimples = 0
    ilSumQtd = 0
    dlOutrosCus = 0
    dlCusCompo = 0
    
    '*****************************************************************************************
    'Se o pedido estiver liberado para modificação e o cliente não for optante pelo SIMPLES
    'baiano, o sistema define que os cálculos não devem envolver nada relacionado ao SIMPLES.
    '*****************************************************************************************
            
    slPedSimples = "N"
    
    '*****************************************************************************************
    'Se o sistema tiver negociado comissão com o representante, a comissão oficial será a
    'negociada e haverá uma notificação na janela do pedido. Se não houver negociação, segue o
    'valor padrão definido.
    '*****************************************************************************************
    
    dlPerComiCalc = dlPerComiN
    LblResultNegocio.Caption = ""

    CalculaDesconto
    
    '*****************************************************************************************
    'Explicar aqui para que serve a função RecalculaDescontoGrid.
    '*****************************************************************************************
    
    RecalculaDescontoGrid
    
    '*****************************************************************************************
    'Começa a calcular e compor o grid com os índices do pedido.
    '*****************************************************************************************
    
    SlTabela = "M"
    
    For Linhas = 1 To grid.rows - 1
        
        If 34.2 <= dDscRegiao Then
            iDscRegiao = iDscRegiao + 1
        Else
            iDscForaRegiao = iDscForaRegiao + 1
        End If
        
        '**************************************************************************************
        'Calcula o preço total de tabela (bruto) do pedido.
        '**************************************************************************************
        dlTotBru = dlTotBru + Trim(grid.TextMatrix(Linhas, 5))
        
        '**************************************************************************************
        'Calcula o peso total de todos os produtos do pedido.
        '**************************************************************************************
        dlTotPes = dlTotPes + Trim(grid.TextMatrix(Linhas, 13))
        
        '**************************************************************************************
        'Calcula o preço total líqüido do pedido (preço líqüido é o preço concedido com o
        'desconto aplicado pelo representante.
        '**************************************************************************************
        dlTotLiq = dlTotLiq + Trim((grid.TextMatrix(Linhas, 5) - (grid.TextMatrix(Linhas, 5) * 0.342) * 1))
        
        '**************************************************************************************
        'Calcula o preço total ideal para o pedido. Preço ideal é o mínimo a que o pedido pode
        'chegar, levando-se em conta todos os descontos possíveis. Descontos concedidos por
        'chave não fazem parte da lista dos possíveis.
        '**************************************************************************************
        dlTotIdeal = dlTotIdeal + Trim(grid.TextMatrix(Linhas, 15))
        
        blAchou = False
        '**************************************************************************************
        'Procura na base de dados o custo do grupo de produto mais recente. Isso é o mesmo que
        'procurar pelo preço do composto do produto em questão, com o objetivo de envolver tal
        'preço nos cálculos dos índices.
        '**************************************************************************************
        
        '**************************************************************************************
        'Em 05/07/2010 os tubos de 100mm passaram a contar com um grupo de produto exclusivo.
        'Assim, todos os pedidos emitidos até 05/07/2010 são consultados informando o código
        '110 como sendo do grupo de produto dos tubos de 100mm. Do dia 5 em diante informa-se o
        'código 113. Como não havia grupo 113 em datas anteriores à informada, a consulta na
        'sua forma original estava gerando um erro.
        '**************************************************************************************
        
        If grid.TextMatrix(Linhas, 12) = 113 Then
        
            If Datped < #7/5/2010# Then
            
                sgQuery = "select a.*, b.* from GRUPO_PRODUTO a, CUSTO_GRUPO_PRODUTO b"
                sgQuery = sgQuery + "    Where a.IdeGrp = " & Trim(grid.TextMatrix(Linhas, 12))
                sgQuery = sgQuery + "      and a.IdeGrp = b.IdeGrp"
                sgQuery = sgQuery + "      and b.datativ = (select max(datativ) from CUSTO_GRUPO_PRODUTO"
                sgQuery = sgQuery + "                        Where IdeGrp = " & Trim(grid.TextMatrix(Linhas, 12))
                sgQuery = sgQuery + "                          and datativ <= convert(datetime,'" & "05/07/2010" & "',103))"
                
            Else
        
                sgQuery = "select a.*, b.* from GRUPO_PRODUTO a, CUSTO_GRUPO_PRODUTO b"
                sgQuery = sgQuery + "    Where a.IdeGrp = " & Trim(grid.TextMatrix(Linhas, 12))
                sgQuery = sgQuery + "      and a.IdeGrp = b.IdeGrp"
                sgQuery = sgQuery + "      and b.datativ = (select max(datativ) from CUSTO_GRUPO_PRODUTO"
                sgQuery = sgQuery + "                        Where IdeGrp = " & Trim(grid.TextMatrix(Linhas, 12))
                sgQuery = sgQuery + "                          and datativ <= convert(datetime,GETDATE(),103))"
                
            End If
            
        Else
        
            sgQuery = "select a.*, b.* from GRUPO_PRODUTO a, CUSTO_GRUPO_PRODUTO b"
            sgQuery = sgQuery + "    Where a.IdeGrp = " & Trim(grid.TextMatrix(Linhas, 12))
            sgQuery = sgQuery + "      and a.IdeGrp = b.IdeGrp"
            sgQuery = sgQuery + "      and b.datativ = (select max(datativ) from CUSTO_GRUPO_PRODUTO"
            sgQuery = sgQuery + "                        Where IdeGrp = " & Trim(grid.TextMatrix(Linhas, 12))
            sgQuery = sgQuery + "                          and datativ <= convert(datetime,GETDATE(),103))"
            
        End If
            
        Call Consulta(sgQuery)
        
        '*************************************************************************************
        'Retorna um erro se não encontrar preço de composto para o produto em questão. Encerra
        'a rotina depois disso.
        '*************************************************************************************
        
        If Rs.EOF = True Then
                
            sgQuery = "Grupo não encontrado [" & Trim(grid.TextMatrix(Linhas, 12)) & "], informe ao administrador do sistema, CalculaMargem"
                
            MsgBox sgQuery, vbExclamation + vbOKOnly, "Atenção!"
                
            Exit Sub
            
        End If
        
        ilI = GrdIndice.rows
        
        '*************************************************************************************
        'Define custo de frete. Se for FOB, custo zero; se não, informa-se o preço cobrado.
        '*************************************************************************************
        
        dlPerCusFrtCalc = 0
        
        '*************************************************************************************
        'Armazena em variáveis, respectivamente: custo unitário de aquisição do produto, custo
        'unitário adicional de aquisição do produto e alíquota de imposto federal.
        '*************************************************************************************
        
        dlCusUntQtd = Trim(grid.TextMatrix(Linhas, 15))
        dlCusAdiQtd = Trim(grid.TextMatrix(Linhas, 16))
        dlAlqImpFed = Trim(grid.TextMatrix(Linhas, 17))
        
        '*************************************************************************************
        'Armazena a margem de lucro obtida para o produto.
        '*************************************************************************************
        
        dlMrgPrd = Format(Trim(grid.TextMatrix(Linhas, 14)), "##0.00")
        
        '*************************************************************************************
        'Soma todos os fatores envolvidos na composição de preço do produto, ou seja, custos e
        'margem de lucro. Pela ordem, são os seguintes: índice de custo fixo médio (pessoal,
        'máquinário, etc.), custo financeiro, comissão de representante, frete, margem de
        'lucro, ICMS, provisão de devedores duvidosos e impostos federais.
        '*************************************************************************************
        
        dlSumGrd = Rs!IdxFix + dlPerCusFin + dlPerComiCalc + dlPerCusFrtCalc + dlMrgPrd + PerCusIcm + dlIdxPDD + dlAlqImpFed
        
        '*************************************************************************************
        'Soma todos os custos envolvidos na composição de preço do produto. É o mesmo cálculo
        'da linha anterior, porém sem margem de lucro.
        '*************************************************************************************
        
        dlSumSubGrd = Rs!IdxFix + dlPerCusFin + dlPerComiCalc + dlPerCusFrtCalc + PerCusIcm + dlIdxPDD + dlAlqImpFed
        
        '*************************************************************************************
        'Armazena em variável a quantidade vendida do produto atual.
        '*************************************************************************************
                
        ilSumQtd = 1
        
        '*************************************************************************************
        'Se o produto foi comprado para revenda, o fator de cálculo será igual ao custo de
        'aquisição dividido pelo valor antigo do fator; o custo do composto será igual ao
        'custo de aquisição multiplicado pela quantidade de peças vendidas. Se o produto foi
        'produzido pela Unocann, o fator de cálculo será igual ao custo do composto dividido
        'pelo antigo valor do fator. O novo valor do fator será armazenado na coluna 10 do
        'grid de índices.
        '*************************************************************************************
        
        dlIDX = (100 - dlSumGrd) / 100
        
        If blVlUnt = True Then
            dlIDX = Format(dlCusUntQtd / dlIDX, "###,###,##0.00")
            dlCusCompo = dlCusUntQtd * ilSumQtd
        Else
            dlIDX = Format(Rs!ValCusUnt / dlIDX, "##0.00")
        End If
        '*************************************************************************************
        'Calcula o percentual de participação do custo do composto sobre o preço de venda,
        'descontados os outros custos e a margem de lucro. O resultado será um preço de
        'composto mais alto que o cadastrado, já que tal preço terá sofrido os impactos dos
        'outros custos.
        '*************************************************************************************
        If dlIDX > 0 Then
            
        dlIDX = Format(Rs!ValCusUnt / dlIDX, "##0.00")
        End If
        
        '*************************************************************************************
        'Armazena em variável o peso bruto do produto.
        '*************************************************************************************
        
        dlSumPes = Trim(grid.TextMatrix(Linhas, 13))
        
        '*************************************************************************************
        'Calcula a margem de lucro.
        '*************************************************************************************
        
        '*************************************************************************************
        'Armazena em variável o preço total líqüido do produto.
        '*************************************************************************************
        
        dlSumGrd = Format((grid.TextMatrix(Linhas, 5) - (grid.TextMatrix(Linhas, 5) * 0.342) * 1), "##,##0.00")
        
        '*************************************************************************************
        'O fator de cálculo passa a ser a soma de todos os elementos de composição de preço do
        'produto, sem margem de lucro.
        '*************************************************************************************
        
        dlIDX = dlSumSubGrd
        
        '*************************************************************************************
        'Calcula custo de composto, outros custos de produção, valor da margem (em moeda) e
        'margem real de lucro (em percentual). As variáveis de cálculo mudam levando-se em
        'consideração se o produto foi adquirido para revenda ou produzido pela Unocann.
        '*************************************************************************************
        
        '*********************************************************************************
        'Calcula outros custos multiplicando o preço líqüido do produto pelo fator de
        'cálculo e dividinho o resultado por 100. Essa conta faz sentido, já que o fator
        'de cálculo foi obtido a partir da composição de preço do produto sem a margem de
        'lucro.
        '*********************************************************************************
        dlOutrosCus = Format((dlSumGrd * dlIDX) / 100, "####0.00")
        
        '*********************************************************************************
        'Calcula o custo do composto multiplicando o preço cobrado pelo fornecedor externo
        'pela quantidade vendida; multiplica-se também a quantidade vendida pelo custo
        'adicional de aquisição do produto. O custo do composto é a soma dos dois
        'resultados.
        '*********************************************************************************
        dlCusCompo = (dlCusUntQtd * ilSumQtd) + (ilSumQtd * dlCusAdiQtd)
        
        '*********************************************************************************
        'Calcula a margem de lucro em moeda subtraindo o preço líqüido do produto pelo
        'resultado da soma dos custos (composto e outros).
        '*********************************************************************************
        dlValMargem = dlSumGrd - (dlOutrosCus + dlCusCompo)
        dlValMargem = Format(dlValMargem, "0.00")
        
        '*********************************************************************************
        'Calcula a margem real de lucro (em percentual) dividindo o valor do lucro em
        'moeda pelo preço total líqüido e dividindo seu resultado por 100.
        '*********************************************************************************
        dlMargemReal = (dlValMargem / dlSumGrd) * 100
        
        '*************************************************************************************
        'Armazena no grid dos índices, pela ordem: margem real de lucro (em percentual), valor
        'da margem de lucro (em moeda) e a soma dos fatores que compõe o preço (sem margem).
        '*************************************************************************************
        grid.TextMatrix(Linhas, 15) = Format(dlMargemReal, "##0.00")
            
        Rs.Close
            
        Set Rs = Nothing
        
    Next Linhas

    Me.Refresh

    Exit Sub

TrataErro:

    'Rotina_Erro "CalculaIndice"
End Sub

Function EnableLinhaNF()

    ''MskNumNf.Enabled = True
    'BtoProduto.Enabled = True
    'MskSerie.Enabled = True
    'Bto_Aplica.Enabled = True
    'MskDatEmiNf.Enabled = True
    GrdNotaCliente.Enabled = True
    
End Function

Function LimpaLinhaNF()

    'MskNumNf.Limpar
    'MskSerie.Limpar
    'MskVlrUnit.Limpar
    ''MskDatEmiNf.Limpar
    LblRotaRec.Caption = ""
    ilNumTab = 0
    blimpa = True
    ''VSValUnit.Value = 0
    ''LblUnit.Caption = "Valor Unitário"
    
End Function

Private Sub Bto_Aplica_Click()

    iDscForaRegiao = 0

    If bgBloqPed = True Then
        Exit Sub
    End If

    If AplicaDescontoGrid = False Then
        
        'MskDatEmiNf.SetFocus
        
        Exit Sub
        
    End If
  
    '*******************************************************************************
    'UPDATE: 03/07/2017 Não tem necessidade de calcular indice neste momento
    '*******************************************************************************
    
    'If CalculaIndice = False Then
        'LimpaGeral
    'End If
  
'    If MskMargem.Texto < 8 Then
'        BtoGrava.Enabled = False
'        MsgBox "Esse pedido está fora da política de descontos da Unocann, " & vbCrLf & _
'        "ele só poderá ser gravado quando estiver adequado à política da empresa.", vbCritical, "ATENÇÃO!"
'        Me.Caption = "UNOCANN Tubos e Conexões - Simulação de Pedidos  [" & Trim(slNomRep) & "]"
'    Else
'        Me.Caption = "UNOCANN Tubos e Conexões - Manutenção de Pedidos  [" & Trim(slNomRep) & "]"
'
'    End If
  
  
    ''MskDatEmiNf.Limpar
    'BtoAdiNF.SetFocus
  
End Sub

Private Sub BtoAdiNF_Click()

    Dim dlPesBru As Double
    Dim dlDesc   As Double
    
    If bgBloqPed = True Then
        Exit Sub
    End If
    
    '*****************************************************************************************
    'Zera as variáveis de cálculo e definição das cores. Define a comissão do representante.
    '*****************************************************************************************
    
    dlPerComiNeg = 0
    slClasCor = ""
    blFechaComi = False
    dlPerComiCalc = dlPerComiN
    
    '*****************************************************************************************
    'A lista de pedido não pode conter mais de 65 itens.
    '*****************************************************************************************
    
    If GrdNotaCliente.rows > 65 Then
    
        MsgBox "É pertimido a inclusão de até 65 itens por pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        LimpaLinhaNF
        
        'MskNumNf.SetFocus
        
        Exit Sub
        
    End If
    
    '*****************************************************************************************
    '
    '*****************************************************************************************
    
    If Me.ActiveControl.Name = "BtoSair" Or Me.ActiveControl.Name = "BtoLimpaCTRC" Then
        Exit Sub
    End If
    
    '*****************************************************************************************
    'Faz consistência do campo "Desconto". Se estiver vazio, recebe um valor numérico (zero).
    '*****************************************************************************************
    
    If Trim(MskDatEmiNf) = "" Then
        MskDatEmiNf = 0
    End If
'-----------------------------------------------------------------
If ilCodCnd = 1 Or ilCodCnd = 12 Or ilCodCnd = 24 Then 'A vista ou 14 dias
   bAVista = True
Else
   bAVista = False
End If

    If bAVista = True Then 'estou trabalhando a vista
            
                If Trim(MskDatEmiNf) <= dDscRegiao Then
                    iDscRegiao = iDscRegiao + 1
                Else
                    iDscForaRegiao = iDscForaRegiao + 1
                End If
                
             Else
            
                If Trim(MskDatEmiNf) <= dDscRegiao Then
                    iDscRegiao = iDscRegiao + 1
                Else
                    iDscForaRegiao = iDscForaRegiao + 1
                End If
            
            
            End If
'----------------------------------------------------------------
    '*****************************************************************************************
    'Faz a consistência do campo "Código do Produto".
    '*****************************************************************************************
    
    If Trim(MskNumNf) = "" Or Trim(MskNumNf) = 0 Then
    
        MsgBox "Informe o Código do produto.", vbExclamation + vbOKOnly, "Atenção!"
        
        'MskNumNf.SetFocus
        
        Exit Sub
        
    End If
    
    '*****************************************************************************************
    'Se o usuário indicou que o pedido é de um Kit Irrigação, o sistema verifica se o produto
    'a ser inserido na lista pode fazer parte de um kit. Se não puder, não entra na lista.
    '*****************************************************************************************
    
    If ChkKit.Value = 1 And ilFlgKit = 0 Then
    
        MsgBox "Este Produto não compõe a linha de irrigação !", vbExclamation + vbOKOnly, "Atenção!"
        
        'MskNumNf.SetFocus
        
        Exit Sub
        
    End If
    
    '*****************************************************************************************
    'Faz a consistência do campo "Quantidade".
    '*****************************************************************************************
    
    If Trim(MskSerie) = "" Or Trim(MskSerie) = 0 Then
    
        MsgBox "Informe a Quantidade do produto.", vbExclamation + vbOKOnly, "Atenção!"
        
        'MskSerie.SetFocus
        
        Exit Sub
        
    End If
    
    '*****************************************************************************************
    'Aplica formato de moeda ao desconto e avalia se seu valor respeita o limite permitido
    'para a venda.
    '*****************************************************************************************
    
    dlDesc = Format(Trim(MskDatEmiNf), "##0.00")
    
    If dlDesc > dlSumDscItem Then
    
        MsgBox "Desconto informado maior que o permitido para esta venda.", vbExclamation + vbOKOnly, "Atenção!"
        
        ''MskDatEmiNf.SetFocus
        
        Exit Sub
        
    End If
    
    '*****************************************************************************************
    'Verifica se o produto atual já existe na lista. Pode até haver inclusão duplicada de um
    'mesmo produto, porém a quantidade informada deve ser diferente daquela que já foi
    'inserida.
    '*****************************************************************************************
    
    If blModificar = False Then
    
        For Linhas = 1 To GrdNotaCliente.rows - 1
        
            If GrdNotaCliente.TextMatrix(Linhas, 0) = Trim(MskNumNf) And GrdNotaCliente.TextMatrix(Linhas, 1) = Trim(MskSerie) Then
            
                MsgBox "Nota Fiscal existente.", vbExclamation + vbOKOnly, "Atenção!"
                
                ''MskDatEmiNf.SetFocus
                
                Exit Sub
                
            End If
            
        Next Linhas
        
    End If
    
    '*****************************************************************************************
    'A rotina a seguir insere os itens do pedido no grid.
    '*****************************************************************************************
    
    '*****************************************************************************************
    'Segue abaixo uma lista com a posição de cada informação no grid:
    '0: Código do produto
    '1: Descrição do produto
    '2: Quantidade vendida por embalagem
    '3: Quantidade vendida
    '4: Preço unitário (bruto ou com desconto?)
    '5: Identificação da tabela aplicada (vazio para tabela normal, A e B)
    '6: Percentual de desconto aplicado para o produto
    '7: Valor total líqüido do produto
    '8: Código do grupo de produto
    '9: Valor total bruto do produto
    '10: Peso bruto do produto
    '11: Preço ideal da venda
    '12: Preço unitário bruto
    '13: Margem de lucro
    '14: Kit Irrigação?
    '15: Custo unitário de aquisição do produto
    '16: Custo unitário adicional de aquisição do produto
    '17: Alíquota de imposto federal
    '18: Demonstração gráfica da margem do produto. Margem individual.
    '*****************************************************************************************
    
    If blModificar = True Then
    
        '*************************************************************************************
        'Consolida as alterações realizadas no pedido.
        '*************************************************************************************
    
        GrdNotaCliente.RowSel = ilCelula
        GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 3) = IIf(Trim(MskSerie) = "", "0", MskSerie)
        GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 4) = IIf(Trim(MskVlrUnit) = "", "0", Format(MskVlrUnit, "##,###,##0.00"))
        
        If ilNumTab = 0 Then
            
            GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 5) = ""
            
        Else
        
            If ilNumTab = 1 Then
                GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 5) = "A"
            Else
                GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 5) = "B"
            End If
        
        End If
        
        ilQtdEmb = GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 2)
        
        GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 6) = IIf(Trim(MskDatEmiNf) = "", "0", Format(MskDatEmiNf, "##0.00"))
'-----------------------------------------------------------
If ilCodCnd = 1 Or ilCodCnd = 12 Or ilCodCnd = 24 Then 'A vista ou 14 dias
   bAVista = True
Else
   bAVista = False
End If
        
        If bAVista = True Then 'estou trabalhando a vista
            
                If GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 6) <= dDscRegiao Then
                    iDscRegiao = iDscRegiao + 1
                Else
                    iDscForaRegiao = iDscForaRegiao + 1
                End If
                
             Else
            
                If GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 6) <= dDscRegiao Then
                    iDscRegiao = iDscRegiao + 1
                Else
                    iDscForaRegiao = iDscForaRegiao + 1
                End If
            
            
            End If
'-----------------------------------------------------------
        
        
        If Trim(MskDatEmiNf) = 0 Then
            dlValItem = Trim(MskVlrUnit) * (Trim(MskSerie) * ilQtdEmb)
        Else
            dlValItem = (MskVlrUnit * (1 - (MskDatEmiNf / 100))) * (Trim(MskSerie) * ilQtdEmb)
        '    dlValItem = (MskVlrUnit.Texto - ((MskVlrUnit.Texto * MskDatEmiNf) / 100)) * (Trim(MskSerie.Texto) * ilQtdEmb)
        End If
        
        GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 18) = Format(dlValItem, "##,###,##0.00")
        GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 9) = Format(Trim(MskVlrUnit) * (Trim(MskSerie) * ilQtdEmb), "##,###,##0.00")
        
        dlPesBru = dlPesUnt * (Trim(MskSerie) * ilQtdEmb)
        
        GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 10) = Format(dlPesBru, "###,##0.0000")
        
        VlIdealItem = (dlValUntN - ((dlValUntN * dlSumDscItemORIG) / 100)) * (Trim(MskSerie) * ilQtdEmb)
        
        GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 11) = Format(VlIdealItem, "###,##0.00")
        
    Else
    
        '*************************************************************************************
        'Inclui o produto na lista.
        '*************************************************************************************
    
        ilind = GrdNotaCliente.rows
        
        GrdNotaCliente.rows = GrdNotaCliente.rows + 1
        GrdNotaCliente.TextMatrix(ilind, 0) = Format(Trim(MskNumNf), "0000")
        GrdNotaCliente.TextMatrix(ilind, 1) = Trim(LblRotaRec.Caption)
        GrdNotaCliente.TextMatrix(ilind, 2) = ilQtdEmb
        GrdNotaCliente.TextMatrix(ilind, 3) = IIf(Trim(MskSerie) = "", "0", MskSerie)
        GrdNotaCliente.TextMatrix(ilind, 4) = IIf(Trim(MskVlrUnit) = "", "0", Format(MskVlrUnit, "##,###,##0.00"))
        
        '*************************************************************************************
        'Identifica a tabela aplicada: vazio para tabela normal ou tabelas A e B.
        '*************************************************************************************
        
        If ilNumTab = 0 Then
            
            GrdNotaCliente.TextMatrix(ilind, 5) = ""
            
        Else
        
            If ilNumTab = 1 Then
                GrdNotaCliente.TextMatrix(ilind, 5) = "A"
            Else
                GrdNotaCliente.TextMatrix(ilind, 5) = "B"
            End If
            
        End If
        
        GrdNotaCliente.TextMatrix(ilind, 6) = IIf(Trim(MskDatEmiNf) = "", "0", Format(MskDatEmiNf, "##0.00"))
        
        '*************************************************************************************
        'Se houver desconto aplicado ao produto, o valor total do item será o preço bruto,
        'menos o desconto aplicado, vezes a quantidade vendida. Se não houver desconto, o
        'valor total do item será o preço bruto vezes a quantidade vendida.
        '*************************************************************************************
        
        If Trim(MskDatEmiNf) = 0 Then
            dlValItem = Trim(MskVlrUnit) * (Trim(MskSerie) * ilQtdEmb)
        Else
            dlValItem = (MskVlrUnit * (1 - (MskDatEmiNf / 100))) * (Trim(MskSerie) * ilQtdEmb)
            'dlValItem = (MskVlrUnit - ((MskVlrUnit.Texto * MskDatEmiNf) / 100)) * (Trim(MskSerie.Texto) * ilQtdEmb)
        End If
        
        GrdNotaCliente.TextMatrix(ilind, 18) = Format(dlValItem, "##,###,##0.00")
        GrdNotaCliente.TextMatrix(ilind, 8) = ilIdeGrp
        GrdNotaCliente.TextMatrix(ilind, 9) = Format(Trim(MskVlrUnit) * (Trim(MskSerie) * ilQtdEmb), "##,###,##0.00")
        
        '*************************************************************************************
        'O cálculo do peso bruto do produto é simples: multiplica-se o peso unitário da peça
        'pelo resultado da quantidade de embalagens vendidas vezes a quantidade de produtos
        'que compõe cada embalagem.
        '*************************************************************************************
        
        dlPesBru = dlPesUnt * (Trim(MskSerie) * ilQtdEmb)
        
        GrdNotaCliente.TextMatrix(ilind, 10) = Format(dlPesBru, "###,##0.0000")
        
        '*************************************************************************************
        'A linha a seguir calcula aquilo que foi chamado de "valor ideal" do produto: trata-se
        'do preço total líqüido. Importante frisar que o desconto é a soma de todos aqueles
        'encontrados para o perfil da venda, independente de haver chave para o pedido ou não.
        '*************************************************************************************
        
        VlIdealItem = (dlValUntN - ((dlValUntN * dlSumDscItemORIG) / 100)) * (Trim(MskSerie) * ilQtdEmb)
        
        GrdNotaCliente.TextMatrix(ilind, 11) = Format(VlIdealItem, "###,##0.00")
        GrdNotaCliente.TextMatrix(ilind, 12) = Format(dlValUntN, "###,##0.00")
        GrdNotaCliente.TextMatrix(ilind, 13) = Format(dlMrgPrd, "##0.000")
        GrdNotaCliente.TextMatrix(ilind, 14) = ilFlgKit
        GrdNotaCliente.TextMatrix(ilind, 15) = Format(dlValCusUntQtd, "###,##0.00")
        GrdNotaCliente.TextMatrix(ilind, 16) = Format(dlValCusAdicQtd, "###,##0.00")
        GrdNotaCliente.TextMatrix(ilind, 17) = Format(dlAlqImpFed, "###,##0.000")
        
    End If
    
    '*****************************************************************************************
    'Limpra os campos por onde o produto foi inserido.
    '*****************************************************************************************
    
    LimpaLinhaNF
    
    '*****************************************************************************************
    'Habilita os campos para uma possível inserção de novo produto. Também desabilita o campo
    'que permite a marcação de Kit Irrigação; como o primeiro item já foi aceito, o pedido
    'atual já foi definido como Kit e não pode mais deixar esse status.
    '*****************************************************************************************
    
    ''MskNumNf.Enabled = True
    'BtoProduto.Enabled = True
    'MskSerie.Enabled = True
    'Bto_Aplica.Enabled = True
    'BtoExcNF.Enabled = False
    'BtoAdiNF.Enabled = True
    'MskNumNf.SetFocus
    
    blModificar = False
    
    If bgSimula = False Then
        BtoGrava.Enabled = True
    End If
    
    GrdNotaCliente.Enabled = True
    
    ilNumTab = 0
    
    LblResultNegocio.Caption = ""
    ChkKit.Enabled = False
    
    '*****************************************************************************************
    'Calcula os índices do pedido. Se houver qualquer problema durante o cálculo, todo o
    'formulário será limpo e o pedido anulado.
    '*****************************************************************************************
    
    If CalculaIndice = False Then
        
        LimpaGeral
        
    Else
        
        ilind = ilind - 3
        
            GrdNotaCliente.col = 19
            GrdNotaCliente.row = ilind
            
         If APLICA = 1 Then
            
         Else
            If GrdIndice.TextMatrix(ilind, 15) < 10 Then
                GrdNotaCliente.CellBackColor = &HFF&
            ElseIf GrdIndice.TextMatrix(ilind, 15) >= 10 And GrdIndice.TextMatrix(ilind, 15) < 15 Then
                GrdNotaCliente.CellBackColor = &HFF00&
            ElseIf GrdIndice.TextMatrix(ilind, 15) >= 15 Then
                GrdNotaCliente.CellBackColor = &HFF0000
            End If
        End If
    End If
    
'    If MskMargem.Texto < 8 Then
'        BtoGrava.Enabled = False
'        MsgBox "Esse pedido está fora da política de descontos da Unocann, " & vbCrLf & _
'               "ele só poderá ser gravado quando estiver adequado à política da empresa.", vbCritical, "ATENÇÃO!"
'                    Me.Caption = "UNOCANN Tubos e Conexões - Simulação de Pedidos  [" & Trim(slNomRep) & "]"
'
'    Else
'            Me.Caption = "UNOCANN Tubos e Conexões - Manutenção de Pedidos  [" & Trim(slNomRep) & "]"
'
'    End If
    
    
End Sub
'-------------------------------------------------------------
Private Sub TrocaCIF_FOB()

    If bgBloqPed = True Then
        Exit Sub
    End If

    dlPerComiNeg = 0
    slClasCor = ""
    dlPerComiCalc = dlPerComiN
    blFechaComi = False
    
    If GrdNotaCliente.rows = 1 Then
        Exit Sub
    End If
    
    If Opt_CIF.Value = True Then
    
        'MskNumNf.Enabled = True
        dlPerDesFOBReal = 0
        TxtTransp.Text = "UNOCANN TRANSPORTES LTDA"
        
        CalculaDesconto
        
        If CalculaIndice = False Then
            LimpaGeral
        End If
        
        'MskNumNf.SetFocus

    End If

    If Opt_FOB.Value = True Then
            
        'MskNumNf.Enabled = True
        dlPerDesFOBReal = dlPerDesFOB
        TxtTransp.Text = "O PROPRIO"
        
        CalculaDesconto
        
        If CalculaIndice = False Then
            LimpaGeral
        End If
        
        'MskNumNf.SetFocus
            
            
    End If

    
'    sgQuery = MsgBox("Deseja remover o produto " & GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 0) & " ?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenção!")
    
'    If sgQuery = vbNo Then
 '       GoTo Continua
 '   End If
    
    
Continua:
    
'    LimpaLinhaNF
'
'    MskNumNf.Enabled = True
'    BtoProduto.Enabled = True
'    MskSerie.Enabled = True
'    'Bto_Aplica.Enabled = True
'    BtoExcNF.Enabled = False
'    BtoAdiNF.Enabled = True
'    MskNumNf.SetFocus
'    BtoAdiNF.Caption = "&Adicionar"
'
'    blModificar = False
'
'    If CalculaIndice = False Then
'        LimpaGeral
'    End If
    
End Sub





'-------------------------------------------------------------



Private Sub BtoExcNF_Click()

    If bgBloqPed = True Then
        Exit Sub
    End If

    dlPerComiNeg = 0
    slClasCor = ""
    dlPerComiCalc = dlPerComiN
    blFechaComi = False
    
    If GrdNotaCliente.rows = 1 Then
        Exit Sub
    End If
    
    sgQuery = MsgBox("Deseja remover o produto " & GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 0) & " ?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenção!")
    
    If sgQuery = vbNo Then
        GoTo Continua
    End If
    
    If GrdNotaCliente.RowSel > 0 Then
        
        If GrdNotaCliente.rows > 2 Then
            
            GrdNotaCliente.RemoveItem GrdNotaCliente.RowSel
            
        Else
        
            GrdNotaCliente.rows = 1
            'ChkKit.Enabled = True
            GrdNotaCliente.Enabled = False
            
        End If
        
    End If
    
Continua:
    
    LimpaLinhaNF
    
    ''MskNumNf.Enabled = True
    'BtoProduto.Enabled = True
    'MskSerie.Enabled = True
    'Bto_Aplica.Enabled = True
    'BtoExcNF.Enabled = False
    'BtoAdiNF.Enabled = True
    'MskNumNf.SetFocus
    'BtoAdiNF.Caption = "&Adicionar"
    
    blModificar = False
    
    If CalculaIndice = False Then
        LimpaGeral
    End If
    
End Sub

Private Sub BtoGrava_Click()
    
    DoEvents
    
    '*********************************************
    'Pedido com apenas tubo de 100 não pode ser gravado
    'A não ser que sua margem supere os 9%
    '***********************************************
    
    If dlMargemGeral < 9 Then
        If bSo100 = True Then
            MsgBox "Pedido com item apenas TUBO DE 100MM," & vbCr & _
            "não pode ser gravado," & vbCr & _
            "inclua novos itens ou feche o Pedido!", vbExclamation + vbOKOnly, "Atenção!"
         '   MskNroPedido
            Exit Sub
        End If
        
        If (dlTotLiq * 0.7) < dTotb100 Then
            MsgBox "Esse Pedido está com item TUBO DE 100MM," & vbCr & _
            "maior que 70% do total permitido para venda!", vbExclamation + vbOKOnly, "Atenção!"
         '   MskNroPedido
            Exit Sub
        End If
    End If
    
    
    '*****************************************************************************************
    'Verifica se a condição de pagamento foi informada.
    '*****************************************************************************************
    
    If Trim(ilCodCnd) = "" Or Trim(ilCodCnd) = 0 Or Trim(CboCondPag.Criterio) = "" Then
        
        MsgBox "Informe a Condição de pagamento", vbExclamation + vbOKOnly, "Atenção!"
        
        CboCondPag.SetFocus
        
        Exit Sub
        
   End If
   
    If bgBloqPed = True Then
        Exit Sub
    End If
            
    cboCli_Consultar
    
    '*****************************************************************************************
    'Verifica se o cliente foi informado.
    '*****************************************************************************************
    
    If Trim(slremet) = "" Or ilCodCli = 0 Or Trim(ilCodCli) = "" Then
        
        MsgBox "Informe o Cliente.", vbExclamation + vbOKOnly, "Atenção!"
        
        slremet = ""
        ilCodCli = 0
        
        cboCli.Habilitado = True
        cboCli.SetFocus
        
        Exit Sub
        
    End If
      
    cboCli_LostFocus
   
    blleitura = False
    
    '*****************************************************************************************
    'Faz a leitura dos dados do cliente e exibe na guia correspondente. Apura também dados
    'sobre descontos, tributos e calcula os índices do pedido.
    '*****************************************************************************************
    
    If LeituraCliente = False Then
        LimpaGeral
    End If
   
    If blDescZero = True And d5.Texto > 0 Then
    
        sgQuery = MsgBox("Existe item sem Desconto, Deseja voltar no pedido?", vbQuestion + vbYesNo + vbDefaultButton1, "Atenção!")
        
        If sgQuery = vbYes Then
            Exit Sub
        End If
        
    End If

    slAceita = False
    dlComiSug = 0
    
    If slFlgSugComi = "S" And bgPedMKT = False Then
    
        If EquilibraComissao = False Then
            Exit Sub
        End If
        
    End If
    
    If ConferePrazo = False Then
        Exit Sub
    End If
    
    If ChkKit.Value = 1 Then
    
        If ValidaKit = False Then
            Exit Sub
        End If
        
    End If
    
    If slAceita = True Then
    
        sgQuery = MsgBox("Confirma Pedido?", vbQuestion + vbYesNo + vbDefaultButton1, "Atenção!")
        
        If sgQuery = vbNo Then
            Exit Sub
        End If
        
    End If
    
    If Trim(TxtTransp.Text) = "" Then
        
        MsgBox "Informe a Transportadora !", vbExclamation + vbOKOnly, "Atenção!"
        
        TxtTransp.SetFocus
        
        Exit Sub
        
    End If
   
'----------------------------------------------
    'If iDscForaRegiao < 1 Then
    
        'BtoGrava.Enabled = False
        
        'Exit Sub
        
    'Else
        
        'If iDscForaRegiao >= 1 Then
            
            If Date >= CDate("04/02/2017") Then
                
                If dlMargemGeral <= 7.99 Then
                    MsgBox "Esse pedido está FORA da política comercial da Unocann. " & vbCrLf & _
                    "Corrija os descontos praticados até que essa mensagem não apareça.", vbCritical, "Atenção!"
                    
                    Me.Caption = "UNOCANN Tubos e Conexões - Simulação de Pedidos  [" & Trim(slNomRep) & "]"
                
                    Exit Sub
                
                ElseIf dlMargemGeral > 8 And dlMargemGeral <= 8.5 Then
                
                    MsgBox "Esse pedido está PRÓXIMO da política comercial da Unocann. " & vbCrLf & vbCrLf & _
                    "Reveja os descontos praticados e mix de produtos até que a cor fique VERDE.", vbInformation, "Atenção!"
            
           '     MsgBox "Reveja os descontos praticados e mix de produtos até que a cor fique VERDE.", vbInformation, "Atenção!"
                       
           '     MsgBox "Reveja os descontos praticados e mix de produtos até que a cor fique VERDE.", vbInformation, "Atenção!"
                       
                       
                Me.Caption = "UNOCANN Tubos e Conexões - Simulação de Pedidos  [" & Trim(slNomRep) & "]"
        
                
                End If
            
            Else
            
                If dlMargemGeral <= 7 Then
                    MsgBox "Esse pedido está FORA da política comercial da Unocann. " & vbCrLf & _
                    "Corrija os descontos praticados até que essa mensagem não apareça.", vbCritical, "Atenção!"
                    
                    Me.Caption = "UNOCANN Tubos e Conexões - Simulação de Pedidos  [" & Trim(slNomRep) & "]"
                
                    Exit Sub
                
                ElseIf dlMargemGeral > 7 And dlMargemGeral <= 8.5 Then
                
                    MsgBox "Esse pedido está PRÓXIMO da política comercial da Unocann. " & vbCrLf & vbCrLf & _
                    "Reveja os descontos praticados e mix de produtos até que a cor fique VERDE.", vbInformation, "Atenção!"
            
           '     MsgBox "Reveja os descontos praticados e mix de produtos até que a cor fique VERDE.", vbInformation, "Atenção!"
                       
           '     MsgBox "Reveja os descontos praticados e mix de produtos até que a cor fique VERDE.", vbInformation, "Atenção!"
                       
                       
                Me.Caption = "UNOCANN Tubos e Conexões - Simulação de Pedidos  [" & Trim(slNomRep) & "]"
        
                
                End If
           End If
      'Else
            
            'Me.Caption = "UNOCANN Tubos e Conexões - Manutenção de Pedidos  [" & Trim(slNomRep) & "]"
        
      'End If

    'End If

    

    
'----------------------------------------------
   
    If bgPedMKT = True Then
        
        GravaCTRCTMK
        
        Me.Refresh
        
        MsgBox "Pedido " & MskNroPedido.Text & " incluido com sucesso!", vbExclamation + vbOKOnly, "Atenção!"
        
        BtoSair_Click
        
        Exit Sub
        
    Else
        
        GravaCTRC
    End If
       
    If bgConsultaPed = True Then
    
        Unload Me
        
        Set FrmConhecimento = Nothing
        
        bgBloqPed = False
        bgConsultaPed = False
        
        FrmPosiPed.Enabled = True
        FrmPosiPed.Show
        
        Exit Sub
        
    End If
   
    LimpaGeral
    
    Me.Refresh
   
End Sub

Private Sub BtoLimpaCTRC_Click()

    DoEvents
    
    If bgConsultaPed = True Then
        Exit Sub
    End If
   
    If sgFlagOper = "A" Then
        
        sgQuery = MsgBox("Deseja Atualizar o Pedido?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenção!")
        
        If sgQuery = vbYes Then
            
            BtoGrava_Click
            
            Exit Sub
        
        End If
    
    End If

    LimpaGeral
    
    DoEvents
    
End Sub

Private Sub BtoLimpaNF_Click()

    If bgBloqPed = True Then
        Exit Sub
    End If

    'BtoLimpaNF.Caption = "&Limpar"
    'BtoImprime.Enabled = False
    
    LimpaLinhaNF
    EnableLinhaNF
    
    'BtoExcNF.Enabled = False
    'BtoAdiNF.Enabled = True
    'MskNumNf.SetFocus
    
    blModificar = False
    
End Sub

Private Sub BtoProduto_Click()
    
    CarregaGridGrupo
    
    FraGrupo.Visible = True
    
    DoEvents
    
    GrdGrupo.RowSel = 1
    GrdGrupo_Click
    GrdGrupo.SetFocus
    
End Sub

Private Sub BtoSair_Click()
    
    DoEvents
     
    If sgFlagOper = "A" And bgBloqPed = False Then
        
        sgQuery = MsgBox("Deseja Atualizar o Pedido?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenção!")
        
        If sgQuery = vbYes Then
            
            If iDscForaRegiao < 1 Then
                BtoGrava.Enabled = True
            Else
                If iDscForaRegiao > 1 Then
                
                    If dlMargemGeral <= 7 Then
                
                    MsgBox "Esse pedido está FORA da política comercial da Unocann. " & vbCrLf & _
                                      "Corrija os descontos praticados até que essa mensagem não apareça.", vbCritical, "Atenção!"

                    ' "Deseja GRAVAR esse Pedido mesmo assim ?", vbExclamation + vbYesNo + vbDefaultButton2, "Atenção!")
                    Me.Caption = "UNOCANN Tubos e Conexões - Simulação de Pedidos  [" & Trim(slNomRep) & "]"
                    
                   ' If sgQuery = vbNo Then
                        Exit Sub
                   ' End If
                
                ElseIf dlMargemGeral >= 7 And dlMargemGeral <= 9 Then
                
                     MsgBox "Esse pedido está PRÓXIMO da política comercial da Unocann. " & vbCrLf & _
                             "Reveja os descontos praticados e mix de produtos para facilitar a liberação." & vbCrLf & vbCrLf & _
                             "Deseja GRAVAR esse Pedido mesmo assim ?", vbInformation, "Atenção!"
                             
                    MsgBox "Reveja os descontos praticados e mix de produtos até que a cor fique VERDE.", vbInformation, "Atenção!"
                   
            MsgBox "Reveja os descontos praticados e mix de produtos até que a cor fique VERDE.", vbInformation, "Atenção!"
                             
                Me.Caption = "UNOCANN Tubos e Conexões - Simulação de Pedidos  [" & Trim(slNomRep) & "]"
'
                 End If
                    
          Else
                    
                    Me.Caption = "UNOCANN Tubos e Conexões - Manutenção de Pedidos  [" & Trim(slNomRep) & "]"
                    
          End If
                
        End If
                
            
        End If
            
            
            
            
            '            If MskMargem.Texto < 8 Then
            '                    BtoGrava.Enabled = False
            '                    MsgBox "Esse pedido está fora da política de descontos da Unocann, " & vbCrLf & _
            '                    "ele só poderá ser gravado quando estiver adequado à política da empresa.", vbCritical, "ATENÇÃO!"
            '                    Me.Caption = "UNOCANN Tubos e Conexões - Simulação de Pedidos  [" & Trim(slNomRep) & "]"
            '                Else
            '                    Me.Caption = "UNOCANN Tubos e Conexões - Manutenção de Pedidos  [" & Trim(slNomRep) & "]"
            '                    BtoGrava_Click
            '                End If
            
        '    Exit Sub
            
        'End If
    
    End If

    Unload Me
    
    Set FrmConhecimento = Nothing
 
    If bgConsultaPed = True Then
    
        bgBloqPed = False
        bgConsultaPed = False
        
        If igTela = "Monit" Then
            
            igTela = ""
            
            FrmPosMonit.Enabled = True
            FrmPosMonit.Show
            
        Else
            
            If igTela = "PosPed" Then
                FrmPosiPed.Enabled = True
                FrmPosiPed.Show
            Else
                FrmTMKPrincipal.Enabled = True
                FrmTMKPrincipal.Show
            End If
            
        End If
    
    End If
 
    'lgSeqLig = 0

    iDscForaRegiao = 0

    If bgPedMKT = True Then
        FrmTMKPrincipal.Enabled = True
        FrmTMKPrincipal.Show
    End If

End Sub

Private Sub CboCondPag_Consultar()

    'CboCondPag.query = "Select Condição = DscCnd, Cod = CodCnd, 'Prazo Médio' = PrzMed  From CONDICAO " & _
    '"Where blqcnd <> 'S' and " & IIf(IsNumeric(CboCondPag.Criterio), "CodCnd", "DscCnd") & " Like '" & CboCondPag.Criterio & "%'"

    CboCondPag.query = "Select Condição = DscCnd, Cod = CodCnd, 'Prazo Médio' = PrzMed  From CONDICAO " & _
    "Where blqcnd <> 'S' and DscCnd Like '" & CboCondPag.Criterio & "%'"

End Sub

Private Sub CboCondPag_GotFocus()
    
    DoEvents
    
    If ControleLostFocus = False Then
        Exit Sub
    End If
    
    If Me.ActiveControl.Name = "BtoSair" Or Me.ActiveControl.Name = "BtoLimpaCTRC" Or Me.ActiveControl.Name = "MskNroPedido" Or blLG = True Then
        Exit Sub
    End If

    'If Trim(slremet) = "" Or ilCodCli = 0 Or Trim(ilCodCli) = "" Then
        
        'MsgBox "Informe o Cliente.", vbExclamation + vbOKOnly, "Atenção!"
        
        'slremet = ""
        'ilCodCli = 0
        
        'cboCli.Habilitado = True
        'cboCli.SetFocus
        
        'Exit Sub
    
    'End If


    If Trim(cboCli.codigo) > 0 Then
    
        SSTConhec.TabEnabled(1) = True
        SSTConhec.TabEnabled(2) = True

        'slremet = cboCli.Criterio
        'ilCodCli = cboCli.Codigo
        'blleitura = False
        
        'MskNumNf.Enabled = True
        'MskSerie.Enabled = True
        ''Bto_Aplica.Enabled = True
        'MskDatEmiNf.Enabled = True
        
        'If blModificar = False Then
            'MskNumNf.Enabled = True
            'BtoProduto.Enabled = True
        'End If
        
    'Else
        
        'slremet = ""
        
    End If

    'If LeituraCliente = False Then
        'LimpaGeral
    'End If

    'If blVencidos = True And sgFlagOper <> "A" And blRetornoDupls = False Then
        
        'MsgBox "Existe(m) Título(s) vencido(s) para este cliente!", vbExclamation + vbOKOnly, "Atenção!"
        
        'blVencidos = False
        'blRetornoDupls = True
        
        'SSTConhec.Tab = 2
        
        'Exit Sub
        
    'End If

    Call SelecionaTudo
    
End Sub

Private Sub CboCondPag_LostFocus()
    
    If ControleLostFocus = False Then
        Exit Sub
    End If
    
    If Me.ActiveControl.Name = "BtoSair" Or Me.ActiveControl.Name = "BtoLimpaCTRC" Or Me.ActiveControl.Name = "MskNroPedido" Then
        Exit Sub
    End If
  
    'If Trim(slremet) = "" Or ilCodCli = 0 Or Trim(ilCodCli) = "" Then

    If (Trim(slremet) = "" Or Trim(cboCli.codigo) = "") And bgSimula = False Then
        
        MsgBox "Informe o Cliente.", vbExclamation + vbOKOnly, "Atenção!"
        
        slremet = ""
        ilCodCli = 0
        
        cboCli.Habilitado = True
        cboCli.SetFocus
        
        Exit Sub
        
    End If

    If (Trim(cboCli.codigo) > 0 And Trim(cboCli.codigo) <> "") Or bgSimula = True Then
    
        slremet = cboCli.Criterio
        ilCodCli = IIf(cboCli.codigo = "", 0, cboCli.codigo)
        blleitura = False
        
        ''MskNumNf.Enabled = True
        'MskSerie.Enabled = True
        ''Bto_Aplica.Enabled = True
        'MskDatEmiNf.Enabled = True
        
        If blModificar = False Then
            ''MskNumNf.Enabled = True
            ''BtoProduto.Enabled = True
        End If
    
    Else
        
        slremet = ""
        
    End If
    
    SSTConhec.TabEnabled(1) = True
    SSTConhec.TabEnabled(2) = True

    If LeituraCliente = False Then
        
        Exit Sub
        
        'LimpaGeral
        
    End If

    If blVencidos = True And sgFlagOper <> "A" And blRetornoDupls = False Then
     
        MsgBox "Existe(m) Título(s) vencido(s) para este cliente!", vbExclamation + vbOKOnly, "Atenção!"
        
        blVencidos = False
        blRetornoDupls = True
        
        'SSTConhec.Tab = 2
        
        'Exit Sub
        
    End If
  
    If Trim(CboCondPag.Criterio) <> "" And Trim(CboCondPag.codigo) <> "" And Trim(CboCondPag.codigo) > 0 Then
    
        If LeituraCondicao = False Then
            
            LimpaGeral
            
            Exit Sub
            
        End If
        
    End If
  
    Opt_FOB.Enabled = True
    Opt_CIF.Enabled = True
  
    'MskNumNf.SetFocus
    
End Sub

Private Sub cboCli_Consultar()

    '*****************************************************************************************
    'Pesquisa os clientes encontrados com as expressões digitadas pelo usuário e armazena o
    'primeiro da lista em variável.
    '*****************************************************************************************
    
    slremet = ""
    
    cboCli.query = "Select NomCli As Cliente, CodCli As Código, CgcCli as CNPJ, FlgContr As Contribuinte From Cliente Where " & IIf(IsNumeric(cboCli.Criterio), "CodCli", "NomCli") & " Like '" & cboCli.Criterio & "%' order by " & IIf(IsNumeric(cboCli.Criterio), "CodCli", "NomCli")
    
    slremet = Trim(cboCli.Criterio)
    
End Sub

Private Sub cboCli_GotFocus()
    
    'On Error Resume Next
    
    'If Me.ActiveControl.Name = "BtoSair" Or _

        'Me.ActiveControl.Name = "BtoLimpaCTRC" Or _
        'Me.ActiveControl.Name = "MskNroPedido" Or _

        'blLG = True Then
        
        'Exit Sub
    'End If

    'If Trim(MskNroPedido.Text) = "" Or Trim(MskNroPedido.Text) = 0 Then
        
        'MsgBox "Informe o Número do Pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        'MskNroPedido.SetFocus
        
        'Exit Sub
    
    'End If

    'Call SelecionaTudo
    
End Sub

Private Sub cboCli_LostFocus()
    
    '*****************************************************************************
    'Define o nome e o código do cliente selecionado.
    '*****************************************************************************
    
    If Me.ActiveControl.Name = "BtoSair" Or Me.ActiveControl.Name = "BtoLimpaCTRC" Then
        Exit Sub
    End If
    
    If Trim(cboCli.codigo) > 0 And Trim(cboCli.codigo) <> "" Then
    
        slremet = cboCli.Criterio
        ilCodCli = cboCli.codigo
        
    End If
            
End Sub

Private Sub CmdCancelar_Click()

    Dim slResp As String

    slResp = MsgBox("Confirma Cancelamento do Pedido?", vbQuestion + vbYesNo + vbDefaultButton1, "Atenção!")
    
    If slResp = vbNo Then
        Exit Sub
    End If

    sgQuery = "Update PEDIDO set SitPed = 'C', DatLib = convert(datetime,'" & Format(CDate(Date) & " " & Time, "dd/mm/yyyy hh:mm:ss") & "',103), "
    sgQuery = sgQuery & " DatAtu = convert(datetime,getdate(),103)"
    sgQuery = sgQuery & " where nroped = '" & Trim(MskNroPedido.Text) & "'"
  
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing

    If bgConsultaPed = True Then
        
        Unload Me
        
        Set FrmConhecimento = Nothing
        
        bgBloqPed = False
        bgConsultaPed = False
        
        FrmPosiPed.Enabled = True
        FrmPosiPed.Show
        
        Exit Sub
    
    End If
 
    LimpaGeral

End Sub

Private Sub CmdDN_Click()
    
    Dim Num&
    Num& = ScrollText&(TxtNegocio, 1)
    
End Sub

Private Sub CmdImpr_Click()

    Dim slNomeImpr As String
    Dim logon As Integer
    
    #Const fDebug = True
    
    If APLICA = 1 Then
        
        sgQuery = " SELECT"
        sgQuery = sgQuery & " PEDIDO.NroPed, PEDIDO.Datped, PEDIDO.CIFOB, PEDIDO.NomTra, PEDIDO.TexObs,PEDIDO.ChvDsc,"
        sgQuery = sgQuery & " ITEM_PEDIDO.VlrIte, CONDICAO.DscCnd,CLIENTE.CodCli, CLIENTE.NomCli, CLIENTE.EndCli, CLIENTE.BaiCli,CLIENTE.CidCli, CLIENTE.CepCli,"
        sgQuery = sgQuery & " CLIENTE.CgcCli , CLIENTE.InsCli, CLIENTE.UFCli, CLIENTE.FlgContr, CLIENTE.FlgSIMBa, REPRESENTANTE.NomRep"
        sgQuery = sgQuery & " From"
        sgQuery = sgQuery & " PEDIDO , ITEM_PEDIDO , CONDICAO , CLIENTE, REPRESENTANTE "
        sgQuery = sgQuery & " where PEDIDO.nroped = ITEM_PEDIDO.nroped"
        sgQuery = sgQuery & "   and PEDIDO.codcnd = CONDICAO.codcnd"
        sgQuery = sgQuery & "   and PEDIDO.codcli = CLIENTE.codcli"
        sgQuery = sgQuery & "   and PEDIDO.CodRep = REPRESENTANTE.CodRep"
        sgQuery = sgQuery & "   and PEDIDO.nroped = '" & Trim(MskNroPedido.Text) & "'"
        
    Else
        
        sgQuery = " SELECT"
        sgQuery = sgQuery & " PEDIDO.NroPed, PEDIDO.Datped, PEDIDO.CIFOB, PEDIDO.NomTra, PEDIDO.TexObs, PEDIDO.ChvDsc"
        sgQuery = sgQuery & " ITEM_PEDIDO.VlrIte, CONDICAO.DscCnd,CLIENTE.CodCli, CLIENTE.NomCli, CLIENTE.EndCli, CLIENTE.BaiCli,CLIENTE.CidCli, CLIENTE.CepCli,"
        sgQuery = sgQuery & " CLIENTE.CgcCli , CLIENTE.InsCli, CLIENTE.UFCli, CLIENTE.FlgContr, CLIENTE.FlgSIMBa, REPRESENTANTE.NomRep"
        sgQuery = sgQuery & " From"
        sgQuery = sgQuery & " PEDIDO , ITEM_PEDIDO , CONDICAO , CLIENTE, REPRESENTANTE, USUARIO  "
        sgQuery = sgQuery & " where PEDIDO.nroped = ITEM_PEDIDO.nroped"
        sgQuery = sgQuery & "   and PEDIDO.codcnd = CONDICAO.codcnd"
        sgQuery = sgQuery & "   and PEDIDO.codcli = CLIENTE.codcli"
        sgQuery = sgQuery & "   and PEDIDO.CodRep = REPRESENTANTE.CodRep"
        sgQuery = sgQuery & "   and PEDIDO.CodUsuLib *= USUARIO.CodUsu"
        sgQuery = sgQuery & "   and PEDIDO.nroped = '" & Trim(MskNroPedido.Text) & "'"
        
        'sgQuery = "SELECT P.NroPed, P.DatPed, P.CIFOB, P.NomTra, P.TexObs, P.ChvDsc, IP.VlrIte, CD.DscCnd, CL.CodCli, CL.NomCli, CL.EndCli, CL.BaiCli, CL.CidCli, CL.CepCli, CL.CgcCli , CL.InsCli, CL.UFCli, CL.FlgContr, CL.FlgSIMBa, R.NomRep "
        'sgQuery = sgQuery & "FROM Pedido P "
        'sgQuery = sgQuery & "INNER JOIN Item_Pedido IP ON IP.NroPed = P.NroPed "
        'sgQuery = sgQuery & "INNER JOIN Condicao CD ON P.CodCnd = CD.CodCnd "
        'sgQuery = sgQuery & "INNER JOIN Cliente CL ON P.CodCli = CL.CodCli "
        'sgQuery = sgQuery & "INNER JOIN Representante R ON R.CodRep = P.CodRep "
        'sgQuery = sgQuery & "LEFT OUTER JOIN Usuario U ON U.CodUsu = P.CodUsuLib "
        'sgQuery = sgQuery & "WHERE P.NroPed = " & MskNroPedido.Text
        
    End If

    If APLICA = 1 Then
        slNomeImpr = App.Path & "\Relatorios\Pedido.rpt"
    Else
        slNomeImpr = App.Path & "\Relatorios\PedidoMatriz.rpt"
    End If

    With rptcontprop
        
        .ReportFileName = slNomeImpr
        .DiscardSavedData = True
        .SQLQuery = sgQuery
        If APLICA = 1 Then
            '.Connect = "DSN=" & "unocann" & ";UID=" & "sa" & ";PWD=" & "unocann2017;server = UNOCANN-PC"
            .Connect = "DSN=" & "unocann" & ";UID=" & "sa" & ";PWD=" & "sysadmpss1"
        Else
            .Connect = "DSN=" & "unocann" & ";UID=" & "sa" & ";PWD=" & "#unoforte5600!"
        End If
        
        If APLICA = 1 Then
            .WindowState = crptMaximized
        Else
            .Destination = crptToPrinter
        End If
        
        .Action = 1
        
    End With

    If APLICA = 0 Then
        
        sgQuery = "Update PEDIDO set flgimpr = 'S' where nroped = " & Trim(MskNroPedido.Text)
        
        Set Rs = Conexao.Execute(sgQuery)
        Set Rs = Nothing
        
    End If
    
End Sub

Private Sub CmdLibera_Click()

    If slFlgAlt = "L" Then
    
        MsgBox "Este pedido aguarda alteração solicitada pelo comercial", vbExclamation + vbOKOnly, "Atenção!"
        
        Exit Sub
        
    End If

    sgQuery = "Update PEDIDO set DatlibUno = convert(datetime,getdate(),103), CodUsulib = " & LgCodUsuSis
    sgQuery = sgQuery & " where nroped = " & Trim(MskNroPedido.Text)
    
    Set Rs = Conexao.Execute(sgQuery)
    Set Rs = Nothing
  
    CmdLibera.Visible = False
    CmdImpr.Visible = True
    CmdImpr.Enabled = True

End Sub

Private Sub CmdRetorna_Click()

    FraGrupo.Visible = False
    
    'If MskNumNf.Enabled = True Then
        'MskNumNf.SetFocus
    'End If
    
    'If MskSerie.Enabled = True Then
        'MskSerie.SetFocus
    'End If
    
End Sub

Private Sub cmdUP_Click()
    
    Dim Num&
    Num& = ScrollText&(TxtNegocio, -1)

End Sub


Private Sub Form_Activate()


    If bgSimula = True Then
    
        FraParametro.Visible = True
        cboCli.Habilitado = False
        ' BtoGrava.Enabled = False
        
    
    End If
    
    ilCodRep = Trim(sgRepresentante)
    
    Activate_Ped
    
    If ilCodCnd = 1 Or ilCodCnd = 12 Or ilCodCnd = 24 Then 'A vista ou 14 dias
       bAVista = True
        
        If bEKit = True Then 'Venda a vista Kit irrigação
        
            If ilCodRep = 2 Or ilCodRep = 7 Or ilCodRep = 8 Then
                dDscRegiao = 26.87
            ElseIf ilCodRep = 600 Or ilCodRep = 800 Or ilCodRep = 1001 Or ilCodRep = 2100 Or ilCodRep = 6000 Then
                dDscRegiao = 24.15
            ElseIf ilCodRep = 905 Or ilCodRep = 5000 Or ilCodRep = 5001 Then
                dDscRegiao = 22.37
'            Else
'                dDscRegiao = 24.15
            End If
        
        Else
            
            If ilCodRep = 2 Or ilCodRep = 7 Or ilCodRep = 8 Then
                dDscRegiao = 22.37
            ElseIf ilCodRep = 600 Or ilCodRep = 800 Or ilCodRep = 1001 Or ilCodRep = 2100 Or ilCodRep = 6000 Then
                dDscRegiao = 19.69
            ElseIf ilCodRep = 905 Or ilCodRep = 5000 Or ilCodRep = 5001 Then
                dDscRegiao = 20.36
'            Else
'                dDscRegiao = 22.37
            End If
                
        End If

    Else 'A Prazo
       bAVista = False
             If bEKit = True Then 'Venda a Normal Kit irrigação
        
            If ilCodRep = 2 Or ilCodRep = 7 Or ilCodRep = 8 Then
                dDscRegiao = 24.56
            ElseIf ilCodRep = 600 Or ilCodRep = 800 Or ilCodRep = 1001 Or ilCodRep = 2100 Or ilCodRep = 6000 Then
                dDscRegiao = 21.8
            ElseIf ilCodRep = 905 Or ilCodRep = 5000 Or ilCodRep = 5001 Then
                dDscRegiao = 19.96
'            Else
'                dDscRegiao = 21.8
            End If
        
        Else
            
            If ilCodRep = 2 Or ilCodRep = 7 Or ilCodRep = 8 Then
                dDscRegiao = 19.96
            ElseIf ilCodRep = 600 Or ilCodRep = 800 Or ilCodRep = 1001 Or ilCodRep = 2100 Or ilCodRep = 6000 Then
                dDscRegiao = 17.2
            ElseIf ilCodRep = 905 Or ilCodRep = 5000 Or ilCodRep = 5001 Then
                dDscRegiao = 17.89
'            Else
'                dDscRegiao = 19.96
            End If
                
        End If

    End If

'    If ilCodCnd = 1 Or ilCodCnd = 12 Then 'A vista ou 14 dias
'        dDscRegiao = IIf(Trim(d5.Texto) = "", 0, Trim(d5.Texto))
'        bAVista = True
'    Else 'A Prazo
'        dDscRegiao = IIf(Trim(d1.Texto) = "", 0, Trim(d1.Texto))
'        bAVista = False
'    End If

 
End Sub

Private Sub Form_Click()
    
    DoEvents
    
End Sub

Private Sub Form_DblClick()
    
    DoEvents
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

    '*****************************************************************************
    'No momento em que uma tecla for pressionada, o programa avalia qual controle
    'está ativo. Se for o grid onde os produtos são inseridos, se a tecla
    'pressionada for ENTER executa-se a rotina de duplo-clique do controle. Se o
    'ativo não for o grid, a rotina EventoEnter vai evitar que ENTER provoque
    'efeitos indesejáveis no campo "Observações".
    '*****************************************************************************
    
    If Me.ActiveControl.Name = "cboCli" Or Me.ActiveControl.Name = "CboCondPag" Then
    
        Call EventoEnter(KeyAscii)
        
    End If
    
    
     
End Sub

Private Sub Form_Load()

    On Error GoTo TratarErro
    
    ControleLostFocus = True
    auxChange = False
    ControleAtualizaGrid = False
    
    CarregaTemporaria
    
    ilCodRep = Trim(sgRepresentante)
    
    If bgConsultaPed = False Then
        Activate_Ped
    End If
    
    'TUBOS E CONEXÕES COLETOR ESGOTO OCRE
    CarregaGrid "463,531,530,532,591,245,246,247,593,403,1093,439,633,1207,508,442,273,444,325,211,335,495,435,516,515,221,513,251,433,324,252,501,509,253,285,511,470,510,499,494", tubo_tubos_conexoes_coletor_esgoto
    ConfiguraFlexGrid tubo_tubos_conexoes_coletor_esgoto
    DefineClassificacao tubo_tubos_conexoes_coletor_esgoto
    
    'TUBOS E CONEXÕES DEFOFO
    CarregaGrid "526,527,528,529,618,617,512,248,249,523", tubo_conexoes_defofo
    ConfiguraFlexGrid tubo_conexoes_defofo
    DefineClassificacao tubo_conexoes_defofo
    
    'TUBOS E CONEXÕES PBA
    CarregaGrid "637,734,735,641,687,688,1094,226,399,398,397,228,256,386,388,383,384,195,197,198,202,203,286,213,118,237,119,385,172,157,229,361,402,401,156,360,357,359,358,354,212,257,374,79,99,209,209,362,155", tubo_conexoes_pba
    ConfiguraFlexGrid tubo_conexoes_pba
    DefineClassificacao tubo_conexoes_pba
    
    'TUBOS E CONEXÕES PREDIAL
    CarregaGrid "11,7,8,9,223,572,217,768,621,3418,62,61,60,59,574,601,45,44,39,38,101,116,108,341,78,77,76,75,107,216,469,80,81,92,82,96,97,98,272,46,47,64,48,63,65,49,66,102,128,127,126,114,129,130,222,255,306,606,443,607,446,445,461,310,368,601,116,608,609,610,611,612,471,3449,3448,334,600,598,599,3310,193,192,194,352,759", tubos_conexoes_predial
    ConfiguraFlexGrid tubos_conexoes_predial
    DefineClassificacao tubos_conexoes_predial
    
    'TUBOS E CONEXÕES ROSCÁVEIS
    CarregaGrid "12,13,67,68,14,103,171,159,160,173,174,177,178", tubos_conexoes_roscaveis
    ConfiguraFlexGrid tubos_conexoes_roscaveis
    DefineClassificacao tubos_conexoes_roscaveis
    
    'TUBOS E CONEXÕES IRRIGAÇÃO AZUIS
    CarregaGrid "219,220,259,268,260,261,262,264,265,266,294,295,296,289,290,298,415,297,292,293,331,219,314,328,326,327,705,706,707,301,302,317,299,300,307,666,604,715,717,718,692,643,719,291", tubos_conexoes_irri_azuis
    ConfiguraFlexGrid tubos_conexoes_irri_azuis
    DefineClassificacao tubos_conexoes_irri_azuis
    
    'TUBOS E CONEXÕES ÁGUA
    CarregaGrid "1,2,3,4,5,6,89,93,94,110,100,490,57,58,186,122,123,722,224,488,485,484,504,503,699,700,502,492,33,34,35,43,36,111,106,113,83,84,85,86,87,88,188,214,184,37,41,50,40,42,109,115,121,483,481,482,696,697,698,708,227,460,720,313,721,3459,3460,3664,3462,3461,3463,496,497,475,459,468,465,467,440,474,709,710,711,712,713,714,726,727,728,3462", tubos_conexoes_agua
    ConfiguraFlexGrid tubos_conexoes_agua
    DefineClassificacao tubos_conexoes_agua
    '*****************************************************************************
    'Posiciona o formulário no canto superior esquerdo do MDI e define suas
    'medidas, de modo que a área com os cálculos liberadores do pedido fiquem
    'ocultas.
    '*****************************************************************************
    
    Me.Left = 0
    Me.Top = 0
    Me.Height = 9030
    Me.Width = 12465

    Status.Visible = False

    '*************************************************************************************
    'Se o usuário tiver perfil de administrador, poderá ter acesso aos índices.
    '*************************************************************************************
    
    If sgFlgUsu = "L" And APLICA = 0 Then
        FrmConhecimento.BorderStyle = 2
    End If
    
    iDscRegiao = 0 'True
    
    '*************************************************************************************
    'Carrega campos e formata o ambiente.
    '*************************************************************************************
    
    SSTConhec.TabVisible(3) = False

    blCarregou = False
    blLG = False
    blRetornoDupls = False
    
    Set cboCli.Conexao = Conexao
    Set CboCondPag.Conexao = Conexao
    
    Datped = CDate(Date)
   
    'MskNumNf.TipodeDados numero
    'MskSerie.TipodeDados numero
    'MskVlrUnit.TipodeDados numero
    'MskDatEmiNf.TipodeDados numero
    MskMargem.TipodeDados numero
    d6.TipodeDados Literal
    d7.TipodeDados Literal
    SSTConhec.TabEnabled(1) = False
    SSTConhec.TabEnabled(2) = False
    FraGrupo.Visible = False
   
    blEbahia = False
    slPedSimples = ""
    slFlgAlt = ""
   
    GrdNotaCliente.TextMatrix(0, 0) = "Cód."
    GrdNotaCliente.ColWidth(0) = 550
    GrdNotaCliente.TextMatrix(0, 1) = "Descrição do Produto"
    GrdNotaCliente.ColWidth(1) = 4830
    GrdNotaCliente.TextMatrix(0, 2) = "Emb."
    GrdNotaCliente.ColWidth(2) = 500
    GrdNotaCliente.TextMatrix(0, 3) = "Qtde"
    GrdNotaCliente.ColWidth(3) = 600
    GrdNotaCliente.TextMatrix(0, 4) = "Vlr Unitário"
    GrdNotaCliente.ColWidth(4) = 1100
    GrdNotaCliente.TextMatrix(0, 5) = ""
    GrdNotaCliente.ColWidth(5) = 300
    GrdNotaCliente.TextMatrix(0, 6) = "%Desc"
    GrdNotaCliente.ColWidth(6) = 660
    'GrdNotaCliente.TextMatrix(0, 7) = "Vlr do Item"
    'GrdNotaCliente.ColWidth(7) = 1100
    GrdNotaCliente.TextMatrix(0, 7) = "Vlr c/Desc."
    GrdNotaCliente.ColWidth(7) = 1300
    GrdNotaCliente.TextMatrix(0, 8) = "Grupo"
    GrdNotaCliente.ColWidth(8) = 1
    GrdNotaCliente.TextMatrix(0, 9) = "Vlr Bruto"
    GrdNotaCliente.ColWidth(9) = 1
    GrdNotaCliente.TextMatrix(0, 10) = "Peso"
    GrdNotaCliente.ColWidth(10) = 1
    GrdNotaCliente.TextMatrix(0, 11) = "Vlr Ideal"
    GrdNotaCliente.ColWidth(11) = 1
    GrdNotaCliente.TextMatrix(0, 12) = "Vlr Tabela N"
    GrdNotaCliente.ColWidth(12) = 1
    GrdNotaCliente.TextMatrix(0, 13) = "% Margem"
    GrdNotaCliente.ColWidth(13) = 1
    GrdNotaCliente.TextMatrix(0, 14) = "Irriga"
    GrdNotaCliente.ColWidth(14) = 1
    GrdNotaCliente.TextMatrix(0, 15) = "CusUntQtd"
    GrdNotaCliente.ColWidth(15) = 1
    GrdNotaCliente.TextMatrix(0, 16) = "CusAdiQtd"
    GrdNotaCliente.ColWidth(16) = 1
    GrdNotaCliente.TextMatrix(0, 17) = "AlqImpFed"
    GrdNotaCliente.ColWidth(17) = 1
    GrdNotaCliente.TextMatrix(0, 18) = "Vlr do Item"
    GrdNotaCliente.ColWidth(18) = 1100
    If APLICA = 1 Then
        GrdNotaCliente.TextMatrix(0, 19) = ""
    Else
        GrdNotaCliente.TextMatrix(0, 19) = "Margem"
    End If
    
    GrdNotaCliente.ColWidth(19) = 800
        
    GrdIndice.TextMatrix(0, 0) = "Cod"
    GrdIndice.ColWidth(0) = 450
    GrdIndice.TextMatrix(0, 1) = "Descrição do Produto"
    GrdIndice.ColWidth(1) = 3040
    GrdIndice.TextMatrix(0, 2) = "Fixo"
    GrdIndice.ColWidth(2) = 500
    GrdIndice.TextMatrix(0, 3) = "Fin."
    GrdIndice.ColWidth(3) = 470
    GrdIndice.TextMatrix(0, 4) = "Comi."
    GrdIndice.ColWidth(4) = 530
    GrdIndice.TextMatrix(0, 5) = "Frete"
    GrdIndice.ColWidth(5) = 500
    GrdIndice.TextMatrix(0, 6) = "PDD"
    GrdIndice.ColWidth(6) = 500
    GrdIndice.TextMatrix(0, 7) = "Marg."
    GrdIndice.ColWidth(7) = 500
    GrdIndice.TextMatrix(0, 8) = "ICM"
    GrdIndice.ColWidth(8) = 500
    GrdIndice.TextMatrix(0, 9) = "V.Unt"
    GrdIndice.ColWidth(9) = 520
    GrdIndice.TextMatrix(0, 10) = "IDX"
    GrdIndice.ColWidth(10) = 500
    GrdIndice.TextMatrix(0, 11) = "Vlr Mínimo"
    GrdIndice.ColWidth(11) = 920
    GrdIndice.TextMatrix(0, 12) = "Vlr Pedido"
    GrdIndice.ColWidth(12) = 900
    GrdIndice.TextMatrix(0, 13) = "Vlr Prev."
    GrdIndice.ColWidth(13) = 900
    GrdIndice.TextMatrix(0, 14) = "Peso"
    GrdIndice.ColWidth(14) = 750
    GrdIndice.TextMatrix(0, 15) = "Marg(%)"
    GrdIndice.ColWidth(15) = 720
    GrdIndice.TextMatrix(0, 16) = "Vlr Margem"
    GrdIndice.ColWidth(16) = 1
    GrdIndice.TextMatrix(0, 17) = "% Custo"
    GrdIndice.ColWidth(17) = 1
    
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
    GrdGrupo.TextMatrix(0, 0) = "Grupo"
    GrdGrupo.ColWidth(0) = 6000
    GrdGrupo.TextMatrix(0, 1) = ""
    GrdGrupo.ColWidth(1) = 1
    GrdProduto.TextMatrix(0, 0) = ""
    GrdProduto.ColWidth(0) = 150
    GrdProduto.TextMatrix(0, 1) = "Cód."
    GrdProduto.ColWidth(1) = 350
    GrdProduto.TextMatrix(0, 2) = "Descrição do Produto"
    GrdProduto.ColWidth(2) = 3600
    GrdProduto.TextMatrix(0, 3) = "Emb."
    GrdProduto.ColWidth(3) = 400
    GrdProduto.TextMatrix(0, 4) = "Preço Unit."
    GrdProduto.ColWidth(4) = 1100
    GrdProduto.TextMatrix(0, 5) = "  "
    GrdProduto.ColWidth(5) = 1100
    GrdProduto.TextMatrix(0, 6) = "  "
    GrdProduto.ColWidth(6) = 1100
    GrdEntrega.TextMatrix(0, 0) = "Cód."
    GrdEntrega.ColWidth(0) = 500
    GrdEntrega.TextMatrix(0, 1) = "Descrição do Produto"
    GrdEntrega.ColWidth(1) = 3300
    GrdEntrega.TextMatrix(0, 2) = "Qtd Ped."
    GrdEntrega.ColWidth(2) = 750
    GrdEntrega.TextMatrix(0, 3) = "Qtd Fat."
    GrdEntrega.ColWidth(3) = 700
    GrdEntrega.TextMatrix(0, 4) = "Tot.Entreg."
    GrdEntrega.ColWidth(4) = 1000
    GrdEntrega.TextMatrix(0, 5) = "N.Fiscal"
    GrdEntrega.ColWidth(5) = 700
    GrdEntrega.TextMatrix(0, 6) = "Dt.Emissão"
    GrdEntrega.ColWidth(6) = 980
    GrdEntrega.TextMatrix(0, 7) = ""
    GrdEntrega.ColWidth(7) = 250
    GrdEntrega.TextMatrix(0, 8) = "Ped.Saldo"
    GrdEntrega.ColWidth(8) = 900
    GrdEntrega.TextMatrix(0, 9) = "Qtde "
    GrdEntrega.ColWidth(9) = 600
    GrdEntrega.TextMatrix(0, 10) = "Qtd Fat."
    GrdEntrega.ColWidth(10) = 700
    GrdEntrega.TextMatrix(0, 11) = "N.Fiscal"
    GrdEntrega.ColWidth(11) = 700
    GrdEntrega.TextMatrix(0, 12) = "Dt.Emissão"
    GrdEntrega.ColWidth(12) = 1000
    ''MskNumNf.Enabled = False
    'MskSerie.Enabled = False
    ''Bto_Aplica.Enabled = False
    'MskVlrUnit.Enabled = False
    'MskDatEmiNf.Enabled = False
    ''BtoAdiNF.Enabled = False
    ''BtoExcNF.Enabled = False
    ''BtoLimpaNF.Enabled = False
    BtoGrava.Enabled = False
    GrdNotaCliente.Enabled = False
    Opt_FOB.Enabled = False
    Opt_CIF.Enabled = False
    LblSimBahia.Visible = False
    ChkKit.Enabled = True
    ChkKit.Value = 0
    
    '*************************************************************************************
    'Zera variáveis de cálculos.
    '*************************************************************************************
   
    slremet = ""
    ilCodCnd = 0
    ilQtdPar = 0
    dlValUntN = 0
    dlValUntA = 0
    dlValUntB = 0
    dlValItem = 0
    dlPerCusFin = 0
    dlPerDesCnd = 0
    ilQtdParCnd = 0
    ilPrzMed = 0
    ilNumTab = 0
    ilind = 0
    Linhas = 0
    dlPerDesRep = 0
    ilIdeGrp = 0
    dlPerDesPadrao = 0
    dlPerComiNeg = 0
    slClasCor = ""
    blleitura = False
    blimpa = False
    
    LblTextoideal.Visible = False
    vl3.Visible = False
    LblIdeal.Visible = False
    lblpercideal.Visible = False
    d9.Visible = False
    TxtTransp.Text = "UNOCANN TRANSPORTES LTDA"

    sgFlagOper = "I"
    slIrriga = ""
    QtdTubo = 0
    QtdAspe = 0
    QtdConx = 0
    ilDescChave = 0
    slChave = ""
   
    Opt_CIF.Value = True
    TxtTransp.Text = "UNOCANN TRANSPORTES LTDA"
   
    If bgConsultaPed = True Then
        MskNroPedido.Text = igNroPed
    End If
   
    Me.Show
   
    Exit Sub

TratarErro:

    Rotina_Erro "Form_Load"
    
End Sub

Private Sub GrdGrupo_Click()

    If GrdGrupo.RowSel = 0 Then
        Exit Sub
    End If
    
    ilGrupo = GrdGrupo.TextMatrix(GrdGrupo.RowSel, 1)
    
    CarregaGridProduto
    
End Sub

Private Sub GrdGrupo_DblClick()

    If GrdGrupo.RowSel = 0 Then
        Exit Sub
    End If
    
    ilGrupo = GrdGrupo.TextMatrix(GrdGrupo.RowSel, 1)
    
    CarregaGridProduto
    
End Sub

Private Sub GrdGrupo_SelChange()

    If GrdGrupo.RowSel = 0 Then
        Exit Sub
    End If
    
    ilGrupo = GrdGrupo.TextMatrix(GrdGrupo.RowSel, 1)
    
    CarregaGridProduto
    
End Sub

Private Sub GrdNotaCliente_DblClick()
    
    If bgBloqPed = True Then
        Exit Sub
    End If

    If GrdNotaCliente.RowSel = 0 Then
        Exit Sub
    End If
    
    If GrdNotaCliente.SelectionMode = flexSelectionByRow Then
        GrdNotaCliente.SelectionMode = flexSelectionFree
        GrdNotaCliente.Refresh
        GrdNotaCliente_Click
        'tubos_conexoes_irri_azuis_SelChange
        Exit Sub
    End If
    ilCelula = GrdNotaCliente.RowSel
    
    MskNumNf = GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 0)
    Me.Refresh
    LblRotaRec.Caption = GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 1)
    MskSerie = GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 3)
    MskVlrUnit = GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 4)
    
    If Trim(GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 5)) = "" Then
    
        ilNumTab = 0
        'LblUnit.Caption = "Valor Unitário"
    
    Else
    
        If Trim(GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 5)) = "A" Then
            ilNumTab = 1
            'LblUnit.Caption = "Valor Unitário - A"
        Else
            ilNumTab = 2
            'LblUnit.Caption = "Valor Unitário - B"
        End If
        
    End If
    
    MskDatEmiNf = GrdNotaCliente.TextMatrix(GrdNotaCliente.RowSel, 6)
    ''MskNumNf.Enabled = False
    'BtoAdiNF.Enabled = True
    'BtoAdiNF.Caption = "&Alterar"
    'BtoExcNF.Enabled = True
    
    blModificar = True
    
    'BtoProduto.Enabled = False
    'MskSerie.Enabled = True
    'MskSerie.SetFocus
    
    ''VSValUnit.Value = ilNumTab
    
    DoEvents
    
End Sub

Private Sub GrdNotaCliente_RowColChange()

    'If MSFlexGrid1.Col > 5 Then
        'MSFlexGrid1.ScrollBars = flexScrollBarBoth
    'Else
        'MSFlexGrid1.ScrollBars = flexScrollBarNone
    'End If
    
End Sub

Private Sub GrdProduto_DblClick()

    If bgBloqPed = True Then
        Exit Sub
    End If

    'If MskNumNf.Enabled = True Then
        
        'MskNumNf.Texto = GrdProduto.TextMatrix(GrdProduto.RowSel, 1)
        'FraGrupo.Visible = False
        
        'MskNumNf_LostFocus
        'DoEvents
        
        'MskSerie.SetFocus
        
    'End If
    
End Sub

Private Sub Label24_Click()
    
    DoEvents
    
End Sub

Private Sub Label24_DblClick()
    
    DoEvents
    
End Sub

Private Sub lblNegocio_Click()
    
    DoEvents
    
End Sub

Private Sub lblNegocio_DblClick()
    
    DoEvents
    
End Sub

Private Sub MskDatEmiNf_GotFocus()
    
    If ControleLostFocus = False Then
        Exit Sub
    End If
    
    'If Trim(MskSerie.Texto) = "" Or Trim(MskSerie.Texto) = 0 Then
        
        'MsgBox "Informe o Quantidade do produto", vbExclamation + vbOKOnly, "Atenção!"
        
        'MskSerie.SetFocus
        
        'Exit Sub
    
    'End If

    If LeituraCliente = False Then
        LimpaGeral
    End If

    Call SelecionaTudo

End Sub

Private Sub MskDatEmiNf_LostFocus()

    If Trim(MskDatEmiNf) = "" Then
        MskDatEmiNf = 0
    End If

End Sub

Private Sub MskNroPedido_GotFocus()
    
    blLG = False
    
    Call SelecionaTudo

End Sub

Private Sub MskNroPedido_LostFocus()

    If ControleLostFocus = False Then
        Exit Sub
    End If
    
    Dim i As Integer
    
    'If Me.ActiveControl.Name = "BtoSair" Or Me.ActiveControl.Name = "BtoLimpaCTRC" Or Me.ActiveControl.Name = "BtoProduto" Or blLG = True Or blRetornoDupls = True Then
        'Exit Sub
   ' End If
  
    If bgBloqPed = True And blCarregou = True Then
        Exit Sub
    End If
  
    If bgPedMKT = True And bgSimula = False Then
        
        cboCli.Criterio = igCodCli
        
        ilCodCli = igCodCli
        
        If LeituraCliente = False Then
            
            LimpaGeral
            
            Exit Sub
            
        End If
        
        cboCli.Habilitado = False
        CboCondPag.Habilitado = True
        CboCondPag.SetFocus
        
        Exit Sub
        
    End If
  
    blCarregou = True
  
    If Trim(MskNroPedido.Text) = "" Or Trim(MskNroPedido.Text) = 0 Then
        
        CmdCancelar.Enabled = False
        CmdImpr.Enabled = False
        
        bgBloqPed = False
        
        LblBloqPed.Caption = "Este pedido não pode ser alterado"
        LblBloqPed.Visible = False
        
        If bgSimula = False Then
        
            If Numero_Ped = False Then
                Exit Sub
            End If
            
        End If
    
    Else
    
        '*********************************************************************************
        'A função CarregaTela expõe os dados do pedido selecionado. Se houver alguma falha
        'no carregamento seu valor será False, e a função LimpaGeral retorna o ambiente ao
        'seu estado original.
        '*********************************************************************************
     
        If CarregaTela = False Then
            
            LimpaGeral
            
            Exit Sub
            
        End If
     
        SSTConhec.TabVisible(3) = True
        
        If sgFlagOper = "A" Then
            
            Opt_FOB.Enabled = True
            Opt_CIF.Enabled = True
            
            '*****************************************************************************
            'Se o cliente se responsabilizou pela entrega do pedido, não deve ser cobrado
            'frete.
            '*****************************************************************************
            
            If Opt_FOB.Value = True Then
                dlPerDesFOBReal = dlPerDesFOB
            Else
                dlPerDesFOBReal = 0
            End If
            
            ilCodCli = cboCli.codigo
            
            CboCondPag.Habilitado = True
            
            blleitura = False
            
            SSTConhec.TabEnabled(1) = True
            SSTConhec.TabEnabled(2) = True
            
            '*****************************************************************************
            '
            '*****************************************************************************
            
            If LeituraCliente = False Then
                
                LimpaGeral
                
                Exit Sub
                
            End If
            
            If LeituraCondicao = False Then
                
                LimpaGeral
                
                Exit Sub
                
            End If
            
            If CalculaIndice = False Then
                
                LimpaGeral
                
                Exit Sub
                
            End If
            
            GrdNotaCliente.col = 19
                    
            For i = 1 To GrdIndice.rows - 1
            
                GrdNotaCliente.row = i
                
                If APLICA = 1 Then
                
                Else
        
                    If GrdIndice.TextMatrix(i, 15) < 10 Then
                        GrdNotaCliente.CellBackColor = &HFF&
                    ElseIf GrdIndice.TextMatrix(i, 15) >= 10 And GrdIndice.TextMatrix(i, 15) < 15 Then
                        GrdNotaCliente.CellBackColor = &HFF00&
                    ElseIf GrdIndice.TextMatrix(i, 15) >= 15 Then
                        GrdNotaCliente.CellBackColor = &HFF0000
                    End If
                
                End If
                
            Next
            
            CmdImpr.Enabled = True
            ''MskNumNf.Enabled = True
            
            LimpaLinhaNF
            
            ''MskNumNf.Enabled = True
            ''BtoProduto.Enabled = True
            'MskSerie.Enabled = True
            ''Bto_Aplica.Enabled = True
            ''BtoExcNF.Enabled = False
            ''BtoAdiNF.Enabled = True
            'MskNumNf.SetFocus
            blModificar = False
            ChkKit.Enabled = False
            BtoGrava.Enabled = True
            GrdNotaCliente.Enabled = True
            
            ilNumTab = 0
            
            'LblResultNegocio.Caption = ""
            
        End If
    
    End If
      
    If bgBloqPed = True Then
        
        ''MskNumNf.Enabled = False
        ''MskNumNf.Enabled = False
        ''BtoProduto.Enabled = False
        'MskSerie.Enabled = False
        ''Bto_Aplica.Enabled = False
        ''BtoExcNF.Enabled = False
        ''BtoAdiNF.Enabled = False
        blModificar = False
        ChkKit.Enabled = False
        BtoGrava.Enabled = False
        'GrdNotaCliente.Enabled = False
        
        If blImpr = False Then
            CmdImpr.Enabled = False
        End If
        
        cboCli.Habilitado = False
        CboCondPag.Habilitado = False
        Opt_CIF.Enabled = False
        Opt_FOB.Enabled = False
        TxtTransp.Enabled = False
        'MskDatEmiNf.Enabled = False
        BtoLimpaCTRC.Enabled = False
    
    End If
  
    If blVencidos = True And sgFlagOper <> "A" And blRetornoDupls = False Then
    
        MsgBox "Existe(m) Título(s) vencido(s) para este cliente!", vbExclamation + vbOKOnly, "Atenção!"
        
        blVencidos = False
        blRetornoDupls = True
        
        'SSTConhec.Tab = 2
        
    End If
  
    If bgSimula = True Then
        CboUFCli.SetFocus
    End If
  
End Sub

Private Sub MskNumNf_GotFocus()
    
    If ControleLostFocus = False Then
        Exit Sub
    End If
    
    If Me.ActiveControl.Name = "BtoSair" Or Me.ActiveControl.Name = "BtoLimpaCTRC" Or blLG = True Then
        Exit Sub
    End If
    
    'If Trim(CboCondPag.Criterio) = "" Then
        
        'MsgBox "Informe a Condição de pagamento", vbExclamation + vbOKOnly, "Atenção!"
        
        'CboCondPag.SetFocus
        
        'Exit Sub
        
    'End If
  
    If (Trim(MskNroPedido.Text) = "" Or Trim(MskNroPedido.Text) = 0) And bgPedMKT = False Then
     
        MsgBox "Informe o Número do Pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        'MskNroPedido.SetFocus
        
        Exit Sub
        
    End If

   ' If (Trim(MskNroPedido.Text) < Seqini Or Trim(MskNroPedido.Text) > SeqFim) And APLICA = 1 And sgFlagOper <> "A" Then
    Dim x As Double
    x = Val(Mid(MskNroPedido.Text, 2, 5))
    If x < Val(Mid(Seqini, 2, 5)) Or x > Val(Mid(SeqFim, 2, 5)) And APLICA = 1 And sgFlagOper <> "A" Then
     
     
     
        MsgBox "Número do pedido fora do intervalo permitido, contate o administrador do sistema", vbExclamation + vbOKOnly, "Atenção!"
        
        'MskNroPedido.SetFocus
        
        Exit Sub
        
    End If
  
    If Trim(slremet) = "" Or Trim(cboCli.codigo) = "" And bgSimula = False Then
    
        Exit Sub
    
    Else
    
        If LeituraCliente = False Then
            LimpaGeral
        End If
        
    End If

    Call SelecionaTudo
    
    'BtoProduto.Enabled = True
    
    'blModificar = False
    
    'CboCondPag.Habilitado = False
    'cboCli.Habilitado = False
    MskNroPedido.Enabled = False
    'BtoAdiNF.Enabled = True
    'BtoAdiNF.Caption = "&Adicionar"
    'BtoLimpaNF.Enabled = True
    blimpa = True
    'VSValUnit.Value = 0
    'MskDatEmiNf.Limpar
    
End Sub

Private Sub MskNumNf_LostFocus()

    If Me.ActiveControl.Name = "Opt_CIF" Or Me.ActiveControl.Name = "Opt_FOB" Or blLG = True Or blModificar = True Then
        Exit Sub
    End If
  
    If Me.ActiveControl.Name = "BtoSair" Or Me.ActiveControl.Name = "BtoLimpaCTRC" Or blLG = True Then
        Exit Sub
    End If

    'If Trim(CboCondPag.Criterio) = "" Then
        
        'MsgBox "Informe a Condição de pagamento", vbExclamation + vbOKOnly, "Atenção!"
        
        'CboCondPag.SetFocus
        
        'Exit Sub
    
    'End If
  
    If (Trim(MskNroPedido.Text) = "" Or Trim(MskNroPedido.Text) = 0) And bgPedMKT = False Then
        
        MsgBox "Informe o Número do Pedido", vbExclamation + vbOKOnly, "Atenção!"
        
        'MskNroPedido.SetFocus
        
        Exit Sub
        
    End If
  
    If (Trim(MskNroPedido.Text) < Seqini Or Trim(MskNroPedido.Text) > SeqFim) And APLICA = 1 And sgFlagOper <> "A" Then
        
        MsgBox "Número do pedido fora do intervalo permitido, contate o administrador do sistema", vbExclamation + vbOKOnly, "Atenção!"
        
        'MskNroPedido.SetFocus
        
        Exit Sub
        
    End If
  
    blModificar = False
    
    'CboCondPag.Habilitado = False
    'cboCli.Habilitado = False
    MskNroPedido.Enabled = False
    'BtoAdiNF.Enabled = True
    'BtoAdiNF.Caption = "&Adicionar"
    'BtoLimpaNF.Enabled = True
    blimpa = True
    'VSValUnit.Value = 0
    'MskDatEmiNf.Limpar
  
    'If Trim(slremet) = "" Or Trim(cboCli.Codigo) = "" Then
        
        'Exit Sub
        
    'Else
    
        'If LeituraCliente = False Then
            'LimpaGeral
        'End If
    
    'End If

End Sub

Private Sub MskSerie_GotFocus()

    On Error Resume Next
    
    If ControleLostFocus = False Then
        Exit Sub
    End If
    
    Dim Linhas As Integer
    
    DoEvents

    If Me.ActiveControl.Name = "BtoSair" Or Me.ActiveControl.Name = "BtoLimpaCTRC" Or Me.ActiveControl.Name = "BtoLimpaNF" Or blLG = True Or bgBloqPed = True Or MskNroPedido.Text = "" Then
        Exit Sub
    End If

    If Trim(MskNumNf) = "" Or Trim(MskNumNf) = "" Then
        Exit Sub
    End If
    
    sgQuery = "SELECT a.*, b.* from PRODUTO a, PRECO_PRODUTO b"
    sgQuery = sgQuery + "   WHERE a.flgsitu = 'N'"
    sgQuery = sgQuery + "     and a.Codprd = " & Trim(MskNumNf)
    sgQuery = sgQuery + "     and a.codprd = b.codprd"
    sgQuery = sgQuery + "     and b.datativ = (select max(datativ) from preco_produto"
    sgQuery = sgQuery + "                       Where Codprd = " & Trim(MskNumNf)
    sgQuery = sgQuery + "                         and datativ <= convert(datetime,'" & Trim(Datped) & "',103))"
  
    Call Consulta(sgQuery)
    
    If Rs.EOF Then
    
        MsgBox "Produto inexistente ou fora de linha", vbExclamation + vbOKOnly, "Atenção!"
        
        Rs.Close
        
        Set Rs = Nothing
        
        'MskNumNf.SetFocus
        
        Exit Sub
        
    End If
  
    LblRotaRec.Caption = IIf(IsNull(Rs!DSCPRD), "", Rs!DSCPRD)
    
    ilIdeGrp = IIf(Trim(Rs!IdeGrp) = "", 0, Trim(Rs!IdeGrp))
    dlPesUnt = IIf(Trim(Rs!PesUnt) = "", 0, Trim(Rs!PesUnt))
    dlValUntN = IIf(Trim(Rs!ValUntN) = "", 0, Trim(Rs!ValUntN))
    dlValUntA = IIf(Trim(Rs!ValUntA) = "", 0, Trim(Rs!ValUntA))
    dlValUntB = IIf(Trim(Rs!ValUntB) = "", 0, Trim(Rs!ValUntB))
    dlMrgPrd = IIf(Trim(Rs!MrgPrd) = "", 0, Trim(Rs!MrgPrd))
    dlValCusUntQtd = IIf(Trim(Rs!valcusuntqtd) = "", 0, Trim(Rs!valcusuntqtd))
    dlValCusAdicQtd = IIf(Trim(Rs!valcusadicqtd) = "", 0, Trim(Rs!valcusadicqtd))
    dlAlqImpFed = IIf(Trim(Rs!AlqImpFed) = "", 0, Trim(Rs!AlqImpFed))
    ilQtdEmb = IIf(Trim(Rs!QtdEmb) = "", 1, Trim(Rs!QtdEmb))
    ilFlgKit = Trim(Rs!FlgKit)
    
    Rs.Close
    
    Set Rs = Nothing
  
    'Acha desconto promocional destacado para o produto ou grupo (por representante)
    
    'dlPerDesPrd = 0
    
    'sgQuery = "select PerDsc from Desconto_promocional where CodRep = " & Trim(ilCodRep)
    'sgQuery = sgQuery + " and IdeGrp = " & ilIdeGrp
    'sgQuery = sgQuery + " and Codprd = " & Trim(MskNumNf.Texto)
    
    'Call consulta(sgQuery)
    
    'If Not Rs.EOF Then
        
        'dlPerDesPrd = IIf(Trim(Rs!PerDsc) = "", 0, Trim(Rs!PerDsc))
    
    'Else
        
        'Rs.Close
        
        'Set Rs = Nothing
        
        'sgQuery = "select PerDsc from Desconto_promocional where CodRep = " & Trim(ilCodRep)
        'sgQuery = sgQuery + " and IdeGrp = " & ilIdeGrp
        'sgQuery = sgQuery + " and Codprd is null "
        
        'Call consulta(sgQuery)
        
        'If Not Rs.EOF Then
            'dlPerDesPrd = IIf(Trim(Rs!PerDsc) = "", 0, Trim(Rs!PerDsc))
        'End If
    
    'End If
    
    'If dlPerDesPrd = 0 Then
        dlPerDesPrd = dlPerDesRep
    'End If
    
    'Rs.Close
    
    'Set Rs = Nothing
  
    CalculaDesconto
    Call SelecionaTudo
    
End Sub

Private Sub Opt_CIF_Click()

    ''MskNumNf.Enabled = True
    dlPerDesFOBReal = 0
    TxtTransp.Text = "UNOCANN TRANSPORTES LTDA"
    
    CalculaDesconto
    
    If CalculaIndice = False Then
        LimpaGeral
    End If
    
    'MskNumNf.SetFocus

End Sub

Private Sub Opt_FOB_Click()
    
    ''MskNumNf.Enabled = True
    dlPerDesFOBReal = dlPerDesFOB
    TxtTransp.Text = "O PROPRIO"
    
    CalculaDesconto
    
    If CalculaIndice = False Then
        LimpaGeral
    End If
    
    'MskNumNf.SetFocus

End Sub

Private Sub SSTConhec_Click(PreviousTab As Integer)

    DoEvents
    
    If slFlgAlt = "N" And SSTConhec.Tab = 1 Then
        
        slFlgAlt = "O"
        
        sgQuery = "Update PEDIDO set flgalt = 'O'"
        sgQuery = sgQuery & " where nroped = '" & Trim(MskNroPedido.Text) & "'"
        
        Set Rs = Conexao.Execute(sgQuery)
        Set Rs = Nothing
        
    End If
    
End Sub

Private Sub SSTConhec_DblClick()
    
    DoEvents
    
    If slFlgAlt = "N" And SSTConhec.Tab = 1 Then
     
        slFlgAlt = "O"
        
        sgQuery = "Update PEDIDO set flgalt = 'O'"
        sgQuery = sgQuery & " where nroped = " & Trim(MskNroPedido.Text)
        
        Set Rs = Conexao.Execute(sgQuery)
        Set Rs = Nothing
    
    End If
    
End Sub

Private Sub SSTConhec_GotFocus()
    
    DoEvents
    
    If ControleLostFocus = False Then
        Exit Sub
    End If
    
    If slFlgAlt = "N" And SSTConhec.Tab = 1 Then
     
        slFlgAlt = "O"
        
        sgQuery = "Update PEDIDO set flgalt = 'O'"
        sgQuery = sgQuery & " where nroped = '" & Trim(MskNroPedido.Text) & "'"
        
        Set Rs = Conexao.Execute(sgQuery)
        Set Rs = Nothing
        
    End If

End Sub

Private Sub TxtChave_LostFocus()

    If Me.ActiveControl.Name = "BtoSair" Or Me.ActiveControl.Name = "BtoLimpaCTRC" Or blLG = True Then
        Exit Sub
    End If
    
    '*********************************************************************************
    'Em 22/09/2008 a validação a seguir me deu problemas com erros muito loucos. É bom
    'ficar de sobreaviso.
    '*********************************************************************************
    
    If Trim(TxtChave.Text) = "" Or Trim(TxtChave.Text) = 0 Then
        Exit Sub
    End If
    
    If Trim(slChave) = "" Then
        slChave = ""
        ilDescChave = 0
    End If
    
    If DecrypChave = True Then
        
        LblC.Caption = ilDescChave
        
        dlSumDscItem = ilDescChave
        
        LblC.Visible = True
    
    Else
    
        SSTConhec.Tab = 1
        TxtChave.SetFocus
        
    End If
        
End Sub

Private Sub TxtNegocio_Click()
    
    DoEvents
    
End Sub

Private Sub TxtNegocio_DblClick()
    
    DoEvents
    
End Sub

Private Sub TxtObserva_Click()
    
    DoEvents
    
End Sub

Private Sub TxtObserva_DblClick()
    
    DoEvents
    
End Sub

Private Sub TxtObserva_GotFocus()
    
    DoEvents
    
End Sub

Private Sub TxtPwdoper_GotFocus()
    
    Call SelecionaTudo
    
End Sub

Private Sub TxtPwdoper_LostFocus()
    
    'If Me.ActiveControl.Name = "BtoSair" Or _

        'Me.ActiveControl.Name = "BtoLimpaCTRC" Then
        'bgSenComi = True
        
        'Exit Sub
    
    'End If

    If Trim(TxtPwdoper.Text) <> "" Then
    
        sgQuery = "SELECT PWDCOMPARE('" & Trim(TxtPwdoper.Text) & "',SenUsu, 0) AS Senha_OK from usuario where codusu = " & LgCodUsuSis
        
        Consulta sgQuery
        
        If Rs("Senha_OK") = 1 Then
            
            bgSenOK = True
            bgSenComi = True
            
            lblmens.Visible = True
            
        Else
            
            lblmens.Visible = True
            'TxtPwdoper.SetFocus
            
            Exit Sub
        
        End If
    
    Else
    
        bgSenComi = True
    
    End If

End Sub

Private Sub TxtTransp_GotFocus()

    DoEvents
    
    Call SelecionaTudo
    
End Sub

Private Sub VSValUnit_Change()

    If blimpa = True Then
    
        'VSValUnit.Value = 0
        
        blimpa = False
        
        Exit Sub
        
    End If
 
    
End Sub


