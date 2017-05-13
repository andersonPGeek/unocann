VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Agenda 
   Caption         =   "Agenda de compromisso."
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Frm_Agenda.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text3 
      Height          =   2055
      Left            =   5400
      TabIndex        =   85
      Text            =   "Text3"
      Top             =   5640
      Width           =   5415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5400
      TabIndex        =   83
      Text            =   "Text1"
      Top             =   960
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5400
      TabIndex        =   81
      Text            =   "Text1"
      Top             =   360
      Width           =   5415
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   25
      Left            =   14880
      TabIndex        =   80
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   25
      Left            =   14520
      TabIndex        =   79
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   24
      Left            =   14880
      TabIndex        =   78
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   24
      Left            =   14520
      TabIndex        =   77
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   23
      Left            =   14880
      TabIndex        =   76
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   23
      Left            =   14520
      TabIndex        =   75
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   22
      Left            =   14880
      TabIndex        =   74
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   22
      Left            =   14520
      TabIndex        =   73
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   21
      Left            =   14880
      TabIndex        =   72
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   21
      Left            =   14520
      TabIndex        =   71
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   20
      Left            =   14880
      TabIndex        =   70
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   20
      Left            =   14520
      TabIndex        =   69
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   19
      Left            =   14880
      TabIndex        =   68
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   19
      Left            =   14520
      TabIndex        =   67
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   18
      Left            =   14880
      TabIndex        =   66
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   18
      Left            =   14520
      TabIndex        =   65
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   17
      Left            =   14880
      TabIndex        =   64
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   17
      Left            =   14520
      TabIndex        =   63
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   16
      Left            =   14880
      TabIndex        =   62
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   16
      Left            =   14520
      TabIndex        =   61
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   15
      Left            =   14880
      TabIndex        =   60
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   15
      Left            =   14520
      TabIndex        =   59
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   14
      Left            =   14880
      TabIndex        =   58
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   14
      Left            =   14520
      TabIndex        =   57
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   13
      Left            =   14880
      TabIndex        =   56
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   13
      Left            =   14520
      TabIndex        =   55
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   12
      Left            =   14880
      TabIndex        =   54
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   12
      Left            =   14520
      TabIndex        =   53
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   11
      Left            =   14880
      TabIndex        =   52
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   11
      Left            =   14520
      TabIndex        =   51
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   10
      Left            =   14880
      TabIndex        =   50
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   10
      Left            =   14520
      TabIndex        =   49
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   9
      Left            =   14880
      TabIndex        =   48
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   9
      Left            =   14520
      TabIndex        =   47
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   8
      Left            =   14880
      TabIndex        =   46
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   8
      Left            =   14520
      TabIndex        =   45
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   7
      Left            =   14880
      TabIndex        =   44
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   7
      Left            =   14520
      TabIndex        =   43
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   6
      Left            =   14880
      TabIndex        =   42
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   6
      Left            =   14520
      TabIndex        =   41
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   5
      Left            =   14880
      TabIndex        =   40
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   5
      Left            =   14520
      TabIndex        =   39
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   4
      Left            =   14880
      TabIndex        =   38
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   4
      Left            =   14520
      TabIndex        =   37
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   3
      Left            =   14880
      TabIndex        =   36
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   3
      Left            =   14520
      TabIndex        =   35
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   2
      Left            =   14880
      TabIndex        =   34
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   2
      Left            =   14520
      TabIndex        =   33
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   1
      Left            =   14880
      TabIndex        =   32
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   1
      Left            =   14520
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Codigo 
      Height          =   285
      Index           =   0
      Left            =   14880
      TabIndex        =   30
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Ano 
      Height          =   285
      Index           =   0
      Left            =   14520
      TabIndex        =   29
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   960
      TabIndex        =   28
      Top             =   7335
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   960
      TabIndex        =   27
      Top             =   7095
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   960
      TabIndex        =   26
      Top             =   6840
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   960
      TabIndex        =   25
      Top             =   6580
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   960
      TabIndex        =   24
      Top             =   6255
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   960
      TabIndex        =   23
      Top             =   6015
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   960
      TabIndex        =   22
      Top             =   5775
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   960
      TabIndex        =   21
      Top             =   5445
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   960
      TabIndex        =   20
      Top             =   5175
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   960
      TabIndex        =   19
      Top             =   4935
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   960
      TabIndex        =   18
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   960
      TabIndex        =   17
      Top             =   4400
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   960
      TabIndex        =   16
      Top             =   4095
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   960
      TabIndex        =   15
      Top             =   3855
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   960
      TabIndex        =   14
      Top             =   3615
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   960
      TabIndex        =   13
      Top             =   3340
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   960
      TabIndex        =   12
      Top             =   3020
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   8
      Left            =   960
      TabIndex        =   11
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   960
      TabIndex        =   10
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   960
      TabIndex        =   9
      Top             =   2240
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   960
      TabIndex        =   8
      Top             =   2000
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   7
      Top             =   1695
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   960
      TabIndex        =   6
      Top             =   1400
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   1170
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   855
      Width           =   3975
   End
   Begin VB.TextBox Txt_Compromisso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   3975
   End
   Begin VB.PictureBox Dta_Agenda 
      Height          =   315
      Left            =   1560
      ScaleHeight     =   255
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   190
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   8160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3975
      Left            =   5400
      TabIndex        =   87
      Top             =   1320
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7011
      _Version        =   393217
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Assunto"
      Height          =   255
      Left            =   5400
      TabIndex        =   86
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Contato"
      Height          =   255
      Left            =   5400
      TabIndex        =   84
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Cliente"
      Height          =   255
      Left            =   5400
      TabIndex        =   82
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Frm_Agenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rst_Agenda                  As rdoResultset
Dim nodTreeView                 As Node


Private Sub Dta_Agenda_Click()
    Sl_Desc = "select * "
    Sl_Desc = Sl_Desc & " from Tba_ClienteAgenda "
    Set Rst_Agenda = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
    While Rst_Agenda.EOF = False
          Ano(Rst_Agenda!Age_Hr_Codigo) = Rst_Agenda!Age_Ano
          Codigo(Rst_Agenda!Age_Hr_Codigo) = Rst_Agenda!Age_Codigo
          Txt_Compromisso(Rst_Agenda!Age_Hr_Codigo) = Rst_Agenda!Age_Assunto
          Rst_Agenda.MoveNext
    Wend
    Rst_Agenda.Close

End Sub

Private Sub Txt_Compromisso_Click(Index As Integer)
   ImageList1.ListImages.Clear

   TreeView1.LabelEdit = tvwManual ' Set property to manual.
   TreeView1.Nodes.Clear
   TreeView1.HotTracking = True
   ImageList1.ListImages.Clear
   
 ''TreeView1.Style = tvwTreelinesPlusMinusText
   Sl_Pai = "'0'"
    
   Set nodTreeView = TreeView1.Nodes.Add(, , Sl_Pai, "Histórico de contatos com a empresa", 0)
     
   Sl_Desc = "select * "
   Sl_Desc = Sl_Desc & " from Tba_Contatos       b,"
   Sl_Desc = Sl_Desc & "      Tba_ClienteAgenda  a"
   Sl_Desc = Sl_Desc & "  where b.Con_Cliente = a.Cli_Cliente"
   Sl_Desc = Sl_Desc & "    and b.Con_Codigo  = a.Con_Codigo"
   Set Rst_Agenda = Cn.OpenResultset(Sl_Desc, rdOpenDynamic, rdConcurRowVer)
   While Rst_Agenda.EOF = False
         Sl_Pai = "'0'"
         Sl_Filho = "'" & Format(Rst_Agenda!Age_Dt_Contato + 3, "dd-mm-yyyy") & "'"
     
         Set nodTreeView = TreeView1.Nodes.Add(Sl_Pai, tvwChild, Sl_Filho, Format(Rst_Agenda!Age_Dt_Contato, "dd-mm-yyyy") & "  -  " & Rst_Agenda!Con_Nome, 0)
          
         
         Rst_Agenda.MoveNext
   Wend
   Rst_Agenda.Close
   TreeView1.Nodes.Item(1).Expanded = True
   Exit Sub

End Sub

 
