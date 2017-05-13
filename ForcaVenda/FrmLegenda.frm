VERSION 5.00
Begin VB.Form FrmLegenda 
   BackColor       =   &H00400000&
   Caption         =   "Legenda"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Enabled         =   0   'False
      Height          =   2175
      Left            =   5040
      TabIndex        =   16
      Top             =   0
      Width           =   4695
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Left            =   600
         TabIndex        =   20
         Text            =   "A"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   19
         Text            =   "FrmLegenda.frx":0000
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   525
         Left            =   600
         TabIndex        =   22
         Text            =   "N"
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   21
         Text            =   "FrmLegenda.frx":0023
         Top             =   120
         Width           =   3255
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   555
         Left            =   960
         TabIndex        =   18
         Text            =   "S"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "FrmLegenda.frx":0044
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5535
      Begin VB.TextBox Text34 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Left            =   120
         TabIndex        =   15
         Text            =   "999999"
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox Text33 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1440
         TabIndex        =   14
         Text            =   "Pedido não enviado à Unocann"
         Top             =   120
         Width           =   2895
      End
      Begin VB.TextBox Text32 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   120
         TabIndex        =   13
         Text            =   "999999"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text31 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "FrmLegenda.frx":006B
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text30 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   120
         TabIndex        =   11
         Text            =   "999999"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1440
         TabIndex        =   10
         Text            =   "Pedido Faturado "
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   120
         TabIndex        =   9
         Text            =   "999999"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "FrmLegenda.frx":00A5
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Left            =   120
         TabIndex        =   7
         Text            =   "999999"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text25 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1440
         TabIndex        =   6
         Text            =   "Pedido Cancelado"
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Left            =   120
         TabIndex        =   5
         Text            =   "999999"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox Text23 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "FrmLegenda.frx":00D8
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   525
         Left            =   120
         TabIndex        =   3
         Text            =   "999999"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "FrmLegenda.frx":0112
         Top             =   3000
         Width           =   2895
      End
   End
   Begin VB.CommandButton BtoSair 
      BackColor       =   &H0080FFFF&
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
      Height          =   855
      Left            =   8280
      Picture         =   "FrmLegenda.frx":014F
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "FrmLegenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtoSair_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Me.Left = 1500
    Me.Top = 4000
    Me.Height = 4170
    Me.Width = 9480

End Sub
