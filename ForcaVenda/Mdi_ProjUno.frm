VERSION 5.00
Begin VB.Form Mdi_ProjUno 
   Caption         =   "U N O C A N N  -  Força de Vendas"
   ClientHeight    =   6870
   ClientLeft      =   1320
   ClientTop       =   1905
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   10365
   Begin VB.Menu Mnu_Cadastros 
      Caption         =   "&Cadastros"
      Begin VB.Menu Mnu_Clientes 
         Caption         =   "C&lientes"
      End
      Begin VB.Menu Mnu_representantes 
         Caption         =   "&Representantes"
      End
      Begin VB.Menu Mnu_produtos 
         Caption         =   "&Produtos"
      End
   End
   Begin VB.Menu Mnu_tabelas 
      Caption         =   "&Tabelas"
      Begin VB.Menu Mnu_parametros 
         Caption         =   "Parâ&metros"
      End
      Begin VB.Menu mnu_tributacao 
         Caption         =   "Tri&butação"
      End
      Begin VB.Menu mnu_precos 
         Caption         =   "Preç&os"
      End
   End
   Begin VB.Menu mnu_gerenciamento 
      Caption         =   "&Gerenciamento"
      Begin VB.Menu mnu_importabase 
         Caption         =   "&Importa Base"
      End
      Begin VB.Menu mnu_gerainterface 
         Caption         =   "Gera inter&face"
      End
   End
   Begin VB.Menu mnu_consultas 
      Caption         =   "Consu&ltas"
   End
End
Attribute VB_Name = "Mdi_ProjUno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Mnu_Clientes_Click()
'  CarregaForm FrmConhecimento, "teste"
FrmConhecimento.Show


End Sub
