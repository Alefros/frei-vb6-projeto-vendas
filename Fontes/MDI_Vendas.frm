VERSION 5.00
Begin VB.MDIForm MDI_Vendas 
   BackColor       =   &H8000000C&
   Caption         =   "Projeto Vendas"
   ClientHeight    =   6900
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9240
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_Cadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu mnu_Clientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnu_Loca 
         Caption         =   "Localizações"
      End
      Begin VB.Menu mnutr01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_For 
         Caption         =   "Fornecedores"
      End
      Begin VB.Menu mnu_categoria 
         Caption         =   "Categorias"
      End
      Begin VB.Menu mnu_Prod 
         Caption         =   "Produtos"
      End
      Begin VB.Menu mnu_Vend 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu mnu_ped 
         Caption         =   "Pedidos"
      End
      Begin VB.Menu mnutr02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Sair 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "MDI_Vendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
            Call abrir_banco
End Sub

Private Sub mnu_categoria_Click()
            frmcategoria.Show
End Sub

Private Sub mnu_Clientes_Click()
            frm_cadastro.Show
End Sub

Private Sub mnu_For_Click()
            'frm_fornecedores.Show
End Sub

Private Sub mnu_Loca_Click()
            Frmloca.Show
End Sub

Private Sub mnu_ped_Click()
            frmPedidos.Show
End Sub

Private Sub mnu_Prod_Click()
            frmprodutos.Show
End Sub

Private Sub mnu_Sair_Click()
            End
End Sub

Private Sub mnu_Vend_Click()
            frmVendedores.Show
End Sub
