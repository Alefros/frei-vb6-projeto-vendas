VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_cadastro 
   Caption         =   "Cadastro e Controle de Clientes"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid mfgcli 
      Height          =   1815
      Left            =   120
      TabIndex        =   24
      Top             =   5040
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   4
      ScrollBars      =   0
      FormatString    =   "Código|Nome                    |CPF                         |     Celular                    "
   End
   Begin VB.Frame fra_controles 
      Caption         =   "Controles"
      Height          =   1215
      Left            =   120
      TabIndex        =   54
      Top             =   6960
      Width           =   9735
      Begin VB.CommandButton cmdnovo 
         Caption         =   "Novo"
         Height          =   855
         Left            =   7440
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdgravar 
         Caption         =   "Gravar"
         Height          =   855
         Left            =   5400
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdexcluir 
         Caption         =   "Excluir"
         Height          =   855
         Left            =   3360
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdalterar 
         Caption         =   "Alterar"
         Height          =   855
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fra_contatos 
      Caption         =   "Contatos"
      Height          =   2175
      Left            =   120
      TabIndex        =   44
      Top             =   2760
      Width           =   9735
      Begin VB.TextBox txtramal 
         Height          =   285
         Left            =   8760
         TabIndex        =   23
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtfccel 
         Height          =   285
         Left            =   3360
         TabIndex        =   22
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox txtfccomercial 
         Height          =   285
         Left            =   3360
         TabIndex        =   20
         Top             =   1200
         Width           =   4575
      End
      Begin VB.TextBox txtfcrecados 
         Height          =   285
         Left            =   3360
         TabIndex        =   18
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox txtfcres 
         Height          =   285
         Left            =   3360
         TabIndex        =   16
         Top             =   240
         Width           =   4575
      End
      Begin MSMask.MaskEdBox mskres 
         Height          =   375
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "9999-9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskrecados 
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "9999-9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskcomercial 
         Height          =   375
         Left            =   1080
         TabIndex        =   19
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "9999-9999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskcel 
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Top             =   1680
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "9999-9999"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_ramal 
         Caption         =   "Ramal"
         Height          =   255
         Left            =   8160
         TabIndex        =   53
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lbl_falarcom_cel 
         Caption         =   "Falar com"
         Height          =   255
         Left            =   2400
         TabIndex        =   52
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lbl_falarcom_com 
         Caption         =   "Falar com"
         Height          =   255
         Left            =   2400
         TabIndex        =   51
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lbl_falarcom_rec 
         Caption         =   "Falar com"
         Height          =   255
         Left            =   2400
         TabIndex        =   50
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbl_falarcom_res 
         Caption         =   "Falar com"
         Height          =   255
         Left            =   2400
         TabIndex        =   49
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl_residencial 
         Caption         =   "Residencial"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl_recados 
         Caption         =   "Recados"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lbl_comercial 
         Caption         =   "Comercial"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lbl_celular 
         Caption         =   "Celular"
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   1800
         Width           =   495
      End
   End
   Begin VB.Frame fra_fumante 
      Caption         =   "Fumante"
      Height          =   735
      Left            =   8400
      TabIndex        =   43
      Top             =   1920
      Width           =   1455
      Begin VB.OptionButton optnao 
         Alignment       =   1  'Right Justify
         Caption         =   "Não"
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optsim 
         Caption         =   "Sim"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fra_sexo 
      Caption         =   "Sexo"
      Height          =   735
      Left            =   6000
      TabIndex        =   42
      Top             =   1920
      Width           =   2295
      Begin VB.OptionButton optmasculino 
         Alignment       =   1  'Right Justify
         Caption         =   "Masculino"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optfeminino 
         Caption         =   "Feminino"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fra_documentacao 
      Caption         =   "Documentação"
      Height          =   735
      Left            =   120
      TabIndex        =   38
      Top             =   1920
      Width           =   5775
      Begin MSMask.MaskEdBox mskrg 
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "99.999.999-9"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskcpf 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   14
         Mask            =   "999.999.999-99"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msknascimento 
         Height          =   375
         Left            =   4560
         TabIndex        =   10
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_nascimento 
         Caption         =   "Nascimento"
         Height          =   255
         Left            =   3600
         TabIndex        =   41
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl_cpf 
         Caption         =   "CPF"
         Height          =   255
         Left            =   1800
         TabIndex        =   40
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lbl_rg 
         Caption         =   "RG"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame fra_dados 
      Caption         =   "Dados Pessoais"
      Height          =   1695
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   9735
      Begin VB.ComboBox Cbo_estado 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtcidade 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtbairro 
         Height          =   285
         Left            =   6960
         TabIndex        =   6
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtcomplemento 
         Height          =   285
         Left            =   3360
         TabIndex        =   5
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtnumero 
         Height          =   285
         Left            =   9000
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtloca 
         Height          =   285
         Left            =   3360
         TabIndex        =   2
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txtnome 
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   8895
      End
      Begin MSMask.MaskEdBox mskcep 
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "99999-999"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_complemento 
         Caption         =   "Complemento"
         Height          =   255
         Left            =   2280
         TabIndex        =   37
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbl_bairro 
         Caption         =   "Bairro"
         Height          =   255
         Left            =   6360
         TabIndex        =   36
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbl_cidade 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lbl_estado 
         Caption         =   "Estado"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbl_numero 
         Caption         =   "Nº"
         Height          =   255
         Left            =   8640
         TabIndex        =   33
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lbl_logradouro 
         Caption         =   "Logradouro"
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lbl_cep 
         Caption         =   "CEP"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lbl_nome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frm_cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l_estados As Integer

Dim l_codcli As Integer

Dim l_codcli2 As Integer
Private Sub carregarlista()
            'mfgcli.Rows = 2
            'mfgcli.Clear
            'mfgcli.FormatString = "Código|Nome|CPF|Celular"
            On Error Resume Next
            If tabcli.State = adStateOpen Then tabcli.Close
            tabcli.Open "clientes", conectar, adOpenKeyset, adLockOptimistic
            Do Until tabcli.EOF
                       mfgcli.TextMatrix(mfgcli.Rows - 1, 0) = tabcli!codigo
                       mfgcli.TextMatrix(mfgcli.Rows - 1, 1) = tabcli!nome
                       mfgcli.TextMatrix(mfgcli.Rows - 1, 2) = Format(tabcli!CPF, "&&&.&&&.&&&-&&")
                       mfgcli.TextMatrix(mfgcli.Rows - 1, 3) = Format(tabcli!tel_cel, "(&&) &&&&-&&&&")
                       mfgcli.RowData(mfgcli.Rows - 1) = tabcli!codigo
                       tabcli.MoveNext
                       mfgcli.Rows = mfgcli.Rows + 1
                       Loop
                       mfgcli.Rows = mfgcli.Rows - 1
End Sub

Private Sub Cbo_estado_Click()
            'tab_ufs.MoveNext
            'l_estados = Cbo_estado.ItemData(cbo_Uf.ListIndex)
End Sub

Private Sub Cbo_estado_GotFocus()
            'Cbo_estado = tab_ufs!nome 'carrega a primeira UF na combo estados
            
            'l_estados = Cbo_estado.ItemData(Cbo_estado.ListIndex)
            'cbo_Uf.ItemData (cbo_Uf.ListIndex)
End Sub
Private Sub Limpar_cliente()
            
            Call Desabilitar_Mascara
            mskcep = Clear
            mskres = Clear
            mskrecados = Clear
            mskcomercial = Clear
            mskcel = Clear
            mskrg = Clear
            mskcpf = Clear
            msknascimento = Clear
                txtnome = Clear
                txtloca = Clear
                txtnumero = Clear
                txtcomplemento = Clear
                txtbairro = Clear
                txtcidade = Clear
                txtfcres = Clear
                txtfcrecados = Clear
                txtfccomercial = Clear
                txtfccel = Clear
                txtramal = Clear
            Call Habilitar_Mascara
            txtnome.SetFocus
            
End Sub

Private Sub cmdalterar_Click()

End Sub

Private Sub cmdexcluir_Click()

End Sub

Private Sub cmdGravar_Click()
            status = "Gravadas"
            If txtnome = Empty Then
               MsgBox "Não é possível gravar as informações." & Chr(13) & "Digite primeiro o nome do cliente", vbInformation
               Exit Sub
            End If
               
            Call gravar_cliente
            Call box_1
            Call Limpar_cliente
            Call carregarlista
            
            
End Sub

Private Sub cmdnovo_Click()
            'cmdgravar.Enabled = True
            'Call Fechar_Arquivos
            'Call Abrir_Arquivos
            Call Limpar_cliente
            'Call Carregar_Lista
            
End Sub

Private Sub Form_Load()
            Call carregarlista
            
            If tabcli.State = adStateOpen Then tabcli.Close
            tabcli.Open "clientes", conectar, adOpenKeyset, adLockOptimistic
            
            
            If tab_loca.State = adStateOpen Then tab_loca.Close
            tab_loca.Open "localizacoes", conectar, adOpenKeyset, adLockOptimistic
            
            
            If tab_ufs.State = adStateOpen Then tab_ufs.Close
            tab_ufs.Open "ufs", conectar, adOpenKeyset, adLockOptimistic
            
            
            If tab_cid.State = adStateOpen Then tab_cid.Close
            tab_cid.Open "cidades", conectar, adOpenKeyset, adLockOptimistic
            
            
            Do While tab_ufs.EOF = False     ' ComboBox da Guia CIDADES'
            Cbo_estado.AddItem tab_ufs!nome
            Cbo_estado.ItemData(Cbo_estado.NewIndex) = tab_ufs!codigo
            tab_ufs.MoveNext
            Loop
            
           
          
End Sub
Private Sub gravar_cliente()
            Call Desabilitar_Mascara
            
            If status <> "Alteradas" Then tabcli.AddNew
            tabcli!nome = txtnome.Text
            tabcli!tel_res = mskres.Text
            tabcli!CPF = mskcpf.Text
            tabcli!numero = txtnumero.Text
            tabcli!complemento = txtcomplemento.Text
            tabcli!tel_com = mskcomercial.Text
            tabcli!tel_cel = mskcel.Text
            tabcli!tel_recados = mskrecados.Text
            tabcli!rg = mskrg.Text
            tabcli!CPF = mskcpf.Text
            'tabcli!sexo=
            tabcli!nascimento = msknascimento.Text
            tabcli!cep = mskcep.Text
            tabcli!tel_cel = mskcel.Text
            
            
            
            Call Habilitar_Mascara
End Sub


Private Sub fra_controles_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub mfgcli_Click()
            l_codcli = mfgcli.RowData(mfgcli.Row)
                If tabcli.State = adStateOpen Then tabcli.Close
                tabcli.Open "select * from clientes where codigo =" & l_codcli
                    If tabcli.RecordCount <> 0 Then Call Habilitar_Mascara
                   
            'L_Linha = mfgVend.Row
            'l_codcli = mfgcli.Row
            
            'L_CodVend = mfgVend.TextMatrix(L_Linha, 0)
            'l_codcli = mfgcli.TextMatrix(l_codcli, 0)
            
            'If Tab_Vend.State = adStateOpen Then Tab_Vend.Close
            'If tabcli.State = adStateOpen Then tabcli.Close
            
            'Tab_Vend.Open "Select * From Vendedores Where Inscricao = " & L_CodVend
            'tabcli.Open "Select * From clientes Where codigo = " & l_codcli
            
            'Call Mostrar_Vend
            'Call desativar
            
            'cmdGravar.Enabled = False
          
           
            
                
                
End Sub
Private Sub Exibir()
            
            Call desativar
            mskcep = tabcli!cep
            mskrg = tabcli!rg
            mskcpf = tabcli!CPF
            msknascimento = tabcli!nascimento
            mskres = tabcli!tel_res
            mskcomercial = tabcli!tel_com
            mskcel = tabcli!tel_cel
            Call ativar
            
            txtnome = tabcli!nome
            txtloca = tab_loca!logradouro
            txtnumero = tabcli!numero
            'txt_estado = tab_ufs!nome
            
            'txtcomplemento = tabcli!complemento
            
            txtcidade = tab_cid!nome

            
End Sub
Private Sub Habilitar_Mascara()
            mskcep.PromptInclude = True
            mskrg.PromptInclude = True
            mskcpf.PromptInclude = True
            msknascimento.PromptInclude = True
            mskres.PromptInclude = True
            mskrecados.PromptInclude = True
            mskcomercial.PromptInclude = True
            mskcel.PromptInclude = True
End Sub

Private Sub Desabilitar_Mascara()
            mskcep.PromptInclude = False
            mskres.PromptInclude = False
            mskrecados.PromptInclude = False
            mskcomercial.PromptInclude = False
            mskcel.PromptInclude = False
            mskrg.PromptInclude = False
            mskcpf.PromptInclude = False
            msknascimento.PromptInclude = False
            
End Sub
