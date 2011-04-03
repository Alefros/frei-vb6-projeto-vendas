VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Frmloca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro e controle de localizações"
   ClientHeight    =   4740
   ClientLeft      =   3585
   ClientTop       =   2775
   ClientWidth     =   8400
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Cmd_Ant 
      Caption         =   "Anterior"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   24
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Prox 
      Caption         =   "Próximo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   23
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Ultimo 
      Caption         =   "Último"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   22
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Primeiro 
      Caption         =   "Primeiro"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   21
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Excluir 
      BackColor       =   &H80000009&
      Caption         =   "Excluir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton Cmd_Incluir 
      BackColor       =   &H80000009&
      Caption         =   "Incluir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton cmd_Alterar 
      BackColor       =   &H80000009&
      Caption         =   "Alterar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton cmd_Novo 
      BackColor       =   &H80000009&
      Caption         =   "Novo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Estados"
      TabPicture(0)   =   "Frmloca.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "txt_Uf"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Cidades"
      TabPicture(1)   =   "Frmloca.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbo_Uf"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txt_Cid"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Bairros"
      TabPicture(2)   =   "Frmloca.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txt_Bairros"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cbo_Cid"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Localizações"
      TabPicture(3)   =   "Frmloca.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(1)=   "Label7"
      Tab(3).Control(2)=   "Label8"
      Tab(3).Control(3)=   "cbo_Bairro"
      Tab(3).Control(4)=   "txt_Loca"
      Tab(3).Control(5)=   "msk_Cep"
      Tab(3).ControlCount=   6
      Begin MSMask.MaskEdBox msk_Cep 
         Height          =   375
         Left            =   -72960
         TabIndex        =   19
         Top             =   2040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   9
         Mask            =   "99999-999"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txt_Loca 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   -72960
         TabIndex        =   16
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cbo_Bairro 
         Height          =   315
         Left            =   -72960
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cbo_Cid 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "Frmloca.frx":0070
         Left            =   1800
         List            =   "Frmloca.frx":0072
         TabIndex        =   13
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txt_Bairros 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cbo_Uf 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "Frmloca.frx":0074
         Left            =   -72960
         List            =   "Frmloca.frx":0076
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txt_Cid 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   -72960
         TabIndex        =   8
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txt_Uf 
         Height          =   495
         Left            =   -72960
         MaxLength       =   2
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Logradouro"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74520
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "CEP"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   18
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Bairros"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Cidades"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Uf"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   10
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Frmloca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l_uf As Integer

Dim l_coduf As Integer

Dim l_codcid As Integer

Dim l_codbair As Integer
Private Sub cbo_Bairro_Click()
            l_codbair = cbo_Bairro.ItemData(cbo_Bairro.ListIndex)
End Sub
Private Sub cbo_Cid_Click()
            l_codcid = cbo_Cid.ItemData(cbo_Cid.ListIndex)
End Sub
Private Sub cbo_Uf_Click()
            l_coduf = cbo_Uf.ItemData(cbo_Uf.ListIndex) ' ComboBox Guia CIDADES onde ela pega a informação da entidade UF'
End Sub
Private Sub cmd_Alterar_Click()
            status = "Alteradas "
            If SSTab1.Tab = 0 Then              'ENTIDADE UFs - GUIA ESTADOS'
                l_uf = tab_ufs!codigo
                If tab_ufs.State = adStateOpen Then
                    tab_ufs.Close
                    tab_ufs.Open "select * from ufs where nome ='" & txt_Uf & "'"   '* = seleção de todos os campos da entidade/ seleciona todos os campos da entidade ufs, onde o nome for igual ao que é digitado na caixinha de texto ;  '2ª instrução em SQL''
                        If tab_ufs.RecordCount <> 0 Then
                            If MsgBox("Deseja realmente alterar esta unidade federativa?", vbQuestion + vbYesNo + vbDefaultButton2, "vendas") = vbYes Then
                                ElseIf tab_ufs.RecordCount = 0 Then
                                tab_ufs.Close
                                tab_ufs.Open "update ufs set nome = '" & txt_Uf & "' where codigo = " & l_uf
                                Exit Sub
                End If
                        End If
                            End If
                
                ElseIf SSTab1.Tab = 1 Then          'ENTIDADE CIDADEs - GUIA CIDADES'
                
                    
                    ElseIf SSTab1.Tab = 2 Then          'ENTIDADE BAIRROs - GUIA BAIRROS'
                    
        
                        ElseIf SSTab1.Tab = 3 Then          'ENTIDADE LOCALIZACOEs - GUIA LOCALIZAÇÕES'
                    
            
            End If
            
            Call box_1
            Call limpar
End Sub
Private Sub Cmd_Ant_Click()
            tab_ufs.MovePrevious               'ENTIDADE UFs - GUIA ESTADOS'
            If tab_ufs.BOF = True Then
                tab_ufs.MoveLast
            End If
            txt_Uf = tab_ufs!nome
            
            tab_cid.MovePrevious            'ENTIDADE CIDADEs - GUIA CIDADES'
            If tab_cid.BOF = True Then
                tab_cid.MoveLast
            End If
            txt_Cid = tab_cid!nome
End Sub
Private Sub cmd_Excluir_Click()
            status = "Excluidas "
            If SSTab1.Tab = 0 Then                  'ENTIDADE UFs - GUIA ESTADOS'
            If tab_ufs.State = adStateOpen Then
               tab_ufs.Close
            End If
            tab_ufs.Open "select * from ufs where nome ='" & txt_Uf & "'"   '* = seleção de todos os campos da entidade/ seleciona todos os campos da entidade ufs, onde o nome for igual ao que é digitado na caixinha de texto ;  '2ª instrução em SQL''
                If tab_ufs.RecordCount = 0 Then
                    MsgBox "Esta unidade federativa não esta cadastrada, por favor verificar!", vbExclamation
                    Exit Sub
                End If
               If MsgBox("Deseja realmente excluir esta unidade federativa.", vbQuestion + vbYesNo + vbDefaultButton2, "Vendas") = vbYes Then    '3ª instrução em SQL'
                  conectar.Execute "Delete from ufs where nome ='" & txt_Uf & "'"     '3ª instrução em SQL'
                  Else
                  Exit Sub
                End If
            
                ElseIf SSTab1.Tab = 1 Then          'ENTIDADE CIDADEs - GUIA CIDADES'
                       If tab_cid.State = adStateOpen Then
                        tab_cid.Close
                       End If
            tab_cid.Open "select * from cidades where nome ='" & txt_Cid & "'"   '* = seleção de todos os campos da entidade/ seleciona todos os campos da entidade ufs, onde o nome for igual ao que é digitado na caixinha de texto ;  '2ª instrução em SQL''
                If tab_cid.RecordCount = 0 Then
                    MsgBox ("Esta cidade não está cadastrada, por favor verificar!"), vbExclamation
                End If
                Exit Sub
               If MsgBox("Deseja realmente excluir esta cidade.", vbQuestion + vbYesNo + vbDefaultButton2, "Vendas") = vbYes Then   '3ª instrução em SQL'
                  conectar.Execute "Delete from cidades where nome ='" & txt_Cid & "'"     '3ª instrução em SQL'
                  Else
                  Exit Sub
                End If
                    
                    ElseIf SSTab1.Tab = 2 Then          'ENTIDADE BAIRROs - GUIA BAIRROS'
                       If tab_bar.State = adStateOpen Then
                        tab_bar.Close
                       End If
            tab_bar.Open "select * from bairros where nome ='" & txt_Bairros & "'"   '* = seleção de todos os campos da entidade/ seleciona todos os campos da entidade ufs, onde o nome for igual ao que é digitado na caixinha de texto ;  '2ª instrução em SQL''
                If tab_bar.RecordCount = 0 Then
                    MsgBox ("Este bairro não está cadastrado, por favor verificar!"), vbExclamation
                End If
                Exit Sub
               If MsgBox("Deseja realmente excluir este bairro.", vbQuestion + vbYesNo + vbDefaultButton2, "Vendas") = vbYes Then   '3ª instrução em SQL'
                  conectar.Execute "Delete from bairros where nome ='" & txt_Bairros & "'"     '3ª instrução em SQL'
                  Else
                  Exit Sub
                End If
                
                    ElseIf SSTab1.Tab = 3 Then          'ENTIDADE LOCALIZACOEs - GUIA LOCALIZAÇÕES'
                            If tab_loca.State = adStateOpen Then
                        tab_loca.Close
                       End If
            tab_loca.Open "select * from localizacoes where nome ='" & txt_Loca & "'"   '* = seleção de todos os campos da entidade/ seleciona todos os campos da entidade ufs, onde o nome for igual ao que é digitado na caixinha de texto ;  '2ª instrução em SQL''
                If tab_loca.RecordCount = 0 Then
                    MsgBox ("Este logradouro não está cadastrado, por favor verificar!"), vbExclamation
                End If
                Exit Sub
               If MsgBox("Deseja realmente excluir esta localização.", vbQuestion + vbYesNo + vbDefaultButton2, "Vendas") = vbYes Then   '3ª instrução em SQL'
                  conectar.Execute "Delete from localizacoes where nome ='" & txt_Loca & "'"     '3ª instrução em SQL'
                  Else
                  Exit Sub
                End If
            
                End If
            Call box_1
            Call limpar
            txt_Uf.SetFocus
            Cmd_excluir.Enabled = False
End Sub
Private Sub Cmd_Incluir_Click()
            status = "Incluidas "
            If SSTab1.Tab = 0 Then          'ENTIDADE UFs - GUIA ESTADOS'
                If tab_ufs.State = adStateOpen Then
                   tab_ufs.Close
                End If
                    tab_ufs.Open "select * from ufs where nome ='" & txt_Uf & "'"   '* = seleção de todos os campos da entidade/ seleciona todos os campos da entidade ufs, onde o nome for igual ao que é digitado na caixinha de texto ;  '2ª instrução em SQL''
                    If tab_ufs.RecordCount <> 0 Then
                        MsgBox "Atenção! Esta unidade federativa já foi cadastrada, por favor, verificar.", vbExclamation
                    Exit Sub
                    Else
                    conectar.Execute "Insert into ufs(nome) Values('" & txt_Uf & "')"   '1ª instrução em SQL'
                    End If
            
            
            ElseIf SSTab1.Tab = 1 Then          'ENTIDADE CIDADEs - GUIA CIDADES'
                   If tab_cid.State = adStateOpen Then
                        tab_cid.Close
                   End If
                   tab_cid.Open "select * from cidades where nome ='" & txt_Cid & "'"   '* = seleção de todos os campos da entidade/ seleciona todos os campos da entidade ufs, onde o nome for igual ao que é digitado na caixinha de texto ;  '2ª instrução em SQL''
                    If tab_cid.RecordCount <> 0 Then
                        MsgBox "Atenção! Esta cidade já foi cadastrada, por favor, verificar.", vbExclamation
                        Exit Sub
                    Else
                        conectar.Execute "Insert into cidades(nome, cod_uf) Values('" & txt_Cid & "', '" & l_coduf & "')"
                    End If
                
                ElseIf SSTab1.Tab = 2 Then          'ENTIDADE BAIRROs - GUIA BAIRROS'
                       If tab_bar.State = adStateOpen Then
                        tab_bar.Close
                   End If
                   tab_bar.Open "select * from bairros where nome ='" & txt_Bairros & "'"   '* = seleção de todos os campos da entidade/ seleciona todos os campos da entidade ufs, onde o nome for igual ao que é digitado na caixinha de texto ;  '2ª instrução em SQL''
                    If tab_bar.RecordCount <> 0 Then
                        MsgBox "Atenção! Este bairro já foi cadastrado, por favor, verificar.", vbExclamation
                        Exit Sub
                    Else
                        conectar.Execute "Insert into bairros (nome, cod_cid) Values('" & txt_Bairros & "', '" & l_codcid & "')"
                    End If
                
                    ElseIf SSTab1 = 3 Then          'ENTIDADE LOCALIZACOEs - GUIA LOCALIZAÇÕES'
                    If tab_loca.State = adStateOpen Then
                        tab_loca.Close
                    End If
                   tab_loca.Open "select * from localizacoes where logradouro ='" & txt_Loca & "'"   '* = seleção de todos os campos da entidade/ seleciona todos os campos da entidade ufs, onde o nome for igual ao que é digitado na caixinha de texto ;  '2ª instrução em SQL''
                    If tab_loca.RecordCount <> 0 Then
                        MsgBox "Atenção! Este logradouro já foi cadastrado, por favor, verificar.", vbExclamation
                        Exit Sub
                    Else
                        conectar.Execute "Insert into localizacoes (logradouro, cod_bairro) Values('" & txt_Loca & "', '" & l_bair & "')"
                    End If
                    
            End If
            Call box_1
            Call limpar
            txt_Uf.SetFocus
            Cmd_Incluir.Enabled = False
End Sub
Private Sub limpar()
            If SSTab1.Tab = 0 Then
               txt_Uf = Clear
            ElseIf SSTab1.Tab = 1 Then
                   txt_Cid = Clear
                   cbo_Uf.Text = Clear
                ElseIf SSTab1.Tab = 2 Then
                       txt_Bairros = Clear
                       cbo_Cid.Text = Clear
                    ElseIf SSTab1.Tab = 3 Then
                           txt_Loca = Clear
                           cbo_Bairro.Text = Clear
                           msk_Cep.PromptInclude = False
                           msk_Cep.Text = Clear
                           msk_Cep.PromptInclude = True
            End If
End Sub
Private Sub cmd_Novo_Click()
            Call limpar
End Sub
Private Sub Cmd_Primeiro_Click()
            tab_ufs.MoveFirst
            txt_Uf = tab_ufs!nome
            
            tab_cid.MoveFirst
            txt_Cid = tab_cid!nome
            cbo_Uf = tab_ufs!nome
            
            
End Sub
Private Sub Cmd_Prox_Click()
            tab_ufs.MoveNext
            If tab_ufs.EOF = True Then
               tab_ufs.MoveFirst
            End If
            txt_Uf = tab_ufs!nome
            
            tab_cid.MoveNext
            If tab_cid.EOF = True Then
               tab_cid.MoveFirst
            End If
            txt_Cid = tab_cid!nome
            cbo_Uf = tab_ufs!nome
End Sub
Private Sub Cmd_Ultimo_Click()
            tab_ufs.MoveLast
            txt_Uf = tab_ufs!nome
            
            tab_cid.MoveLast
            txt_Cid = tab_cid!nome
            cbo_Uf = tab_ufs!nome            ' " ! " E O QUE SEPARA A VARIÁVEL DO CAMPO AO QUAL VOCÊ QUER IDENTIFICAR'
End Sub
Private Sub Form_Load()
            If tab_ufs.State = adStateOpen Then tab_ufs.Close
            tab_ufs.Open "ufs", conectar, adOpenKeyset, adLockOptimistic 'adOpenKeyset = ajuste do teclado'
            
            If tab_cid.State = adStateOpen Then tab_cid.Close
            tab_cid.Open "cidades", conectar, adOpenKeyset, adLockOptimistic
            
            If tab_bar.State = adStateOpen Then tab_bar.Close
            tab_bar.Open "bairros", conectar, adOpenKeyset, adLockOptimistic
            
            If tab_loca.State = adStateOpen Then tab_loca.Close
            tab_loca.Open "localizacoes", conectar, adOpenKeyset, adLockOptimistic
            Do While tab_ufs.EOF = False     ' ComboBox da Guia CIDADES'
            cbo_Uf.AddItem tab_ufs!nome
            cbo_Uf.ItemData(cbo_Uf.NewIndex) = tab_ufs!codigo
            tab_ufs.MoveNext
            Loop
                    Do While tab_cid.EOF = False    ' ComboBox da Guia BAIRROS'
                     cbo_Cid.AddItem tab_cid!nome
                     cbo_Cid.ItemData(cbo_Cid.NewIndex) = tab_cid!codigo
                     tab_cid.MoveNext
                    Loop
                        Do While tab_bar.EOF = False    ' ComboBox da Guia LOCALIZAÇÕES'
                     cbo_Bairro.AddItem tab_bar!nome
                     cbo_Bairro.ItemData(cbo_Bairro.NewIndex) = tab_bar!codigo
                     tab_bar.MoveNext
                    Loop
                
End Sub

Private Sub txt_Bairros_Change()
            Cmd_Incluir.Enabled = True
End Sub
Private Sub txt_Cid_Change()
            Cmd_Incluir.Enabled = True 'Proibe gravar registro em branco, na caixa de texto Nome da UF'
                Cmd_Incluir.Enabled = True          'Proibe gravar registro em branco, na caixa de texto Nome da UF'
                If txt_Cid = Empty Then          'BOTÃO INCLUIR'
                    Cmd_Incluir.Enabled = False
                End If
                            
            Cmd_excluir.Enabled = True          'BOTÃO EXCLUIR'
            If txt_Cid = Empty Then
                    Cmd_excluir.Enabled = False
            End If
                 
                 Cmd_alterar.Enabled = True             'BOTÃO ALTERAR'
            If txt_Cid = Empty Then
                    Cmd_alterar.Enabled = False
            End If
End Sub
Private Sub txt_Uf_Change()
            Cmd_Incluir.Enabled = True          'Proibe gravar registro em branco, na caixa de texto Nome da UF'
                If txt_Uf = Empty Then          'BOTÃO INCLUIR'
                    Cmd_Incluir.Enabled = False
                End If
                 If Len(txt_Uf) <> 2 Then   'LEN= contador de caracteres, essa linha não permite a inclusão de uma letra só'
                Cmd_Incluir.Enabled = False
             End If
            
            Cmd_excluir.Enabled = True          'BOTÃO EXCLUIR'
            If txt_Uf = Empty Then
                    Cmd_excluir.Enabled = False
            End If
                 If Len(txt_Uf) <> 2 Then   'LEN= contador de caracteres, essa linha não permite a inclusão de uma letra só'
                Cmd_excluir.Enabled = False
                 End If
                 
                 Cmd_alterar.Enabled = True             'BOTÃO ALTERAR'
            If txt_Uf = Empty Then
                    Cmd_alterar.Enabled = False
            End If
                 If Len(txt_Uf) <> 2 Then   'LEN= contador de caracteres, essa linha não permite a inclusão de uma letra só'
                Cmd_alterar.Enabled = False
                 End If
End Sub
Private Sub txt_Uf_KeyPress(KeyAscii As Integer)
            If KeyAscii < 65 Or KeyAscii > 90 Then
                If KeyAscii < 97 Or KeyAscii > 122 Then
                KeyAscii = 8
            End If
                End If
End Sub
Private Sub txt_Uf_LostFocus()
            txt_Uf = UCase(txt_Uf) ' lostfocus + UCase serve para: ao digitar em minusculas, ele tranforma as informações em maiusculas'
End Sub

