VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmprodutos 
   Caption         =   "Produtos"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Controles"
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
      TabIndex        =   16
      Top             =   5880
      Width           =   9495
      Begin VB.CommandButton Command4 
         Caption         =   "Alterar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   20
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Excluir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Gravar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   18
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   17
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   9495
      Begin MSFlexGridLib.MSFlexGrid Mfg_produtos 
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   1720
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações dos produtos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   9495
      Begin MSMask.MaskEdBox Msk_validade 
         Height          =   375
         Left            =   7320
         TabIndex        =   3
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox Txt_descricao 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   1320
         Width           =   4815
      End
      Begin VB.TextBox Txt_unidade 
         Height          =   375
         Left            =   4560
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Txt_preco 
         Height          =   375
         Left            =   7320
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox Cbo_categoria 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Text            =   "Selecione aqui a categoria"
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Txt_produto 
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   8055
      End
      Begin VB.Label Label8 
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Unidade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Preço unitário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Categoria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Validade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nome"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Validade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmprodutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub Form_Load()

End Sub
