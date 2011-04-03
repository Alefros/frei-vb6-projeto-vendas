VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmcategoria 
   Caption         =   "Cadastro de Categorias"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_excluir 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comandos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   6735
      Begin VB.CommandButton Cmd_buscar 
         Caption         =   "Buscar"
         Height          =   735
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Cmd_alterar 
         Caption         =   "Alterar"
         Height          =   495
         Left            =   4080
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_gravar 
         Caption         =   "Gravar"
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_novo 
         Caption         =   "Novo"
         Height          =   495
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Categorias cadastradas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   6735
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2295
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4048
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações das categorias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6735
      Begin VB.TextBox Txt_cod_cat 
         Height          =   375
         Left            =   4920
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Txt_categoria 
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Categoria"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmcategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_gravar_Click()
            if tab_cat.State
End Sub

Private Sub Form_Load()
            tab_cat.Open "Categorias", conectar, adOpenKeyset, adLockOptimistic
End Sub



