VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cliente"
   ClientHeight    =   6960
   ClientLeft      =   6375
   ClientTop       =   3720
   ClientWidth     =   11055
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11055
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FmeBotoes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   30
      TabIndex        =   48
      Top             =   0
      Width           =   10935
      Begin VB.CommandButton cmd_Adicionar 
         Caption         =   " &Novo"
         Height          =   675
         Left            =   120
         Picture         =   "frmCliente.frx":0BC2
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_Pesquisar 
         Caption         =   "&Pesquisar"
         Height          =   675
         Left            =   1200
         Picture         =   "frmCliente.frx":10F4
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_Limpar 
         Caption         =   " &Limpar"
         Height          =   675
         Left            =   2280
         Picture         =   "frmCliente.frx":1626
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_Gravar 
         Caption         =   "&Gravar"
         Enabled         =   0   'False
         Height          =   675
         Left            =   3360
         Picture         =   "frmCliente.frx":1B58
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_Excluir 
         Caption         =   "&Excluir"
         Enabled         =   0   'False
         Height          =   675
         Left            =   7560
         Picture         =   "frmCliente.frx":208A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmd_Sair 
         Caption         =   "&Sair"
         Height          =   675
         Left            =   4440
         Picture         =   "frmCliente.frx":25BC
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FmeClientes 
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
      Height          =   5715
      Left            =   30
      TabIndex        =   27
      Top             =   1080
      Width           =   10935
      Begin VB.ListBox lstRazaoSocial 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   780
         ItemData        =   "frmCliente.frx":26B6
         Left            =   1800
         List            =   "frmCliente.frx":26BD
         TabIndex        =   29
         Top             =   5520
         Visible         =   0   'False
         Width           =   8880
      End
      Begin VB.TextBox txtSite 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1935
         MaxLength       =   40
         TabIndex        =   14
         Top             =   3165
         Width           =   3405
      End
      Begin VB.TextBox txtCidEndPrincipal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1935
         MaxLength       =   60
         TabIndex        =   7
         Top             =   1908
         Width           =   3435
      End
      Begin VB.TextBox txtNroEndPrincipal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1491
         Width           =   960
      End
      Begin VB.TextBox txtInscMunicipal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1935
         MaxLength       =   30
         TabIndex        =   12
         Top             =   2742
         Width           =   1820
      End
      Begin VB.TextBox txtCgc_Cpf 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9000
         MaxLength       =   18
         TabIndex        =   9
         Top             =   1908
         Width           =   1820
      End
      Begin VB.TextBox txtFone1Principal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         MaxLength       =   9
         TabIndex        =   10
         Top             =   2325
         Width           =   1020
      End
      Begin VB.TextBox txtCelular 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3915
         MaxLength       =   9
         TabIndex        =   11
         Top             =   2325
         Width           =   1620
      End
      Begin VB.TextBox TxtInscest 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5865
         MaxLength       =   30
         TabIndex        =   13
         Top             =   2742
         Width           =   1820
      End
      Begin VB.TextBox txtContatoPrincipal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5535
         MaxLength       =   20
         TabIndex        =   17
         Top             =   3993
         Width           =   2190
      End
      Begin VB.TextBox txtObsPrincipal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1935
         MaxLength       =   200
         TabIndex        =   18
         Top             =   4380
         Width           =   8895
      End
      Begin VB.TextBox TxtRazaoSocial 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1935
         MaxLength       =   45
         TabIndex        =   0
         Top             =   657
         Width           =   6135
      End
      Begin VB.TextBox txtComplEndPrincipal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3900
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1491
         Width           =   2460
      End
      Begin VB.TextBox txtEndPrincipal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4515
         MaxLength       =   60
         TabIndex        =   3
         Top             =   1074
         Width           =   6285
      End
      Begin VB.TextBox txtUFEndPrincipal 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7260
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1908
         Width           =   480
      End
      Begin VB.TextBox txtBaiEndPrincipal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7260
         MaxLength       =   45
         TabIndex        =   6
         Top             =   1491
         Width           =   3540
      End
      Begin VB.Frame Frame5 
         Caption         =   "Formas de Bloqueio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   585
         Left            =   5865
         TabIndex        =   28
         Top             =   -15
         Width           =   5055
         Begin VB.VScrollBar Scroll 
            Height          =   255
            Left            =   4680
            TabIndex        =   53
            Top             =   240
            Width           =   135
         End
         Begin VB.CheckBox chk_Bloqueado 
            Caption         =   "Crédito Bloqueado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   1
            ToolTipText     =   "Se Verificado Notifica na Tela de Pedidos"
            Top             =   285
            Width           =   1695
         End
         Begin VB.Label lblDiasEmAtraso 
            Alignment       =   2  'Center
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   4200
            TabIndex        =   52
            Top             =   285
            Width           =   255
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Dias de Atraso em C. Pend. :"
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
            Height          =   240
            Left            =   2040
            TabIndex        =   51
            Top             =   300
            Width           =   2070
         End
      End
      Begin VB.TextBox txtemail 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1935
         MaxLength       =   60
         TabIndex        =   15
         Top             =   3570
         Width           =   5775
      End
      Begin MSMask.MaskEdBox mskCepEndPrincipal 
         Height          =   360
         Left            =   1935
         TabIndex        =   2
         Top             =   1074
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#####-###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskCliDesde 
         Height          =   360
         Left            =   1935
         TabIndex        =   16
         Top             =   3993
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Frame fmeTipoCli 
         Height          =   615
         Left            =   1920
         TabIndex        =   54
         Top             =   4710
         Width           =   4815
         Begin VB.TextBox txtPagDia 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4200
            TabIndex        =   20
            Top             =   180
            Width           =   525
         End
         Begin VB.TextBox txtPrazo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1650
            TabIndex        =   19
            Top             =   180
            Width           =   525
         End
         Begin VB.Label Label11 
            Caption         =   "Pagamento dia :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   2460
            TabIndex        =   56
            Top             =   180
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Prazo em dias :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   60
            TabIndex        =   55
            Top             =   180
            Width           =   1545
         End
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pendente :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         TabIndex        =   57
         Top             =   4860
         Width           =   1755
      End
      Begin VB.Label lblcodCli 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2010
         TabIndex        =   50
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   1770
      End
      Begin VB.Label lblCNPJ 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ /CPF:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   7770
         TabIndex        =   47
         Top             =   1908
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Fone1:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         TabIndex        =   46
         Top             =   2325
         Width           =   1740
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Celular:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   3000
         TabIndex        =   45
         Top             =   2325
         Width           =   750
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Caption         =   "Email :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   135
         TabIndex        =   44
         Top             =   3576
         Width           =   1755
      End
      Begin VB.Label lblInscMun 
         Alignment       =   1  'Right Justify
         Caption         =   "Insc. Municipal :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         TabIndex        =   43
         Top             =   2742
         Width           =   1740
      End
      Begin VB.Label LblInscEst 
         Alignment       =   1  'Right Justify
         Caption         =   "Insc. Estadual :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   4110
         TabIndex        =   42
         Top             =   2742
         Width           =   1710
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Contato :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   4410
         TabIndex        =   41
         Top             =   3990
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Observaï¿½ï¿½es:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         TabIndex        =   40
         Top             =   4410
         Width           =   1755
      End
      Begin VB.Label lblRazaoSocial 
         Alignment       =   1  'Right Justify
         Caption         =   "Razao Social :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         TabIndex        =   39
         Top             =   657
         Width           =   1770
      End
      Begin VB.Label lblCepPrincipal 
         Alignment       =   1  'Right Justify
         Caption         =   "CEP :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         TabIndex        =   38
         Top             =   1074
         Width           =   1755
      End
      Begin VB.Label lblComplEndPrincipal 
         Alignment       =   1  'Right Justify
         Caption         =   "Compl. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   2940
         TabIndex        =   37
         Top             =   1491
         Width           =   915
      End
      Begin VB.Label lblNroEndPrincipal 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         TabIndex        =   36
         Top             =   1491
         Width           =   1755
      End
      Begin VB.Label lblEndPrincipal 
         Alignment       =   1  'Right Justify
         Caption         =   "Endereç0 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   3330
         TabIndex        =   35
         Top             =   1074
         Width           =   1155
      End
      Begin VB.Label lblCidEndPrincipal 
         Alignment       =   1  'Right Justify
         Caption         =   "Cidade :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         TabIndex        =   34
         Top             =   1908
         Width           =   1755
      End
      Begin VB.Label lblUFEndPrincipal 
         Alignment       =   1  'Right Justify
         Caption         =   "UF :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   6510
         TabIndex        =   33
         Top             =   1908
         Width           =   600
      End
      Begin VB.Image ImgPesqCli 
         Height          =   240
         Left            =   8190
         Picture         =   "frmCliente.frx":26D1
         Top             =   700
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblSite 
         Alignment       =   1  'Right Justify
         Caption         =   "Site :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         TabIndex        =   32
         Top             =   3159
         Width           =   1755
      End
      Begin VB.Label lblBairro 
         Alignment       =   1  'Right Justify
         Caption         =   "Bairro :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   6285
         TabIndex        =   31
         Top             =   1491
         Width           =   825
      End
      Begin VB.Label lblCliDesde 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente Desde :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         TabIndex        =   30
         Top             =   3993
         Width           =   1755
      End
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   1095
      Left            =   30
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cependprincipal As String
Dim cependcorresp As String
Dim cependcobr As String
Dim cependent As String
Dim valida_digito As String
Dim cdcontato As Double
Dim LinhaSelecionada As Double
Dim CodCliente As Double
Dim CodVendedor As Double
Dim Tipo As String

Public IncluirCliente As Integer

Private Function TrataTeclas(strSkeycode As Integer)
    KeyCode = strSkeycode
    FmeBotoes.Enabled = True
    
    Select Case KeyCode
        Case 27
            If cmd_Gravar.Enabled = True Then
                cmd_Gravar.SetFocus
                Exit Function
            ElseIf cmd_Adicionar.Enabled = True Then
                cmd_Adicionar.SetFocus
                Exit Function
            ElseIf cmd_Pesquisar.Enabled = True Then
                cmd_Pesquisar.SetFocus
                Exit Function
            ElseIf cmd_Limpar.Enabled = True Then
                cmd_Limpar.SetFocus
                Exit Function
            ElseIf cmd_Cancelar.Enabled = True Then
                cmd_Cancelar.SetFocus
                Exit Function
            ElseIf cmd_Excluir.Enabled = True Then
                cmd_Excluir.SetFocus
                Exit Function
            End If
    End Select

End Function

Private Sub AdicionarDados()
    'Call GetDados
    Tipo = "I"
    Liberacampo
    CodVendedor = 1
    
    cmd_Adicionar.Enabled = False
    cmd_Pesquisar.Enabled = False
    cmd_Gravar.Enabled = True
    txtCidEndPrincipal.Text = "Sï¿½o Paulo"
    txtUFEndPrincipal.Text = "SP"
    FmeClientes.Enabled = True
End Sub

Private Sub Excluir()
mensagem = MsgBox("Confirma realmente a EXCLUSï¿½O deste Cliente ? ", vbCritical + vbYesNo + vbDefaultButton2, "Responda-me")
            If mensagem = vbYes Then
                Tipo = "E"
                SQL_Registro
                LimpaCampo
            End If
End Sub

Private Sub Pesquisa()
    FmeClientes.Enabled = True
    Call GetDados
    Tipo = "A"
    ImgPesqCli.Visible = True
    cmd_Adicionar.Enabled = False
    cmd_Pesquisar.Enabled = False
    TxtRazaoSocial.Enabled = True
    TxtRazaoSocial.SetFocus
'*
'*** Fabio Reinert (Alemï¿½o) - 08/2017 - Liberar o CPF/CNPJ para pesquisa - inicio
'*
    txtCgc_Cpf.Enabled = True
'*
'*** Fabio Reinert (Alemï¿½o) - 08/2017 - Liberar o CPF/CNPJ para pesquisa - Fim
'*
    
End Sub

Private Sub Sair()
    
    If Consultas <> "S" Then
        'Fecha todos Recordsets
        If Consultas <> "S" Then
            If cmd_Sair.Enabled = True And Tipo = "A" Or Tipo = "I" Then
                mensagem = MsgBox("Informaï¿½ï¿½es nï¿½o foram gravadas. Deseja Sair assim mesmo ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
                If mensagem = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    Consultas = ""
    Unload Me
    
End Sub

Private Sub cmbRepresentante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MskCliDesde.SetFocus
End If
End Sub

Private Sub cmd_Adicionar_Click()
Call Too_Botoes(1)
lblcodCli.Caption = ""
MskCliDesde.Text = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub cmd_Excluir_Click()
Call Too_Botoes(4)
End Sub



Private Sub cmd_Gravar_Click()
Call Too_Botoes(5)
End Sub

Private Sub cmd_Limpar_Click()
Call Too_Botoes(3)
End Sub



Private Sub cmd_Pesquisar_Click()
Call Too_Botoes(2)
End Sub

Private Sub cmd_Sair_Click()
Call Too_Botoes(6)
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape And cmd_Gravar.Enabled = True Then
        mensagem = MsgBox("Salvar Dados ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
        If mensagem = vbNo Then
            If IncluirCliente = 1 Then
                Unload Me
                Exit Sub
            End If
        Else
            cmd_Gravar_Click
        Exit Sub
        End If
End If

If KeyCode = vbKeyF6 And cmd_Gravar.Enabled = True Then
        mensagem = MsgBox("Salvar Dados ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
        If mensagem = vbNo Then
            Exit Sub
        Else
        cmd_Gravar_Click
        Exit Sub
        End If
    ElseIf KeyCode = vbKeyF2 And cmd_Adicionar.Enabled = True Then
        cmd_Adicionar_Click
        Exit Sub
    ElseIf KeyCode = vbKeyF3 And cmd_Pesquisar.Enabled = True Then
        cmd_Pesquisar_Click
        Exit Sub
    ElseIf KeyCode = vbKeyF4 And cmd_Limpar.Enabled = True Then
        cmd_Limpar_Click
        Exit Sub
    ElseIf KeyCode = vbKeyF5 And cmd_Excluir.Enabled = True Then
        cmd_Excluir_Click
        Exit Sub
    ElseIf KeyCode = vbKeyF7 And cmd_Sair.Enabled = True Then
        cmd_Sair_Click
        Exit Sub
    ElseIf KeyCode = vbKeyEscape And cmd_Sair.Enabled = True Then
        mensagem = MsgBox("Informaï¿½ï¿½es nï¿½o Salvas. Deseja Sair assim mesmo ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
        If mensagem = vbNo Then
            TxtRazaoSocial.SetFocus
            Exit Sub
            Else
            Unload Me
        End If
    End If

End Sub

Private Sub Form_Load()
    
'    Me.Left = (frmMenu.Width - Me.Width) / 2
'    Me.Top = ((frmMenu.Height - Me.Height) / 2)
    
    KeyPreview = True
    Tipo = ""
    cependprincipal = ""
    cependcorresp = ""
    cependcobr = ""
    cependent = ""
    cdcontato = 0
    LinhaSelecionada = 0
    
    'Call Carrega_Colunas_Itens_Compra
    
    'para incluir cliente atraves do form Saidas
    If IncluirCliente = 1 Then
        Call AdicionarDados
    End If
    
    If Consultas = "S" Then
        cmd_Adicionar.Enabled = False
        cmd_Pesquisar.Enabled = False
        ImgPesqCli.Visible = True
        FmeClientes.Enabled = True
        TxtRazaoSocial.Enabled = True
'        TxtRazaoSocial.SetFocus
    End If
'*** Fabio Reinert - 08/2017 - Posiciona o listview dos clientes na posiï¿½ï¿½o do textbox do nome - Inicio
    lstRazaoSocial.Top = TxtRazaoSocial.Top
    lstRazaoSocial.Left = TxtRazaoSocial.Left
'*** Fabio Reinert - 08/2017 - Posiciona o listview dos clientes na posiï¿½ï¿½o do textbox do nome - Fim
'***
'*** Houve remanejamento dos campos na tela para acomodar o campo para telefone Celular
'*** Foi baixado o campo com a lista de compras do cliente
'***

End Sub

Private Function TeclasAtalhoTxt(strKeyCode As Integer)
On Error GoTo TeclasErroTeclasTxt
    
    KeyCode = strKeyCode
    
    If KeyCode = vbKeyEscape And Consultas = "S" Then
        Unload Me
        Exit Function
    End If

    If KeyCode = vbKeyF4 Or KeyCode = 27 Then
        If MsgBox("Salvar dados do Cliente..? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Call LimpaCampo
            Exit Function
        Else
            Call GravarDados
            Exit Function
        End If
    ElseIf KeyCode = vbKeyF5 And cmd_Excluir.Enabled = True Then
        If MsgBox("Deseja Excluir Cliente ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Salvar ?") = vbNo Then
            Exit Function
        Else
            Call Excluir
            Exit Function
        End If
    ElseIf KeyCode = vbKeyF6 And cmd_Gravar.Enabled = True Then '27 = esc
        If MsgBox("Salvar dados do Cliente..? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Function
        Else
            Call GravarDados
            Exit Function
        End If
    End If

Exit Function

TeclasErroTeclasTxt:
    Call ErrosGeraisLog(Now, Me.Name, "TeclasAtalhoTxt", Err.Description, Err.Number)
    Erro "Teclas de Atalho Txt"
    
End Function

Private Function AtalhoToollbar(strKeyCode As Integer)
On Error GoTo TeclasErroTeclasTxt
    
    KeyCode = strKeyCode
    
    If KeyCode = vbKeyF2 Then
        Call AdicionarDados
        Exit Function
    ElseIf KeyCode = vbKeyF3 Then
        Call Pesquisa
        Exit Function
    ElseIf KeyCode = vbKeyF4 Then
        If MsgBox("Deseja Limpar Dados ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Limpar ?") = vbNo Then
            Exit Function
         Else
            Call LimpaCampo
            Exit Function
         End If
    ElseIf KeyCode = vbKeyF5 And cmd_Excluir.Enabled = True Then
        If MsgBox("Deseja Excluir Cliente ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Excluir ?") = vbNo Then
            Exit Function
        Else
            Call Excluir
            Exit Function
        End If
    ElseIf KeyCode = vbKeyF6 And cmd_Gravar.Enabled = True Then '27 = esc
        If MsgBox("Salvar dados do Cliente..? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Function
        Else
            Call GravarDados
            Exit Function
        End If
    ElseIf KeyCode = vbKeyEscape Then
        Call Sair
        Exit Function
    End If

Exit Function

TeclasErroTeclasTxt:
    Call ErrosGeraisLog(Now, Me.Name, "AtalhoToollbar", Err.Description, Err.Number)
    Erro "Teclas de Atalho Txt"
    
End Function

Private Sub GetDados()
    'cmbRepresentante.Clear
    lstRazaoSocial.Clear
    
'    sql = "SELECT * FROM tab_clientes "
'    sql = sql & " ORDER BY RAZAO_SOCIAL"
'    Set Rstemp = New ADODB.Recordset
'    Rstemp.Open sql, Cnn, 1, 2
'    If Rstemp.RecordCount > 0 Then
'        Rstemp.MoveLast
'        Rstemp.MoveFirst
'        While Not Rstemp.EOF
'            cmbRepresentante.AddItem UCase(Rstemp!RAZAO_SOCIAL)
'            cmbRepresentante.ItemData(cmbRepresentante.NewIndex) = Rstemp(0)
'            Rstemp.MoveNext
'        Wend
'
'    End If
'    Rstemp.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'tipo = "A"
    'frmMenu.menu_cadastro_cliente.Enabled = True
    'Consultas = ""
    'IncluirCliente = 0
End Sub


Private Sub ImgPesqCli_Click()

    frmCliente.MousePointer = 11
       
    tipo_pesq = "N"
    sql = "SELECT * FROM tab_Clientes "
   ' sql = sql & " where razao_social >= '" & Trim(FiltraAspasSimples(TxtRazaoSocial.Text)) & "' "
   ' sql = sql & "AND razao_social <= '" & Trim(FiltraAspasSimples(TxtRazaoSocial.Text)) & "Z' "
'    sql = sql & " Where razao_social Like '%" & FiltraAspasSimples(Trim(TxtRazaoSocial.Text)) & "%'"
    If IsNumeric(TxtRazaoSocial.Text) Then
        sql = sql & " Where CODIGO = " & Me.TxtRazaoSocial.Text
    Else
        
        sql = sql & " Where RAZAO_SOCIAL Like '%" & FiltraAspasSimples(Trim(TxtRazaoSocial.Text)) & "%'"
    End If
    sql = sql & " order by RAZAO_SOCIAL  asc "
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open sql, Cnn, 1, 2
    If Rstemp.RecordCount > 0 Then
        lstRazaoSocial.Visible = True
        lstRazaoSocial.Clear
        lstRazaoSocial.SetFocus
        While Not Rstemp.EOF
            lstRazaoSocial.AddItem UCase(Rstemp!RAZAO_SOCIAL)
            lstRazaoSocial.ItemData(lstRazaoSocial.NewIndex) = Rstemp!Codigo
            Rstemp.MoveNext
        Wend
        
        lstRazaoSocial.Selected(0) = True
    Else
        MsgBox "Cliente nï¿½o Encontrado.", vbInformation, "Aviso"
    End If
    Rstemp.Close
    frmCliente.MousePointer = 0


End Sub

Private Sub Label24_Click()
Screen.MousePointer = 11
Label24.ForeColor = &H80000012

Dim Html As String
    
On Error GoTo TrataErro
    iexp = Environ("WINDIR") & "\explorer.exe  "
    Html = Shell(iexp & ("http://www.sintegra.gov.br"), vbMaximizedFocus)
    Screen.MousePointer = 1
Exit Sub

TrataErro:

If Err.Number <> 0 Then
    Err.Clear
    Screen.MousePointer = 1
End If
End Sub

Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Label24.ForeColor = &HFF0000
End Sub


Private Sub lstRazaoSocial_DblClick()
    If lstRazaoSocial.ListIndex <> -1 Then
        ImgPesqCli.Visible = False
        CodCliente = lstRazaoSocial.ItemData(lstRazaoSocial.ListIndex)
        strSql = "Select * from tab_Clientes WHERE Codigo = " & CodCliente
        lblcodCli.Caption = CodCliente
        Set Rstemp = New ADODB.Recordset
        Rstemp.Open strSql, Cnn, 1, 2
        strSql = ""

        Call MontaCampos
        lstRazaoSocial.Visible = False
        If Consultas <> "S" Then
            cmd_Excluir.Enabled = True
            cmd_Gravar.Enabled = True
            Liberacampo
            lstRazaoSocial.Visible = False
            mskCepEndPrincipal.SetFocus
        End If
    End If

End Sub

Private Sub lstRazaoSocial_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    lstRazaoSocial.Clear
    lstRazaoSocial.Visible = False
    TxtRazaoSocial.SetFocus
End If
End Sub


Private Sub lstRazaoSocial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call lstRazaoSocial_DblClick
End If
End Sub


Private Sub MskCliDesde_GotFocus()
Call SelText(MskCliDesde)
TrataMensagensTeclas
End Sub

Private Sub MskCliDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If MskCliDesde.Text <> "__/__/____" Then
            If Not IsDate(MskCliDesde.Text) Then
                MsgBox "Data Invï¿½lida.", vbInformation, Me.Caption
                MskCliDesde.Text = "__/__/____"
                MskCliDesde.SetFocus
                Exit Sub
            End If
        End If
        SendKeys "{tab}"
    End If
End Sub

Private Function Too_Botoes(mi_transacao As Integer)
    
    Select Case mi_transacao

        Case 1 'Novo
           
            Call AdicionarDados
            If IncluirCliente <> 1 Then
            txtCidEndPrincipal.Text = "São Paulo"
            txtUFEndPrincipal.Text = "SP"
            TxtRazaoSocial.SetFocus
    End If
           
        Case 2  'Pesquisa
                
            Call Pesquisa

        Case 3   'Limpar

            Call LimpaCampo

        Case 4   'Excluir
            
            Call Excluir
        
        Case 5   'Gravar
            
            Call GravarDados

        Case 6   'Sair

           Call Sair

    End Select

End Function

Private Sub GravarDados()
        If Len(Trim(TxtRazaoSocial.Text)) = 0 Then
            MsgBox "Obrigatï¿½rio informar Razao Social.", vbInformation, Me.Caption
            TxtRazaoSocial.SetFocus
            Exit Sub
        End If
        
        Call SQL_Registro
        
        If IncluirCliente = 1 Then
            frmSaidas.txtCodCliente.Text = CodCliente
            frmSaidas.txtcliente.Text = UCase(Trim(TxtRazaoSocial.Text))
            IncluirCliente = 0
            Unload Me
            frmSaidas.CarregaClienteCodigo
            Exit Sub
        ElseIf IncluirCliente = 2 Then
            frmCadPets.txtCodigoDono = CodCliente
            frmCadPets.txtDono.Text = UCase(Trim(TxtRazaoSocial.Text))
            IncluirCliente = 0
            Unload Me
            'frmSaidas.CarregaClienteCodigo
            Exit Sub

            LimpaCampo
        End If
End Sub

'*** Fabio Reinert - 09/2017 - Botoes de radio para caso de pendentes.
Private Sub OptMensal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtObsPrincipal.SetFocus
    End If
End Sub

Private Sub OptPagDia_Click()
    txtPagDia.Enabled = True
    txtPagDia.SetFocus
End Sub

Private Sub OptPagDia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPagDia.Enabled = True
        txtPagDia.SetFocus
    End If
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyUp Then
        txtPagDia.Enabled = False
        OptMensal.SetFocus
    End If
    If KeyAscii = vbKeyRight Or KeyAscii = vbKeyDown Then
        txtPagDia.Enabled = True
        txtPagDia.SetFocus
    End If

End Sub

Private Sub Scroll_Change()
    lblDiasEmAtraso.Caption = Scroll.Value
End Sub

Private Sub txtCelular_GotFocus()
    txtCelular = SemFormatoTel(txtCelular.Text)
    Call SelText(txtCelular)
    TrataMensagensTeclas
End Sub

Private Sub txtCelular_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        If txtCelular.Text = "" Then
            SendKeys "{tab}"
            Exit Sub
        End If
        
        If Len(txtCelular.Text) < 7 Then
            MsgBox "Nï¿½mero de Celular do Cliente invï¿½lido !", 64, "Aviso..."
            txtCelular.SetFocus
            Exit Sub
        Else
            txtCelular.Text = FormatCEL(txtCelular)
        End If
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtCelular_LostFocus()
    txtCelular.Text = FormatCEL(txtCelular)
End Sub

Private Sub txtCgc_Cpf_GotFocus()
Dim cont        As Integer
Dim strAux      As String

strAux = ""

txtCgc_Cpf.Text = SemFormatoCPF_CNPJ(txtCgc_Cpf.Text)

Call SelText(txtCgc_Cpf)
If Tipo = "A" Then
    'LblTeclas.Caption = "Digite o CPF/CNPJ a pesquisar e dï¿½ [ENTER]"
Else
    TrataMensagensTeclas
End If

End Sub


Private Sub txtCgc_Cpf_KeyPress(KeyAscii As Integer)
Dim CPF_CNPJ        As String
Dim CodKey          As Integer

CodKey = KeyAscii

If KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If

If CodKey = 13 Then
'*
'*** Fabio Reinert (Alemï¿½o) - 08/2017 - Alteraï¿½ï¿½o para pesquisa de cliente por CNPJ/CPF tambï¿½m - Inicio
'*
'    If txtCgc_Cpf.Text <> "" Then
'        CPF_CNPJ = SemFormatoCPF_CNPJ(txtCgc_Cpf.Text)
'        If Len(CPF_CNPJ) = 11 Then
'            If Not (ValidaCPF(CPF_CNPJ)) Then
'                If MsgBox("CPF Invï¿½lido !" & vbNewLine & "Deseja Prosseguir...?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                    txtCgc_Cpf.SetFocus
'                    Exit Sub
'                Else
'                    txtFone1Principal.SetFocus
'                End If
'            End If
'        ElseIf Len(CPF_CNPJ) = 14 Then
'            If Not (ValidaCGC(CPF_CNPJ)) Then
'                If MsgBox("CNPJ Invï¿½lido !" & vbNewLine & "Deseja Prosseguir...?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                    txtCgc_Cpf.SetFocus
'                    Exit Sub
'                Else
'                    txtFone1Principal.SetFocus
'                End If
'            End If
'        Else
'            If MsgBox("CNPJ Invï¿½lido !" & vbNewLine & "Deseja Prosseguir...?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                txtCgc_Cpf.SetFocus
'                Exit Sub
'            Else
'                SendKeys "{tab}"
'            End If
'        End If
'    Else
'        'MsgBox "O campo CPF/CNPJ, ï¿½ Obrigatï¿½rio !", 64, "Aviso"
'        'txtCgc_Cpf.SetFocus
'        'Exit Sub
'    End If
'
'    If Len(txtCgc_Cpf.Text) = 11 Then
'        txtCgc_Cpf.Text = FormatCPF_CNPJ(txtCgc_Cpf.Text)
'    ElseIf Len(txtCgc_Cpf.Text) = 14 Then
'        txtCgc_Cpf.Text = FormatCPF_CNPJ(txtCgc_Cpf.Text)
'    End If
'    txtFone1Principal.SetFocus
    
    If txtCgc_Cpf.Text = "" Then
        'mskCepEndPrincipal.SetFocus
        txtFone1Principal.SetFocus
        Exit Sub
    End If
    CPF_CNPJ = SemFormatoCPF_CNPJ(txtCgc_Cpf.Text)
    If Len(CPF_CNPJ) = 11 Then
        If Not (ValidaCPF(CPF_CNPJ)) Then
            If MsgBox("CPF Invï¿½lido !" & vbNewLine & "Deseja Prosseguir...?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                txtCgc_Cpf.SetFocus
                Exit Sub
            End If
        End If
    ElseIf Len(CPF_CNPJ) = 14 Then
        If Not (ValidaCGC(CPF_CNPJ)) Then
            If MsgBox("CNPJ Invï¿½lido !" & vbNewLine & "Deseja Prosseguir...?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                txtCgc_Cpf.SetFocus
                Exit Sub
            End If
        End If
    Else
        If MsgBox("CPF/CNPJ Invï¿½lido !" & vbNewLine & "Deseja Prosseguir...?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            txtCgc_Cpf.SetFocus
            Exit Sub
         End If
    End If
    If Tipo = "A" Then
        If Len(TxtRazaoSocial.Text) = 0 Then
            If Cnn.State = adStateOpen Then
                Cnn.Close
            End If
            Call sConectaBanco
            strSql = "Select * from Cliente WHERE CGC_CPF = '" & FormatCPF_CNPJ(CPF_CNPJ) & "'"
            'lblcodCli.Caption = CodCliente
            Set Rstemp = New ADODB.Recordset
            Rstemp.Open strSql, Cnn, 1, 2
            strSql = ""
            If Rstemp.RecordCount = 0 Then
                MsgBox "CPF/CNPJ nï¿½o encontrado!", vbOKOnly, "Aviso"
                txtCgc_Cpf.SetFocus
            Else
                Call MontaCampos
                Call Liberacampo
                lblcodCli.Caption = Rstemp!Codigo
                mskCepEndPrincipal.SetFocus
            End If
        Else
            txtCgc_Cpf.Text = FormatCPF_CNPJ(txtCgc_Cpf.Text)
            txtFone1Principal.SetFocus
        End If
        
    End If
'*
'*** Fabio Reinert (Alemï¿½o) - 08/2017 - Alteraï¿½ï¿½o para pesquisa de cliente por CNPJ/CPF tambï¿½m - Fim
'*

End If
End Sub


Private Sub txtCgc_Cpf_LostFocus()
Dim CPF_CNPJ        As String

If InStr(txtCgc_Cpf.Text, ".") = 0 Then
    If Len(txtCgc_Cpf.Text) = 11 Then
        txtCgc_Cpf.Text = FormatCPF_CNPJ(txtCgc_Cpf.Text)
    ElseIf Len(txtCgc_Cpf.Text) = 14 Then
        txtCgc_Cpf.Text = FormatCPF_CNPJ(txtCgc_Cpf.Text)
    End If
End If

End Sub


Private Sub txtContatoPrincipal_GotFocus()
    Call SelText(txtContatoPrincipal)
    TrataMensagensTeclas
End Sub


Private Sub txtContatoPrincipal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtObsPrincipal.SetFocus
    End If
End Sub

Private Sub txtemail_Change()
    txtemail.Text = LCase(txtemail.Text)
    txtemail.SelStart = Len(txtemail.Text)
End Sub


Private Sub txtFone1Principal_GotFocus()
txtFone1Principal = SemFormatoTel(txtFone1Principal.Text)
Call SelText(txtFone1Principal)
TrataMensagensTeclas
End Sub

Private Sub txtFone1Principal_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If

If KeyAscii = 13 Then
    If txtFone1Principal.Text = "" Then
    '    MsgBox "O campo telefone do Cliente, ï¿½ obrigatï¿½rio !", 64, "Aviso..."
        txtFone2Principal.SetFocus
        Exit Sub
    End If
    'ElseIf Len(txtFone1Principal.Text) < 7 Then
    If Len(txtFone1Principal.Text) < 7 Then
        MsgBox "Nï¿½mero de telefone do Cliente Invï¿½lido !", 64, "Aviso..."
        txtFone1Principal.SetFocus
        Exit Sub
    Else
        txtFone1Principal.Text = FormatTEL(txtFone1Principal)
    End If
        SendKeys "{tab}"
End If
  
End Sub

Private Sub txtFone1Principal_LostFocus()
txtFone1Principal.Text = FormatTEL(txtFone1Principal)
End Sub

Private Sub txtFone2Principal_GotFocus()
txtFone2Principal = SemFormatoTel(txtFone2Principal.Text)
Call SelText(txtFone2Principal)
TrataMensagensTeclas
End Sub

Private Sub txtFone2Principal_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If

If KeyAscii = 13 Then
    If txtFone2Principal.Text = "" Then
    '    MsgBox "O campo telefone do Cliente, ï¿½ obrigatï¿½rio !", 64, "Aviso..."
         SendKeys "{tab}"
        Exit Sub
    End If
    'ElseIf Len(txtFone2Principal.Text) < 7 Then
    If Len(txtFone2Principal.Text) < 7 Then
        MsgBox "Nï¿½mero de telefone do Cliente invï¿½lido !", 64, "Aviso..."
        txtFone2Principal.SetFocus
        Exit Sub
    Else
        txtFone2Principal.Text = FormatTEL(txtFone2Principal)
    End If

    SendKeys "{tab}"
End If
End Sub

Private Sub txtFone2Principal_LostFocus()
txtFone2Principal.Text = FormatTEL(txtFone2Principal)
End Sub

Private Sub txtObsPrincipal_GotFocus()
    Call SelText(txtObsPrincipal)
End Sub

Private Sub txtObsPrincipal_KeyDown(KeyCode As Integer, Shift As Integer)
'
'*** Fabio Reinert - 09/2017 - Processo alterado para inclusï¿½o dos clientes pendentes - Inicio
'***
'
    If KeyCode = vbKeyDown Then
        txtPrazo.SetFocus
    ElseIf KeyCode = vbKeyUp Then
        txtContatoPrincipal.SetFocus
    End If
'*** Fabio Reinert - 09/2017 - Processo alterado para inclusï¿½o dos clientes pendentes - Fim
'***

End Sub

Private Sub txtObsPrincipal_KeyPress(KeyAscii As Integer)
'
'*** Fabio Reinert - 09/2017 - Processo alterado para inclusï¿½o dos clientes pendentes - Inicio
'***
'    Char = Chr(KeyAscii)
'    KeyAscii = Asc(UCase(Char))
'    If KeyAscii = 13 Then
'        If MsgBox("Salvar Dados...?", vbQuestion + vbYesNo + vbDefaultButton1, "") = vbNo Then
'            Exit Sub
'            LimpaCampo
'        Else
'            cmd_Gravar_Click
'        End If
'    End If
'
    If KeyAscii = 13 Then
        txtPrazo.Enabled = True
        txtPrazo.SetFocus
    End If

'*** Fabio Reinert - 09/2017 - Processo alterado para inclusï¿½o dos clientes pendentes - Fim
'***
End Sub

Private Sub txtPagDia_GotFocus()
    Call SelText(txtPagDia)
End Sub

Private Sub txtPagDia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
        txtPagDia.Enabled = False
        txtPrazo.SetFocus
    End If
    If KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
        TxtRazaoSocial.SetFocus
    End If

End Sub

Private Sub txtPagDia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtPagDia.Text) = "" Then
            If MsgBox("Dia em branco, deseja prosseguir assim mesmo? ", vbYesNo, "Responda-me") = vbNo Then
                txtPagDia.SetFocus
                Exit Sub
            Else
                cmd_Gravar_Click
                Exit Sub
            End If
        Else
            If Val(txtPagDia.Text) < 1 Or Val(txtPagDia.Text) > 30 Then
                MsgBox "Dia invï¿½lido, favor informar um dia vï¿½lido", vbOKOnly, "aviso"
                txtPagDia.SetFocus
                Exit Sub
            End If
        End If
        cmd_Gravar_Click
    Else
        If KeyAscii = 46 Then
            KeyAscii = 44
        End If
        If KeyAscii = 44 <> 0 Then
            KeyAscii = 0
            Exit Sub
        End If
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
            KeyAscii = 0
        End If
    End If

    
End Sub

Private Sub txtPrazo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If Trim(txtPrazo.Text) = "" Then
            txtPagDia.Enabled = True
            txtPagDia.SetFocus
        Else
            cmd_Gravar_Click
        End If
    ElseIf KeyCode = vbKeyUp Then
        txtPagDia.Enabled = False
        txtObsPrincipal.SetFocus
    End If
End Sub

Private Sub txtPrazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtPrazo.Text) = "" Then
            txtPagDia.Enabled = True
            txtPagDia.SetFocus
        Else
            If Val(txtPrazo.Text) < 1 Or Val(txtPrazo.Text) > 30 Then
                If MsgBox("Dia invï¿½lido: " & txtPrazo.Text & ". deseja prosseguir? ", vbYesNo, "Responda-me") = vbNo Then
                    txtPagDia.SetFocus
                    Exit Sub
                End If
            End If
            txtPagDia.Text = ""     'aqui vai zerar o campo dia de pagamento pois entrou com dias de prazo
            cmd_Gravar_Click
        End If
    Else
        If KeyAscii = 46 Then
            KeyAscii = 44
        End If
        If KeyAscii = 44 <> 0 Then
            KeyAscii = 0
            Exit Sub
        End If
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtRazaoSocial_GotFocus()
    SelText TxtRazaoSocial
    If Tipo <> "I" Then
      '  LblTeclas.Caption = "Digite o texto a pesquisar e dï¿½ [Enter]/Pesquisa CPF/CNPJ vï¿½ atï¿½ o campo com o mouse ou [TAB] 2 vezes "
    End If
End Sub

Private Sub TrataMensagensTeclas()
   ' LblTeclas.Caption = ""
   ' LblTeclas.Caption = " [ F6 ] Salvar [ F4 ] Limpa dados Tela "
'    lblTeclas2.Caption = "[ esc ] Sair"
    
End Sub

Private Sub TxtRazaoSocial_KeyDown(KeyCode As Integer, Shift As Integer)
'Call TeclasAtalhoTxt(KeyCode)
End Sub

Private Sub Txtrazaosocial_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        
        If Tipo = "A" Or Consultas = "S" Then
            ImgPesqCli_Click
        Else
            mskCepEndPrincipal.SetFocus
        End If
    End If
End Sub

Private Sub mskcependprincipal_GotFocus()
    SelText mskCepEndPrincipal
    TrataMensagensTeclas
End Sub

Private Sub mskcependprincipal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If cependprincipal <> mskCepEndPrincipal.Text Then
'            Set dbCepBrasil = New ADODB.Connection
'            dbCepBrasil.CursorLocation = adUseServer
'            dbCepBrasil.Open "File Name=" & App.Path & "\CNN_CEP.udl;"
'
'            sql = "SELECT * FROM CEP_BRASIL where "
'            sql = sql & " CEP = '" & mskCepEndPrincipal.Text & "'"
'            sql = sql & " ORDER BY UF,CIDADE,LOGRADOURO,CEP"
'            Set Rstemp = New ADODB.Recordset
'            Rstemp.Open sql, dbCepBrasil, adOpenForwardOnly, adLockReadOnly
'            If Rstemp.RecordCount > 0 Then
'                txtEndPrincipal.Text = Trim(Rstemp!TIPO_LOGRADOURO & ". " & Rstemp!LOGRADOURO & ",")
'                If Not IsNull(Rstemp!Bairro) Then
'                    txtBaiEndPrincipal.Text = Rstemp!Bairro
'                End If
'                If Not IsNull(Rstemp!UF) Then
'                    txtUFEndPrincipal.Text = Rstemp!UF
'                End If
'                If Not IsNull(Rstemp!Cidade) Then
'                    txtCidEndPrincipal.Text = Format(Rstemp!Cidade)
'                End If
'            End If
'            Rstemp.Close
'
'            dbCepBrasil.Close
'            Set dbCepBrasil = Nothing
'        End If
        txtEndPrincipal.SetFocus
    End If
End Sub

Private Sub txtendprincipal_GotFocus()
    SelText txtEndPrincipal
    TrataMensagensTeclas
End Sub

Private Sub txtEndPrincipal_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        txtNroEndPrincipal.SetFocus
    End If
End Sub

Private Sub txtnroendprincipal_GotFocus()
    SelText txtNroEndPrincipal
    TrataMensagensTeclas
End Sub

Private Sub txtnroEndPrincipal_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        txtComplEndPrincipal.SetFocus
    End If
End Sub

Private Sub txtcomplendprincipal_GotFocus()
    SelText txtComplEndPrincipal
    TrataMensagensTeclas
End Sub

Private Sub txtComplEndPrincipal_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        txtBaiEndPrincipal.SetFocus
    End If
End Sub

Private Sub txtbaiendprincipal_GotFocus()
    SelText txtBaiEndPrincipal
    TrataMensagensTeclas
End Sub

Private Sub txtbaiendprincipal_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        txtCidEndPrincipal.SetFocus
    End If
End Sub

Private Sub txtcidendprincipal_GotFocus()
    SelText txtCidEndPrincipal
    TrataMensagensTeclas
End Sub

Private Sub txtcidendprincipal_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        txtUFEndPrincipal.SetFocus
    End If
End Sub

Private Sub txtufendprincipal_GotFocus()
    SelText txtUFEndPrincipal
End Sub

Private Sub txtUFendprincipal_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        Call sConectaBanco
        If Len(Trim(txtUFEndPrincipal.Text)) > 0 Then
            sql = "SELECT * FROM tab_Estados where Sigla = '" & txtUFEndPrincipal.Text & "'"
            Set Rstemp = New ADODB.Recordset
            Rstemp.Open sql, Cnn, 1, 2
            If Rstemp.RecordCount = 0 Then
                MsgBox "Estado invï¿½lido.", vbInformation, Me.Caption
                Rstemp.Close
                txtUFEndPrincipal.Text = ""
                txtUFEndPrincipal.SetFocus
                Exit Sub
            End If
            Rstemp.Close
        End If
        txtCgc_Cpf.SetFocus
    End If
End Sub


Private Sub txtinscest_GotFocus()
    SelText TxtInscest
    TrataMensagensTeclas
End Sub

Private Sub txtinscest_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtinscmunicipal_GotFocus()
    SelText txtInscMunicipal
    TrataMensagensTeclas
End Sub

Private Sub txtinscmunicipal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub


Private Sub txtsite_GotFocus()
    SelText txtSite
    TrataMensagensTeclas
End Sub

Private Sub txtsite_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(LCase(Char))
    If KeyAscii = 13 Then
        txtemail.SetFocus
    End If
End Sub

Private Sub txtemail_GotFocus()
Call SelText(txtemail)
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtemail.Text <> "" And InStr(txtemail.Text, "@") = 0 Then
            MsgBox "e-mail do Cliente invï¿½lido !", 64, "Aviso..."
            txtemail.SetFocus
            txtemail_GotFocus
            Exit Sub
        End If
        MskCliDesde.SetFocus
    End If
End Sub

Private Sub Liberacampo()
    FmeClientes.Enabled = True
    TxtRazaoSocial.Enabled = True
    mskCepEndPrincipal.Enabled = True
    txtEndPrincipal.Enabled = True
    txtNroEndPrincipal.Enabled = True
    txtComplEndPrincipal.Enabled = True
    txtBaiEndPrincipal.Enabled = True
    txtCidEndPrincipal.Enabled = True
    txtUFEndPrincipal.Enabled = True
    txtCgc_Cpf.Enabled = True
    txtFone1Principal.Enabled = True
    txtCelular.Enabled = True
    txtContatoPrincipal.Enabled = True
    TxtInscest.Enabled = True
    txtInscMunicipal.Enabled = True
    txtSite.Enabled = True
    txtemail.Enabled = True
   ' cmbRepresentante.Enabled = True
    MskCliDesde.Enabled = True
    chk_Bloqueado.Enabled = True
    txtObsPrincipal.Enabled = True
    chk_Bloqueado.Enabled = True
    '*** Fabio Reinert - 08/2017 - Inclusï¿½o do campo telefone celular - inicio
    txtCelular.Enabled = True
    '*** Fabio Reinert - 08/2017 - Inclusï¿½o do campo telefone celular - Fim
    '*** Fabio Reinert - 09/2017 - Inclusï¿½o dos campos para pendentes - Inicio
    txtPrazo.Enabled = True
    txtPagDia.Enabled = False     'ï¿½ enabled false e depende do campo prazo
    '*** Fabio Reinert - 09/2017 - Inclusï¿½o dos campos para pendentes - Fim
End Sub

Private Sub LimpaCampo()
    Tipo = ""
    ImgPesqCli.Visible = False
    
    'List_Itens_Compra.ListItems.Clear
    
    lstRazaoSocial.Clear
    lstRazaoSocial.Visible = False
   
    lblcodCli.Caption = ""
    'LblTeclas.Caption = ""
    cependprincipal = ""
    cependcorresp = ""
    cependcobr = ""
    cependent = ""
    cdcontato = 0
    LinhaSelecionada = 0
    IncluirCliente = 0

    chk_Bloqueado.Value = False
    
    FmeClientes.Enabled = False
    chk_Bloqueado.Enabled = False
    TxtRazaoSocial.Enabled = False
    mskCepEndPrincipal.Enabled = False
    txtEndPrincipal.Enabled = False
    txtNroEndPrincipal.Enabled = False
    txtComplEndPrincipal.Enabled = False
    txtBaiEndPrincipal.Enabled = False
    txtCidEndPrincipal.Enabled = False
    txtUFEndPrincipal.Enabled = False
    txtCgc_Cpf.Enabled = False
    MskCliDesde.Enabled = False
    TxtInscest.Enabled = False
    txtInscMunicipal.Enabled = False
    txtSite.Enabled = False
    txtemail.Enabled = False
    txtContatoPrincipal.Enabled = False
    txtFone1Principal.Enabled = False
    txtCelular.Enabled = False
    'cmbRepresentante.Enabled = False
    txtObsPrincipal.Enabled = False
    cmd_Sair.Enabled = True
    
    TxtRazaoSocial.Text = ""
    TxtRazaoSocial.Text = ""
    mskCepEndPrincipal.Text = "_____-___"
    txtEndPrincipal.Text = ""
    txtNroEndPrincipal.Text = ""
    txtComplEndPrincipal.Text = ""
    txtBaiEndPrincipal.Text = ""
    txtCidEndPrincipal.Text = ""
    txtUFEndPrincipal.Text = ""
    txtCgc_Cpf.Text = ""
    MskCliDesde.Text = "__/__/____"
    TxtInscest.Text = ""
    txtInscMunicipal.Text = ""
    txtSite.Text = ""
    txtemail.Text = ""
    txtFone1Principal.Text = ""
    txtFone2Principal.Text = ""
    txtContatoPrincipal.Text = ""
    txtObsPrincipal.Text = ""
    lblDiasEmAtraso.Caption = ""
    cmbRepresentante.Clear
    '*** Fabio Reinert - 08/2017 - Inclusï¿½o do campo telefone celular - Inicio
    txtCelular.Text = ""
    '*** Fabio Reinert - 08/2017 - Inclusï¿½o do campo telefone celular - Fim
    '*** Fabio Reinert - 09/2017 - Inclusï¿½o dos campos para pendentes - Inicio
    txtPrazo.Text = ""
    txtPagDia.Text = ""     'Sï¿½o enabled false e dependem do chkpendente
    '*** Fabio Reinert - 09/2017 - Inclusï¿½o dos campos para pendentes - Fim
    
    If Consultas <> "S" Then
        cmd_Adicionar.Enabled = True
        cmd_Pesquisar.Enabled = True
        cmd_Excluir.Enabled = False
        cmd_Gravar.Enabled = False
    Else
        ImgPesqCli.Visible = True
        FmeClientes.Enabled = True
        TxtRazaoSocial.Enabled = True
        TxtRazaoSocial.SetFocus
        Exit Sub
    End If
 
End Sub

Private Sub SQL_Registro()
    On Error GoTo TrataErros

    strSql = ""
    Select Case Tipo
        
        Case "I"
            NovoCodigo = Select_Max("Cliente", "CODIGO")
            CodCliente = NovoCodigo
            CodVendedor = 1
            strSql = "INSERT INTO Cliente VALUES ("
            strSql = strSql & CodCliente & ","
            strSql = strSql & "'" & FiltraAspasSimples(Trim(TxtRazaoSocial.Text)) & "',"
            If Len(Trim(mskCepEndPrincipal.Text)) = "_____-___" Then
                strSql = strSql & "null,"
            Else
                strSql = strSql & "'" & mskCepEndPrincipal.Text & "',"
            End If
            If Len(Trim(txtEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "'" & FiltraAspasSimples(txtEndPrincipal.Text) & "',"
            Else
                strSql = strSql & "null,"
            End If
            If Len(Trim(txtNroEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "'" & FiltraAspasSimples(txtNroEndPrincipal.Text) & "',"
            Else
                strSql = strSql & "null,"
            End If
            If Len(Trim(txtComplEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "'" & FiltraAspasSimples(txtComplEndPrincipal.Text) & "',"
            Else
                strSql = strSql & "null,"
            End If
            If Len(Trim(txtBaiEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "'" & FiltraAspasSimples(txtBaiEndPrincipal.Text) & "',"
            Else
                strSql = strSql & "null,"
            End If
            If Len(Trim(txtCidEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "'" & FiltraAspasSimples(txtCidEndPrincipal.Text) & "',"
            Else
                strSql = strSql & "null,"
            End If
            If Len(Trim(txtUFEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "'" & txtUFEndPrincipal.Text & "',"
            Else
                strSql = strSql & "null,"
            End If
            If txtCgc_Cpf.Text <> "__.___.___/____-__" Then
                strSql = strSql & "'" & txtCgc_Cpf.Text & "',"
            Else
                strSql = strSql & "null,"
            End If
            If Len(Trim(TxtInscest.Text)) <> 0 Then
                strSql = strSql & "'" & UCase(TxtInscest.Text) & "',"
            Else
                strSql = strSql & "null,"
            End If
            If Len(Trim(txtInscMunicipal.Text)) <> 0 Then
                strSql = strSql & "'" & txtInscMunicipal.Text & "',"
            Else
                strSql = strSql & "null,"
            End If
            If Len(Trim(txtSite.Text)) <> 0 Then
                strSql = strSql & "'" & txtSite.Text & "',"
            Else
                strSql = strSql & "null,"
            End If
            If Len(Trim(txtemail.Text)) <> 0 Then
                strSql = strSql & "'" & txtemail.Text & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            If Len(Trim(txtFone1Principal.Text)) <> 0 Then
                strSql = strSql & "'" & txtFone1Principal.Text & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            If Len(Trim(txtFone2Principal.Text)) <> 0 Then
                strSql = strSql & "'" & txtFone2Principal.Text & "',"
            Else
                strSql = strSql & "null,"
            End If
                        

            '*** Fabio Reinert - 08/2017 - inclusï¿½o do campo: telefone celular - Fim
            
            If Len(Trim(txtContatoPrincipal.Text)) <> 0 Then
                strSql = strSql & "'" & Len(txtContatoPrincipal.Text) & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            If Len(Trim(txtObsPrincipal.Text)) <> 0 Then
                strSql = strSql & "'" & txtObsPrincipal.Text & "',"
            Else
                strSql = strSql & "null,"
            End If
            
             If MskCliDesde.Text <> "__/__/____" Then
                strSql = strSql & "'" & Format(MskCliDesde.Text, "mm/dd/yyyy") & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            'strSql = strSql & cmbRepresentante.ItemData(cmbRepresentante.ListIndex) & ","
            strSql = strSql & CodVendedor & ","
            If lblDiasEmAtraso.Caption <> "" And Val(lblDiasEmAtraso.Caption) <> 0 Then
                strSql = strSql & lblDiasEmAtraso.Caption & ","
            Else
                strSql = strSql & "null,"
            End If
            strSql = strSql & IIf(chk_Bloqueado = 1, "'S'", "'N'") & ","
                        
            '*** Fabio Reinert - 08/2017 - inclusï¿½o do campo: telefone celular - Inicio
            If Len(Trim(txtCelular.Text)) <> 0 Then
                strSql = strSql & "'" & txtCelular.Text & "',"
            Else
                strSql = strSql & "null,"
            End If
            '*** Fabio Reinert - 08/2017 - inclusï¿½o do campo: telefone celular - Fim
            
            '*** Fabio Reinert - 09/2017 - inclusï¿½o do campo: Prazo de pagamento pendente - Inicio
            If Len(Trim(txtPrazo.Text)) <> 0 Then
                strSql = strSql & "'" & txtPrazo.Text & "',"
            Else
                strSql = strSql & "null,"
            End If
            '*** Fabio Reinert - 09/2017 - inclusï¿½o do campo: Prazo de pagamento pendente - Fim
            
            '*** Fabio Reinert - 09/2017 - inclusï¿½o do campo: Dia  para pagamento pendente - Inicio
            If Len(Trim(txtPagDia.Text)) <> 0 Then
                strSql = strSql & "'" & txtPagDia.Text & "')"
            Else
                strSql = strSql & "null)"
            End If
            '*** Fabio Reinert - 09/2017 - inclusï¿½o do campo: Dia  para pagamento pendente - Fim
            
            Cnn.Execute strSql
            Cnn.CommitTrans
        
        Case "A"
            
            CodVendedor = 1
            strSql = ""
            strSql = " update cliente set "
            
            strSql = strSql & "RAZAO_SOCIAL = '" & FiltraAspasSimples(Trim(TxtRazaoSocial.Text)) & "',"
            
            If Len(Trim(mskCepEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "CEP_PRINCIPAL = '" & mskCepEndPrincipal & "',"
            Else
                strSql = strSql & "CEP_PRINCIPAL = null,"
            End If
            
            If Len(Trim(txtEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "ENDERECO_PRINCIPAL = '" & FiltraAspasSimples(txtEndPrincipal.Text) & "',"
            Else
                strSql = strSql & "ENDERECO_PRINCIPAL = null,"
            End If
            
            If Len(Trim(txtNroEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "NRO_END_PRINCIPAL = '" & FiltraAspasSimples(txtNroEndPrincipal.Text) & "',"
            Else
                strSql = strSql & "NRO_END_PRINCIPAL = null,"
            End If
            
            If Len(Trim(txtComplEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "COMPL_END_PRINCIPAL = '" & FiltraAspasSimples(txtComplEndPrincipal.Text) & "',"
            Else
                strSql = strSql & "COMPL_END_PRINCIPAL = null,"
            End If
            
            If Len(Trim(txtBaiEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "BAIRRO_END_PRINCIPAL = '" & FiltraAspasSimples(txtBaiEndPrincipal.Text) & "',"
            Else
                strSql = strSql & "BAIRRO_END_PRINCIPAL = null,"
            End If
            
            If Len(Trim(txtCidEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "CIDADE_END_PRINCIPAL = '" & FiltraAspasSimples(txtCidEndPrincipal.Text) & "',"
            Else
                strSql = strSql & "CIDADE_END_PRINCIPAL = null,"
            End If
            
            If Len(Trim(txtUFEndPrincipal.Text)) <> 0 Then
                strSql = strSql & "UF_END_PRINCIPAL = '" & txtUFEndPrincipal.Text & "',"
            Else
                strSql = strSql & "UF_END_PRINCIPAL = null,"
            End If
            
            If Len(Trim(TxtInscest.Text)) <> 0 Then
                strSql = strSql & "INSC_ESTADUAL = '" & UCase(TxtInscest.Text) & "',"
            Else
                strSql = strSql & "INSC_ESTADUAL = null,"
            End If
            
            If Len(Trim(txtInscMunicipal.Text)) <> 0 Then
                strSql = strSql & "INSC_MUNICIPAL = '" & txtInscMunicipal.Text & "',"
            Else
                strSql = strSql & "INSC_MUNICIPAL = null,"
            End If
            
            If Len(Trim(txtCgc_Cpf.Text)) <> 0 Then
                strSql = strSql & "Cgc_Cpf = '" & txtCgc_Cpf.Text & "',"
            Else
                strSql = strSql & "Cgc_Cpf = null,"
            End If
            
            If Len(Trim(txtSite.Text)) <> 0 Then
                strSql = strSql & "SITE = '" & txtSite.Text & "',"
            Else
                strSql = strSql & "SITE = null,"
            End If
            
            If Len(Trim(txtemail.Text)) <> 0 Then
                strSql = strSql & "EMAIL = '" & txtemail.Text & "',"
            Else
                strSql = strSql & "EMAIL = null,"
            End If
            
            If Len(Trim(txtFone1Principal.Text)) <> 0 Then
                strSql = strSql & "fone1 = '" & txtFone1Principal.Text & "',"
            Else
                strSql = strSql & "fone1 = null,"
            End If
            
            If Len(Trim(txtFone2Principal.Text)) <> 0 Then
                strSql = strSql & "fone2 = '" & txtFone2Principal.Text & "',"
            Else
                strSql = strSql & "fone2 = null,"
            End If
            
            If Len(Trim(txtContatoPrincipal.Text)) <> 0 Then
                strSql = strSql & "contato = '" & txtContatoPrincipal.Text & "',"
            Else
                strSql = strSql & "contato = null,"
            End If

            If Len(Trim(cmbRepresentante.Text)) <> 0 Then
                strSql = strSql & "COD_REPRESENTANTE = " & CodVendedor & ","
            Else
                strSql = strSql & "COD_REPRESENTANTE = null,"
            End If
            
            If lblDiasEmAtraso.Caption <> "" And Val(lblDiasEmAtraso.Caption) <> 0 Then
                strSql = strSql & "DIAS_ATRASO = " & lblDiasEmAtraso.Caption & ","
            Else
                strSql = strSql & "DIAS_ATRASO = null,"
            End If
            
            If MskCliDesde.Text <> "__/__/____" Then
                strSql = strSql & "CLIENTE_DESDE = '" & Format(MskCliDesde.Text, "mm/dd/yyyy") & "',"
            Else
                strSql = strSql & "CLIENTE_DESDE = null,"
            End If
            
            If Len(txtObsPrincipal.Text) <> 0 Then
                strSql = strSql & "OBS = '" & txtObsPrincipal.Text & "',"
            Else
                strSql = strSql & "OBS = null,"
            End If
            
            strSql = strSql & "bloqueado = " & IIf(chk_Bloqueado = 1, "'S'", "'N'")
            
            '*** Fabio Reinert - 08/2017 - Inclusï¿½o do campo: prazo pagamento pendente - Inicio
            If Len(Trim(txtPrazo.Text)) <> 0 Then
                strSql = strSql & ",prazopend = '" & txtPrazo.Text & "',"
            Else
                strSql = strSql & ",prazopend = null,"
            End If
            '*** Fabio Reinert - 08/2017 - Inclusï¿½o do campo: prazo pagamento pendente - Fim
            
            '*** Fabio Reinert - 08/2017 - Inclusï¿½o do campo: dia pagamento pendente - Inicio
            If Len(Trim(txtPagDia.Text)) <> 0 Then
                strSql = strSql & "pagtodia = '" & txtPagDia.Text & "'"
            Else
                strSql = strSql & "pagtodia = null"
            End If
            '*** Fabio Reinert - 08/2017 - Inclusï¿½o do campo: dia pagamento pendente - Fim
            
            strSql = strSql & " where CODIGO = " & CodCliente
            
            Cnn.Execute strSql
            Cnn.CommitTrans
        
        
        Case "E"
        
            strSql = ""
            strSql = "DELETE from Cliente "
            strSql = strSql & " where CODIGO = " & CodCliente
            Cnn.Execute strSql
            Cnn.CommitTrans
    End Select

    cmd_Sair.Enabled = False

    Exit Sub
    
TrataErros:
    Call ErrosGeraisLog(Now, Me.Name, "SQL_Registro", Err.Description, Err.Number)
    'Erro " sqlRegistro "
    
End Sub

Private Sub MontaCampos()
On Error GoTo TrataerroMontacampos

    If Rstemp.RecordCount <> 0 Then
        
        S_chk_Bloqueado = IIf(IsNull(Rstemp("Bloqueado")), "N", Rstemp("Bloqueado"))
        chk_Bloqueado.Value = IIf(Rstemp("Bloqueado") = "S", 1, 0)
       
        If Not IsNull(Rstemp!DIAS_ATRASO) Then
           lblDiasEmAtraso.Caption = Rstemp!DIAS_ATRASO
            Scroll.Value = Rstemp!DIAS_ATRASO
        Else
            Scroll.Value = 0
        End If
        
        TxtRazaoSocial.Text = UCase(Rstemp!RAZAO_SOCIAL)
        
        If Not IsNull(Rstemp!CEP_PRINCIPAL) Then
            mskCepEndPrincipal.Text = Rstemp!CEP_PRINCIPAL
        End If
        If Not IsNull(Rstemp!ENDERECO_PRINCIPAL) Then
            txtEndPrincipal.Text = UCase(Rstemp!ENDERECO_PRINCIPAL)
        End If
        If Not IsNull(Rstemp!NRO_END_PRINCIPAL) Then
            txtNroEndPrincipal.Text = Rstemp!NRO_END_PRINCIPAL
        End If
        If Not IsNull(Rstemp!COMPL_END_PRINCIPAL) Then
            txtComplEndPrincipal.Text = Rstemp!COMPL_END_PRINCIPAL
        End If
        If Not IsNull(Rstemp!BAIRRO_END_PRINCIPAL) Then
            txtBaiEndPrincipal.Text = UCase(Rstemp!BAIRRO_END_PRINCIPAL)
        End If
        If Not IsNull(Rstemp!CIDADE_END_PRINCIPAL) Then
            txtCidEndPrincipal.Text = UCase(Rstemp!CIDADE_END_PRINCIPAL)
        End If
        If Not IsNull(Rstemp!UF_END_PRINCIPAL) Then
            txtUFEndPrincipal.Text = Rstemp!UF_END_PRINCIPAL
        End If
        If Not IsNull(Rstemp!CGC_CPF) Then
            txtCgc_Cpf.Text = Rstemp!CGC_CPF
        End If
        If Not IsNull(Rstemp!FONE1) Then
            txtFone1Principal.Text = Rstemp!FONE1
        End If
        If Not IsNull(Rstemp!fone2) Then
            txtFone2Principal.Text = Rstemp!fone2
        End If
        
        '*** Fabio Reinert - 08/2017 - Inclusao do campo: Telefone Celular - Inicio
        If Not IsNull(Rstemp!CELULAR) Then
            txtCelular.Text = Rstemp!CELULAR
        End If
        '*** Fabio Reinert - 08/2017 - Inclusao do campo: Telefone Celular - Fim
        
        If Not IsNull(Rstemp!contato) Then
            txtContatoPrincipal.Text = UCase(Rstemp!contato)
        End If
        If Not IsNull(Rstemp!INSC_ESTADUAL) Then
            TxtInscest.Text = Rstemp!INSC_ESTADUAL
        End If
        If Not IsNull(Rstemp!INSC_MUNICIPAL) Then
            txtInscMunicipal.Text = Rstemp!INSC_MUNICIPAL
        End If
        If Not IsNull(Rstemp!SITE) Then
            txtSite.Text = Rstemp!SITE
        End If
        If Not IsNull(Rstemp!EMAIL) Then
            txtemail.Text = Rstemp!EMAIL
        End If
        If Not IsNull(Rstemp!CLIENTE_DESDE) Then
            MskCliDesde.Text = Format(Rstemp!CLIENTE_DESDE, "dd/mm/yyyy")
        End If
'
'*** Fabio Reinert - 09/2017 - Inclusï¿½o dos campos: prazo pagto e Pagto.dia - Inicio
'
        If Not IsNull(Rstemp!PRAZOPEND) Then
            txtPrazo.Text = Rstemp!PRAZOPEND
        End If
        
        If Not IsNull(Rstemp!PAGTODIA) Then   'Se tem dia de pagamento entï¿½o
            txtPagDia.Text = Rstemp!PAGTODIA  ' e coloca o conteudo do dia na tela
        End If
'
'*** Fabio Reinert - 09/2017 - Inclusï¿½o dos campos: Tipo de cliente e Pagto.dia - Fim
'
        
        If Not IsNull(Rstemp!obs) Then
            txtObsPrincipal.Text = UCase(Rstemp!obs)
        End If
    
        If Not IsNull(Rstemp!COD_REPRESENTANTE) Then
            sql = ""
            sql = sql & "SELECT * FROM Representante where Codigo = " & Rstemp!COD_REPRESENTANTE
            Set RsTemp1 = New ADODB.Recordset
            RsTemp1.Open sql, Cnn, 1, 2
            If RsTemp1.RecordCount > 0 Then
                For I = 0 To cmbRepresentante.ListCount - 1
                    If cmbRepresentante.ItemData(I) = Rstemp!COD_REPRESENTANTE Then
                        cmbRepresentante.ListIndex = I
                        Exit For
                    End If
                Next
             End If
             RsTemp1.Close
        End If
        
        'compras
        sql = "SELECT V.SEQUENCIA, V.DATA_NF, V.TOTAL_SAIDA FROM "
        sql = sql & " CLIENTE C LEFT JOIN SAIDAS_PRODUTO V ON C.CODIGO = V.CODIGO_CLIENTE  "
        sql = sql & " WHERE C.CODIGO = " & CodCliente
        sql = sql & " AND V.status_saida is not null "
        sql = sql & " AND EXTRACT(year FROM(V.DATA_NF)) = " & Year(Date)
        sql = sql & " ORDER BY V.DATA_NF"
        Set Rstemp5 = New ADODB.Recordset
        Rstemp5.Open sql, Cnn, 1, 2
        If Rstemp5.RecordCount > 0 Then
            For X = 1 To Rstemp5.RecordCount
                List_Itens_Compra.ListItems.Add X, , Rstemp5!SEQUENCIA
                List_Itens_Compra.ListItems(X).SubItems(1) = Format(Rstemp5!Data_NF, "dd/mm/yyyy")
                List_Itens_Compra.ListItems(X).SubItems(2) = Format(Rstemp5!TOTAL_SAIDA, "###,##0.00")
                total = Format(total + Rstemp5!TOTAL_SAIDA, "###,##0.00")
                Rstemp5.MoveNext
            Next X
        End If
        Rstemp5.Close
        Set Rstemp5 = Nothing
    End If

Exit Sub

TrataerroMontacampos:
If Err.Number <> 0 Then
    Call ErrosGeraisLog(Now, Me.Name, "MontaCampos", Err.Description, Err.Number)
    Err.Clear
End If
End Sub

Private Sub VScroll1_Change()

End Sub


