VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCdPgto 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forma de Pagamento"
   ClientHeight    =   8490
   ClientLeft      =   6180
   ClientTop       =   2985
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCdPgto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel_BOLETO 
      Height          =   6735
      Left            =   120
      TabIndex        =   33
      Top             =   1200
      Visible         =   0   'False
      Width           =   7215
      _Version        =   65536
      _ExtentX        =   12726
      _ExtentY        =   11880
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.Frame Frame4 
         Caption         =   "Nro. Nota Fiscal "
         Height          =   855
         Left            =   600
         TabIndex        =   47
         Top             =   120
         Width           =   5895
         Begin VB.TextBox txtNroNF 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   120
            MaxLength       =   9
            TabIndex        =   48
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label lbl_mensagem_grupo 
            Caption         =   "Digite 0 caso n�o emita Nota Fiscal."
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   1800
            TabIndex        =   49
            Top             =   400
            Width           =   3555
         End
      End
      Begin VB.TextBox txtVlr_Parcela 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   360
         Left            =   2385
         MaxLength       =   10
         TabIndex        =   39
         Top             =   1470
         Width           =   950
      End
      Begin VB.Frame Frame3 
         Height          =   4095
         Left            =   600
         TabIndex        =   43
         Top             =   1920
         Width           =   5895
         Begin MSComctlLib.ListView LIST_ITENS_FORMA_PGTO 
            Height          =   3735
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   5610
            _ExtentX        =   9895
            _ExtentY        =   6588
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FlatScrollBar   =   -1  'True
            FullRowSelect   =   -1  'True
            PictureAlignment=   4
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.TextBox txtIntervalo 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5145
         MaxLength       =   3
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txtQtParc 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2385
         MaxLength       =   2
         TabIndex        =   37
         Top             =   1080
         Width           =   630
      End
      Begin MSMask.MaskEdBox MskVcto 
         Height          =   360
         Left            =   5145
         TabIndex        =   41
         Top             =   1470
         Width           =   1155
         _ExtentX        =   2037
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
      Begin VB.CommandButton  cmd_Limpar 
         Height          =   435
         Left            =   6600
         TabIndex        =   46
         ToolTipText     =   "Limpa todos os campos"
         Top             =   240
         Width           =   375
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1      End
      Begin VB.Label lblTTCompra 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Parcela :"
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   600
         TabIndex        =   45
         Top             =   1470
         Width           =   1740
      End
      Begin VB.Label lbl_intervalo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Intervalo. Vcto :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   3360
         TabIndex        =   42
         Top             =   1080
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label lbl_qtparce 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Qtde Parcelas :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   600
         TabIndex        =   40
         Top             =   1080
         Width           =   1740
      End
      Begin VB.Label Label26 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Compra :"
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   0
         Left            =   630
         TabIndex        =   36
         Top             =   6210
         Width           =   1470
      End
      Begin VB.Label lblTotCompra 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   2160
         TabIndex        =   35
         Top             =   6210
         Width           =   975
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vencimento :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   3360
         TabIndex        =   34
         Top             =   1470
         Width           =   1740
      End
   End
   Begin VB.CommandButton cmd_Gravar 
      Caption         =   "&Gravar"
      Height          =   375
      Left            =   5880
      Picture         =   "frmCdPgto.frx":0464
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7005
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame SSFrame1 
      Height          =   7365
      Left            =   15
      TabIndex        =   8
      Top             =   960
      Width           =   7470
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Frame fme_Totais 
         Caption         =   "Totais"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   2175
         Left            =   3720
         TabIndex        =   23
         Top             =   3840
         Width           =   3615
         Begin VB.TextBox txt_Desconto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2040
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   780
            Width           =   1440
         End
         Begin VB.Label lbl_Sub_Tot_Pedido 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   2040
            TabIndex        =   30
            Top             =   360
            Width           =   1440
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sub Total Pedido :"
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1830
         End
         Begin VB.Label lblTroco 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2040
            TabIndex        =   28
            Top             =   1620
            Width           =   1440
         End
         Begin VB.Label LabelTroco 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   120
            TabIndex        =   27
            Top             =   1620
            Width           =   1830
         End
         Begin VB.Label lblTotPedido 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   2040
            TabIndex        =   26
            Top             =   1200
            Width           =   1440
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Valor a Receber :"
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   1830
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Desconto :"
            ForeColor       =   &H00000080&
            Height          =   360
            Left            =   120
            TabIndex        =   24
            Top             =   780
            Width           =   1830
         End
      End
      Begin VB.TextBox txtCodSeq 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   5730
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.Frame fmeListaFormaPgto 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3375
         Begin MSComctlLib.ListView LIST_DETALHESPGTO 
            Height          =   4815
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   8493
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FlatScrollBar   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Formas de Pagamento"
         ForeColor       =   &H00800000&
         Height          =   2895
         Left            =   3720
         TabIndex        =   9
         Top             =   855
         Width           =   3615
         Begin VB.TextBox txtVlrCartLoja 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2040
            TabIndex        =   5
            Text            =   "0.00"
            Top             =   2040
            Width           =   1440
         End
         Begin VB.TextBox txtVlrdinheiro 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2040
            TabIndex        =   21
            Text            =   "0.00"
            Top             =   300
            Width           =   1440
         End
         Begin VB.TextBox txtPendente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2040
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   2460
            Width           =   1440
         End
         Begin VB.TextBox txtVlrCartEletron 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2040
            TabIndex        =   4
            Text            =   "0.00"
            Top             =   1560
            Width           =   1440
         End
         Begin VB.TextBox txtVlrCartCredito 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2040
            TabIndex        =   3
            Text            =   "0.00"
            Top             =   1140
            Width           =   1440
         End
         Begin VB.TextBox txtVlrCheques 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   2040
            TabIndex        =   2
            Text            =   "0.00"
            Top             =   720
            Width           =   1440
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cart. Loja :"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   50
            Top             =   2040
            Width           =   1830
         End
         Begin VB.Label lblDinheiro 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Dinheiro :"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   22
            Top             =   300
            Width           =   1830
         End
         Begin VB.Label lblCheques 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cheque(s) :"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1830
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pendente :"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   12
            Top             =   2460
            Width           =   1830
         End
         Begin VB.Label lblCart_Credito 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cart. Cr�dito :"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   11
            Top             =   1140
            Width           =   1830
         End
         Begin VB.Label lblCart_Eletron 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cart. Eletr�nico :"
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   120
            TabIndex        =   10
            Top             =   1560
            Width           =   1830
         End
      End
      Begin VB.Frame SSFrame2 
         Height          =   1110
         Left            =   120
         TabIndex        =   14
         Top             =   6120
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
         _ExtentY        =   1958
         _StockProps     =   14
         Caption         =   "Tecla de Atalho"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label lblAjuda2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   480
            TabIndex        =   16
            Top             =   720
            Width           =   6165
         End
         Begin VB.Label lblTeclas 
            Alignment       =   2  'Center
            Caption         =   "Digite o n�mero do  Pedido e tecle [  Enter ]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   645
            TabIndex        =   15
            Top             =   480
            Width           =   5925
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   31
         Top             =   5400
         Width           =   3375
         Begin VB.CheckBox chk_Imp_Forma_Pgto 
            Caption         =   "Imprimir forma Pgto "
            ForeColor       =   &H00000080&
            Height          =   240
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   2535
         End
      End
      Begin VB.Label lblped 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N�  Pedido :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   3720
         TabIndex        =   17
         Top             =   360
         Width           =   1950
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Recebimento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   240
      Width           =   7455
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmCdPgto.frx":20B6
      Top             =   45
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmCdPgto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalDinheiro As Double
Dim TotalCheque As Double
Dim TotalCartCredito As Double
Dim TotalCartEletro As Double
Dim TotalPedido As Double
Dim TotalRecebido As Double
Dim ItemChecado As Boolean
Dim Msg As String
Dim resposta As String
Dim TotPedido As Double
Dim contPedido As Double
Dim CodSeq As Double
Dim CodCliente As Double
Dim TotalPendente As Double
Dim Total_Desconto As Double
Dim flag_Imp_Forma_Pgto As Boolean

Public strNumped As String
Public str_cnpj_cpf As String


Dim cont As Integer
Dim lvListItems_Itens As MSComctlLib.ListItem
Dim ItemSelecionado As Integer 'ListView
Dim LinhaSelecionada As Integer 'ListView

Dim Total_Parcelas As Double

Dim conta_Insert As Integer 'controlar o n� de insert ou teclas enter
Dim flag_altera_forma_pgto_Boleto As Boolean

Dim total_Troco As String
Dim TotalCartLoja As Double



Public Sub CalcValRecebidos()
    
 On Error GoTo trataAqui
    
    TotalRecebido = 0
    TotalDinheiro = 0
    TotalCheque = 0
    TotalCartCredito = 0
    TotalCartEletro = 0
    TotalPendente = 0
    Total_Desconto = 0
    lblTroco.Caption = ""
    TotalCartLoja = 0
    
    lblTotPedido.Caption = Format(CCur(lbl_Sub_Tot_Pedido.Caption) - CCur(txt_Desconto.Text), "0.00")
    
    TotalPedido = Format(lbl_Sub_Tot_Pedido.Caption, "#,##0.00;(#,##0.00")

    If IsNumeric(txtVlrdinheiro.Text) = True Then
        txtVlrdinheiro.Text = Format(txtVlrdinheiro.Text, "#,##0.00;(#,##0.00")
        TotalDinheiro = txtVlrdinheiro.Text
    End If
    
    If IsNumeric(txtVlrCheques.Text) = True Then
        txtVlrCheques.Text = Format(txtVlrCheques.Text, "#,##0.00;(#,##0.00")
        TotalCheque = txtVlrCheques.Text
    End If
    
    If IsNumeric(txtVlrCartCredito.Text) = True Then
        txtVlrCartCredito.Text = Format(txtVlrCartCredito.Text, "#,##0.00;(#,##0.00")
        TotalCartCredito = txtVlrCartCredito.Text
    End If
    
    If IsNumeric(txtVlrCartEletron.Text) = True Then
         txtVlrCartEletron.Text = Format(txtVlrCartEletron.Text, "#,##0.00;(#,##0.00")
         TotalCartEletro = txtVlrCartEletron.Text
    End If
    
    If IsNumeric(txtPendente.Text) = True Then
        txtPendente.Text = Format(txtPendente.Text, "0.00")
        TotalPendente = txtPendente.Text
    End If
    
    If IsNumeric(txtVlrCartLoja.Text) = True Then
         txtVlrCartLoja.Text = Format(txtVlrCartLoja.Text, "#,##0.00;(#,##0.00")
         TotalCartLoja = txtVlrCartLoja.Text
    End If
    
    If IsNumeric(txt_Desconto.Text) = True Then
        txt_Desconto.Text = Format(txt_Desconto.Text, "0.00")
        Total_Desconto = txt_Desconto.Text
        'txt_Desconto.SetFocus
    End If
    
    TotalRecebido = (TotalDinheiro + TotalCheque + TotalCartCredito + TotalCartEletro + TotalPendente + TotalCartLoja)
        
    lblTroco.Caption = TotalRecebido - (TotalPedido - Total_Desconto)
    
    If TotalRecebido < CCur(lblTotPedido) Then
        LabelTroco.ForeColor = &H80&
        lblTroco.ForeColor = &H80&
        LabelTroco.Caption = "Falta :"
        lblTroco.Caption = Format(lblTroco, "0.00")
        Exit Sub
    ElseIf TotalRecebido > CCur(lblTotPedido) Then
        LabelTroco.ForeColor = &H800000
        lblTroco.ForeColor = &H800000
        LabelTroco.Caption = "Troco :"
        lblTroco.Caption = Format(lblTroco, "0.00")
        Exit Sub
    ElseIf TotalRecebido = CCur(lblTotPedido) Then
        LabelTroco.ForeColor = &H800000
        lblTroco.ForeColor = vbBlack
        LabelTroco.Caption = "Nenhum :"
        lblTroco.Caption = Format(lblTroco.Caption, "0.00")
        Exit Sub
    End If

Exit Sub

trataAqui:
Call ErrosGeraisLog(Date, Me.Name, "CalcValRecebidos", Err.Description, Err.Number)
Erro "Nos Valores Recebidos"

End Sub



Private Sub MontaPedido()
    lblTotPedido.Caption = 0
    CodSeq = 0
    TotPedido = 0
    contPedido = 0
    
    CodSeq = Val(txtCodSeq.Text)
    
    sql = "Select sequencia, CODIGO_CLIENTE, total_saida from saidas_produto "
    sql = sql & " where SEQUENCIA =  " & CodSeq
    sql = sql & " and (STATUS_SAIDA) IS NULL"
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Cnn, 1, 2
    If Rs.RecordCount <> 0 Then
        Rs.MoveLast
        Rs.MoveFirst
        CodCliente = Rs(1)
        lbl_Sub_Tot_Pedido.Caption = Format(Rs(2), "###,##0.00")
        lblTotPedido.Caption = Format(Rs(2), "###,##0.00")
        txtCodSeq.Enabled = False
        fmeListaFormaPgto.Enabled = True
        'SendKeys "{Tab}"
        LIST_DETALHESPGTO.SetFocus
        lblTeclas.Caption = "Selecione as Op��es de Pagamento com a Tecla [ Espa�o ] "
        lblAjuda2.Caption = " ou clique do [ Mouse ] e Tecle [ Enter ]"
    Else
        lblTeclas.Caption = ""
        lblTroco.Caption = ""
        LIST_DETALHESPGTO.ListItems.Clear
        Call GetDados
        MsgBox "Pedido n�o Encontrado...!", vbInformation, "Aviso"
        txtCodSeq_GotFocus
    End If
    
    Rs.Close
    Set Rs = Nothing
   
End Sub

Private Sub SQL_Registro()
    Dim cont        As Integer
    Dim intAux      As String
    Dim x As Integer
    Dim i As Integer
    Dim SomaCheques As Double
    Dim DifPagCheque As Double
    Dim selFormaPgto As String
    Dim vlrtroco As Double
    Dim strDataPedido As String
    
    Screen.MousePointer = 11
    
    selFormaPgto = ""
    total_Troco = "0,00"
    
    On Error GoTo TrataErroInsertFormPgto
    
    SomaCheques = 0
    DifChequePagCheque = 0
    vlrtroco = 0

    sql = ""
    tipo = "I"
    
    Select Case tipo

    Case "I"
        ''Procedure completa select com retorno
'        sql = "CREATE PROCEDURE SP_PRODUTO ("
'        sql = sql + " CODIGO_PRODUTO VARCHAR(13))"
'        sql = sql + " RETURNS ("
'        sql = sql + " OUT_CODIGO DOUBLE PRECISION, OUT_DESCRICAO VARCHAR(80), OUT_PRECO DOUBLE PRECISION )"
'        sql = sql + " AS BEGIN FOR "
'        sql = sql + " Select CODIGO_INTERNO, DESCRICAO, PRECO "
'        sql = sql & " FROM PRODUTO WHERE CODIGO_INTERNO = :CODIGO_PRODUTO INTO :OUT_CODIGO, OUT_DESCRICAO, :OUT_PRECO "
'        sql = sql + "  DO suspend; END "
'        Cnn.Execute sql
'
'        Set mobjCmd = New ADODB.Command
'        Set mobjCmd.ActiveConnection = Cnn
'
'        Call ClearCommandParameters
'
'        mobjCmd.CommandType = adCmdStoredProc
'
'        Codigo = 1026
'        'IN-parameters
'        mobjCmd.Parameters.Append mobjCmd.CreateParameter("CODIGO_PRODUTO", adVarChar, adParamInput, 14, Codigo)
'
'        'OUT -Parameters
'        mobjCmd.Parameters.Append mobjCmd.CreateParameter("OUT_CODIGO", adDouble, adParamOutput) 'RETORNA_PARAMETRO DO CAMPO
'        mobjCmd.Parameters.Append mobjCmd.CreateParameter("OUT_DESCRICAO", adBSTR, adParamOutput)
'        mobjCmd.Parameters.Append mobjCmd.CreateParameter("OUT_PRECO", adDouble, adParamOutput)
''        mobjCmd.Parameters.Append mobjCmd.CreateParameter("@VLR_TOT_CUST", adDouble, adParamInput, 8, FormatNumber(RsTemp1!TOT_ITEN_CUSTO, 3))
''        mobjCmd.Parameters.Append mobjCmd.CreateParameter("@VLR_TOT_VEND", adDouble, adParamInput, 8, FormatNumber(RsTemp1!TOT_ITEN_VENDA, 3))
''        mobjCmd.Parameters.Append mobjCmd.CreateParameter("@PERC_LUCRO", adDouble, adParamInput, 8, PERC_LUCRO_ITEM)
''        mobjCmd.Parameters.Append mobjCmd.CreateParameter("@PERC_PARTICIP_PROD", adDouble, adParamInput, 8, FormatNumber(PARTICIPACAO_PRODUTO, 2))
'        mobjCmd.CommandText = "SP_PRODUTO"
'        mobjCmd.Execute
'        'RETORNA OS PARAMETROS
'        strCodigo = mobjCmd.Parameters("OUT_CODIGO") 'RETORNA PARAMETRO  - adParamOutput
'        Descricao = mobjCmd.Parameters("OUT_DESCRICAO") 'RETORNA PARAMETRO  - adParamOutput
'        Preco = mobjCmd.Parameters("OUT_PRECO") 'RETORNA PARAMETRO  - adParamOutput
'        'ou outra forma de retorno pelo index
'        strOutputParam0 = mobjCmd.Parameters(0).Value
'        strOutputParam1 = mobjCmd.Parameters(1).Value
'        strOutputParam2 = mobjCmd.Parameters(2).Value
'        strOutputParam3 = mobjCmd.Parameters(3).Value
        
'*******Procedure completa select com retorno
'*************************************************************************************************
'*************************************************************************************************

               
'        sql = "Create PROCEDURE SP_UPDATE_FORMA_PGTO1000 (NRO_PEDIDO DOUBLE PRECISION, FORMA_PGTO VARCHAR(5), STR_STATUS_SAIDA VARCHAR(1))"
'        sql = sql + " AS BEGIN "
'        sql = sql + " UPDATE SAIDAS_PRODUTO SET FORMAPGTO=:FORMA_PGTO, STATUS_SAIDA =:STR_STATUS_SAIDA WHERE SEQUENCIA=:NRO_PEDIDO; END"
'        Cnn.Execute sql
'
'
''        gTransacao = True
''        Cnn.BeginTrans
''
''        contador = 0
'
'        ClearCommandParameters
'        cmd.ActiveConnection = Cnn
'        cmd.CommandType = adCmdStoredProc
''
'        'IN-parameters
'        cmd.Parameters.Append cmd.CreateParameter("SEQUENCIA", adDouble, adParamInput, , CCur(txtCodSeq.Text))
'        cmd.Parameters.Append cmd.CreateParameter("FORMA_PGTO", adBSTR, adParamInput, , "R")
'        cmd.Parameters.Append cmd.CreateParameter("STATUS_SAIDA", adBSTR, adParamInput, , "S")
'        cmd.CommandText = "SP_UPDATE_FORMA_PGTO1000"
'        cmd.Execute
        
        
'        sql = "Create PROCEDURE SP_INSERT_RECE_PAGA(STR_CODIGO DOUBLE PRECISION, STR_DATA VARCHAR(10), STR_DESCRICAO VARCHAR(40), "
'        sql = sql + " STR_VALOR VARCHAR(10), STR_DATA_BAIXA VARCHAR(10), STR_VALOR_BAIXA VARCHAR(10), STR_TIPO_MOVIMENTACAO VARCHAR(1), "
'        sql = sql + " STR_TP_FAVORECIDO VARCHAR(1), STR_SEQUENCIA DOUBLE PRECISION )"
'        sql = sql + " AS BEGIN "
'        sql = sql + " INSERT INTO RECE_PAGA (CODIGO, DATA, DESCRICAO, VALOR, DATA_BAIXA, VALOR_BAIXA, TIPO_MOVIMENTACAO,"
'        sql = sql + " TP_FAVORECIDO, SEQUENCIA) VALUES (:STR_CODIGO, :STR_DATA, :STR_DESCRICAO, :STR_VALOR, :STR_DATA_BAIXA, "
'        sql = sql + " :STR_VALOR_BAIXA, :STR_TIPO_MOVIMENTACAO, :STR_TP_FAVORECIDO, :STR_SEQUENCIA); END"
'        Cnn.Execute sql
        
'        ClearCommandParameters
'        cmd.ActiveConnection = Cnn
'        cmd.CommandType = adCmdStoredProc
'
'        '***** INSERE NO CONTAS A RECEBER RECE_PAGA
'        NovoCodigo = Select_Max("rece_paga", "SEQUENCIA")
'
'        'IN-parameters
'        cmd.Parameters.Append cmd.CreateParameter("STR_CODIGO", adDouble, adParamInput, , CodCliente)
'        cmd.Parameters.Append cmd.CreateParameter("STR_DATA", adBSTR, adParamInput, , Format(Date, "MM/DD/yyyy"))
'        cmd.Parameters.Append cmd.CreateParameter("STR_DESCRICAO", adBSTR, adParamInput, , "Recebimento Entrada n� " & Format(txtCodSeq.Text, "0000"))
'        cmd.Parameters.Append cmd.CreateParameter("STR_VALOR", adBSTR, adParamInput, , Troca_Virg_Zero(Format(lblTotPedido.Caption, "0.00")))
'        cmd.Parameters.Append cmd.CreateParameter("STR_DATA_BAIXA", adBSTR, adParamInput, , Format(Date, "MM/DD/yyyy"))
'        cmd.Parameters.Append cmd.CreateParameter("STR_VALOR_BAIXA", adBSTR, adParamInput, , Troca_Virg_Zero(Format(lblTotPedido.Caption, "0.00")))
'        cmd.Parameters.Append cmd.CreateParameter("STR_TIPO_MOVIMENTACAO", adBSTR, adParamInput, , "R")
'        cmd.Parameters.Append cmd.CreateParameter("STR_TP_FAVORECIDO", adBSTR, adParamInput, , "C")
'        cmd.Parameters.Append cmd.CreateParameter("STR_CODIGO", adDouble, adParamInput, , CCur(NovoCodigo))
'        cmd.CommandText = "SP_INSERT_RECE_PAGA"
'        cmd.Execute
'
'        DoEvents
'        Cnn.CommitTrans
        
'        Cnn.Rollback
'        gTransacao = False

        'Exit Sub
        
        conta_Insert = conta_Insert + 1
        If conta_Insert > 1 Then
            Unload Me
            Exit Sub
        End If
        
        
        If ValidaDados = False Then Exit Sub:
        
        gTransacao = True
        Cnn.BeginTrans
        
        contador = 0
            
        '***** INSERE NO CONTAS A RECEBER RECE_PAGA
        NovoCodigo = Select_Max("rece_paga", "SEQUENCIA")
        
        sql = "Select data_nf from saidas_produto where sequencia = " & CDbl(txtCodSeq.Text)
        Set Rstemp = New ADODB.Recordset
        Rstemp.Open sql, Cnn, 1, 2
        If Rstemp.RecordCount > 0 Then
            strDataPedido = Format(Rstemp!Data_NF, "dd/mm/yyyy")
        End If
        
        'sql = "Insert into Rece_Paga values ( "
        sql = "Insert into Rece_Paga (CODIGO,DATA,DESCRICAO,VALOR,TIPO_MOVIMENTACAO,TP_FAVORECIDO, SEQUENCIA) values ( "
        sql = sql & CodCliente & ","
        sql = sql & "'" & Format(strDataPedido, "mm/dd/yyyy") & "',"
        sql = sql & "'Recebimento Entrada n� " & Format(txtCodSeq.Text, "0000") & "',"
        sql = sql & "'" & Troca_Virg_Zero(Format(lblTotPedido.Caption, "0.00")) & "',"
        sql = sql & "'R',"
        sql = sql & "'C'," 'TP_FAVORECIDO
        'sql = sql & "NULL,NULL,"
        sql = sql & NovoCodigo & ")"
        'sql = sql & "NULL,NULL,NULL)"
        Cnn.Execute sql
        
        For i = 1 To LIST_DETALHESPGTO.ListItems.Count
            If LIST_DETALHESPGTO.ListItems(i).Checked = True Then
                contador = contador + 1
                sql = "INSERT INTO FORMA_PGTO VALUES ("
                sql = sql & CodSeq & ","
                sql = sql & CodCliente & ","
                sql = sql & LIST_DETALHESPGTO.ListItems(i).Text & ","   'cod_FORMAPGTO
                sql = sql & "'S'," 'SAIDA
                'sql = sql & "'" & i & "/" & LIST_DETALHESPGTO.ListItems.Count & "',"
                sql = sql & "NULL,"
                'sql = sql & "'" & Format(Date, "mm/dd/yyyy") & "',"
                sql = sql & "'" & Format(strDataPedido, "mm/dd/yyyy") & "',"
                If LIST_DETALHESPGTO.ListItems(i).Text = 1 Then
                    sql = sql & "'" & Troca_Virg_Zero(Format(txtVlrdinheiro.Text, "0.00")) & "',"
                ElseIf LIST_DETALHESPGTO.ListItems(i).Text = 2 Then
                    'sql = sql & "'" & Troca_Virg_Zero(Format(txtVlrCheques.Text, "0.00")) & "',"
                    '27/10/2010
                    sql2 = "SELECT * FROM CHEQUES WHERE SEQUENCIA = " & CDbl(txtCodSeq.Text) & " ORDER BY SEQUENCIA "
                    Set RsTemp1 = New ADODB.Recordset
                    RsTemp1.Open sql2, Cnn, 1, 2
                    If RsTemp1.RecordCount > 0 Then
                        RsTemp1.MoveLast
                        RsTemp1.MoveFirst
                        x = 0
                        While Not RsTemp1.EOF
                            x = x + 1 'QTDE DE CHEQUES
                            sql = "INSERT INTO FORMA_PGTO VALUES ("
                            sql = sql & CodSeq & ","
                            sql = sql & CodCliente & ","
                            sql = sql & LIST_DETALHESPGTO.ListItems(i).Text & ","   'cod_FORMAPGTO
                            sql = sql & "'S'," 'SAIDA
                            sql = sql & "'" & x & "/" & RsTemp1.RecordCount & "',"
                            sql = sql & "'" & Format(RsTemp1!VENCIMENTO, "mm/dd/yyyy") & "',"
                            sql = sql & "'" & Troca_Virg_Zero(Format(RsTemp1!valor, "0.00")) & "',"
                            sql = sql & "'" & UCase(NomeUsuario) & "')"
                            Cnn.Execute sql
                            RsTemp1.MoveNext
                        Wend
                    End If
                    RsTemp1.Close
                    Set RsTemp1 = Nothing
                ElseIf LIST_DETALHESPGTO.ListItems(i).Text = 3 Then
                    sql = sql & "'" & Troca_Virg_Zero(Format(txtVlrCartCredito.Text, "0.00")) & "',"
                ElseIf LIST_DETALHESPGTO.ListItems(i).Text = 4 Then
                    sql = sql & "'" & Troca_Virg_Zero(Format(txtVlrCartEletron.Text, "0.00")) & "',"
                ElseIf LIST_DETALHESPGTO.ListItems(i).Text = 5 Then
                    sql = sql & "'" & Troca_Virg_Zero(Format(txtPendente.Text, "0.00")) & "',"
                    selFormaPgto = "R"
                ElseIf LIST_DETALHESPGTO.ListItems(i).Text = 7 Then
                    sql = sql & "'" & Troca_Virg_Zero(Format(txtVlrCartLoja.Text, "0.00")) & "',"
                End If
                
                If LIST_DETALHESPGTO.ListItems(i).Text <> 2 Then
                    sql = sql & "'" & UCase(NomeUsuario) & "')"
                    Cnn.Execute sql
                End If
            End If
        Next i
        
        '"UPDATE OR INSERT INTO EMPLOYEE (ID, NAME)  VALUES (:ID, :NAME)  RETURNING OLD.Name"
        
        sql = " UPDATE SAIDAS_PRODUTO SET "
        If selFormaPgto = "R" Then
            sql = sql & " FormaPgto = 'R'" & ","
        End If
        sql = sql & " STATUS_SAIDA = 'S',"
        sql = sql & " TOTAL_DESCONTO = TOTAL_DESCONTO + " & Troca_Virg_Zero(txt_Desconto.Text) & ","
        sql = sql & " TOTAL_SAIDA = '" & Troca_Virg_Zero(lblTotPedido.Caption) & "'"
        sql = sql & " WHERE SEQUENCIA = " & CodSeq
        Cnn.Execute sql
        
        NovoCodigo = Select_Max("CTAS_PENDENTE", "SEQUENCIA")
        If txtPendente.Text <> "" And txtPendente.Text <> "0,00" And CDbl(txtPendente) > 0 Then
            sql = "Insert into CTAS_PENDENTE values ( "
            sql = sql & NovoCodigo & ","
            sql = sql & CodCliente & ","
            'sql = sql & "'" & Format(Date, "mm/dd/yyyy") & "','"
            sql = sql & "'" & Format(strDataPedido, "mm/dd/yyyy") & "','"
            sql = sql & "Pendencia Ref Pedido n� " & Format(txtCodSeq.Text, "0000") & "',"
            sql = sql & "'" & Troca_Virg_Zero(Format(txtPendente.Text, "0.00")) & "',null)"
            Cnn.Execute sql
        End If
    
        DoEvents
        Cnn.CommitTrans
        gTransacao = False
        Screen.MousePointer = 1
        
        If flag_Imp_Forma_Pgto = True Then
            If MsgBox("Imprimir forma de pagamento...?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me") = vbYes Then
                Call Imprime_Forma_Pgto
            End If
        End If
        
        Dim Msg As String
        Msg = IIf(retImpFiscal <> 0 And retImpFiscal <> 5, "A Impressora Fiscal est� pronta para Imprimir...?", "O SAT est� conectado ao computador?")
        
        If flagImpFiscalSelecionada = True Then
            'If MsgBox("Dados Atualizados com Sucesso." & vbNewLine & vbNewLine & "Impressora Fiscal Pronta para Imprimir...?", vbQuestion + vbYesNo + vbDefaultButton2, "Responda-me") = vbYes Then
            If MsgBox("Aten��o..." & vbNewLine & vbNewLine & Msg, vbQuestion + vbYesNo + vbDefaultButton2, "Responda-me") = vbYes Then
                str_cnpj_cpf = ""
                frm_CNPJ_CPF.Show 1
                If retImpFiscal = 1 Then
                    Call ImprimeCupom_Fiscal(CodSeq, txtVlrdinheiro, txtVlrCheques, txtVlrCartCredito, txtVlrCartEletron, txtPendente, str_cnpj_cpf, "0.00")
                ElseIf retImpFiscal = 5 Then 'sat
                            Call CarregarConfig_SAT
                            
                            Dim aRetorno As String
                            Dim arrRetorno() As String
                            Dim retorno_ConsultarSAT As String
                            Dim Imprimir As String
                            
                            Screen.MousePointer = 1
                            
                            'consultar sat operacional avaliar melhor se realmente � necess�rio usar a rotina abaixo
                            aRetorno = spdSAT.ConsultarStatusOperacional(NumeroSessao)
                            If Len(aRetorno) = 0 Then
                                MsgBox spdSAT.ConsultarStatusOperacional(NumeroSessao)
                                MsgBox spdSAT.ConsultarSAT(NumeroSessao)
                            End If
                            arrRetorno = Split(aRetorno, "|")
                            If (UBound(arrRetorno) > 7) Then
                                edtUltIDAutorizado = arrRetorno(8)
                                autorizado = UCase(arrRetorno(2))
                                qrcode = arrRetorno(11)
                            End If
                        
                            If InStr(aRetorno, "NAO_CONECTADO") Then
                                'MsgBox "Sem comunica��o"
                                lbl_Situacao_SAT = "SAT Desconectado...!"
                                'Exit Function
                                MsgBox aRetorno, vbInformation, "Aviso"
                                Exit Sub
                            ElseIf InStr(aRetorno, "CONECTADO") Then
                                'MsgBox "ok, SAT Conectado...!", vbInformation, "Aviso"
                                lbl_Situacao_SAT = "SAT Conectado...!"
                            ElseIf InStr(aRetorno, "Erro ao tentar abrir a porta de comunicacao") Then
                                MsgBox aRetorno, vbInformation, "Aviso"
                                'lbl_Situacao_SAT = "SAT Conectado...!"
                                Exit Sub
                            Else 'qualquer erro desconhecido
                                MsgBox aRetorno, vbInformation, "Aviso"
                                Exit Sub
                            End If

ConsultarSAT:
                            'consultar SAT
                            aRetorno = spdSAT.ConsultarSAT(NumeroSessao)
                            
                            'varre array
                            arrRetorno = Split(aRetorno, "|")
                            If (UBound(arrRetorno) > 3) Then
                                sessao = arrRetorno(0)
                                retorno_ConsultarSAT = arrRetorno(1)
                                mensagem_Sat = UCase(arrRetorno(2))
                            End If
                            
                            Select Case retorno_ConsultarSAT
                                Case "08000"    'SAT em opera��o
                                    'ok, tudo normal com o sat
                                    If GeraDataset_Produtos(CodSeq, frmCdPgto.str_cnpj_cpf) = True Then
                                        'MsgBox "OK"
                                    End If
                                Case "08098"    'SAT em processamento.Tente novamente.
                                       'impress�o cupom caso algo d� errado
                                        Imprimir = MsgBox("N�o foi poss�vel realizar a impress�o do Cupom Fiscal Eletr�nico," & vbNewLine & "Imprimir Cupom n�o fiscal ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
                                        If (Imprimir = vbYes) Then
                                            Call frmSaidas.ImprimePedidoCupon
                                            Imprimir = MsgBox("Deseja imprimir o cupom novamente?", vbYesNo + vbDefaultButton2 + vbInformation + vbOKOnly, "Aviso")
                                            If (Imprimir = vbYes) Then Call frmSaidas.ImprimePedidoCupon
                                        End If
                                        Call Erro_ConsultaSAT("Pedido: " & CStr(CodSeq), "Data: " & Now, Me.Name, "spdSAT.ConsultarSAT", "Ret. SAT: " & retorno_ConsultarSAT, "Mensagem SAT: " & CStr(mensagem_Sat))
                                Case "08099"    'Erro desconhecido.
                                        Imprimir = MsgBox("N�o foi poss�vel realizar a impress�o do Cupom Fiscal Eletr�nico," & vbNewLine & "Imprimir Cupom n�o fiscal ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
                                        If (Imprimir = vbYes) Then
                                            frmSaidas.ImprimePedidoCupon
                                            Imprimir = MsgBox("Deseja imprimir o cupom novamente?", vbYesNo + vbDefaultButton2 + vbInformation + vbOKOnly, "Aviso")
                                            If (Imprimir = vbYes) Then frmSaidas.ImprimePedidoCupon
                                        End If
                                    Call Erro_ConsultaSAT("Pedido: " & CStr(CodSeq), "Data: " & Now, Me.Name, "spdSAT.ConsultarSAT", "Ret. SAT: " & retorno_ConsultarSAT, "Mensagem SAT: " & CStr(mensagem_Sat))
                                    MsgBox mensagem_Sat & vbNewLine & "Erro: " & retorno_ConsultarSAT, vbInformation, "Aviso"
                            End Select
                    
                'Else
                    'Call Imprime_Cupom_Fiscal_Sweda(CodSeq, txtVlrdinheiro, txtVlrCheques, txtVlrCartCredito, txtVlrCartEletron, txtPendente, str_cnpj_cpf, "0.00")
                End If
                Unload frm_CNPJ_CPF
                Set frm_CNPJ_CPF = Nothing
            End If
            If flag_Gaveta_Bematec = True Then
                Retorno = Bematech_FI_AcionaGaveta()
                'Fun��o que analisa o retorno da impressora
                Call VerificaRetornoImpressora("", "", "Acionamento da Gaveta")
            ElseIf flag_Gaveta_Elgin = True Then
                'Retorno = Elgin_AcionaGaveta()
                'Call TrataRetorno(Retorno)
            End If
        Else
            MsgBox "Dados Atualizados com Sucesso.", vbInformation, "Aviso"
        End If
        
        If IsNumeric(lblTroco.Caption) And CCur(lblTroco.Caption) > 0 Then
            total_Troco = lblTroco.Caption
        Else
            total_Troco = "0,00"
        End If
        
    Case "A"

            For cont = 1 To List_Itens_FormPgto.ListItems.Count
                FlexList_CondPgto.Row = cont
                sql = "UPDATE COND_PGTO SET"
                sql = sql & " DESCRICAO = '" & FiltraAspasSimples(txt_Descr_Pgto) & "'"
                If Len(Trim(txtQtParc.Text)) > 0 Then
                    sql = sql & ",QTDE_PARCELAS = " & txtQtParc.Text
                Else
                    sql = sql & ",QTDE_PARCELAS = null"
                End If
                
                FlexList_CondPgto.Col = 0
                If Len(Trim(FlexList_CondPgto.Text)) > 0 Then
                    sql = sql & ",VCTO_01 = '" & Trim(FlexList_CondPgto.Text) & "'"
                Else
                    sql = sql & ",null"
                End If
                
                FlexList_CondPgto.Col = 2
                If Len(Trim(FlexList_CondPgto.Text)) > 0 Then
                    sql = sql & ",Num_Parc = '" & Trim(FlexList_CondPgto.Text) & "'"
                Else
                    sql = sql & ",null"
                End If
                
                FlexList_CondPgto.Col = 1
                If Len(Trim(FlexList_CondPgto.Text)) > 0 Then
                    sql = sql & ",Intervalo_Parc = '" & Trim(FlexList_CondPgto.Text) & "'"
                Else
                    sql = sql & ",null"
                End If
                
                sql = sql & ",FLAG_LCTO_VALE = '" & FlagPendText & "'"
                sql = sql & " WHERE Codigo = " & lblNrCodigo.Caption
                Cnn.Execute sql
            Next
            
    Case "E"
         
            Msg = "Confirma REALMENTE a Exclus�o desta CONDI��O DE PAGAMENTO ?"
            Estilo = vbYesNo + vbCritical + vbDefaultButton2
            Titulo = Me.Caption
            resposta = MsgBox(Msg, Estilo, Titulo)
            
            If resposta = 6 Then ' Sim
                sql = "DELETE FROM COND_PGTO "
                sql = sql & " WHERE Codigo = " & lblNrCodigo.Caption
                Cnn.Execute sql
            End If
    
    End Select
 
    Screen.MousePointer = 1

Exit Sub

TrataErroInsertFormPgto:

If gTransacao = True Then Cnn.RollbackTrans
    If Err.Number <> 0 Then
        'MsgBox Err.Source
      
        Call ErrosGeraisLog(Date, Me.Name, "MontaPedidos", Err.Description, Err.Number)
        Erro "Gravar Dados Forma Pgto"
    End If
Screen.MousePointer = 1
End Sub

Private Sub Imprime_Forma_Pgto()
    'inicia impressao do titulo
    Printer.FontName = "Arial"
    Printer.FontSize = 5
    Printer.Print
    Printer.FontSize = 8
    Printer.Print Space(1)
    Printer.Print "=========================================="
    Printer.Print (UCase(Mid(NomeEmpresa, 1, 42)))
    Printer.Print "=========================================="
    Printer.Print Tab(1); TIPO_PEDIDO & " " & Format(CodSeq, "0000"); Tab(20); "DATA: "; Tab(26); Format(Date, "dd/mm/yyyy"); Tab(38); Format(Now, "hh:mm")
    Printer.Print "=========================================="
    Printer.Print Tab(15); "SUB TOTAL:"; Tab(32); Format(Format(Me.lbl_Sub_Tot_Pedido.Caption, "##,##0.00"), "@@@@@@@@")
       If CCur(txt_Desconto.Text) > 0 Then
        Printer.Print Tab(15); "DESCONTO:"; Tab(32); Format(Format(txt_Desconto.Text, "##,##0.00"), "@@@@@@@@")
    End If
    If CCur(Me.lbl_Sub_Tot_Pedido.Caption) <> CCur(Me.lblTotPedido.Caption) Then
        Printer.Print Tab(15); "TOTAL:"; Tab(32); Format(Format(Me.lblTotPedido.Caption, "##,##0.00"), "@@@@@@@@")
    End If
    Printer.Print "=========================================="
    Printer.Print "FORMA DE PAGAMENTO"
    Printer.Print "=========================================="
    If CCur(txtVlrdinheiro.Text) > 0 Then
        Printer.Print Tab(15); "DINHEIRO:"; Tab(32); Format(Format(txtVlrdinheiro.Text, "##,##0.00"), "@@@@@@@@")
    End If
    
    If CCur(txtVlrCheques.Text) > 0 Then
        Printer.Print Tab(15); "CHEQUE(S):"; Tab(32); Format(Format(txtVlrCheques.Text, "##,##0.00"), "@@@@@@@@")
    End If

    If CCur(txtVlrCartCredito.Text) > 0 Then
        Printer.Print Tab(15); "CART. CREDITO:"; Tab(32); Format(Format(txtVlrCartCredito.Text, "##,##0.00"), "@@@@@@@@")
    End If
    
    If CCur(txtVlrCartEletron.Text) > 0 Then
        Printer.Print Tab(15); "CART. ELETR.:"; Tab(32); Format(Format(txtVlrCartEletron.Text, "##,##0.00"), "@@@@@@@@")
    End If
    
    If CCur(txtPendente.Text) > 0 Then
        Printer.Print Tab(15); "PENDENTE.:"; Tab(32); Format(Format(txtPendente.Text, "##,##0.00"), "@@@@@@@@")
    End If
    Printer.Print "=========================================="
    Printer.Print Tab(15); "VALOR PAGO:"; Tab(32); Format(Format(Me.lblTotPedido.Caption, "##,##0.00"), "@@@@@@@@")
    
    If CCur(Me.lblTroco.Caption) > 0 Then
        Printer.Print Tab(15); "TROCO:"; Tab(32); Format(Format(Me.lblTroco.Caption, "##,##0.00"), "@@@@@@@@")
    End If
   
    Printer.Print "========================================="
   ' Printer.Print "CUPOM SEM VALOR FISCAL"
   ' Printer.Print "========================================="
    Printer.Print
    Printer.Print "CONFERIDO POR:"
    Printer.Print "ASS:-------------------------------------"
    
    For i = 1 To 11
        Printer.Print Space(1)
    Next i
    
    Printer.EndDoc

End Sub
Private Function ValidaDados() As Boolean
    If TotalRecebido < CCur(lblTotPedido) Then
        MsgBox "Somat�ria a Receber, diverge do Valor Total a Pagar. ", vbCritical, "Aviso"
        For i = 1 To LIST_DETALHESPGTO.ListItems.Count
             LIST_DETALHESPGTO.ListItems.Item(i).Checked = True
             Exit For
        Next
        txtVlrdinheiro.Enabled = True
        txtVlrdinheiro.SetFocus
        txtVlrdinheiro_GotFocus
        If txtVlrCheques.Enabled = True Then
            txtVlrCheques.SetFocus
        ElseIf txtVlrCartCredito.Enabled = True Then
            txtVlrCartCredito.SetFocus
        ElseIf txtVlrCartEletron.Enabled = True Then
            txtVlrCartEletron.SetFocus
        ElseIf txtPendente.Enabled = True Then
            txtPendente.SetFocus
        End If
        ValidaDados = False
        Exit Function
    End If
    
    If TotalRecebido = 0 Then
        MsgBox "Digite um Pedido V�lido?", vbCritical, "Aviso"
        ValidaDados = False
        txtCodSeq.SetFocus
        Exit Function
    End If
    
    If Val(lblTroco.Caption) > 0 And txtPendente.Enabled = True Then
        MsgBox "Se h� troco, n�o h� Pend�ncia." & vbNewLine & "Verifique.", vbCritical, "Aviso"
        txtPendente.SetFocus
        ValidaDados = False
        Exit Function
    End If
    
    ValidaDados = True
    
End Function

Private Sub chk_Imp_Forma_Pgto_Click()
If chk_Imp_Forma_Pgto.Value = Checked Then
    flag_Imp_Forma_Pgto = True
    WriteIniFile App.Path & "\SisAdven.ini", "imp_forma_pgto", "Chk", "1"
Else
    WriteIniFile App.Path & "\SisAdven.ini", "imp_forma_pgto", "Chk", "0"
    flag_Imp_Forma_Pgto = False
End If

End Sub

Private Sub chk_Imp_Forma_Pgto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub cmd_Gravar_Click()
        Call FormatText(Me)
        Call CalcValRecebidos
        Msg = "Salvar Dados...?"
        resposta = MsgBox(Msg, vbQuestion + vbYesNoCancel + vbDefaultButton1, "Responda-me")
        If resposta = 7 Then 'nao
            Call VerificaDeletaLancaCheques
            Call LimpaControles(Me)
            strNumped = ""
            Unload Me
            Exit Sub
        ElseIf resposta = 6 Then 'sim
            If CodCliente = 1 And Val(txtPendente.Text) > 0 Then
                MsgBox "Pend�ncia Recusada para Cliente Consumidor." & vbCrLf & "Selecione outra Forma de Pagamento", vbCritical, "Opera��o n�o Permitida"
                fmeListaFormaPgto.Enabled = True
                txtPendente.Text = "0.00"
                txtPendente.Enabled = False
                LabelTroco.ForeColor = vbBlue
                lblTroco.ForeColor = vbBlue
                LabelTroco.Caption = "Saldo :"
                LIST_DETALHESPGTO.SetFocus
                Exit Sub
            End If
            If ValidaDados = True Then
                SQL_Registro
                tipo = ""
                strNumped = ""
                Unload Me
                Exit Sub
            Else
                Exit Sub
            End If
        ElseIf resposta = 2 Then ' cancelar
            Call VerificaDeletaLancaCheques
            Call LimpaControles(Me)
            strNumped = ""
            lblTotPedido.Caption = "0.00"
            lblTroco.Caption = "0.00"
            LabelTroco.ForeColor = vbBlue
            LabelTroco.Caption = "Saldo :"
            LIST_DETALHESPGTO.ListItems.Clear
            Call GetDados
            txtCodSeq.Text = ""
            lblTeclas.Caption = "Selecione as Op��es de Pagamento com a Tecla [ Espa�o ou Mouse ] e Tecle [ Enter ]"
            fmeListaFormaPgto.Enabled = False
            lblTeclas.Caption = "Digite o n�mero do  Pedido e tecle [  Enter ] "
            txtCodSeq.Enabled = True
            txtCodSeq.SetFocus
            Exit Sub
        End If
End Sub


Private Sub cmd_Limpar_Click()
Call CarregaListItens_Forma_Pgto
txtQtParc.Text = ""
txtIntervalo.Text = ""
txtVlr_Parcela.Text = ""
MskVcto.Text = "__/__/____"
MskVcto.Mask = "##/##/####"
'lbl_Total_Parcelas.Caption = ""
txtQtParc.Enabled = True
lbl_intervalo.Visible = False
txtIntervalo.Visible = False
txtVlr_Parcela.Enabled = False
MskVcto.Enabled = False
txtQtParc.Enabled = True
txtQtParc.SetFocus

End Sub

Private Sub fmeListaFormaPgto_Click()
If txtVlrdinheiro.Enabled = True Then
        txtVlrdinheiro.SetFocus
    ElseIf txtVlrCheques.Enabled = True Then
        txtVlrCheques.SetFocus
    ElseIf txtVlrCartCredito.Enabled = True Then
        txtVlrCartCredito.SetFocus
    ElseIf txtVlrCartEletron.Enabled = True Then
        txtVlrCartEletron.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Call FormatText(Me)
        Call CalcValRecebidos
        Msg = "Salvar Dados...?"
        resposta = MsgBox(Msg, vbQuestion + vbYesNoCancel + vbDefaultButton1, "Responda-me")
        If resposta = 7 Then 'nao
            Call VerificaDeletaLancaCheques
            Call LimpaControles(Me)
            strNumped = ""
            Unload Me
            Exit Sub
        ElseIf resposta = 6 Then 'sim
            If CodCliente = 1 And Val(txtPendente.Text) > 0 Then
                MsgBox "Pend�ncia Recusada para Cliente Consumidor." & vbCrLf & "Selecione outra Forma de Pagamento", vbCritical, "Opera��o n�o Permitida"
                fmeListaFormaPgto.Enabled = True
                txtPendente.Text = "0.00"
                txtPendente.Enabled = False
                LabelTroco.ForeColor = vbBlue
                lblTroco.ForeColor = vbBlue
                LabelTroco.Caption = "Saldo :"
                LIST_DETALHESPGTO.SetFocus
                Exit Sub
            End If
            If ValidaDados = True Then
                Call SQL_Registro
                tipo = ""
                strNumped = ""
                Unload Me
                If CCur(total_Troco) > 0 Then
                    MsgBox "Troco R$ " & total_Troco, vbInformation, "Troco"
                End If
                Exit Sub
            Else
                Exit Sub
            End If
        ElseIf resposta = 2 Then ' cancelar
            Call VerificaDeletaLancaCheques
            Call LimpaControles(Me)
            strNumped = ""
            lblTotPedido.Caption = "0.00"
            lblTroco.Caption = "0.00"
            LabelTroco.ForeColor = vbBlue
            LabelTroco.Caption = "Saldo :"
            LIST_DETALHESPGTO.ListItems.Clear
            Call GetDados
            txtCodSeq.Text = ""
            lblTeclas.Caption = "Selecione as Op��es de Pagamento com a Tecla [ Espa�o ou Mouse ] e Tecle [ Enter ]"
            fmeListaFormaPgto.Enabled = False
            lblTeclas.Caption = "Digite o n�mero do  Pedido e tecle [  Enter ] "
            txtCodSeq.Enabled = True
            txtCodSeq.SetFocus
            Exit Sub
        End If
    End If

End Sub



Private Sub VerificaDeletaLancaCheques()
    
If Len(txtCodSeq.Text) > 0 Then
    sql = "select * from  cheques "
    sql = sql & " where sequencia = " & Trim(txtCodSeq.Text)
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Cnn, 1, 2
    If Rs.RecordCount > 0 Then
        sql = " delete from cheques "
        sql = sql & " where sequencia = " & Trim(txtCodSeq.Text)
        Cnn.Execute (sql)
    End If
    Rs.Close
End If
    
End Sub

Private Function VerificaChecados() As Boolean

    Dim i As Integer
    Dim contador As Integer
    Dim vlrMensagem As Integer
    Dim FlagLancaCheque As Boolean
    
    txt_Desconto.Enabled = False
    VerificaChecados = False
    FlagLancaCheque = False

    For i = 1 To LIST_DETALHESPGTO.ListItems.Count
        If LIST_DETALHESPGTO.ListItems(i).Checked = True Then
            Select Case i
                Case 1
                    txtVlrdinheiro.Enabled = True
                Case 2
                    'txtVlrCheques.Enabled = True
                    FlagLancaCheque = True
                Case 3
                    txtVlrCartCredito.Enabled = True
                Case 4
                    txtVlrCartEletron.Enabled = True
                Case 5
                    txtPendente.Enabled = True
                Case 6
                    If CodCliente = 1 Then
                        MsgBox "N�o � Poss�vel gerar Boleto(s) para o Cliente Consumidor.", vbCritical, "Opera��o n�o Permitida"
                        fmeListaFormaPgto.Enabled = True
                        LabelTroco.ForeColor = vbBlue
                        lblTroco.ForeColor = vbBlue
                        LabelTroco.Caption = "Saldo :"
                        LIST_DETALHESPGTO.SetFocus
                        Exit Function
                    End If
                    KeyPreview = False
                    lblTotCompra.Caption = Format(lbl_Sub_Tot_Pedido.Caption, "###,##0.00")
                    SSPanel_BOLETO.Visible = True
                    SSFrame1.Enabled = False
                    '''txtNroNF.Text = Select_Max("saidas_produto", "nf")
                    txtNroNF.Text = Select_Max("nfe", "nro_nf")
                    txtNroNF.SetFocus
                Case 7
                    txtVlrCartLoja.Enabled = True
           End Select
           VerificaChecados = True
           txt_Desconto.Enabled = True
        End If
    Next i
    
    If VerificaChecados = False Then
        VerificaChecados = False
        Exit Function
    Else
        If FlagLancaCheque = True Then
            FrmCheques.nroPedido = Val(txtCodSeq.Text)
            FrmCheques.mi_transacao = 1
            FrmCheques.CodCliente = CodCliente
            FrmCheques.cmd_Gravar.Enabled = False
            FrmCheques.cmd_Sair.Enabled = False
            FrmCheques.Show 1
        ElseIf txtVlrdinheiro.Enabled = True Then
            txtVlrdinheiro.SetFocus
        'ElseIf txtVlrCheques.Enabled = True Then
        '    txtVlrCheques.SetFocus
        ElseIf txtVlrCartCredito.Enabled = True Then
            txtVlrCartCredito.SetFocus
        ElseIf txtVlrCartEletron.Enabled = True Then
            txtVlrCartEletron.SetFocus
        ElseIf txtPendente.Enabled = True Then
            If txtVlrdinheiro.Enabled = False And txtVlrCheques.Enabled = False And txtVlrCartCredito.Enabled = False And txtVlrCartEletron.Enabled = False Then
                txtPendente.Text = lblTotPedido.Caption
            Exit Function
            Else
                txtPendente.SetFocus
            End If
        ElseIf txtVlrCartLoja.Enabled = True Then
            txtVlrCartLoja.SetFocus
        End If
    End If
    VerificaChecados = True
End Function


Public Sub FormatText(formAtivo As Form)
Dim MyCtrls As Object

For Each MyCtrls In formAtivo.Controls
    If TypeOf MyCtrls Is TextBox Then
        If MyCtrls.Text <> txtCodSeq.Text Then
            MyCtrls.Text = Format(MyCtrls.Text, "0.00")
        End If
    End If
Next MyCtrls
End Sub

Private Sub GetDados()
Dim cont As Integer
cont = 0

LIST_DETALHESPGTO.ListItems.Clear
    
    sql = "Select * from FORMAS "
    Set Rstemp3 = New ADODB.Recordset
    Rstemp3.Open sql, Cnn, 1, 2
    If Rstemp3.RecordCount > 0 Then
        cont = 1
        While Not Rstemp3.EOF
            If Not IsNull(Rstemp3!Codigo) Then
                LIST_DETALHESPGTO.ListItems.Add (cont), , Rstemp3!Codigo
            Else
                LIST_DETALHESPGTO.ListItems.Add (cont), , ""
            End If
            If Not IsNull(Rstemp3!Descricao) Then
                LIST_DETALHESPGTO.ListItems(cont).SubItems(1) = UCase(Rstemp3!Descricao)
            Else
                LIST_DETALHESPGTO.ListItems(cont).SubItems(1) = "Descri��o n�o Encontrada"
            End If
            cont = cont + 1
            Rstemp3.MoveNext
        Wend
        LIST_DETALHESPGTO.ListItems.Add (cont), , 6
        LIST_DETALHESPGTO.ListItems(cont).SubItems(1) = "BOLETO BANC�RIO"
        
        cont = cont + 1
        LIST_DETALHESPGTO.ListItems.Add (cont), , 7
        LIST_DETALHESPGTO.ListItems(cont).SubItems(1) = "CART�O LOJA"
    End If
    
    Rstemp3.Close
    Set Rstemp3 = Nothing
End Sub


Private Sub CarregaListDetalhesPgto()
    With LIST_DETALHESPGTO
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add , "C1", "C�digo", 1000, lvwColumnLeft
        .ColumnHeaders.Add , "C2", "Descri��o", 2200, lvwColumnLeft
      '  .ColumnHeaders.Add , "C3", "Total", 1350, lvwColumnRight
    End With

End Sub

Private Sub Form_Load()
'Me.Left = (frmSaidas.Width - Me.Width) / 2
'Me.Top = ((frmSaidas.Height - Me.Height) / 2)

'Call CentralizaJanela(Me)

'lblTotPedido.Caption = Format(frmSaidas.valTotalPedido, "0.00")
If strNumped <> "" Then
    txtCodSeq.Text = strNumped
End If
Call CarregaListDetalhesPgto
Call CarregaListItens_Forma_Pgto
Call GetDados

Mensagem_Final_Cupom = ReadIniFile(App.Path & "\SisAdven.ini", "Men_Promoc", "", "")
Mensagem_Final_Cupom = UCase(Mid(Mensagem_Final_Cupom, 1, 492))

'12-07-2016 removido a pergunta
'''If ReadIniFile(App.Path & "\SisAdven.ini", "imp_forma_pgto", "Chk", "0") = 0 Then
'''    chk_Imp_Forma_Pgto.Value = Unchecked
'''    flag_Imp_Forma_Pgto = False
'''Else
'''    chk_Imp_Forma_Pgto.Value = Checked
'''    flag_Imp_Forma_Pgto = True
'''End If

chk_Imp_Forma_Pgto.Value = Unchecked
flag_Imp_Forma_Pgto = False

KeyPreview = True

conta_Insert = 0

End Sub

Private Sub CarregaListItens_Forma_Pgto()
    With LIST_ITENS_FORMA_PGTO
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add , "C1", "Parcela", 1000, lvwColumnLeft
        .ColumnHeaders.Add , "C2", "Valor", 1250, lvwColumnRight
        .ColumnHeaders.Add , "C3", "Vencimento", 1350, lvwColumnLeft
        .ColumnHeaders.Add , "C4", "Intervalo", 1850, lvwColumnLeft
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
txtCodSeq.Text = ""
strNumped = ""
End Sub

Private Sub Form_Resize()
'Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Retorno = Bematech_FI_FechaPortaSerial()

End Sub

Private Sub lblValorTTCompra_Click()

End Sub

Private Sub LIST_DETALHESPGTO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        'Unload Me
    ElseIf KeyCode = 13 Then
        fme_Totais.Enabled = False
        If VerificaChecados = False Then Exit Sub:
        fmeListaFormaPgto.Enabled = False
        lblTeclas.Caption = "Digite o Valor e Tecle [ Enter ]  ou  [ Esc ] para Finalizar"
        
        lblTeclas.Visible = True
        lblAjuda2.Visible = False
        fme_Totais.Enabled = True
    End If
End Sub

Private Sub LIST_ITENS_FORMA_PGTO_DblClick()
 For i = LIST_ITENS_FORMA_PGTO.ListItems.Count To 1 Step -1
        If LIST_ITENS_FORMA_PGTO.ListItems(i).Selected = True Then
            flag_altera_forma_pgto_Boleto = True
            LinhaSelecionada = 1
            txtQtParc.Enabled = False
            txtIntervalo.Enabled = False
            'txtVlr_Parcela.Enabled = False
           ' txtQtParc.Text = LIST_ITENS_FORMA_PGTO.ListItems.Item(i).Text
            txtQtParc.Text = LIST_ITENS_FORMA_PGTO.ListItems.Count
            txtIntervalo.Text = LIST_ITENS_FORMA_PGTO.ListItems.Item(i).SubItems(3)
            txtVlr_Parcela.Text = Format(CCur(LIST_ITENS_FORMA_PGTO.ListItems.Item(i).SubItems(1)), "0.00")
            MskVcto.Text = LIST_ITENS_FORMA_PGTO.ListItems.Item(i).SubItems(2)
            'txtIntervalo.SetFocus
            MskVcto.Enabled = True
            MskVcto.SetFocus
            Exit For
        End If
    Next i
End Sub

Private Sub LIST_ITENS_FORMA_PGTO_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Index = 0 Then Exit Sub
    ItemSelecionado = Item.Index
End Sub

Private Sub LIST_ITENS_FORMA_PGTO_KeyDown(KeyCode As Integer, Shift As Integer)
'''    If KeyCode = vbKeyDelete Then
'''        For i = 1 To LIST_ITENS_FORMA_PGTO.ListItems.Count
'''            If LIST_ITENS_FORMA_PGTO.ListItems.Item(i).Selected = True Then
'''                If MsgBox("Excluir o Item...?", vbYesNo + vbQuestion + vbDefaultButton1, "Responda-me") = vbYes Then
'''                    LIST_ITENS_FORMA_PGTO.ListItems.Remove (i)
'''                    cont = cont - 1
'''                    txtIntervalo.SetFocus
'''                    Exit For
'''                End If
'''            End If
'''        Next
'''    End If
End Sub


Private Sub LIST_ITENS_FORMA_PGTO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If LIST_ITENS_FORMA_PGTO.ListItems.Count > 0 Then
        LIST_ITENS_FORMA_PGTO_DblClick
    End If
End If
End Sub

Private Sub MskVcto_GotFocus()
Call SelText(MskVcto)
End Sub

Private Sub MskVcto_KeyPress(KeyAscii As Integer)
Dim total As Double
Dim aux_Intervalo_anterior As Integer
Dim data_Anterior As String

On Error GoTo Trata_Erro

total = 0
aux_Intervalo_anterior = CCur(txtIntervalo.Text)

    If KeyAscii = 13 Then
        If Not IsNumeric(Me.txtNroNF.Text) Then
            MsgBox "Informe o N�mero da Nota Fiscal...!", vbInformation, "Aviso"
            txtNroNF.SetFocus
            Exit Sub
        End If
    
        If IsDate(MskVcto.Text) = False Then
            MsgBox "Data Inv�lida...!", vbInformation, "Aviso"
            MskVcto.SetFocus
            Exit Sub
        End If
        
        If CCur(txtQtParc.Text) = 0 Then
            VlParcela = lbl_Sub_Tot_Pedido.Caption
        Else
            VlParcela = FormatNumber(lbl_Sub_Tot_Pedido.Caption / CCur(txtQtParc.Text), 2)
        End If
        ContParc = CCur(txtQtParc.Text)
                       
        strData = Format(Date, "dd/mm/yyyy")
        
        Total_Parcelas = 0
        
        If flag_altera_forma_pgto_Boleto = False Then
            For i = 1 To CInt(txtQtParc.Text)
                i = LIST_ITENS_FORMA_PGTO.ListItems.Count + 1
                LIST_ITENS_FORMA_PGTO.ListItems.Add i, , i
                
                If ContParc = LIST_ITENS_FORMA_PGTO.ListItems.Count Then
                    str_Diferenca_nas_Parcelas = Format(CCur(lblTotCompra.Caption) - CCur(Total_Parcelas), "0.00")
                    LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(1) = Format(str_Diferenca_nas_Parcelas, "###,##0.00")
                Else
                    LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(1) = Format(txtVlr_Parcela.Text, "###,##0.00")
                End If
                
                If i = 1 Then
                    'calcula intervalo em dias
                    strIntervalo = DateDiff("d", Date, MskVcto.Text)
                    LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(3) = strIntervalo
                    LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(2) = MskVcto.Text
                Else
                    'exibe data de vencimento atrav�s do intervalo de dias
                    'LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(2) = Format(DateAdd("d", CCur(Me.txtIntervalo), strdata), "dd/mm/yyyy")
                    LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(2) = Format(DateAdd("d", aux_Intervalo_anterior, Date), "dd/mm/yyyy")
                    LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(3) = aux_Intervalo_anterior
                    strData = Format(LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(2), "dd/mm/yyyy")
                End If
                aux_Intervalo_anterior = aux_Intervalo_anterior + CCur(txtIntervalo.Text)
                Total_Parcelas = Total_Parcelas + CCur(LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(1))
            Next
        Else
            For i = LIST_ITENS_FORMA_PGTO.ListItems.Count To 1 Step -1
                If LIST_ITENS_FORMA_PGTO.ListItems(i).Selected = True Then
                    strIntervalo = DateDiff("d", Date, MskVcto.Text)
                    txtIntervalo.Text = strIntervalo
                    LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(3) = strIntervalo
                    LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(2) = MskVcto.Text

                    Exit For
                End If
            Next i
        End If
        
        flag_altera_forma_pgto_Boleto = False
        txtQtParc.Enabled = False
        txtIntervalo.Enabled = False
        txtVlr_Parcela.Enabled = False
        MskVcto.Enabled = False
        LIST_ITENS_FORMA_PGTO.SetFocus

        If MsgBox("Salvar Dados..? ", vbQuestion + vbYesNo + vbDefaultButton2, "Responda-me") = vbYes Then
            Dim SEQUENCIA As Double
            SEQUENCIA = Select_Max("Rece_Paga", "SEQUENCIA")
            'For i = LIST_ITENS_FORMA_PGTO.ListItems.Count To 1 Step -1  'decrecente
            Cnn.BeginTrans
            gTransacao = True
            For i = 1 To LIST_ITENS_FORMA_PGTO.ListItems.Count
                sql = "Insert into Rece_Paga  (CODIGO,DATA,DESCRICAO,VALOR,TIPO_MOVIMENTACAO,TP_FAVORECIDO, SEQUENCIA,PEDIDO,COD_CLIENTE) values ( "
                sql = sql & CodCliente & ","
                'vencimento
                sql = sql & "'" & Format(LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(2), "mm/dd/yyyy") & "',"
                sql = sql & "'BOLETO BANC�RIO - Pedido n� " & CodSeq & " Parcela " & i & "/" & LIST_ITENS_FORMA_PGTO.ListItems.Count & "',"
                sql = sql & "'" & Troca_Virg_Zero(Format(LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(1), "0.00")) & "',"
                sql = sql & "'R',"
                sql = sql & "'C',"
                sql = sql & SEQUENCIA & ","
                sql = sql & CodSeq & ","
                sql = sql & CodCliente & ")"
                SEQUENCIA = SEQUENCIA + 1
                Cnn.Execute sql
                
                'FORMA_PGTO
                sql = "INSERT INTO FORMA_PGTO VALUES ("
                sql = sql & CodSeq & ","
                sql = sql & CodCliente & ","
                sql = sql & "6" & ","   'cod_FORMAPGTO
                sql = sql & "'S'," 'SAIDA
                'sql = sql & "NULL,"
                sql = sql & "'" & i & "/" & LIST_ITENS_FORMA_PGTO.ListItems.Count & "'," ' PARCELA
                'sql = sql & "'" & Format(Date, "mm/dd/yyyy") & "',"
                sql = sql & "'" & Format(LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(2), "mm/dd/yyyy") & "',"
                'sql = sql & "'" & Troca_Virg_Zero(Format(lblTotCompra.Caption, "0.00")) & "',"
                sql = sql & "'" & Troca_Virg_Zero(Format(LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(1), "0.00")) & "',"
                sql = sql & "'" & UCase(NomeUsuario) & "')"
                Cnn.Execute sql
            Next
            
            sql = " UPDATE SAIDAS_PRODUTO SET "
            sql = sql & " STATUS_SAIDA = 'S',"
            sql = sql & " NF = " & Me.txtNroNF.Text
            sql = sql & " WHERE SEQUENCIA = " & CodSeq
            Cnn.Execute sql
 
            Cnn.CommitTrans
            gTransacao = False
            DoEvents
            
            Call ImprimimeBoletos(CodCliente, CodSeq)
            
            If CCur(txtNroNF.Text) > 0 Then
                Call GeraNFE
            Else
                If flagImpFiscalSelecionada = True Then
                    If MsgBox("Dados Atualizados com Sucesso." & vbNewLine & vbNewLine & "Impressora Fiscal Pronta para Imprimir...?", vbQuestion + vbYesNo + vbDefaultButton2, "Responda-me") = vbYes Then
                        frm_CNPJ_CPF.Show 1
                        Call ImprimeCupom_Fiscal(CodSeq, "0,00", "0,00", "0,00", "0,00", "0,00", str_cnpj_cpf, lblTotCompra.Caption)
                    End If
                Else
                    MsgBox "Dados Atualizados com Sucesso.", vbInformation, "Aviso"
                End If
            End If
            
            Unload Me
            Exit Sub
        
        End If
    End If
        
    
    Select Case KeyAscii
        Case 48 To 57
        Case 8
    Case Else
        KeyAscii = 0
    End Select

Exit Sub

Trata_Erro:

If gTransacao = True Then Cnn.RollbackTrans

End Sub

Private Sub GeraNFE()

Dim x As Integer
Dim Desconto As Double

With Frm_NFe
'    If IsNumeric(Me.txtCodSeq.Text) Then
'        .Carrega_Pedido (Me.txtCodSeq.Text)
'    Else
'        .Carrega_Pedido (StrNroPedidosNFE)
'    End If
'    '.Carrega_Colunas_Itens_Pedido
'    .lbl_Vlr_Desconto.Caption = Me.txtDesconto.Text
'    .lbl_Vlr_Frete.Caption = Me.txtVlrFrete.Text
'    .lbl_Vlr_Total_NF.Caption = Me.lblValorTTCompra.Caption
'    Desconto = Format(((txtDesconto.Text) / CCur(lblValorTTCompra.Caption)) * 100, "0.00")
'    For X = 1 To Me.List_Itens_Pedido.ListItems.Count
'        .List_Itens_Pedido.ListItems.Add X, , Me.List_Itens_Pedido.ListItems(X).Text
'
'        .List_Itens_Pedido.ListItems(X).SubItems(1) = Me.List_Itens_Pedido.ListItems(X).SubItems(1)
'        .List_Itens_Pedido.ListItems(X).SubItems(2) = GetCampo("select unidade from produto where codigo_interno = '" & Me.List_Itens_Pedido.ListItems(X).Text & "'", "unidade")
'        .List_Itens_Pedido.ListItems(X).SubItems(3) = List_Itens_Pedido.ListItems(X).SubItems(2)
'        .List_Itens_Pedido.ListItems(X).SubItems(4) = List_Itens_Pedido.ListItems(X).SubItems(3) 'unitario
'        .List_Itens_Pedido.ListItems(X).SubItems(5) = List_Itens_Pedido.ListItems(X).SubItems(4) 'total
'        .List_Itens_Pedido.ListItems(X).SubItems(6) = GetCampo("select codigo from produto where codigo_interno = '" & Me.List_Itens_Pedido.ListItems(X).Text & "'", "codigo")
'        .Lbl_Total_Itens.Caption = Me.Lbl_Total_Itens.Caption
'    Next X
'
'    .Show 1
End With


With FrmConsultaPedidos
    .MontaPedidoNfe (Me.txtCodSeq.Text)
    .Show 1
    '.cmd_Gerar_NFe.SetFocus
End With

End Sub

Private Sub ImprimimeBoletos(Optional ByVal cod_cli As Double, Optional ByVal Pedido As Double)

On Error GoTo tRATA_ERR

Dim CobreBemX As CobreBemX.ContaCorrente

Dim Boleto As Object
Dim i As Integer


Dim varUltimoNN As String   'dar um select no banco do ultimo numero de boleto gerado

If Dir(App.Path & "\Remessa", vbDirectory) = "" Then
    On Error Resume Next
    MkDir App.Path & "\Remessa"
End If

Set CobreBemX = New ContaCorrente

Dim diasProtestoBoleto As Integer
Dim PercentualJurosDiaAtraso, PercentualMultaAtraso As Double
Dim str_instrucao As String
Dim str_sequencia_processamento_remessa As String

PercentualJurosDiaAtraso = 0
PercentualMultaAtraso = 0

sql = "SELECT * FROM CONTA_CORRENTE_BOLETO "
Set Rstemp = New ADODB.Recordset
Rstemp.Open sql, Cnn, 1, 2
If Rstemp.RecordCount > 0 Then
    While Not Rstemp.EOF
        'Banco = CobreBemX.NumeroBanco
        CobreBemX.ArquivoLicenca = Rstemp!ArquivoLicenca
        CobreBemX.CodigoAgencia = Rstemp!AgenciaCEDENTE
        CobreBemX.NumeroContaCorrente = Rstemp!ContaCorrenteCedente
        CobreBemX.CodigoCedente = Rstemp!CodigoCedente
        CobreBemX.InicioNossoNumero = Rstemp!InicioNossoNumero
        CobreBemX.FimNossoNumero = Rstemp!FimNossoNumero
        
        'CobreBemX.ProximoNossoNumero = Left(Rstemp!ProximoNossoNumero, 8)
        CobreBemX.ProximoNossoNumero = Rstemp!ProximoNossoNumero
        
        If Not IsNull(Rstemp!PercentualJurosDiaAtraso) Then
            PercentualJurosDiaAtraso = CCur(Rstemp!PercentualJurosDiaAtraso)
        End If
        
        If Not IsNull(Rstemp!PercentualMultaAtraso) Then
            PercentualMultaAtraso = CCur(Rstemp!PercentualMultaAtraso)
        End If
        
        CobreBemX.PadroesBoleto.PadroesBoletoImpresso.ArquivoLogotipo = Rstemp!CAMINHO_LOGOTIPO_BOLETO_IMP
        If Not IsNull(Rstemp!DIAS_PROTESTO) Then
            diasProtestoBoleto = Rstemp("DIAS_PROTESTO")
        Else
            diasProtestoBoleto = 0
        End If
        
        If Not IsNull(Rstemp!INSTRUCAO1) Then
            str_instrucao = Rstemp("INSTRUCAO1")
        End If
        
        If Not IsNull(Rstemp!Instrucao2) Then
            str_instrucao = str_instrucao & "<br> " & Rstemp("INSTRUCAO2")
        End If
        
        If Not IsNull(Rstemp!Instrucao3) Then
            str_instrucao = str_instrucao & "<br> " & Rstemp("INSTRUCAO3")
        End If
        
        If Not IsNull(Rstemp!SEQUENCIA_PROCESSAMENTO) Then
            str_sequencia_processamento_remessa = (Rstemp("SEQUENCIA_PROCESSAMENTO") + 1)
        Else
            str_sequencia_processamento_remessa = "1"
            sql = "UPDATE CONTA_CORRENTE_BOLETO SET SEQUENCIA_PROCESSAMENTO = 1 "
            Cnn.Execute sql
        End If
        Rstemp.MoveNext
    Wend
Else
   ' CobreBemX.CodigoAgencia = "8088"
   ' CobreBemX.NumeroContaCorrente = "05663-8"
   ' CobreBemX.CodigoCedente = "92082835"
   ' CobreBemX.InicioNossoNumero = "00000001"
   ' CobreBemX.FimNossoNumero = "99999999"
   ' CobreBemX.ProximoNossoNumero = "00001000"
End If

Rstemp.Close

str_sequencia_processamento_remessa = Format(str_sequencia_processamento_remessa, "0000")

'para gerar arquivo remessa
'CobreBemX.ArquivoRemessa.Arquivo = "CDR" & Format(Date, "DD-MM-YYYY") & "_" & str_sequencia_processamento_remessa & ".txt" '
'CobreBemX.ArquivoRemessa.Arquivo = Format(Date, "DDMM") & "_" & str_sequencia_processamento_remessa & ".txt"
CobreBemX.ArquivoRemessa.Arquivo = Format(Date, "DDMM") & str_sequencia_processamento_remessa & ".txt"
CobreBemX.ArquivoRemessa.Diretorio = App.Path & "\Remessa\"
CobreBemX.ArquivoRemessa.Layout = "CNAB400"
'CobreBemX.ArquivoRemessa.SEQUENCIA = "0000001" 'GRAVA A SEQUENCIA PARA O BANCO ITAU
CobreBemX.ArquivoRemessa.SEQUENCIA = str_sequencia_processamento_remessa 'GRAVA A SEQUENCIA PARA O BANCO ITAU
CobreBemX.PadroesBoleto.PadroesBoletoImpresso.CaminhoImagensCodigoBarras = App.Path & "\ImagensBoleto\"


sql = "SELECT RECE_PAGA.PEDIDO,RECE_PAGA.SEQUENCIA,RECE_PAGA.DATA,CLIENTE.CODIGO,CLIENTE.RAZAO_SOCIAL,CLIENTE.ENDERECO_PRINCIPAL,"
sql = sql & " CLIENTE.NRO_END_PRINCIPAL,CLIENTE.CEP_PRINCIPAL,CLIENTE.CGC_CPF,CLIENTE.BAIRRO_END_PRINCIPAL,CLIENTE.CIDADE_END_PRINCIPAL,"
sql = sql & " CLIENTE.UF_END_PRINCIPAL,RECE_PAGA.VALOR "
sql = sql & " FROM RECE_PAGA,CLIENTE WHERE RECE_PAGA.COD_CLIENTE=CLIENTE.CODIGO "
sql = sql & " and RECE_PAGA.COD_CLIENTE=" & cod_cli
sql = sql & " and RECE_PAGA.PEDIDO=" & Pedido
sql = sql & " ORDER BY 2 "
Set Rstemp = New ADODB.Recordset
Rstemp.Open sql, Cnn, 1, 2
If Rstemp.RecordCount > 0 Then
    Rstemp.MoveLast
    Rstemp.MoveFirst
    
    For i = 1 To Rstemp.RecordCount
        Set Boleto = CobreBemX.DocumentosCobranca.Add
                
        'Especie Doc trocar de RC para DMI
        Boleto.TipoDocumentoCobranca = "DM"
        
        If diasProtestoBoleto > 0 Then
            Boleto.DiasProtesto = diasProtestoBoleto
        End If
        If IsNumeric(Me.txtNroNF.Text) = True Then
            If CCur(txtNroNF.Text) = 0 Then
                Boleto.NumeroDocumento = "PI. " & Format(Pedido, "0000")
            Else
                Boleto.NumeroDocumento = "N.F " & Format(Me.txtNroNF.Text, "0000")
            End If
        Else
            Boleto.NumeroDocumento = "PI. " & Format(Pedido, "0000")
        End If
        Boleto.NomeSacado = VerificaNulo(Rstemp!RAZAO_SOCIAL)
        If Not IsNull(Rstemp!CGC_CPF) Then
            Boleto.CPFSacado = Rstemp!CGC_CPF
        Else
            Boleto.CPFSacado = "00.000.000/0000-00"
        End If
        If Not IsNull(Rstemp!NRO_END_PRINCIPAL) Then
            Boleto.EnderecoSacado = Replace(Rstemp!ENDERECO_PRINCIPAL, ",", "") & ", " & Rstemp!NRO_END_PRINCIPAL
        Else
            Boleto.EnderecoSacado = VerificaNulo(Rstemp!ENDERECO_PRINCIPAL)
        End If
        Boleto.BairroSacado = VerificaNulo(Rstemp!BAIRRO_END_PRINCIPAL)
        Boleto.CidadeSacado = VerificaNulo(Rstemp!CIDADE_END_PRINCIPAL)
        Boleto.EstadoSacado = VerificaNulo(Rstemp!UF_END_PRINCIPAL)
        If Not IsNull(Rstemp!CEP_PRINCIPAL) Then
            Boleto.CepSacado = Replace(Rstemp!CEP_PRINCIPAL, "-", "")
        Else
            Boleto.CepSacado = Replace("00000-000", "-", "")
        End If
        Boleto.DataDocumento = Format(Date, "DD/MM/YYYY")
        Boleto.DataVencimento = Rstemp!Data
        Boleto.ValorDocumento = Format(Rstemp!valor, "0.00")
        If PercentualJurosDiaAtraso <> 0 Then
            Boleto.PercentualJurosDiaAtraso = PercentualJurosDiaAtraso '0.33
        Else
            Boleto.PercentualJurosDiaAtraso = 0.33
        End If
        If PercentualMultaAtraso <> 0 Then
            Boleto.PercentualMultaAtraso = PercentualMultaAtraso
        Else
            Boleto.PercentualMultaAtraso = 2
        End If
        Boleto.PercentualDesconto = 0
        Boleto.ValorOutrosAcrescimos = 0
        Boleto.PadroesBoleto.Demonstrativo = "Referente a compras na loja<br><b>" & "Parcela " & i & "/" & Rstemp.RecordCount & "</b>"
        'Boleto.PadroesBoleto.InstrucoesCaixa = "<br><br>N�o dispensar juros e multa ap�s o vencimento"
        Boleto.PadroesBoleto.InstrucoesCaixa = str_instrucao
        'Boleto.PadroesBoleto.Demonstrativo = "<b>REFERENTE AO SISTEMA " & UCase(rd.Item("nome_Sistema")) & "</b><br>PROTESTAR 3 DIAS APOS O VENCIMENTO <BR>NAO RECEBER APOS O DIA " & Convert.ToDateTime(Boleto.DataVencimento).AddDays(3) & _
        '"<br>NAO DISPENSAR JUROS E MULTA APOS O VENCIMENTO</b>"
        'Boleto.PadroesBoleto.InstrucoesCaixa = "<b>REFERENTE AO SISTEMA " & UCase(rd.Item("nome_Sistema")) & "</b><br>PROTESTAR 3 DIAS APOS O VENCIMENTO <BR>NAO RECEBER APOS O DIA " & Convert.ToDateTime(Boleto.DataVencimento).AddDays(3) & _
        '"<br>NAO DISPENSAR JUROS E MULTA APOS O VENCIMENTO</b>"
        CobreBemX.CalcularDadosBoletos
        
        If Left(CobreBemX.NumeroBanco, 3) = "237" Then    'BRADESCO
            sql = "UPDATE RECE_PAGA SET NOSSO_NUMERO='" & Left(Boleto.NossoNumero, 11) & "'"
        ElseIf Left(CobreBemX.NumeroBanco, 3) = "341" Then    'itau
            sql = "UPDATE RECE_PAGA SET NOSSO_NUMERO='" & Boleto.NossoNumero & "'"    'NRO TESTE ARQUIVO RETORNO TESTADO = 231018155
        End If
        sql = sql & " WHERE PEDIDO = " & CodSeq & " AND SEQUENCIA = " & Rstemp!SEQUENCIA
        Cnn.Execute sql
        Rstemp.MoveNext
    Next i
End If

Rstemp.Close
Set Rstemp = Nothing

CobreBemX.CalcularDadosBoletos

Banco = Left(CobreBemX.NumeroBanco, 3)

ULTIMO_NOSSO_NUMERO_GEROU = Boleto.NossoNumero

If Banco = "237" Then
    'bradesco
    sql = "UPDATE CONTA_CORRENTE_BOLETO SET PROXIMONOSSONUMERO='" & Left(Boleto.NossoNumero, 11) & "'"
    Cnn.Execute sql

ElseIf Banco = "341" Then
    'itau
    sql = "UPDATE CONTA_CORRENTE_BOLETO SET PROXIMONOSSONUMERO='" & Left(Boleto.NossoNumero, 8) & "'"
    Cnn.Execute sql

ElseIf Banco = "399" Then
    'HSBC
    sql = "UPDATE CONTA_CORRENTE_BOLETO SET PROXIMONOSSONUMERO='" & Left(Boleto.NossoNumero, 11) & "'"
    Cnn.Execute sql
End If

sql = "UPDATE CONTA_CORRENTE_BOLETO SET SEQUENCIA_PROCESSAMENTO = '" & str_sequencia_processamento_remessa & "'" 'SEQUENCIA_PROCESSAMENTO + 1 "
Cnn.Execute sql

CobreBemX.ImprimeBoletos
'CobreBemX.GravaArquivoRemessa

Set CobreBemX = Nothing

 
Exit Sub

tRATA_ERR:
    If Err.Number <> 0 Then
        If Err.Number = 91 Then
            'MsgBox "Ocorreu um Erro e o boleto n�o foi gerado." & vbNewLine & "V� na tela de Cadastro de Cedente e configure corretamente os campos existentes... ", vbCritical, "Aviso"
            MsgBox "Ocorreu um Erro e o boleto n�o foi gerado." & vbNewLine & vbNewLine & "Compartilhar a pasta c:\Sistema SisAdven", vbCritical, "Aviso"
        Else
            MsgBox "Ocorreu um Erro e o boleto n�o foi gerado." & "N� Erro: " & "  Descri��o: " & Err.Description, vbCritical, "Aviso"
        End If
        
        Err.Clear
        Screen.MousePointer = 1
    End If


End Sub



Private Sub Totais_Opcao_Boleto()

Total_Parcelas = 0

    For i = 1 To LIST_ITENS_FORMA_PGTO.ListItems.Count
        Total_Parcelas = Total_Parcelas + CCur(LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(1))
    Next i
    
    lbl_Total_Parcelas.Caption = "Total das Parcelas : " + Format(Total_Parcelas, "###,##0.00")
End Sub

Private Sub txt_Desconto_GotFocus()
Call SelText(txt_Desconto)
End Sub


Private Sub txt_Desconto_KeyPress(KeyAscii As Integer)
 '   If KeyAscii = 13 Then
 '       If IsNumeric(txt_Desconto.Text) Then
 '           txt_Desconto.Text = Format(txt_Desconto.Text, "0.00")
 '       Else
 '           txt_Desconto.Text = "0,00"
 '       End If
 '       lblTotPedido.Caption = Format(CCur(lbl_Sub_Tot_Pedido.Caption) - CCur(txt_Desconto.Text), "0.00")
 '   End If
    
    If KeyAscii = 13 Then
        lblTeclas.Caption = "[ Esc ] Salvar, Fechar Janela, ou Cancelar "
        txt_Desconto_LostFocus
        SendKeys "{Tab}"
    End If
    
    KeyAscii = ConsisteTeclaValorNumerico(txt_Desconto.Text, KeyAscii)
End Sub


Private Sub txt_Desconto_LostFocus()
    Call CalcValRecebidos
End Sub

Private Sub txtCodSeq_GotFocus()
Call SelText(txtCodSeq)
End Sub


Private Sub txtCodSeq_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CodSeq = Val(txtCodSeq.Text)
    sql = "Select sequencia, CODIGO_CLIENTE, total_saida from saidas_produto "
    sql = sql & " where SEQUENCIA =  " & CodSeq
    sql = sql & " and STATUS_SAIDA = 'S' "
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open sql, Cnn, 1, 2
    If Rstemp.RecordCount > 0 Then
        fme_Totais.Enabled = False
        MsgBox "Pedido j� Recebido..", vbInformation, "Aviso"
        txtCodSeq_GotFocus
        Rstemp.Close
        Exit Sub
    End If
        
    Rstemp.Close
    Set Rstemp = Nothing
    
    txtCodSeq_LostFocus
End If

    Select Case KeyAscii
        Case 48 To 57
        Case 8
        Case Else
            KeyAscii = 0
    End Select
End Sub


Private Sub txtCodSeq_LostFocus()
Call MontaPedido
End Sub



Private Sub txtIntervalo_GotFocus()
Call SelText(txtIntervalo)
End Sub

Private Sub txtIntervalo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not IsNumeric(txtIntervalo.Text) Then
        MskVcto.Enabled = False
        MskVcto.Mask = "##/##/####"
        Exit Sub
    End If
    
    If LIST_ITENS_FORMA_PGTO.ListItems.Count = 0 Then
        Me.txtVlr_Parcela.Text = FormatNumber(lbl_Sub_Tot_Pedido.Caption / CCur(txtQtParc.Text), 2)
        ContParc = CCur(txtQtParc.Text)
    Else
        Dim SOMA_PARCELAS As Double
        SOMA_PARCELAS = 0
        Total_Parcelas = 0
        For i = 1 To LIST_ITENS_FORMA_PGTO.ListItems.Count
            Total_Parcelas = Total_Parcelas + CCur(LIST_ITENS_FORMA_PGTO.ListItems(i).SubItems(1))
            If CCur(txtQtParc.Text) = CCur(LIST_ITENS_FORMA_PGTO.ListItems.Count + 1) Then
                txtVlr_Parcela.Text = Format(CCur(lblTotCompra.Caption) - Total_Parcelas, "0.00")
            Else
                 Me.txtVlr_Parcela.Text = FormatNumber(lbl_Sub_Tot_Pedido.Caption / CCur(txtQtParc.Text), 2)
            End If
        Next i
        
       ' If CCur(txtQtParc.Text) = CCur(LIST_ITENS_FORMA_PGTO.ListItems.Count) Then
       '     txtVlr_Parcela.Text = Format(CCur(lblTotCompra.Caption) - Total_Parcelas, "0.00")
       ' End If
        
        
    End If
    
    strData = Format(Date, "dd/mm/yyyy")
    
    DATA_VENCIMENTO = Format(DateAdd("d", txtIntervalo.Text, strData), "dd/mm/yyyy")
    
    MskVcto.Text = Format(DATA_VENCIMENTO, "dd/mm/yyyy")
    'MskVcto.Enabled = True
    'MskVcto.SetFocus
    txtVlr_Parcela.Enabled = True
    SendKeys "{tab}"
End If
End Sub

Private Sub txtNroNF_GotFocus()
Call SelText(txtNroNF)
End Sub


Private Sub txtNroNF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(Me.txtNroNF.Text) = True Then
        LIST_ITENS_FORMA_PGTO.ListItems.Clear
        txtQtParc.Enabled = True
        txtQtParc.SetFocus
    End If
    Select Case KeyAscii
    Case 48 To 57
    Case 8
    Case Else
        KeyAscii = 0
    End Select
End Sub


Private Sub txtPendente_GotFocus()
Call SelText(txtPendente)
End Sub


Private Sub txtPendente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lblTeclas.Caption = "[ Esc ] Salva, Fechar Janela, ou Cancela "
        Call txtPendente_LostFocus
    End If

    KeyAscii = ConsisteTeclaValorNumerico(txtPendente.Text, KeyAscii)
End Sub


Private Sub txtPendente_LostFocus()
Call CalcValRecebidos
End Sub

Private Sub txtQtParc_Change()

If IsNumeric(txtQtParc.Text) = True Then
    
    
    If IsNumeric(txtQtParc.Text) = True Then
        If CCur(txtQtParc.Text) > 1 Then
            lbl_intervalo.Visible = True
            txtIntervalo.Visible = True
            txtIntervalo.Text = "30"
        ElseIf CCur(txtQtParc.Text) = 0 Then
            lbl_intervalo.Visible = False
            txtIntervalo.Visible = False
        Else
            lbl_intervalo.Visible = False
            txtIntervalo.Visible = False
            txtIntervalo.Text = "30"
        End If
    End If
End If
End Sub

Private Sub txtQtParc_GotFocus()
Call SelText(txtQtParc)
End Sub

Private Sub txtQtParc_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If IsNumeric(txtQtParc.Text) = True Then
            txtQtParc.Enabled = False
            If CCur(txtQtParc.Text) > 1 Then
                txtIntervalo.Enabled = True
                txtIntervalo.SetFocus
            ElseIf CCur(txtQtParc.Text) = 0 Then
                txtQtParc.Enabled = True
                txtQtParc.SetFocus
            ElseIf CCur(txtQtParc.Text) = 1 Then
                'MskVcto.Text = Format(Date, "dd/mm/yyyy")
                'MskVcto.Enabled = True
                'MskVcto.SetFocus
                txtVlr_Parcela.Text = lbl_Sub_Tot_Pedido.Caption
                'MskVcto.Text = Format(Date, "dd/mm/yyyy")
                MskVcto.Text = Format(DateAdd("d", CCur(Me.txtIntervalo), Date), "dd/mm/yyyy")
                txtVlr_Parcela.Enabled = True
                txtVlr_Parcela.SetFocus
            End If
        End If
    End If
    Select Case KeyAscii
    Case 48 To 57
    Case 8
    Case Else
      KeyAscii = 0
    End Select
End Sub


Private Sub txtVlr_Parcela_GotFocus()
Call SelText(txtVlr_Parcela)
End Sub


Private Sub txtVlr_Parcela_KeyPress(KeyAscii As Integer)
On Error GoTo trataAqui
    
    If KeyAscii = 13 Then
        If txtVlr_Parcela.Text = "" Then
            txtVlr_Parcela.Text = "0"
        End If
        If txtVlr_Parcela.Text <> "" Or Val(txtVlr_Parcela) <> 0 Then
            If IsNumeric(txtVlr_Parcela.Text) = True Then
                txtVlr_Parcela.Text = Format(txtVlr_Parcela.Text, "0.00")
                MskVcto.Enabled = True
                SendKeys "{tab}"
            Else
                MsgBox "Valor da Parcela Inv�lido...!", vbCritical, "Erro"
                txtVlr_Parcela.SetFocus
            End If
        Else
            txtVlr_Parcela.Text = "0.00"
            SendKeys "{tab}"
        End If
    End If
    
    KeyAscii = ConsisteTeclaValorNumerico(txtVlr_Parcela.Text, KeyAscii)

Exit Sub

trataAqui:
    Err.Clear

End Sub


Private Sub txtVlrCartCredito_GotFocus()
Call SelText(txtVlrCartCredito)
End Sub


Private Sub txtVlrCartCredito_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        lblAjuda2.Caption = "[ Esc ] Salvar, Fechar Janela, ou Cancelar "
        Call txtVlrCartCredito_LostFocus
        SendKeys "{Tab}"
    End If
    
    KeyAscii = ConsisteTeclaValorNumerico(txtVlrCartCredito, KeyAscii)
End Sub


Private Sub txtVlrCartCredito_LostFocus()
Call CalcValRecebidos
End Sub

Private Sub txtVlrCartEletron_GotFocus()
Call SelText(txtVlrCartEletron)
End Sub


Private Sub txtVlrCartEletron_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lblTeclas.Caption = "[ Esc ] Salvar, Fechar Janela, ou Cancelar "
        Call txtVlrCartEletron_LostFocus
        SendKeys "{Tab}"
    End If
    
    KeyAscii = ConsisteTeclaValorNumerico(txtVlrCartEletron, KeyAscii)
End Sub


Private Sub txtVlrCartEletron_LostFocus()
Call CalcValRecebidos
End Sub


Private Sub txtVlrCartLoja_GotFocus()
Call SelText(txtVlrCartLoja)
End Sub


Private Sub txtVlrCartLoja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lblTeclas.Caption = "[ Esc ] Salvar, Fechar Janela, ou Cancelar "
        Call txtVlrCartLoja_LostFocus
        SendKeys "{Tab}"
    End If
    
    KeyAscii = ConsisteTeclaValorNumerico(txtVlrCartLoja, KeyAscii)
End Sub


Private Sub txtVlrCartLoja_LostFocus()
Call CalcValRecebidos
End Sub

Private Sub txtVlrCheques_GotFocus()
Call SelText(txtVlrCheques)

End Sub


Private Sub txtVlrCheques_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        lblTeclas.Caption = "[ Esc ] Salvar, Fechar Janela, ou Cancelar "
        Call txtVlrCheques_LostFocus
        SendKeys "{Tab}"
    End If
    
    KeyAscii = ConsisteTeclaValorNumerico(txtVlrCheques, KeyAscii)
End Sub


Private Sub txtVlrCheques_LostFocus()
Call CalcValRecebidos
End Sub

Private Sub txtVlrdinheiro_GotFocus()
Call SelText(txtVlrdinheiro)
End Sub

Private Sub txtVlrdinheiro_KeyPress(KeyAscii As Integer)
saldo = 0

TotalPedido = 0

If KeyAscii = 13 Then
    lblTeclas.Caption = "[ Esc ] Salvar, Fechar Janela, ou Cancelar "
    txtVlrdinheiro_LostFocus
    SendKeys "{Tab}"
End If

KeyAscii = ConsisteTeclaValorNumerico(txtVlrdinheiro, KeyAscii)
End Sub


Private Sub txtVlrdinheiro_LostFocus()
Call CalcValRecebidos
End Sub




