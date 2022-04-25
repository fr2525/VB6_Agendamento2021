VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAgenda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agenda Pet Shop"
   ClientHeight    =   8925
   ClientLeft      =   3390
   ClientTop       =   3495
   ClientWidth     =   14535
   ForeColor       =   &H00008000&
   Icon            =   "frmAgenda.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   14535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_adicionar 
      Caption         =   "Novo Atendimento"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   4200
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Criar novo atendimento"
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1395
   End
   Begin VB.CommandButton cmd_Sair 
      Caption         =   "Sair"
      Height          =   855
      Left            =   13260
      Picture         =   "frmAgenda.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Sair do sistema"
      Top             =   480
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1230
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameCadastros 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cadastros"
      Height          =   1305
      Left            =   6000
      TabIndex        =   17
      Top             =   240
      Width           =   7095
      Begin VB.CommandButton cmd_relatos 
         Caption         =   "Relatórios"
         Height          =   855
         Left            =   5880
         Picture         =   "frmAgenda.frx":0106
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Relatórios"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Cmd_formas 
         Caption         =   "Formas Pgto"
         Height          =   855
         Left            =   4920
         Picture         =   "frmAgenda.frx":0208
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Cadastro de formas de pagamento"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_tipos 
         Caption         =   "Tipos de pets"
         Height          =   855
         Left            =   2040
         Picture         =   "frmAgenda.frx":0E4A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Cadastro de Tipos de Pets"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_Novo_cli 
         Caption         =   "Novo Cli"
         Height          =   735
         Left            =   5520
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmd_servicos 
         Caption         =   "Serviços"
         Height          =   855
         Left            =   3960
         Picture         =   "frmAgenda.frx":1F14
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Cadastro de Serviços"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_pets 
         Caption         =   "Pets"
         Height          =   855
         Left            =   3000
         Picture         =   "frmAgenda.frx":2356
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Cadastro de Pets"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_clientes 
         Caption         =   "Clientes"
         Height          =   855
         Left            =   1080
         Picture         =   "frmAgenda.frx":2A02
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cadastro de clientes "
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_ajustes 
         Caption         =   "Ajustes"
         Height          =   855
         Left            =   120
         Picture         =   "frmAgenda.frx":3644
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Ajustes"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   150
      TabIndex        =   1
      Top             =   600
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   -2147483647
      CalendarTitleBackColor=   -2147483632
      CalendarTitleForeColor=   16776960
      CalendarTrailingForeColor=   128
      Format          =   109248512
      CurrentDate     =   42902
      MaxDate         =   401768
      MinDate         =   36892
   End
   Begin VB.Frame frameDetalhe 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   6525
      Left            =   3600
      TabIndex        =   2
      Top             =   1680
      Width           =   6645
      Begin VB.Frame frameOperacao 
         BackColor       =   &H00C0FFC0&
         Height          =   1005
         Left            =   1680
         TabIndex        =   41
         Top             =   5460
         Width           =   2955
         Begin VB.CommandButton cmd_voltar 
            Caption         =   "Voltar"
            Height          =   615
            Left            =   180
            Picture         =   "frmAgenda.frx":3A86
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmd_gravar 
            Caption         =   "Gravar"
            Enabled         =   0   'False
            Height          =   615
            Left            =   1080
            Picture         =   "frmAgenda.frx":3FB8
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmd_limpar 
            Caption         =   "Limpar"
            Enabled         =   0   'False
            Height          =   615
            Left            =   1980
            Picture         =   "frmAgenda.frx":44EA
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox CmbPets 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   2445
      End
      Begin VB.ComboBox cmbServicos 
         Height          =   315
         ItemData        =   "frmAgenda.frx":4A1C
         Left            =   1290
         List            =   "frmAgenda.frx":4A1E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2640
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ComboBox CmbHorario 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmAgenda.frx":4A20
         Left            =   1230
         List            =   "frmAgenda.frx":4A22
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2160
         Width           =   1125
      End
      Begin VB.TextBox txtTipoAtend 
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
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   10
         Top             =   2640
         Width           =   5010
      End
      Begin VB.TextBox txtObserv 
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
         Height          =   720
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   4680
         Width           =   6180
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
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
         Left            =   4890
         MaxLength       =   20
         TabIndex        =   7
         Top             =   3090
         Width           =   1410
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000005&
         Caption         =   "Cuidados Especiais"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   240
         TabIndex        =   11
         Top             =   3540
         Width           =   6135
         Begin VB.Label lblEspecial1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   240
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1620
         End
      End
      Begin VB.TextBox txtHrSaida 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
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
         Left            =   5250
         MaxLength       =   2
         TabIndex        =   19
         Top             =   2130
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtMinSaida 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
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
         Left            =   5910
         MaxLength       =   2
         TabIndex        =   20
         Top             =   2130
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Dono"
         Height          =   1425
         Left            =   300
         TabIndex        =   23
         Top             =   90
         Width           =   6135
         Begin VB.ComboBox cmbDonos 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Text            =   "cmbDonos"
            Top             =   240
            Width           =   4545
         End
         Begin VB.TextBox txtProp 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   9
            Top             =   210
            Width           =   5130
         End
         Begin VB.Label lblFone2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fone2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   240
            Left            =   3360
            TabIndex        =   39
            Top             =   1080
            Width           =   660
         End
         Begin VB.Label lblFone1 
            BackStyle       =   0  'Transparent
            Caption         =   "Fone1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   240
            Left            =   1320
            TabIndex        =   38
            Top             =   1080
            Width           =   1860
         End
         Begin VB.Label lblEndereco 
            BackStyle       =   0  'Transparent
            Caption         =   "Endereco + bairro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   5055
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefones :"
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
            Height          =   240
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   1170
         End
      End
      Begin VB.Label lbltipo 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   40
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5640
         TabIndex        =   22
         Top             =   2160
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saída :"
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
         Height          =   240
         Left            =   4500
         TabIndex        =   18
         Top             =   2190
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horário :"
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
         Height          =   240
         Left            =   210
         TabIndex        =   16
         Top             =   2220
         Width           =   900
      End
      Begin VB.Label lbl_Animal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pet :"
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
         Height          =   270
         Left            =   630
         TabIndex        =   15
         Top             =   1740
         Width           =   480
      End
      Begin VB.Label lbl_TipoAtend 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serviço :"
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
         Height          =   240
         Left            =   165
         TabIndex        =   14
         Top             =   2640
         Width           =   930
      End
      Begin VB.Label lbl_Obseerv 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Observ :"
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
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   4380
         Width           =   900
      End
      Begin VB.Label lbl_Valor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Valor :"
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
         Left            =   3840
         TabIndex        =   12
         Top             =   3120
         Width           =   1080
      End
   End
   Begin VB.Frame framLista 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Atendimentos"
      Height          =   6705
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   13995
      Begin MSComctlLib.ListView LIst_Atendimentos 
         Height          =   6135
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   10821
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[DEL]    para Excluir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   435
      Left            =   10800
      TabIndex        =   25
      Top             =   8370
      Width           =   3375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Tecle   [ENTER]   para Incluir/Alterar/Baixar/Desfazer           "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   465
      Left            =   270
      TabIndex        =   26
      Top             =   8370
      Width           =   9315
   End
   Begin VB.Menu pMnuLista 
      Caption         =   "Selecione"
      Visible         =   0   'False
      Begin VB.Menu pMnuAlterar 
         Caption         =   "&Alterar"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu pMnuExcluir 
         Caption         =   "&Excluir"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu pmnuBaixar 
         Caption         =   "&Baixa atendimento"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu pmnuDesfaz 
         Caption         =   "&Desfazer atendimento"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu pmnuReceber 
         Caption         =   "&Receber"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu pmnuVacina 
         Caption         =   "&Vacina"
      End
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim aLista_Horario(48) As String
Dim iTipoOperacao As Integer
Dim lvListItems_Itens As MSComctlLib.ListItem
Dim bClick_Listview As Boolean

Dim gNomeEmpresa, gEnderecoEmpresa, gCGC_EMPRESA, _
    gCEP_EMPRESA, gFone1Empresa, gFone2Empresa, _
    gemailEmpresa
Dim Complete As New clsAutoComplete
Public gOpcao As Integer
Public bRecebido As Boolean
Public bVacina As Boolean
Public dProximaVacina As Date
Public sDescVacina As String
Public nValorRecebido As Currency
Public sFormaPagto As String

Const IND_HORARIO = 0
Const IND_SITUACAO = 1
Const IND_PET = 2
Const IND_TIPO_PET = 3
Const IND_NOME_DONO = 4
Const IND_TELEFONE = 5
Const IND_OBSERV = 6
Const IND_DESC_ATEND = 7
Const IND_VALOR = 8
Const IND_TIPO_ATEND = 9
Const IND_HORA_SAIDA = 10
Const IND_COD_DONO = 11
Const IND_VLR_REC = 12
Const IND_COD_PET = 13
Const IND_VACINA = 14

Private Sub Carrega_Colunas_Atendimentos()
    With LIst_Atendimentos
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Horário", 700, lvwColumnLeft
        .ColumnHeaders.Add 2, , "Situação", 1170, lvwColumnLeft
        .ColumnHeaders.Add 3, , "Pet", 1400, lvwColumnLeft
        .ColumnHeaders.Add 4, , "Tipo de Pet ", 1200, lvwColumnLeft
        .ColumnHeaders.Add 5, , "Dono ", 2700, lvwColumnLeft
        .ColumnHeaders.Add 6, , "Telefone ", 1200, lvwColumnLeft
        .ColumnHeaders.Add 7, , "Observ", 2000, lvwColumnLeft
        .ColumnHeaders.Add 8, , "Atendimento", 3200, lvwColumnLeft
        .ColumnHeaders.Add 9, , "Valor", 900, lvwColumnRight
        .ColumnHeaders.Add 10, , "tipo_atend", 0, lvwColumnLeft
        .ColumnHeaders.Add 11, , "Hr.Saida", 0, lvwColumnLeft
        .ColumnHeaders.Add 12, , "IdCli", 0, lvwColumnLeft
        .ColumnHeaders.Add 13, , "Vlr. Recebido", 1200, lvwColumnLeft
        .ColumnHeaders.Add 14, , "idAnimal", 0, lvwColumnLeft
        .ColumnHeaders.Add 15, , "Vacina", 0, lvwColumnLeft
    End With
End Sub

Private Sub MontaAtendimentos(pData As String)
    
    Dim I, X, y, z As Integer
    Dim sColor, sOldColor As String
    
    Dim sHora As String   'Para comparar com a hora do combo
    
    LIst_Atendimentos.ListItems.Clear
    Call sConectaBanco
    strSql = ""
    strSql = strSql & " SELECT A.dt_atend,a.idanimal,a.tipo_atend, a.HORA_SAIDA,A.OBSERVA, A.HORA_VACINA "
    strSql = strSql & ",B.id,B.ID_CLI,B.NOME,B.TIPO_ANI,C.DESCRICAO AS TIPOPET, D.DESCRICAO AS SERVICO "
    strSql = strSql & ", D.VACINA, A.VALOR,A.VALOR_RECEBIDO, E.RAZAO_SOCIAL "
    strSql = strSql & ",  E.ENDERECO_PRINCIPAL, E.NRO_END_PRINCIPAL "
    strSql = strSql & ", E.BAIRRO_END_PRINCIPAL, E.FONE1, e.fone2 "
    strSql = strSql & " FROM TAB_ATENDIMENTOS A , TAB_PETS B, tab_tipos_pets C, "
    strSql = strSql & " TAB_SERVICOS D, tab_clientes E"
    strSql = strSql & " WHERE A.DT_ATEND >= cdate('" & pData & " 00:00:00')"
    strSql = strSql & " AND A.DT_ATEND <= cdate('" & pData & " 23:59:58')"
    strSql = strSql & " AND A.idanimal = b.id "
    strSql = strSql & " AND b.id_cli = e.id "
    strSql = strSql & " AND b.TIPO_ANI = C.id "
    strSql = strSql & " AND A.tipo_atend = d.id "
    strSql = strSql & " ORDER BY a.dt_atend "
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    
    If Rstemp.RecordCount <> 0 Then
        Rstemp.MoveLast
        Rstemp.MoveFirst
        sHora = Mid(Rstemp!DT_ATEND, 12, 5)
    Else
        sHora = "99:99"
    End If
        'fmeListaPedidos.Visible = True
        '**** Aqui nós vamos preencher a lista de atendimentos
        '     deixando os horarios vazios do combo de horarios sem preencher com os dados
        LIst_Atendimentos.ListItems.Clear
        I = 0
        CmbHorario.ListIndex = I
        Do While I < CmbHorario.ListCount
            If CmbHorario.text = sHora Then  '*** horario com atendimento
                LIst_Atendimentos.ListItems.Add I + 1, , sHora
                'List_Atendimentos.ListItems(i + 1).Add i + 1, , sHora
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_SITUACAO) = "PENDENTE"
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_PET) = Trim(Rstemp!Nome)
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_TIPO_PET) = Trim(Rstemp!TIPOPET)
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_NOME_DONO) = Trim(Rstemp!RAZAO_SOCIAL)
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_TELEFONE) = IIf(IsNull(Rstemp!FONE1), "", Rstemp!FONE1)
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_OBSERV) = IIf(IsNull(Rstemp!observa), "", Rstemp!observa)
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_DESC_ATEND) = Trim(Rstemp!SERVICO)
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_VALOR) = Format(Rstemp!Valor, "##,##0.00")
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_TIPO_ATEND) = Trim(Rstemp!TIPO_ATEND)
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_HORA_SAIDA) = Rstemp!HORA_SAIDA
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_COD_DONO) = Rstemp!id_cli
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_VLR_REC) = IIf(IsNull(Rstemp!valor_recebido), "0", Format(Rstemp!valor_recebido, "###,##0.00"))
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_COD_PET) = Rstemp!idAnimal
                LIst_Atendimentos.ListItems(I + 1).SubItems(IND_VACINA) = IIf(IsNull(Rstemp!VACINA), "", Rstemp!VACINA)
                sOldColor = LIst_Atendimentos.ForeColor
                sColor = sOldColor
                
                If Trim(Left(Rstemp!HORA_SAIDA, 2)) <> "" Then
                    sColor = vbMagenta
                    LIst_Atendimentos.ListItems(I + 1).SubItems(IND_SITUACAO) = "ATENDIDO"
                End If
                If Rstemp!valor_recebido > 0 Then
                    sColor = &H8000&
                    LIst_Atendimentos.ListItems(I + 1).SubItems(IND_SITUACAO) = "RECEBIDO"
                End If
                If Rstemp!VACINA = "S" And LIst_Atendimentos.ListItems(I + 1).SubItems(IND_SITUACAO) = "ATENDIDO" And Rstemp!hora_vacina = "  :  " Then
                    sColor = vbBlue
                    LIst_Atendimentos.ListItems(I + 1).SubItems(IND_SITUACAO) = "VACINA"
                Else
                    If LIst_Atendimentos.ListItems(I + 1).SubItems(IND_SITUACAO) = "VACINA" And Rstemp!hora_vacina <> "  :  " Then
                        sColor = vbMagenta
                        LIst_Atendimentos.ListItems(I + 1).SubItems(IND_SITUACAO) = "ATENDIDO"
                    End If
                End If
                LIst_Atendimentos.ListItems(I + 1).ForeColor = sColor
                For z = 1 To 7
                    LIst_Atendimentos.ListItems(I + 1).ListSubItems(z).ForeColor = sColor
                Next
                I = I + 1
                If I < CmbHorario.ListCount Then
                    CmbHorario.ListIndex = I
                End If
            ElseIf CmbHorario.text < sHora Then
                LIst_Atendimentos.ListItems.Add I + 1, , CmbHorario.text
                LIst_Atendimentos.ListItems(I + 1).SubItems(1) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(2) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(3) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(4) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(5) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(6) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(7) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(8) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(9) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(10) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(11) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(12) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(13) = ""
                LIst_Atendimentos.ListItems(I + 1).SubItems(14) = ""
                I = I + 1
                If I < CmbHorario.ListCount Then
                    CmbHorario.ListIndex = I
                End If
            Else
                Rstemp.MoveNext
                If Rstemp.EOF Then
                    sHora = "99:99"
                Else
                    sHora = Mid(Rstemp!DT_ATEND, 12, 5)
                End If
            End If
        Loop
    'End If
         
    Rstemp.Close
    Set Rstemp = Nothing
    
End Sub


Private Sub Busca_Atendimento(pData As String)
    
    Call sConectaBanco
    strSql = ""
    strSql = strSql & " SELECT A.dt_atend,a.idanimal,a.tipo_atend,a.HORA_SAIDA, a.observa, A.HORA_VACINA "
    strSql = strSql & ",B.id,B.ID_CLI,B.NOME,B.TIPO_ANI,C.DESCRICAO AS TIPOPET"
    strSql = strSql & ", b.cuidados_especiais, D.DESCRICAO AS SERVICO "
    strSql = strSql & ", A.VALOR, A.VALOR_RECEBIDO, E.RAZAO_SOCIAL , E.ENDERECO_PRINCIPAL, E.NRO_END_PRINCIPAL "
    strSql = strSql & ", E.BAIRRO_END_PRINCIPAL, E.FONE1, E.fone2 "
    strSql = strSql & " FROM TAB_ATENDIMENTOS A , tab_pets B, tab_tipos_pets C, "
    strSql = strSql & " TAB_SERVICOS D, tab_clientes E"
    strSql = strSql & " WHERE A.DT_ATEND >= cdate('" & pData & "')"
    strSql = strSql & " AND A.DT_ATEND <= cdate('" & pData & "')"
    strSql = strSql & " AND A.idanimal = b.id "
    strSql = strSql & " AND b.id_cli = e.id "
    strSql = strSql & " AND b.TIPO_ANI = C.id "
    strSql = strSql & " AND A.tipo_atend = d.id "
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    If Rstemp.RecordCount > 0 Then
        Call Carrega_campos_texto
    Else
        'MsgBox "Sem Atendimentos para a data selecionada", vbOKOnly
        'fmeListaPedidos.Visible = False
        Call Limpa_Campos_Texto
    End If
    
    Rstemp.Close
    Set Rstemp = Nothing
    'Cnn.Close
    
End Sub

Private Sub Busca_Dono(pId_cli As Double)
    
    Call sConectaBanco
    strSql = ""
    strSql = strSql & " SELECT id, RAZAO_SOCIAL , ENDERECO_PRINCIPAL, NRO_END_PRINCIPAL  "
    strSql = strSql & ", BAIRRO_END_PRINCIPAL, fone1, fone2 "
    strSql = strSql & " FROM tab_clientes"
    strSql = strSql & " WHERE id = " & pId_cli
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    If Rstemp.RecordCount > 0 Then
        Call Carrega_Endereco
    Else
        'MsgBox "Sem Atendimentos para a data selecionada", vbOKOnly
        'fmeListaPedidos.Visible = False
        Call Limpa_Endereco
    End If
    
    Rstemp.Close
    Set Rstemp = Nothing
   ' Cnn.Close
    
End Sub
'
Private Sub Busca_tipo(pId As Double)
    
    Call sConectaBanco
    strSql = ""
    strSql = strSql & " SELECT descricao"
    strSql = strSql & " FROM tab_tipos_pets a, tab_pets b"
    strSql = strSql & " WHERE b.id = " & pId
    strSql = strSql & " AND a.id = b.tipo_ani"
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    If Rstemp.RecordCount > 0 Then
        lbltipo.Caption = Rstemp!Descricao
    Else
        lbltipo.Caption = ""
    End If
    
    Rstemp.Close
    Set Rstemp = Nothing
End Sub

'****


Private Sub Carrega_Endereco()
    lblEndereco.Caption = IIf(IsNull(Rstemp!ENDERECO_PRINCIPAL), "", Trim(Rstemp!ENDERECO_PRINCIPAL)) & _
                       " , " & IIf(IsNull(Rstemp!NRO_END_PRINCIPAL), "", Trim(Rstemp!NRO_END_PRINCIPAL)) & _
                       " - " & IIf(IsNull(Rstemp!BAIRRO_END_PRINCIPAL), "", Trim(Rstemp!BAIRRO_END_PRINCIPAL))
    lblFone1.Caption = IIf(IsNull(Rstemp!FONE1), "", Trim(Rstemp!FONE1))
    lblFone2.Caption = IIf(IsNull(Rstemp!fone2), "", Trim(Rstemp!fone2))
  
End Sub

Private Sub Limpa_Endereco()
    txtEndereco.text = ""
    txtBairro.text = ""
    txtFone1.text = ""
    txtFone2.text = ""
  
End Sub

Private Sub cmbDonos_Change()
    If cmbDonos.ListIndex <> -1 Then
        If Not fCarrega_Pets(cmbDonos.ItemData(cmbDonos.ListIndex)) Then
            CmbPets.Clear
        End If
        Call Busca_Dono(cmbDonos.ItemData(cmbDonos.ListIndex))
    End If
End Sub

Private Sub cmbDonos_GotFocus()
    Call SelText(cmbDonos)
End Sub

Private Sub cmbDonos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
       If cmbDonos.ListIndex <> -1 Then
           If Not fCarrega_Pets(cmbDonos.ItemData(cmbDonos.ListIndex)) Then
               CmbPets.Clear
               MsgBox "Sem Pets cadastrados para o cliente, favor cadastrar", vbOKOnly, "Aviso"
               cmbDonos.SetFocus
               Exit Sub
           End If
           Call Busca_Dono(cmbDonos.ItemData(cmbDonos.ListIndex))
           CmbPets.SetFocus
       Else
           MsgBox "Favor escolher um dono ou então cadastrar um novo", vbOKOnly, "Aviso"
           cmbDonos.SetFocus
           Exit Sub
       End If
   ElseIf KeyCode = vbKeyEscape Then
        If MsgBox("Deseja mesmo sair e descartar as alterações? ", vbYesNo + vbQuestion, "Responda-me") = vbYes Then
            cmd_Voltar_Click
        End If
   
   End If
    
End Sub

Private Sub cmbDonos_KeyPress(KeyAscii As Integer)

    Dim cb As Long
    Dim FindString As String
    Const CB_ERR = (-1)
    Const CB_FINDSTRING = &H14C
    With cmbDonos
        If KeyAscii < 32 Or KeyAscii > 127 Then
            Exit Sub
        End If
        
        If .SelLength = 0 Then
            FindString = .text & Chr$(KeyAscii)
        Else
            FindString = Left$(.text, .SelStart) & Chr$(KeyAscii)
        End If
        
        cb = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal FindString)
        
        If cb <> CB_ERR Then
            .ListIndex = cb
            .SelStart = Len(FindString)
            .SelLength = Len(.text) - .SelStart
        End If
        KeyAscii = 0
    End With
End Sub


Private Sub CmbHorario_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
        If CmbHorario.ListIndex <> -1 Then
            'fCarrega_Pets (cmbDonos.ItemData(cmbDonos.ListIndex))
            cmbServicos.SetFocus
        End If
   ElseIf KeyCode = vbKeyEscape Then
        If MsgBox("Deseja mesmo sair e descartar as alterações? ", vbYesNo + vbQuestion, "Responda-me") = vbYes Then
            cmd_Voltar_Click
        End If
    End If
End Sub

Private Sub CmbHorario_Validate(Cancel As Boolean)
    If CmbHorario.ListIndex = -1 Then
        MsgBox "Favor escolher um horario para o atendimento", vbOKOnly, "Aviso"
        Cancel = True
    End If
End Sub

Private Sub CmbPets_Click()
    If CmbPets.ListIndex <> -1 Then
        'fCarrega_Pets (cmbDonos.ItemData(cmbDonos.ListIndex))
        Call Busca_tipo(CmbPets.ItemData(CmbPets.ListIndex))
    End If
End Sub

Private Sub CmbPets_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
        If CmbPets.ListIndex <> -1 Then
            'fCarrega_Pets (cmbDonos.ItemData(cmbDonos.ListIndex))
            Call Busca_tipo(CmbPets.ItemData(CmbPets.ListIndex))
            If CmbHorario.Enabled = True Then
                CmbHorario.SetFocus
            Else
                cmbServicos.SetFocus
            End If
        End If
   ElseIf KeyCode = vbKeyEscape Then
        If MsgBox("Deseja mesmo sair e descartar as alterações? ", vbYesNo + vbQuestion, "Responda-me") = vbYes Then
            cmd_Voltar_Click
        End If
    End If
End Sub

Private Sub cmbServicos_Change()
     txtValor.text = Format(fCarrega_Servicos(cmbServicos.ItemData(cmbServicos.ListIndex)), "###,##0.00")
End Sub

Private Sub cmbServicos_Click()
     txtValor.text = Format(fCarrega_Servicos(cmbServicos.ItemData(cmbServicos.ListIndex)), "###,##0.00")
End Sub

Private Sub cmbServicos_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
        If cmbServicos.ListIndex <> -1 Then
            'fCarrega_Pets (cmbDonos.ItemData(cmbDonos.ListIndex))
            
            txtValor.SetFocus
        End If
   ElseIf KeyCode = vbKeyEscape Then
        If MsgBox("Deseja mesmo sair e descartar as alterações? ", vbYesNo + vbQuestion, "Responda-me") = vbYes Then
            cmd_Voltar_Click
        End If
   End If
End Sub

Private Sub cmd_Adicionar_Click()
    'DTPicker1.Enabled = True
    If DateDiff("D", DTPicker1.Value, Date) > 10 Then
        If MsgBox("Data do agendamento está muito diferente da data atual." & vbCrLf & "Deseja prosseguir assim mesmo?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
    End If
    
    If Not fCarrega_Servicos() Then
        MsgBox "Sem serviços cadastrados, favor cadastrar", vbOKOnly, "Aviso"
        Exit Sub
    End If
    If Not fCarrega_Tipos() Then
        MsgBox "Sem Tipos de animais cadastrados, favor cadastrar", vbOKOnly, "Aviso"
        Exit Sub
    End If
   
    If Not fCarregadonos() Then
        MsgBox "Sem donos de pet cadastrados, favor cadastrar", vbOKOnly, "Aviso"
        Exit Sub
    End If
    cmbDonos.Visible = True
    cmbDonos.Top = txtProp.Top
    cmbDonos.Left = txtProp.Left
   ' cmbDonos.Height = txtProp.Height
    cmbDonos.Width = txtProp.Width
    cmbDonos.ListIndex = 0
    txtProp.Visible = False
    
    'framLista.Visible = False
    framLista.Enabled = False
    frameDetalhe.Enabled = True
    frameDetalhe.Visible = True
    frameDetalhe.Left = 4260
    
    CmbHorario.Enabled = True
    
    txtTipoAtend.Visible = False
    cmbServicos.Enabled = True
    cmbServicos.Visible = True
    cmbServicos.ListIndex = 0
    TxtObserv.Enabled = True
    txtValor.Enabled = True
    
    TxtObserv.text = ""
    txtValor.text = "0.00"
    
    cmd_pets.Enabled = False
    cmd_servicos.Enabled = False
    cmd_tipos.Enabled = False
    cmd_ajustes.Enabled = False
    cmd_Relatos.Enabled = False
    
    'If Not fCarrega_Pets() Then
    'End If
    'If Not fCarrega_Servicos() Then
    'End If
    
    Call Carrega_Combo_Horario_Livre
    
    Posiciona_Combo_Horario (LIst_Atendimentos.SelectedItem)
    
    cmbServicos.ListIndex = 0
    Call cmbServicos_Click
    
    cmd_Adicionar.Enabled = False
    cmd_Adicionar.Visible = False
    cmd_Gravar.Enabled = True
    cmd_Voltar.Enabled = True
    cmd_Voltar.Visible = True
    
    cmbDonos.Enabled = True
    'cmbDonos.ListIndex = 0
        
    cmbDonos.SetFocus
    
    DoEvents
    
    iTipoOperacao = 1

End Sub

Private Sub cmd_Ajustes_Click()
    Dim Form As Form
    Me.Visible = False
    Set Form = New frmAjustes
    frmAjustes.Show vbModal
    Me.Visible = True
    Set Form = Nothing
    Call sCarrega_Agenda
End Sub

Private Sub cmd_Clientes_Click()
    Consulta = "S"
    frmCliente.IncluirCliente = 2
    frmCliente.MskCliDesde.text = Format(Date, "dd/mm/yyyy")
    frmCliente.Show 1

End Sub

Private Sub Cmd_formas_Click()
    Dim Form As Form
    Me.Visible = False
    Set Form = New frmFormas
    Form.Show vbModal
    cmd_servicos.Enabled = True
    cmd_pets.Enabled = True
    Me.Visible = True
    Set Form = Nothing
    'Call Form_Load
    Call sCarrega_Agenda

End Sub

Private Sub cmd_Gravar_Click()
    If ioperacao = 1 Then
        If CmbPets.ListIndex = -1 Then
            MsgBox "Pet inválido. Favor corrigir", vbOKOnly, "Aviso"
            CmbPets.SetFocus
            Exit Sub
        End If
        If cmbDonos.ListIndex = -1 Then
            MsgBox "Escolha um cliente por favor", vbOKOnly, "Aviso"
            cmbDonos.SetFocus
            Exit Sub
        End If
        If cmbServicos.ListIndex = -1 Then
            MsgBox "Favor escolher um tipo de atendimento", vbOKOnly, "Aviso"
            Cancel = True
        End If
    End If

    If fGravar_Atendimento() Then
        cmd_Adicionar.Enabled = True
        cmd_Adicionar.Visible = True
        cmd_Gravar.Enabled = False
        cmd_Limpar.Enabled = False
        cmd_Voltar.Enabled = False
        cmd_Voltar.Visible = False
        cmd_pets.Enabled = True
        cmd_servicos.Enabled = True
        cmd_tipos.Enabled = True
        cmd_ajustes.Enabled = True
        Call Desabilita(Me)
        CmbPets.Enabled = True
        'Call Form_Load
        Call Carrega_Combo_Horario
        Call sCarrega_Agenda
        Call Carrega_Combo_Horario_Livre
        iTipoOperacao = 0
    Else
        Exit Sub
    End If
    DTPicker1.Enabled = True
    framLista.Enabled = True
    frameDetalhe.Visible = False
    LIst_Atendimentos.Enabled = True
    LIst_Atendimentos.SetFocus
End Sub

Private Function fGravar_Atendimento()
    
    fGravar_Atendimento = False
    
    'Operacao:
    ' 1 - Inclusao
    ' 2 - Alteração
    ' 3 - Exclusão
    ' 4 - Baixa
    ' 5 - Recebimento
    ' 6 - Desfazer a baixa
    
    If iTipoOperacao < 3 Then      'Operacao 1- Inclusao 2 - Alteração 3 - Exclusão
        
        If cmbDonos.ListIndex = -1 Then
           MsgBox "Favor escolher um dono de pet", vbOKOnly, "Aviso"
           cmbDonos.SetFocus
           Exit Function
        End If
        
        If CmbPets.ListIndex = -1 Then
           MsgBox "Favor escolher pet", vbOKOnly, "Aviso"
           CmbPets.SetFocus
           Exit Function
        End If
        
        If CmbHorario.ListIndex = -1 Then
           MsgBox "Favor escolher um horário", vbOKOnly, "Aviso"
           CmbHorario.SetFocus
           Exit Function
        End If
        
        If cmbServicos.ListIndex = -1 Then
           MsgBox "Favor escolher um serviço a ser executado", vbOKOnly, "Aviso"
           CmbHorario.SetFocus
           Exit Function
        End If
        
        If Len(txtValor.text) = 0 Or _
           txtValor.text = "" Or _
           txtValor.text = "0.00" Or _
           Val(txtValor.text) = 0 Then
            If MsgBox("Valor está em branco ou zerado, deseja prosseguir assim mesmo? ", _
                    vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me") = vbNo Then
                txtValor.SetFocus
                Exit Function
            End If
        End If
    End If
    fGravar_Atendimento = True
    
'   strSql = "Insert INTO tab_atendimentos (dt_atend,idanimal,tipo_atend,valor,tempo_atend,operador, dt_Atualiza) "
'            " values (Now , idanimal,tipo_atend,valor,tempo_atend,operador, date) "

    Call sConectaBanco
    
    On Error GoTo Erro_fGravar_Atendimento
    If iTipoOperacao = 1 Then
        strSql = "INSERT INTO tab_atendimentos (DT_ATEND " & _
                 ",IDANIMAL " & _
                 ",TIPO_ATEND " & _
                 ",VALOR " & _
                 ",HORA_SAIDA " & _
                 ",OBSERVA " & _
                 ",OPERADOR " & _
                 ",DT_ATUALIZA)"
        strSql = strSql + " VALUES( '" & Format(DTPicker1.Value & " " & CmbHorario.text, "yyyy/mm/dd hh:mm:ss") & _
                        "'," & CmbPets.ItemData(CmbPets.ListIndex) & _
                        "," & cmbServicos.ItemData(cmbServicos.ListIndex) & _
                        "," & Val(txtValor.text) & _
                        ",'" & "  :  " & _
                        "','" & TxtObserv.text & _
                        "','" & NomeUsuario & _
                        "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"

        sOrdem = ReadIniFile(App.Path & "\Petshop.ini", "ORDEM", "", "")
        
        If sOrdem = "S" Then
            If MsgBox("Atendimento incluído com sucesso. deseja imprimi-lo ?", _
                    vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me") = vbYes Then
                Call sImprimeAtendimento
            End If
        Else
            'MsgBox "Atendimento incluído com sucesso.", vbOKOnly, "Aviso"
        End If

    ElseIf iTipoOperacao = 2 Or iTipoOperacao = 3 Then   'Caso seja alteração vai excluir o registro e posteriormente grava um novo
        'Primeiro vai apagar o registro do atendimento pois pode ser que trocou de horario
        strSql = "DELETE FROM tab_atendimentos WHERE DT_ATEND = '" & Format(DTPicker1.Value & " " & LIst_Atendimentos.SelectedItem, "yyyy/mm/dd hh:mm:ss") & _
                                          "' AND IDANIMAL = " & LIst_Atendimentos.SelectedItem.SubItems(IND_COD_PET)

        Cnn.Execute strSql
        
        If iTipoOperacao = 2 Then
            strSql = "INSERT INTO tab_atendimentos (DT_ATEND " & _
                     ",IDANIMAL " & _
                     ",TIPO_ATEND " & _
                     ",VALOR " & _
                     ",HORA_SAIDA " & _
                     ",OBSERVA " & _
                     ",OPERADOR " & _
                     ",DT_ATUALIZA)"
            strSql = strSql + " VALUES( '" & Format(DTPicker1.Value & " " & CmbHorario.text, "yyyy/mm/dd hh:mm:ss") & _
                            "'," & LIst_Atendimentos.SelectedItem.SubItems(IND_COD_PET) & _
                            "," & cmbServicos.ItemData(cmbServicos.ListIndex) & _
                            "," & Val(txtValor.text) & _
                            ",'" & "  :  " & _
                            "','" & TxtObserv.text & _
                            "','" & NomeUsuario & _
                            "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        End If
    ElseIf iTipoOperacao = 4 Then
        If LIst_Atendimentos.SelectedItem.SubItems(1) = "VACINA" Then
            strSql = "UPDATE tab_atendimentos SET HORA_SAIDA = '" & txtHrSaida & ":" & txtMinSaida & _
                                              "', HORA_VACINA = '" & Left(Time, 5) & _
                                              "',OPERADOR = '" & NomeUsuario & _
                                              "', DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                              "' WHERE DT_ATEND = '" & Format(DTPicker1.Value, "yyyy/mm/dd") & " " & LIst_Atendimentos.SelectedItem & ":00'"
        Else
            strSql = "UPDATE tab_atendimentos SET HORA_SAIDA = '" & txtHrSaida & ":" & txtMinSaida & _
                                              "', HORA_VACINA = '  :  " & _
                                              "',OPERADOR = '" & NomeUsuario & _
                                              "', DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                              "' WHERE DT_ATEND = '" & Format(DTPicker1.Value, "yyyy/mm/dd") & " " & LIst_Atendimentos.SelectedItem & ":00'"
        End If
        Cnn.Execute strSql
        Cnn.CommitTrans
        'Vai gravar o registro de vacinação do animal atendido
        If LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO) = "VACINA" Then
            strSql = "INSERT INTO tab_vacinas (IDANIMAL " & _
                     ",DT_ATEND " & _
                     ",DESCRICAO " & _
                     ",VALOR " & _
                     ",DT_PROXIMA " & _
                     ",OPERADOR " & _
                     ",DT_ATUALIZA)"
            strSql = strSql + " VALUES(" & LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO) & _
                              ",'" & Format(DTPicker1.Value & " " & CmbHorario.text, "yyyy/mm/dd hh:mm:ss") & _
                              "','" & txtDescVacina.text & _
                              "'," & Val(txtValor.text) & _
                              ",'" & Format(txtProximaVac.text, "yyyy/mm/dd") & _
                              "','" & NomeUsuario & _
                            "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        
        End If
    
    ElseIf iTipoOperacao = 5 Then
        strSql = "UPDATE tab_atendimentos SET VALOR_RECEBIDO = " & nValorRecebido & _
                                          ",OPERADOR = '" & NomeUsuario & _
                                          "', DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                          "', FORMAPAGTO = '" & sFormaPagto & _
                                          "' WHERE DT_ATEND = '" & Format(DTPicker1.Value, "yyyy/mm/dd") & " " & LIst_Atendimentos.SelectedItem & ":00'"
    
        'Aqui vai chamar a rotina de imprimir o recibo
        'Call sImprimeRecibo
        
    ElseIf iTipoOperacao = 6 Then
        strSql = "UPDATE tab_atendimentos SET HORA_SAIDA = '  :  '" & _
                                          ",OPERADOR = '" & NomeUsuario & _
                                          "', DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                          "' WHERE DT_ATEND = '" & Format(DTPicker1.Value, "yyyy/mm/dd") & " " & LIst_Atendimentos.SelectedItem & ":00'"
    Else
        strSql = "UPDATE tab_atendimentos SET HORA_VACINA = '  :  '" & _
                                          ",OPERADOR = '" & NomeUsuario & _
                                          "', DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                          "' WHERE DT_ATEND = '" & Format(DTPicker1.Value, "yyyy/mm/dd") & " " & LIst_Atendimentos.SelectedItem & ":00'"
        Cnn.Execute strSql
        Cnn.CommitTrans
                                          
        strSql = "DELETE FROM tab_vacinas WHERE IDANIMAL  = " & LIst_Atendimentos.SelectedItem.SubItems(IND_COD_PET) & _
                 " AND DT_ATEND >= '" & Format(DTPicker1.Value, "yyyy/mm/dd") & " 00:00:00'" & _
                 " AND DT_ATEND <= '" & Format(DTPicker1.Value, "yyyy/mm/dd") & " 23:59:59'"
                 
    End If
    Cnn.Execute strSql
    Cnn.CommitTrans
    'Cnn.Close
    Exit Function

Erro_fGravar_Atendimento:
    fGravar_Atendimento = False
End Function

Private Sub cmd_Limpar_Click()
    Call Limpa_Campos_Texto
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True

End Sub

Private Sub cmd_Medic_Click()
    MsgBox "Aguarde... em desenvolvimento", vbOKOnly, "Aviso"
End Sub


'Private Sub cmd_Novo_cli_Click()
'    Me.Visible = False
'    Dim formPets As Form
'    Set formPets = New frmCadPets
'    'formPets.cmd_Adicionar_Click
'    formPets.Show vbModal
'    Me.Visible = True
'    Call fCarregadonos
'    If cmbDonos.ListIndex > -1 Then
'        Call fCarrega_Pets(cmbDonos.ItemData(cmbDonos.ListIndex))
'    End If
'    cmbDonos.SetFocus
'    Set formPets = Nothing
'
'End Sub

Private Sub cmd_Pets_Click()
    Dim Form As Form
    Me.Visible = False
    Set Form = New frmCadPets
    Form.Show vbModal
    cmd_tipos.Enabled = True
    cmd_servicos.Enabled = True
    Me.Visible = True
    Set Form = Nothing
    'Call Form_Load
    Call sCarrega_Agenda
End Sub

Private Sub cmd_Receber_Click()
    If Val(txtRecebido.text) = 0 Then
        If MsgBox("Valor está em branco, confirma? ", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me") = vbNo Then
            Exit Sub
        End If
    End If
    If LIST_DETALHESPGTO.SelectedItem = 0 Then
        MsgBox "Favor escolher uma forma de pagamento", vbOKOnly, "Aviso"
        LIST_DETALHESPGTO.SetFocus
        Exit Sub
    End If
    Call cmd_Gravar_Click
End Sub

Private Sub cmd_Relatos_Click()
    'MsgBox "Aguarde... Em desenvolvimento", vbOKOnly, "Aviso"
    FrmRelatos.Show vbModal
    'Cmd_Servicos.Enabled = True
    'cmd_pets.Enabled = True
    'Call Form_Load
    Call sCarrega_Agenda
End Sub



Private Sub cmd_Sair_Click()
    Unload Me
End Sub

Private Sub Cmd_Servicos_Click()
    Dim Form As Form
    Me.Visible = False
    Set Form = New frmServicos
    Form.Show vbModal
    cmd_tipos.Enabled = True
    cmd_pets.Enabled = True
    Me.Visible = True
    Set Form = Nothing
    'Call Form_Load
    Call sCarrega_Agenda
End Sub

Private Sub cmd_Tipos_Click()
    Dim Form As Form
    Me.Visible = False
    Set Form = New frmTipos
    Form.Show vbModal
    cmd_servicos.Enabled = True
    cmd_pets.Enabled = True
    Me.Visible = True
    Set Form = Nothing
    'Call Form_Load
    Call sCarrega_Agenda
End Sub


Private Sub cmd_Voltar_Click()
    'Call Form_Load
    Call Carrega_Combo_Horario
    Call sCarrega_Agenda
    Call Carrega_Combo_Horario_Livre
    cmbServicos.Enabled = False
    cmbServicos.Visible = False
    'txtAnimal.Visible = True
    txtTipoAtend.Visible = True
    DTPicker1.Enabled = True
    TxtObserv.Enabled = True
    
    cmd_pets.Enabled = False
    cmd_servicos.Enabled = False
    cmd_tipos.Enabled = False
    
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = False
    cmd_Limpar.Enabled = False

    cmd_Voltar.Enabled = False
    cmd_Voltar.Visible = False
    cmd_Adicionar.Enabled = True
    cmd_Adicionar.Visible = True
    Call Desabilita(Me)
    CmbPets.Enabled = True
   ' framLista.Width = 13995
    LIst_Atendimentos.Width = 13757
   ' framLista.Enabled = True
   ' framLista.Visible = True
    frameDetalhe.Visible = False
    framLista.Enabled = True
    framLista.Visible = True
    LIst_Atendimentos.Enabled = True
    LIst_Atendimentos.Visible = True
    LIst_Atendimentos.SetFocus
    cmd_pets.Enabled = True
    cmd_servicos.Enabled = True
    cmd_tipos.Enabled = True
    cmd_ajustes.Enabled = True
  
End Sub

Private Sub Command1_Click()
    Call sImprimeRecibo
End Sub

Private Sub DTPicker1_Change()
   
   Call Carrega_Combo_Horario
   Call MontaAtendimentos(Format(DTPicker1.Value, "yyyy/mm/dd"))
   If LIst_Atendimentos.ListItems.Count > 0 Then
       List_Atendimentos_ItemClick LIst_Atendimentos.ListItems(1)
       Call Busca_Atendimento(Format(DTPicker1.Value, "yyyy/mm/dd") & " " & "12:00:00")
       Call Posiciona_Combo_Horario(Left(LIst_Atendimentos.ListItems(1), 5))
   End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then ' Se vc der um ENTER
        If LIst_Atendimentos.Enabled = True Then
           LIst_Atendimentos.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        If MsgBox("Deseja mesmo sair?", vbYesNo + vbQuestion, "Responda-me") = vbYes Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Activate()
    If DTPicker1.Enabled Then
        DTPicker1.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        If MsgBox("Deseja mesmo sair?", vbYesNo + vbQuestion, "Responda-me") = vbYes Then
            Unload Me
        End If
    End If
        
End Sub

Private Sub Form_Load()
    Dim sArgs() As String
   
    Dim iLoop As Integer
    'Assuming that the arguments passed from
    'command line will have space in between,
    'you can also use comma or other things...
    'If Len(Command$) = 0 Then
    '    frmAcesso.Show vbModal
    'Else
    '    sArgs = Split(Command$, ",")
    '    sysNomeAcesso = sArgs(0)
    '    sysSenha = sArgs(1)
    'End If
           
    'Call sCria_tabelas   --- Retirado pois o banco vai junto
    
    '*** Fabio Reinert - 04/2017 - Nova Checagem de segurança - Inicio
    'Call sValidaCliente
    '*** Fabio Reinert - 04/2017 - Nova Checagem de segurança - Fim
    
    bClick_Listview = True
       
    'Aqui vai pegar o nome da empresa
    'If Len(Trim(NomeEmpresa)) = 0 Then
        gcEmpresa = "Empresa Testes"
        gcEndereco = "Rua Que sobe e desc, s/n"
    'End If
    
    'Titulo = "==============================================================================="
    'Empresa = UCase(NomeEmpresa)
    'Endereco = UCase(Endereco)
    'Fone = Fone1Empresa & " - " & Fone2Empresa
    'Moldura = "'"
    'linhaDupla = "==============================================================================="
    'LinhaSilmples = "-------------------------------------------------------------------------------"
       
    'Aqui vai posicionar os controles e os tamanhos deles
    framLista.Width = 13995
    LIst_Atendimentos.Width = 13757
    framLista.Visible = True
    frameDetalhe.Visible = False
   
    'Aqui vai carregar o combo com os horarios
    Call Carrega_Combo_Horario
   
    Call Carrega_Colunas_Atendimentos
      
    'Call MontaAtendimentos(Format(Date, "yyyy/mm/dd"))
   
    If Not fCarrega_Servicos() Then
        MsgBox "Sem serviços cadastrados, favor cadastrar", vbOKOnly, "Aviso"
    End If
    
    If Not fCarrega_Tipos() Then
        MsgBox "Sem Tipos de animais cadastrados, favor cadastrar", vbOKOnly, "Aviso"
    End If
   
    If Not fCarregadonos() Then
        MsgBox "Sem Donos de pet cadastrados, favor cadastrar", vbOKOnly, "Aviso"
    End If
    fCarrega_Pets (cmbDonos.ItemData(0))
    'Posiciona os combos nas coordenadas dos campos texto
   
    cmbServicos.Top = txtTipoAtend.Top
    cmbServicos.Left = txtTipoAtend.Left
    cmbServicos.Visible = False
   
    'cmd_Voltar.Top = cmd_Adicionar.Top
    'cmd_Voltar.Left = cmd_Adicionar.Left
    'cmd_Voltar.Width = cmd_Adicionar.Width
    'cmd_Voltar.Height = cmd_Adicionar.Height
    'cmd_Voltar.Enabled = False
    'cmd_Voltar.Visible = True
    cmd_Adicionar.Enabled = True
    cmd_Adicionar.Visible = True
    
    DTPicker1.Value = Now()
    Call sCarrega_Agenda
    
'    Call MontaAtendimentos(Format(Date, "yyyy/mm/dd"))
'
'    If List_Atendimentos.ListItems.Count > 0 Then
'        List_Atendimentos_ItemClick List_Atendimentos.ListItems(1)
'    End If
   
   
End Sub

Private Sub sCarrega_Agenda()
    
    Call Carrega_Combo_Horario
    Call MontaAtendimentos(Format(DTPicker1.Value, "yyyy/mm/dd"))
    Call Carrega_Combo_Horario_Livre
    
    If LIst_Atendimentos.ListItems.Count > 0 Then
        List_Atendimentos_ItemClick LIst_Atendimentos.ListItems(1)
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}"
    End If
    
    If KeyCode = vbKeyEscape Then
        If cmd_Gravar.Enabled = True Then
            mensagem = MsgBox("Salvar Dados ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
            If mensagem = vbYes Then
                Call cmd_Gravar_Click
            End If
            framLista.Enabled = True
            framLista.Visible = True
            frameDetalhe.Enabled = False
            frameDetalhe.Visible = False
            cmd_Gravar.Enabled = False
            cmd_Limpar.Enabled = False
            cmd_Adicionar.Visible = True
            cmd_Adicionar.Enabled = True
            cmd_Voltar.Visible = False
            cmd_pets.Enabled = True
            cmd_servicos.Enabled = True
            cmd_tipos.Enabled = True
            cmd_ajustes.Enabled = True
            DTPicker1.SetFocus
            Exit Sub
        Else
            Unload Me
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
    ElseIf KeyCode = vbKeyF4 And cmd_Limpar.Enabled = True Then
        Exit Sub
    ElseIf KeyCode = vbKeyF7 And cmd_Sair.Enabled = True Then
        cmd_Sair_Click
        Exit Sub
    ElseIf KeyCode = vbKeyEscape And cmd_Sair.Enabled = True Then
        mensagem = MsgBox("Informações não Salvas. Deseja Sair assim mesmo ?", _
                        vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
        If mensagem = vbNo Then
            Exit Sub
        Else
            Unload Me
        End If
    End If
     
End Sub

Private Function fCarrega_Pets(pCodpesq As Double)
   
  fCarrega_Pets = True
  Call sConectaBanco
  If Rstemp.State = adStateOpen Then
      Rstemp.Close
   End If

   strSql = "SELECT a.id,a.id_cli,a.nome as nomepet"
   strSql = strSql & " FROM tab_pets a"
   strSql = strSql & " where a.id_cli = " & pCodpesq & " ORDER BY a.nome"
   
   Rstemp.Open strSql, Cnn, adOpenKeyset
   If Rstemp.BOF And Rstemp.EOF Then
       fCarrega_Pets = False
       Exit Function
   End If
    With Rstemp
        CmbPets.Clear
        Do Until .EOF 'percorre o recordset ate o fim
          'inclui os itens correspondentes
          CmbPets.AddItem Rstemp!nomepet
          CmbPets.ItemData(CmbPets.NewIndex) = Rstemp!id
          .MoveNext
        Loop
    End With
    
    Rstemp.Close
'    Cnn.Close
End Function
'
'***** Carrega Donos
Private Function fCarregadonos()
  Dim Indice As Double
  fCarregadonos = True
  Call sConectaBanco
  If Rstemp.State = adStateOpen Then
      Rstemp.Close
   End If

   strSql = "SELECT a.id,a.razao_social as nome "
   strSql = strSql & " FROM  tab_clientes a"
   strSql = strSql & " ORDER BY a.razao_social "
   
   Rstemp.Open strSql, Cnn, adOpenKeyset
   If Rstemp.BOF And Rstemp.EOF Then
       fCarregadonos = False
       Exit Function
   End If
   
    With Rstemp
        cmbDonos.Clear
        Do Until .EOF 'percorre o recordset ate o fim
          'inclui os itens correspondentes
          cmbDonos.AddItem Rstemp!Nome
          cmbDonos.ItemData(cmbDonos.NewIndex) = Rstemp!id
          .MoveNext
        Loop
        
        Rstemp.Close
    End With

End Function


Private Sub List_Atendimentos_DblClick()
    If Len(LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO)) = 0 Then
        Call cmd_Adicionar_Click
    End If
End Sub

Private Sub List_Atendimentos_ItemClick(ByVal Item As MSComctlLib.ListItem)
     If Len(Trim(Item.text)) = 0 Then
          Call Busca_Atendimento(Format(DTPicker1.Value, "yyyy/mm/dd") & " 08:00")
          Call Posiciona_Combo_Horario("08:00")
      Else
          Call Busca_Atendimento(Format(DTPicker1.Value, "yyyy/mm/dd") & " " & Item.text)
          Call Posiciona_Combo_Horario(Item.text)
      End If
End Sub

Private Sub List_Atendimentos_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyUp And LIst_Atendimentos.SelectedItem.Index = 1 Then
        DTPicker1.SetFocus
        Exit Sub
    
    End If
    If KeyCode = vbKeyReturn Then
        'List_Atendimentos.Index
        If Len(Trim(LIst_Atendimentos.SelectedItem.SubItems(1))) = 0 Then
            KeyCode = 0
            KeyAscii = 0
            Call cmd_Adicionar_Click
            Exit Sub
        End If
        KeyCode = 0
        KeyAscii = 0
        Dim FrmOpcoes As Form
        Set FrmOpcoes = New frmOptions
        
        FrmOpcoes.OptAlterar.Enabled = False
        FrmOpcoes.OptBaixar.Enabled = False
        FrmOpcoes.OptDesfazer.Enabled = False
        FrmOpcoes.OptReceber.Enabled = False
        FrmOpcoes.OptVacina.Enabled = False
        FrmOpcoes.OptNada.Enabled = True
        FrmOpcoes.Top = LIst_Atendimentos.SelectedItem.Top + 5100
        FrmOpcoes.Left = LIst_Atendimentos.SelectedItem.Left + 2800
'
        Select Case UCase(LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO))
            Case "PENDENTE"
                FrmOpcoes.OptAlterar.Enabled = True
                FrmOpcoes.OptBaixar.Enabled = True
                FrmOpcoes.OptDesfazer.Enabled = False
                FrmOpcoes.OptReceber.Enabled = False
                If UCase(LIst_Atendimentos.SelectedItem.SubItems(IND_VACINA)) = "S" Then
                    FrmOpcoes.OptVacina.Enabled = True
                Else
                    FrmOpcoes.OptVacina.Enabled = False
                End If
            Case "ATENDIDO"
                If UCase(LIst_Atendimentos.SelectedItem.SubItems(IND_VACINA)) = "S" Then
                    FrmOpcoes.OptAlterar.Enabled = False
                    FrmOpcoes.OptBaixar.Enabled = False
                    FrmOpcoes.OptDesfazer.Enabled = True
                    FrmOpcoes.OptReceber.Enabled = True
                    FrmOpcoes.OptVacina.Enabled = False
                Else
                    FrmOpcoes.OptAlterar.Enabled = False
                    FrmOpcoes.OptBaixar.Enabled = False
                    FrmOpcoes.OptDesfazer.Enabled = True
                    FrmOpcoes.OptReceber.Enabled = True
                    FrmOpcoes.OptVacina.Enabled = False
                End If
            Case "VACINA"
                FrmOpcoes.OptAlterar.Enabled = False
''**** Testar a baixa de vacina   ***** Inicio
                FrmOpcoes.OptBaixar.Enabled = True
                FrmOpcoes.OptDesfazer.Enabled = False
                FrmOpcoes.OptReceber.Enabled = False
''**** Testar a baixa de vacina    *****  Fim
'
''                    sep2.Visible = False
''                    pmnuBaixar.Visible = False
''                    sep3.Visible = False
''                    pmnuDesfaz.Visible = True
''                    sep4.Visible = True
''                    pmnuReceber.Visible = True
            Case "RECEBIDO"
                 MsgBox "Atendimento já recebido, não pode ser alterado/excluido", vbOKOnly, "Aviso"
                 Exit Sub
        End Select
            
        FrmOpcoes.Show vbModal
        Select Case gOpcao
            Case 1
                'Alterar
                Call pMnuAlterar_Click
            Case 2
                'Baixar
                Call pmnuBaixar_Click
            Case 3
                'Desfazer
                Call pmnuDesfaz_Click
            Case 4
                'receber
                Call pmnuReceber_Click
            Case 5
                'Vacina
                Call pmnuVacina_Click
            Case 6
            'Nada
        End Select
        Call sCarrega_Agenda
        'List_Atendimentos.SetFocus
'            PopupMenu pMnuLista
    ElseIf KeyCode = vbKeyDelete Then
            If Len(Trim(LIst_Atendimentos.SelectedItem.SubItems(1))) > 0 Then
                Call pMnuExcluir_Click
            End If
    ElseIf KeyCode = vbKeyEscape Then
        If MsgBox("Deseja mesmo sair?", vbYesNo + vbQuestion, "Responda-me") = vbYes Then
            Unload Me
        End If
    End If
End Sub

Private Sub List_Atendimentos_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    Set lvListItems_Itens = LIst_Atendimentos.HitTest(X, y)

    'Check if a record was selected
    If lvListItems_Itens Is Nothing Then
        If LIst_Atendimentos.ListItems.Count > 0 Then
            LIst_Atendimentos.SelectedItem.Selected = False
        End If
        'se não estiver item selecionado desabilita menus
        If Button = 2 Then
            If UCase(LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO)) = "ATENDIDO" Then
                pmnuDesfaz.Visible = True
                sep2.Visible = True
            Else
                pmnuDesfaz.Visible = False
                sep2.Visible = False
            End If
            PopupMenu pMnuLista, , , , pMnuAlterar
        End If
        Exit Sub
    Else
        'Habilita menus
        lvListItems_Itens.Selected = True
        If Button = 2 Then
            If Len(LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO)) = 0 Then
                Exit Sub
            End If
            pMnuAlterar.Visible = True
            sep1.Visible = True
            pMnuExcluir.Visible = True
            sep2.Visible = True
            pmnuBaixar.Visible = True
            sep3.Visible = True
            pmnuDesfaz.Visible = True
            sep4.Visible = True
            pmnuReceber.Visible = True
            sep5.Visible = False
            pmnuVacina.Visible = False
            
            Select Case UCase(LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO))
                Case "PENDENTE"
                    pMnuAlterar.Visible = True
                    sep1.Visible = True
                    pMnuExcluir.Visible = True
                    sep2.Visible = True
                    pmnuBaixar.Visible = True
                    sep3.Visible = False
                    pmnuDesfaz.Visible = False
                    sep4.Visible = False
                    pmnuReceber.Visible = False
                    If UCase(LIst_Atendimentos.SelectedItem.SubItems(IND_VACINA)) = "S" Then
                        sep5.Visible = True
                        pmnuVacina.Visible = True
                    Else
                        sep5.Visible = False
                        pmnuVacina.Visible = False
                    End If
                Case "ATENDIDO"
                    If UCase(LIst_Atendimentos.SelectedItem.SubItems(IND_VACINA)) = "S" Then
                        pMnuAlterar.Visible = False
                        sep1.Visible = False
                        pMnuExcluir.Visible = False
                        sep2.Visible = False
                        pmnuBaixar.Visible = False
                        sep3.Visible = False
                        pmnuDesfaz.Visible = True
                        sep4.Visible = True
                        pmnuReceber.Visible = True
                        sep5.Visible = False
                        pmnuVacina.Visible = False
                    Else
                        pMnuAlterar.Visible = False
                        sep1.Visible = False
                        pMnuExcluir.Visible = False
                        sep2.Visible = False
                        pmnuBaixar.Visible = False
                        sep3.Visible = False
                        pmnuDesfaz.Visible = True
                        sep4.Visible = True
                        pmnuReceber.Visible = True
                        sep5.Visible = False
                        pmnuVacina.Visible = False
                    End If
                Case "VACINA"
                    pMnuAlterar.Visible = False
                    sep1.Visible = False
                    pMnuExcluir.Visible = False
'**** Testar a baixa de vacina   ***** Inicio
                    sep2.Visible = False
                    pmnuBaixar.Visible = True
                    sep3.Visible = True
                    pmnuDesfaz.Visible = True
                    sep4.Visible = False
                    pmnuReceber.Visible = False
'**** Testar a baixa de vacina    *****  Fim

'                    sep2.Visible = False
'                    pmnuBaixar.Visible = False
'                    sep3.Visible = False
'                    pmnuDesfaz.Visible = True
'                    sep4.Visible = True
'                    pmnuReceber.Visible = True
                Case "RECEBIDO"
                     MsgBox "Atendimento já recebido, não pode ser alterado/excluido", vbOKOnly, "Aviso"
                     Exit Sub
            End Select
            
            PopupMenu pMnuLista, , 3000, 6000
        End If
    End If
End Sub

Private Function fCarrega_Servicos(Optional pPesq As Double)
   fCarrega_Servicos = True
   Call sConectaBanco
   strSql = "select ID,descricao, valor FROM tab_servicos "
   If pPesq > 0 Then
       strSql = strSql & " WHERE ID = " & pPesq
   End If
   strSql = strSql & " ORDER BY descricao"
   If RsTemp1.State = adStateOpen Then
       RsTemp1.Close
   End If
   RsTemp1.Open strSql, Cnn, adOpenKeyset
   If RsTemp1.BOF And RsTemp1.EOF Then
      fCarrega_Servicos = False
      Exit Function
   End If
   '**** Caso seja apenas pesquisa de preço não carregar o Combo novamente ***
    If pPesq = 0 Then
        Carrega_Combo_Servicos
    Else
        fCarrega_Servicos = RsTemp1!Valor
    End If
   
   RsTemp1.Close
  ' Cnn.Close
   
End Function

Private Sub Carrega_Combo_Servicos()
 With RsTemp1
      .MoveFirst
      cmbServicos.Clear
      Do While Not .EOF
        cmbServicos.AddItem Trim(RsTemp1!Descricao)
        'cmbServicos.List(cmbServicos.ListIndex, 2) = RsTemp1!valor
        cmbServicos.ItemData(cmbServicos.NewIndex) = RsTemp1!id
        .MoveNext
      Loop
  End With
     
End Sub


Private Function fCarrega_Tipos()
   fCarrega_Tipos = True
   Call sConectaBanco
   If Rstemp2.State = adStateOpen Then
      Rstemp2.Close
   End If
   strSql = "select ID,descricao FROM tab_tipos_pets"
   Rstemp2.Open strSql, Cnn, adOpenKeyset
   If Rstemp2.BOF And Rstemp2.EOF Then
      fCarrega_Tipos = False
      Exit Function
   End If
   'Carrega_Combo_Servicos
   Rstemp2.Close
'   Cnn.Close
   
End Function

Private Sub Carrega_Combo_Horario()
   Dim I As Integer
   Dim X As Integer
   Dim sHora, sHoraInicio, sHoraFim, sDuracao As String
   Dim nVagas As Integer
   Dim nDuracao As Double
   Dim bAchou As Boolean
   
   bAchou = True
   'Antes deve ser carregada a lista de atendimentos do dia para depois vermos qual horario está livre
    sHoraInicio = ReadIniFile(App.Path & "\PetShop.ini", "HORA_INICIO", "", "")
    If sHoraInicio = "" Then
        sHoraInicio = "08:00"
        bAchou = False
    End If
    sHoraFim = ReadIniFile(App.Path & "\PetShop.ini", "HORA_FIM", "", "")
    If sHoraFim = "" Then
        sHoraFim = "18:00"
        bAchou = False
    End If
    
    sDuracao = ReadIniFile(App.Path & "\PetShop.ini", "DURACAO", "", "")
    If sDuracao = "" Then
        sDuracao = "030"
        bAchou = False
    End If
    
    If Not bAchou Then
        Call sMostraAviso("Atenção operador, Parametros configurados ", _
                          "Hora Inicial: " & sHoraInicio & " Hora Final: " & sHoraFim, _
                          "Duração do atendimento: " & sDuracao, _
                          "Se quiser alterar clique no botão AJUSTES da tela de agendamento")
        WriteIniFile App.Path & "\Petshop.ini", "HORA_INICIO", "", sHoraInicio
        WriteIniFile App.Path & "\Petshop.ini", "HORA_FIM", "", sHoraFim
        WriteIniFile App.Path & "\Petshop.ini", "DURACAO", "", sDuracao
   End If
   
   nDuracao = Val(sDuracao)    'Duração em minutos do serviço / Atendimento
   sHora = sHoraInicio
   CmbHorario.Clear
   I = 0
   Do While Left(sHora, 5) <= Left(sHoraFim, 5)
       aLista_Horario(I) = Mid(sHora, 1, 5)
       CmbHorario.AddItem (aLista_Horario(I))
       sHora = DateAdd("n", nDuracao, CDate(sHora))
       I = I + 1
   Loop
   
End Sub

Private Sub Carrega_Combo_Horario_Livre()
   Dim I As Integer
   Dim X As Integer
   Dim sHora, sHoraInicio, sHoraFim, sDuracao As String
   Dim nVagas As Integer
   Dim nDuracao As Double
   Dim bAchou As Boolean
   
   CmbHorario.Clear
   For X = 1 To LIst_Atendimentos.ListItems.Count
       If LIst_Atendimentos.ListItems(X).SubItems(1) = "" Then
           CmbHorario.AddItem (LIst_Atendimentos.ListItems(X))
       End If
       'sHora = DateAdd("n", nDuracao, CDate(sHora))
   Next

End Sub

Private Sub LIST_DETALHESPGTO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmd_Receber.SetFocus
    ElseIf KeyCode = vbKeyUp And LIST_DETALHESPGTO.SelectedItem.Index = 1 Then
        txtRecebido.SetFocus
    End If
End Sub

Private Sub LIST_DETALHESPGTO_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmd_Receber.SetFocus
    ElseIf KeyAscii = vbKeyUp And LIST_DETALHESPGTO.SelectedItem.Index = 1 Then
        txtRecebido.SetFocus
    End If

End Sub

Private Sub LIST_DETALHESPGTO_LostFocus()
    cmd_Receber.SetFocus
End Sub

Private Sub pMnuAlterar_Click()
    If Len(LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO)) = 0 Then
       Exit Sub
    End If
    If LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO) <> "PENDENTE" Then
        MsgBox "Atendimento não pode ser alterado", vbOKOnly, "Aviso"
    Else
        cmbDonos.Visible = True
        cmbDonos.Top = txtProp.Top
        cmbDonos.Left = txtProp.Left
        ' cmbDonos.Height = txtProp.Height
        cmbDonos.Width = txtProp.Width
        Call Posiciona_Combo_Codigo(cmbDonos, LIst_Atendimentos.SelectedItem.SubItems(IND_COD_DONO))
        txtProp.Visible = False
        'framLista.Visible = False
        framLista.Enabled = False
        frameDetalhe.Enabled = True
        frameDetalhe.Visible = True
        If Not fCarrega_Pets(LIst_Atendimentos.SelectedItem.SubItems(IND_COD_DONO)) Then
            MsgBox "Não existem PETS cadastrados para esse nome"
            Unload Me
        End If
        cmbDonos.Enabled = True
        Call Posiciona_Combo_Codigo(CmbPets, LIst_Atendimentos.SelectedItem.SubItems(IND_COD_PET))
        CmbPets.Enabled = True
        'CmbHorario.ListIndex = 1
        Call Posiciona_Combo(CmbHorario, LIst_Atendimentos.SelectedItem.text)
        Call Posiciona_Combo_Codigo(cmbServicos, LIst_Atendimentos.SelectedItem.SubItems(IND_TIPO_ATEND))
        'Call Carrega_Combo_Horario_Livre
        'CmbHorario.Enabled = True
        'CmbHorario.SetFocus
        'Call cmbDonos_LostFocus
        cmbServicos.Enabled = True
        cmbServicos.Visible = True
        txtTipoAtend.Visible = False
        TxtObserv.Enabled = True
        txtValor.Enabled = True
        cmd_Voltar.Enabled = True
        cmd_Voltar.Visible = True
        cmd_Adicionar.Visible = False
        cmd_Adicionar.Enabled = False
        cmd_Gravar.Enabled = True
        cmd_Limpar.Enabled = True
        DTPicker1.Enabled = False
        cmbDonos.SetFocus
        iTipoOperacao = 2
    End If
    
End Sub

Private Sub pMnuExcluir_Click()
    If Len(LIst_Atendimentos.SelectedItem.SubItems(1)) = 0 Then
       Exit Sub
    End If
    If LIst_Atendimentos.SelectedItem.SubItems(1) <> "PENDENTE" Then
        MsgBox "Atendimento Não pode ser excluido. Status =  " & _
               LIst_Atendimentos.SelectedItem.SubItems(1), vbOKOnly, "Aviso"
    Else
        If MsgBox("Tem certeza que deseja excluir o atendimento para o PET:  " & Chr(13) & Chr(10) & _
              Trim(LIst_Atendimentos.SelectedItem.SubItems(2)) & " - " & _
              Trim(LIst_Atendimentos.SelectedItem.SubItems(3)) & Chr(13) & Chr(10) & _
              " PROPRIETÁRIO = " & Trim(LIst_Atendimentos.SelectedItem.SubItems(4)), _
              vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me") = vbYes Then
            iTipoOperacao = 3
            Call Posiciona_Combo(cmbServicos, LIst_Atendimentos.SelectedItem.SubItems(5))
            Call fGravar_Atendimento
            'Call Form_Load
            Call Carrega_Combo_Horario
            Call sCarrega_Agenda
            Call Carrega_Combo_Horario_Livre
        End If
    End If

End Sub

Private Sub pmnuBaixar_Click()
    If Len(LIst_Atendimentos.SelectedItem.SubItems(1)) = 0 Then
       Exit Sub
    End If
    If LIst_Atendimentos.SelectedItem.SubItems(1) = "ATENDIDO" Then
        MsgBox "Atendimento já foi baixado", vbOKOnly, "Aviso"
    Else
        If LIst_Atendimentos.SelectedItem.SubItems(1) = "VACINA" Then
            bVacina = False
            Dim fVacina As Form
            Set fVacina = New frmVacina
            fVacina.txtDtVacina.text = Format(DTPicker1.Value, "DD/MM/YYYY")
            fVacina.txtProximaVac.text = Format(DateAdd("D", 30, DTPicker1.Value), "DD/MM/YYYY")
            fVacina.txtDescVacina.text = LIst_Atendimentos.SelectedItem.SubItems(5)
            fVacina.Show vbModal
            Unload fVacina
            Set fVacina = Nothing
            If bVacina Then
                iTipoOperacao = 4
                txtHrSaida = Left(Time, 2)
                txtMinSaida = Mid(Time, 4, 2)
                Call cmd_Gravar_Click
            Else
                MsgBox "Próxima vacina não foi registrada", vbOKOnly + vbInformation, "Aviso"
            End If
        Else
            iTipoOperacao = 4
            txtHrSaida = Left(Time, 2)
            txtMinSaida = Mid(Time, 4, 2)
            'txtHrSaida.SetFocus
            Call cmd_Gravar_Click
        End If
    End If
End Sub


Private Sub pmnuReceber_Click()
    
    If Len(LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO)) = 0 Then
       Exit Sub
    End If
    If LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO) = "ATENDIDO" Then
        'cmd_Gravar.Enabled = True
        'cmd_Limpar.Enabled = True
        'cmd_Adicionar.Enabled = False
        'cmd_Adicionar.Visible = False
        'cmd_Voltar.Enabled = True
        'cmd_Voltar.Visible = True
        bRecebido = False
        Dim fReceber As Form
        Set fReceber = New frmReceberPet
        fReceber.txtRecebido.text = LIst_Atendimentos.SelectedItem.SubItems(IND_VALOR)
        fReceber.Show vbModal
        Unload fReceber
        Set fReceber = Nothing
        If bRecebido Then
            iTipoOperacao = 5
            Call cmd_Gravar_Click
        Else
            MsgBox "Recebimento não foi efetuado", vbOKOnly + vbInformation, "Aviso"
        End If
    ElseIf LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO) = "VACINA" Then
        cmd_Gravar.Enabled = True
        cmd_Limpar.Enabled = True
        cmd_Adicionar.Enabled = False
        cmd_Adicionar.Visible = False
        cmd_Voltar.Enabled = True
        cmd_Voltar.Visible = True
        iTipoOperacao = 5
        Call cmd_Gravar_Click
    Else
        MsgBox "Atendimento não pode ter recebimento. Status =  " & _
               LIst_Atendimentos.SelectedItem.SubItems(1), vbOKOnly, "Aviso"
    End If
End Sub

Private Sub pmnuDesfaz_Click()
    If LIst_Atendimentos.SelectedItem.SubItems(IND_HORA_SAIDA) = "  :  " Then
        MsgBox "Nada a processar", vbOKOnly, "Aviso"
    Else
        txtHrSaida.Enabled = True
        txtHrSaida.text = ""
        txtMinSaida.Enabled = True
        txtMinSaida.text = ""
        cmd_Gravar.Enabled = True
        cmd_Limpar.Enabled = True
        cmd_Adicionar.Enabled = False
        cmd_Voltar.Enabled = True
        cmd_Voltar.Visible = True
        If LIst_Atendimentos.SelectedItem.SubItems(IND_VACINA) = "S" Then
           iTipoOperacao = 7
        Else
           iTipoOperacao = 6
        End If
        Call cmd_Gravar_Click
    End If
End Sub

Private Sub pmnuVacina_Click()
    If Len(LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO)) = 0 Then
       Exit Sub
    End If
    If LIst_Atendimentos.SelectedItem.SubItems(IND_SITUACAO) = "ATENDIDO" Then
        MsgBox "Atendimento já foi baixado", vbOKOnly, "Aviso"
    Else
        iTipoOperacao = 4
        'txtHrSaida = Left(Time, 2)
        'txtMinSaida = Mid(Time, 4, 2)
        Call cmd_Gravar_Click
    End If

End Sub

Private Sub txtAnimal_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Carrega_campos_texto()
    
    'txtAnimal.Text = Trim(Rstemp!nome)
    lbltipo.Caption = Rstemp!TIPOPET
    txtProp.text = Rstemp!RAZAO_SOCIAL
    Call Carrega_Endereco
    txtTipoAtend.text = Rstemp!SERVICO
    txtValor.text = Format(Rstemp!Valor, "##,##0.00")
'    txtRecebido.Text = Format(Rstemp!valor_recebido, "##,##0.00")
    txtHrSaida.text = Left(Rstemp!HORA_SAIDA, 2)
    txtMinSaida.text = Mid(Rstemp!HORA_SAIDA, 4, 2)
    TxtObserv.text = IIf(IsNull(Rstemp!observa), "", Rstemp!observa)
    If Len(Rstemp!cuidados_Especiais) > 0 Then
        lblEspecial1.Caption = Mid(Rstemp!cuidados_Especiais, 1, 40)
    Else
        lblEspecial1.Caption = ""
    End If
    
End Sub

Private Sub Limpa_Campos_Texto()
    
    lbltipo.Caption = ""
    txtProp.text = ""
    
    lblEndereco.Caption = ""
    lblFone1.Caption = ""
    lblFone2.Caption = ""
    txtTipoAtend.text = ""
    txtValor.text = "0,00"
   ' txtRecebido.Text = "0,00"
    txtHrSaida.text = ""
    txtMinSaida.text = ""
    TxtObserv.text = ""
    lblEspecial1.Caption = ""
    
End Sub


Private Sub Posiciona_Combo_Horario(pHorario As String)
    For I = 0 To CmbHorario.ListCount - 1
        CmbHorario.ListIndex = I
        If CmbHorario.text >= pHorario Then
            Exit For
        End If
    Next
    
End Sub

Private Sub txtDescVacina_Validate(Cancel As Boolean)
    If Len(Trim(txtDescVacina)) = 0 Then
        MsgBox "Descrição da vacina não pode estar em branco", vbOKOnly, "Aviso"
        Cancel = True
    End If
End Sub

Private Sub txtHrSaida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    If KeyAscii = 44 And InStr(txtHrSaida.text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If Len(Trim(txtHrSaida.text)) = 0 Then
            MsgBox "Obrigatório Informar Hora da Saida.", vbInformation, "Aviso"
            txtHrSaida.SetFocus
            Exit Sub
        End If
        Sendkeys "{tab}"
    End If

End Sub

Private Sub txtHrSaida_Validate(Cancel As Boolean)
    If Val(txtHrSaida.text) < 0 Or Val(txtHrSaida.text) > 23 Then
       MsgBox "Informe uma Hora Válida.", vbExclamation, "Aviso "
       Cancel = True       ' aqui esta o pulo do gato , o foco não se move
    End If
End Sub

Private Sub txtMinSaida_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtMinSaida.text)) = 0 Then
            MsgBox "Obrigatório Informar minutos da Saida.", vbInformation, "Aviso"
            txtMinSaida.SetFocus
            Exit Sub
        End If
        Sendkeys "{tab}"
    End If
    
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    If KeyAscii = 44 And InStr(txtHrSaida.text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMinSaida_LostFocus()
    cmd_Gravar.SetFocus
End Sub

Private Sub txtMinSaida_Validate(Cancel As Boolean)
    If Val(txtMinSaida.text) < 0 Or Val(txtMinSaida.text) > 59 Then
       MsgBox "Informe Minutos Válidos.", vbExclamation, "Aviso"
       Cancel = True       ' aqui esta o pulo do gato , o foco não se move
    End If
    
End Sub

Private Sub TxtObserv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmd_Gravar.SetFocus
   ElseIf KeyCode = vbKeyEscape Then
        If MsgBox("Deseja mesmo sair e descartar as alterações? ", vbYesNo + vbQuestion, "Responda-me") = vbYes Then
            cmd_Voltar_Click
        End If
    End If
End Sub

Private Sub txtObserv_LostFocus()
    If MsgBox("Salvar dados?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me") = vbYes Then
        'iTipoOperacao = 1
        Call cmd_Gravar_Click
    End If
End Sub

Private Sub txtProp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtProp.text)) = 0 Then
            MsgBox "Obrigatório Informar o dono do Pet.", vbInformation, "Aviso"
            txtProp.SetFocus
            Exit Sub
        Else
            Call sConectaBanco
            If Rstemp.State = adStateOpen Then
                Rstemp.Close
            End If

            strSql = "SELECT a.id,a.razao_social as nome "
            strSql = strSql & " FROM  tab_clientes a"
            strSql = strSql & " WHERE a.razao_social = " & Trim(txtProp.text)
   
            Rstemp.Open strSql, Cnn, adOpenKeyset
            If Rstemp.BOF And Rstemp.EOF Then
                MsgBox "Dono não cadastrado", vbInformation, "Aviso"
                Exit Sub
            End If
            With Rstemp
                Do Until .EOF 'percorre o recordset ate o fim
                    If cmbDonos.ItemData(cmbDonos.ListIndex) = Rstemp!id Then
                       Rstemp.MoveLast
                    End If
                Loop
                Rstemp.Close
            End With

        End If
        Sendkeys "{tab}"
    End If
    
End Sub

Private Sub txtProximaVac_LostFocus()
    cmd_Gravar_Vacina.SetFocus
End Sub

Private Sub txtRecebido_GotFocus()
    SelText (txtRecebido)
End Sub

Private Sub txtRecebido_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        LIST_DETALHESPGTO.SetFocus
    End If
    If KeyCode = vbKeyEscape Then
        cmd_Voltar_Click
    End If
End Sub

Private Sub txtRecebido_LostFocus()
  '  iTipoOperacao = 5
  '  Call cmd_Gravar_Click
  LIST_DETALHESPGTO.SetFocus
End Sub

Private Sub Posiciona_Combo_Codigo(pCombo As ComboBox, pCodigo As String)
For I = 0 To pCombo.ListCount - 1
    pCombo.ListIndex = I
    'If UCase(Trim(pCombo.Text)) = UCase(Trim(pTexto)) Then
    If Trim(pCombo.ItemData(pCombo.ListIndex)) = Trim(pCodigo) Then
        Exit For
    End If
Next
End Sub

Private Sub Posiciona_Combo(pCombo As ComboBox, pTexto As String)
For I = 0 To pCombo.ListCount - 1
    pCombo.ListIndex = I
    If UCase(Trim(pCombo.text)) = UCase(Trim(pTexto)) Then
        Exit For
    End If
Next
End Sub

Private Sub sImprimeRecibo()
    Call sCarregaNomeEmpresa
    Relatorios.Formulas(0) = "RAZAO_SOCIAL = '" & gNomeEmpresa & "'"
    Relatorios.Formulas(1) = "DONO = '" & LIst_Atendimentos.SelectedItem.SubItems(4) & "'"
    Relatorios.Formulas(2) = "NOME_PET = '" & LIst_Atendimentos.SelectedItem.SubItems(2) & "'"
    Relatorios.Formulas(3) = "SERVICO = '" & LIst_Atendimentos.SelectedItem.SubItems(5) & "'"
    Relatorios.Formulas(4) = "VALOR = '" & LIst_Atendimentos.SelectedItem.SubItems(10) & "'"
    Relatorios.ReportFileName = App.Path & "\relrecibo.rpt"
   ' Relatorios.WindowTitle = frmRelCad.Caption & " - " & OptCliente.Caption

    On Error GoTo SaiImp
    'SelecPrint.Action = 5
    'Relatorios.PrintReport
    Relatorios.Action = 1
    Screen.MousePointer = 1
    Exit Sub
    
SaiImp:
    If Err.Number = 32755 Then
        Err.Clear
        Screen.MousePointer = 1
        Exit Sub
    Else
        MsgBox "Erro : " & Err.Description, vbOKOnly, "Aviso"
    End If
    Screen.MousePointer = 1

End Sub

Private Sub sImprimeAtendimento()
    Dim PaginaInicial, Paginafinal, numerodecopias, I
'    Dim impressora As Printer
'    CommonDialog1.CancelError = True
    On Error GoTo TrataErro_Imprime

    'mostra a janela para impressora
'    CommonDialog1.ShowPrinter
'    'Captura os valores definidos pelo usuário na janela
'    PaginaInicial = CommonDialog1.FromPage
'    Paginafinal = CommonDialog1.ToPage
'    numerodecopias = CommonDialog1.Copies
   Screen.MousePointer = 11
   Printer.Orientation = vbPRORLandscape
   Printer.ScaleMode = 7
   PageNumber = 1
'   Call CabecPa(PageNumber, True)   ' Calls PrintHeader Routine
      
   'DADOS CADASTRAIS
            
   Printer.FontName = "courier new"
   Printer.FontSize = 6
   Printer.FontItalic = False
   Printer.FontBold = True
   Printer.CurrentX = cMarginLeft
   Printer.CurrentY = 2.5
   Printer.Print "Dados para Atendimento:"
   Printer.FontBold = False
   Printer.CurrentY = Printer.CurrentY + 0.25
   
'   'Nome do Pet
   Printer.FontBold = True
   Printer.Print Tab(0); "Pet......... : ";
   Printer.Print Tab(18); txtAnimal.text;
   Printer.Print
   
   'tipo de animal
   Printer.Print Tab(0); "Tipo........ : ";
   Printer.Print Tab(18); lbltipo.Caption;
   Printer.Print
   
   'Proprietário
   Printer.FontBold = True
   Printer.Print Tab(0); "Proprietário : ";
   Printer.Print Tab(18); txtProp.text;
   Printer.Print
   
   'Endereço
   Printer.Print Tab(0); "Endereço....: ";
   Printer.Print Tab(18); lblEndereco.text;
   Printer.Print
   
 '  'Bairro
 '  Printer.FontBold = True
 '  Printer.Print Tab(0); "               ";
 '  Printer.Print Tab(18); txtBairro.text;
      
   'Tipo de atendimento
   Printer.FontBold = True
   Printer.Print Tab(0); "Servico...... :";
   Printer.Print Tab(18); txtTipoAtend.text;
   Printer.Print
   
'   'data vencimento

'   Printer.FontBold = True
'   Printer.Print Tab(108); "Data Vecto: ";
'   Printer.FontBold = False
'   Printer.Print Tab(120); tblliberacao!dta_LibFinal;
'
'   'taxa de liberacao
'   Printer.FontBold = True
'   Printer.Print Tab(133); "Taxa de Liberação: ";
'   Printer.FontBold = False
'   Printer.Print Tab(153); Format(tblliberacao!vlr_LibCotacao, "#0.000000");
'
'   'Valor emprestimo
'   Printer.FontBold = True
'   Printer.Print Tab(164); "Valor Empréstimo:  (Real): ";
'   Printer.FontBold = False
'   Printer.Print Tab(191); Format(tblliberacao!vlr_libEmprestimo, "#0.000000");
'
'   'atividade
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.FontBold = True
'   Printer.Print Tab(0); "Atividade: ";
'   Printer.FontBold = False
'   Printer.Print Tab(13); Mid(tblliberacao!desativ, 1, 20);
'
'   'tipo de emprestimo
'   Printer.FontBold = True
'   Printer.Print Tab(50); "Tipo de Emprestimo: ";
'   Printer.FontBold = False
'   Printer.Print Tab(74); Mid(tblliberacao!emprestimo, 1, 20);
'
'   'cliente
'   Printer.FontBold = True
'   Printer.Print Tab(100); "Cliente: ";
'   Printer.FontBold = False
'   Printer.Print Tab(112); "Banco de Crédito Nacional";
'
'   'valor liberado
'   Printer.FontBold = True
'   Printer.Print Tab(164); "Valor Liberado   (moeda): ";
'   Printer.FontBold = False
'   Printer.Print Tab(204 - Len(Format(tblliberacao!vlr_LibLiberado, "#0.000000"))); _
'                               Format(tblliberacao!vlr_LibLiberado, "#0.000000");
'
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.Print Tab(0); "Av. Das Nações Unidas, 12901 - do 2º ao 12º, 14º e 15º andares - Torre Oeste - São Paulo -SP ";
'
'   'moeda
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.FontBold = True
'   Printer.Print Tab(0); "Moeda: ";
'   Printer.FontBold = False
'   Printer.Print Tab(10); tblliberacao!Moeda
'
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.Print String(210, "_")
'
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.FontBold = True
'   Printer.Print "Carência"
'   Printer.FontBold = True
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.Print
'   Printer.FontBold = False
'
'   'carencia
'   Do While Not tblJrsPas.EOF And tblJrsPas!dta_jrsfinal < tblliberacao!dta_libcarencia
'      'primeira carência
'      Printer.Print Tab(2); "(" + Format(tblJrsPas!nro_Jrsparcela, "00") + ")";
'      Printer.Print Tab(7); tblJrsPas!dta_jrsfinal;
'      tblJrsPas.MoveNext
'      'segunda carência
'      Printer.Print Tab(32); "(" + Format(tblJrsPas!nro_Jrsparcela, "00") + ")";
'      Printer.Print Tab(37); tblJrsPas!dta_jrsfinal;
'      tblJrsPas.MoveNext
'      'terceira carência
'      Printer.Print Tab(62); "(" + Format(tblJrsPas!nro_Jrsparcela, "00") + ")";
'      Printer.Print Tab(67); tblJrsPas!dta_jrsfinal;
'      tblJrsPas.MoveNext
'      'quarta carência
'      Printer.Print Tab(92); "(" + Format(tblJrsPas!nro_Jrsparcela, "00") + ")";
'      Printer.Print Tab(97); tblJrsPas!dta_jrsfinal
'      tblJrsPas.MoveNext
'   Loop
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.Print String(210, "_")
'   If Printer.CurrentY >= Printer.ScaleHeight - cMarginBottom Then
'      Call SaltaPagina1
'   End If
'
'   '---------------------------------principal-----------------------------
'
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.FontBold = True
'   Printer.Print "PARCELA DE PRINCIPAL"
'   Printer.CurrentY = Printer.CurrentY + 0.15
'   Printer.Print "Nr Parc    data Vecto     data Pagto     Valor da Parcela(M)    Data Liquidação   Taxa da Liquidação   Valor Liquidado"
'   Printer.FontBold = False
'   Printer.Print
'   Do While Not tblPrincipalPas.EOF()
'      Printer.Print Tab(2); Format(tblPrincipalPas!nro_PriParcela, "00");
'      Printer.Print Tab(12); tblPrincipalPas!dta_priFinal; Tab(27); tblPrincipalPas!dta_pripagto;
'      Printer.Print Tab(60 - Len(Format(tblPrincipalPas!vlr_priparcela, "##,######0.000000"))); Format(tblPrincipalPas!vlr_priparcela, "##,######0.000000");
'      Printer.Print Tab(68); IIf(IsNull(tblPrincipalPas!dta_priLiquidado), " ", tblPrincipalPas!dta_priLiquidado);
'      Printer.Print Tab(95 - Len(Format(tblPrincipalPas!txa_PriReajuste, "#0000000"))); IIf(IsNull(tblPrincipalPas!dta_priLiquidado), " ", Format(tblPrincipalPas!txa_PriReajuste, "#0.000000"));
'      Printer.Print Tab(119 - Len(Format(tblPrincipalPas!vlr_priLiquidado, "##,##0.00"))); IIf(IsNull(tblPrincipalPas!dta_priLiquidado), " ", Format(tblPrincipalPas!vlr_priLiquidado, "##,##0.00"))
'      tblPrincipalPas.MoveNext
'      If Printer.CurrentY >= Printer.ScaleHeight - cMarginBottom Then
'         Call SaltaPagina1
'         Printer.FontBold = True
'         Printer.Print "PARCELA DE PRINCIPAL"
'         Printer.CurrentY = Printer.CurrentY + 0.15
'         Printer.Print "Nr Parc    data Vecto     data Pagto     Valor da Parcela(M)    Data Liquidação   Taxa da Liquidação   Valor Liquidado"
'         Printer.FontBold = False
'      End If
'   Loop
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.Print String(210, "_")
'   If Printer.CurrentY >= Printer.ScaleHeight - cMarginBottom Then
'      Call SaltaPagina1
'   End If
'
'   '-------------------------------juros-------------------------------
'
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.FontBold = True
'   Printer.Print "PARCELA DE JUROS"
'   Printer.CurrentY = Printer.CurrentY + 0.15
'   Printer.Print "Nr Parc   Data Inicial   Data Vecto     Vlr Juros (M)     Taxa de Juros(a.a.)   Valor Liquidado     Taxa Liquidação"
'   Printer.CurrentY = Printer.CurrentY + 0.15
'   Printer.FontBold = False
'   tblJrsPas.MoveFirst
'   Do While Not tblJrsPas.EOF()
'      Printer.Print Tab(2); Format(tblJrsPas!nro_Jrsparcela, "00");
'      Printer.Print Tab(12); tblJrsPas!dta_JrsInicio;
'      Printer.Print Tab(26); tblJrsPas!dta_jrsfinal;
'      Printer.Print Tab(54 - Len(Format(tblJrsPas!vlr_JrsParcela, "##,######0.000000"))); Format(tblJrsPas!vlr_JrsParcela, "##,######0.000000");
'      Printer.Print Tab(75 - Len(Format(tblJrsPas!txa_JrsJuros, "###0.000000"))); Format(tblJrsPas!txa_JrsJuros, "###0.000000");
'      Printer.Print Tab(96 - Len(Format(tblJrsPas!vlr_jrsLiquidado, "##,##0.00"))); IIf(IsNull(tblJrsPas!dta_jrsLiquidado), " ", Format(tblJrsPas!vlr_jrsLiquidado, "##,##0.00"));
'      Printer.Print Tab(116 - Len(Format(tblJrsPas!txa_JrsReajuste, "#0.000000"))); IIf(IsNull(tblJrsPas!dta_jrsLiquidado), " ", Format(tblJrsPas!txa_JrsReajuste, "#0.000000"))
'      tblJrsPas.MoveNext
'      If Printer.CurrentY >= Printer.ScaleHeight - cMarginBottom Then
'         Call SaltaPagina1
'         Printer.FontBold = True
'         Printer.Print "PARCELA DE JUROS"
'         Printer.CurrentY = Printer.CurrentY + 0.15
'         Printer.Print "Nr Parc   Data Inicial   Data Vecto     Vlr Juros (M)     Taxa de Juros(a.a.)   Valor Liquidado     Taxa Liquidação"
'         Printer.CurrentY = Printer.CurrentY + 0.15
'         Printer.FontBold = False
'      End If
'   Loop
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.Print String(210, "_")
'   If Printer.CurrentY >= Printer.ScaleHeight - cMarginBottom Then
'      Call SaltaPagina1
'   End If
'
'   '-------------------------------IR-------------------------------
'
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.FontBold = True
'   Printer.Print "PARCELA DE IR"
'   Printer.CurrentY = Printer.CurrentY + 0.3
'   Printer.Print "Nr Parc   Data Inicial   Data Vecto     Vlr IR   (M)     Taxa de I.R. (a.a.)   Valor Liquidado     Taxa Liquidação"
'   Printer.CurrentY = Printer.CurrentY + 0.15
'   Printer.FontBold = False
'   Do While Not tblIrPas.EOF()
'      Printer.Print Tab(2); Format(tblIrPas!nro_IrParcela, "00");
'      Printer.Print Tab(12); tblIrPas!dta_irinicio;
'      Printer.Print Tab(26); tblIrPas!dta_IrFinal;
'      Printer.Print Tab(54 - Len(Format(tblIrPas!vlr_IrParcela, "##,######0.000000"))); Format(tblIrPas!vlr_IrParcela, "##,######0.000000");
'      Printer.Print Tab(75 - Len(Format(tblIrPas!txa_ir, "###0.00000000"))); Format(tblIrPas!txa_ir, "###0.00000000");
'      Printer.Print Tab(96 - Len(Format(tblIrPas!vlr_irliquidado, "##,##0.00"))); IIf(IsNull(tblIrPas!dta_IrLiquidado), " ", Format(tblIrPas!vlr_irliquidado, "##,##0.00"));
'      Printer.Print Tab(116 - Len(Format(tblIrPas!txa_IrReajuste, "#0.000000"))); IIf(IsNull(tblIrPas!dta_IrLiquidado), " ", Format(tblIrPas!txa_IrReajuste, "#0.000000"))
'      tblIrPas.MoveNext
'      If Printer.CurrentY >= Printer.ScaleHeight - cMarginBottom Then
'         Call SaltaPagina1
'         Printer.FontBold = True
'         Printer.Print "PARCELA DE IR"
'         Printer.CurrentY = Printer.CurrentY + 0.3
'         Printer.Print ; "Nr Parc   Data Inicial   Data Vecto     Vlr IR   (M)     Taxa de I.R. (a.a.)   Valor Liquidado     Taxa Liquidação"
'         Printer.CurrentY = Printer.CurrentY + 0.15
'         Printer.FontBold = False
'      End If
'   Loop
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.Print String(210, "_")
'   If Printer.CurrentY >= Printer.ScaleHeight - cMarginBottom Then
'      Call SaltaPagina1
'   End If
'
'   '----------------------------comissao--------------------------------------
'
'   Printer.CurrentY = Printer.CurrentY + 0.25
'   Printer.FontBold = True
'   Printer.Print "PARCELA DE COMISSÃO"
'   Printer.CurrentY = Printer.CurrentY + 0.3
'   Printer.Print "Nr Parc   Data Inicial   Data Vecto     Vlr Comissão BNDES(M)    Vlr Comissão BCN (M)    Tx Comissão BNDES (a.a.)   Tx Comissão BCN (a.a.)    Valor Liquidado    Taxa Liquidação"
'   Printer.CurrentY = Printer.CurrentY + 0.15
'   Printer.FontBold = False
'   Do While Not tblComissaoPas.EOF()
'      Printer.Print Format(tblComissaoPas!nro_CmsParcela, "00"); Tab(12); tblComissaoPas!dta_CmsInicio; Tab(26); tblComissaoPas!dta_CmsFinal;
'      Printer.Print Tab(62 - Len(Format(tblComissaoPas!vlr_CmsParcBndes, "##,######0.000000"))); Format(tblComissaoPas!vlr_CmsParcBndes, "##,######0.000000");
'      Printer.Print Tab(86 - Len(Format(tblComissaoPas!vlr_CmsParcBcn, "##,######0.000000"))); Format(tblComissaoPas!vlr_CmsParcBcn, "##,######0.000000");
'      Printer.Print Tab(113 - Len(Format(tblComissaoPas!txa_CmsComissao, "#0.000000"))); Format(tblComissaoPas!txa_CmsComissao, "#0.000000");
'      Printer.Print Tab(139 - Len(Format(tblComissaoPas!txa_CmsBcn, "#0.000000"))); Format(tblComissaoPas!txa_CmsBcn, "#0.000000");
'      Printer.Print Tab(158 - Len(Format(tblComissaoPas!vlr_cmsliquidado, "##,##0.00"))); IIf(IsNull(tblComissaoPas!dta_CmsLiquidado), " ", Format(tblComissaoPas!vlr_cmsliquidado, "##,##0.00"));
'      Printer.Print Tab(178 - Len(Format(tblComissaoPas!txa_CmsReajuste, "#0.000000"))); IIf(IsNull(tblComissaoPas!dta_CmsLiquidado), " ", Format(tblComissaoPas!txa_CmsReajuste, "#0.000000"))
'      tblComissaoPas.MoveNext
'      If Printer.CurrentY >= Printer.ScaleHeight - cMarginBottom Then
'         Call SaltaPagina1
'         Printer.FontBold = True
'         Printer.Print "PARCELA DE COMISSÃO"
'         Printer.CurrentY = Printer.CurrentY + 0.3
'         Printer.Print "Nr Parc   Data Inicial   Data Vecto     Vlr Comissão BNDES(M)    Vlr Comissão BCN (M)    Tx Comissão BNDES (a.a.)   Tx Comissão BCN (a.a.)    Valor Liquidado    Taxa Liquidação"
'         Printer.CurrentY = Printer.CurrentY + 0.15
'         Printer.FontBold = False
'      End If
'   Loop
   Printer.CurrentY = Printer.CurrentY + 0.25
   Printer.Print String(210, "_")
   Printer.EndDoc
   Screen.MousePointer = 0

TrataErro_Imprime:
    Exit Sub
End Sub

Private Sub txtValor_GotFocus()
    SelText txtValor
End Sub

Public Sub sCarregaNomeEmpresa()
    If Cnn.State = adStateOpen Then
        Cnn.Close
    End If
    Call sConectaBanco
    sql = "select * from empresa "
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Cnn, 1, 2
    
    If Rs.RecordCount = 0 Then Exit Sub
    Rs.MoveLast
    Rs.MoveFirst
    While Not Rs.EOF
        gNomeEmpresa = IIf(IsNull(Rs("RazaoSocial_Empresa")), "", Rs("RazaoSocial_Empresa"))
        gEnderecoEmpresa = IIf(IsNull(Rs("Endereco_Empresa")), "", Rs("Endereco_Empresa"))
        gCGC_EMPRESA = IIf(IsNull(Rs("Cgc_Cpf")), "", Rs("Cgc_Cpf"))
        gCEP_EMPRESA = IIf(IsNull(Rs("Cep_Empresa")), "", Rs("Cep_Empresa"))
        gFone1Empresa = IIf(IsNull(Rs("fone1_Empresa")), "", Rs("fone1_Empresa"))
        gFone2Empresa = IIf(IsNull(Rs("fone2_Empresa")), "", Rs("fone2_Empresa"))
        gemailEmpresa = IIf(IsNull(Rs("E_MAIL_EMPRESA")), "", Rs("E_MAIL_EMPRESA"))
        Rs.MoveNext
    Wend
    'Status.Panels(2).Text = Trim(UCase(NomeEmpresa))
    
    Rs.Close
    Set Rs = Nothing
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        cmbServicos.SetFocus
    End If
    If KeyCode = vbKeyReturn Then
        TxtObserv.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        If MsgBox("Deseja mesmo sair e descartar as alterações? ", vbYesNo + vbQuestion, "Responda-me") = vbYes Then
            cmd_Voltar_Click
        End If
    End If
End Sub

Private Sub txtValor_LostFocus()
    txtValor.text = Format(txtValor.text, "###,##0.00")
End Sub


Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub
