VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetalhe 
   Caption         =   "Detalhe"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameOperacao 
      Caption         =   "Operação"
      Height          =   1335
      Left            =   1530
      TabIndex        =   29
      Top             =   8100
      Width           =   3945
      Begin VB.CommandButton cmd_Voltar 
         Caption         =   "&Voltar (Alt+V)"
         Height          =   915
         Left            =   1980
         Picture         =   "frmDetalhe.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Desfaz a digitação"
         Top             =   690
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmd_Gravar 
         Caption         =   "&Gravar (Alt+G)"
         Enabled         =   0   'False
         Height          =   915
         Left            =   2640
         Picture         =   "frmDetalhe.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Grava o atendimento"
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton Cmd_limpar 
         Caption         =   "&Limpar (Alt + L)"
         Enabled         =   0   'False
         Height          =   915
         Left            =   1350
         Picture         =   "frmDetalhe.frx":053C
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Limpa campos do atendimento"
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Adicionar 
         Caption         =   "&Novo (Alt+N)"
         Height          =   915
         Left            =   120
         Picture         =   "frmDetalhe.frx":097E
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Novo atendimento"
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Frame frameDetalhe 
      Caption         =   "Detalhe"
      Height          =   7665
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   6645
      Begin VB.Frame Frame6 
         Caption         =   "Proprietário"
         Height          =   2025
         Left            =   300
         TabIndex        =   13
         Top             =   1230
         Width           =   6135
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
            TabIndex        =   18
            Top             =   210
            Width           =   5880
         End
         Begin VB.TextBox txtEndereco 
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
            TabIndex        =   17
            Top             =   660
            Width           =   5880
         End
         Begin VB.TextBox txtBairro 
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
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   16
            Top             =   1110
            Width           =   4470
         End
         Begin VB.TextBox txtFone1 
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
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   15
            Top             =   1560
            Width           =   1890
         End
         Begin VB.TextBox txtFone2 
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
            Left            =   4110
            MaxLength       =   50
            TabIndex        =   14
            Top             =   1560
            Width           =   1890
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
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
            Left            =   180
            TabIndex        =   20
            Top             =   1590
            Width           =   1170
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Bairro :"
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
            Left            =   600
            TabIndex        =   19
            Top             =   1230
            Width           =   765
         End
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
         Left            =   5610
         MaxLength       =   2
         TabIndex        =   12
         Top             =   3300
         Visible         =   0   'False
         Width           =   390
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
         Left            =   5010
         MaxLength       =   2
         TabIndex        =   11
         Top             =   3300
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtTipo 
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
         Top             =   810
         Width           =   4860
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cuidados Especiais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   885
         Left            =   270
         TabIndex        =   8
         Top             =   4740
         Width           =   6135
         Begin VB.Label lblEspecial1 
            AutoSize        =   -1  'True
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
            Left            =   150
            TabIndex        =   9
            Top             =   360
            Width           =   60
         End
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
         Left            =   1290
         MaxLength       =   20
         TabIndex        =   7
         Top             =   4260
         Width           =   1410
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
         Height          =   1680
         Left            =   1500
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   5760
         Width           =   4740
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
         TabIndex        =   5
         Top             =   3780
         Width           =   5010
      End
      Begin VB.TextBox txtAnimal 
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
         TabIndex        =   4
         Top             =   300
         Width           =   4860
      End
      Begin VB.ComboBox CmbHorario 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmDetalhe.frx":0DC0
         Left            =   1290
         List            =   "frmDetalhe.frx":0DC2
         TabIndex        =   3
         Text            =   "00:00"
         Top             =   3300
         Width           =   1125
      End
      Begin VB.ComboBox cmbServicos 
         Height          =   315
         ItemData        =   "frmDetalhe.frx":0DC4
         Left            =   4290
         List            =   "frmDetalhe.frx":0DC6
         TabIndex        =   2
         Text            =   "Servicos"
         Top             =   5850
         Visible         =   0   'False
         Width           =   3165
      End
      Begin MSComctlLib.ListView ListaPets 
         Height          =   1095
         Left            =   420
         TabIndex        =   1
         ToolTipText     =   "Duplo Clique para escolher o PET"
         Top             =   6450
         Visible         =   0   'False
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   1931
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lbl_Valor 
         Alignment       =   2  'Center
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
         Left            =   300
         TabIndex        =   28
         Top             =   4290
         Width           =   1080
      End
      Begin VB.Label lbl_Obseerv 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Left            =   360
         TabIndex        =   27
         Top             =   5730
         Width           =   900
      End
      Begin VB.Label lbl_TipoAtend 
         AutoSize        =   -1  'True
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
         Left            =   225
         TabIndex        =   26
         Top             =   3810
         Width           =   930
      End
      Begin VB.Label lbl_Animal 
         AutoSize        =   -1  'True
         Caption         =   "Pet :"
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
         Left            =   660
         TabIndex        =   25
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Left            =   270
         TabIndex        =   24
         Top             =   3330
         Width           =   900
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo :"
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
         Left            =   570
         TabIndex        =   23
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hora Saída :"
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
         Left            =   3600
         TabIndex        =   22
         Top             =   3360
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Left            =   5370
         TabIndex        =   21
         Top             =   3300
         Visible         =   0   'False
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmDetalhe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
