VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1230
   ClientLeft      =   2565
   ClientTop       =   1215
   ClientWidth     =   7785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraOpcoes 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   420
      Width           =   7515
      Begin VB.OptionButton OptVacina 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vacina"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4782
         TabIndex        =   13
         Top             =   330
         Width           =   975
      End
      Begin VB.OptionButton OptReceber 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Receber"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3624
         TabIndex        =   12
         Top             =   330
         Width           =   1035
      End
      Begin VB.OptionButton OptDesfazer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Desfazer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2376
         TabIndex        =   11
         Top             =   330
         Width           =   1125
      End
      Begin VB.OptionButton OptBaixar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Baixar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1218
         TabIndex        =   10
         Top             =   330
         Width           =   960
      End
      Begin VB.OptionButton OptAlterar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alterar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   9
         Top             =   330
         Width           =   945
      End
      Begin VB.CommandButton cmd_opt_ok 
         Caption         =   "Ok"
         Height          =   345
         Left            =   6810
         TabIndex        =   8
         Top             =   240
         Width           =   555
      End
      Begin VB.OptionButton OptNada 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Voltar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5880
         TabIndex        =   7
         Top             =   330
         Value           =   -1  'True
         Width           =   795
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   4
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O que deseja fazer com esse atendimento? "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   180
      TabIndex        =   14
      Top             =   90
      Width           =   7365
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdAlterar_Click()
    frmAgenda.gOpcao = "A"
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    frmAgenda.gOpcao = "C"
    Unload Me
End Sub

Private Sub cmdExcluir_Click()
    frmAgenda.gOpcao = "E"
    Unload Me
End Sub

Private Sub cmd_opt_ok_Click()
    If OptAlterar = True Then
        frmAgenda.gOpcao = 1
    ElseIf OptBaixar = True Then
        frmAgenda.gOpcao = 2
    ElseIf OptDesfazer = True Then
        frmAgenda.gOpcao = 3
    ElseIf OptReceber = True Then
        frmAgenda.gOpcao = 4
    ElseIf OptVacina = True Then
        frmAgenda.gOpcao = 5
    ElseIf OptNada = True Then
        frmAgenda.gOpcao = 6
    End If
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If OptAlterar = True Then
        frmAgenda.gOpcao = 1
    ElseIf OptBaixar = True Then
        frmAgenda.gOpcao = 2
    ElseIf OptDesfazer = True Then
        frmAgenda.gOpcao = 3
    ElseIf OptReceber = True Then
        frmAgenda.gOpcao = 4
    ElseIf OptVacina = True Then
        frmAgenda.gOpcao = 5
    ElseIf OptNada = True Then
        frmAgenda.gOpcao = 6
    End If
    cmd_opt_ok.SetFocus

End Sub

Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

