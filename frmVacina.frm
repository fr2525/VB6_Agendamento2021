VERSION 5.00
Begin VB.Form frmVacina 
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   ScaleHeight     =   2325
   ScaleWidth      =   9615
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraVacina 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vacina"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.TextBox txtDtVacina 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3150
         TabIndex        =   1
         Text            =   "99/99/9999"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox txtDescVacina 
         Height          =   405
         Left            =   3150
         TabIndex        =   2
         Top             =   960
         Width           =   6015
      End
      Begin VB.TextBox txtProximaVac 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3150
         TabIndex        =   3
         Text            =   "99/99/9999"
         Top             =   1530
         Width           =   1095
      End
      Begin VB.CommandButton cmd_Gravar_Vacina 
         Caption         =   "Gravar"
         Height          =   435
         Left            =   6480
         TabIndex        =   4
         Top             =   1530
         Width           =   1275
      End
      Begin VB.CommandButton cmd_Pular 
         Caption         =   "Cancelar"
         Height          =   435
         Left            =   7890
         TabIndex        =   5
         Top             =   1530
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data da vacinação :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   270
         TabIndex        =   8
         Top             =   390
         Width           =   2730
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1410
         TabIndex        =   7
         Top             =   975
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Próxima Vacina :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   600
         TabIndex        =   6
         Top             =   1560
         Width           =   2370
      End
   End
End
Attribute VB_Name = "frmVacina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Gravar_Vacina_Click()
    frmAgenda.bVacina = True
    frmAgenda.dProximaVacina = txtProximaVac.Text
    frmAgenda.sDescVacina = txtDescVacina.Text
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyEscape Then
        Msg = "Deseja Salvar Dados ...?"
        resposta = MsgBox(Msg, vbQuestion + vbYesNoCancel + vbDefaultButton1, "Inclusão de Dados")
        If resposta = vbNo Then  'nao
            frmAgenda.bVacina = False
        ElseIf resposta = vbYes Then 'sim
            Call cmd_Gravar_Vacina_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' Se vc der um ENTER
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If
End Sub

Private Sub txtDescVacina_GotFocus()
    SelText txtDescVacina
End Sub

Private Sub txtDescVacina_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub txtDtVacina_GotFocus()
    SelText txtDtVacina
End Sub

Private Sub txtProximaVac_GotFocus()
    SelText txtProximaVac
    txtProximaVac.Text = DateAdd("M", 1, txtDtVacina.Text)
End Sub
