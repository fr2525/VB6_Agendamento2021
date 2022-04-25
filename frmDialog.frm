VERSION 5.00
Begin VB.Form frmDialog 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reativação"
   ClientHeight    =   4965
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7665
   ForeColor       =   &H00000000&
   Icon            =   "frmDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCNPJ 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   435
      Left            =   1800
      MaxLength       =   18
      TabIndex        =   6
      Text            =   "11.111.111/0001-93"
      Top             =   2040
      Width           =   3795
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Digite a senha informada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1463
      TabIndex        =   4
      Top             =   2700
      Width           =   4455
      Begin VB.TextBox txtContraSenha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   4155
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Fechar"
      Height          =   615
      Left            =   3990
      TabIndex        =   5
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   615
      Left            =   2175
      TabIndex        =   3
      Top             =   3900
      Width           =   1215
   End
   Begin VB.Label Lbltitulo2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Entre em contato com a Reinert Informática e solicite uma senha para liberação informando o CNPJ abaixo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   180
      TabIndex        =   1
      Top             =   840
      Width           =   7275
   End
   Begin VB.Label LblTitulo1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Liberação do sistema necessária"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   7215
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nContSenhas As Integer

Option Explicit

Private Sub CancelButton_Click()
    Ok = False
    'sRestauraSegur
    'Call Fecha_Formularios
    'End
    Unload Me
End Sub


Private Sub Form_Activate()
nContSenhas = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' Se vc der um ENTER
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If
End Sub

Private Sub OKButton_Click()
Ok = False
    
    nContSenhas = nContSenhas + 1
    If nContSenhas < 4 Then
        If fGeraSenha(Trim(SemFormatoCPF_CNPJ(Me.txtCNPJ.Text))) = Trim(Me.txtContraSenha.Text) Then
            Ok = True
            nContSenhas = 4
            Unload Me
            Exit Sub
        Else
            MsgBox "Senha incorreta. Tente novamente", vbOKOnly, "Aviso"
            Me.txtContraSenha.SetFocus
            Exit Sub
        End If
'    Next
    Else
        MsgBox "Total de tentativas de senha excedido, saindo do sistema", vbCritical, "Aviso"
        'Unload frmMenu
        End
    End If
End Sub


Private Sub txtCNPJ_GotFocus()
  With txtCNPJ
      .SelStart = 0
      .SelLength = Len(.Text)
   End With

End Sub


Private Sub txtContraSenha_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
       OKButton_Click
    Else
       If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
           KeyAscii = 0
       End If
   End If
End Sub
