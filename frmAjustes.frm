VERSION 5.00
Begin VB.Form frmAjustes 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Configuração"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   6045
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkOrdem 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprime ordem de coleta?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   14
      Top             =   2850
      Width           =   3225
   End
   Begin VB.TextBox txtFim2 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   4620
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1770
      Width           =   420
   End
   Begin VB.TextBox txtInicio2 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   4620
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1170
      Width           =   420
   End
   Begin VB.CommandButton cmd_Gravar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Gravar (Alt+G)"
      Height          =   765
      Left            =   1620
      Picture         =   "frmAjustes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   1395
   End
   Begin VB.CommandButton cmd_Sair 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Sair (Alt+S)"
      Height          =   765
      Left            =   3120
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmAjustes.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   1395
   End
   Begin VB.TextBox txtDuracao 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   3990
      MaxLength       =   3
      TabIndex        =   8
      Top             =   2310
      Width           =   1050
   End
   Begin VB.TextBox txtFim 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   3990
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1770
      Width           =   420
   End
   Begin VB.TextBox txtInicio 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   4020
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1170
      Width           =   420
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4380
      TabIndex        =   13
      Top             =   1740
      Width           =   195
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4380
      TabIndex        =   12
      Top             =   1110
      Width           =   195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Minutos)"
      Height          =   195
      Left            =   5130
      TabIndex        =   11
      Top             =   2370
      Width           =   645
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Configurações da agenda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Left            =   1380
      TabIndex        =   3
      Top             =   270
      Width           =   3600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Término das Atividades :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1260
      TabIndex        =   2
      Top             =   1800
      Width           =   2595
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio das Atividades :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Duração do atendimento :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1185
      TabIndex        =   0
      Top             =   2370
      Width           =   2670
   End
End
Attribute VB_Name = "frmAjustes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sHoraInicio, sHoraFim, sDuracao, sOrdem As String
Dim sTipoEmpresa As String

Private Sub cmd_Gravar_Click()
    If txtInicio.Text > 23 Then
        MsgBox "Hora de inicio invalida.", vbOKOnly, "Aviso"
        txtInicio.SetFocus
        Exit Sub
    End If
    If txtInicio2.Text > 59 Then
        MsgBox "Minutos de inicio invalidos.", vbOKOnly, "Aviso"
        txtInicio.SetFocus
        Exit Sub
    End If
    If txtFim.Text > 23 Then
        MsgBox "Hora final invalida.", vbOKOnly, "Aviso"
        txtFim.SetFocus
        Exit Sub
    End If
    If txtFim.Text > 59 Then
        MsgBox "Minutos Finais invalidos.", vbOKOnly, "Aviso"
        txtFim.SetFocus
        Exit Sub
    End If
    
    If Left(txtInicio.Text, 2) & Right(txtInicio2.Text, 2) > Left(txtFim.Text, 2) & Right(txtFim2.Text, 2) Then
        MsgBox "Hora Inicial maior que a hora final.", vbOKOnly, "Aviso"
        txtInicio.SetFocus
        Exit Sub
    End If
    
    sHoraInicio = txtInicio.Text & ":" & txtInicio2.Text
    sHoraFim = txtFim.Text & ":" & txtFim2.Text
    
    WriteIniFile App.Path & "\Petshop.ini", "HORA_INICIO", "", sHoraInicio
    WriteIniFile App.Path & "\Petshop.ini", "HORA_FIM", "", sHoraFim
    WriteIniFile App.Path & "\Petshop.ini", "DURACAO", "", sDuracao
    WriteIniFile App.Path & "\Petshop.ini", "ORDEM", "", IIf(chkOrdem.Value = 1, "S", "N")
    MsgBox "Operação efetuada com sucesso", vbInformation, " vbOkOnly"
    Unload Me
    
End Sub

Private Sub cmd_Sair_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyEscape And cmd_Gravar.Enabled = True Then
      mensagem = MsgBox("Salvar Dados ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
      If mensagem = vbNo Then
          Unload Me
          Exit Sub
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

ElseIf KeyCode = vbKeyF7 And cmd_Sair.Enabled = True Then
    cmd_Sair_Click
    Exit Sub
ElseIf KeyCode = vbKeyEscape And cmd_Sair.Enabled = True Then
    mensagem = MsgBox("Informações não Salvas. Deseja Sair assim mesmo ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
    If mensagem = vbNo Then
        Exit Sub
    Else
        Unload Me
    End If
End If


End Sub

Private Sub Form_Load()

    sTipoEmpresa = ReadIniFile(App.Path & "\Petshop.ini", "TIPO_EMPRESA", "", "")
    If sTipoEmpresa = "PETSHOP" Then
        sHoraInicio = ReadIniFile(App.Path & "\Petshop.ini", "HORA_INICIO", "", "")
        sHoraFim = ReadIniFile(App.Path & "\Petshop.ini", "HORA_FIM", "", "")
        sDuracao = ReadIniFile(App.Path & "\Petshop.ini", "DURACAO", "", "")
        sOrdem = ReadIniFile(App.Path & "\Petshop.ini", "ORDEM", "", "")
        
        If sHoraInicio = "" Then
            sHoraInicio = "00"
            shorainicio2 = "00"
        End If
        
        If sHoraFim = "" Then
            sHoraFim = "23"
            shorafim2 = "59"
        End If
        If sDuracao = "" Then
            sDuracao = "30"
        End If
    Else
        MsgBox "Empresa não é tipo PESHOP", vbCritical, "Aviso!"
        End
    End If
    
    txtDuracao.Text = Val(sDuracao)
    txtInicio.Text = Format(Left(sHoraInicio, 2), "00")
    txtInicio2.Text = Format(Mid(sHoraInicio, 4, 2), "00")
    txtFim.Text = Format(Left(sHoraFim, 2), "00")
    txtFim2.Text = Format(Mid(sHoraFim, 4, 2), "00")
    If sOrdem = "S" Then
        chkOrdem.Value = 1
    Else
        chkOrdem.Value = 0
    End If
End Sub

Private Sub txtDuracao_GotFocus()
    SelText txtDuracao
End Sub

Private Sub txtDuracao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmd_Gravar.SetFocus
    End If

End Sub

Private Sub txtFim_GotFocus()
   SelText txtFim
End Sub

Private Sub txtFim2_GotFocus()
   SelText txtFim2
End Sub

Private Sub txtFim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFim2.SetFocus
    End If

End Sub

Private Sub txtFim2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDuracao.SetFocus
    End If

End Sub

Private Sub txtInicio_GotFocus()
    SelText txtInicio
End Sub

Private Sub txtInicio2_GotFocus()
    SelText txtInicio2
End Sub

Private Sub txtInicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtInicio2.SetFocus
    End If
End Sub

Private Sub txtInicio2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFim.SetFocus
    End If
End Sub

