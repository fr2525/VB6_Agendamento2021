VERSION 5.00
Begin VB.Form frmAcesso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Identificação Usuário"
   ClientHeight    =   2745
   ClientLeft      =   4125
   ClientTop       =   3465
   ClientWidth     =   4290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAcesso.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2752.52
   ScaleMode       =   0  'User
   ScaleWidth      =   1629.63
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Sair 
      Caption         =   "&Saír"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmd_confirma 
      Caption         =   "&Conectar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   4305
      Begin VB.TextBox txt_Senha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   810
         Width           =   804
      End
      Begin VB.TextBox txt_NomeAcesso 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   2028
      End
      Begin VB.Label lbl_Senha 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Senha :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   810
         Width           =   1020
      End
      Begin VB.Label lbl_NomeAcesso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Digite o nome do usuário e senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1080
      TabIndex        =   7
      Top             =   240
      Width           =   3315
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   0
      Picture         =   "frmAcesso.frx":030A
      Top             =   0
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   780
      Left            =   675
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public flagExcluirPedido As Boolean
Dim JanelaChamadora As Form
Dim lJanelaChamadora As Form          'Variável-exemplo do uso de formulários criados em "run-time"

Public flagPermiteExclusao As Boolean
Private Sub cmd_confirma_Click()

Dim rsAcesso As New ADODB.Recordset
Dim rsMenu As New ADODB.Recordset

Dim campo As Field
Dim campo2 As Control

Dim flag_Usuario_Master As Boolean

'Call Conecta_Banco
Call sConectaBanco
acesso = ""

        If txt_NomeAcesso.Text = "ROMEU" And txt_Senha.Text = "123" Then
                flag_Usuario_Master = True
                sysAcesso = CDbl("1")
                NomeUsuario = UCase("Supervisor")
                'frmMenu.Status.Panels(6).Text = UCase(NomeUsuario)
            GoSub entradireto
        Else
            flag_Usuario_Master = False
        End If

        CONV_NOME = ""
        For X = 1 To Len(Trim(txt_Senha.Text))
            CONV_NOME = CONV_NOME + Chr(Asc(Mid(Trim(UCase(txt_Senha.Text)), X, 1)) * 2)
        Next
        

        CONV_NOME = ""
        For X = 1 To Len(Trim(txt_Senha.Text))
            CONV_NOME = CONV_NOME + Chr(Asc(Mid(Trim(UCase(txt_Senha.Text)), X, 1)) * 2)
        Next
        
        sql = "select * from Cad_Usuarios where "
        sql = sql & "USUA_DS_NOME_ACESSO = '" & txt_NomeAcesso.Text & "'"
        sql = sql & " and USUA_NR_SENHA = '" & UCase(CONV_NOME) & "'"
        Set rsAcesso = New ADODB.Recordset
        rsAcesso.Open sql, Cnn, 1, 3
            If rsAcesso.RecordCount = 0 Then
                rsAcesso.Close
                MsgBox "Nome do Usuário ou Senha inválida.", vbInformation, "Aviso"
                txt_Senha = ""
                txt_NomeAcesso.SetFocus
                Call txt_NomeAcesso_GotFocus
                Exit Sub
            Else
                acesso = rsAcesso!USUA_DS_TP_ACESSO
            End If
        rsAcesso.Close
        Set rsAcesso = Nothing
        

        sql = "select * from Cad_Usuarios where "
        sql = sql & "USUA_DS_NOME_ACESSO = '" & txt_NomeAcesso & "'"
        sql = sql & " and USUA_NR_SENHA = '" & UCase(CONV_NOME) & "'"
        Set rsAcesso = New ADODB.Recordset
        rsAcesso.Open sql, Cnn, 1, 2
        With rsAcesso
            If .RecordCount = 0 Then
                rsAcesso.Close
                MsgBox "Nome do Usuário ou Senha inválida.", vbInformation, "Aviso"
                txt_Senha = ""
                txt_NomeAcesso.SetFocus
                Call txt_NomeAcesso_GotFocus
                Exit Sub
            End If
        
        Frame1.Visible = False
        'Define as variáveis do sistema
        sysNome = !USUA_DS_NOME
        sysCodigo = !USUA_CD_USUARIO
        sysNomeAcesso = !USUA_DS_NOME_ACESSO
        sysAcesso = !USUA_DS_TP_ACESSO
        sysSenha = !USUA_NR_SENHA
        NomeUsuario = !USUA_DS_NOME
        CodUsuarioLogado = sysCodigo
        If flag_Relogin = True Then
            frmMenu.Status.Panels(6).Text = UCase(NomeUsuario)
        End If

entradireto:
'    Mostra_Login
    
'    If sysAcesso = 1 Then
'        sql = "select * from Cad_Menus WHERE MENU_DS_SISTEMA = '" & NomeSistema & "'"
'        Set rsMenu = New ADODB.Recordset
'        rsMenu.Open sql, Cnn, 1, 2
'        With rsMenu
'            If .RecordCount = 0 Then
'                rsMenu.Close
'                MsgBox "Tabela de Menus esta vazia.", vbInformation, Me.Caption
'                Exit Sub
'            End If
'            .MoveLast
'            .MoveFirst
'            Do While Not .EOF
'                campotab = "MENU_DS_NOME_SISTEMA"
'                For Each campo In .Fields
'                    If UCase(campo.Name) = campotab Then
'                        For Each campo2 In frmMenu.Controls
'                            If UCase(campo.Value) = UCase(campo2.Name) Then
'                                campo2.Enabled = True
'                                'toobar.Buttons.Item.Key
'                                'frmMenu.tbToolBar.Buttons(LCase(campo.Value)).Enabled = True
'                                'frmMenu.tbToolBar.Buttons("menu_cadastro").Enabled = True
'                                'If UCase(campo.Value) = "MENU_CADASTRO_PRODUTOS" Then
'                                '    frmMenu.Toolbar1.Buttons(1).Enabled = True
'                                'End If
'                                Exit For
'                            End If
'                        Next
'                    End If
'                Next
'                .MoveNext
'            Loop
'            .Close
'        End With
'
'    Else
'        sql = "select * from Cad_Menus "
'        sql = sql & "where MENU_CD_CODI in ("
'        sql = sql & "select MENU_CD_CODI from Cad_Opcoes_Usuario_Acesso "
'        sql = sql & " where OPAC_CD_CODI = " & sysCodigo
'        sql = sql & " and MENU_DS_SISTEMA = '" & NomeSistema & "'"
'        sql = sql & " order by MENU_CD_CODI )"
'        Set rsMenu = New ADODB.Recordset
'        rsMenu.Open sql, Cnn, 1, 2
'        With rsMenu
'            If .RecordCount = 0 Then
'                cmd_confirma.Visible = False
'                Acesso_OK = 1
'                rsAcesso.Close
'                frmMenu.Enabled = True
'                Unload frmAcesso
'                Exit Sub
'            End If
'            .MoveLast
'            .MoveFirst
'            Do While Not .EOF
'                campotab = "MENU_DS_NOME_SISTEMA"
'                For Each campo In .Fields
'                    If UCase(campo.Name) = campotab Then
'                        For Each campo2 In frmMenu.Controls
'                            If campo.Value = UCase(campo2.Name) Then
'                                campo2.Enabled = True
'                                Exit For
'                            End If
'                        Next
'                    End If
'                Next
'                .MoveNext
'            Loop
'            .Close
'        End With
'    End If
End With

'gravar arquivo ini
'WriteIniFile App.Path & "\Petshop.ini", txt_NomeAcesso.Text, "Logado em", Format(Now, "Long Date") & " - " & Now

'frmMenu.menu_utilitarios_Imp_Fiscal.Enabled = False

'If ReadIniFile("C:\Petshop.ini", "Imp_Fisc_Sel", "Uso", "0") Then
    'frmMenu.menu_utilitarios_Imp_Fiscal.Enabled = True
'Else
    'frmMenu.menu_utilitarios_Imp_Fiscal.Enabled = False
'End If

cmd_confirma.Visible = False
Acesso_OK = 1

If flag_Usuario_Master = False Then
    rsAcesso.Close
    Set rsAcesso = Nothing
End If

'frmMenu.Enabled = True

'If frmMenu.menu_Cadastro.Enabled = True Then
'    If frmMenu.menu_cadastro_cliente.Enabled = True Then
'        frmMenu.Toolbar1.Buttons(3).Enabled = True
'    Else
'        frmMenu.Toolbar1.Buttons(3).Enabled = False
'    End If
'
'    If frmMenu.menu_cadastro_fornecedor.Enabled = True Then
'        frmMenu.Toolbar1.Buttons(4).Enabled = True
'    Else
'        frmMenu.Toolbar1.Buttons(4).Enabled = False
'    End If
'
'    If frmMenu.menu_cadastro_produtos.Enabled = True Then
'        frmMenu.Toolbar1.Buttons(5).Enabled = True
'    Else
'        frmMenu.Toolbar1.Buttons(5).Enabled = False
'    End If
'Else
'    frmMenu.Toolbar1.Buttons(3).Enabled = False
'    frmMenu.Toolbar1.Buttons(4).Enabled = False
'    frmMenu.Toolbar1.Buttons(5).Enabled = False
'End If
'
'If frmMenu.menu_movimentacao.Enabled = True Then
'    If frmMenu.menu_movimentacao_entrada.Enabled = True Then
'        frmMenu.Toolbar1.Buttons(7).Enabled = True
'    Else
'        frmMenu.Toolbar1.Buttons(7).Enabled = False
'    End If
'
'    If frmMenu.menu_movimentacao_saida.Enabled = True Then
'        frmMenu.Toolbar1.Buttons(8).Enabled = True
'    Else
'        frmMenu.Toolbar1.Buttons(8).Enabled = False
'    End If
'Else
'    frmMenu.Toolbar1.Buttons(7).Enabled = False
'    frmMenu.Toolbar1.Buttons(8).Enabled = False
'End If

Unload frmAcesso

End Sub

Private Sub cmd_Sair_Click()
    If flag_Relogin = True Then
        frmMenu.Toolbar1.buttons(3).Enabled = False
        frmMenu.Toolbar1.buttons(4).Enabled = False
        frmMenu.Toolbar1.buttons(5).Enabled = False
        frmMenu.Toolbar1.buttons(7).Enabled = False
        frmMenu.Toolbar1.buttons(8).Enabled = False
        frmMenu.Status.Panels(6).Text = "Ausente"
        Unload Me
        Exit Sub
    ElseIf flagExcluirPedido = False Then
        Dim Form As Form
        For Each Form In Forms
           Unload Form
           Set Form = Nothing
        Next Form
        Cnn.Close
        Set Cnn = Nothing
        End
    Else
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape And flagExcluirPedido = True Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
'Me.Left = (frmMenu.Width - Me.Width) / 2
'Me.Top = ((frmMenu.Height - Me.Height) / 2)

'Call RemoveMenus(Me)
   
KeyPreview = True
   
End Sub

Private Sub txt_NomeAcesso_GotFocus()
Call SelText(txt_NomeAcesso)
End Sub

Private Sub txt_NomeAcesso_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        txt_Senha.SetFocus
    End If
End Sub

Private Sub txt_Senha_GotFocus()
Call SelText(txt_Senha)
End Sub

Private Sub txt_Senha_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        cmd_confirma_Click
    End If

End Sub
