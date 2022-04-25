VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCadPets 
   Caption         =   "Cadastro de Pets"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   12960
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmd_Voltar 
      Caption         =   "&Retornar"
      Height          =   765
      Left            =   6600
      Picture         =   "frmCadPets.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6000
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Limpar 
      Caption         =   "&Limpar"
      Height          =   765
      Left            =   1470
      Picture         =   "frmCadPets.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5850
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Sair 
      Caption         =   "Sair"
      Height          =   765
      Left            =   5280
      Picture         =   "frmCadPets.frx":0A64
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5850
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Excluir 
      Caption         =   "&Excluir"
      Height          =   765
      Left            =   4011
      Picture         =   "frmCadPets.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5850
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Gravar 
      Caption         =   "&Gravar"
      Height          =   765
      Left            =   2744
      Picture         =   "frmCadPets.frx":1090
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5850
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Adicionar 
      Caption         =   "&Novo"
      Height          =   765
      Left            =   210
      Picture         =   "frmCadPets.frx":15C2
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5850
      Width           =   1155
   End
   Begin VB.Frame FrameDados 
      Caption         =   "Detalhes"
      Height          =   6405
      Left            =   6720
      TabIndex        =   6
      Top             =   180
      Width           =   5985
      Begin VB.CommandButton cmd_novo_Dono 
         Caption         =   "Novo"
         Height          =   495
         Left            =   5220
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox txtCodigoDono 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3900
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   780
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame FrameClie 
         BackColor       =   &H00FFFFFF&
         Height          =   945
         Left            =   1050
         TabIndex        =   15
         Top             =   5190
         Visible         =   0   'False
         Width           =   4605
         Begin VB.PictureBox Picture3 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   0
            Picture         =   "frmCadPets.frx":1AF4
            ScaleHeight     =   255
            ScaleWidth      =   10980
            TabIndex        =   17
            Top             =   0
            Width           =   11040
            Begin VB.PictureBox Pic_FecharFmeListaProd 
               AutoSize        =   -1  'True
               Height          =   270
               Left            =   8970
               Picture         =   "frmCadPets.frx":AD0A
               ScaleHeight     =   210
               ScaleWidth      =   240
               TabIndex        =   18
               ToolTipText     =   "Fechar"
               Top             =   0
               Width           =   300
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lista de Clientes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   225
               Left            =   -30
               TabIndex        =   19
               Top             =   0
               Width           =   4545
            End
         End
         Begin ComctlLib.ListView List_RazaoSocial 
            Height          =   465
            Left            =   30
            TabIndex        =   16
            Top             =   450
            Visible         =   0   'False
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   820
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.TextBox txtDono 
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
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   14
         Top             =   360
         Width           =   4110
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
         Left            =   180
         MaxLength       =   50
         TabIndex        =   13
         Top             =   750
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.TextBox txtPet 
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
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1080
         Width           =   4110
      End
      Begin VB.ComboBox cmbTipos 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmCadPets.frx":AFEC
         Left            =   3660
         List            =   "frmCadPets.frx":AFEE
         TabIndex        =   3
         Top             =   1890
         Width           =   2085
      End
      Begin VB.TextBox TxtObserv 
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
         Height          =   870
         Left            =   1590
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   3690
         Width           =   4110
      End
      Begin VB.TextBox txtCuidEspec 
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
         Height          =   570
         Left            =   1590
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2610
         Width           =   4110
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   1590
         TabIndex        =   2
         Top             =   1860
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483644
         CalendarTrailingForeColor=   16711935
         Format          =   163119105
         CurrentDate     =   42908
      End
      Begin VB.Label lbl_Animal 
         BackStyle       =   0  'Transparent
         Caption         =   "Pet  :"
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
         Left            =   375
         TabIndex        =   12
         Top             =   1110
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dono :"
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
         Height          =   330
         Left            =   210
         TabIndex        =   11
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Height          =   360
         Left            =   3000
         TabIndex        =   10
         Top             =   1890
         Width           =   600
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Nasc.  :"
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
         Left            =   375
         TabIndex        =   9
         Top             =   1890
         Width           =   1170
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Observ. :"
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
         Height          =   330
         Left            =   480
         TabIndex        =   8
         Top             =   3630
         Width           =   1035
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuidados Especiais :"
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
         Height          =   480
         Left            =   345
         TabIndex        =   7
         Top             =   2640
         Width           =   1200
      End
   End
   Begin MSComctlLib.ListView lstPets 
      Height          =   5205
      Left            =   150
      TabIndex        =   0
      ToolTipText     =   "Bot�o direito para Alterar/Excluir"
      Top             =   240
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   9181
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuEdicao 
      Caption         =   "Edicao"
      Visible         =   0   'False
      Begin VB.Menu mnuAlterar 
         Caption         =   "&Alterar"
      End
      Begin VB.Menu mnuExcluir 
         Caption         =   "&Excluir"
      End
   End
End
Attribute VB_Name = "frmCadPets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iTipoOperacao As Integer
Dim CodCliente As String

Private Sub Nomes_Colunas()
    With lstPets
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "C�digo", 0, lvwColumnLeft
        .ColumnHeaders.Add 2, , "Nome do Pet", 1300, lvwColumnLeft
        .ColumnHeaders.Add 3, , "Tipo do Pet", 1800, lvwColumnLeft
        .ColumnHeaders.Add 4, , "Propriet�rio", 4000, lvwColumnLeft
        'Colunas com tamanho zero apenas para guardar os valores
        .ColumnHeaders.Add 5, , "id_cli", 0, lvwColumnLeft
        .ColumnHeaders.Add 6, , "tipo_ani", 0, lvwColumnLeft
        .ColumnHeaders.Add 7, , "dt_nasc", 0, lvwColumnLeft
        .ColumnHeaders.Add 8, , "Pedigree", 0, lvwColumnLeft
        .ColumnHeaders.Add 9, , "observacoes", 0, lvwColumnLeft
        .ColumnHeaders.Add 10, , "cuidados_especiais", 0, lvwColumnLeft
    End With
End Sub

Private Sub Dados_Colunas()
    
    Call sConectaBanco
    strSql = ""
    strSql = strSql & " SELECT A.ID" & _
                      ",A.Id_cli " & _
                      ",A.Tipo_ani " & _
                      ",A.nome" & _
                      ",A.dt_nasc " & _
                      ",A.pedigree " & _
                      ",A.observacoes" & _
                      ",A.Cuidados_especiais" & _
                      ",B.descricao AS Tipo_Animal" & _
                      ",c.razao_social AS Nome_dono" & _
                      " FROM tab_pets A, tab_tipos_pets B, tab_clientes C " & _
                      " WHERE a.id_cli = c.id " & _
                      " AND a.tipo_ani = b.Id " & _
                      " ORDER BY a.nome "

    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    If Rstemp.RecordCount <> 0 Then
        Rstemp.MoveLast
        Rstemp.MoveFirst
        'fmeListaPedidos.Visible = True
        
        For X = 1 To Rstemp.RecordCount
            With lstPets
                .ListItems.Add X, , Rstemp!id
                
                If Not IsNull(Rstemp!Nome) Then
                    .ListItems(X).SubItems(1) = Trim(Rstemp!Nome)
                Else
                    .ListItems(X).SubItems(1) = ""
                End If
                
                If Not IsNull(Rstemp!tipo_Animal) Then
                    'object.SubItems(index) [= string]
                    .ListItems(X).SubItems(2) = Trim(Rstemp!tipo_Animal)
                Else
                    .ListItems(X).SubItems(2) = "SEM TIPO"
                End If
                If Not IsNull(Rstemp!NOME_DONO) Then
                    .ListItems(X).SubItems(3) = Trim(Rstemp!NOME_DONO)
                Else
                    .ListItems(X).SubItems(3) = "SEM DONO"
                End If
                '*********************************************
                If Not IsNull(Rstemp!id_cli) Then
                    .ListItems(X).SubItems(4) = Rstemp!id_cli
                Else
                    .ListItems(X).SubItems(4) = ""
                End If
                        
                If Not IsNull(Rstemp!tipo_ani) Then
                    .ListItems(X).SubItems(5) = Rstemp!tipo_ani
                Else
                    .ListItems(X).SubItems(5) = ""
                End If
                If Not IsNull(Rstemp!DT_NASC) Then
                    .ListItems(X).SubItems(6) = Rstemp!DT_NASC
                Else
                    .ListItems(X).SubItems(6) = ""
                End If
                        
                If Not IsNull(Rstemp!pedigree) Then
                    .ListItems(X).SubItems(7) = Rstemp!pedigree
                Else
                    .ListItems(X).SubItems(7) = ""
                End If

                If Not IsNull(Rstemp!observacoes) Then
                    .ListItems(X).SubItems(8) = Rstemp!observacoes
                Else
                    .ListItems(X).SubItems(8) = ""
                End If
                        
                If Not IsNull(Rstemp!cuidados_Especiais) Then
                    .ListItems(X).SubItems(9) = Rstemp!cuidados_Especiais
                Else
                    .ListItems(X).SubItems(9) = ""
                End If
        
            End With
            Rstemp.MoveNext
        Next
'    Else
'        MsgBox "Sem registros", vbOKOnly
    End If
    
    Rstemp.Close
    Set Rstemp = Nothing
    
End Sub


Private Sub Carrega_campos_texto()
    If lstPets.ListItems.Count > 0 Then
        txtCodigoDono.text = Trim(lstPets.SelectedItem.ListSubItems.Item(4))
        txtPet.text = Trim(lstPets.SelectedItem.ListSubItems.Item(1))
        txtTipo.text = Trim(lstPets.SelectedItem.ListSubItems.Item(2))
        txtDono.text = Trim(lstPets.SelectedItem.ListSubItems.Item(3))
        DTPicker1.Value = Trim(lstPets.SelectedItem.ListSubItems.Item(6))
        TxtObserv.text = Trim(lstPets.SelectedItem.ListSubItems.Item(8))
        txtCuidEspec.text = Trim(lstPets.SelectedItem.ListSubItems.Item(9))
    End If
    
End Sub

Private Sub cmbTipos_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
       txtCuidEspec.SetFocus
    End If

End Sub

Private Sub cmbTipos_LostFocus()
    If cmbTipos.ListIndex = -1 Then
        MsgBox "Favor entrar com um tipo de Pet", vbCritical, "Aviso"
        cmbTipos.SetFocus
    Else
        txtCuidEspec.SetFocus
    End If
End Sub

Private Sub cmd_Adicionar_Click()
    txtDono.Visible = True
    txtTipo.Visible = False
    DTPicker1.Enabled = True
    cmd_novo_Dono.Enabled = True
    Call Habilita(Me)
    txtDono.text = ""
    txtPet.text = ""
    'DTPicker1.value = ""
    txtCuidEspec.text = ""
    TxtObserv.text = ""
    'cmbDonos.Visible = True
    cmbTipos.Visible = True
    'If Not DadosCBOtabela(cmbDonos, "cliente", "razao_social", "codigo") Then
    '    MsgBox "N�o existem clientes para exibir. Favor cadastrar", vbCritical, "Aviso"
    '    Unload Me
    '    Exit Sub
    'End If
    
    If Not DadosCBOtabela(cmbTipos, "tab_tipos_pets", "descricao", "ID") Then
        MsgBox "N�o existem Tipos de Pets para exibir. Favor cadastrar", vbCritical, "Aviso"
        Unload Me
        Exit Sub
    End If
    
    cmd_Adicionar.Enabled = False
    cmd_Adicionar.Visible = False
    cmd_Gravar.Enabled = True
    cmd_Excluir.Enabled = False
    cmd_Voltar.Enabled = True
    cmd_Voltar.Visible = True
    lstPets.Enabled = False
    'cmbDonos.SetFocus
    iTipoOperacao = 1
    txtDono.SetFocus
End Sub

Private Sub cmd_Excluir_Click()

    If MsgBox("Tem certeza que deseja excluir o Pet: " & _
                            Trim(txtPet.text) & Chr(13) & Chr(10) & _
                            " pertencente ao cliente: " & Trim(txtDono.text), _
                            vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        Call sConectaBanco
        strSql = ""
        strSql = strSql & " SELECT count(*) as contador FROM TAB_atendimentos "
        strSql = strSql & " WHERE idAnimal = " & lstPets.SelectedItem.text
        If Rstemp.State = adStateOpen Then
            Rstemp.Close
        End If
        Set Rstemp = New ADODB.Recordset
        Rstemp.Open strSql, Cnn, 1, 2
        If Rstemp!Contador > 0 Then
            MsgBox "Exclus�o n�o permitida para Pet com atendimento efetuado para ele", vbCritical, "Aviso"
        Else
            If fExcluir_Pet() Then
                cmd_Adicionar.Enabled = True
                cmd_Excluir.Enabled = True
                cmd_Gravar.Enabled = False
                lstPets.ListItems.Clear
                Call Dados_Colunas
                If lstPets.ListItems.Count > 0 Then
                    lstPets.ListItems(1).Selected = True
                    txtPet.text = Trim(lstPets.SelectedItem.ListSubItems.Item(1))
                End If
            Else
                MsgBox "Erro ao excluir o Pet: " & Err.Description
            End If
        End If
    Else
        cmd_Adicionar.Enabled = True
        cmd_Excluir.Enabled = False
        cmd_Gravar.Enabled = False
        lstPets.ListItems.Clear
        Call Dados_Colunas
        If lstPets.ListItems.Count > 0 Then
            lstPets.ListItems(1).Selected = True
            txtPet.text = Trim(lstPets.SelectedItem.ListSubItems.Item(1))
        End If
    End If
    lstPets.Enabled = True
End Sub

Private Sub cmd_Gravar_Click()
    
    If Len(txtPet.text) = 0 Or txtPet.text = "" Then
        MsgBox "Nome do Pet inv�lido. Favor corrigir", vbOKOnly
        txtPet.SetFocus
        Exit Sub
    End If
    
    If cmbTipos.ListIndex = -1 Then
        MsgBox "Favor escolher um tipo de animal para o Pet.", vbOKOnly
        cmbTipos.SetFocus
        Exit Sub
    End If
    
    If fGravar_Pet() Then
        cmd_Adicionar.Enabled = True
        cmd_Adicionar.Visible = True
        cmd_Voltar.Enabled = False
        cmd_Voltar.Visible = False
        cmd_Gravar.Enabled = False
        cmd_Limpar.Enabled = False
        'cmd_Excluir.Enabled = true
        lstPets.ListItems.Clear
        Call Dados_Colunas
        'lstPets.ListItems(1).Selected = True
        Call Carrega_campos_texto
        Call Desabilita(Me)
        
        cmbTipos.Visible = False
        txtDono.Visible = True
        txtTipo.Visible = True
        DTPicker1.Enabled = False
        txtPet.Visible = True
        txtPet.text = Trim(lstPets.SelectedItem.ListSubItems.Item(1))
        'Call cmd_Limpar_Click
    Else
        MsgBox "Erro ao incluir o tipo de PET: " & Err.Description
    End If
    lstPets.Enabled = True
    
End Sub

Private Sub cmd_Limpar_Click()
    FrameDados.Enabled = True
    txtDono.text = ""
    txtCuidEspec.text = ""
    TxtObserv.text = ""
    txtPet.text = ""
    'txtServico.SetFocus
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True
    
End Sub

Private Sub cmd_novo_Dono_Click()
    Consulta = "S"
    frmCliente.IncluirCliente = 2
    frmCliente.MskCliDesde.text = Format(Date, "dd/mm/yyyy")
    frmCliente.TxtRazaoSocial.text = txtDono.text
    frmCliente.Show 1
'
'    Me.Visible = False
'    frmCliente.IncluirCliente = 2
'    frmCliente.Show vbModal
'    Me.Visible = True
'    txtPet.SetFocus
End Sub

Private Sub cmd_Sair_Click()
   Unload Me
End Sub

Private Sub cmd_Voltar_Click()
    Call Form_Load
    'cmd_Voltar.Top = cmd_Adicionar.Top
    'cmd_Voltar.Left = cmd_Adicionar.Left
    cmd_Voltar.Enabled = False
    cmd_Voltar.Visible = False
    cmd_Adicionar.Enabled = True
    cmd_Adicionar.Visible = True
    cmd_Gravar.Enabled = False
    cmd_Limpar.Enabled = False
    cmd_novo_Dono.Enabled = False
    
    Desabilita Me
    DTPicker1.Enabled = False
    lstPets.Enabled = True
    
End Sub
Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       cmbTipos.SetFocus
    End If
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmbTipos.SetFocus
    End If
End Sub

Private Sub Form_Load()
    
    Call Nomes_Colunas
    Call Dados_Colunas
    'txtDono.Top = cmbDonos.Top
    'txtDono.Left = cmbDonos.Left
    txtDono.Visible = True
    txtTipo.Top = cmbTipos.Top
    txtTipo.Left = cmbTipos.Left
    txtTipo.Visible = True
    cmd_Voltar.Visible = False
    cmd_Adicionar.Enabled = True
    cmd_Adicionar.Visible = True
    cmd_Voltar.Top = cmd_Adicionar.Top
    cmd_Voltar.Left = cmd_Adicionar.Left
    Me.Height = 7185
    
    'List_RazaoSocial.Height = 4305
    'List_RazaoSocial.Visible = False
    FrameDados.Height = lstPets.Height
    FrameClie.Top = txtDono.Top
    FrameClie.Left = txtDono.Left
    FrameClie.Height = 4305
    FrameClie.Visible = False
    List_RazaoSocial.Height = 4300
    'List_RazaoSocial.Visible = False

    
    Desabilita Me
    
    'lstServicos.ListItems = 1
    If lstPets.ListItems.Count > 0 Then
        txtPet.text = Trim(lstPets.SelectedItem.ListSubItems.Item(1))
        txtTipo.text = Trim(lstPets.SelectedItem.ListSubItems.Item(2))
        txtDono.text = Trim(lstPets.SelectedItem.ListSubItems.Item(3))
        DTPicker1.Value = Trim(lstPets.SelectedItem.ListSubItems.Item(6))
        TxtObserv.text = Trim(lstPets.SelectedItem.ListSubItems.Item(8))
        txtCuidEspec.text = Trim(lstPets.SelectedItem.ListSubItems.Item(9))
    End If

End Sub

Private Function fGravar_Pet()
    
    fGravar_Pet = True
    
    On Error GoTo Erro_fGravar_Pet
    Call sConectaBanco
    'ID,DESCRICAO,valor,tempo_est
    If iTipoOperacao = 1 Then
        strSql = "INSERT INTO tab_pets (ID_CLI,TIPO_ANI, NOME, DT_NASC, OBSERVACOES,CUIDADOS_ESPECIAIS,OPERADOR, DT_ATUALIZA)"
        strSql = strSql + " VALUES( " & List_RazaoSocial.SelectedItem.SubItems(1)
        strSql = strSql + "," & cmbTipos.ItemData(cmbTipos.ListIndex)
        strSql = strSql + ",'" & txtPet.text & "','" & Format(DTPicker1.Value, "yyyy/mm/dd") & "'"
        strSql = strSql + ",'" & TxtObserv.text & "','" & txtCuidEspec & "'"
        strSql = strSql + ",'" & NomeUsuario & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        
    Else
        strSql = "UPDATE tab_pets SET ID_ClI = " & Val(lstPets.SelectedItem.SubItems(4)) & _
                                        ",TIPO_ANI = " & cmbTipos.ItemData(cmbTipos.ListIndex) & _
                                        ",NOME = '" & Trim(UCase(txtPet.text)) & _
                                        "',DT_NASC = '" & Format(DTPicker1.Value, "yyyy/mm/dd") & _
                                        "',OBSERVACOES =  '" & Trim(UCase(TxtObserv.text)) & _
                                        "',CUIDADOS_ESPECIAIS = '" & Trim(UCase(txtCuidEspec)) & _
                                        "',OPERADOR = '" & NomeUsuario & _
                                        "',DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "'" & _
                                        " WHERE ID = " & lstPets.SelectedItem.text
                                                 
    End If
    Cnn.Execute strSql
   ' Cnn.Close
    Exit Function
    
Erro_fGravar_Pet:
    fGravar_Pet = False
End Function

Private Function fExcluir_Pet()
    
    fExcluir_Pet = True
    
    On Error GoTo Erro_fExcluir_Pet
    
    strSql = "DELETE from tab_pets WHERE ID = " & lstPets.SelectedItem.text
    Cnn.Execute strSql
    Exit Function
Erro_fExcluir_Pet:
    fExcluir_Pet = False
End Function

Private Sub lstPets_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call Carrega_campos_texto
End Sub

Private Sub lstPets_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = 2 Then
        lstPets.SetFocus
        'mnuEdicao.Visible = True
        PopupMenu mnuEdicao, lstPets.SelectedItem.Left + y, lstPets.SelectedItem.Top + X
    End If

End Sub

Private Sub mnuAlterar_Click()
    Dim inicio
    Dim Contador
    Dim Valor
    Dim meio
    
    'txtDono.Visible = False
    
    txtTipo.Visible = False
    DTPicker1.Enabled = True
    txtPet.Enabled = True
    txtCuidEspec.Enabled = True
    TxtObserv.Enabled = True
    txtPet.SetFocus
    'txtDono.Text = ""
    'DTPicker1.value = ""
    'txtCuidEspec.Text = ""
    'TxtObserv.Text = ""
    'cmbDonos.Enabled = True
    'cmbDonos.Visible = True
    cmbTipos.Enabled = True
    cmbTipos.Visible = True
    'DadosCBOtabela cmbDonos, "cliente", "razao_social", "codigo"
    DadosCBOtabela cmbTipos, "tab_tipos_pets", "descricao", "ID"
    
    cmd_Gravar.Enabled = True
    cmd_Limpar.Enabled = True
    cmd_Adicionar.Enabled = False
    cmd_Adicionar.Visible = False
    cmd_Voltar.Enabled = True
    cmd_Voltar.Visible = True

    inicio = 1
    'contador = cmbDonos.ListCount - 1
    'valor = lstPets.SelectedItem.SubItems(4)
    'For i = inicio To contador
    '    'cmbDonos.ListIndex = inicio
    '    If valor = cmbDonos.ItemData(i) Then
    '        cmbDonos.ListIndex = i
    '        Exit For
    '    End If
    'Next
    
    inicio = 0
    Contador = cmbTipos.ListCount - 1
    Valor = lstPets.SelectedItem.SubItems(5)
    For I = inicio To Contador
        'cmbDonos.ListIndex = inicio
        If Valor = cmbTipos.ItemData(I) Then
            cmbTipos.ListIndex = I
            Exit For
        End If
    Next
    'cmbDonos.SetFocus
    iTipoOperacao = 2

End Sub

Private Sub mnuExcluir_Click()
    Call cmd_Excluir_Click
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
ElseIf KeyCode = vbKeyF2 And cmd_Adicionar.Enabled = True Then
    cmd_Adicionar_Click
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
    mensagem = MsgBox("Informa��es n�o Salvas. Deseja Sair assim mesmo ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
    If mensagem = vbNo Then
        Exit Sub
    Else
        Unload Me
    End If
End If
  
'*
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
ElseIf KeyCode = vbKeyF2 And cmd_Adicionar.Enabled = True Then
    cmd_Adicionar_Click
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
    mensagem = MsgBox("Informa��es n�o Salvas. Deseja Sair assim mesmo ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
    If mensagem = vbNo Then
        Exit Sub
    Else
        Unload Me
    End If
End If

End Sub

Private Sub txtCuidEspec_GotFocus()
    SelText txtCuidEspec
End Sub

Private Sub txtCuidEspec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtObserv.SetFocus
    End If
End Sub

Private Sub txtCuidEspec_LostFocus()
    txtCuidEspec.text = UCase(Trim(txtCuidEspec.text))
End Sub

Private Sub txtDono_GotFocus()
   SelText txtDono
End Sub

Private Sub txtDono_KeyPress(KeyAscii As Integer)
    
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))

    If KeyAscii = 13 Then
        'If IsNumeric(txtDono.Text) = True And Trim(txtDono.Text) <> "" And Trim(Val(txtDono.Text)) <> 0 Then
        '    txtDono.Text = Val(txtDono.Text)
        '    Call txtDono_KeyPress(13)
        '    Exit Sub
        'End If
        If Trim(txtDono.text) <> "" Then
            txtDono.text = Trim(UCase(txtDono.text))
        End If
        If CarregaListaCliente = True Then
            FrameClie.Visible = True
            List_RazaoSocial.Visible = True
            List_RazaoSocial.SetFocus
        Else
            FrameClie.Visible = False
            List_RazaoSocial.Visible = False
        End If
    End If

End Sub

'Private Sub txtDono_LostFocus()
'    txtPet.SetFocus
'End Sub

Private Sub TxtObserv_GotFocus()
    SelText TxtObserv
End Sub

Private Sub TxtObserv_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        mensagem = MsgBox("Salvar Dados ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
        If mensagem = vbNo Then
            Unload Me
            Exit Sub
        Else
           cmd_Gravar_Click
           Exit Sub
        End If
    End If

End Sub

Private Sub txtPet_GotFocus()
    SelText txtPet
End Sub

Private Sub txtPet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DTPicker1.SetFocus
    End If
End Sub

Private Sub txtPet_LostFocus()
    txtPet.text = UCase(Trim(txtPet.text))
End Sub

Private Sub Carrega_Colunas_Clientes()

    With List_RazaoSocial
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Cliente", 5950, lvwColumnLeft
        .ColumnHeaders.Add 2, , "C�digo", 0, lvwColumnLeft
    End With
End Sub


Function CarregaListaCliente() As Boolean
Dim sql      As String
On Error GoTo TrataPesquisa

    CarregaListaCliente = False
    
    List_RazaoSocial.ListItems.Clear
    
    Call Carrega_Colunas_Clientes
    
    I = 0
    
'    sql = ""
'    sql = sql & "Select CODIGO, RAZAO_SOCIAL "
'    sql = sql & " From CLIENTE "
'    sql = sql & " Where RAZAO_SOCIAL Like '%" & FiltraAspasSimples(Trim(txtcliente.Text)) & "%'"
'    sql = sql & " order by RAZAO_SOCIAL  asc "

    sql = "Select id, RAZAO_SOCIAL, CGC_CPF, CEP_PRINCIPAL, BAIRRO_END_PRINCIPAL FROM tab_CLIENTEs "
    sql = sql & " Where RAZAO_SOCIAL Like '%" & Trim(txtDono.text) & "%'"
    'sql = sql & " OR CGC_CPF Like '%" & FormatCPF_CNPJ(Trim(txtcliente.Text)) & "%'"
    'sql = sql & " OR CEP_PRINCIPAL Like '%" & Trim(txtcliente.Text) & "%'"
    'sql = sql & " OR ENDERECO_PRINCIPAL Like '%" & Trim(txtcliente.Text) & "%'"
    sql = sql & " order by RAZAO_SOCIAL  asc "
    
    If Rstemp.State = adStateOpen Then
        Rstemp.Close
    End If
    If Cnn.State = adStateOpen Then
        Cnn.Close
    End If
    
    Call sConectaBanco
    
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open sql, Cnn, 1, 2
    If Rstemp.RecordCount > 0 Then
        Do Until Rstemp.EOF
            I = I + 1
            List_RazaoSocial.ListItems.Add I, , UCase(Rstemp!RAZAO_SOCIAL)
            'If Not IsNull(Rstemp!BAIRRO_END_PRINCIPAL) Then
            '    List_RazaoSocial.ListItems(i).SubItems(1) = UCase(Trim(Rstemp!BAIRRO_END_PRINCIPAL))
            'Else
            '    List_RazaoSocial.ListItems(i).SubItems(1) = ""
            'End If
            List_RazaoSocial.ListItems(I).SubItems(1) = UCase(Trim(Rstemp!id))
            Rstemp.MoveNext
        Loop
    Else
        If MsgBox("Cliente n�o  Encontrado. " & vbNewLine & "Deseja Efetuar Cadastro..?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me") = vbNo Then
            CarregaListaCliente = False
            txtDono.SetFocus
            txtDono_GotFocus
            Exit Function
        Else
            Call cmd_novo_Dono_Click
            CarregaListaCliente = False
            Exit Function
        End If
    End If
    
    CarregaListaCliente = True

Exit Function

TrataPesquisa:

    MsgBox "Erro ao carregar lista de clientes. Descri��o do Erro: " & Chr(10) & Chr(13) & Err.Description, vbOKOnly, "Aviso"
    txtDono.SetFocus
    Exit Function

End Function

Private Sub List_RazaoSocial_DblClick()

If List_RazaoSocial.ListItems.Count > 0 Then
    For I = List_RazaoSocial.ListItems.Count To 1 Step -1
       If List_RazaoSocial.ListItems(I).Selected = True Then
           'CodCliente = List_RazaoSocial.ListItems.Item(i).Text
           'VARRE O VALOR DA COLUNA 2 = C�DIGO
           CodCliente = List_RazaoSocial.SelectedItem.SubItems(1)
           Exit For
       End If
    Next I
    
    If Rstemp.State = adStateOpen Then
        Rstemp.Close
    End If
    If Cnn.State = adStateOpen Then
        Cnn.Close
    End If
    
    Call sConectaBanco
    
    sql = "SELECT * FROM tab_Clientes WHERE id = " & CodCliente
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open sql, Cnn, 1, 2
    If Rstemp.RecordCount > 0 Then
        txtCodigoDono.text = CodCliente
        txtDono.text = List_RazaoSocial.SelectedItem
        List_RazaoSocial.Visible = False
        FrameClie.Visible = False
        
    End If
    Rstemp.Close
    txtPet.SetFocus
End If
Exit Sub

trataCliente:
  ' If Err.Number <> 0 Then
    'Erro "Selecionar cliente na lista"
    Exit Sub '
    'End If
End Sub

'Private Sub List_RazaoSocial_ItemClick(ByVal Item As MSComctlLib.ListItem)
'  itemPed_Orc = Item.Index
'End Sub


Private Sub List_RazaoSocial_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then
        Call List_RazaoSocial_DblClick
    ElseIf KeyCode = vbKeyEscape Then
'        Pic_FecharFmeListaProd_Click
    End If
End Sub


Private Sub List_RazaoSocial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List_RazaoSocial_DblClick
End If
End Sub

