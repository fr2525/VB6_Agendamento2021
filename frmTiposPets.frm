VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTipos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tipos de Pets"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   5685
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_Sair 
      Caption         =   "Sair"
      Height          =   765
      Left            =   4470
      Picture         =   "frmTiposPets.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   915
   End
   Begin VB.CommandButton cmd_Excluir 
      Caption         =   "&Excluir"
      Height          =   765
      Left            =   3720
      Picture         =   "frmTiposPets.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmd_Gravar 
      Caption         =   "&Gravar"
      Enabled         =   0   'False
      Height          =   765
      Left            =   3050
      Picture         =   "frmTiposPets.frx":062C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   915
   End
   Begin VB.CommandButton cmd_Limpar 
      Caption         =   "&Limpar"
      Height          =   765
      Left            =   1630
      Picture         =   "frmTiposPets.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   915
   End
   Begin VB.CommandButton cmd_Voltar 
      Caption         =   "&Retornar"
      Height          =   765
      Left            =   1680
      Picture         =   "frmTiposPets.frx":1090
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Adicionar 
      Caption         =   "&Novo"
      Height          =   765
      Left            =   210
      Picture         =   "frmTiposPets.frx":15C2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   915
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
      Left            =   180
      MaxLength       =   50
      TabIndex        =   1
      Top             =   4020
      Width           =   5220
   End
   Begin MSComctlLib.ListView lstTipos 
      Height          =   3645
      Left            =   180
      TabIndex        =   0
      ToolTipText     =   "Botï¿½o direito para Alterar/Excluir"
      Top             =   240
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   6429
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
      Caption         =   "Ediï¿½ï¿½o"
      Visible         =   0   'False
      Begin VB.Menu mnuAlterar 
         Caption         =   "&Alterar"
      End
      Begin VB.Menu mnuExcluir 
         Caption         =   "&Excluir"
      End
   End
End
Attribute VB_Name = "frmTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iTipoOperacao As Integer
Dim lvListItems_Itens As MSComctlLib.ListItem

Private Sub Carrega_Colunas_Tipos()
    With lstTipos
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Código", 300, lvwColumnLeft
        .ColumnHeaders.Add 2, , "Descrição", 4900, lvwColumnLeft
    End With
End Sub

Private Sub MontaColunas_Tipos()
    
    Call sConectaBanco
    strSql = ""
    strSql = strSql & " SELECT ID,DESCRICAO FROM TAB_tipos_pets ORDER BY DESCRICAO"
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    If Rstemp.RecordCount <> 0 Then
        Rstemp.MoveLast
        Rstemp.MoveFirst
 
        For X = 1 To Rstemp.RecordCount
            lstTipos.ListItems.Add X, , Rstemp!id
            
            If Not IsNull(Rstemp!Descricao) Then
                lstTipos.ListItems(X).SubItems(1) = Rstemp!Descricao
            Else
                lstTipos.ListItems(X).SubItems(1) = ""
            End If
            Rstemp.MoveNext
        Next
      
    Else
        MsgBox "Sem registros", vbOKOnly
    End If
    
    Rstemp.Close
    Set Rstemp = Nothing
    
End Sub

Private Sub cmd_Adicionar_Click()
    
    cmd_Adicionar.Enabled = False
    cmd_Adicionar.Visible = False
    cmd_Gravar.Enabled = True
    cmd_Excluir.Enabled = False
    cmd_Voltar.Enabled = True
    cmd_Voltar.Visible = True
    iTipoOperacao = 1
    txtAnimal.Enabled = True
    txtAnimal.text = ""
    txtAnimal.SetFocus
    lstTipos.Enabled = False
End Sub

Private Sub cmd_Excluir_Click()
    If Len(txtAnimal.text) = 0 Or txtAnimal.text = "" Then
       MsgBox "Tipo de Animal invï¿½lido. Favor corrigir", vbOKOnly
       txtAnimal.SetFocus
       Exit Sub
    End If
    
    If MsgBox("Tem certeza que deseja excluir o tipo de animal: " & Chr(13) & Chr(10) & _
                            Trim(lstTipos.SelectedItem.ListSubItems.Item(1)), _
                            vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me") = vbYes Then
       Call sConectaBanco
       strSql = ""
       strSql = strSql & " SELECT count(*) as contador FROM TAB_pets "
       strSql = strSql & " WHERE tipo_ani = " & lstTipos.SelectedItem.text
       If Rstemp.State = adStateOpen Then
           Rstemp.Close
       End If
       Set Rstemp = New ADODB.Recordset
       Rstemp.Open strSql, Cnn, 1, 2
       If Rstemp!Contador > 0 Then
            MsgBox "Exclusï¿½o nï¿½o permitida para tipo de animal em uso", vbCritical, "Aviso"
        Else
            If fExcluir_Tipo_Pet() Then
                cmd_Adicionar.Enabled = True
                cmd_Voltar.Enabled = False
                cmd_Gravar.Enabled = False
                cmd_Limpar.Enabled = False
                cmd_Voltar.Enabled = False
                cmd_Voltar.Visible = False
                lstTipos.ListItems.Clear
                Call MontaColunas_Tipos
                If lstTipos.ListItems.Count > 0 Then
                    lstTipos.ListItems(1).Selected = True
                    txtAnimal.text = Trim(lstTipos.SelectedItem.ListSubItems.Item(1))
                End If
            Else
                MsgBox "Erro ao excluir o tipo de PET: " & Err.Description
            End If
        End If
        Cnn.Close
    End If
End Sub

Private Sub cmd_Gravar_Click()
    If Len(txtAnimal.text) = 0 Or txtAnimal.text = "" Then
       MsgBox "Tipo de Animal invï¿½lido. Favor corrigir", vbOKOnly
       txtAnimal.SetFocus
       Exit Sub
    End If
    If fGravar_Tipo_Pet() Then
        cmd_Adicionar.Enabled = True
        cmd_Adicionar.Visible = True
        cmd_Gravar.Enabled = False
        cmd_Limpar.Enabled = False
        cmd_Voltar.Enabled = False
        cmd_Voltar.Visible = False
        lstTipos.ListItems.Clear
        Call MontaColunas_Tipos
        lstTipos.ListItems(1).Selected = True
        txtAnimal.text = Trim(lstTipos.SelectedItem.ListSubItems.Item(1))
        txtAnimal.Enabled = False
    Else
        MsgBox "Erro ao incluir o tipo de PET: " & Err.Description
    End If
    lstTipos.Enabled = True
    
End Sub

Private Sub cmd_Limpar_Click()
    txtAnimal.text = ""
    'txtAnimal.SetFocus
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True
End Sub

Private Sub cmd_Sair_Click()
    Unload Me
End Sub

Private Sub cmd_Voltar_Click()
    Call Form_Load
    cmd_Voltar.Enabled = False
    cmd_Voltar.Visible = False
    cmd_Adicionar.Enabled = True
    cmd_Adicionar.Visible = True
    cmd_Gravar.Enabled = False
    cmd_Limpar.Enabled = False
    Desabilita Me
    lstTipos.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}"
    End If
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
        mensagem = MsgBox("Informaï¿½ï¿½es nï¿½o Salvas. Deseja Sair assim mesmo ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
        If mensagem = vbNo Then
            Exit Sub
        Else
            Unload Me
        End If
    End If

End Sub

Private Sub Form_Load()
    Call Carrega_Colunas_Tipos
    Call MontaColunas_Tipos
    'lstTipos.ListItems = 1
    If lstTipos.ListItems.Count > 0 Then
        txtAnimal.text = Trim(lstTipos.SelectedItem.ListSubItems.Item(1))
    End If
    cmd_Voltar.Top = cmd_Adicionar.Top
    cmd_Voltar.Left = cmd_Adicionar.Left
    cmd_Voltar.Enabled = False
    cmd_Voltar.Visible = False
    cmd_Adicionar.Enabled = True
    cmd_Adicionar.Visible = True
    Desabilita Me
End Sub

Private Function fGravar_Tipo_Pet()
    
    If Len(txtAnimal.text) = 0 Or txtAnimal.text = "" Then
       MsgBox "Tipo de Animal invï¿½lido. Favor corrigir", vbOKOnly
       txtAnimal.SetFocus
       Exit Function
    End If
    
    fGravar_Tipo_Pet = True
    Call sConectaBanco
    On Error GoTo Erro_fGravar_Tipo_Pet
    If iTipoOperacao = 1 Then
        strSql = "INSERT INTO tab_tipos_pets (DESCRICAO, OPERADOR, DT_ATUALIZA)"
        strSql = strSql + " VALUES( '" & UCase(txtAnimal.text) & "','" & NomeUsuario & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        
    Else
        strSql = "UPDATE tab_tipos_pets SET DESCRICAO = '" & UCase(txtAnimal.text) & _
                                          "',OPERADOR = '" & NomeUsuario & _
                                          "', DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                          "' WHERE ID = " & lstTipos.SelectedItem.text
    End If
    Cnn.Execute strSql
    Cnn.Close
    lstTipos.Enabled = True
    
    Exit Function
Erro_fGravar_Tipo_Pet:
    fGravar_Tipo_Pet = False
End Function

Private Function fExcluir_Tipo_Pet()
    
    fExcluir_Tipo_Pet = True
    
    On Error GoTo Erro_fExcluir_Tipo_Pet
    
    strSql = "DELETE from tab_tipoS_pets WHERE ID = " & lstTipos.SelectedItem.text
    Cnn.Execute strSql
    Exit Function
Erro_fExcluir_Tipo_Pet:
    fExcluir_Tipo_Pet = False
End Function

Private Sub lstTipos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtAnimal.text = Trim(lstTipos.SelectedItem.ListSubItems.Item(1))
End Sub

Private Sub lstTipos_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If lstTipos.ListItems.Count > 0 Then
'            SendKeys "{tab}"
'        End If
'    Else
'        Call lstTipos_ItemClick
'    End If
End Sub

Private Sub lstTipos_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'   If Button = 2 Then
'        lstTipos.SetFocus
'        mnuEdicao.Visible = True
'        PopupMenu mnuEdicao, lstTipos.SelectedItem.Left + y, lstTipos.SelectedItem.Top + x
'    End If
    
    Set lvListItems_Itens = lstTipos.HitTest(X, y)

    'Check if a record was selected
    If lvListItems_Itens Is Nothing Then
        If lstTipos.ListItems.Count > 0 Then
            lstTipos.SelectedItem.Selected = False
        End If
        'se nï¿½o estiver item selecionado desabilita menus
        If Button = 2 Then
            PopupMenu mnuEdicao, , , , mnuAlterar
        End If
        Exit Sub
    Else
        'Habilita menus
        lvListItems_Itens.Selected = True
        If Button = 2 Then
            PopupMenu mnuEdicao, , , , mnuAlterar
        End If
    End If

End Sub

Private Sub mnuAlterar_Click()
    txtAnimal.Enabled = True
    cmd_Gravar.Enabled = True
    cmd_Limpar.Enabled = True
    cmd_Adicionar.Enabled = False
    cmd_Adicionar.Visible = False
    cmd_Voltar.Enabled = True
    cmd_Voltar.Visible = True
    txtAnimal.SetFocus
    iTipoOperacao = 2
End Sub

Private Sub mnuExcluir_Click()
    Call cmd_Excluir_Click
End Sub

Private Sub txtAnimal_GotFocus()
    SelText txtAnimal
End Sub

Private Sub txtAnimal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mensagem = MsgBox("Salvar Dados ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
        If mensagem = vbNo Then
            Unload Me
            Exit Sub
        Else
            If Len(txtAnimal.text) = 0 Then
                MsgBox "Favor digitar um tipo de animal vï¿½lido", vbOKOnly
                Cancel = True
                txtAnimal.SetFocus
                Exit Sub
            Else
                cmd_Gravar_Click
                Exit Sub
            End If
        End If
    End If

End Sub

Private Sub txtAnimal_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub txtAnimal_LostFocus()
'    If Len(txtAnimal.Text) = 0 Then
'       MsgBox "Favor digitar um tipo de animal vï¿½lido", vbOKOnly
'       Cancel = True
'    End If
End Sub



