VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFormas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Formas de pagamento"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6255
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSCommand cmd_Sair 
      Height          =   765
      Left            =   4710
      TabIndex        =   7
      Top             =   240
      Width           =   1155
      Caption         =   "Sair"
   End
   Begin Threed.SSCommand cmd_Excluir 
      Height          =   765
      Left            =   2160
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "&Excluir"
      Picture         =   "frmTipos.frx":010A
   End
   Begin Threed.SSCommand cmd_Gravar 
      Height          =   765
      Left            =   3210
      TabIndex        =   5
      Top             =   240
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "&Gravar"
      Picture         =   "frmTipos.frx":028C
   End
   Begin Threed.SSCommand cmd_Limpar 
      Height          =   765
      Left            =   1710
      TabIndex        =   4
      Top             =   240
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "&Limpar"
      Picture         =   "frmTipos.frx":040E
   End
   Begin Threed.SSCommand cmd_Voltar 
      Height          =   765
      Left            =   750
      TabIndex        =   3
      Top             =   5340
      Visible         =   0   'False
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "&Retornar"
      Picture         =   "frmTipos.frx":0860
   End
   Begin Threed.SSCommand cmd_Adicionar 
      Height          =   765
      Left            =   210
      TabIndex        =   2
      Top             =   240
      Width           =   1155
      _Version        =   65536
      _ExtentX        =   2037
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "&Novo"
      Picture         =   "frmTipos.frx":096A
   End
   Begin VB.TextBox txtForma 
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
      Top             =   5010
      Width           =   5700
   End
   Begin MSComctlLib.ListView lstFormas 
      Height          =   3645
      Left            =   180
      TabIndex        =   0
      ToolTipText     =   "Bot�o direito para Alterar/Excluir"
      Top             =   1200
      Width           =   5670
      _ExtentX        =   10001
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
      Caption         =   "Edi��o"
      Visible         =   0   'False
      Begin VB.Menu mnuAlterar 
         Caption         =   "&Alterar"
      End
      Begin VB.Menu mnuExcluir 
         Caption         =   "&Excluir"
      End
   End
End
Attribute VB_Name = "frmFormas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iTipoOperacao As Integer
Dim lvListItems_Itens As MSComctlLib.ListItem

Private Sub Carrega_colunas_formas()
    With lstFormas
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "C�digo", 300, lvwColumnLeft
        .ColumnHeaders.Add 2, , "Descri��o", 4900, lvwColumnLeft
    End With
End Sub

Private Sub Montacolunas_formas()
    
    Call sConectaBanco
    strSql = ""
    strSql = strSql & " SELECT ID,DESCRICAO FROM tab_formas_pagto ORDER BY DESCRICAO"
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    If Rstemp.RecordCount <> 0 Then
        Rstemp.MoveLast
        Rstemp.MoveFirst
        'fmeListaPedidos.Visible = True
        
        For X = 1 To Rstemp.RecordCount
            lstFormas.ListItems.Add X, , Rstemp!id
            
            If Not IsNull(Rstemp!Descricao) Then
                lstFormas.ListItems(X).SubItems(1) = Rstemp!Descricao
            Else
                lstFormas.ListItems(X).SubItems(1) = ""
            End If
            Rstemp.MoveNext
        Next
        'lstFormas.SetFocus
       
    Else
        MsgBox "Sem registros", vbOKOnly
        'fmeListaPedidos.Visible = False
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
    txtForma.Enabled = True
    txtForma.Text = ""
    txtForma.SetFocus
    lstFormas.Enabled = False
End Sub

Private Sub cmd_Excluir_Click()
    If Len(txtForma.Text) = 0 Or txtForma.Text = "" Then
       MsgBox "Forma de Pagamento inv�lida. Favor corrigir", vbOKOnly
       txtForma.SetFocus
       Exit Sub
    End If
    
    If MsgBox("Tem certeza que deseja excluir essa forma de pagamento: " & Chr(13) & Chr(10) & _
                            Trim(lstFormas.SelectedItem.ListSubItems.Item(1)), _
                            vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me") = vbYes Then
       If fExcluir_Forma() Then
          cmd_Adicionar.Enabled = True
          cmd_Voltar.Enabled = False
          cmd_Gravar.Enabled = False
          cmd_Limpar.Enabled = False
          cmd_Voltar.Enabled = False
          cmd_Voltar.Visible = False
          lstFormas.ListItems.Clear
          Call Montacolunas_formas
          If lstFormas.ListItems.Count > 0 Then
              lstFormas.ListItems(1).Selected = True
              txtForma.Text = Trim(lstFormas.SelectedItem.ListSubItems.Item(1))
          End If
       Else
          MsgBox "Erro ao excluir a forma de pagamento: " & Err.Description
       End If
       Cnn.Close
    End If
End Sub

Private Sub cmd_Gravar_Click()
    If Len(txtForma.Text) = 0 Or txtForma.Text = "" Then
       MsgBox "Forma de pagamento inv�lida. Favor corrigir", vbOKOnly
       txtForma.SetFocus
       Exit Sub
    End If
    If fGravar_Forma() Then
        cmd_Adicionar.Enabled = True
        cmd_Adicionar.Visible = True
        cmd_Gravar.Enabled = False
        cmd_Limpar.Enabled = False
        cmd_Voltar.Enabled = False
        cmd_Voltar.Visible = False
        lstFormas.ListItems.Clear
        Call Montacolunas_formas
        lstFormas.ListItems(1).Selected = True
        txtForma.Text = Trim(lstFormas.SelectedItem.ListSubItems.Item(1))
        txtForma.Enabled = False
    Else
        MsgBox "Erro ao incluir a Forma de pagamento: " & Err.Description
    End If
    lstFormas.Enabled = True
    
End Sub

Private Sub cmd_Limpar_Click()
    txtForma.Text = ""
    'txtForma.SetFocus
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
    lstFormas.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
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
        mensagem = MsgBox("Informa��es n�o Salvas. Deseja Sair assim mesmo ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
        If mensagem = vbNo Then
            Exit Sub
        Else
            Unload Me
        End If
    End If

End Sub

Private Sub Form_Load()
    Call Carrega_colunas_formas
    Call Montacolunas_formas
    'lstFormas.ListItems = 1
    If lstFormas.ListItems.Count > 0 Then
        txtForma.Text = Trim(lstFormas.SelectedItem.ListSubItems.Item(1))
    End If
    cmd_Voltar.Top = cmd_Adicionar.Top
    cmd_Voltar.Left = cmd_Adicionar.Left
    cmd_Voltar.Enabled = False
    cmd_Voltar.Visible = False
    cmd_Adicionar.Enabled = True
    cmd_Adicionar.Visible = True
    Desabilita Me
End Sub

Private Function fGravar_Forma()
    
    If Len(txtForma.Text) = 0 Or txtForma.Text = "" Then
       MsgBox "Forma de pagamento inv�lida. Favor corrigir", vbOKOnly
       txtForma.SetFocus
       Exit Function
    End If
    
    fGravar_Forma = True
    Call sConectaBanco
    On Error GoTo Erro_fGravar_Forma
    If iTipoOperacao = 1 Then
        strSql = "INSERT INTO tab_formas_pagto (DESCRICAO, OPERADOR, DT_ATUALIZA)"
        strSql = strSql + " VALUES( '" & UCase(txtForma.Text) & "','" & NomeUsuario & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        
    Else
        strSql = "UPDATE tab_formas_pagto SET DESCRICAO = '" & UCase(txtForma.Text) & _
                                          "',OPERADOR = '" & NomeUsuario & _
                                          "', DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                          "' WHERE ID = '" & lstFormas.SelectedItem.Text & "'"
    End If
    Cnn.Execute strSql
    Cnn.Close
    lstFormas.Enabled = True
    
    Exit Function
Erro_fGravar_Forma:
    fGravar_Forma = False
End Function

Private Function fExcluir_Forma()
    
    fExcluir_Forma = True
    
    On Error GoTo Erro_fExcluir_Forma
    
    strSql = "DELETE from tab_formas_pagto WHERE ID = '" & lstFormas.SelectedItem.Text & "'"
    Cnn.Execute strSql
    Exit Function
Erro_fExcluir_Forma:
    fExcluir_Forma = False
End Function

Private Sub lstFormas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtForma.Text = Trim(lstFormas.SelectedItem.ListSubItems.Item(1))
End Sub

Private Sub lstFormas_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If lstFormas.ListItems.Count > 0 Then
'            SendKeys "{tab}"
'        End If
'    Else
'        Call lstFormas_ItemClick
'    End If
End Sub

Private Sub lstFormas_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'   If Button = 2 Then
'        lstFormas.SetFocus
'        mnuEdicao.Visible = True
'        PopupMenu mnuEdicao, lstFormas.SelectedItem.Left + y, lstFormas.SelectedItem.Top + x
'    End If
    
    Set lvListItems_Itens = lstFormas.HitTest(X, y)

    'Check if a record was selected
    If lvListItems_Itens Is Nothing Then
        If lstFormas.ListItems.Count > 0 Then
            lstFormas.SelectedItem.Selected = False
        End If
        'se n�o estiver item selecionado desabilita menus
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
    txtForma.Enabled = True
    cmd_Gravar.Enabled = True
    cmd_Limpar.Enabled = True
    cmd_Adicionar.Enabled = False
    cmd_Adicionar.Visible = False
    cmd_Voltar.Enabled = True
    cmd_Voltar.Visible = True
    txtForma.SetFocus
    iTipoOperacao = 2
End Sub

Private Sub mnuExcluir_Click()
    Call cmd_Excluir_Click
End Sub

Private Sub txtForma_GotFocus()
    SelText txtForma
End Sub

Private Sub txtForma_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mensagem = MsgBox("Salvar Dados ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
        If mensagem = vbNo Then
            Unload Me
            Exit Sub
        Else
            If Len(txtForma.Text) = 0 Then
                MsgBox "Favor digitar uma forma de pagamento v�lida ", vbOKOnly
                Cancel = True
                txtForma.SetFocus
                Exit Sub
            Else
                cmd_Gravar_Click
                Exit Sub
            End If
        End If
    End If

End Sub

Private Sub txtForma_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub txtForma_LostFocus()
'    If Len(txtForma.Text) = 0 Then
'       MsgBox "Favor digitar um tipo de animal v�lido", vbOKOnly
'       Cancel = True
'    End If
End Sub



