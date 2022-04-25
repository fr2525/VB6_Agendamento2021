VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServicos 
   Caption         =   "Serviços"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkVacina 
      Alignment       =   1  'Right Justify
      Caption         =   "Vacina?"
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
      Height          =   405
      Left            =   3900
      TabIndex        =   4
      Top             =   4500
      Width           =   1275
   End
   Begin VB.CommandButton cmd_Voltar 
      Caption         =   "Retornar"
      Height          =   765
      Left            =   2790
      Picture         =   "frmTiposAtend.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5730
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Adicionar 
      Caption         =   "Novo"
      Height          =   765
      Left            =   180
      Picture         =   "frmTiposAtend.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   1155
   End
   Begin VB.CommandButton Cmd_limpar 
      Caption         =   "Limpar"
      Enabled         =   0   'False
      Height          =   765
      Left            =   1455
      Picture         =   "frmTiposAtend.frx":0A64
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Gravar 
      Caption         =   "Gravar"
      Enabled         =   0   'False
      Height          =   765
      Left            =   2760
      Picture         =   "frmTiposAtend.frx":0F96
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Sair 
      Caption         =   "Sair"
      Height          =   765
      Left            =   4080
      Picture         =   "frmTiposAtend.frx":14C8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Width           =   1155
   End
   Begin VB.CommandButton cmd_Excluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   765
      Left            =   1170
      Picture         =   "frmTiposAtend.frx":15C2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5580
      Visible         =   0   'False
      Width           =   1155
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
      Left            =   2340
      MaxLength       =   3
      TabIndex        =   3
      Top             =   4560
      Width           =   750
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
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
      Left            =   270
      MaxLength       =   14
      TabIndex        =   2
      Top             =   4560
      Width           =   1860
   End
   Begin VB.TextBox txtServico 
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
      Left            =   240
      MaxLength       =   50
      TabIndex        =   1
      Top             =   3840
      Width           =   4980
   End
   Begin MSComctlLib.ListView lstservicos 
      Height          =   3135
      Left            =   180
      TabIndex        =   0
      ToolTipText     =   "Botão direito para Alterar/Excluir"
      Top             =   240
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   5530
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
   Begin VB.Label Label3 
      Caption         =   "minutos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   3150
      TabIndex        =   9
      Top             =   4590
      Width           =   540
   End
   Begin VB.Label Label2 
      Caption         =   "Duração :"
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
      Height          =   330
      Left            =   2340
      TabIndex        =   8
      Top             =   4290
      Width           =   1050
   End
   Begin VB.Label Label1 
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
      Height          =   210
      Left            =   270
      TabIndex        =   7
      Top             =   4290
      Width           =   840
   End
   Begin VB.Label lbl_Animal 
      Caption         =   "Descrição :"
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
      Left            =   255
      TabIndex        =   6
      Top             =   3510
      Width           =   1380
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
Attribute VB_Name = "frmServicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iTipoOperacao As Integer
Dim lvListItems_Itens As MSComctlLib.ListItem

Private Sub Nomes_Colunas()
    With lstservicos
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Código", 0, lvwColumnLeft
        .ColumnHeaders.Add 2, , "Descrição", 4000, lvwColumnLeft
        .ColumnHeaders.Add 3, , "Valor", 1400, lvwColumnRight
        .ColumnHeaders.Add 4, , "Duração", 900, lvwColumnRight
        .ColumnHeaders.Add 5, , "Vacina?", 900, lvwColumnRight
    End With
End Sub


Private Sub Dados_Colunas()
    
    Call sConectaBanco
    strSql = ""
    strSql = strSql & " SELECT ID,DESCRICAO,valor,tempo_est,vacina FROM TAB_servicos ORDER BY DESCRICAO"
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    If Rstemp.RecordCount <> 0 Then
        Rstemp.MoveLast
        Rstemp.MoveFirst
        'fmeListaPedidos.Visible = True
        
        For X = 1 To Rstemp.RecordCount
            With lstservicos
                .ListItems.Add X, , Rstemp!id
                
                If Not IsNull(Rstemp!Descricao) Then
                    .ListItems(X).SubItems(1) = Rstemp!Descricao
                Else
                    .ListItems(X).SubItems(1) = ""
                End If
                If Not IsNull(Rstemp!Valor) Then
                    .ListItems(X).SubItems(2) = Format(Rstemp!Valor, "###,##0.00")
                Else
                    .ListItems(X).SubItems(2) = "0.00"
                End If
                If Not IsNull(Rstemp!TEMPO_EST) Then
                    .ListItems(X).SubItems(3) = Format(Rstemp!TEMPO_EST, "000")
                Else
                    .ListItems(X).SubItems(3) = "00"
                End If
                
                If Not IsNull(Rstemp!VACINA) Then
                    .ListItems(X).SubItems(4) = IIf(Rstemp!VACINA = "S", "SIM", "NAO")
                Else
                    .ListItems(X).SubItems(4) = "NAO"
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




Private Sub chkVacina_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmd_Gravar.SetFocus
    End If
End Sub

Private Sub chkVacina_LostFocus()
   ' cmd_Gravar.SetFocus
End Sub

Private Sub cmd_Adicionar_Click()
    txtServico.Enabled = True
    txtDuracao.Enabled = True
    txtValor.Enabled = True
    txtServico.SetFocus
    txtServico.text = ""
    txtDuracao.text = ""
    txtValor.text = Format("0.00", "###,##0.00")
    chkVacina.Value = 0
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True
    cmd_Excluir.Enabled = False
    lstservicos.Enabled = False
    cmd_Voltar.Visible = True
    iTipoOperacao = 1

End Sub

Private Sub cmd_Excluir_Click()
    If Len(txtServico.text) = 0 Or txtServico.text = "" Then
       MsgBox "Descrição de serviço inválida. Favor corrigir", vbOKOnly
       txtServico.SetFocus
       Exit Sub
    End If
    
    If MsgBox("Tem certeza que deseja excluir o Serviço: " & Chr(13) & Chr(10) & _
                            Trim(lstservicos.SelectedItem.ListSubItems.Item(1)), vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        If fExcluir_Servico() Then
            cmd_Adicionar.Enabled = True
            cmd_Voltar.Enabled = False
            cmd_Gravar.Enabled = False
            cmd_Limpar.Enabled = False
            cmd_Voltar.Enabled = False
            cmd_Voltar.Visible = False
            lstservicos.ListItems.Clear
            Call Dados_Colunas
            If lstservicos.ListItems.Count > 0 Then
                lstservicos.ListItems(1).Selected = True
                txtServico.text = Trim(lstservicos.SelectedItem.ListSubItems.Item(1))
            End If
        Else
            MsgBox "Erro ao excluir o Serviço: " & Err.Description
        End If
    End If

End Sub

Private Sub cmd_Gravar_Click()
    
    If Len(txtServico.text) = 0 Or txtServico.text = "" Then
        MsgBox "Descrição de serviço inválida. Favor corrigir", vbOKOnly
        txtServico.SetFocus
        Exit Sub
    End If
    
    If Val(txtValor.text) = 0 Then
        If MsgBox("Campo Valor do serviço não está preenchido. " & Chr(13) & Chr(10) & "Deseja continuar e gravar assim mesmo? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            txtValor.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtDuracao.text) = 0 Then
        If MsgBox("Campo Tempo de duração do serviço não está preenchido. " & Chr(13) & Chr(10) & "Deseja continuar e gravar assim mesmo? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            txtDuracao.SetFocus
            Exit Sub
        End If
    End If
    
    If fGravar_Servico() Then
        cmd_Adicionar.Enabled = True
        cmd_Adicionar.Visible = True
        cmd_Voltar.Enabled = False
        cmd_Gravar.Enabled = False
        cmd_Limpar.Enabled = False
        cmd_Voltar.Enabled = False
        cmd_Voltar.Visible = False
        'cmd_Excluir.Enabled = true
        lstservicos.ListItems.Clear
        Call Dados_Colunas
        lstservicos.ListItems(1).Selected = True
        txtServico.text = Trim(lstservicos.SelectedItem.ListSubItems.Item(1))
        'Call cmd_Limpar_Click
    Else
        MsgBox "Erro ao incluir o serviço: " & Err.Description
    End If
    lstservicos.Enabled = True
End Sub

Private Sub cmd_Limpar_Click()
    txtServico.text = ""
    txtValor.text = "0.00"
    txtDuracao.text = ""
    'txtServico.SetFocus
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True
    
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
    Desabilita Me
    lstservicos.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Then
'      SendKeys "{TAB}"
'  End If
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
    mensagem = MsgBox("Informações não Salvas. Deseja Sair assim mesmo ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
    If mensagem = vbNo Then
        Exit Sub
    Else
        Unload Me
    End If
End If


End Sub

Private Sub Form_Load()
    Call Nomes_Colunas
    Call Dados_Colunas
    'lstServicos.ListItems = 1
    If lstservicos.ListItems.Count > 0 Then
        txtServico.text = Trim(lstservicos.SelectedItem.ListSubItems.Item(1))
        txtValor.text = Format(lstservicos.SelectedItem.ListSubItems.Item(2), "###,##0.00")
        txtDuracao.text = lstservicos.SelectedItem.ListSubItems.Item(3)
    End If
    cmd_Voltar.Top = cmd_Adicionar.Top
    cmd_Voltar.Left = cmd_Adicionar.Left
    cmd_Voltar.Visible = False
End Sub

Private Sub lstservicos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtServico.text = Trim(lstservicos.SelectedItem.ListSubItems.Item(1))
    txtValor.text = Format(lstservicos.SelectedItem.ListSubItems.Item(2), "###,##0.00")
    txtDuracao.text = lstservicos.SelectedItem.ListSubItems.Item(3)
    chkVacina.Value = IIf(lstservicos.SelectedItem.ListSubItems.Item(4) = "SIM", 1, 0)
    iTipoOperacao = 2
End Sub

Private Sub lstservicos_KeyPress(KeyAscii As Integer)
    If lstservicos.ListItems.Count > 0 Then
        Sendkeys "{tab}"
    End If
End Sub

Private Sub lstservicos_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'   If Button = 2 Then
'        lstservicos.SetFocus
'        'mnuEdicao.Visible = True
'        PopupMenu mnuEdicao, lstservicos.SelectedItem.Left + y, lstservicos.SelectedItem.Top + x
'    End If
    
    Set lvListItems_Itens = lstservicos.HitTest(X, y)

    'Check if a record was selected
    If lvListItems_Itens Is Nothing Then
        If lstservicos.ListItems.Count > 0 Then
            lstservicos.SelectedItem.Selected = False
        End If
        'se não estiver item selecionado desabilita menus
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
    txtServico.Enabled = True
    txtDuracao.Enabled = True
    txtValor.Enabled = True
    cmd_Gravar.Enabled = True
    cmd_Limpar.Enabled = True
    cmd_Adicionar.Enabled = False
    cmd_Adicionar.Visible = False
    cmd_Voltar.Enabled = True
    cmd_Voltar.Visible = True
    txtServico.SetFocus
    iTipoOperacao = 2
End Sub

Private Sub mnuExcluir_Click()
    Call cmd_Excluir_Click
End Sub

Private Sub txtDuracao_GotFocus()
    SelText txtDuracao
End Sub

Private Sub txtDuracao_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        chkVacina.SetFocus
    End If

    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    If KeyAscii = 44 And InStr(txtDuracao.text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If

End Sub

Private Function fGravar_Servico()
    
    If Len(txtServico.text) = 0 Or txtServico.text = "" Then
       MsgBox "Descrição do serviço inválida. Favor corrigir", vbOKOnly
       txtServico.SetFocus
       Exit Function
    End If
    
    fGravar_Servico = True
    
    On Error GoTo Erro_fGravar_Servico
    
    Call sConectaBanco
    
    'ID,DESCRICAO,valor,tempo_est
    If iTipoOperacao = 1 Then
        strSql = "INSERT INTO tab_servicos (DESCRICAO, VALOR, TEMPO_EST, VACINA, OPERADOR, DT_ATUALIZA)"
        strSql = strSql & " VALUES( '" & UCase(txtServico.text) & "',"
        strSql = strSql & Replace(txtValor.text, ",", ".") & "," & txtDuracao.text & ",'"
        strSql = strSql & IIf(chkVacina.Value = 0, "N", "S") & "','"
        strSql = strSql + NomeUsuario & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        
    Else
        strSql = "UPDATE tab_servicos SET DESCRICAO = '" & UCase(txtServico.text) & _
                                          "',VALOR =   " & Replace(txtValor.text, ",", ".") & _
                                          ",tempo_est = " & txtDuracao.text & _
                                          ",vacina = '" & IIf(chkVacina.Value = 0, "N", "S") & _
                                          "',OPERADOR = '" & NomeUsuario & _
                                          "', DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                          "' WHERE ID = " & lstservicos.SelectedItem.text
                                          
    End If
    Cnn.Execute strSql
    Cnn.Close
    Exit Function
    
Erro_fGravar_Servico:
    fGravar_Servico = False
End Function

Private Function fExcluir_Servico()
    
    fExcluir_Servico = True
    
    On Error GoTo Erro_fExcluir_Servico
    Call sConectaBanco
    strSql = ""
    strSql = strSql & " SELECT count(*) as contador FROM TAB_atendimentos "
    strSql = strSql & " WHERE tipo_atend = " & lstservicos.SelectedItem.text
    If Rstemp.State = adStateOpen Then
        Rstemp.Close
    End If
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    If Rstemp!Contador > 0 Then
         MsgBox "Exclusão não permitida para tipo de serviço com atendimento efetuado para ele", vbCritical, "Aviso"
     Else
        strSql = "DELETE from tab_servicos WHERE ID = '" & lstservicos.SelectedItem.text & "'"
        Cnn.Execute strSql
    End If
    
    Exit Function
    
Erro_fExcluir_Servico:
    fExcluir_Servico = False
End Function

Private Sub txtServico_GotFocus()
     Call SelText(txtServico)
End Sub

Private Sub txtServico_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        txtValor.SetFocus
    End If

End Sub

Private Sub txtValor_GotFocus()
    Call SelText(txtValor)
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
   
   'Char = Chr(KeyAscii)
   'KeyAscii = Asc(UCase(Char))
   If KeyAscii = 13 Then
       txtDuracao.SetFocus
       Exit Sub
   End If
 
 If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(txtValor.text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtValor_LostFocus()
    txtValor.text = Format(txtValor.text, "###,##0.00")
End Sub
