VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmReceberPet 
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   ScaleHeight     =   4320
   ScaleWidth      =   3825
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraRecebido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recebimento"
      ClipControls    =   0   'False
      ForeColor       =   &H00008000&
      Height          =   4395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3885
      Begin VB.TextBox txtRecebido 
         Alignment       =   1  'Right Justify
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
         Height          =   390
         Left            =   2190
         MaxLength       =   20
         TabIndex        =   1
         Top             =   180
         Width           =   1410
      End
      Begin VB.CommandButton cmd_Receber 
         Caption         =   "Receber"
         Height          =   375
         Left            =   1620
         TabIndex        =   3
         ToolTipText     =   "Receber o valor do serviço"
         Top             =   3810
         Width           =   975
      End
      Begin VB.CommandButton cmd_retornar 
         Caption         =   "Voltar"
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         ToolTipText     =   "Retorna a agenda sem fazer o recebimento"
         Top             =   3810
         Width           =   975
      End
      Begin MSComctlLib.ListView LIST_DETALHESPGTO 
         Height          =   2925
         Left            =   180
         TabIndex        =   2
         Top             =   780
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   5159
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Recebido :"
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
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmReceberPet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_Receber_Click()
    If Val(txtRecebido.Text) = 0 Then
        If MsgBox("Valor está em branco, confirma? ", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me") = vbNo Then
            Exit Sub
        End If
    End If
    If LIST_DETALHESPGTO.SelectedItem = 0 Then
        MsgBox "Favor escolher uma forma de pagamento", vbOKOnly, "Aviso"
        LIST_DETALHESPGTO.SetFocus
        Exit Sub
    End If
    frmAgenda.nValorRecebido = CCur(txtRecebido.Text)
    frmAgenda.sFormaPagto = LIST_DETALHESPGTO.SelectedItem.SubItems(1)
    frmAgenda.bRecebido = True
    Unload Me
End Sub

Private Sub cmd_retornar_Click()
    frmAgenda.bRecebido = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then
        Msg = "Deseja Salvar Dados ...?"
        resposta = MsgBox(Msg, vbQuestion + vbYesNoCancel + vbDefaultButton1, "Inclusão de Dados")
        If resposta = vbNo Then  'nao
            frmAgenda.bRecebido = False
        ElseIf resposta = vbYes Then 'sim
            frmAgenda.nValorRecebido = CCur(txtRecebido.Text)
            frmAgenda.sFormaPagto = LIST_DETALHESPGTO.SelectedItem
            frmAgenda.bRecebido = True
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ' Se vc der um ENTER
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Call sCarregaCabPgto
    Call sCarregaListaPagto
End Sub
'
'*****************************************************
'
Private Sub sCarregaCabPgto()
    With LIST_DETALHESPGTO
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add , "C1", "Código", 800, lvwColumnLeft
        .ColumnHeaders.Add , "C2", "Descrição", 2600, lvwColumnLeft
      '  .ColumnHeaders.Add , "C3", "Total", 1350, lvwColumnRight
    End With

End Sub
'*
'**************************************************************
'
Private Sub sCarregaListaPagto()
Dim cont As Integer
cont = 0

LIST_DETALHESPGTO.ListItems.Clear
    
    Call sConectaBanco
    
    sql = "Select * from tab_formas_pagto "
    Set Rstemp3 = New ADODB.Recordset
    Rstemp3.Open sql, Cnn, 1, 2
    If Rstemp3.RecordCount > 0 Then
        cont = 1
        While Not Rstemp3.EOF
            If Not IsNull(Rstemp3!id) Then
                LIST_DETALHESPGTO.ListItems.Add (cont), , Rstemp3!id
            Else
                LIST_DETALHESPGTO.ListItems.Add (cont), , ""
            End If
            If Not IsNull(Rstemp3!Descricao) Then
                LIST_DETALHESPGTO.ListItems(cont).SubItems(1) = UCase(Rstemp3!Descricao)
            Else
                LIST_DETALHESPGTO.ListItems(cont).SubItems(1) = "Descrição não Encontrada"
            End If
            cont = cont + 1
            Rstemp3.MoveNext
        Wend
    End If
    
    Rstemp3.Close
    Set Rstemp3 = Nothing
    
End Sub

Private Sub txtRecebido_GotFocus()
    SelText txtRecebido
End Sub
