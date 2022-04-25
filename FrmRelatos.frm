VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmRelatos 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Relatórios em Excel"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Voltar 
      Caption         =   "Voltar"
      Height          =   855
      Left            =   6270
      Picture         =   "FrmRelatos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmd_Ok 
      Caption         =   "Ok"
      Height          =   855
      Left            =   5100
      Picture         =   "FrmRelatos.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame fraPets 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pet para fechamento:"
      Enabled         =   0   'False
      Height          =   2595
      Left            =   210
      TabIndex        =   10
      Top             =   5550
      Visible         =   0   'False
      Width           =   7065
      Begin MSComctlLib.ListView ListaPets 
         Height          =   2085
         Left            =   90
         TabIndex        =   11
         ToolTipText     =   "Duplo Clique para escolher o PET"
         Top             =   240
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   3678
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selecionado:"
      Height          =   705
      Left            =   300
      TabIndex        =   8
      Top             =   3420
      Width           =   6975
      Begin VB.Label lblRelSelec 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.CommandButton cmd_Sair 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Sair (Alt+S)"
      Height          =   855
      Left            =   6270
      Picture         =   "FrmRelatos.frx":0244
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   945
   End
   Begin VB.CommandButton cmd_Relatos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   5100
      Picture         =   "FrmRelatos.frx":033E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   945
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   2490
      TabIndex        =   1
      Top             =   4350
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   609
      _Version        =   393216
      Format          =   117964801
      CurrentDate     =   42736
   End
   Begin MSComctlLib.ListView lst_Relatos 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   750
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483647
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nome do Relatório"
         Object.Width           =   12171
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "NomeSubRotina"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   2490
      TabIndex        =   2
      Top             =   4830
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   609
      _Version        =   393216
      Format          =   117899265
      CurrentDate     =   42736
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Período Final :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   435
      Left            =   480
      TabIndex        =   7
      Top             =   4830
      Width           =   1905
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Período Inicial :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   435
      Left            =   360
      TabIndex        =   6
      Top             =   4350
      Width           =   1935
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Escolha abaixo o relatório desejado "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   525
      Left            =   960
      TabIndex        =   5
      Top             =   180
      Width           =   5445
   End
End
Attribute VB_Name = "FrmRelatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTotalRel As Double

Private Sub cmd_Ok_Click()
    Call sImprimeFechaMes
    fraPets.Visible = False
    fraPets.Enabled = False
    DTPicker1.Enabled = True
    DTPicker2.Enabled = True
    lst_Relatos.Visible = True
    'cmd_Relatos.Caption = "&Imprimir!"
    'cmd_Relatos.Picture = LoadPicture(App.Path & "\impressora.ico")
    'cmd_Sair.Caption = "&Sair"
    'cmd_Sair.Picture = LoadPicture(App.Path & "\close.bmp")
    lst_Relatos.Enabled = True
    lst_Relatos.SetFocus
    cmd_Ok.Visible = False
    cmd_voltar.Visible = False
    cmd_Relatos.Visible = True
    cmd_Sair.Visible = True

End Sub

Private Sub cmd_Relatos_Click()
'****************************************************************

    nTotalRel = 0
    
    MousePointer = 11
    Select Case UCase(Trim(lst_Relatos.SelectedItem.SubItems(1)))
        Case UCase("AtendPeriodo")  'AtendPeriodo
            Call sGeraAtendPeriodo
        Case UCase("FechaMesPet")
            Call sGeraFechaMesPet
        Case UCase("VacVencPeriodo")
            Call sGeraVacVencPeriodo
    End Select
'        Call sImprimeFechaMes
'        fraPets.Visible = False
'        fraPets.Enabled = False
'        DTPicker1.Enabled = True
'        DTPicker2.Enabled = True
'        lst_Relatos.Visible = True
'        lst_Relatos.Visible = True
'        lst_Relatos.Enabled = True
'        lst_Relatos.SetFocus
End Sub

Private Sub sGeraAtendPeriodo()
    Call sConectaBanco

    strSql = "SELECT A.dt_atend,a.idanimal,a.tipo_atend, a.HORA_SAIDA,A.OBSERVA, A.HORA_VACINA " & _
                " ,B.id,B.ID_CLI,B.NOME,B.TIPO_ANI,C.DESCRICAO AS TIPOPET, D.DESCRICAO AS SERVICO " & _
                ", D.VACINA, A.VALOR,A.VALOR_RECEBIDO, E.RAZAO_SOCIAL" & _
                " ,E.FONE1, E.FONE2" & _
                " FROM TAB_ATENDIMENTOS A , tab_pets B, tab_tipos_pets C," & _
                " TAB_SERVICOS D, tab_clientes E" & _
                " Where  a.hora_saida <> '  :  ' " & _
                " AND A.idanimal = b.id " & _
                " AND b.id_cli = e.codigo " & _
                " AND b.tipo_ani = c.id" & _
                " AND a.tipo_atend = d.id " & _
                " AND a.dt_atend between '" & Format(DTPicker1.Value, "YYYY/MM/DD 00:00:00") & _
                "' AND '" & Format(DTPicker2.Value, "YYYY/MM/DD 23:59:59") & "'"
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    If Rstemp.RecordCount = 0 Then
        MousePointer = 0
        MsgBox "Sem movimentação no período", vbOKOnly, "Atenção"
        Exit Sub
    End If
   
'  relFecha.Destination = 0
'  relFecha.ReportFileName = "rptfechaperiodo.rpt"
'  relFecha.Action = 1
'  With relFecha
'        .Reset
'        .WindowShowZoomCtl = True
'        '.WindowControlBox = False
'        .PageZoom (100)
'        .WindowShowExportBtn = True
'        .WindowShowPrintBtn = True
'        .WindowShowPrintSetupBtn = True
'        .WindowShowRefreshBtn = True
'        .WindowShowCloseBtn = True
'        .WindowShowGroupTree = False
'        .WindowState = crptMaximized
'        .PageZoom (100)
'        '.WindowTitle = Me.Caption
'        strSql = "SELECT TAB_ATENDIMENTOS.DT_ATEND,TAB_ATENDIMENTOS.VALOR,tab_pets.NOME" & _
'                    ",TAB_SERVICOS.DESCRICAO,tab_tipos_pets.DESCRICAO " & _
'                    " From TAB_ATENDIMENTOS, tab_pets, TAB_SERVICOS,tab_tipos_pets  " & _
'                    " Where TAB_ATENDIMENTOS.IDANIMAL = tab_pets.ID " & _
'                    " AND TAB_ATENDIMENTOS.TIPO_ATEND = TAB_SERVICOS.ID " & _
'                    " AND tab_pets.TIPO_ANI = tab_tipos_pets.ID " & _
'                    " AND tab_atendimentos.dt_atend >= '" & Format(DTPicker1.Value & " 00:00:00", "yyyy/mm/dd hh:mm:ss") & "'" & _
'                    " AND tab_atendimentos.dt_atend <= '" & Format(DTPicker2.Value & " 23:59:59", "yyyy/mm/dd hh:mm:ss") & "'"
'
'        .SQLQuery = strSql
'        '.SelectionFormula = "{representante.codigo} = " & cboInicial.ItemData(cboInicial.ListIndex)
'        '.ReportFileName = App.Path & "\relatendimentos.rpt"
'        .ReportFileName = App.Path & "\relatendimentos.rpt"
'        '.Connect = "DSN=cnn_firebird;UID=SYSDBA;PWD=masterkey"
'        '.SelectionFormula = "{INVENTARIO.estoque} < 0 "
'        '.WindowTitle = frmRelVenda.Caption
'        .Formulas(0) = "PERI_DE = '" & DTPicker1.Value & "'"
'        '.Formulas(0) = "PERIODO = 'Período : De " & DTPicker1.Value & Space(5) & "Ate " & DTPicker2.Value & "'"
'
'        .Formulas(1) = "PERI_ATE = '" & DTPicker2.Value & "'"
'        '.Formulas(2) = "subTITULO = 'Relatório Inventário Contagem de Estoque'"
'
'        '.RetrieveDataFiles
'        .Action = 1
'    End With

'  FrmCompras.CrRelcomp.Destination = 0 'Vídeo
'  CristalSelect = "{tab_Clientes.negativo} = True"
'  FrmCompras.CrRelcomp.SelectionFormula = CristalSelect
'  FrmCompras.CrRelcomp.Formulas(0) = "nomeloja = '" & gNome & "'"
'  FrmCompras.CrRelcomp.ReportFileName = gPathRel & "\rlisnegra.rpt"
'  FrmCompras.CrRelcomp.Action = 1
'
'''' Cria a componente da classe application
'''' inclui um novo arquivo e uma nova planilha
    Set oexcel = CreateObject("Excel.Application")
    oexcel.Workbooks.Add 'inclui o workbook
    Set objExlSht = oexcel.ActiveWorkbook.Sheets(1)

    oexcel.Columns(1).columnWidth = 20
    oexcel.Columns(2).columnWidth = 20
    oexcel.Columns(3).columnWidth = 50
    oexcel.Columns(4).columnWidth = 50
    oexcel.Columns(5).columnWidth = 30
    oexcel.Columns(6).columnWidth = 15
    X = 1

    oexcel.cells(X, 1).Font.Size = 12
    oexcel.cells(X, 1).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 1).Font.Bold = False
    oexcel.cells(X, 1).Value = "Emissão: " & Date
    oexcel.cells(X, 3).Font.Size = 20
    oexcel.cells(X, 3).Font.Color = RGB(180, 0, 0)
    oexcel.cells(X, 3).Font.Bold = True
    oexcel.range("C1:h1").Merge (True)
    oexcel.range("C1:h1").Value = "Atendimentos por período"
    X = X + 1
    oexcel.range("C1:h1").Merge (False)
    '
    oexcel.cells(X, 1).Font.Size = 12
    oexcel.cells(X, 1).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 1).Font.Bold = False
    oexcel.cells(X, 1).Value = "Dt.Atendimento"
    '
    oexcel.cells(X, 2).Font.Size = 12
    oexcel.cells(X, 2).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 2).Font.Bold = False
    oexcel.cells(X, 2).Value = "Nome do PET"
    '
    oexcel.cells(X, 3).Font.Size = 12
    oexcel.cells(X, 3).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 3).Font.Bold = False
    oexcel.cells(X, 3).Value = "Serviço"
    '
    oexcel.cells(X, 4).Font.Size = 12
    oexcel.cells(X, 4).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 4).Font.Bold = False
    oexcel.cells(X, 4).Value = "Proprietário"
    '
    oexcel.cells(X, 5).Font.Size = 12
    oexcel.cells(X, 5).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 5).Font.Bold = False
    oexcel.cells(X, 5).Value = "Telefone"
    '
    oexcel.cells(X, 6).Font.Size = 12
    oexcel.cells(X, 6).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 6).Font.Bold = False
    oexcel.cells(X, 6).Value = "Valor"
    X = X + 1
    Do While Not Rstemp.EOF
        oexcel.range("C1:h1").Merge (False)
        'oexcel.cells(x, 1).Font.Size = 12
        'oexcel.cells(x, 1).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 1).Font.Bold = False
        oexcel.cells(X, 1).Value = Format(Rstemp!DT_ATEND, "dd/mm/yyyy hh:mm:ss")
        '
        'oexcel.cells(x, 2).Font.Size = 12
        'oexcel.cells(x, 2).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 2).Font.Bold = False
        oexcel.cells(X, 2).Value = Rstemp!Nome
        '
        'oexcel.cells(x, 3).Font.Size = 12
        'oexcel.cells(x, 3).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 3).Font.Bold = False
        oexcel.cells(X, 3).Value = Rstemp!SERVICO
        '
        'oexcel.cells(x, 4).Font.Size = 12
        'oexcel.cells(x, 4).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 4).Font.Bold = False
        oexcel.cells(X, 4).Value = Rstemp!RAZAO_SOCIAL
        '
        'oexcel.cells(x, 5).Font.Size = 12
        'oexcel.cells(x, 5).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 5).Font.Bold = False
        oexcel.cells(X, 5).Value = Rstemp!FONE1 & " - " & Rstemp!FONE2
        '
        'oexcel.cells(x, 6).Font.Size = 12
        'oexcel.cells(x, 6).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 6).Font.Bold = False
        oexcel.cells(X, 6).HorizontalAlignment = 4
        oexcel.cells(X, 6).Value = Format(Rstemp!Valor, "###,##0.00")

        nTotalRel = nTotalRel + Rstemp!Valor
        Rstemp.MoveNext
        X = X + 1
        
    Loop
    X = X + 2
    oexcel.cells(X, 5).Font.Size = 14
    oexcel.cells(X, 5).Font.Bold = True
    oexcel.cells(X, 5).HorizontalAlignment = 4
    oexcel.cells(X, 5).Value = " TOTAL  "
    oexcel.cells(X, 6).Font.Size = 14
    oexcel.cells(X, 6).Font.Bold = True
    oexcel.cells(X, 6).HorizontalAlignment = 4
    oexcel.cells(X, 6).Value = Format(nTotalRel, "###,##0.00")
    
    On Error GoTo erro_sGeraAtendPeriodo

    objExlSht.SaveAs App.Path & "\AtendimentosPeriodo.xls"
    oexcel.Visible = True

    MousePointer = 0
    Exit Sub
    
erro_sGeraAtendPeriodo:
    MsgBox "Favor ver se a planilha do Excel dos atendimentos está aberta e feche-a", vbOKOnly, "Aviso"
    
'''
'''Rstemp.Close
'''Set Rstemp = Nothing

End Sub

Private Sub cmd_Sair_Click()
   Unload Me
End Sub

Private Sub sGeraFechaMesPet()
    lblTitulo.Caption = "Escolha o Pet para fechamento"
    fraPets.Top = lst_Relatos.Top
    fraPets.Left = lst_Relatos.Left
    fraPets.Height = lst_Relatos.Height
    fraPets.Width = lst_Relatos.Width
    fraPets.Visible = True
    fraPets.Enabled = True
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    lst_Relatos.Visible = False
    cmd_Relatos.Visible = False
    cmd_Sair.Visible = False
    cmd_Ok.Visible = True
    cmd_voltar.Visible = True
    Call Carrega_Colunas_Pets
    Call fCarrega_Pets
    ListaPets.SetFocus
    
End Sub

'********************************************************************
'*** Criar outra sub para esse processo abaixo

'*******************************************************************
Private Sub sImprimeFechaMes()
    
    Call sConectaBanco
    
    strSql = "SELECT A.dt_atend,a.idanimal,a.tipo_atend, a.HORA_SAIDA,A.OBSERVA, A.HORA_VACINA " & _
                " ,B.id,B.ID_CLI,B.NOME,B.TIPO_ANI,C.DESCRICAO AS TIPOPET, D.DESCRICAO AS SERVICO " & _
                ", D.VACINA, A.VALOR,A.VALOR_RECEBIDO, E.RAZAO_SOCIAL" & _
                " ,E.FONE1, E.FONE2" & _
                " FROM TAB_ATENDIMENTOS A , tab_pets B, tab_tipos_pets C," & _
                " TAB_SERVICOS D, tab_clientes E" & _
                " Where  a.hora_saida <> '  :  ' " & _
                " AND A.idanimal = b.id " & _
                " AND b.id_cli = e.codigo " & _
                " AND b.tipo_ani = c.id" & _
                " AND a.tipo_atend = d.id " & _
                " AND a.dt_atend between '" & Format(DTPicker1.Value, "YYYY/MM/DD 00:00:00") & _
                "' AND '" & Format(DTPicker2.Value, "YYYY/MM/DD 23:59:59") & "'" & _
                " AND a.idanimal = " & ListaPets.SelectedItem.text
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    If Rstemp.RecordCount = 0 Then
        MousePointer = 0
        MsgBox "Sem movimentação no período", vbOKOnly, "Atenção"
        Exit Sub
    End If
' Cria a componente da classe application
' inclui um novo arquivo e uma nova planilha
    Set oexcel = CreateObject("Excel.Application")
    oexcel.Workbooks.Add 'inclui o workbook
    Set objExlSht = oexcel.ActiveWorkbook.Sheets(1)
    
    oexcel.Columns(1).columnWidth = 20
    oexcel.Columns(2).columnWidth = 20
    oexcel.Columns(3).columnWidth = 50
    oexcel.Columns(4).columnWidth = 50
    oexcel.Columns(5).columnWidth = 30
    oexcel.Columns(6).columnWidth = 15
    X = 1
    
    oexcel.cells(X, 1).Font.Size = 12
    oexcel.cells(X, 1).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 1).Font.Bold = False
    oexcel.cells(X, 1).Value = "Emissão: " & Date
    oexcel.cells(X, 3).Font.Size = 20
    oexcel.cells(X, 3).Font.Color = RGB(180, 0, 0)
    oexcel.cells(X, 3).Font.Bold = True
    oexcel.range("C1:h1").Merge (True)
    oexcel.range("C1:h1").Value = "Fechamento do periodo do PET: " & Trim(ListaPets.SelectedItem.SubItems(1))
    X = X + 1
    oexcel.range("C1:h1").Merge (False)
    '
    oexcel.cells(X, 1).Font.Size = 12
    oexcel.cells(X, 1).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 1).Font.Bold = False
    oexcel.cells(X, 1).Value = "Dt.Atendimento"
    '
    oexcel.cells(X, 2).Font.Size = 12
    oexcel.cells(X, 2).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 2).Font.Bold = False
    oexcel.cells(X, 2).Value = "Nome do PET"
    '
    oexcel.cells(X, 3).Font.Size = 12
    oexcel.cells(X, 3).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 3).Font.Bold = False
    oexcel.cells(X, 3).Value = "Serviço"
    '
    oexcel.cells(X, 4).Font.Size = 12
    oexcel.cells(X, 4).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 4).Font.Bold = False
    oexcel.cells(X, 4).Value = "Proprietário"
    '
    oexcel.cells(X, 5).Font.Size = 12
    oexcel.cells(X, 5).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 5).Font.Bold = False
    oexcel.cells(X, 5).Value = "Telefone"
    '
    oexcel.cells(X, 6).Font.Size = 12
    oexcel.cells(X, 6).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 6).Font.Bold = False
    oexcel.cells(X, 6).Value = "Valor"
    X = X + 1
    Do While Not Rstemp.EOF
        oexcel.range("C1:h1").Merge (False)
        'oexcel.cells(x, 1).Font.Size = 12
        'oexcel.cells(x, 1).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 1).Font.Bold = False
        oexcel.cells(X, 1).Value = Format(Rstemp!DT_ATEND, "dd/mm/yyyy hh:mm:ss")
        '
        'oexcel.cells(x, 2).Font.Size = 12
        'oexcel.cells(x, 2).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 2).Font.Bold = False
        oexcel.cells(X, 2).Value = Rstemp!Nome
        '
        'oexcel.cells(x, 3).Font.Size = 12
        'oexcel.cells(x, 3).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 3).Font.Bold = False
        oexcel.cells(X, 3).Value = Rstemp!SERVICO
        '
        'oexcel.cells(x, 4).Font.Size = 12
        'oexcel.cells(x, 4).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 4).Font.Bold = False
        oexcel.cells(X, 4).Value = Rstemp!RAZAO_SOCIAL
        '
        'oexcel.cells(x, 5).Font.Size = 12
        'oexcel.cells(x, 5).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 5).Font.Bold = False
        oexcel.cells(X, 5).Value = Rstemp!FONE1 & " - " & Rstemp!FONE2
        '
        'oexcel.cells(x, 6).Font.Size = 12
        'oexcel.cells(x, 6).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 6).Font.Bold = False
        oexcel.cells(X, 6).HorizontalAlignment = 4
        oexcel.cells(X, 6).Value = Format(Rstemp!Valor, "###,##0.00")
        
        nTotalRel = nTotalRel + Rstemp!Valor
        Rstemp.MoveNext
        X = X + 1
        
    Loop
    X = X + 2
    oexcel.cells(X, 5).Font.Size = 14
    oexcel.cells(X, 5).Font.Bold = True
    oexcel.cells(X, 5).HorizontalAlignment = 4
    oexcel.cells(X, 5).Value = " TOTAL  "
    oexcel.cells(X, 6).Font.Size = 14
    oexcel.cells(X, 6).Font.Bold = True
    oexcel.cells(X, 6).HorizontalAlignment = 4
    oexcel.cells(X, 6).Value = Format(nTotalRel, "###,##0.00")

    On Error GoTo erro_sGeraFechaMes

    objExlSht.SaveAs App.Path & "\FechamentoPorPet.xls"
    oexcel.Visible = True

    MousePointer = 0

    Rstemp.Close
    Set Rstemp = Nothing
    Exit Sub

erro_sGeraFechaMes:
    MsgBox "Favor ver se a planilha do Excel do fechamento está aberta e feche-a", vbOKOnly, "Aviso"

End Sub

Private Sub sGeraVacVencPeriodo()

    Call sConectaBanco
    strSql = "select a.idanimal,a.dt_atend,a.descricao,a.valor, a.dt_proxima, b.nome " & _
            ",c.descricao as tipopet, d.razao_social as nomedono, d.fone1, d.fone2 " & _
            " from tab_vacinas a, tab_pets b, tab_tipos_pets c, tab_clientes d" & _
            " where a.dt_proxima between '" & Format(DTPicker1.Value, "YYYY/MM/DD 00:00:00") & _
            "' AND '" & Format(DTPicker2.Value, "YYYY/MM/DD 23:59:59") & "'" & _
            " AND b.id_cli = d.codigo "
            
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    If Rstemp.RecordCount = 0 Then
        MousePointer = 0
        MsgBox "Sem movimentação no período", vbOKOnly, "Atenção"
        Exit Sub
    End If
' Cria a componente da classe application
' inclui um novo arquivo e uma nova planilha
    Set oexcel = CreateObject("Excel.Application")
    oexcel.Workbooks.Add 'inclui o workbook
    Set objExlSht = oexcel.ActiveWorkbook.Sheets(1)
    
    oexcel.Columns(1).columnWidth = 20
    oexcel.Columns(2).columnWidth = 20
    oexcel.Columns(3).columnWidth = 50
    oexcel.Columns(4).columnWidth = 50
    oexcel.Columns(5).columnWidth = 30
    oexcel.Columns(6).columnWidth = 15
    X = 1
    
    oexcel.cells(X, 1).Font.Size = 12
    oexcel.cells(X, 1).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 1).Font.Bold = False
    oexcel.cells(X, 1).Value = "Emissão: " & Date
    oexcel.cells(X, 3).Font.Size = 20
    oexcel.cells(X, 3).Font.Color = RGB(180, 0, 0)
    oexcel.cells(X, 3).Font.Bold = True
    oexcel.range("C1:h1").Merge (True)
    oexcel.range("C1:h1").Value = "Vacinas que irão vencer no período"
    X = X + 1
    oexcel.range("C1:h1").Merge (False)
    '
    oexcel.cells(X, 1).Font.Size = 12
    oexcel.cells(X, 1).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 1).Font.Bold = False
    oexcel.cells(X, 1).Value = "Dt.Vencto"
    '
    oexcel.cells(X, 2).Font.Size = 12
    oexcel.cells(X, 2).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 2).Font.Bold = False
    oexcel.cells(X, 2).Value = "Nome do PET"
    '
    oexcel.cells(X, 3).Font.Size = 12
    oexcel.cells(X, 3).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 3).Font.Bold = False
    oexcel.cells(X, 3).Value = "Descrição da Vacina"
    '
    oexcel.cells(X, 4).Font.Size = 12
    oexcel.cells(X, 4).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 4).Font.Bold = False
    oexcel.cells(X, 4).Value = "Proprietário"
    '
    oexcel.cells(X, 5).Font.Size = 12
    oexcel.cells(X, 5).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 5).Font.Bold = False
    oexcel.cells(X, 5).Value = "Telefone"
    '
    oexcel.cells(X, 6).Font.Size = 12
    oexcel.cells(X, 6).Font.Color = RGB(0, 0, 255)
    oexcel.cells(X, 6).Font.Bold = False
    oexcel.cells(X, 6).Value = "Valor"
    X = X + 1
    Do While Not Rstemp.EOF
        oexcel.range("C1:h1").Merge (False)
        'oexcel.cells(x, 1).Font.Size = 12
        'oexcel.cells(x, 1).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 1).Font.Bold = False
        oexcel.cells(X, 1).Value = Format(Rstemp!DT_PROXIMA, "dd/mm/yyyy hh:mm:ss")
        '
        'oexcel.cells(x, 2).Font.Size = 12
        'oexcel.cells(x, 2).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 2).Font.Bold = False
        oexcel.cells(X, 2).Value = Rstemp!Nome
        '
        'oexcel.cells(x, 3).Font.Size = 12
        'oexcel.cells(x, 3).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 3).Font.Bold = False
        oexcel.cells(X, 3).Value = Rstemp!Descricao
        '
        'oexcel.cells(x, 4).Font.Size = 12
        'oexcel.cells(x, 4).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 4).Font.Bold = False
        oexcel.cells(X, 4).Value = Rstemp!nomedono
        '
        'oexcel.cells(x, 5).Font.Size = 12
        'oexcel.cells(x, 5).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 5).Font.Bold = False
        oexcel.cells(X, 5).Value = Rstemp!FONE1 & " - " & Rstemp!FONE2
        '
        'oexcel.cells(x, 6).Font.Size = 12
        'oexcel.cells(x, 6).Font.Color = RGB(0, 0, 255)
        'oexcel.cells(x, 6).Font.Bold = False
        'oexcel.cells(x, 6).HorizontalAlignment = 4
        'oexcel.cells(x, 6).Value = Format(Rstemp!valor, "###,##0.00")
        
        'nTotalRel = nTotalRel + Rstemp!valor
        Rstemp.MoveNext
        X = X + 1
        
    Loop
    'x = x + 2
    'oexcel.cells(x, 5).Font.Size = 14
    'oexcel.cells(x, 5).Font.Bold = True
    'oexcel.cells(x, 5).HorizontalAlignment = 4
    'oexcel.cells(x, 5).Value = " TOTAL  "
    'oexcel.cells(x, 6).Font.Size = 14
    'oexcel.cells(x, 6).Font.Bold = True
    'oexcel.cells(x, 6).HorizontalAlignment = 4
    'oexcel.cells(x, 6).Value = Format(nTotalRel, "###,##0.00")

    On Error GoTo erro_sGeraVacVenc
    
    objExlSht.SaveAs App.Path & "\VencimentoVacinas.xls"
    oexcel.Visible = True

    MousePointer = 0
    
    Rstemp.Close
    Set Rstemp = Nothing
    Exit Sub

erro_sGeraVacVenc:
    MsgBox "Favor ver se a planilha do Excel das vacinas está aberta e feche-a", vbOKOnly, "Aviso"
    
End Sub

Private Sub cmd_Voltar_Click()
    fraPets.Visible = False
    cmd_Relatos.Visible = True
    cmd_Sair.Visible = True
    cmd_Ok.Visible = False
    cmd_voltar.Visible = False
    lst_Relatos.Visible = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}"
    End If
    If KeyCode = vbKeyEscape Then
        mensagem = MsgBox("Sair mesmo ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
        If mensagem = vbYes Then
            Unload Me
            Exit Sub
        Else
            Exit Sub
        End If
    End If

'    If KeyCode = vbKeyF6 And cmd_Gravar.Enabled = True Then
'        mensagem = MsgBox("Salvar Dados ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
'        If mensagem = vbNo Then
'            Exit Sub
'        Else
'            cmd_Gravar_Click
'            Exit Sub
'        End If
'    ElseIf KeyCode = vbKeyF2 And cmd_Adicionar.Enabled = True Then
'        cmd_Adicionar_Click
'        Exit Sub
'    ElseIf KeyCode = vbKeyF4 And Cmd_limpar.Enabled = True Then
'        cmd_Limpar_Click
'        Exit Sub
'    ElseIf KeyCode = vbKeyF5 And cmd_Excluir.Enabled = True Then
'        cmd_Excluir_Click
'        Exit Sub
'    ElseIf KeyCode = vbKeyF7 And cmd_Sair.Enabled = True Then
'        cmd_Sair_Click
'        Exit Sub
'    ElseIf KeyCode = vbKeyEscape And cmd_Sair.Enabled = True Then
'        mensagem = MsgBox("Informações não Salvas. Deseja Sair assim mesmo ?", vbQuestion + vbYesNo + vbDefaultButton1, "Responda-me")
'        If mensagem = vbNo Then
'            Exit Sub
'        Else
'            Unload Me
'        End If
'    End If

End Sub

Private Sub Form_Load()
    
    nTipoRel = 0
    
    Me.Height = 5760
    fraPets.Top = lst_Relatos.Top
    fraPets.Left = lst_Relatos.Left
    fraPets.Height = lst_Relatos.Height
    fraPets.Width = lst_Relatos.Width
    fraPets.Visible = False
    cmd_Ok.Top = cmd_Relatos.Top
    cmd_Ok.Left = cmd_Relatos.Left
    cmd_voltar.Top = cmd_Sair.Top
    cmd_voltar.Left = cmd_Sair.Left
    
    Call sConectaBanco
    strSql = ""
    strSql = strSql & " SELECT NomeRelato, NomeSubRotina FROM TAB_relatorios ORDER BY NomeRelato "
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, Cnn, 1, 2
    If Rstemp.RecordCount <> 0 Then
        Rstemp.MoveLast
        Rstemp.MoveFirst
        'fmeListaPedidos.Visible = True
        
        For X = 1 To Rstemp.RecordCount
            lst_Relatos.ListItems.Add X, , Rstemp!NomeRelato
            
            If Not IsNull(Rstemp!NomeSubRotina) Then
                lst_Relatos.ListItems(X).SubItems(1) = Rstemp!NomeSubRotina
            Else
                lst_Relatos.ListItems(X).SubItems(1) = ""
            End If
            Rstemp.MoveNext
        Next
    Else
        'MsgBox "Tabela de modelos de relatórios sem registros", vbOKOnly
        On Error GoTo Erro_Cria_Tab_relatos
        
        strSql = ""
        strSql = strSql & "INSERT INTO tab_relatorios (nomerelato,nomesubrotina,operador,dt_atualiza) "
        strSql = strSql & " VALUES ('Fechamento Mensal por PET','FechaMesPet'"
        strSql = strSql & ",'" & sysNomeAcesso & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        Cnn.Execute strSql
        Cnn.CommitTrans
        '
        strSql = ""
        strSql = strSql & "INSERT INTO tab_relatorios (nomerelato,nomesubrotina,operador,dt_atualiza) "
        strSql = strSql & " VALUES ('Vacinas que vencerão no período','VacVencPeriodo'"
        strSql = strSql & ",'" & sysNomeAcesso & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = ""
        strSql = strSql & "INSERT INTO tab_relatorios (nomerelato,nomesubrotina,operador,dt_atualiza) "
        strSql = strSql & " VALUES ('Atendimentos do período','AtendPeriodo'"
        strSql = strSql & ",'" & sysNomeAcesso & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        'fmeListaPedidos.Visible = False
    End If
    
    Rstemp.Close
    Set Rstemp = Nothing
    DTPicker1.Value = DateAdd("d", -30, Date)
    DTPicker2.Value = Date
    Exit Sub
    
Erro_Cria_Tab_relatos:
    Call sMostraErro("frmRelatos", Err.Number, Err.Description)
    Call Fecha_Formularios
    End
    
End Sub
'
'    Dim myExcelFile As New clsExcelFile
'    Dim FileName$
'
'    With myExcelFile
'        'Create the new spreadsheet
'        FileName$ = App.Path & "\FechaMesPet.xls"
'        .CreateFile FileName$
'
'        'set a Password for the file. If set, the rest of the spreadsheet will
'        'be encrypted. If a password is used it must immediately follow the
'        'CreateFile method.
'        'This is different then protecting the spreadsheet (see below).
'        'NOTE: For some reason this function does not work. Excel will
'        'recognize that the file is password protected, but entering the password
'        'will not work. Also, the file is not encrypted. Therefore, do not use
'        'this function until I can figure out why it doesn't work. There is not
'        'much documentation on this function available.
'        '.SetFilePassword "PAUL"
'
'        'specify whether to print the gridlines or not
'        'this should come before the setting of fonts and margins
'        .PrintGridLines = False
'
'        'it is a good idea to set margins, fonts and column widths
'        'prior to writing any text/numerics to the spreadsheet. These
'        'should come before setting the fonts.
'
'        .SetMargin xlsTopMargin, 1.5   'set to 1.5 inches
'        .SetMargin xlsLeftMargin, 1.5
'        .SetMargin xlsRightMargin, 1.5
'        .SetMargin xlsBottomMargin, 1.5
'
'        'to insert a Horizontal Page Break you need to specify the row just
'        'after where you want the page break to occur. You can insert as many
'        'page breaks as you wish (in any order).
'        .InsertHorizPageBreak 10
'        .InsertHorizPageBreak 20
'
'        'set a default row height for the entire spreadsheet (1/20th of a point)
'        .SetDefaultRowHeight 14
'
'        'Up to 4 fonts can be specified for the spreadsheet. This is a
'        'limitation of the Excel 2.1 format. For each value written to the
'        'spreadsheet you can specify which font to use.
'
'        .setFont "Arial", 10, xlsNoFormat              'font0
'        .setFont "Arial", 10, xlsBold                  'font1
'        .setFont "Arial", 10, xlsBold + xlsUnderline   'font2
'        .setFont "Courier", 16, xlsBold + xlsItalic    'font3
'
'        'Column widths are specified in Excel as 1/256th of a character.
'        .SetColumnWidth 1, 5, 18
'
'        'Set special row heights for row 1 and 2
'        .SetRowHeight 1, 30
'        .SetRowHeight 2, 30
'
'        'set any header or footer that you want to print on
'        'every page. This text will be centered at the top and/or
'        'bottom of each page. The font will always be the font that
'        'is specified as font0, therefore you should only set the
'        'header/footer after specifying the fonts through SetFont.
'        .SetHeader "Fecha Mes por Pet"
'        .SetFooter "Novavia - Excel Class"
'
'        'write a normal left aligned string using font3 (Courier Italic)
'        .WriteValue xlsText, xlsFont3, xlsLeftAlign, xlsNormal, 1, 1, "Relatório de fechamnento do mes por Pet"
'        .WriteValue xlsText, xlsFont1, xlsLeftAlign, xlsNormal, 2, 1, "Novavia Automação"
'
'        'write some data to the spreadsheet
'        'Use the default format #3 "#,##0" (refer to the WriteDefaultFormats function)
'        'The WriteDefaultFormats function is compliments of Dieter Hauk in Germany.
'        .WriteValue xlsinteger, xlsFont0, xlsLeftAlign, xlsNormal, 6, 1, 2000, 3
'
'        'write a cell with a shaded number with a bottom border
'        .WriteValue xlsnumber, xlsFont1, xlsrightAlign + xlsBottomBorder + xlsShaded, xlsNormal, 7, 1, 12123.456, 4
'
'        'write a normal left aligned string using font2 (bold & underline)
'        .WriteValue xlsText, xlsFont2, xlsLeftAlign, xlsNormal, 8, 1, "This is a test string"
'
'        'write a locked cell. The cell will not be able to be overwritten, BUT you
'        'must set the sheet PROTECTION to on before it will take effect!!!
'        .WriteValue xlsText, xlsFont3, xlsLeftAlign, xlsLocked, 9, 1, "This cell is locked"
'
'        'fill the cell with "F"'s
'        .WriteValue xlsText, xlsFont0, xlsFillCell, xlsNormal, 10, 1, "F"
'
'        'write a hidden cell to the spreadsheet. This only works for cells
'        'that contain formula. Text, Number, Integer value text can not be hidden
'        'using this feature. It is included here for the sake of completeness.
'        .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsHidden, 11, 1, "If this were a formula it would be hidden!"
'
'        'write some dates to the file. NOTE: you need to write dates as xlsNumber
'        Dim d As Date
'        d = "15/01/2001"
'        .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, 15, 1, d, 12
'
'        d = "31/12/1999"
'        .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, 16, 1, d, 12
'
'        d = "01/04/2002"
'        .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, 17, 1, d, 12
'
'        d = "21/10/1998"
'        .WriteValue xlsnumber, xlsFont0, xlsCentreAlign, xlsNormal, 18, 1, d, 12
'
'        'PROTECT the spreadsheet so any cells specified as LOCKED will not be
'        'overwritten. Also, all cells with HIDDEN set will hide their formula.
'        'PROTECT does not use a password.
'        .ProtectSpreadsheet = False 'False | True
'
'        'Finally, close the spreadsheet
'        .CloseFile
'
'        MsgBox "Excel BIFF Spreadsheet created." & vbCrLf & "Filename: " & FileName$, vbInformation + vbOKOnly, "Excel Class"
'    End With
'
'    Exit Sub
'
'Erro_FechaMesPet:
'    Debug.Print "Número: " & Err.Number & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Linha: " & Erl & vbCrLf
'
'End Sub

Private Sub lst_Relatos_GotFocus()
   With lst_Relatos
       If .ListItems.Count > 0 Then
           For iY = 1 To .ListItems.Count
                .ListItems(iY).ForeColor = vbBlue
           Next
       End If
   End With
End Sub

Private Sub lst_Relatos_ItemClick(ByVal Item As MSComctlLib.ListItem)

    lblRelSelec.Caption = lst_Relatos.SelectedItem.text

End Sub

Private Sub Carrega_Colunas_Pets()
    With ListaPets
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "CodPet", 0, lvwColumnLeft
        .ColumnHeaders.Add 2, , "Nome do Pet", 1900, lvwColumnLeft
        .ColumnHeaders.Add 3, , "Tipo do Pet", 1460, lvwColumnLeft
        .ColumnHeaders.Add 4, , "Dono do Pet", 3900, lvwColumnLeft
        .ColumnHeaders.Add 5, , "Cuidados", 0, lvwColumnLeft
        .ColumnHeaders.Add 6, , "Id_Cli", 6, lvwColumnLeft
        .Height = 1990
    End With
End Sub

Private Function fCarrega_Pets(Optional nomepesq As String)
   
  fCarrega_Pets = True
  Call sConectaBanco
  If Rstemp.State = adStateOpen Then
      Rstemp.Close
   End If

   strSql = "SELECT a.id,a.id_cli,a.nome as nomepet,a.cuidados_especiais, b.razao_social as nomedono"
   strSql = strSql & ",A.tipo_ani,c.descricao as TIPO"
   strSql = strSql & " FROM tab_pets a, tab_clientes b, tab_tipos_pets c "
   strSql = strSql & " WHERE a.id_cli = b.codigo "
   strSql = strSql & " AND a.tipo_ani = c.id "
   If Len(nomepesq) > 0 Then
       strSql = strSql & " AND a.nome like '%" & nomepesq & "%'"
   End If
   strSql = strSql & " ORDER BY a.nome "
   
   Rstemp.Open strSql, Cnn, adOpenKeyset
   If Rstemp.BOF And Rstemp.EOF Then
       fCarrega_Pets = False
       Exit Function
   End If
   
   Carrega_List_Pets
   Rstemp.Close
   Cnn.Close
   
End Function

Private Sub Carrega_List_Pets()

 With Rstemp
    If .RecordCount <> 0 Then
      .MoveLast
      .MoveFirst
      ListaPets.ListItems.Clear
      For X = 1 To Rstemp.RecordCount
            If Not IsNull(Rstemp!id) Then
                ListaPets.ListItems.Add X, , Left(Rstemp!id, 5)
            Else
                ListaPets.ListItems.Add X, , ""
            End If

            ListaPets.ListItems(X).SubItems(1) = Rstemp!nomepet
            ListaPets.ListItems(X).SubItems(2) = Rstemp!Tipo
            ListaPets.ListItems(X).SubItems(3) = Rstemp!nomedono
            ListaPets.ListItems(X).SubItems(4) = Rstemp!cuidados_Especiais
            ListaPets.ListItems(X).SubItems(5) = Rstemp!id_cli
            Rstemp.MoveNext
        Next
    End If
  End With
  ListaPets.Height = 2000
End Sub


Private Sub lst_Relatos_LostFocus()
   Dim boAchou As Boolean
   With lst_Relatos
       boAchou = False
       If .ListItems.Count > 0 Then
           For iY = 1 To .ListItems.Count
               If .ListItems(iY).Selected = True Then
                   boAchou = True
                   iX = iY
               Else
                   '.ListItems(iY).ForeColor = vbBlack
               End If
           Next
       End If
   
       If boAchou = True Then
          .ListItems(iX).ForeColor = vbRed
       End If
   End With
    
End Sub
