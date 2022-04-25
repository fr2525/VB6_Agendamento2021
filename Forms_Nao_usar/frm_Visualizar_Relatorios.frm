VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frm_Visualizar_Relatorios 
   ClientHeight    =   7605
   ClientLeft      =   1725
   ClientTop       =   4410
   ClientWidth     =   13395
   LinkTopic       =   "Form3"
   ScaleHeight     =   7605
   ScaleWidth      =   13395
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_Enviar_por_email 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      Picture         =   "frm_Visualizar_Relatorios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Enviar por email usando Outlook"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      Picture         =   "frm_Visualizar_Relatorios.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Abre em PDF"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton But1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Impressora"
      Height          =   375
      Left            =   11080
      TabIndex        =   1
      ToolTipText     =   "Escolher Impressora local ou Rede"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   10080
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   11565
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9045
      lastProp        =   500
      _cx             =   15954
      _cy             =   20399
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frm_Visualizar_Relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Appl As New CRAXDRT.Application
Public Report As New CRAXDRT.Report

Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Public CrxApp As New CRAXDRT.Application
Public CrxRpt As New CRAXDRT.Report
Public CrxSubRpt As New CRAXDRT.Report

Dim lJanelaChamadora As Form

Public Function Rel_Compras(ByVal CodSeq As Double)

On Error GoTo trataErroImpressao

Dim strCodPedido As String
Screen.MousePointer = 11

    strSql = "DELETE from relCompras"
    Cnn.Execute strSql
    
    strSql = "Select * from Itens_Compra where sequencia = " & CodSeq
    Set RsTemp1 = New ADODB.Recordset
    RsTemp1.Open strSql, Cnn, adOpenStatic, adLockReadOnly
    If RsTemp1.RecordCount > 0 Then
        RsTemp1.MoveLast
        RsTemp1.MoveFirst
        While Not RsTemp1.EOF
            strSql = ""
            strSql = "INSERT INTO relCompras VALUES ("
            strSql = strSql & CodSeq & ","
            strSql = strSql & RsTemp1(1) & ","
            strSql = strSql & "'" & RsTemp1(2) & "',"
            strSql = strSql & "'" & Troca_Virg_Zero(Format(RsTemp1(3), "0.000")) & "',"
            strSql = strSql & "'" & Troca_Virg_Zero(Format(RsTemp1(4), "0.00")) & "')"
            
            Cnn.Execute strSql
            RsTemp1.MoveNext
        Wend
    Else
        MsgBox "Pedido de Compras não Encontrado..!", vbInformation, "Aviso"
        Screen.MousePointer = 1
        Exit Function
    End If
        
    RsTemp1.Close
    If Not RsTemp1 Is Nothing Then
        Set RsTemp1 = Nothing
    End If
        
    'relatorio
    Set Report = Appl.OpenReport(App.Path & "\relCompras.Rpt")
    CRViewer91.ReportSource = Report
    Report.RecordSelectionFormula = "{COMPRA_PRODUTO.SEQUENCIA} = " & Format(CodSeq, "0000")
    
    'formulas do relatorio
    Dim i As Integer
    Dim nomeFormula As String
    For i = 1 To Report.FormulaFields.Count
        nomeFormula = UCase(Report.FormulaFields.Item(i).Name)
        If nomeFormula = "{@SEQUENCIA}" Then
            Report.FormulaFields.Item(i).Text = "'" + Format(CodSeq, "0000") + "'"
        ElseIf nomeFormula = "{@PEDIDO_ORCAMENTO}" Then
            'Report.FormulaFields.Item(i).Text = FMLNOME1
        End If
    Next i
    
    'atualiza o banco de dados
    Report.DiscardSavedData
    'abre janela impressora
    'Report.PrinterSetup (0)
    
    ' abre rewlatorio
    CRViewer91.ViewReport
    
    Do While CRViewer91.IsBusy      'ZOOM METHOD DOES NOT WORK WHILE
        DoEvents                    'REPORT IS LOADING, SO WE MUST PAUSE
    Loop
    
    Screen.MousePointer = 1
Exit Function

trataErroImpressao:
Screen.MousePointer = 1

MsgBox Err.Number & " Ocorreu um erro descrição : " & Err.Description, vbInformation, "Aviso"
Err.Clear
End Function

Public Function Rel_Comissao_Individual(ByVal data_inicio As String, ByVal data_fim As String, ByVal cod_vendedor As Double)
On Error GoTo trataErroImpressao

Set lJanelaChamadora = Screen.ActiveForm

If lJanelaChamadora.Name <> "FrmCompras" Then
    Command3.Visible = False
End If

    sql = "DELETE FROM REL_COMISSAO"
    Cnn.Execute sql
    
    sql = "SELECT * FROM SAIDAS_PRODUTO_FAT WHERE DATA_PREV_ENT IS NOT NULL "
    sql = sql & " AND DATA_PREV_ENT >= '" & Format(data_inicio, "mm/dd/yyyy") & "'"
    sql = sql & " AND DATA_PREV_ENT <= '" & Format(data_fim, "mm/dd/yyyy") & "'"
    sql = sql & " AND CODIGO_REPRESENTANTE = " & cod_vendedor
    '''sql = sql & " AND STATUS_NF = 'S'" 'nota faturada
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open sql, Cnn, adOpenStatic, adLockReadOnly
    If Rstemp.RecordCount > 0 Then
        Rstemp.MoveLast
        Rstemp.MoveFirst
        While Not Rstemp.EOF
            sql = "INSERT INTO REL_COMISSAO values( "
            sql = sql & Rstemp!CODIGO_REPRESENTANTE & ","
            sql = sql & Rstemp!SEQUENCIA & ","
            sql = sql & "'" & Format(Rstemp!Data_NF, "MM/DD/YYYY") & "',"
            sql = sql & "'" & Troca_Virg_Zero(Format(Rstemp!TOTAL_SAIDA, "0.00")) & "',"
            If Not IsNull(Rstemp!PERC_COMISSAO) Then
                sql = sql & "'" & Troca_Virg_Zero(Format(Rstemp!PERC_COMISSAO, "0.00")) & "',"
                sql = sql & "'" & Troca_Virg_Zero(Format(Rstemp!TOTAL_COMISSAO, "0.00")) & "')"
            Else
                sql = sql & "0,0)"
            End If
            Cnn.Execute sql
            Rstemp.MoveNext
        Wend
    End If
    
    Set Report = Appl.OpenReport(App.Path & "\relcomiss_ind.Rpt")
    CRViewer91.ReportSource = Report
   ' Report.RecordSelectionFormula = "{relcomiss_ind.@PERIODO} = " '" + "Período de: " & data_inicio & Space(15) & " Até: " & data_fim & "'"
    
    'formulas do relatorio
    Dim i As Integer
    Dim nomeFormula As String
    For i = 1 To Report.FormulaFields.Count
        nomeFormula = UCase(Report.FormulaFields.Item(i).Name)
        If nomeFormula = "{@PERIODO}" Then
            Report.FormulaFields.Item(i).Text = "'" + "Período de: " & data_inicio & Space(15) & " Até: " & data_fim & "'"
        End If
    Next i
    
    'atualiza o banco de dados
    Report.DiscardSavedData
   
    ' abre rewlatorio
    CRViewer91.ViewReport
    
    Do While CRViewer91.IsBusy      'ZOOM METHOD DOES NOT WORK WHILE
        DoEvents                    'REPORT IS LOADING, SO WE MUST PAUSE
    Loop
    
    Screen.MousePointer = 1
Exit Function

trataErroImpressao:
Screen.MousePointer = 1

MsgBox Err.Number & " Ocorreu um erro descrição : " & Err.Description, vbInformation, "Aviso"
Err.Clear
End Function


Public Function Rel_Comissao_Geral(ByVal data_inicio As String, ByVal data_fim As String)
On Error GoTo trataErroImpressao
    
    Set lJanelaChamadora = Screen.ActiveForm
    If lJanelaChamadora.Name <> "FrmCompras" Then
        Command3.Visible = False
    End If

    Set Report = Appl.OpenReport(App.Path & "\Relcomissao.Rpt")
    CRViewer91.ReportSource = Report
    'formulas do relatorio
    Dim i As Integer
    Dim nomeFormula As String
    For i = 1 To Report.FormulaFields.Count
        nomeFormula = UCase(Report.FormulaFields.Item(i).Name)
        If nomeFormula = "{@PERIODO}" Then
            Report.FormulaFields.Item(i).Text = "'" + "Período de: " & data_inicio & Space(15) & " Até: " & data_fim & "'"
        End If
    Next i
    
    'atualiza o banco de dados
    Report.DiscardSavedData
   
    ' abre rewlatorio
    CRViewer91.ViewReport
    
    Do While CRViewer91.IsBusy      'ZOOM METHOD DOES NOT WORK WHILE
        DoEvents                    'REPORT IS LOADING, SO WE MUST PAUSE
    Loop

    
Screen.MousePointer = 1
    
Exit Function

trataErroImpressao:
Screen.MousePointer = 1

MsgBox Err.Number & " Ocorreu um erro descrição : " & Err.Description, vbInformation, "Aviso"
Err.Clear
End Function

Public Function Rel_Lucros(ByVal data_inicio As String, ByVal data_fim As String)
On Error GoTo trataErroImpressao
    
    Set lJanelaChamadora = Screen.ActiveForm
    If lJanelaChamadora.Name <> "FrmCompras" Then
        Command3.Visible = False
    End If
    
    TOT_UNIT_VENDA = 0
    TOT_UNI_CUSTO = 0
    TOT_ITEM = 0
    TOT_GERAL_LIQ = 0
    TOT_COMISS = 0
    
    VlrSubTotalPedido = 0
    VlrFrete = 0
    VlrDesconto = 0
    
    Cnn.Execute "Delete from REL_LUCRO_PROD"
    
    sql = " SELECT SUM(TOTAL_SAIDA) AS TOT_VENDA, SUM(TOTAL_DESCONTO) AS TOT_DESC, SUM(TOTAL_COMISSAO) AS TOT_COMISSAO "
    sql = sql & " FROM SAIDAS_PRODUTO_FAT "
    sql = sql & " WHERE DATA_PREV_ENT >= " & "'" & Format(data_inicio, "mm/dd/yyyy") & "'"
    sql = sql & " AND DATA_PREV_ENT <= " & "'" & Format(data_fim, "mm/dd/yyyy") & "'"
    'sql = sql & " AND STATUS_NF = 'S'"
    sql = sql & " AND NATUREZA = 1"
    Set Rstemp2 = New ADODB.Recordset
    Rstemp2.Open sql, Cnn, 1, 2
    If Rstemp2.RecordCount > 0 Then
        If Not IsNull(Rstemp2!TOT_VENDA) Then
            TOT_GERAL_LIQ = Format(Rstemp2!TOT_VENDA, "###,##0.00")
        Else
            TOT_GERAL_LIQ = 0
        End If
        If Not IsNull(Rstemp2!TOT_DESC) Then
            TOT_DESCONTO = Format(Rstemp2!TOT_DESC, "###,##0.00")
        Else
            TOT_DESCONTO = 0
        End If
        If Not IsNull(Rstemp2!TOT_COMISSAO) Then
            TOT_COMISS = Format(Rstemp2!TOT_COMISSAO, "###,##0.00")
        Else
            TOT_COMISS = 0
        End If
        TOT_VENDA_BRUTA = CCur(TOT_GERAL_LIQ) + CCur(TOT_DESCONTO)
    End If
    
    Rstemp2.Close
    Set Rstemp2 = Nothing
    
    sql = "SELECT B.CODIGO_PRODUTO, SUM(B.QTDE) as TQtde, SUM(B.VALOR_TOTAL) AS TOT_ITEN_VENDA, SUM(B.VALOR_CUSTO_TOTAL) AS TOT_ITEN_CUSTO "
    sql = sql & " FROM ITENS_SAIDA_FAT B, SAIDAS_PRODUTO_FAT A "
    sql = sql & " WHERE A.DATA_PREV_ENT >= " & "'" & Format(data_inicio, "mm/dd/yyyy") & "'"
    sql = sql & " AND A.DATA_PREV_ENT <= " & "'" & Format(data_fim, "mm/dd/yyyy") & "'"
    'sql = sql & " AND STATUS_NF = 'S'"
    sql = sql & " AND NATUREZA = 1"
    sql = sql & " AND A.SEQUENCIA = B.SEQUENCIA "
    sql = sql & " GROUP BY B.CODIGO_PRODUTO "
    Set RsTemp1 = New ADODB.Recordset
    RsTemp1.Open sql, Cnn, 1, 2
    If RsTemp1.RecordCount > 0 Then
        RsTemp1.MoveLast
        RsTemp1.MoveFirst
        Screen.MousePointer = 11
        Contador = 1
        While Not RsTemp1.EOF
            PERC_LUCRO_ITEM = 0
            If RsTemp1!TOT_ITEN_CUSTO > 0 Then
                PERC_LUCRO_ITEM = CCur((RsTemp1!TOT_ITEN_VENDA - RsTemp1!TOT_ITEN_CUSTO) / RsTemp1!TOT_ITEN_CUSTO * 100)
            ElseIf RsTemp1!TOT_ITEN_CUSTO = 0 And RsTemp1!TOT_ITEN_VENDA = 0 Then
                PERC_LUCRO_ITEM = 0
            ElseIf RsTemp1!TOT_ITEN_CUSTO = 0 And RsTemp1!TOT_ITEN_VENDA > 0 Then
                PERC_LUCRO_ITEM = 100
            End If
            
            If Not IsNull(RsTemp1!TOT_ITEN_CUSTO) Then
                TOT_CUSTO = CCur(TOT_CUSTO + RsTemp1!TOT_ITEN_CUSTO)
            End If
            
            sql = "INSERT INTO REL_LUCRO_PROD VALUES ("
            sql = sql & RsTemp1!CODIGO_PRODUTO & ","
            sql = sql & "'" & Troca_Virg_Zero(CCur(RsTemp1!TQTDE)) & "',"
            sql = sql & "'" & Troca_Virg_Zero(CCur(RsTemp1!TOT_ITEN_CUSTO)) & "',"
            sql = sql & "'" & Troca_Virg_Zero(CCur(RsTemp1!TOT_ITEN_VENDA)) & "',"
            sql = sql & "'" & Troca_Virg_Zero(PERC_LUCRO_ITEM) & "')"
            Cnn.Execute sql
            RsTemp1.MoveNext
            Contador = Contador + 1
        Wend
    Else
        MsgBox "Nenhum registro foi encontrado..!", vbInformation, "Aviso"
        Screen.MousePointer = 1
        RsTemp1.Close
        Exit Function
    End If
        
    RsTemp1.Close
    Set RsTemp1 = Nothing
    
    TOTAL_DESCONTOS = CCur(TOT_COMISS) + CCur(TOT_CUSTO)
                
    'TOT_LUCRO = Format(TOT_VENDA_BRUTA - TOTAL_DESCONTOS, "###,##0.00")
    TOT_LUCRO = Format(TOT_VENDA_BRUTA - CCur(TOT_DESCONTO) - CCur(TOT_COMISS) - TOT_CUSTO, "###,##0.00")
    str_PORCENT_LUCRO = Format(Format(TOT_LUCRO / TOT_GERAL_LIQ) * 100, "0.000")

    Set Report = Appl.OpenReport(App.Path & "\REL_LC_PROD.Rpt")
    CRViewer91.ReportSource = Report
    
    'Report.RecordSelectionFormula = "{COMPRA_PRODUTO.SEQUENCIA} = " & Format(CodSeq, "0000")
    
    'formulas do relatorio
    Dim i As Integer
    Dim nomeFormula As String
    For i = 1 To Report.FormulaFields.Count
        nomeFormula = UCase(Report.FormulaFields.Item(i).Name)
        If nomeFormula = "{@PERI_DE}" Then
            Report.FormulaFields.Item(i).Text = "'" + data_inicio & "'"
        ElseIf nomeFormula = "{@PERI_ATE}" Then
            Report.FormulaFields.Item(i).Text = "'" + data_fim & "'"
        ElseIf nomeFormula = "{@TITULO}" Then
            Report.FormulaFields.Item(i).Text = "'Relatório Resumo de Lucros'"
        ElseIf nomeFormula = "{@VENDA_BRUTA}" Then
            Report.FormulaFields.Item(i).Text = "'" & Format(TOT_VENDA_BRUTA, "###,##0.00") & "'"
        ElseIf nomeFormula = "{@DESCONTO}" Then
            Report.FormulaFields.Item(i).Text = "'" & Format(TOT_DESCONTO, "###,##0.00") & "'"
        ElseIf nomeFormula = "{@VENDA_LIQ}" Then
            Report.FormulaFields.Item(i).Text = "'" & Format(TOT_VENDA_BRUTA - CCur(TOT_DESCONTO), "###,##0.00") & "'"
        ElseIf nomeFormula = "{@TOT_COMISS}" Then
            Report.FormulaFields.Item(i).Text = "'" & Format(TOT_COMISS, "###,##0.00") & "'"
        ElseIf nomeFormula = "{@L_LIQ}" Then
            Report.FormulaFields.Item(i).Text = "'" & Format(TOT_VENDA_BRUTA - CCur(TOT_DESCONTO) - TOT_COMISS - TOT_CUSTO, "###,##0.00") & "'"
        ElseIf nomeFormula = "{@PERC_LC_LIQ}" Then
            Report.FormulaFields.Item(i).Text = "'" & Format(str_PORCENT_LUCRO, "###,##0.00") & " %" & "'"
        ElseIf nomeFormula = "{@VLR_TOT_CUSTO}" Then
            Report.FormulaFields.Item(i).Text = "'" & Format(TOT_CUSTO, "###,##0.00") & "'"
        End If
    Next i
    
    'atualiza o banco de dados
    Report.DiscardSavedData
   
    ' abre rewlatorio
    CRViewer91.ViewReport
    
    Do While CRViewer91.IsBusy      'ZOOM METHOD DOES NOT WORK WHILE
        DoEvents                    'REPORT IS LOADING, SO WE MUST PAUSE
    Loop

    
Screen.MousePointer = 1
    
Exit Function

trataErroImpressao:
Screen.MousePointer = 1

MsgBox Err.Number & " Ocorreu um erro descrição : " & Err.Description, vbInformation, "Aviso"
Err.Clear
End Function



Public Function Rel_Pedido(ByVal CodSeq As Double)
On Error GoTo trataErroImpressao

Dim strCodPedido As String
Screen.MousePointer = 11

Set lJanelaChamadora = Screen.ActiveForm

If lJanelaChamadora.Name <> "FrmCompras" Then
    Command3.Visible = False
End If

    sql = "DELETE from relpedidos"
    Cnn.Execute sql
    
    sql = "Select * from ITENS_SAIDA_FAT where sequencia = " & CodSeq
    Set RsTemp1 = New ADODB.Recordset
    RsTemp1.Open sql, Cnn, adOpenStatic, adLockReadOnly
    If RsTemp1.RecordCount > 0 Then
        RsTemp1.MoveLast
        RsTemp1.MoveFirst
        While Not RsTemp1.EOF
            sql = "INSERT INTO relpedidos VALUES ("
            sql = sql & CodSeq & ","
            sql = sql & RsTemp1(1) & ","
            sql = sql & "'" & RsTemp1(2) & "',"
            sql = sql & "'" & Troca_Virg_Zero(Format(RsTemp1(3), "0.00")) & "',"
            sql = sql & "'" & Troca_Virg_Zero(Format(RsTemp1(4), "0.00")) & "')"
            Cnn.Execute sql
            RsTemp1.MoveNext
        Wend
    Else
        MsgBox "Pedido/Orçamento não Encontrado..!", vbInformation, "Aviso"
        Screen.MousePointer = 1
        Exit Function
    End If
        
    RsTemp1.Close
    If Not RsTemp1 Is Nothing Then
        Set RsTemp1 = Nothing
    End If

    sql = "SELECT * FROM FORMA_PGTO_PREV where SEQUENCIA = " & CodSeq
    sql = sql & " and TIPO_MOV = 'R'"
    Set RsTemp1 = New ADODB.Recordset
    RsTemp1.Open sql, Cnn, adOpenStatic, adLockReadOnly
    If RsTemp1.RecordCount <> 0 Then
        RsTemp1.MoveLast
        RsTemp1.MoveFirst
        valor1 = ""
        valor2 = ""
        valor3 = ""
        valor4 = ""
        valor5 = ""
        valor6 = ""
        data1 = ""
        data2 = ""
        data3 = ""
        data4 = ""
        data5 = ""
        data6 = ""
        Do While True
            valor1 = Format(RsTemp1!Valor, "0.00")
            data1 = Format(RsTemp1!DATA_VENCIMENTO, "dd/mm/yyyy")
            RsTemp1.MoveNext
            If RsTemp1.EOF Then
                Exit Do
            End If
            valor2 = Format(RsTemp1!Valor, "0.00")
            data2 = Format(RsTemp1!DATA_VENCIMENTO, "dd/mm/yyyy")
            RsTemp1.MoveNext
            If RsTemp1.EOF Then
                Exit Do
            End If
            valor3 = Format(RsTemp1!Valor, "0.00")
            data3 = Format(RsTemp1!DATA_VENCIMENTO, "dd/mm/yyyy")
            RsTemp1.MoveNext
            If RsTemp1.EOF Then
                Exit Do
            End If
            valor4 = Format(RsTemp1!Valor, "0.00")
            data4 = Format(RsTemp1!DATA_VENCIMENTO, "dd/mm/yyyy")
            RsTemp1.MoveNext
            If RsTemp1.EOF Then
                Exit Do
            End If
            valor5 = Format(RsTemp1!Valor, "0.00")
            data5 = Format(RsTemp1!DATA_VENCIMENTO, "dd/mm/yyyy")
            RsTemp1.MoveNext
            If RsTemp1.EOF Then
                Exit Do
            End If
            valor6 = Format(RsTemp1!Valor, "0.00")
            data6 = Format(RsTemp1!DATA_VENCIMENTO, "dd/mm/yyyy")
            RsTemp1.MoveNext
            If RsTemp1.EOF Then
                Exit Do
            End If
        Loop
    End If
    
    RsTemp1.Close
    If Not RsTemp1 Is Nothing Then
        Set RsTemp1 = Nothing
    End If


    
    '************************************************************
    
    'relatorio
    Set Report = Appl.OpenReport(App.Path & "\relpedido.Rpt")
    CRViewer91.ReportSource = Report
    'Report.Database.LogOnServer "dao", "", "banco.mdb", "", "senha"
    Report.RecordSelectionFormula = "{relpedidos.SEQUENCIA} = " & Format(CodSeq, "0000")
    'report.FormulaFields
    
    'formulas do relatorio
    Dim i As Integer
    Dim nomeFormula As String
    For i = 1 To Report.FormulaFields.Count
        nomeFormula = UCase(Report.FormulaFields.Item(i).Name)
        If nomeFormula = "{@SEQUENCIA}" Then
            Report.FormulaFields.Item(i).Text = "'" + Format(CodSeq, "0000") + "'"
        ElseIf nomeFormula = "{@DATA_1}" Then
            Report.FormulaFields.Item(i).Text = "'" + data1 + "'"
        ElseIf nomeFormula = "{@VALOR_1}" Then
            Report.FormulaFields.Item(i).Text = "'" + Format(valor1, "#,###,##0.00") + "'"
        ElseIf nomeFormula = "{@DATA_2}" Then
            Report.FormulaFields.Item(i).Text = "'" + data2 + "'"
        ElseIf nomeFormula = "{@VALOR_2}" Then
            Report.FormulaFields.Item(i).Text = "'" + Format(valor2, "#,###,##0.00") + "'"
        ElseIf nomeFormula = "{@DATA_3}" Then
            Report.FormulaFields.Item(i).Text = "'" + data3 + "'"
        ElseIf nomeFormula = "{@VALOR_3}" Then
            Report.FormulaFields.Item(i).Text = "'" + Format(valor3, "#,###,##0.00") + "'"
        ElseIf nomeFormula = "{@DATA_4}" Then
            Report.FormulaFields.Item(i).Text = "'" + data4 + "'"
        ElseIf nomeFormula = "{@VALOR_4}" Then
            Report.FormulaFields.Item(i).Text = "'" + Format(valor4, "#,###,##0.00") + "'"
        ElseIf nomeFormula = "{@DATA_5}" Then
            Report.FormulaFields.Item(i).Text = "'" + data5 + "'"
        ElseIf nomeFormula = "{@VALOR_5}" Then
            Report.FormulaFields.Item(i).Text = "'" + Format(valor5, "#,###,##0.00") + "'"
        ElseIf nomeFormula = "{@DATA_6}" Then
            Report.FormulaFields.Item(i).Text = "'" + data6 + "'"
        ElseIf nomeFormula = "{@VALOR_6}" Then
            Report.FormulaFields.Item(i).Text = "'" + Format(valor6, "#,###,##0.00") + "'"
        End If
    Next i


    
    'atualiza o banco de dados
    Report.DiscardSavedData
    CRViewer91.ViewReport
    
    Do While CRViewer91.IsBusy      'ZOOM METHOD DOES NOT WORK WHILE
        DoEvents                    'REPORT IS LOADING, SO WE MUST PAUSE
    Loop


    data1 = ""
    data2 = ""
    data3 = ""
    data4 = ""
    data5 = ""
    data6 = ""
    valor1 = ""
    valor2 = ""
    valor3 = ""
    valor4 = ""
    valor5 = ""
    valor6 = ""
    
    Screen.MousePointer = 1
    
Exit Function

trataErroImpressao:
Screen.MousePointer = 1

MsgBox Err.Number & " Ocorreu um erro descrição : " & Err.Description, vbInformation, "Aviso"
Err.Clear
End Function


Private Sub Relatorio_com_Procedure()
    Dim conexao As New ADODB.Connection
    Dim cmd As New ADODB.Command

    
    Dim CRXApplication As New CRAXDRT.Application
    Dim CRXReport As New CRAXDDRT.Report
    Dim CRXDatabase As CRAXDRT.Database
    
    Dim Pedido As ADODB.Parameter


    'conexao.ConnectionString = "Provider=SQLOLEDB.1;Password=;Persist Security Info=True;User ID=sa;Initial Catalog=TopFolha;Data Source=SYSCOMP2\SYSCOMP_2005"
    'conexao.ConnectionString = "Provider=SQLOLEDB.1;Password=;Persist Security Info=True;User ID=sb;Initial Catalog=arqdados;Data Source=core2 "
    conexao.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sb;Initial Catalog=arqdados;Data Source=core2"
    conexao.Open
        
    Set Pedido = cmd.CreateParameter("@COD_EMPRESA", adInteger, adParamInput)
    'Set par_funcionario = cmd.CreateParameter("@COD_FUNCIONARIO", adVarChar, adParamInput)
    
    Pedido = InputBox("Nº Pedido")
    'par_filial = InputBox("Escolha a filial")
    'par_funcionario = InputBox("Escolha o funcionario")
    
    cmd.Parameters.Append Pedido
    'cmd.Parameters.Append par_filial
    'cmd.Parameters.Append par_funcionario

    Set cmd = New ADODB.Command
        With cmd
            Set .ActiveConnection = conexao
            .CommandType = adCmdStoredProc
            .CommandText = "PR_SEL_Pedido" & COD_EMPRESA
            'Set rs = .Execute
        End With
        
    Set Rs = cmd.Execute()
    Do Until Rs.EOF
        Rs.MoveNext
    Loop

    Set CRXReport = CRXApplication.OpenReport(App.Path & "\REPORT4.rpt", 1)
    Set CRXDatabase = CRXReport.Database

    CRXDatabase.SetDataSource Rs, 3, 1

    CR.ReportSource = CRXReport
    CR.Refresh
    CR.ViewReport
    
    
    Set conexao = Nothing
    Set cmd = Nothing
    Set Rs = Nothing



End Sub

Private Sub But1_Click()
Unload frm_Visualizar_Relatorios
End Sub

Private Sub cmd_Enviar_por_email_Click()

On Error Resume Next
If Dir$("\\servidor2000\SisAdven\PDF\") = "" Then
    MkDir "\\servidor2000\SisAdven\PDF"
End If

On Error GoTo Trata_Erro
    'pdf
    Report.ExportOptions.FormatType = crEFTPortableDocFormat
    Report.ExportOptions.DestinationType = crEDTDiskFile
    'Report.ExportOptions.DiskFileName = App.Path & "\ped_compra.pdf"
    Report.ExportOptions.DiskFileName = "\\servidor2000\SisAdven\PDF\ped_compra.pdf"
    Report.ExportOptions.PDFExportAllPages = True
    Report.Export (False)
    
    Dim iexp As String
Dim a As String

iexp = Environ("WINDIR") & "\explorer.exe  "
'Shell iexp & App.Path & "\ped_compra.pdf", vbMaximizedFocus
Shell iexp & "\\servidor2000\SisAdven\PDF\ped_compra.pdf", vbMaximizedFocus

Exit Sub

Trata_Erro:

If Err.Number = -2147206452 Then
    'Kill ("\\servidor2000\SisAdven\HTML\lista_" & temp)
    'On Error GoTo erro_exclusao_pdf
    'Kill App.Path & "\ped_compra.pdf"
Else
    MsgBox "Ocorreu um erro: " & Err.Description & " nro:" & Err.Number, vbCritical, "Erro"
End If

'erro_exclusao_pdf:
End Sub

Public Function Rel_CFe(ByVal data_inicio As String, ByVal data_fim As String, ByVal Titulo As String, ByVal reportName As String)
On Error GoTo trataErroImpressao
    
    Set lJanelaChamadora = Screen.ActiveForm
    If lJanelaChamadora.Name <> "FrmCompras" Then
        Command3.Visible = False
    End If
    
   'reportName = "REPORT1.RPT"
    Set Report = Appl.OpenReport(App.Path & "\" & reportName)
    CRViewer91.ReportSource = Report
    
    'Report.RecordSelectionFormula = "{COMPRA_PRODUTO.SEQUENCIA} = " & Format(CodSeq, "0000")
    
    'formulas do relatorio
    Dim i As Integer
    Dim nomeFormula As String
    For i = 1 To Report.FormulaFields.Count
        nomeFormula = UCase(Report.FormulaFields.Item(i).Name)
        If nomeFormula = "{@PERI_DE}" Then
            Report.FormulaFields.Item(i).Text = "'" & data_inicio & "'"
        ElseIf nomeFormula = "{@PERI_ATE}" Then
            Report.FormulaFields.Item(i).Text = "'" & data_fim & "'"
        ElseIf nomeFormula = "{@TITULO}" Then
            Report.FormulaFields.Item(i).Text = "'" & Titulo & "'"
        End If
    Next i
    
    'atualiza o banco de dados
    Report.DiscardSavedData
   
    ' abre rewlatorio
    CRViewer91.ViewReport
    
    Do While CRViewer91.IsBusy      'ZOOM METHOD DOES NOT WORK WHILE
        DoEvents                    'REPORT IS LOADING, SO WE MUST PAUSE
    Loop

    
Screen.MousePointer = 1
    
Exit Function

trataErroImpressao:
Screen.MousePointer = 1

MsgBox Err.Number & " Ocorreu um erro descrição : " & Err.Description, vbInformation, "Aviso"
Err.Clear
End Function




Private Sub Command1_Click()
On Error Resume Next
Report.PrinterSetup Me.hWnd 'Abre Print Setup
End Sub

Private Sub Command2_Click()
    Report.PrintOut
End Sub

Private Sub Command3_Click()

On Error Resume Next
If Dir$("\\servidor2000\SisAdven\PDF\") = "" Then
    MkDir "\\servidor2000\SisAdven\PDF"
End If

On Error GoTo Trata_Erro
    'pdf
    Report.ExportOptions.FormatType = crEFTPortableDocFormat
    Report.ExportOptions.DestinationType = crEDTDiskFile
    'Report.ExportOptions.DiskFileName = App.Path & "\ped_compra.pdf"
    Report.ExportOptions.DiskFileName = "\\servidor2000\SisAdven\PDF\ped_compra.pdf"
    Report.ExportOptions.PDFExportAllPages = True
    Report.Export (False)
    
    Dim iexp As String
Dim a As String

iexp = Environ("WINDIR") & "\explorer.exe  "
'Shell iexp & App.Path & "\ped_compra.pdf", vbMaximizedFocus
Shell iexp & "\\servidor2000\SisAdven\PDF\ped_compra.pdf", vbMaximizedFocus

Exit Sub

Trata_Erro:

If Err.Number = -2147206452 Then
    'Kill ("\\servidor2000\SisAdven\HTML\lista_" & temp)
    'On Error GoTo erro_exclusao_pdf
    'Kill App.Path & "\ped_compra.pdf"
Else
    MsgBox "Ocorreu um erro: " & Err.Description & " nro:" & Err.Number, vbCritical, "Erro"
End If

'erro_exclusao_pdf:



End Sub



Private Sub Form_Load()


'If lJanelaChamadora.Name = "FrmCompras" Then
    'Call Rel_Compras
'End If

'Call teste
'Exit Sub
'
'Dim strSQL As String
'
'    strSQL = "DELETE from relCompras"
'    Cnn.Execute strSQL
'
'    CodSeq = "404"
'
'    strSQL = ""
'    strSQL = "DELETE from relpedidos"
'    Cnn.Execute strSQL
'
'    strSQL = "Select * from ITENS_SAIDA_FAT "
'    strSQL = strSQL & " where sequencia = " & CodSeq
'    Set RsTemp1 = New ADODB.Recordset
'    RsTemp1.Open strSQL, Cnn, adOpenStatic, adLockReadOnly
'    If RsTemp1.RecordCount > 0 Then
'        RsTemp1.MoveLast
'        RsTemp1.MoveFirst
'        While Not RsTemp1.EOF
'            strSQL = ""
'            strSQL = "INSERT INTO relpedidos VALUES ("
'            strSQL = strSQL & CodSeq & ","
'            strSQL = strSQL & RsTemp1(1) & ","
'            strSQL = strSQL & "'" & RsTemp1(2) & "',"
'            strSQL = strSQL & "'" & Troca_Virg_Zero(Format(RsTemp1(3), "0.00")) & "',"
'            strSQL = strSQL & "'" & Troca_Virg_Zero(Format(RsTemp1(4), "0.00")) & "')"
'
'            Cnn.Execute strSQL
'            RsTemp1.MoveNext
'        Wend
'    Else
'        MsgBox "Pedido/Orçamento não Encontrado..!", vbInformation, "Aviso"
'        Screen.MousePointer = 1
'        Exit Sub
'    End If
'
'    RsTemp1.Close
'    If Not RsTemp1 Is Nothing Then
'        Set RsTemp1 = Nothing
'    End If
'
'    strSQL = "SELECT * FROM FORMA_PGTO_PREV where SEQUENCIA = " & CodSeq
'    strSQL = strSQL & " and TIPO_MOV = 'R'"
'    Set RsTemp1 = New ADODB.Recordset
'    RsTemp1.Open strSQL, Cnn, adOpenStatic, adLockReadOnly
'    If RsTemp1.RecordCount <> 0 Then
'        RsTemp1.MoveLast
'        RsTemp1.MoveFirst
'        valor1 = ""
'        valor2 = ""
'        valor3 = ""
'        valor4 = ""
'        valor5 = ""
'        valor6 = ""
'        data1 = ""
'        data2 = ""
'        data3 = ""
'        data4 = ""
'        data5 = ""
'        data6 = ""
'        Do While True
'            valor1 = Format(RsTemp1!Valor, "0.00")
'            data1 = Format(RsTemp1!DATA_VENCIMENTO, "dd/mm/yyyy")
'            RsTemp1.MoveNext
'            If RsTemp1.EOF Then
'                Exit Do
'            End If
'            valor2 = Format(RsTemp1!Valor, "0.00")
'            data2 = Format(RsTemp1!DATA_VENCIMENTO, "dd/mm/yyyy")
'            RsTemp1.MoveNext
'            If RsTemp1.EOF Then
'                Exit Do
'            End If
'            valor3 = Format(RsTemp1!Valor, "0.00")
'            data3 = Format(RsTemp1!DATA_VENCIMENTO, "dd/mm/yyyy")
'            RsTemp1.MoveNext
'            If RsTemp1.EOF Then
'                Exit Do
'            End If
'            valor4 = Format(RsTemp1!Valor, "0.00")
'            data4 = Format(RsTemp1!DATA_VENCIMENTO, "dd/mm/yyyy")
'            RsTemp1.MoveNext
'            If RsTemp1.EOF Then
'                Exit Do
'            End If
'            valor5 = Format(RsTemp1!Valor, "0.00")
'            data5 = Format(RsTemp1!DATA_VENCIMENTO, "dd/mm/yyyy")
'            RsTemp1.MoveNext
'            If RsTemp1.EOF Then
'                Exit Do
'            End If
'            valor6 = Format(RsTemp1!Valor, "0.00")
'            data6 = Format(RsTemp1!DATA_VENCIMENTO, "dd/mm/yyyy")
'            RsTemp1.MoveNext
'            If RsTemp1.EOF Then
'                Exit Do
'            End If
'        Loop
'    End If
'
'    RsTemp1.Close
'    If Not RsTemp1 Is Nothing Then
'        Set RsTemp1 = Nothing
'    End If
'
'
'
'    '************************************************************
'
'    'relatorio
'    Set Report = Appl.OpenReport(App.Path & "\relpedido.Rpt")
'    CRViewer91.ReportSource = Report
'    'Report.Database.LogOnServer "dao", "", "banco.mdb", "", "senha"
'    Report.RecordSelectionFormula = "{relpedidos.SEQUENCIA} = " & Format(CodSeq, "0000")
'    'report.FormulaFields
'
'    'formulas do relatorio
'    Dim i As Integer
'    Dim nomeFormula As String
'    For i = 1 To Report.FormulaFields.Count
'        nomeFormula = UCase(Report.FormulaFields.Item(i).Name)
'        If nomeFormula = "{@SEQUENCIA}" Then
'            Report.FormulaFields.Item(i).Text = "'" + Format(CodSeq, "0000") + "'"
'        ElseIf nomeFormula = "{@DATA_1}" Then
'            Report.FormulaFields.Item(i).Text = "'" + data1 + "'"
'        ElseIf nomeFormula = "{@VALOR_1}" Then
'            Report.FormulaFields.Item(i).Text = "'" + valor1 + "'"
'        ElseIf nomeFormula = "{@DATA_2}" Then
'            Report.FormulaFields.Item(i).Text = "'" + data2 + "'"
'        ElseIf nomeFormula = "{@VALOR_2}" Then
'            Report.FormulaFields.Item(i).Text = "'" + Format(valor2, "0.00") + "'"
'        ElseIf nomeFormula = "{@DATA_3}" Then
'            Report.FormulaFields.Item(i).Text = "'" + data3 + "'"
'        ElseIf nomeFormula = "{@VALOR_3}" Then
'            Report.FormulaFields.Item(i).Text = "'" + Format(valor3, "0.00") + "'"
'        ElseIf nomeFormula = "{@DATA_4}" Then
'            Report.FormulaFields.Item(i).Text = "'" + data4 + "'"
'        ElseIf nomeFormula = "{@VALOR_4}" Then
'            Report.FormulaFields.Item(i).Text = "'" + Format(valor4, "0.00") + "'"
'        ElseIf nomeFormula = "{@DATA_5}" Then
'            Report.FormulaFields.Item(i).Text = "'" + data5 + "'"
'        ElseIf nomeFormula = "{@VALOR_5}" Then
'            Report.FormulaFields.Item(i).Text = "'" + Format(valor5, "0.00") + "'"
'        ElseIf nomeFormula = "{@DATA_6}" Then
'            Report.FormulaFields.Item(i).Text = "'" + data6 + "'"
'        ElseIf nomeFormula = "{@VALOR_6}" Then
'            Report.FormulaFields.Item(i).Text = "'" + Format(valor6, "0.00") + "'"
'        End If
'    Next i
'
'
'
'    'atualiza o banco de dados
'    Report.DiscardSavedData
'    'abre janela impressora
'    Report.PrinterSetup (0)
'    CRViewer91.ViewReport
'
'    Do While CRViewer91.IsBusy      'ZOOM METHOD DOES NOT WORK WHILE
'        DoEvents                    'REPORT IS LOADING, SO WE MUST PAUSE
'    Loop
'
'
'    data1 = ""
'    data2 = ""
'    data3 = ""
'    data4 = ""
'    data5 = ""
'    data6 = ""
'    valor1 = ""
'    valor2 = ""
'    valor3 = ""
'    valor4 = ""
'    valor5 = ""
'    valor6 = ""
End Sub

Private Sub Form_Resize()
    With CRViewer91
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Appl = Nothing
Set Report = Nothing
'Call FechaRecordsets
End Sub

