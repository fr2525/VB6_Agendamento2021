Attribute VB_Name = "Funcoes"

Public Sub Desabilita(frm As Form)
'Deixa os textbox desabilitados
   Dim I
   
   For I = 0 To frm.Controls.Count - 1
       If TypeOf frm.Controls(I) Is TextBox Then
          frm.Controls(I).Enabled = False
       End If
       'If TypeOf frm.Controls(i) Is MaskEdBox Then
       '   frm.Controls(i).Enabled = False
       'End If
'       If TypeOf frm.Controls(i) Is MSFlexGrid Then
'          frm.Controls(i).Enabled = True
'       End If
       If TypeOf frm.Controls(I) Is ComboBox Then
          frm.Controls(I).Enabled = False
       End If
       If TypeOf frm.Controls(I) Is OptionButton Then
          frm.Controls(I).Enabled = False
       End If
   Next I
   
End Sub

Public Sub Habilita(frm As Form)
 Dim I
 For I = 0 To frm.Controls.Count - 1
    If TypeOf frm.Controls(I) Is TextBox Then
       frm.Controls(I).Enabled = True
    End If
    'If TypeOf frm.Controls(i) Is MaskEdBox Then
    '   frm.Controls(i).Enabled = True
    'End If
'    If TypeOf frm.Controls(i) Is MSFlexGrid Then
'          frm.Controls(i).Enabled = True
'       End If
     If TypeOf frm.Controls(I) Is ComboBox Then
          frm.Controls(I).Enabled = True
     End If
     If TypeOf frm.Controls(I) Is OptionButton Then
          frm.Controls(I).Enabled = True
     End If
Next I

End Sub

'--------------------------------------------------------------------------------------------------
Function ReadIniFile(ByVal sIniFileName As String, ByVal sSection As String, ByVal sItem As String, ByVal sDefault As String) As String
   Dim iRetAmount As Integer   'the amount of characters returned
   Dim sTemp As String

   sTemp = String$(400, 0)   'fill with nulls
   iRetAmount = GetPrivateProfileString(sSection, sItem, sDefault, sTemp, 400, sIniFileName)
   sTemp = Left$(sTemp, iRetAmount)
   ReadIniFile = sTemp
End Function
'
Function WriteIniFile(ByVal sIniFileName As String, ByVal sSection As String, ByVal sItem As String, ByVal sText As String) As Boolean
   Dim I As Integer
   On Error GoTo sWriteIniFileError

   I = WritePrivateProfileString(sSection, sItem, sText, sIniFileName)
   WriteIniFile = True

   Exit Function
sWriteIniFileError:
    WriteIniFile = False
End Function

Public Function DadosCBOtabela(cb As ComboBox, Tabela As String, campo As String, CodId As String) As Boolean
    Dim Selecao As ADODB.Recordset
    Call sConectaBanco
    strTemp = cb.Text
    cb.Clear
    Set Selecao = New ADODB.Recordset
    Selecao.Open "SELECT " & CodId & "," & campo & " FROM " & Tabela & " ORDER BY " & campo, Cnn, adOpenDynamic, adLockReadOnly
    If Selecao.EOF = True Then
        DadosCBOtabela = False
        Selecao.Close
        Exit Function
    End If
    Do While Not Selecao.EOF
        cb.AddItem IIf(IsNull(Selecao(campo)), "", Trim(Selecao(campo)))
        cb.ItemData(cb.NewIndex) = Selecao(CodId)
        Selecao.MoveNext
    Loop
    DadosCBOtabela = True
    Selecao.Close
    Cnn.Close
    cb.Text = strTemp
    strTemp = ""
    
End Function

Public Sub sCria_tabelas()

'*** Ve se existes as tabelas do sistema de Agendamento e se não existirem cria-as
'*
    
    On Error GoTo Erro_sCria_Tabelas
    
    Call sConectaBanco '---> Somente aqui vou usar os comandos de conexao porque se der erro tem que apagar os arquivos de segurança

   ' 'IF Firebird  -- Se é Firebird então:
   ' strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='tab_pets' "
   ' 'ELSE -- Senão se for SQL Server
   ' 'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'tab_pets'"
   ' 'END IF
   ' If Rstemp2.State = adStateOpen Then
   '     Rstemp2.Close
   ' End If
   ' Set Rstemp2 = New ADODB.Recordset
   ' Rstemp2.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
   '
   ' If Rstemp2(0) = 0 Then
       '***************************************************************************************
       '***************************   CRIA A TABELA DE ANIMAIS   ******************************
       '***************************************************************************************
       '
        strSql = "CREATE TABLE IF NOT EXISTS tab_pets (ID int AUTO_INCREMENT PRIMARY KEY NOT NULL" & _
                                            ", Id_Cli int NOT NULL" & _
                                            ", Nome  varchar(50) NOT NULL" & _
                                            ", Tipo_ani Int not null" & _
                                            ", dt_nasc date" & _
                                            ", pedigree CHAR(1)" & _
                                            ", observacoes varchar(200)" & _
                                            ", cuidados_especiais varchar(100)" & _
                                            ", foto varchar(100)" & _
                                            ", dt_ult_visita date" & _
                                            ", operador character(10)" & _
                                            ", dt_Atualiza timestamp )"
        Cnn.Execute strSql
        'Cnn.CommitTrans
            
'        'cria GENERATOR
'        strSql = "CREATE GENERATOR GEN_ANI_ID1 "
'        Cnn.Execute strSql
'        Cnn.CommitTrans
'
'        strSql = "SET GENERATOR GEN_ANI_ID1 TO 0"
'        Cnn.Execute strSql
'        Cnn.CommitTrans
'
'        strSql = " CREATE TRIGGER tab_pets_BI FOR tab_pets ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
'        strSql = strSql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_ANI_ID1, 1); END; "
'        Cnn.Execute strSql
'        Cnn.CommitTrans
'    End If
    
       '***************************************************************************************
       '***************************   CRIA A TABELA DE ANIMAIS   ******************************
       '***************************************************************************************
       '
        strSql = "CREATE TABLE  IF NOT EXISTS tab_clientes  (" & _
                "codigo  int AUTO_INCREMENT primary key, " & _
                "razao_social  varchar(60) NOT NULL, " & _
                "CEP_PRINCIPAL varchar(9) DEFAULT NULL," & _
                "endereco_principal  varchar(60) DEFAULT NULL," & _
                "nro_end_principal varchar(6) DEFAULT NULL," & _
                "bairro_end_principal  varchar(60) DEFAULT NULL," & _
                "cidade_end_principal  varchar(60) DEFAULT NULL," & _
                "UF_end_principal  varchar(2) DEFAULT NULL," & _
                "cgc_cpf  varchar(20) DEFAULT NULL," & _
                "INSC_ESTADUAL varchar(20) DEFAULT NULL," & _
                "INSC_municipAL varchar(20) DEFAULT NULL," & _
                "site varchar(60) DEFAULT NULL," & _
                "email  varchar(100) DEFAULT NULL," & _
                "fone1  varchar(15) DEFAULT NULL," & _
                "fone2  varchar(15) DEFAULT NULL," & _
                "contato  varchar(60) DEFAULT NULL," & _
                "obs  varchar(250) DEFAULT NULL," & _
                "cliente_desde  datetime DEFAULT NULL," & _
                "cod_representante int(8) DEFAULT NULL," & _
                "dias_atraso int(4) DEFAULT NULL," & _
                "bloqueado  tinyint(1) DEFAULT NULL," & _
                "limite  decimal(19,2) DEFAULT NULL," & _
                "operador  varchar(10) DEFAULT NULL," & _
                "data_atualiza  datetime DEFAULT NULL)"
                          
        Cnn.Execute strSql
                  
        '****************************************************************************************
        '****************   CRIA A TABELA DE TIPOS DE ANIMAL (CÃO/GATO/COELHO)  *****************
        '****************************************************************************************
        '
        strSql = "CREATE TABLE IF NOT EXISTS tab_tipos_pets (ID int AUTO_INCREMENT PRIMARY KEY NOT NULL" & _
                                        ", Descricao varchar(50) NOT NULL" & _
                                        ", operador character(10)" & _
                                        ", dt_Atualiza timestamp )"
        Cnn.Execute strSql
        'Cnn.CommitTrans
                
'***********************************************************************************
        '****************************************************************************************
        '*****************  CRIA A TABELA DE SERVICOS - BANHO/TOSA/VACINAS/ETC  *****************
        '****************************************************************************************
        '
        strSql = "CREATE TABLE IF NOT EXISTS tab_servicos (ID int AUTO_INCREMENT PRIMARY KEY NOT NULL" & _
                                        ", Descricao VARchar(50) NOT NULL" & _
                                        ", valor NUMERIC(12,2)" & _
                                        ", TEMPO_EST NUMERIC(12,2)" & _
                                        ", vacina CHAR(1)" & _
                                        ", operador character(10)" & _
                                        ", dt_Atualiza timestamp)"
        Cnn.Execute strSql
        'Cnn.CommitTrans
                
'***********************************************************************************
        '****************************************************************************************
        '********************    CRIA A TABELA DE ATENDIMENTOS    *******************************
        '****************************************************************************************
        '
        strSql = "CREATE TABLE IF NOT EXISTS tab_atendimentos (Dt_atend timestamp NOT NULL" & _
                                                ", IdAnimal int NOT NULL" & _
                                                ", Tipo_Atend int NOT NULL" & _
                                                ", valor NUMERIC(12,2)" & _
                                                ", valor_recebido NUMERIC(12,2)" & _
                                                ", forma_pagto NUMERIC(2) " & _
                                                ", hora_saida CHAR(5)" & _
                                                ", hora_vacina VARCHAR(5)" & _
                                                ", observa VARCHAR(150)" & _
                                                ", vacina  CHAR(1) " & _
                                                ", status  CHAR(1) " & _
                                                ", operador char(10)" & _
                                                ", dt_Atualiza timestamp" & _
                                                ", primary key (dt_atend) )"
        Cnn.Execute strSql
        'Cnn.CommitTrans
     'Não tem auto incremento porque o campo chave é TIMESTAMP

'***********************************************************************************
        strSql = "CREATE TABLE IF NOT EXISTS tab_vacinas (ID int AUTO_INCREMENT PRIMARY KEY NOT NULL " & _
                                        ",IdAnimal int NOT NULL " & _
                                        ",Dt_atend timestamp NOT NULL " & _
                                        ",Descricao VARCHAR(100) NOT NULL " & _
                                        ",valor NUMERIC(12,2) " & _
                                        ",DT_PROXIMA DATE " & _
                                        ",operador character(10) " & _
                                        ",dt_Atualiza timestamp " & _
                                        " )"
        Cnn.Execute strSql
        'Cnn.CommitTrans

'***********************************************************************************
        strSql = "CREATE TABLE IF NOT EXISTS tab_formas_pagto (ID int AUTO_INCREMENT PRIMARY KEY NOT NULL " & _
                                        ",Descricao VARCHAR(100) NOT NULL " & _
                                        ",DT_PROXIMA DATE " & _
                                        ",operador character(10) " & _
                                        ",dt_Atualiza timestamp " & _
                                        " )"
        Cnn.Execute strSql
        'Cnn.CommitTrans
                
'***********************************************************************************
        '****************************************************************************************
        '********************       CRIA A TABELA DE PROMOCOES      *******************************
        '****************************************************************************************
        '
        strSql = "CREATE TABLE IF NOT EXISTS tab_promocoes (ID int AUTO_INCREMENT PRIMARY KEY NOT NULL" & _
                                          ",Dt_inicio timestamp  NOT NULL" & _
                                          ",Dt_fim timestamp NOT NULL" & _
                                          ",IdAnimal int" & _
                                          ",IdTipoAten int" & _
                                          ",Descricao VARCHAR(100) NOT NULL" & _
                                          ",Valor NUMERIC(12,2)" & _
                                          ",porcent NUMERIC(2,2)" & _
                                          ",operador character(10)" & _
                                          ",Dt_Atualiza timestamp" & _
                                          "  )"

        Cnn.Execute strSql
        'Cnn.CommitTrans
        
    
'***********************************************************************************
        '****************************************************************************************
        '********************       CRIA A TABELA DE RELATORIOS   *******************************
        '****************************************************************************************
        '
        strSql = "CREATE TABLE IF NOT EXISTS tab_relatorios (NomeRelato VARCHAR(100)  NOT NULL" & _
                                          ",NomeSubRotina VARCHAR(40) NOT NULL" & _
                                          ",operador character(10)" & _
                                          ",Dt_Atualiza timestamp)"

        Cnn.Execute strSql
        Dim datahora
        datahora = Year(Now) & "/" & fuZeraEsq(Month(Now), 2) & "/" & fuZeraEsq(Day(Now), 2) & " " & fuZeraEsq(Hour(Now), 2) & ":" & fuZeraEsq(Minute(Now), 2) & ":" & fuZeraEsq(Second(Now), 2)
        
        strSql = "INSERT INTO tab_relatorios (NomeRelato" & _
                                          ",NomeSubRotina" & _
                                          ",operador" & _
                                          ",Dt_Atualiza)" & _
                                          " VALUES ('" & "Fechamento Mensal por PET" & _
                                          " ',' " & "FechaMesPet" & _
                                          "','" & sysNomeAcesso & _
                                          "','" & datahora & "')"
        Cnn.Execute strSql
        strSql = "INSERT INTO tab_relatorios (NomeRelato" & _
                                          ",NomeSubRotina" & _
                                          ",operador" & _
                                          ",Dt_Atualiza)" & _
                                          " VALUES ('" & "Vacinas que vencerão no período" & _
                                          " ',' " & "VacVencPeriodo" & _
                                          "','" & sysNomeAcesso & _
                                          "','" & datahora & "')"
        Cnn.Execute strSql
        strSql = "INSERT INTO tab_relatorios (NomeRelato" & _
                                          ",NomeSubRotina" & _
                                          ",operador" & _
                                          ",Dt_Atualiza)" & _
                                          " VALUES ('" & "Atendimentos do período" & _
                                          " ',' " & "AtendPeriodo" & _
                                          "','" & sysNomeAcesso & _
                                          "','" & datahora & "')"
        Cnn.Execute strSql
      
    
'**********************
    If Rstemp.State = adStateOpen Then
        Rstemp.Close
    End If
    If Rstemp2.State = adStateOpen Then
        Rstemp2.Close
    End If
    
'**********************

    If Cnn.State = adStateOpen Then
        Cnn.Close
    End If
         
    
    Exit Sub

Erro_sCria_Tabelas:
    Call sMostraErro("Módulo de Criação das tabelas", Err.Number, Err.Description)
    End
End Sub

Function MyMacroOperator(given$)
Select Case given
    Case "FechaMesPet": MyMacroOperator = FechaMesPet: Exit Function
    'Case "Test2": MyMacroOperator = Test2: Exit Function
    ' e assim com todas as variáveis
    Case Else ' Precisa conter alguma outra coisa.

    MyMacroOperator = 0 ' valor inválido
    End Select
End Function

Public Sub Fecha_Formularios()
    Dim Form As Form
    For Each Form In Forms
       Unload Form
       Set Form = Nothing
    Next Form
End Sub

Public Function FormatCPF_CNPJ(CPF_CNPJ As String)
Dim cont        As Integer

If Len(CPF_CNPJ) = 11 Then
    For cont = 0 To Len(CPF_CNPJ)
        If cont = 3 Or cont = 6 Then
            FormatCPF_CNPJ = FormatCPF_CNPJ & Mid(CPF_CNPJ, cont, 1) & "."
        ElseIf cont = 9 Then
            FormatCPF_CNPJ = FormatCPF_CNPJ & Mid(CPF_CNPJ, cont, 1) & "-"
        ElseIf cont <> 0 Then
            FormatCPF_CNPJ = FormatCPF_CNPJ & Mid(CPF_CNPJ, cont, 1)
        End If
    Next
ElseIf Len(CPF_CNPJ) = 14 Then
    For cont = 0 To Len(CPF_CNPJ)
        If cont = 2 Or cont = 5 Then
            FormatCPF_CNPJ = FormatCPF_CNPJ & Mid(CPF_CNPJ, cont, 1) & "."
        ElseIf cont = 8 Then
            FormatCPF_CNPJ = FormatCPF_CNPJ & Mid(CPF_CNPJ, cont, 1) & "/"
        ElseIf cont = 12 Then
            FormatCPF_CNPJ = FormatCPF_CNPJ & Mid(CPF_CNPJ, cont, 1) & "-"
        ElseIf cont <> 0 Then
            FormatCPF_CNPJ = FormatCPF_CNPJ & Mid(CPF_CNPJ, cont, 1)
        End If
    Next
End If

End Function


Public Function GetCampo(ByVal sql As String, ByVal campo As String) As String

Dim MyTmpRecordset As Recordset
    On Error GoTo Trata_Erro
    
    Set MyTmpRecordset = New ADODB.Recordset
    MyTmpRecordset.Open sql, Cnn, 1, 2
     
    GetCampo = MyTmpRecordset(0)
    
    MyTmpRecordset.Close
    MyTmpRecordset.ActiveConnection = Nothing
      
    Exit Function

Trata_Erro:
 
End Function

Public Function FiltraAspasSimples(STRDESCRICAO)
' SE STRING CONTIVER UMA ASPA SIMPLES INSERE MAIS UMA
' E O BANCO GRAVARÁ CORRETAMENTE
Dim intI As Integer
Dim intPos As Integer
Dim intClick As Integer

    intI = InStr(1, STRDESCRICAO, "'")
    If intI <> 0 Then
        FiltraAspasSimples = FiltraAspasSimples & Mid(STRDESCRICAO, 1, intI) & "'"
    End If
    
    Do While intI <> 0
        intClick = intClick + 1
        intPos = intI
        intI = InStr(intI + 1, STRDESCRICAO, intPos + 1)
        If intI = 0 Then
            FiltraAspasSimples = FiltraAspasSimples & Mid(STRDESCRICAO, intPos + 1)
        Else
            FiltraAspasSimples = FiltraAspasSimples & Mid(STRDESCRICAO, intPos + 1, intI - intPos) & "'"
        End If
    Loop
    
    If intClick = 0 Then
        FiltraAspasSimples = STRDESCRICAO
    End If
      
End Function

Public Function SemFormatoCPF_CNPJ(CPF_CNPJ As String)
Dim cont        As Integer

cont = 1

For cont = 1 To Len(CPF_CNPJ)
    If InStr("0123456789", Mid(CPF_CNPJ, cont, 1)) > 0 Then
        SemFormatoCPF_CNPJ = SemFormatoCPF_CNPJ & Mid(CPF_CNPJ, cont, 1)
    End If
Next

End Function

Public Function SemFormatoTel(tel As String)
    Dim cont As Integer
    
    cont = 1
        
    For cont = 1 To Len(tel)
        If InStr("0123456789", Mid(tel, cont, 1)) > 0 Then
            SemFormatoTel = SemFormatoTel & Mid(tel, cont, 1)
        End If
    Next

End Function

Public Function Select_Max(Tabela As String, campo As String) As Double
    Dim MyTmpRecordset As Recordset
    On Error GoTo Trata_Erro
    
    'Monta um comando sql para selecionar o valor maximo...
    sql = "SELECT MAX(" & campo & ") AS MAXIMO"
    sql = sql & " FROM " & "" & Tabela
    Set MyTmpRecordset = New ADODB.Recordset
    MyTmpRecordset.Open sql, Cnn, 1, 2
     
    Select_Max = MyTmpRecordset(0) + 1
    
    MyTmpRecordset.Close
    MyTmpRecordset.ActiveConnection = Nothing
  
    Exit Function

Trata_Erro:
    Select_Max = 1
End Function

Public Sub ErrosGeraisLog(sData As String, sFormNome As String, _
                          sRotina As String, sErro As String, _
                          intErroNumero As Integer)

    On Local Error Resume Next

    Dim FileFree As Integer

    FileFree = FreeFile
    Open App.Path & "\ErrosGeraisLog.Txt" For Append As #FileFree
    Print #FileFree, sData, sFormNome, sRotina, sErro, intErroNumero
    Close #FileFree

End Sub

Public Sub Erro(sTexto As String)
   sTexto = sTexto
   MsgBox "Ocorreu um Erro ao : " & sTexto & vbNewLine & "Caso Persista o Erro, Entre em Contato." & vbNewLine & vbNewLine & "Descrição do Erro: " & _
      "(" & Err.Description & ")", vbCritical, "Erro Nro.:" & Err
    Err.Clear
    If Not Rstemp Is Nothing Then
        Set Rstemp = Nothing
    End If
    If Not RsTemp1 Is Nothing Then
        Set RsTemp1 = Nothing
    End If
    
    Screen.MousePointer = 1
End Sub

Public Function FormatTEL(tel As String)
Dim NroTel      As String
Dim cont        As Integer

For cont = 1 To Len(tel)
    If Mid(tel, cont, 1) <> "-" Then
        NroTel = NroTel & Mid(tel, cont, 1)
    End If
Next
tel = NroTel

If tel <> "" And Len(tel) >= 7 Then
    FormatTEL = Left(tel, Len(tel) - 4) & "-" & Right(tel, 4)
End If

End Function
'
'RECEBE UMA STRING E A DEVOLVE SOMENTE COM OS SEUS CARACTERES ALFANUMÉRICOS
'
'EX: ? fuLimpaTexto("A B%CD*()_E F-2,45-5.78")
'      ABCDEF245578
'
Function fuLimpaTexto(ByVal vlTexto As String) As String

    Dim vlCont, vlChar As Integer
    Dim vlNewText As String
    
    vlTexto = VBA.Trim(vlTexto)
    vlNewText = ""
      For vlCont = 1 To Len(vlTexto)
          vlChar = Asc(VBA.Mid(vlTexto, vlCont, 1))
            If (vlChar > 47 And vlChar < 58) Or (vlChar > 64 And vlChar < 91) Then
               vlNewText = vlNewText & VBA.Chr(vlChar)
            End If
      Next vlCont
    fuLimpaTexto = vlNewText

End Function

Function fuZeraEsq(ByVal vlCampo As String, vlTam As Integer) As String

    vlCampo = fuLimpaTexto(vlCampo)
    If vlTam <= Len(vlCampo) Then
        vlCampo = Left(vlCampo, vlTam)
    Else
        vlCampo = VBA.String(vlTam - Len(vlCampo), "0") & vlCampo
    End If
    fuZeraEsq = vlCampo

End Function
'***Fabio Reinert - 08/2017 - função baseada na FORMATTEL mas para celular - Inicio
Public Function FormatData(data As String, dia As String, mes As String, ano As String)
    If data = "dd/mm/yyyy" Then
    End If
    If data = "dd/mm/yyyy" Then
    End If
    
    If data = "yyyy/mm/dd" Then
    End If
    If data = "yyyy/mm/dd" Then
    End If
    
End Function
'*** Fabio Reinert - 08/2017 - função baseada na FORMATTEL mas para celular - Inicio
Public Function FormatCEL(tel As String)
    Dim NroTel      As String
    Dim cont        As Integer
    
    For cont = 1 To Len(tel)
        If Mid(tel, cont, 1) <> "-" Then
            NroTel = NroTel & Mid(tel, cont, 1)
        End If
    Next
    tel = NroTel
    
    If tel <> "" And Len(tel) >= 8 Then
        FormatCEL = Left(tel, Len(tel) - 4) & "-" & Right(tel, 4)
    End If
End Function
'*** Fabio Reinert - 08/2017 - função baseada na FORMATTEL mas para celular - Fim

Public Sub sConectaBanco()
         
  On Error GoTo Erro_sConectaBanco
  
  If Cnn.State = adStateOpen Then
    Cnn.Close
  End If
  Set Cnn = New ADODB.Connection
  
  Cnn.Open "FILEDSN=" & App.Path & "\dbPet.dsn;UID=admin;PWD=oyster;"
  'ConDb.Open "FILEDSN=C:\ARQUIVOS DE PROGRAMAS\SMG\dbsmg.dsn;UID=admin;PWD=oyster;"
  gOperador = "Master"
  gnCodOperador = 99
    
'  If Cnn.State = adStateOpen Then
'    Cnn.Close
'  End If
'  Set Cnn = New ADODB.Connection
'  With Cnn
'     .CursorLocation = adUseClient
'     '.Open "File Name=" & App.Path & "\cnn_fire_servidor.udl;"
'     .Open "Driver=MySQL ODBC 8.0 Unicode Driver;Server=localhost;uid=root;pwd=612950;Database=DB_Pet  "
'   End With
Exit Sub

Erro_sConectaBanco:
    Call sMostraErro("sConectaBanco", Err.Number, Err.Description)
    'Call Fecha_Formularios
    End

End Sub





