Attribute VB_Name = "Funcoes"

Public Sub Desabilita(frm As Form)
'Deixa os textbox desabilitados
   Dim i
   
   For i = 0 To frm.Controls.Count - 1
       If TypeOf frm.Controls(i) Is TextBox Then
          frm.Controls(i).Enabled = False
       End If
       'If TypeOf frm.Controls(i) Is MaskEdBox Then
       '   frm.Controls(i).Enabled = False
       'End If
'       If TypeOf frm.Controls(i) Is MSFlexGrid Then
'          frm.Controls(i).Enabled = True
'       End If
       If TypeOf frm.Controls(i) Is ComboBox Then
          frm.Controls(i).Enabled = False
       End If
'       If TypeOf frm.Controls(i) Is OptionButton Then
'          frm.Controls(i).Enabled = False
'       End If
   Next i
   
End Sub

Public Sub Habilita(frm As Form)
 Dim i
 For i = 0 To frm.Controls.Count - 1
    If TypeOf frm.Controls(i) Is TextBox Then
       frm.Controls(i).Enabled = True
    End If
     If TypeOf frm.Controls(i) Is ComboBox Then
          frm.Controls(i).Enabled = True
     End If
     If TypeOf frm.Controls(i) Is OptionButton Then
          frm.Controls(i).Enabled = True
     End If
Next i

End Sub

Public Sub LimpaTexto(frm As Form)
 Dim i
 For i = 0 To frm.Controls.Count - 1
    If TypeOf frm.Controls(i) Is TextBox Then
       frm.Controls(i).Text = ""
    End If
    If TypeOf frm.Controls(i) Is ComboBox Then
          frm.Controls(i).Index = -1
    End If
    If TypeOf frm.Controls(i) Is CheckBox Then
          frm.Controls(i).Value = 0
    End If

'     If TypeOf frm.Controls(i) Is OptionButton Then
'          frm.Controls(i).Enabled = True
'     End If
Next i

End Sub

Public Sub sLimpaFrame(frm As Form)
'Deixa os textbox desabilitados
   Dim i
   
   For i = 0 To frm.Controls.Count - 1
       If TypeOf frm.Controls(i) Is TextBox Then
          'frm.Controls(i).Enabled = False
          Unload frm.Controls(i)
       End If
       'If TypeOf frm.Controls(i) Is MaskEdBox Then
       '   frm.Controls(i).Enabled = False
       'End If
'       If TypeOf frm.Controls(i) Is MSFlexGrid Then
'          frm.Controls(i).Enabled = True
'       End If
       If TypeOf frm.Controls(i) Is CommandButton Then
          'frm.Controls(i).Enabled = False
          Unload frm.Controls(i)
       End If
'       If TypeOf frm.Controls(i) Is OptionButton Then
'          frm.Controls(i).Enabled = False
'       End If
   Next i
   
End Sub

Public Function DadosCBOtabela(Cb As ComboBox, Tabela As String, campo As String, CodId As String) As Boolean
    Dim Selecao As ADODB.Recordset
    Call Conecta_Banco
    StrTemp = Cb.Text
    Cb.Clear
    Set Selecao = New ADODB.Recordset
    Selecao.Open "SELECT " & CodId & "," & campo & " FROM " & Tabela & " ORDER BY " & campo, Cnn, adOpenDynamic, adLockReadOnly
    If Selecao.EOF = True Then
        DadosCBOtabela = False
        Selecao.Close
        Exit Function
    End If
    Do While Not Selecao.EOF
        Cb.AddItem IIf(IsNull(Selecao(campo)), "", Trim(Selecao(campo)))
        Cb.ItemData(Cb.NewIndex) = Selecao(CodId)
        Selecao.MoveNext
    Loop
    DadosCBOtabela = True
    Selecao.Close
    Cnn.Close
    Cb.Text = StrTemp
    StrTemp = ""
    
End Function

Public Sub sCria_tabelas()
'
'*** Fabio Reinert (Alemão) - 09/2017 - Inicio
'*** Ve se existes as tabelas do sistema de Agendamento e se não existirem cria-as
'*
    
    On Error GoTo Erro_sCria_Tabelas
    
    Call Conecta_Banco '---> Somente aqui vou usar os comandos de conexao porque se der erro tem que apagar os arquivos de segurança

    'IF Firebird  -- Se é Firebird então:
    strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='TAB_ANIMAIS' "
    'ELSE -- Senão se for SQL Server
    'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'TAB_ANIMAIS'"
    'END IF
    If Rstemp2.State = adStateOpen Then
        Rstemp2.Close
    End If
    Set Rstemp2 = New ADODB.Recordset
    Rstemp2.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If Rstemp2(0) = 0 Then
       '***************************************************************************************
       '***************************   CRIA A TABELA DE ANIMAIS   ******************************
       '***************************************************************************************
       '
        strSql = "CREATE TABLE tab_animais  (ID integer NOT NULL" & _
                                            ", Id_Cli integer NOT NULL" & _
                                            ", Nome  character(50) NOT NULL" & _
                                            ", Tipo_ani Int not null" & _
                                            ", dt_nasc date" & _
                                            ", pedigree CHAR(1)" & _
                                            ", observacoes varchar(200)" & _
                                            ", cuidados_especiais varchar(100)" & _
                                            ", foto varchar(100)" & _
                                            ", dt_ult_visita date" & _
                                            ", operador character(10)" & _
                                            ", dt_Atualiza timestamp" & _
                                            ", primary key (ID) )"
        Cnn.Execute strSql
        Cnn.CommitTrans
            
        'cria GENERATOR
        strSql = "CREATE GENERATOR GEN_ANI_ID1 "
        Cnn.Execute strSql
        Cnn.CommitTrans
    
        strSql = "SET GENERATOR GEN_ANI_ID1 TO 0"
        Cnn.Execute strSql
        Cnn.CommitTrans
    
        strSql = " CREATE TRIGGER TAB_ANIMAIS_BI FOR TAB_ANIMAIS ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        strSql = strSql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_ANI_ID1, 1); END; "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
    
'***********************************************************************************
    'IF Firebird  -- Se é Firebird então:
    strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='TAB_TIPOS_AN' "
    'ELSE -- Senão se for SQL Server
    'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'TAB_TIPOS_AN'"
    'END IF
    If Rstemp2.State = adStateOpen Then
        Rstemp2.Close
    End If
    Set Rstemp2 = New ADODB.Recordset
    Rstemp2.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If Rstemp2(0) = 0 Then
        '****************************************************************************************
        '****************   CRIA A TABELA DE TIPOS DE ANIMAL (CÃO/GATO/COELHO)  *****************
        '****************************************************************************************
        '
        strSql = "CREATE TABLE tab_tipos_an (ID integer NOT NULL" & _
                                        ", Descricao character(50) NOT NULL" & _
                                        ", operador character(10)" & _
                                        ", dt_Atualiza timestamp" & _
                                        ", primary key (ID) )"
        Cnn.Execute strSql
        Cnn.CommitTrans
                
        'cria GENERATOR
        strSql = "CREATE GENERATOR GEN_TPA_ID1 "
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = "SET GENERATOR GEN_TPA_ID1 TO 0"
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = " CREATE TRIGGER TAB_TIPOS_AN_BI FOR TAB_TIPOS_AN ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        strSql = strSql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_TPA_ID1, 1); END; "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
'***********************************************************************************
    'IF Firebird  -- Se é Firebird então:
    strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='TAB_SERVICOS' "
    'ELSE -- Senão se for SQL Server
    'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'TAB_servicos'"
    'END IF
    If Rstemp2.State = adStateOpen Then
        Rstemp2.Close
    End If
    Set Rstemp2 = New ADODB.Recordset
    Rstemp2.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If Rstemp2(0) = 0 Then
        '****************************************************************************************
        '*****************  CRIA A TABELA DE SERVICOS - BANHO/TOSA/VACINAS/ETC  *****************
        '****************************************************************************************
        '
        strSql = "CREATE TABLE tab_servicos (ID integer NOT NULL" & _
                                        ", Descricao character(50) NOT NULL" & _
                                        ", valor NUMERIC(12,2)" & _
                                        ", TEMPO_EST NUMERIC(12,2)" & _
                                        ", vacina CHAR(1)" & _
                                        ", operador character(10)" & _
                                        ", dt_Atualiza timestamp" & _
                                        ", primary key (ID) )"
        Cnn.Execute strSql
        Cnn.CommitTrans
                
        'cria GENERATOR
        strSql = "CREATE GENERATOR GEN_SERV_ID1 "
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = "SET GENERATOR GEN_SERV_ID1 TO 0"
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = " CREATE TRIGGER TAB_SERVICOS_BI FOR TAB_SERVICOS ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        strSql = strSql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_SERV_ID1, 1); END; "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
    
'***********************************************************************************
    'IF Firebird  -- Se é Firebird então:
    strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='TAB_ATENDIMENTOS' "
    'ELSE -- Senão se for SQL Server
    'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'TAB_atendimentos'"
    'END IF
    If Rstemp2.State = adStateOpen Then
        Rstemp2.Close
    End If
    Set Rstemp2 = New ADODB.Recordset
    Rstemp2.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If Rstemp2(0) = 0 Then
        '****************************************************************************************
        '********************    CRIA A TABELA DE ATENDIMENTOS    *******************************
        '****************************************************************************************
        '
        strSql = "CREATE TABLE tab_atendimentos (Dt_atend timestamp NOT NULL" & _
                                                ", IdAnimal integer NOT NULL" & _
                                                ", Tipo_Atend integer NOT NULL" & _
                                                ", valor NUMERIC(12,2)" & _
                                                ", valor_recebido NUMERIC(12,2)" & _
                                                ", hora_saida CHAR(5)" & _
                                                ", observa VARCHAR(150)" & _
                                                ", operador char(10)" & _
                                                ", dt_Atualiza timestamp" & _
                                                ", primary key (dt_atend) )"
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
    'Não tem auto incremento porque o campo chave é TIMESTAMP

'***********************************************************************************
    'IF Firebird  -- Se é Firebird então:
    strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='TAB_VACINAS' "
    'ELSE -- Senão se for SQL Server
    'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'TAB_vacinas'"
    'END IF
    If Rstemp2.State = adStateOpen Then
        Rstemp2.Close
    End If
    Set Rstemp2 = New ADODB.Recordset
    Rstemp2.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If Rstemp2(0) = 0 Then
        strSql = "CREATE TABLE tab_vacinas (ID integer NOT NULL " & _
                                        ",IdAnimal integer NOT NULL " & _
                                        ",Dt_atend timestamp NOT NULL " & _
                                        ",Descricao VARCHAR(100) NOT NULL " & _
                                        ",valor NUMERIC(12,2) " & _
                                        ",DT_PROXIMA DATE " & _
                                        ",operador character(10) " & _
                                        ",dt_Atualiza timestamp " & _
                                        ",primary key (id)  )"
        Cnn.Execute strSql
        Cnn.CommitTrans
                
        'cria GENERATOR
        strSql = "CREATE GENERATOR GEN_TVAC_ID1 "
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = "SET GENERATOR GEN_TVAC_ID1 TO 0"
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = " CREATE TRIGGER TAB_VACINAS_BI FOR TAB_VACINAS ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        strSql = strSql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_TVAC_ID1, 1); END; "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If

'***********************************************************************************
    'IF Firebird  -- Se é Firebird então:
    strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='TAB_PROMOCOES' "
    'ELSE -- Senão se for SQL Server
    'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'TAB_promocoes'"
    'END IF
    If Rstemp2.State = adStateOpen Then
        Rstemp2.Close
    End If
    Set Rstemp2 = New ADODB.Recordset
    Rstemp2.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If Rstemp2(0) = 0 Then
        '****************************************************************************************
        '********************       CRIA A TABELA DE PROMOCOES      *******************************
        '****************************************************************************************
        '
        strSql = "CREATE TABLE tab_promocoes (ID integer NOT NULL" & _
                                          ",Dt_inicio timestamp  NOT NULL" & _
                                          ",Dt_fim timestamp NOT NULL" & _
                                          ",IdAnimal integer" & _
                                          ",IdTipoAten integer" & _
                                          ",Descricao VARCHAR(100) NOT NULL" & _
                                          ",Valor NUMERIC(12,2)" & _
                                          ",porcent NUMERIC(2,2)" & _
                                          ",operador character(10)" & _
                                          ",Dt_Atualiza timestamp" & _
                                          ",primary key (ID)  )"

        Cnn.Execute strSql
        Cnn.CommitTrans
        
        'cria GENERATOR
        strSql = "CREATE GENERATOR GEN_TPRO_ID1 "
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = "SET GENERATOR GEN_TPRO_ID1 TO 0"
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = " CREATE TRIGGER TAB_PROMOCOES_BI FOR TAB_PROMOCOES ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        strSql = strSql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_TPRO_ID1, 1); END; "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
'
'****************************************************************
'**** Fabio Reinert (Alemão) - Criar a tabela de cartões        *
'****************************************************************
'
    'IF Firebird  -- Se é Firebird então:
    strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='TAB_CARTOES' "
    'ELSE -- Senão se for SQL Server
    'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'TAB_promocoes'"
    'END IF
    If Rstemp2.State = adStateOpen Then
        Rstemp2.Close
    End If
    Set Rstemp2 = New ADODB.Recordset
    Rstemp2.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If Rstemp2(0) = 0 Then

        '***************************************************
        '*                 Tabela tab_Cartao               *
        '***************************************************
        strSql = "CREATE TABLE TAB_CARTOES ( IDCartao INTEGER NOT NULL, " & _
                                            " bandeira VARCHAR(50)," & _
                                            " carenciacredito SMALLINT, " & _
                                            " carenciadebito SMALLINT, " & _
                                            " planodecontas VARCHAR(110), " & _
                                            " codconta INTEGER, " & _
                                            " tx0 FLOAT," & _
                                            " tx1 FLOAT," & _
                                            " tx2 FLOAT, " & _
                                            " tx3 FLOAT, " & _
                                            " tx4 FLOAT, " & _
                                            " tx5 FLOAT, " & _
                                            " tx6 FLOAT, " & _
                                            " tx7 FLOAT, " & _
                                            " tx8 FLOAT, " & _
                                            " tx9 FLOAT, " & _
                                            " tx10 FLOAT, " & _
                                            " tx11 FLOAT, " & _
                                            " tx12 FLOAT " & _
                                            ",primary key (IDCartao) ) "

        Cnn.Execute strSql
        Cnn.CommitTrans
        
        'cria GENERATOR
        strSql = "CREATE GENERATOR GEN_TCAR_ID1 "
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = "SET GENERATOR GEN_TCAR_ID1 TO 0"
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = " CREATE TRIGGER TAB_CARTOES_BI FOR TAB_CARTOES ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        strSql = strSql & "if (NEW.IDCartao is NULL) then NEW.IDCartao = GEN_ID(GEN_TCAR_ID1, 1); END; "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
'
'******************************************************************************
'**** Fabio Reinert (Alemão) - 08/2017 - Criar a tabela de cartões a receber  *
'******************************************************************************
'
    'IF Firebird  -- Se é Firebird então:
    strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='TAB_RECCARTOES' "
    'ELSE -- Senão se for SQL Server
    'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'TAB_RECCARTOES'"
    'END IF
    If Rstemp2.State = adStateOpen Then
        Rstemp2.Close
    End If
    Set Rstemp2 = New ADODB.Recordset
    Rstemp2.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If Rstemp2(0) = 0 Then

        '***************************************************
        '*                 Tabela tab_reccartoes           *
        '***************************************************
        strSql = "CREATE TABLE TAB_RECCARTOES ( SEQUENCIA INTEGER NOT NULL, " & _
                                            " CODLIENTE integer NOT NULL," & _
                                            " TIPO_CARTAO CHAR(1), " & _
                                            " COD_CARTAO integer NOT NULL, " & _
                                            " DT_EMISSAO DATE, " & _
                                            " VALOR NUMERIC912,2) , " & _
                                            " DT_VENCTO DATE," & _
                                            " DT_BAIXA DATE," & _
                                            " OPERADOR VARCHAR(10), " & _
                                            " DT_ATUALIZA ) "

        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
'
'**************************************************************************************
'**** Fabio Reinert (Alemão) - 08/2017 - Criar a tabela de dinheiro recebido a vista  *
'**************************************************************************************
'
    'IF Firebird  -- Se é Firebird então:
    strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='TAB_RECAVISTA' "
    'ELSE -- Senão se for SQL Server
    'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'TAB_RECAVISTA'"
    'END IF
    If Rstemp2.State = adStateOpen Then
        Rstemp2.Close
    End If
    Set Rstemp2 = New ADODB.Recordset
    Rstemp2.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If Rstemp2(0) = 0 Then

        '***************************************************
        '*                 Tabela tab_recavista            *
        '***************************************************
        strSql = "CREATE TABLE TAB_RECAVISTA ( SEQUENCIA INTEGER NOT NULL, " & _
                                            " CODLIENTE integer NOT NULL," & _
                                            " DT_RECEBIDO DATE, " & _
                                            " VALOR NUMERIC(12,2), " & _
                                            " OPERADOR VARCHAR(10), " & _
                                            " DT_ATUALIZA )"

        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
'

    'IF Firebird  -- Se é Firebird então:
    strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='TAB_MOEDAS' "
    'ELSE -- Senão se for SQL Server
    'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'TAB_promocoes'"
    'END IF
    If Rstemp2.State = adStateOpen Then
        Rstemp2.Close
    End If
    Set Rstemp2 = New ADODB.Recordset
    Rstemp2.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If Rstemp2(0) = 0 Then
        '***************************************************
        '*     Tabela de moedas / Formas de pagamento      *
        '***************************************************
        strSql = "CREATE TABLE tab_moedas ( IDMoeda   integer NOT NULL " & _
                                          ",Descricao varChar(50) NOT NULL " & _
                                          ",primary key (IDMoeda) ) "

        Cnn.Execute strSql
        Cnn.CommitTrans
        
        'cria GENERATOR
        strSql = "CREATE GENERATOR GEN_TMOE_ID1 "
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = "SET GENERATOR GEN_TMOE_ID1 TO 0"
        Cnn.Execute strSql
        Cnn.CommitTrans
        
        strSql = " CREATE TRIGGER TAB_MOEDAS_BI FOR TAB_MOEDAS ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        strSql = strSql & "if (NEW.IDMoeda is NULL) then NEW.IDMoeda = GEN_ID(GEN_TMOE_ID1, 1); END; "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If

'
'**************************************************************************************************

    If Rstemp.State = adStateOpen Then
        Rstemp.Close
    End If
    If Rstemp2.State = adStateOpen Then
        Rstemp2.Close
    End If

    If Cnn.State = adStateOpen Then
        Cnn.Close
    End If

    Exit Sub

Erro_sCria_Tabelas:
    Call sMostraErro("Módulo de Criação das tabelas", Err.Number, Err.Description)
    End
End Sub
'*
'*** Fabio Reinert (Alemão) - 08/2017 - Função de validação de data - Inicio
'*
'
Public Function fValidaData(ByVal sData As String, Optional ByVal sFormato As String) As Boolean
    Dim Dia As Integer, Dia_Pos As Integer
    Dim Mes As Integer, Mes_Pos As Integer
    Dim Ano As Integer, Ano_Pos As Integer
    Dim DDOk As Integer, MMOk As Integer
    Dim YYOk As Integer, i As Integer
    Dim m As Integer, Temp As String
    Dim sBst As Boolean

'P/ chamar, no evento desejado:
'VariávelBoolean = ValidaData(Data, Formato)

'Exemplo:
'  Dim bRESP As Boolean
'  bRESP = ValidaData("31/03/2000", "M/D/Y")
'  If bRESP Then
'    MsgBox "A data é válida!!!"
'  Else
'    MsgBox "A data NÃO é válida!!!"
'  End If
'Ele exibirá "A data é válida!!!"

    If IsMissing(sFormato) Then
        sFormato = "DD/MM/YYYY"
        'OU então, você pode pegar o formato
        'que estiver configurado no Windows.
    Else
        If Len(sFormato) = 0 Then
            sFormato = "DD/MM/YYYY"
        End If
    End If
   
    Temp = Replace(sData, "-", "/")
    sData = Temp
    Temp = Replace(sFormato, "-", "/")
    sFormato = Temp
    Temp = ""

    DDOk = 0
    MMOk = 0
    YYOk = 0

    For i = 1 To Len(sFormato)
        If UCase(Mid(sFormato, i, 1)) = "D" Then
            If DDOk > 2 Then
                fValidaData = False
                Exit Function
            Else
                DDOk = DDOk + 1
                If Dia_Pos = 0 Then
                    Dia_Pos = Mes_Pos + Ano_Pos + 1
                End If
            End If
        ElseIf UCase(Mid(sFormato, i, 1)) = "M" Then
            If MMOk > 2 Then
                fValidaData = False
                Exit Function
            Else
                MMOk = MMOk + 1
                If Mes_Pos = 0 Then
                    Mes_Pos = Dia_Pos + Ano_Pos + 1
                End If
            End If
        ElseIf UCase(Mid(sFormato, i, 1)) = "Y" Then
            If YYOk > 4 Then
                fValidaData = False
                Exit Function
            Else
                YYOk = YYOk + 1
                If Ano_Pos = 0 Then
                    Ano_Pos = Dia_Pos + Mes_Pos + 1
                End If
            End If
        Else
            Select Case UCase(Mid(sFormato, i, 1))
            Case "D", "M", "Y", "/"
            Case Else
                fValidaData = False
                Exit Function
            End Select
        End If
    Next i

    If DDOk = 0 Or MMOk = 0 Then
        fValidaData = False
        Exit Function
    End If

    If YYOk = 0 Or YYOk > 4 Then
        fValidaData = False
        Exit Function
    End If

    If Not IsDate(sData) Then
        fValidaData = False
        Exit Function
    End If

    m = 0
    For i = 1 To Len(sData)
        If Mid(sData, i, 1) = "/" Or i = Len(sData) Then
            If i = Len(sData) Then
                Temp = Temp & Mid(sData, i, 1)
            End If
            m = m + 1
            If m = 3 Then
                m = 4
            End If
            If Dia_Pos = m Then
                Dia = Temp
            ElseIf Mes_Pos = m Then
                Mes = Temp
            ElseIf Ano_Pos = m Then
                Ano = Temp
            End If
            Temp = ""
        Else
            Temp = Temp & Mid(sData, i, 1)
        End If
    Next i

    Select Case Mes
    Case 1, 3, 5, 7, 8, 10, 12
        If Dia < 1 Or Dia > 31 Then
            fValidaData = False
            Exit Function
        End If
    Case 4, 6, 9, 11
        If Dia < 1 Or Dia > 31 Then
            fValidaData = False
            Exit Function
        End If
    Case 2
        If Dia < 1 Or Dia > 29 Then
            fValidaData = False
            Exit Function
        ElseIf Dia = 29 Then
            sBst = False
            If Ano = 0 Then
                sBst = True
            ElseIf Ano Mod 4 = 0 Then
                sBst = True
                If Ano Mod 100 = 0 Then
                    sBst = False
                    If Ano Mod 400 = 0 Then
                        sBst = True
                    End If
                End If
            Else
                sBst = False
            End If
            If sBst = False Then
                fValidaData = False
                Exit Function
            End If
        End If
    Case Else
        fValidaData = False
        Exit Function
    End Select
    fValidaData = True

End Function
'*
'*** Fabio Reinert (Alemão) - 08/2017 - Função de validação de data - Fim
'*
'

Public Sub sMostraAviso(Optional ByVal pTitulo As String, Optional ByVal pTexto1 As String, _
                        Optional ByVal pTexto2 As String, _
                        Optional ByVal pTexto3 As String, _
                        Optional ByVal pTexto4 As String)
                        
    Dim fAviso As Form
    If IsMissing(pTexto2) Then
        pTexto2 = ""
    End If
    If IsMissing(pTexto3) Then
        pTexto3 = ""
    End If
    If IsMissing(pTexto4) Then
        pTexto4 = ""
    End If
    If IsMissing(pTitulo) Then
        pTitulo = "Aviso:"
    End If
    Set fAviso = New frmAviso
    fAviso.lblAviso1.Caption = pTexto1
    fAviso.lblAviso2.Caption = pTexto2
    fAviso.lblAviso3.Caption = pTexto3
    fAviso.lblAviso4.Caption = pTexto4
    fAviso.Caption = pTitulo
    fAviso.Show vbModal
    Unload fAviso
    Set fAviso = Nothing
End Sub

Public Sub sMostraErro(Optional ByVal pModulo, Optional ByVal pErroNumero, Optional ByVal pErroDesc)
        
    If pModulo = "" Then
        pModulo = "Geral"
    End If
    If pErroNumero = "" Then
       pErroNumero = Err.Number
    End If
    If pErroDesc = "" Then
       pErroDesc = Err.Description
    End If
    Call sMostraAviso("Atenção - Erro: ", "Contate a Novavia informando o erro abaixo:", _
                      "No.erro: " & pErroNumero & " Descr.: " & pErroDesc, _
                      "Módulo do erro: " & pModulo, "Sistema será encerrado")
    Call Fecha_Formularios
    End
End Sub

