Attribute VB_Name = "ModSeguranca"
''Fabio Reinert - 04/2017 - Declaração publica usada no módulo de segurança. Verificar se tem conexão com a WEB - Inicio
'Declare Function InternetGetConnectedState Lib _
'    "wininet" (ByRef dwflags As Long, ByVal dwReserved As _
'    Long) As Long
''Fabio Reinert - 04/2017 - Declaração publica usada no módulo de segurança. Verificar se tem conexão com a WEB - Fim

'*** Fabio Reinert - 04/2017 - Módulo de segurança - Declaração de variáveis - Inicio
Public CnnLocaWEB As New ADODB.Connection
Public RsSegur As New ADODB.Recordset
Public gCNPJ As String
Public gMinimo As String
Public gMaximo As String
Public gDataServidorWEB As String
Public gDataServidorLocal As String
Public gDataServidor As String
Public Ok As Boolean
Public sCaminho As String
Public gTemWEB As Boolean
'*** Fabio Reinert - 04/2017 - Módulo de segurança - Declaração de variáveis - Fim

'*** Fabio Reinert - Funções do módulo de segurança - Inicio
'*** Também foram incluidos os seguintes forms:
'***      frmAviso, frmCNPJ e frmDialog
'***
Function fGeraSenha(ByVal pCNPJ As String)

    On Error GoTo Erro_fGeraSenha
    
    'em primeiro teste quando o usuario exclui o arquivo ini
    If Len(Trim(pCNPJ)) = 0 Then
        'busca o cnpj no tabela balanca
        Call sLe_Tab_Balanca
        pCNPJ = Decrypt(RsSegur!peso)
    End If

   If Len(Trim(pCNPJ)) = 14 Then
       fGeraSenha = Format(Left(Trim(Str(Int(((Val(pCNPJ) + Val(Mid(pCNPJ, 3, 3)) - Val(Mid(pCNPJ, 6, 3))) * Day(Date)) / Month(Date)))), 14), "00000000000000")
   Else
       fGeraSenha = Format(Left(Trim(Str(Int(((Val(pCNPJ) + Val(Mid(pCNPJ, 3, 3)) - Val(Mid(pCNPJ, 6, 3))) * Day(Date)) / Month(Date)))), 11), "00000000000")
   End If
   Exit Function
   
Erro_fGeraSenha:
   Call sMostraErro("fGeraSenha", Err.Number, Err.Description)
   Call Fecha_Formularios
   End
End Function

Public Function sConectaWEB(Optional ByVal tem_web As Boolean)
    sConectaWEB = False
    Screen.MousePointer = 11
    
    On Error GoTo traErrWEB
    
    Set CnnLocaWEB = New ADODB.Connection
    With CnnLocaWEB
        .CursorLocation = adUseClient
        .Open "Provider=SQLOLEDB.1;Password=billgodes222;Persist Security Info=True;User ID=novavia3;Data Source=sqlserver03.novavia.com.br;Extended Properties=ConnectionTimeout=3;Connect Timeout=2"
    End With
    
    sConectaWEB = True
    
    Screen.MousePointer = 1
    
    Exit Function

traErrWEB:

    If Err.Number <> 0 Then
        Err.Clear
    End If
    
End Function

Public Sub sCriaCNPJ_CPF()
    '**** Subrotina caso seja a primeira execução do programa e tem que criar a tabela de segurança (tab_balanca)
    '**** e também o arquivo fArquivo.ini no caminho da aplicação
    '**** Chama um formulário para atualização do CNPJ/CPF - Inicio
    
    On Error GoTo Erro_Cria_CNPJ
    
    Ok = False
    
    With frmCNPJ
       .Height = 2300
       .txtCNPJ.Text = FormatCPF_CNPJ(gCNPJ)
       .Show 1
    End With
        
    If Not Ok Then
        'Usuario cancelou, então abortar o sistema
        If Dir("\\servidor\Petshop\fArquivo.ini") <> "" Then
            Kill "\\servidor\Petshop\fArquivo.ini"
        End If
        Call sRestauraSegur
        Call Fecha_Formularios
        End
    End If
        
    'Unload fCNPJ
    'Set fCNPJ = Nothing
    Call sMostraAviso("Aviso", "CNPJ/CPF: " & FormatCPF_CNPJ(gCNPJ) & " Criado com sucesso", "Clique em Ok para continuar")
    'MsgBox "CNPJ/CPF: " & FormatCPF_CNPJ(gCNPJ) & " Criado com sucesso. Clique em Ok para continuar", vbOKOnly
    Exit Sub
    
Erro_Cria_CNPJ:
    Call sMostraAviso("ERRO", "Erro na criação do CNPJ:", Err.Description, "Favor contactar a Novavia informado o erro acima")
    
    If Dir("\\servidor\Petshop\fArquivo.ini") <> "" Then
         Kill "\\servidor\Petshop\fArquivo.ini"
    End If
    Call sRestauraSegur
    Call Fecha_Formularios
    End
End Sub

Public Sub sLibera(Optional pErro As Integer)
    
    On Error GoTo Erro_sLibera
    
    If Len(gCNPJ) = 0 Then
        Call sMostraErro("sLibera", "110")
        Call Fecha_Formularios
        Call sRestauraSegur
        End
    End If
    
    Dim fDialog As New frmDialog
    Screen.MousePointer = 1
    If pErro > 0 Then
        If pErro = 333 Then    '*** Caso seja trava na movimentação de pedidos, então o código é 333
                               '*** e vai dar mensagem diferente da normal
           fDialog.Lbltitulo2.Caption = "Entrar em contato com a Novavia informando o CNPJ abaixo."
        Else
           fDialog.Lbltitulo2.Caption = fDialog.Lbltitulo2.Caption & " e o código: " & Str(pErro)
        End If
    End If
    fDialog.txtCNPJ.Text = FormatCPF_CNPJ(gCNPJ)
    fDialog.Show vbModal
    If Ok Then
        strMinimo = Encrypt(Left(gDataServidor, 2) & Mid(gDataServidor, 4, 2) & Right(gDataServidor, 4))
        strMaximo = Encrypt(Left(fSomaMes(gDataServidor, 1), 2) & Mid(fSomaMes(gDataServidor, 1), 4, 2) + Right(fSomaMes(gDataServidor, 1), 4))
        
        Call sConectaBanco
'
        sql = "UPDATE TAB_BALANCA SET PESO = '" & Encrypt(gCNPJ) & "', MINIMO = '" & strMinimo & "',MAXIMO = '" & strMaximo & "',MODELO = '" & Encrypt("0") & "'"
        Cnn.Execute sql
        Cnn.CommitTrans
    Else
       'Usuario cancelou, então abortar o sistema
       Call Fecha_Formularios
       Call sRestauraSegur
       End
    End If
    Unload fDialog
    Set fDialog = Nothing
    Exit Sub
    
Erro_sLibera:
   Call sMostraErro("sLibera", Err.Number, Err.Description)
   Call Fecha_Formularios
   End
End Sub

Public Sub sLeDataServidorLocal()
    
    Dim sDefaultValue As String
    
    On Error GoTo Erro_sLedataServidorLocal
    
    Call sConectaBanco
    
    'pega data do servidor firebird
    sql = "SELECT current_date from rdb$database"
    Set RsSegur = New ADODB.Recordset
    
    If RsSegur.State = adStateOpen Then
       RsSegur.Close
    End If
    RsSegur.Open sql, Cnn, 1, 2
    gDataServidorLocal = RsSegur(0)
    
    If RsSegur.State = adStateOpen Then
       RsSegur.Close
    End If
    
    If Cnn.State = adStateOpen Then
        Cnn.Close
    End If
    Exit Sub

Erro_sLedataServidorLocal:
   Call sMostraErro("sLeDataServidorLocal", Err.Number, Err.Description)
   Call Fecha_Formularios
   End

End Sub

Public Sub sVeTabBalanca()

'*** Ve se existe a tabela TAB_BALANCA e se não existir cria com 1 registro criptografado com o CNPJ e data atual + 30 dias - Inicio
'*** Chamada pela subrotina sValidaCliente
'*
'    Do While True
    
    On Error GoTo Erro_sVeTabBalanca
    
    'Call sConectaBanco ---> Somente aqui vou usar os comandos de conexao porque se der erro tem que apagar os arquivos de segurança
    Set Cnn = New ADODB.Connection
    With Cnn
       .CursorLocation = adUseClient
        '.Open "File Name=" & App.Path & "\cnn_fire_Servidor.udl;"
        .Open "Provider=IBOLE.Provider.v4;Password=masterkey;Persist Security Info=True;Data Source=servidor:c:\Sistema Petshop\ARQDADOS.GDB;Mode=ReadWrite|Share Deny None"
    End With

    'IF Firebird  -- Se é Firebird então:
    strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='TAB_BALANCA' "
    'ELSE -- Senão se for SQL Server
    'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'TAB_BALANCA'"
    'END IF
    If RsSegur.State = adStateOpen Then
        RsSegur.Close
    End If
    Set RsSegur = New ADODB.Recordset
    RsSegur.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
    
    If RsSegur(0) = 0 Then
        '***  TAB_BALANCA INEXISTENTE
        If Dir("\\servidor\Petshop\fArquivo.ini") = "" Then  '*** Se o arquivo .ini não existe então, criamos
            Open "\\servidor\Petshop\fArquivo.ini" For Output As 1
            ' Print 1, " "
            Close 1
            strSql = "CREATE TABLE TAB_BALANCA(PESO VARCHAR(20), MINIMO VARCHAR(10), MAXIMO VARCHAR(10), MARCA VARCHAR(15), MODELO VARCHAR(1))"
            Cnn.Execute strSql
            Cnn.CommitTrans
            strSql = "INSERT INTO TAB_BALANCA (PESO,MINIMO,MAXIMO,MARCA,MODELO) VALUES('" & "" & "','" & "000000000" & "','" & "99999999" & "','FILIZOLA',' ')"
            Cnn.Execute strSql
            Cnn.CommitTrans
            If RsSegur.State = adStateOpen Then
                RsSegur.Close
            End If
            If Cnn.State = adStateOpen Then
                Cnn.Close
            End If
            '**** Não existe a tab_balanca e nem o arquivo .ini então...
            '**** Abre a tela para digitação do CNPJ/CPF para cadastro da empresa
            Call sCriaCNPJ_CPF
            '****
        Else   '*** Arquivo INI existe mas a TAB_BALANCA não.
               '*** Mostrar mensagem de erro e sair do sistema.
            If RsSegur.State = adStateOpen Then
                RsSegur.Close
            End If
            Set RsSegur = Nothing
            If Cnn.State = adStateOpen Then
                Cnn.Close
            End If
            Set Cnn = Nothing
            'Call sRestauraSegur
            Call sMostraAviso("Aviso", "Favor entrar com contato com a Novavia", "", "Informe o codigo de erro: 500")
            Call Fecha_Formularios
            End
         
        End If
    Else
        If Len(gCNPJ) = 0 Then
            sLe_Tab_Balanca
        End If
        If Dir("\\servidor\Petshop\fArquivo.ini") = "" Then
                Call sLibera(555)
                If RsSegur.State = adStateOpen Then
                    RsSegur.Close
                End If
                Set RsSegur = Nothing
                If Cnn.State = adStateOpen Then
                    Cnn.Close
                End If
                Set Cnn = Nothing
                Open "\\servidor\Petshop\fArquivo.ini" For Output As 1
                Close 1
                Call Fecha_Formularios
            End
        Else
            '*** Tab_balanca existe mas verificar se tem registros
            strSql = "SELECT COUNT(*) QTDE FROM TAB_BALANCA"
            Call sConectaBanco
            If RsSegur.State = adStateOpen Then
                RsSegur.Close
            End If
            Set RsSegur = New ADODB.Recordset
            RsSegur.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
            If RsSegur(0) = 0 Then
                sLibera (100)
'            Else
'                Exit Do
            End If
        End If
    End If
'    Loop
    If RsSegur.State = adStateOpen Then
       RsSegur.Close
    End If
    If Cnn.State = adStateOpen Then
        Cnn.Close
    End If

    Set RsSegur = Nothing
    Set Cnn = Nothing
'*** Ve se existe a tabela TAB_BALANCA e se não existir cria com 1 registro criptografado com o CNPJ e data atual + 30 dias - Fim
    Exit Sub
    
Erro_sVeTabBalanca:
    Call sMostraErro("sVeTabBalanca", Err.Number, Err.Description)
    Call sRestauraSegur
    Call Fecha_Formularios
    End

End Sub

Public Sub sValidaCliente()
'
'****
'**** Será feita a validação de segurança e caso haja qualquer problema ira
'**** chamar a rotina (sLibera) que mostra a tela para entrar em contato com a Novavia e pedir uma senha
'**** de liberação do sistema - Caso esteja Offline, ficará liberado por 30 dias
'**** Call sLibera
'****
    On Error GoTo Erro_sValidaCliente
    
    gDataServidorWEB = ""
    gDataServidorLocal = ""
    gDataServidor = ""
    gTemWEB = False
   
    If fcheckInternetConnection() Then  'Função que verifica se tem acesso a WEB
        'sCaminho = "\\servidor\Petshop\"
        Call sLeDataServidorWEB '-----> Pega a data do servidor WEB
        gDataServidor = gDataServidorWEB
        gTemWEB = True
    End If
    
    Call sLeDataServidorLocal '-----> Pega a data do servidor Local
    
'    If gDataServidorWEB <> gDataServidorLocal And gDataServidorWEB <> "" And gDataServidorLocal <> "" Then
'        Call sMostraAviso("Atenção", "Data do servidor WEB diferente da data do sistema", "", "Favor corrigir a data do seu computador")
'        Call Fecha_Formularios
'        End
'    End If
    If gDataServidorWEB = "" Then 'Data do servidor WEB em branco, pode ser porque não te conexão
        gDataServidorWEB = gDataServidorLocal
        gDataServidor = gDataServidorLocal
    Else
        If gDataServidorWEB <> gDataServidorLocal Then
            Call sMostraAviso("Atenção", "Data do servidor WEB: " & gDataServidorWEB, "Data do servidor local: " & gDataServidorLocal, "Diferem. Favor acertar a data do servidor local")
            'Call Fecha_Formularios
            End
        End If
    End If
    If gDataServidorWEB <> VBA.Format(VBA.Date, "dd/mm/yyyy") Then
        Call sMostraAviso("Atenção", "Data do servidor WEB: " & gDataServidorWEB, "Data da estação de trabalho: " & VBA.Date, "Diferem. Favor acertar a data do seu computador")
        'Call Fecha_Formularios
        End
    End If
    
    Call sVeTabBalanca   '----->   Verifica se existe a tab_balanca e a cria se não existir
        
    '**** Acessa a LocalWEB para verificar se o cliente não está inadimplente
    '**** Caso esteja inadimplente ou não esteja cadastrado ou não tenha conexão
    'Call sTrataWEB
    If Not fTrataWEB Then  '**** Caso esteja inadimplente ou não esteja cadastrado ou não tenha conexão segurança local
        Do While True

'    '**** Vou ter que usar um Goto porque o VB6 não tem um comando CONTINUE
'Continua_Validar:

            Call sLe_Tab_Balanca

            '*** Campo PESO da tab_balanca tem o CNPJ/CPF da empresa
            '*** Campo MINIMO tem a ultima data de utilizacao do sistema
            '*** Campo MAXIMO tem a data limite de utilização do sistema
            '*** Campo MARCA contem apenas uma constante (FILIZOLA)
            '*** Campo MODELO tem o status do cliente:
            '***                    = 0 - Normal
            '***                    > 0 - Bloqueado
            '***

            '*** Alguém colocou brancos no CNPJ/CPF na tab_balanca ou tab_balanca  - Mostrar tela de solicitacao de senha
            If RsSegur!peso = "" Then
                '**** Rotina que mostra a tela para entrar em contato com a Novavia e pedir uma senha
                '**** Código de erro 110
                Call sLibera(110)
            '    GoTo Continua_Validar
                Exit Do
            End If

            '*** Alguém colocou o CNPJ/CPF na tab_balanca - Mostrar tela de solicitacao de senha
            If IsNumeric(RsSegur!peso) Then
                Call sLibera(120)
            '    GoTo Continua_Validar
                Exit Do
            End If

            'Descriptografa o CNPJ/CPF que está na tab_balanca
            gCNPJ = Decrypt(RsSegur!peso)   '---> CGC/CPF da tabela local tab_balanca

            '**** Verifica se o CPF/CNPJ é numerico - Se não, mostra a tela de solicitação de senha
            If Not IsNumeric(gCNPJ) Then
                Call sLibera(150)
            '    GoTo Continua_Validar
                Exit Do
            End If

            '*** Cliente
            '*** Alguém colocou a data sem criptografar no campo MINIMO  da Tab_balanca
            If IsDate(VBA.Left(RsSegur!minimo, 2) & "/" & VBA.Mid(RsSegur!minimo, 3, 2) & "/" & VBA.Right(RsSegur!minimo, 4)) Then   '*** Alguém colocou o CGC/CPF na tab_balanca - Abortar
                Call sLibera(130)
            '    GoTo Continua_Validar
                Exit Do
            End If

            '*** Alguém colocou a data sem criptografar no campo MAXIMO da Tab_balanca
            If IsDate(VBA.Left(RsSegur!maximo, 2) & "/" & VBA.Mid(RsSegur!maximo, 3, 2) & "/" & VBA.Right(RsSegur!maximo, 4)) Then   '*** Alguém colocou o CGC/CPF na tab_balanca - Abortar
                Call sLibera(140)
            '    GoTo Continua_Validar
                Exit Do
            End If

            'Campo MODELO (Status) está em branco - Deveria estar com 0 - Mostra a tela de solicitação de senha
            If RsSegur!Modelo = "" Then
                Call sLibera(160)
            '    GoTo Continua_Validar
                Exit Do
            End If

            'Campo MODELO (Status) está diferente de 0 - Mostra a tela de solicitação de senha
            If Decrypt(VBA.Trim(RsSegur!Modelo)) <> "0" Then
                '**** Aqui vai chamar um formulario para atualização da data de utilização do software - Inicio
                Call sLibera(170)
            '    GoTo Continua_Validar
                Exit Do
            End If

            'Descriptografa a data atual (MINIMO)
            gMinimo = VBA.Left(Decrypt(RsSegur!minimo), 2) & "/" & VBA.Mid(Decrypt(RsSegur!minimo), 3, 2) & "/" & VBA.Right(Decrypt(RsSegur!minimo), 4)
            '**** Verificar se a data atual é valida - se invalida, mostra a tela de solicitação de senha
            If Not IsDate(gMinimo) Then
                Call sLibera(180)
            '    GoTo Continua_Validar
                Exit Do
            End If

            'Descriptografa a data atual (MAXIMO)
            gMaximo = VBA.Left(Decrypt(RsSegur!maximo), 2) & "/" & VBA.Mid(Decrypt(RsSegur!maximo), 3, 2) & "/" & VBA.Right(Decrypt(RsSegur!maximo), 4)
            '**** Verificar se a data é valida - se invalida, mostra a tela de solicitação de senha
            If Not IsDate(gMaximo) Then
                Call sLibera(190)
            '    GoTo Continua_Validar
                Exit Do
            End If

            If Cnn.State = adStateOpen Then
                Cnn.Close
            End If
            Set Cnn = Nothing
            Exit Do

        Loop
        '*** Rotina que faz consistencia dos campos de data inicial e final
        Call sTrataLocal

        If RsSegur.State = adStateOpen Then
            RsSegur.Close
        End If

        If Cnn.State = adStateOpen Then
            Cnn.Close
        End If

        Set CnnLocaWEB = Nothing
        Set Cnn = Nothing
        Set RsSegur = Nothing
    End If
    Exit Sub
    
Erro_sValidaCliente:
    Call sMostraErro("sValidaCliente", Err.Number, Err.Description)
    Call Fecha_Formularios
    End
   
End Sub

Public Sub sLe_Tab_Balanca()

    On Error GoTo Erro_sLe_Tab_Balanca
    
    Call sConectaBanco
            
    '*** Reler a tab_balanca pois pode ser que o usuario tenha deixado o PC ligado durante varios dias
    Set RsSegur = New ADODB.Recordset
    strSql = "SELECT PESO,MINIMO,MAXIMO,MODELO FROM TAB_BALANCA"
    If RsSegur.State = adStateOpen Then
        RsSegur.Close
    End If
    RsSegur.Open strSql, Cnn, 1, 2
    '**** Atualiza CNPJ
    gCNPJ = Decrypt(RsSegur!peso)
    '**** Atualiza os campos de Data Inicio e Final da tab_balanca
    gMinimo = Left(Decrypt(RsSegur!minimo), 2) & "/" & Mid(Decrypt(RsSegur!minimo), 3, 2) & "/" & Right(Decrypt(RsSegur!minimo), 4)
    gMaximo = Left(Decrypt(RsSegur!maximo), 2) & "/" & Mid(Decrypt(RsSegur!maximo), 3, 2) & "/" & Right(Decrypt(RsSegur!maximo), 4)
    Exit Sub
    
Erro_sLe_Tab_Balanca:
    If Err.Number = -2147217887 Then
        Call sMostraErro("sLe_Tab_Balanca", "500", "Favor Entrar com contato com a Novavia")
    Else
        Call sMostraErro("sLe_Tab_Balanca", Err.Number, Err.Description)
    End If
    Call Fecha_Formularios
    End
    
End Sub

'*********************************************
Public Sub sTrataLocal()
'*********************************************
    Dim strSql As String
    Dim strMinimo As String
    Dim strMaximo As String
    Dim strCNPJ As String
    
    On Error GoTo Erro_sTrataLocal
'
Volta_Validar_Local:

    Call sLeDataServidorLocal
    
    Call sLe_Tab_Balanca
        
    If RsSegur.RecordCount = 0 Then
        '**** Rotina que mostra a tela para entrar em contato com a Novavia e pedir uma senha
        '**** Código de erro 100
        Call sMostraAviso("Erro", "Falta tabela essencial ao funcionamento do sistema", "Contate a Novavia")
    End If
    
    '**** Data do servidor menor que a data de ultimo uso do sistema
    '**** Atualiza Status para 1, travando e Mostra tela de solicitação de senha
    If CDate(gDataServidor) < CDate(gMinimo) Then
        Call sConectaBanco
        strSql = "UPDATE TAB_BALANCA SET MODELO = '" & Encrypt("1") & "'"
        Cnn.Execute strSql
        Cnn.CommitTrans
        Call sLibera(310)
        GoTo Volta_Validar_Local
    End If

    '**** Data do servidor maior que a data limite para uso do sistema
    '**** Verifica se tem WEB e se tiver verifica se está Ativo, caso negativo
    '**** atualiza Status para 1, travando e Mostra tela de solicitação de senha
    If CDate(gDataServidor) > CDate(gMaximo) Then
        'If Not fValidaWEB() Then
            strMinimo = Encrypt(Left(gDataServidor, 2) + Mid(gDataServidor, 4, 2) + Right(gDataServidor, 4))
            Call sConectaBanco
            strSql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "',MODELO = '" & Encrypt("1") & "'"
            Cnn.Execute strSql
            Cnn.CommitTrans
            Call sLibera(320)
            GoTo Volta_Validar_Local
        'End If
    End If
    
    '**** Data de ultima utilização é maior que a data limite para uso do sistema
    '**** Verifica se tem WEB e se tivr verifica se está Ativo, caso negativo
    '**** atualiza Status para 1, travando e Mostra tela de solicitação de senha
    If CDate(gMinimo) > CDate(gMaximo) Then
        'If Not fValidaWEB() Then
            strMinimo = Encrypt(Left(gMinimo, 2) + Mid(gMinimo, 4, 2) + Right(gMinimo, 4))
            Call sConectaBanco
            strSql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "',MAXIMO = '" & strMaximo & "',MODELO = '" & Encrypt("1") & "'"
            Cnn.Execute strSql
            Cnn.CommitTrans
            Call sLibera(330)
            GoTo Volta_Validar_Local
        'End If
    End If
    
    '**** Atualizar o Banco de dados local, campo minimo com a data atual
    strMinimo = Encrypt(Left(gDataServidor, 2) + Mid(gDataServidor, 4, 2) + Right(gDataServidor, 4))
    strSql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "'"
    Call sConectaBanco
    Cnn.Execute strSql
    Cnn.CommitTrans
    Cnn.Close
    Set Cnn = Nothing
    Exit Sub

Erro_sTrataLocal:
    Call sMostraErro("sTrataLocal", Err.Number, Err.Description)
    Call Fecha_Formularios
    End
    
End Sub

'******************************************************
Public Sub sTrataLocal_Sem_Web()
'******************************************************
    Dim strSql As String
    Dim strMinimo As String
'    Dim strMaximo As String
'    Dim strCNPJ As String
    
    On Error GoTo Erro_sTrataLocal_Sem_Web
'
Volta_Validar_Local_Sem_Web:

    Call sLeDataServidorLocal
    gDataServidor = gDataServidorLocal
    
    Call sLe_Tab_Balanca
        
    If RsSegur.RecordCount = 0 Then
        '**** Rotina que mostra a tela para entrar em contato com a Novavia avisando que não tem tab_balanca
        Call sMostraAviso("Erro", "Falta tabela essencial ao funcionamento do sistema", "Contate a Novavia")
        Call Fecha_Formularios
        End
    End If
    
    '**** Data do servidor menor que a data de ultimo uso do sistema
    '**** Atualiza Status para 1, travando e Mostra tela de solicitação de senha
    If CDate(gDataServidor) < CDate(gMinimo) Then
        Call sConectaBanco
        strSql = "UPDATE TAB_BALANCA SET MODELO = '" & Encrypt("1") & "'"
        Cnn.Execute strSql
        Cnn.CommitTrans
        Call sLibera(310)
        GoTo Volta_Validar_Local_Sem_Web
    End If

    If Date < CDate(gMinimo) Then
        Call sConectaBanco
        strSql = "UPDATE TAB_BALANCA SET MODELO = '" & Encrypt("1") & "'"
        Cnn.Execute strSql
        Cnn.CommitTrans
        Call sLibera(310)
        GoTo Volta_Validar_Local_Sem_Web
    End If

    '**** Data do servidor maior que a data limite para uso do sistema
    '**** Verifica se tem WEB e se tiver verifica se está Ativo, caso negativo
    '**** atualiza Status para 1, travando e Mostra tela de solicitação de senha
    If CDate(gDataServidor) > CDate(gMaximo) Then
            strMinimo = Encrypt(Left(gDataServidor, 2) + Mid(gDataServidor, 4, 2) + Right(gDataServidor, 4))
            Call sConectaBanco
            strSql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "',MODELO = '" & Encrypt("1") & "'"
            Cnn.Execute strSql
            Cnn.CommitTrans
            Call sLibera(320)
            GoTo Volta_Validar_Local_Sem_Web
    End If
    
    '**** Data de ultima utilização é maior que a data limite para uso do sistema
    '**** Verifica se tem WEB e se tivr verifica se está Ativo, caso negativo
    '**** atualiza Status para 1, travando e Mostra tela de solicitação de senha
    If CDate(gMinimo) > CDate(gMaximo) Then
        strMinimo = Encrypt(Left(gMinimo, 2) + Mid(gMinimo, 4, 2) + Right(gMinimo, 4))
        Call sConectaBanco
        strSql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "',MAXIMO = '" & strMaximo & "',MODELO = '" & Encrypt("1") & "'"
        Cnn.Execute strSql
        Cnn.CommitTrans
        Call sLibera(330)
        GoTo Volta_Validar_Local_Sem_Web
    End If
    
    If Date > CDate(gMaximo) Then
        strMinimo = Encrypt(Left(gMinimo, 2) + Mid(gMinimo, 4, 2) + Right(gMinimo, 4))
        Call sConectaBanco
        strSql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "',MAXIMO = '" & strMaximo & "',MODELO = '" & Encrypt("1") & "'"
        Cnn.Execute strSql
        Cnn.CommitTrans
        Call sLibera(330)
        GoTo Volta_Validar_Local_Sem_Web
    End If
    
    '**** Atualizar o Banco de dados local, campo minimo com a data atual
    strMinimo = Encrypt(Left(gDataServidor, 2) + Mid(gDataServidor, 4, 2) + Right(gDataServidor, 4))
    strSql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "'"
    Call sConectaBanco
    Cnn.Execute strSql
    Cnn.CommitTrans
    Cnn.Close
    Set Cnn = Nothing
    Exit Sub

Erro_sTrataLocal_Sem_Web:
    Call sMostraErro("sTrataLocal", Err.Number, Err.Description)
    Call Fecha_Formularios
    End

End Sub

Public Sub sLeDataServidorWEB()
    
    Dim RsWEB As New ADODB.Recordset
    On Error GoTo Erro_sLeDataServidorWeb
    If sConectaWEB = True Then
        'pega data do servidor sql - LocalWEB
        sql = "SELECT  CONVERT(VARCHAR(10), GETDATE(), 103) AS [DD/MM/YYYY]"
        Set RsWEB = New ADODB.Recordset
        RsWEB.Open sql, CnnLocaWEB, 1, 2
        gDataServidorWEB = RsWEB(0)
        RsWEB.Close
        
        If RsWEB.State = adStateOpen Then
            RsWEB.Close
        End If
        If CnnLocaWEB.State = adStateOpen Then
            CnnLocaWEB.Close
            Set CnnLocaWEB = Nothing
        End If
    End If
    Exit Sub

Erro_sLeDataServidorWeb:
    Call sMostraErro("sLeDataServidorWeb", Err.Number, Err.Description)
    Call Fecha_Formularios
    End
    
End Sub

Public Function fTrataWEB()

    Dim RsWEB As New ADODB.Recordset
    
    fTrataWEB = False
    On Error GoTo Erro_fTrataWeb
    
    If Not gTemWEB Then
       Exit Function
    End If
        
    'Call sLe_Tab_Balanca   '*** Ler a tab_balanca e atualizar o CNPJ/CPF
    
    Call sConectaWEB
    
    Set RsWEB = New ADODB.Recordset
    strSql = "SELECT count(*) FROM TB_CLIENTE WHERE CNPJ = '" & FormatCPF_CNPJ(gCNPJ) & "'"
    strSql = strSql & " AND STATUS='A'"
    RsWEB.Open strSql, CnnLocaWEB, 1, 2
        
    If RsWEB(0) = 0 Then
        '*** Cliente está com Status Inativo na Novavia então, atualizar a tabela tab_balanca
        '*** local, campo MAXIMO com a data atual para não permitir acesso sem WEB.
        Call sConectaBanco
        Call sLe_Tab_Balanca
        gCNPJ = Decrypt(RsSegur!peso)
        If RsSegur.State = adStateOpen Then
            RsSegur.Close
        End If
        strMinimo = Encrypt(Left(gDataServidor, 2) + Mid(gDataServidor, 4, 2) + Right(gDataServidor, 4))
        strMaximo = Encrypt(Left(gDataServidor, 2) + Mid(gDataServidor, 4, 2) + Right(gDataServidor, 4))
        sql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "',MAXIMO = '" & strMaximo & "',MODELO = '" & Encrypt("2") & "'"
        Cnn.Execute sql
        Cnn.CommitTrans
        Cnn.Close
        Set Cnn = Nothing
        '*** Mostra a tela de solictação de senha a Novavia
        Call sLibera(210)
    End If
'   '*** Cliente está com Status Ativo na Novavia então, atualizar a tabela tab_balanca
'   '*** local, campo MAXIMO com a data atual mais 30 dias
    Call sConectaBanco
    strMinimo = Encrypt(Left(gDataServidor, 2) + Mid(gDataServidor, 4, 2) + Right(gDataServidor, 4))
    strMaximo = Encrypt(Left(fSomaMes(gDataServidor, 1), 2) & Mid(fSomaMes(gDataServidor, 1), 4, 2) & Right(fSomaMes(gDataServidor, 1), 4))
    gMinimo = Left(gDataServidor, 2) + Mid(gDataServidor, 4, 2) + Right(gDataServidor, 4)
    gMaximo = Left(fSomaMes(gDataServidor, 1), 2) & Mid(fSomaMes(gDataServidor, 1), 4, 2) & Right(fSomaMes(gDataServidor, 1), 4)
    sql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "',MAXIMO = '" & strMaximo & "',MODELO = '" & Encrypt("0") & "'"
    Cnn.Execute sql
    Cnn.CommitTrans
    Cnn.Close
    Set Cnn = Nothing
    
    If RsWEB.State = adStateOpen Then
        RsWEB.Close
    End If
    
    sql = "UPDATE tb_cliente SET DATA_ACESSO = '" & gDataServidor & "'"
    sql = sql & " WHERE  CNPJ = '" & FormatCPF_CNPJ(gCNPJ) & "'"
    CnnLocaWEB.Execute sql
    
    sql = "INSERT INTO tb_Cliente_Log (Data_Acesso,CNPJ ,Empresa,Login,Hostname,IP_TERMINAL) Values ("
    sql = sql & "getdate(),"
    sql = sql & "'" & FormatCPF_CNPJ(gCNPJ) & "',"
    NOME_EMPRESA = Mid(FiltraAspasSimples(GetCampo("select RazaoSocial_Empresa from EMPRESA ", "RazaoSocial_Empresa")), 1, 100)
    If Len(NOME_EMPRESA) > 0 Then
        sql = sql & "'" & NOME_EMPRESA & "',"
    Else
        sql = sql & "NULL,"
    End If
    sql = sql & "'" & sysNomeAcesso & "',"
    sql = sql & "'" & Environ("ComputerName") & "',"
    sql = sql & "'" & STR_IP_COMPUTADOR & "')"
    CnnLocaWEB.Execute sql
    
    CnnLocaWEB.Close
    Set CnnLocaWEB = Nothing
    fTrataWEB = True
    
    Exit Function

Erro_fTrataWeb:
'    If Err.Number = -2147467259 Then
'        'Call sTrataLocal
'        Screen.MousePointer = 1
        Err.Clear
'    Else
'        Screen.MousePointer = 1
'        'Call sMostraAviso("Atenção:", _
'        '                      "Sem acesso ao banco de dados na WEB.", _
'        '                      "Entre em contato com a Novavia ", _
'        '                      "Informe o erro: " & Err.Description, _
'        '                      "Clique em Ok para continuar")
'   '
'        Err.Clear
'        'Call Fecha_Formularios
'        'End
'    End If
End Function

Public Function fSomaMes(cDataInicial As String, nMeses As Integer)
    
'    Dim Months As Double
'    Dim SecondDate As Date
'   StartDate = InputBox("Enter a date")
    fSomaMes = DateAdd("m", nMeses, cDataInicial)
End Function

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
    Call sMostraAviso("Atenção - Erro: ", "Contate a Reinert Informatica informando o erro abaixo:", _
                      "No.erro: " & pErroNumero & " Descr.: " & pErroDesc, _
                      "Módulo do erro: " & pModulo, "Sistema será encerrado")
    'Call Fecha_Formularios
    End
End Sub

Public Function fCentraTexto(pTexto As String, pTam As Double)
    fCentraTexto = Space(Int(pTam / 2) - Int(Len(Trim(pTexto)) / 2)) & Trim(pTexto)
End Function

'*** Fabio Reinert - 04/2017 - As funcoes abaixo foram colocadas por mim pois não foram encontradas no projeto - Inicio
Public Function Encrypt(ByVal icText As String) As String
 Dim icLen As Integer
 Dim icNewText As String
 Dim icChar
 Dim I As Integer
 icChar = ""
    icLen = Len(icText)
    For I = 1 To icLen
        icChar = Mid(icText, I, 1)
        Select Case Asc(icChar)
            Case 65 To 90
                icChar = Chr(Asc(icChar) + 127)
            Case 97 To 122
                icChar = Chr(Asc(icChar) + 121)
            Case 48 To 57
                icChar = Chr(Asc(icChar) + 196)
            Case 32
                icChar = Chr(32)
        End Select
        icNewText = icNewText + icChar
    Next
    Encrypt = icNewText
End Function

Public Function Decrypt(ByVal icText As String) As String
 Dim icLen As Integer
 Dim icNewText As String
 Dim icChar
 Dim I As Integer
 icChar = ""
    icLen = Len(icText)
    For I = 1 To icLen
        icChar = Mid(icText, I, 1)
        Select Case Asc(icChar)
            Case 192 To 217
                icChar = Chr(Asc(icChar) - 127)
            Case 218 To 243
                icChar = Chr(Asc(icChar) - 121)
            Case 244 To 253
                icChar = Chr(Asc(icChar) - 196)
            Case 32
                icChar = Chr(32)
        End Select
        icNewText = icNewText + icChar
    Next
    Decrypt = icNewText
End Function
'*** Fabio Reinert - 04/2017 - As funcoes acima foram colocadas por mim pois não foram encontradas no projeto - Inicio

Public Sub sRestauraSegur()
   Call sConectaBanco
   On Error GoTo erro_SQL
   
   strSql = "SELECT COUNT(*) QTDE FROM RDB$RELATIONS WHERE RDB$FLAGS=1 and RDB$RELATION_NAME='TAB_BALANCA' "
   'ELSE -- Senão se for SQL Server
   'strSql = "SELECT COUNT(*) FROM  INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = " & "'TAB_BALANCA'"
   'END IF
   Call sConectaBanco
   Set Rstemp = New ADODB.Recordset
   Rstemp.Open strSql, Cnn, adOpenForwardOnly, adLockReadOnly
   If Rstemp(0) > 0 Then
       Rstemp.Close
       RsSegur.Close
       Rs.Close
       strSql = "DROP TABLE TAB_BALANCA"
       Cnn.BeginTrans
       Cnn.Execute strSql
       Cnn.CommitTrans
   Else
       'MsgBox "Restauração da segurança já foi efetuada", vbOKOnly, "Aviso"
       'End
   End If
    If Rstemp.State = adStateOpen Then
        Rstemp.Close
    End If
    If Cnn.State = adStateOpen Then
        Cnn.Close
    End If
   
    If Dir("\\servidor\Petshop\fArquivo.ini") <> "" Then
        Kill "\\servidor\Petshop\fArquivo.ini"
    End If

   Exit Sub
   
erro_SQL:
   MsgBox "Favor fechar o sistema de vendas e reexecutar esse programa novamente", vbCritical, "Atenção"
   End

erro_EXE:
   MsgBox "Programa não encontrado", vbCritical, "Atenção"
   Call Fecha_Formularios
   End

End Sub

Public Function fValida_No_Pedido()
    
    On Error GoTo Erro_fValida_No_Pedido
    
    fValida_No_Pedido = True
    
    
    If fValidaWEB = True Then   'Se tem registro Ativo na base de dados da Novavia
        Exit Function           'segue o processo
    End If
        
    'Está inativo na Novavia ou não tem conexão com a WEB ou deu erro na conexão
    '
    Call sConectaBanco     'Vai verificar se na tab_balanca está dentro do período de 30 dias
    
    Set RsSegur = New ADODB.Recordset
    strSql = "SELECT PESO,MINIMO,MAXIMO,MODELO FROM TAB_BALANCA"
    If RsSegur.State = adStateOpen Then
       RsSegur.Close
    End If
    RsSegur.Open strSql, Cnn, 1, 2
    
    'Descriptografa a data atual (MINIMO)
    gMinimo = Left(Decrypt(RsSegur!minimo), 2) & "/" & Mid(Decrypt(RsSegur!minimo), 3, 2) & "/" & Right(Decrypt(RsSegur!minimo), 4)
    
    'Descriptografa a data atual (MAXIMO)
    gMaximo = Left(Decrypt(RsSegur!maximo), 2) & "/" & Mid(Decrypt(RsSegur!maximo), 3, 2) & "/" & Right(Decrypt(RsSegur!maximo), 4)
    
    '**** Data do servidor maior que a data limite para uso do sistema
    '**** Verifica se tem WEB e se tiver verifica se está Ativo, caso negativo
    '**** atualiza Status para 1, travando e Mostra tela de solicitação de senha
    If CDate(gDataServidor) > CDate(gMaximo) Then
        strMinimo = Encrypt(Left(gDataServidor, 2) + Mid(gDataServidor, 4, 2) + Right(gDataServidor, 4))
        Call sConectaBanco
        strSql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "',MODELO = '" & Encrypt("1") & "'"
        Cnn.Execute strSql
        Cnn.CommitTrans
        Call sLibera(333)
        fValida_No_Pedido = True
        Exit Function
    End If
    
    '**** Data de ultima utilização é maior que a data limite para uso do sistema
    '**** Verifica se tem WEB e se tiver verifica se está Ativo, caso negativo
    '**** atualiza Status para 1, travando e Mostra tela de solicitação de senha
    If CDate(gMinimo) > CDate(gMaximo) Then
        strMinimo = Encrypt(Left(gMinimo, 2) + Mid(gMinimo, 4, 2) + Right(gMinimo, 4))
        Call sConectaBanco
        strSql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "',MAXIMO = '" & strMaximo & "',MODELO = '" & Encrypt("1") & "'"
        Cnn.Execute strSql
        Cnn.CommitTrans
        Call sLibera(333)
        fValida_No_Pedido = False
        Exit Function
    End If
    
    '**** Atualizar o Banco de dados local, campo minimo com a data atual
    strMinimo = Encrypt(Left(gDataServidor, 2) + Mid(gDataServidor, 4, 2) + Right(gDataServidor, 4))
    strSql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "'"
    Call sConectaBanco
    Cnn.Execute strSql
    Cnn.CommitTrans
    Cnn.Close
    Set Cnn = Nothing
    Exit Function

Erro_fValida_No_Pedido:
    Call sMostraErro("fValida_No_Pedido", Err.Number, Err.Description)
    Call Fecha_Formularios
    End
    
End Function

Public Function fValidaWEB()
'**** Função de verificação se o cliente está ativo na LocalWEB

    fValidaWEB = False
    On Error GoTo fTrata_Erro_Web
    
    'Call sConectaWEB
    
    Set CnnLocaWEB = New ADODB.Connection
    With CnnLocaWEB
        .CursorLocation = adUseClient
        .Open "Provider=SQLOLEDB.1;Password=billgodes222;Persist Security Info=True;User ID=novavia3;Data Source=sqlserver03.novavia.com.br;Extended Properties=ConnectionTimeout=3;Connect Timeout=2"
    End With

    If RsSegur.State = adStateOpen Then
        RsSegur.Close
    End If
    
    Set RsSegur = New ADODB.Recordset
    strSql = "SELECT count(*) FROM TB_CLIENTE WHERE CNPJ = '" & FormatCPF_CNPJ(gCNPJ) & "'"
    strSql = strSql & " AND STATUS='A'"
    RsSegur.Open strSql, CnnLocaWEB, 1, 2
        
    If RsSegur(0) = 0 Then
        '*** Cliente está com Status Inativo na Novavia então, atualizar a tabela tab_balanca
        '*** local, campo MAXIMO com a data atual para não permitir acesso sem WEB.
        Call sConectaBanco
        strMinimo = Encrypt(Left(gDataServidor, 2) + Mid(gDataServidor, 4, 2) + Right(gDataServidor, 4))
        strMaximo = Encrypt(Left(gDataServidor, 2) + Mid(gDataServidor, 4, 2) + Right(gDataServidor, 4))
        sql = "UPDATE TAB_BALANCA SET MINIMO = '" & strMinimo & "',MAXIMO = '" & strMaximo & "',MODELO = '" & Encrypt("2") & "'"
        Cnn.Execute sql
        Cnn.CommitTrans
        Cnn.Close
        Set Cnn = Nothing
        '*** Mostra a tela de solictação de senha a Novavia
        'Call sLibera(210)
        Exit Function
    End If
    
    If RsSegur.State = adStateOpen Then
        RsSegur.Close
    End If
    
    sql = "UPDATE tb_cliente SET DATA_ACESSO = '" & gDataServidor & "'"
    sql = sql & " WHERE  CNPJ = '" & FormatCPF_CNPJ(gCNPJ) & "'"
    CnnLocaWEB.Execute sql
    
    CnnLocaWEB.Close
    Set CnnLocaWEB = Nothing
    
    fValidaWEB = True
    Exit Function
    
fTrata_Erro_Web:
    Err.Clear
    
End Function

Public Function fcheckInternetConnection() As Boolean
'code to check for internet connection
'by Daniel Isoje

'referenciar no projeto: Microsoft xml 6.0

Screen.MousePointer = 11
On Error Resume Next
 fcheckInternetConnection = False
 'Dim objSvrHTTP As ServerXMLHTTP
 Dim varProjectID, varCatID, strT As String
 'Set objSvrHTTP = New ServerXMLHTTP
 'objSvrHTTP.Open "GET", "http://www.google.com"
 'objSvrHTTP.setRequestHeader "Accept", "application/xml"
 'objSvrHTTP.setRequestHeader "Content-Type", "application/xml"
 'objSvrHTTP.setTimeouts 1000, 1000, 1000, 1000
 'objSvrHTTP.send strT
 'If Err = 0 Then
 '   fcheckInternetConnection = True
 '   gTemWEB = True
 'Else
 '   fcheckInternetConnection = False
 '   gTemWEB = False
 'End If
 Screen.MousePointer = 1
End Function

'*** Fabio Reinert - Funcoes do módulo de segurança - Fim

