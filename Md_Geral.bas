Attribute VB_Name = "Mod_Geral"
'criar arquivo ttf crystal reports
Declare Function CreateFieldDefFile Lib "p2smon.dll" (lpUnk As Object, ByVal Filename As String, ByVal bOverWriteExistingFile As Long) As Long
'exemplo de uso
'Rstemp.Open "select * from clientes", Cnn, adOpenDynamic, adLockOptimistic
'CreateFieldDefFile rs1, App.Path & "\clientes.ttx", 1


'VB6  função API para dar uma pausa na aplicação
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'exemplo uso
'Sleep (1000) ' = 1 segundo

Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

'****************impressora
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrn As Long, pDefault As Any) As Long
Public Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrn As Long, ByVal Level As Long, pDocInfo As DOC_INFO_1) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrn As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Public Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long
Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrn As Long) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
      "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
      lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As _
      String, ByVal nSize As Long, ByVal lpFilename As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
      "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
      lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long

' PARA PEGAR NOME DA MAQUINA
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

''*** Fabio Reinert - 10/2017 - Para pegar a tecla digitada no MDIForm - Inicio
'Public Declare Function GetAsyncKeyState Lib "user32" _
'    (ByVal vKey As Long) As Integer
''*** Fabio Reinert - 10/2017 - Para pegar a tecla digitada no MDIForm - Fim
'
Private Type DOC_INFO_1
   pDocName As String
   pOutputFile As String
   pDatatype As String
End Type

'****************************************************
'variaveis de conexao
Public Cnn As New ADODB.Connection
Public cmd As New ADODB.Command
Public mobjCmd       As ADODB.Command
Public dbCepBrasil As New ADODB.Connection


'*************************
'variaveis para recordsets
Public Rstemp       As New ADODB.Recordset
Public RsTemp1      As New ADODB.Recordset
Public Rstemp2      As New ADODB.Recordset
Public Rstemp3      As New ADODB.Recordset
Public Rstemp5      As New ADODB.Recordset
Public Rstemp6      As New ADODB.Recordset
Public Rs           As New ADODB.Recordset
Public tmpRecordset As New ADODB.Recordset
Public Rstemporario As New ADODB.Recordset


'variaveis pra controle de registro
Global Situacao_Registro As String
Global Dias_Uso_Sistema As Integer
Global ConsultaProd_Ped As Integer
Global flagConsultaPedProd As Boolean

Public Impressoras As Printer
Public gTransacao As Boolean
Public sql  As String
Public tmpSQL As String
Public sysNome As String
Public sysCodigo As String
Public sysNomeAcesso As String
Public sysAcesso As String
Public sysSenha As String
Public tipo_PedidoOrcamento As Integer
Public flag_Relogin As Boolean
Public PstrCP As String

Public gMensagem As String
Public strSql  As String
Public strSql1 As String
Public strSql2 As String
Public strSql3 As String
Public strPesqProdProv As Boolean
Public strFormaPgto As String

Public strDesc          As String
Public intCod           As Integer
Public varAux           As Byte     '1 para tela de Saidas ou 2 para tela de Orçamentos
Public LogAdmin         As Boolean  'Para verificar a senha do Admin
Dim Disk As String

Global Acesso_OK As Integer              'variavel para identificar se a tela de senha ja foi aberta
Global verifica_abertura As Integer      'variavel para identificar se o executavel ja esta aberto
Global NomeSistema As String
Global tipo As String      'Inclusao(I),Altercao(A),Exclusao(E)
Global valida_digito As Integer
Global Tela As Integer
Global tipo_financeiro As String
Global Consultas As String

Global flag_Gaveta_Bematec As Boolean
Global flag_Gaveta_Elgin As Boolean

Global Flag_Exclui_Pedidos As Boolean
Global Flag_Libera_Qtde_Estoque As Boolean

Global SENHA_EXCLUIU_PEDIDO As String

Global NomeUsuario As String
Global CodUsuarioLogado As Integer
Global UsuarioExcluiPedido As Boolean

Global Consulta As String
Global NomeEmpresa As String
Global EnderecoEmpresa As String
Global CGC_EMPRESA, CEP_EMPRESA As String

Global Fone1Empresa, Fone2Empresa As String
Global PercICMSEstado As String ' recebe valor % ICMS pos estado menu Config.
Global StrNomeMaquina As String

Global emailEmpresa As String

Global flagComEstoque As Boolean
Global flagAltPrVenda As Boolean
Global flagDescPedOrc As Boolean
Global flagImpFiscalSelecionada As Boolean
Global flagCursorCodigo As Boolean
Global flagQtde1 As Boolean
Global flagEmitir_Orcamentos

Global Versao_Software As String

Public frmChamou As Boolean
Global IP_Servidor As String
Global IP_Servidor_Relatorios As String

Global Flag_Destoque_Por_Grupo As Boolean

'codigo de barras usado no pdv
Public FarTopMargin         As Single
Public FarRightMargin       As Single
Public MaxRightMargin       As Single
Public MaxLeftMargin        As Single
Public DocumentLayout       As Single
Public PrinterTop           As Currency
Public PrinterBottom        As Currency

Global retImpFiscal As Integer '0=nenhuma, 1=bematec, 2=daruma 3=sweda 5=sat


Private FormX, FormY As Integer

'No Declarations:
Private Declare Function CreateRoundRectRgn Lib _
       "gdi32" (ByVal X1 As Long, ByVal Y1 As _
       Long, ByVal X2 As Long, ByVal Y2 As Long, _
       ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" _
       (ByVal hwnd As Long, ByVal hRgn As Long, _
       ByVal bRedraw As Boolean) As Long
Private Declare Function GetClientRect Lib "user32" _
       (ByVal hwnd As Long, lpRect As Rect) As Long
Private Type Rect
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type

'usa gaveta bematec não fiscal
Global flag_Usa_gaveta_Bematec_MP4000 As Boolean

'usa gaveta bematec MP2500/EPSON não fiscal
Global flag_Usa_gaveta_Bematec_MP2500 As Boolean

Global Altera_Estoque As Boolean

Global flag_Desc_prod_cupom As Boolean
'
'*************************************************************************************
'*** Fabio Reinert ( Alemao) 06/2017 - Inclusão de captura de IP do cliente - Inicio *
'*************************************************************************************
'
Public STR_IP_COMPUTADOR As String
'
'*** Fabio Reinert - 10/2017 - Inclusão de variaveis para autocompletar o combobox - Inicio
'
#If Win32 Then
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, _
         ByVal wParam As Long, lParam As Any) As Long
#Else
    Declare Function SendMessage Lib "User" _
        (ByVal hwnd As Integer, ByVal wMsg As Integer, _
         ByVal wParam As Integer, lParam As Any) As Long
#End If
'
'*** Fabio Reinert - 10/2017 - Inclusão de variaveis para autocompletar o combobox - Fim
'
'
'**************************************************************************************************************
'* Fabio Reinert (Alemao) - 08/2017 - Arrays p/telas de recebimento debito,credito,boletos e cheques - Inicio *
'**************************************************************************************************************
'*
Public aDebitos() As String
Public aCreditos() As String
Public aBoletos() As String
Public aChequesPre() As String
'
'************************************************************************************************************
'* Fabio Reinert (Alemao) - 08/2017 - Arrays p/telas de recebimento debito,credito,boletos e cheques - Fim  *
'************************************************************************************************************
'*

Public Function BuscaIP() As String
Dim NIC As Variant
Dim NICs As Object

On Error GoTo errError

Set NICs = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

For Each NIC In NICs
   If NIC.IPEnabled Then
        BuscaIP = NIC.IpAddress(0)
    End If
Next NIC

'ou
'Dim IPConfig As Variant
'Dim IPConfigSet As Object
'Set IPConfigSet = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE")
'
'For Each IPConfig In IPConfigSet
' If Not IsNull(IPConfig.IPAddress) Then MsgBox IPConfig.IPAddress(0), vbInformation
'Next IPConfig

Exit Function
    
errError:
    
    If Err.Number <> 0 Then
        Err.Clear
    End If
    BuscaIP = ""

End Function
'
'*************************************************************************************
'*** Fabio Reinert ( Alemao) 06/2017 - Inclusão de captura de IP do cliente - Fim    *
'*************************************************************************************
'
Sub Retangulo(m_hWnd As Long, Fator As Byte)
    Dim RGN As Long
    Dim RC As Rect
    
    Call GetClientRect(m_hWnd, RC)
    RGN = CreateRoundRectRgn(RC.Left, RC.Top, RC.Right, RC.Bottom, Fator, Fator)
    SetWindowRgn m_hWnd, RGN, True
    
End Sub
Function DVEAN(vNr As String, Optional vAlerta As Boolean = False) As Boolean
On Error GoTo Erro

Dim i As Integer, vSoma As Long, vMult As Byte, vDV As String

If Len(vNr) Mod 2 = 0 Then vMult = 3 Else vMult = 1

For i = 1 To Len(vNr) - 1
    vSoma = vSoma + CInt(Mid(vNr, i, 1)) * vMult
    If vMult = 1 Then vMult = 3 Else vMult = 1
Next

vDV = IIf(vSoma Mod 10 = 0, 0, ((Int(vSoma / 10) + 1) * 10) - vSoma)

DVEAN = (vDV = Right(vNr, 1))

'If vAlerta = True Then If DVEAN Then MsgBox "Digito verificador válido.", vbInformation, "Código válido" Else MsgBox "Digito verificador INVÁLIDO.", vbExclamation, "Código INVÁLIDO"

'If DVEAN Then MsgBox "Digito verificador válido.", vbInformation, "Código válido" Else MsgBox "Digito verificador INVÁLIDO.", vbExclamation, "Código INVÁLIDO"

Sair:
Exit Function
Erro:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Erro na Função DVEAN"
Resume Sair
End Function

Public Function CalculaImpostoIBPT(ByVal NCM As String, ByVal vlr_Tot_Item As String) As String
        sql = "SELECT ALIQ_NAC FROM ALIQUOTAS_IBPT WHERE EX IS NULL "
        sql = sql & " AND CODIGO = " & NCM
        sql = sql & " AND TABELA=0"
        
        'strAliquotaVtotTrib = Replace(GetCampo(sql, "ALIQ_NAC"), ".", ",")
        
        Set Rstemp5 = New ADODB.Recordset
        
        Rstemp5.Open sql, Cnn, 1, 2
        If Rstemp5.RecordCount > 0 Then
            strAliquotaVtotTrib = Replace(Rstemp5!ALIQ_NAC, ".", ",")
        Else
            strAliquotaVtotTrib = "0"
        End If
        Rstemp5.Close
        Set Rstemp5 = Nothing
        
        If IsNumeric(strAliquotaVtotTrib) Then
            CalculaImpostoIBPT = Format((CCur(vlr_Tot_Item) * CCur(strAliquotaVtotTrib)) / 100, "0.00")
        Else
            CalculaImpostoIBPT = "0,00"        'vTotTrib
        End If
End Function


'Dicas de Visual basic, Microsoft sql server vbnet, desenvolvimento, ASP, ASP NET
'INSCRICAO ESTADUAL

Public Function ValidaInscrEstadual(pUF As String, pInscr As String) As Boolean
   
   ValidaInscrEstadual = False
   
   Dim strBase              As String
   Dim strBase2             As String
   Dim strOrigem            As String
   Dim strDigito1           As String
   Dim strDigito2           As String
   Dim intPos               As Integer
   Dim intValor             As Integer
   Dim intSoma              As Integer
   Dim intResto             As Integer
   Dim intNumero            As Integer
   Dim intPeso              As Integer
   Dim intDig               As Integer
   
   strBase = ""
   strBase2 = ""
   strOrigem = ""
   If Trim(pInscr) = "ISENTO" Then
       ValidaInscrEstadual = True
       Exit Function
   End If
   For intPos = 1 To Len(Trim(pInscr))
        If InStr(1, "0123456789P", Mid$(pInscr, intPos, 1), vbTextCompare) > 0 Then
            strOrigem = strOrigem & Mid$(pInscr, intPos, 1)
        End If
   Next
   Select Case pUF
     Case "AC"    ' Acre
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "01" And Mid$(strBase, 3, 2) <> "00" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ValidaInscrEstadual = True
              End If
          End If
     Case "AL"    ' Alagoas
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "24" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intSoma = intSoma * 10
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto = 10, "0", str(intResto)), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ValidaInscrEstadual = True
              End If
          End If
     Case "AM"    ' Amazonas
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          If intSoma < 11 Then
              strDigito1 = Right(str(11 - intSoma), 1)
          Else
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
          End If
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "AP"    ' Amapa
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intPeso = 0
          intDig = 0
          If Left(strBase, 2) = "03" Then
              intNumero = Val(Left(strBase, 8))
              If intNumero >= 3000001 And _
                 intNumero <= 3017000 Then
                  intPeso = 5
                  intDig = 0
              ElseIf intNumero >= 3017001 And _
                     intNumero <= 3019022 Then
                  intPeso = 9
                  intDig = 1
              ElseIf intNumero >= 3019023 Then
                  intPeso = 0
                  intDig = 0
              End If
              intSoma = intPeso
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              intValor = 11 - intResto
              If intValor = 10 Then
                  intValor = 0
              ElseIf intValor = 11 Then
                  intValor = intDig
              End If
              strDigito1 = Right(str(intValor), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ValidaInscrEstadual = True
              End If
          End If
     Case "BA"    ' Bahia
          strBase = Left(Trim(strOrigem) & "00000000", 8)
          If InStr(1, "0123458", Left(strBase, 1), vbTextCompare) > 0 Then
              intSoma = 0
              For intPos = 1 To 6
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (8 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 10
              strDigito2 = Right(IIf(intResto = 0, "0", str(10 - intResto)), 1)
              strBase2 = Left(strBase, 6) & strDigito2
              intSoma = 0
              For intPos = 1 To 7
                   intValor = Val(Mid$(strBase2, intPos, 1))
                   intValor = intValor * (9 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 10
              strDigito1 = Right(IIf(intResto = 0, "0", str(10 - intResto)), 1)
          Else
              intSoma = 0
              For intPos = 1 To 6
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (8 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito2 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
              strBase2 = Left(strBase, 6) & strDigito2
              intSoma = 0
              For intPos = 1 To 7
                   intValor = Val(Mid$(strBase2, intPos, 1))
                   intValor = intValor * (9 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
          End If
          strBase2 = Left(strBase, 6) & strDigito1 & strDigito2
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "CE"    ' Ceara
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          intValor = 11 - intResto
          If intValor > 9 Then
              intValor = 0
          End If
          strDigito1 = Right(str(intValor), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "DF"    ' Distrito Federal
          strBase = Left(Trim(strOrigem) & "0000000000000", 13)
          If Left(strBase, 3) = "073" Then
              intSoma = 0
              intPeso = 2
              For intPos = 11 To 1 Step -1
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * intPeso
                   intSoma = intSoma + intValor
                   intPeso = intPeso + 1
                   If intPeso > 9 Then
                       intPeso = 2
                   End If
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
              strBase2 = Left(strBase, 11) & strDigito1
              intSoma = 0
              intPeso = 2
              For intPos = 12 To 1 Step -1
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * intPeso
                   intSoma = intSoma + intValor
                   intPeso = intPeso + 1
                   If intPeso > 9 Then
                       intPeso = 2
                   End If
              Next
              intResto = intSoma Mod 11
              strDigito2 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
              strBase2 = Left(strBase, 12) & strDigito2
              If strBase2 = strOrigem Then
                  ValidaInscrEstadual = True
              End If
          End If
     Case "ES"    ' Espirito Santo
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "GO"    ' Goias
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If InStr(1, "10,11,15", Left(strBase, 2), vbTextCompare) > 0 Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              If intResto = 0 Then
                  strDigito1 = "0"
              ElseIf intResto = 1 Then
                  intNumero = Val(Left(strBase, 8))
                  strDigito1 = Right(IIf(intNumero >= 10103105 And intNumero <= 10119997, "1", "0"), 1)
              Else
                  strDigito1 = Right(str(11 - intResto), 1)
              End If
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ValidaInscrEstadual = True
              End If
          End If
     Case "MA"    ' Maranhão
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "12" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ValidaInscrEstadual = True
              End If
          End If
     Case "MT"    ' Mato Grosso
          strBase = Left(Trim(strOrigem) & "0000000000", 10)
          intSoma = 0
          intPeso = 2
          For intPos = 10 To 1 Step -1
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * intPeso
               intSoma = intSoma + intValor
               intPeso = intPeso + 1
               If intPeso > 9 Then
                   intPeso = 2
               End If
          Next
          intResto = intSoma Mod 11
          strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
          strBase2 = Left(strBase, 10) & strDigito1
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "MS"    ' Mato Grosso do Sul
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "28" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ValidaInscrEstadual = True
              End If
          End If
     Case "MG"    ' Minas Gerais
          strBase = Left(Trim(strOrigem) & "0000000000000", 13)
          strBase2 = Left(strBase, 3) & "0" & Mid$(strBase, 4, 8)
          intNumero = 2
          For intPos = 1 To 12
               intValor = Val(Mid$(strBase2, intPos, 1))
               intNumero = IIf(intNumero = 2, 1, 2)
               intValor = intValor * intNumero
               If intValor > 9 Then
                   strDigito1 = Format(intValor, "00")
                   intValor = Val(Left(strDigito1, 1)) + _
                              Val(Right(strDigito1, 1))
               End If
               intSoma = intSoma + intValor
          Next
          intValor = intSoma
          While Right(Format(intValor, "000"), 1) <> "0"
              intValor = intValor + 1
          Wend
          strDigito1 = Right(Format(intValor - intSoma, "00"), 1)
          strBase2 = Left(strBase, 11) & strDigito1
          intSoma = 0
          intPeso = 2
          For intPos = 12 To 1 Step -1
               intValor = Val(Mid$(strBase2, intPos, 1))
               intValor = intValor * intPeso
               intSoma = intSoma + intValor
               intPeso = intPeso + 1
               If intPeso > 11 Then
                   intPeso = 2
               End If
          Next
          intResto = intSoma Mod 11
          strDigito2 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
          strBase2 = strBase2 & strDigito2
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "PA"    ' Para
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "15" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ValidaInscrEstadual = True
              End If
          End If
     Case "PB"    ' Paraiba
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          intValor = 11 - intResto
          If intValor > 9 Then
              intValor = 0
          End If
          strDigito1 = Right(str(intValor), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "PE"    ' Pernambuco
          strBase = Left(Trim(strOrigem) & "00000000000000", 14)
          intSoma = 0
          intPeso = 2
          For intPos = 13 To 1 Step -1
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * intPeso
               intSoma = intSoma + intValor
               intPeso = intPeso + 1
               If intPeso > 9 Then
                   intPeso = 2
               End If
          Next
          intResto = intSoma Mod 11
          intValor = 11 - intResto
          If intValor > 9 Then
              intValor = intValor - 10
          End If
          strDigito1 = Right(str(intValor), 1)
          strBase2 = Left(strBase, 13) & strDigito1
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "PI"    ' Piaui
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "PR"    ' Parana
          strBase = Left(Trim(strOrigem) & "0000000000", 10)
          intSoma = 0
          intPeso = 2
          For intPos = 8 To 1 Step -1
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * intPeso
               intSoma = intSoma + intValor
               intPeso = intPeso + 1
               If intPeso > 7 Then
                   intPeso = 2
               End If
          Next
          intResto = intSoma Mod 11
          strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          intSoma = 0
          intPeso = 2
          For intPos = 9 To 1 Step -1
               intValor = Val(Mid$(strBase2, intPos, 1))
               intValor = intValor * intPeso
               intSoma = intSoma + intValor
               intPeso = intPeso + 1
               If intPeso > 7 Then
                   intPeso = 2
               End If
          Next
          intResto = intSoma Mod 11
          strDigito2 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
          strBase2 = strBase2 & strDigito2
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "RJ"    ' Rio de Janeiro
          strBase = Left(Trim(strOrigem) & "00000000", 8)
          intSoma = 0
          intPeso = 2
          For intPos = 7 To 1 Step -1
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * intPeso
               intSoma = intSoma + intValor
               intPeso = intPeso + 1
               If intPeso > 7 Then
                   intPeso = 2
               End If
          Next
          intResto = intSoma Mod 11
          strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
          strBase2 = Left(strBase, 7) & strDigito1
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "RN"    ' Rio Grande do Norte
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "20" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intSoma = intSoma * 10
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto > 9, "0", str(intResto)), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ValidaInscrEstadual = True
              End If
          End If
     Case "RO"    ' Rondonia
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          strBase2 = Mid$(strBase, 4, 5)
          intSoma = 0
          For intPos = 1 To 5
               intValor = Val(Mid$(strBase2, intPos, 1))
               intValor = intValor * (7 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          intValor = 11 - intResto
          If intValor > 9 Then
              intValor = intValor - 10
          End If
          strDigito1 = Right(str(intValor), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "RR"    ' Roraima
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "24" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 9
              strDigito1 = Right(str(intResto), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ValidaInscrEstadual = True
              End If
          End If
     Case "RS"    ' Rio Grande do Sul
          strBase = Left(Trim(strOrigem) & "0000000000", 10)
          intNumero = Val(Left(strBase, 3))
          If intNumero > 0 And intNumero < 468 Then
              intSoma = 0
              intPeso = 2
              For intPos = 9 To 1 Step -1
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * intPeso
                   intSoma = intSoma + intValor
                   intPeso = intPeso + 1
                   If intPeso > 9 Then
                       intPeso = 2
                   End If
              Next
              intResto = intSoma Mod 11
              intValor = 11 - intResto
              If intValor > 9 Then
                  intValor = 0
              End If
              strDigito1 = Right(str(intValor), 1)
              strBase2 = Left(strBase, 9) & strDigito1
              If strBase2 = strOrigem Then
                  ValidaInscrEstadual = True
              End If
          End If
     Case "SC"    ' Santa Catarina
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "SE"    ' Sergipe
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          intValor = 11 - intResto
          If intValor > 9 Then
              intValor = 0
          End If
          strDigito1 = Right(str(intValor), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "SP"    ' São Paulo
          If Left(strOrigem, 1) = "P" Then
              strBase = Left(Trim(strOrigem) & "0000000000000", 13)
              strBase2 = Mid$(strBase, 2, 8)
              intSoma = 0
              intPeso = 1
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * intPeso
                   intSoma = intSoma + intValor
                   intPeso = intPeso + 1
                   If intPeso = 2 Then
                       intPeso = 3
                   End If
                   If intPeso = 9 Then
                       intPeso = 10
                   End If
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(str(intResto), 1)
              strBase2 = Left(strBase, 8) & strDigito1 & Mid$(strBase, 11, 3)
          Else
              strBase = Left(Trim(strOrigem) & "000000000000", 12)
              intSoma = 0
              intPeso = 1
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * intPeso
                   intSoma = intSoma + intValor
                   intPeso = intPeso + 1
                   If intPeso = 2 Then
                       intPeso = 3
                   End If
                   If intPeso = 9 Then
                       intPeso = 10
                   End If
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(str(intResto), 1)
              strBase2 = Left(strBase, 8) & strDigito1 & Mid$(strBase, 10, 2)
              intSoma = 0
              intPeso = 2
              For intPos = 11 To 1 Step -1
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * intPeso
                   intSoma = intSoma + intValor
                   intPeso = intPeso + 1
                   If intPeso > 10 Then
                       intPeso = 2
                   End If
              Next
              intResto = intSoma Mod 11
              strDigito2 = Right(str(intResto), 1)
              strBase2 = strBase2 & strDigito2
          End If
          If strBase2 = strOrigem Then
              ValidaInscrEstadual = True
          End If
     Case "TO"    ' Tocantins
          strBase = Left(Trim(strOrigem) & "00000000000", 11)
          If InStr(1, "01,02,03,99", Mid$(strBase, 3, 2), vbTextCompare) > 0 Then
              strBase2 = Left(strBase, 2) & Mid$(strBase, 5, 6)
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase2, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", str(11 - intResto)), 1)
              strBase2 = Left(strBase, 10) & strDigito1
              If strBase2 = strOrigem Then
                  ValidaInscrEstadual = True
              End If
          End If
   End Select
   
End Function







Public Function VerificaAtualizacoesNFE()

    On Error GoTo TRATA_TABELA_LOTE_NFE
    sql = "SELECT COUNT(*) FROM LOTE_NFE "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    On Error GoTo TRATA_DATA_EMISSAO_NFE
    sql = "SELECT DATA_EMISSAO FROM NFE "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
   
    'DEVOLUCAO_NFE
    On Error GoTo TRATA_DEVOLUCAO_NFE
    sql = "SELECT * FROM DEVOLUCAO_NFE "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
CONTINUA_TRATA_CAMPOS_NOVOS_EMPRESA:
    
    On Error GoTo TRATA_CAMPOS_NOVOS_EMPRESA
    sql = "SELECT NOME_FANTASIA FROM EMPRESA "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
        
    'NOVA VIEW VIEW_LISTA_PROD
    On Error GoTo TRATA_VIEW_LISTA_PROD
    sql = "SELECT * FROM VIEW_LISTA_PROD "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    Rstemp6.Close
    Set Rstemp6 = Nothing
    

TRATA_TABELA_LOTE_NFE:
    If Err.Number <> 0 Then
        Set Rstemp6 = Nothing
        sql = "CREATE TABLE LOTE_NFE (ID INTEGER NOT NULL, LOTE DOUBLE PRECISION, NRO_RECIBO VARCHAR(20), PRIMARY KEY (ID))"
        Cnn.Execute sql
        
        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_LOT_ID1 "
        Cnn.Execute sql
        sql = "SET GENERATOR GEN_LOT_ID1 TO 1"
        Cnn.Execute sql
        
        'cria TRIGGER PARA AUTONUMERADOR
        sql = " CREATE TRIGGER LOTE_NFE_BI FOR LOTE_NFE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_LOT_ID1, 1); END  "
        Cnn.Execute sql
        
        sql = "INSERT INTO LOTE_NFE (LOTE)"
        sql = sql & "values (1)"
        Cnn.Execute sql
        
        'TABELA NFE
        '''sql = "CREATE TABLE NFE (ID INTEGER NOT NULL, CHAVE_NFE VARCHAR(50), NRO_LOTE VARCHAR(100), NRO_RECIBO VARCHAR(20), NRO_PROTOCOLO VARCHAR(20), NRO_PEDIDO DOUBLE PRECISION, NRO_NF DOUBLE PRECISION, NRO_CANCELAMENTO_NF VARCHAR(20), STATUS VARCHAR(20), PRIMARY KEY (ID))"
        sql = "CREATE TABLE NFE (ID INTEGER NOT NULL, CHAVE_NFE VARCHAR(50), NRO_LOTE VARCHAR(100), NRO_RECIBO VARCHAR(20), NRO_PROTOCOLO VARCHAR(20), NRO_PEDIDO BLOB SUB_TYPE 1, NRO_NF DOUBLE PRECISION, NRO_CANCELAMENTO_NF VARCHAR(20), STATUS VARCHAR(20), DATA_EMISSAO DATE, TOTAL_NF DOUBLE PRECISION, PRIMARY KEY (ID))"
        Cnn.Execute sql
        
        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_NFE_ID1 "
        Cnn.Execute sql
        sql = "SET GENERATOR GEN_NFE_ID1 TO 0"
        Cnn.Execute sql
        
        'cria TRIGGER PARA AUTONUMERADOR
        sql = " CREATE TRIGGER NFE_BI FOR NFE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_NFE_ID1, 1); END  "
        Cnn.Execute sql
        
'        sql = "ALTER TABLE NFE ADD FOREIGN KEY (NRO_PEDIDO) REFERENCES SAIDAS_PRODUTO (SEQUENCIA)"
'        Cnn.Execute sql
    End If
    
TRATA_DATA_EMISSAO_NFE:
    If Err.Number <> 0 Then
        sql = "ALTER TABLE NFE ADD DATA_EMISSAO DATE "
'        Cnn.Execute sql
    End If
    
    
TRATA_DEVOLUCAO_NFE:
    If Err.Number <> 0 Then
        
''        sql = "CREATE TABLE DEVOLUCAO_NFE (ID INTEGER NOT NULL, SEQUENCIA DOUBLE PRECISION NOT NULL,DATA_NF DATE,NF DOUBLE PRECISION,"
''        sql = sql & " COD_FORNECEDOR DOUBLE PRECISION,TOTAL_SAIDA DOUBLE PRECISION,PRIMARY KEY (ID))"
''        Cnn.Execute sql
''
''        sql = "CREATE TABLE ITENS_DEVOLUCAO_NFE (ID INTEGER NOT NULL, SEQUENCIA DOUBLE PRECISION, CODIGO_PRODUTO DOUBLE PRECISION, QTDE DOUBLE PRECISION,"
''        sql = sql & "VALOR_UNITARIO DOUBLE PRECISION,VALOR_TOTAL DOUBLE PRECISION, PRIMARY KEY (ID))"
'''        Cnn.Execute sql
''
''        'sql = "ALTER TABLE ITENS_DEVOLUCAO_NFE ADD FOREIGN KEY (SEQUENCIA) REFERENCES DEVOLUCAO_NFE (SEQUENCIA)"
''        'Cnn.Execute sql
''
''
''        ''''''''''''''''''''
''         'Primary Keys     SITE REFEENCIA = "http://www.firebirdsql.org/dotnetfirebird/create-a-new-database-from-an-sql-script.html"
''        'sql = "ALTER TABLE DEVOLUCAO_NFE ADD PRIMARY KEY (ID)"
''        'Cnn.Execute sql
''
''        ' Indices
''        sql = "CREATE INDEX ID_X ON DEVOLUCAO_NFE (ID)"
''        Cnn.Execute sql
''
''        'cria GENERATOR
''        sql = "CREATE GENERATOR GEN_SEQ_DEVOLUCAO_NFE "
''        Cnn.Execute sql
''
''        sql = "SET GENERATOR GEN_SEQ_DEVOLUCAO_NFE TO 0"
''        Cnn.Execute sql
''
''         'cria TRIGGER PARA AUTONUMERADOR
''        sql = " CREATE TRIGGER DEVOLUCAO_NFE FOR DEVOLUCAO_NFE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
''        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_SEQ_DEVOLUCAO_NFE, 1); END  "
''        Cnn.Execute sql
''
''
''        'ITENS_DEVOLUCAO_NFE
''        sql = "CREATE INDEX ID_Y ON ITENS_DEVOLUCAO_NFE (ID)"
''        Cnn.Execute sql
''
''        'cria GENERATOR
''        sql = "CREATE GENERATOR GEN_SEQ_ITENS_DEVOLUCAO_NFE "
''        Cnn.Execute sql
''
''        sql = "SET GENERATOR GEN_SEQ_ITENS_DEVOLUCAO_NFE TO 0"
''        Cnn.Execute sql
''
''         'cria TRIGGER PARA AUTONUMERADOR
''        sql = " CREATE TRIGGER ITENS_DEVOLUCAO_NFE FOR ITENS_DEVOLUCAO_NFE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
''        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_SEQ_ITENS_DEVOLUCAO_NFE, 1); END  "
''        Cnn.Execute sql
''
'''        sql = "SELECT * FROM CAD_MENUS WHERE MENU_CD_CODI = 136 "
'''        Set Rstemp6 = New ADODB.Recordset
'''        Rstemp6.Open sql, Cnn, 1, 2
'''        If Rstemp6.RecordCount = 0 Then
'''            sql = "INSERT INTO CAD_MENUS VALUES("
'''            sql = sql & "'SisAdven',"
'''            sql = sql & "136,"
'''            Menu = UCase("menu_movimentacao_nfe_Devolucao")
'''            sql = sql & "'" & Menu & "',"
'''            sql = sql & "'Movimentação - NFe Devolução')"
'            frmMenu.menu_movimentacao_nfe_Devolucao.Enabled = True
'            Cnn.Execute sql
'        End If
'
'        Rstemp6.Close
'        Set Rstemp6 = Nothing
        GoSub CONTINUA_TRATA_CAMPOS_NOVOS_EMPRESA:
        
    End If
    
    
TRATA_CAMPOS_NOVOS_EMPRESA:
    If Err.Number <> 0 Then
                
        sql = "ALTER TABLE EMPRESA ADD NOME_FANTASIA VARCHAR(30), ADD INSC_ESTADUAL VARCHAR(19), ADD NRO_ENDERECO VARCHAR(9)"
        Cnn.Execute sql
        Call VerificaAtualizacoesNFE
    End If
    
    
TRATA_VIEW_LISTA_PROD:
    If Err.Number <> 0 Then
        sql = "CREATE VIEW VIEW_LISTA_PROD (CODIGO,CODIGO_INTERNO,DESCRICAO,VLRCUSTO,PRECO,UNIDADE,SALDO_EM_ESTOQUE, "
        sql = sql & "MARCA,ULTIMA_VENDA,ULTIMA_COMPRA) AS "
        sql = sql & " select A.CODIGO,A.CODIGO_INTERNO,A.DESCRICAO,A.VLRCUSTO, A.PRECO,A.UNIDADE,B.SALDO_EM_ESTOQUE,M.DESCRICAO AS MARCA, A.ULTIMA_VENDA,"
        sql = sql & " A.ULTIMA_COMPRA FROM PRODUTO A, ESTOQUE B, MARCAS M  WHERE A.CODIGO = B.CODIGO_PRODUTO"
        sql = sql & " AND M.CODIGO = A.MARCA  ORDER BY A.DESCRICAO ASC "
        Cnn.Execute sql
        
        Cnn.Close
        Set Cnn = Nothing
        Call Conecta_Banco
    End If
 
End Function


Public Function VerificaAtualizacoes_CST()

    ' VERIFICA SE O CAMPO EXISTE SE NAO CRIA
    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'CHEQUES' AND  rdb$field_name='NRO_CHEQUES'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "ALTER TABLE CHEQUES ADD NRO_CHEQUES BLOB SUB_TYPE 1"
        Cnn.Execute sql
        Cnn.CommitTrans
    End If
    Rstemp6.Close
    Set Rstemp6 = Nothing

    sql = "SELECT  COUNT(*) FROM RDB$RELATIONS WHERE  RDB$RELATION_NAME = 'MUNICIPIOS'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
'        sql = "CREATE TABLE MUNICIPIOS (CUF INTEGER,UF VARCHAR(2),XUF VARCHAR(120) CHARACTER SET WIN1252, CMUN  VARCHAR(7) , XMUN  VARCHAR(120) CHARACTER SET WIN1252, primary key(cmun));"
'        Cnn.Execute sql
        Dim LineofText As String
        Open App.Path & "\municipios-insert-firebird.sql" For Input As #1
        Do While Not EOF(1)
            Line Input #1, LineofText
            'Debug.Print LineofText
            Cnn.Execute LineofText
        Loop
        Close #1
    End If
    
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    sql = "SELECT  COUNT(*) FROM RDB$RELATIONS WHERE  RDB$RELATION_NAME =  'TAB_CSOSN' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = ""
        sql = sql & "CREATE TABLE TAB_CSOSN " & vbCr
        sql = sql & "("
        sql = sql & "   IDCSOSN INTEGER," & vbCr
        sql = sql & "   IDSTATUS INTEGER," & vbCr
        sql = sql & "   CSOSN VARCHAR(3)," & vbCr
        sql = sql & "   DESCRICAO VARCHAR(200)," & vbCr
        sql = sql & "   ENTRADA VARCHAR(1)," & vbCr
        sql = sql & "   ATIVO VARCHAR(1)," & vbCr
        sql = sql & "   ICMS VARCHAR(1)," & vbCr
        sql = sql & "   ISENTO VARCHAR(1)," & vbCr
        sql = sql & "   ICMSSUBST VARCHAR(1)," & vbCr
        sql = sql & "   IPI VARCHAR(1)" & vbCr
        sql = sql & ")"
        
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "INSERT INTO TAB_CSOSN (IDCSOSN, IDSTATUS, CSOSN, DESCRICAO, ENTRADA, ATIVO, ICMS, ISENTO, ICMSSUBST, IPI)"
        sql = sql & "VALUES (1,242,101,'Tributada pelo Simples Nacional com Permissão de Crédito','E','S','S','N','N','S');"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "INSERT INTO TAB_CSOSN (IDCSOSN, IDSTATUS, CSOSN, DESCRICAO, ENTRADA, ATIVO, ICMS, ISENTO, ICMSSUBST, IPI)"
        sql = sql & "VALUES (2,242,102,'Tributada pelo Simples Nacional sem Permissão de Crédito','E','S','N','S','N','N');"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "INSERT INTO TAB_CSOSN (IDCSOSN, IDSTATUS, CSOSN, DESCRICAO, ENTRADA, ATIVO, ICMS, ISENTO, ICMSSUBST, IPI)"
        sql = sql & "VALUES (3,242,103,'Isenção do ICMS no Simples Nacional para Faixa de Receita Bruta','E','S','N','S','N','N');"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "INSERT INTO TAB_CSOSN (IDCSOSN, IDSTATUS, CSOSN, DESCRICAO, ENTRADA, ATIVO, ICMS, ISENTO, ICMSSUBST, IPI)"
        sql = sql & "VALUES (4,242,201,'Tributada pelo Simples Nacional com Permissão de Crédito e com cobrança do ICMS por Substituição Tributária','E','S','S','N','S','N');"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "INSERT INTO TAB_CSOSN (IDCSOSN, IDSTATUS, CSOSN, DESCRICAO, ENTRADA, ATIVO, ICMS, ISENTO, ICMSSUBST, IPI)"
        sql = sql & "VALUES (5,242,202,'Tributada pelo Simples Nacional sem Permissão de Crédito e com cobrança do ICMS por Substituição Tributária','E','S','S','N','S','N');"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = ""
        sql = sql & "INSERT INTO TAB_CSOSN (IDCSOSN, IDSTATUS, CSOSN, DESCRICAO, ENTRADA, ATIVO, ICMS, ISENTO, ICMSSUBST, IPI)"
        sql = sql & "VALUES (6,242,203,'Isenção do ICMS no Simples Nacional para Faixa de Receita Bruta e com Cobrança de ICMS por Substituição Tribuária','E','S','N','N','S','N');"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = ""
        sql = sql & "INSERT INTO TAB_CSOSN (IDCSOSN, IDSTATUS, CSOSN, DESCRICAO, ENTRADA, ATIVO, ICMS, ISENTO, ICMSSUBST, IPI)"
        sql = sql & "VALUES (7,242,300,'Imune','E','S','N','N','N','N');"
        Cnn.Execute sql
        Cnn.CommitTrans
        sql = ""
        
        sql = sql & "INSERT INTO TAB_CSOSN (IDCSOSN, IDSTATUS, CSOSN, DESCRICAO, ENTRADA, ATIVO, ICMS, ISENTO, ICMSSUBST, IPI)"
        sql = sql & "VALUES (8,242,400,'Não Tributada Pelo Simples Nacional','E','S','N','N','N','N');"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "INSERT INTO TAB_CSOSN (IDCSOSN, IDSTATUS, CSOSN, DESCRICAO, ENTRADA, ATIVO, ICMS, ISENTO, ICMSSUBST, IPI)"
        sql = sql & "VALUES (9,242,500,'ICMS Cobrado Anteriormente por Substituição Tributária (Substituído) ou por Antecipação','E','S','N','N','N','N');"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "INSERT INTO TAB_CSOSN (IDCSOSN, IDSTATUS, CSOSN, DESCRICAO, ENTRADA, ATIVO, ICMS, ISENTO, ICMSSUBST, IPI)"
        sql = sql & "VALUES (10,242,900,'Outros','E','S','N','N','N','N');"
        Cnn.Execute sql
        Cnn.CommitTrans
        
    End If
    
    cont = 0
    
    '''    'Tabela Nova, para Acompanhamento de CFOP
    '''    '------------------------------------------------------
    '''    'verifica coluna se a coluna existe caso nao cria
    '''    '------------------------------------------------------
    
    'sql = "DROP TABLE cad_naturezas "
    'Cnn.Execute sql
    'cnn.Execute sql
    
    sql = "SELECT  COUNT(*) FROM RDB$RELATIONS WHERE  RDB$RELATION_NAME = 'CAD_NATUREZAS' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
'''        sql = "ALTER TABLE PRODUTO ADD CFOP CHAR(4)"
'''        Cnn.Execute sql
'''        cnn.Execute sql
'''
'''        Cnn.Execute "UPDATE PRODUTO SET CFOP = '5405'"
'''        cnn.Execute "UPDATE PRODUTO SET CFOP = '5405'"
        
        sql = "CREATE TABLE CAD_NATUREZAS (" & vbCr
        sql = sql & "  idNatureza smallint  NOT NULL," & vbCr
        sql = sql & "  idUF INT  DEFAULT  0 NOT NULL," & vbCr 'INT DEFAULT 99 NOT NULL,
        sql = sql & "  CFOP char(4) DEFAULT  NULL," & vbCr
        sql = sql & "  Natureza varchar(45)  DEFAULT  NULL," & vbCr
        sql = sql & "  Observacao varchar(255)  DEFAULT  NULL," & vbCr
        sql = sql & "  Inciso varchar(255)  DEFAULT  NULL," & vbCr
        sql = sql & "  ICMS INT DEFAULT 0 NOT NULL," & vbCr
        sql = sql & "  ISubst INT DEFAULT 0 NOT NULL," & vbCr
        sql = sql & "  Flag INT DEFAULT 0 NOT NULL," & vbCr
        sql = sql & "  Ativo char(1)  DEFAULT 'N' NOT NULL," & vbCr
        sql = sql & "  Estoque char(1)  DEFAULT 'N' NOT NULL," & vbCr
        sql = sql & "  Custo char(1)  DEFAULT 'N' NOT NULL," & vbCr
        sql = sql & "  CMedio char(1)  DEFAULT 'N' NOT NULL," & vbCr
        sql = sql & "  PVenda char(1)  DEFAULT 'N' NOT NULL," & vbCr
        sql = sql & "  Result char(1)  DEFAULT 'N' NOT NULL," & vbCr
        sql = sql & "  CReceber char(1)  DEFAULT 'N' NOT NULL," & vbCr
        sql = sql & "  CPagar char(1)  DEFAULT 'N' NOT NULL," & vbCr
        sql = sql & "  Entrada char(1)  DEFAULT 'N' NOT NULL," & vbCr
        sql = sql & "  Saida char(1)  DEFAULT 'N' NOT NULL," & vbCr
        sql = sql & "  PMinimo char(1)  DEFAULT 'N' NOT NULL, " & vbCr
        sql = sql & "  PRIMARY KEY  (idNatureza));" & vbCr
        Cnn.Execute sql
        Cnn.CommitTrans
        
        Open App.Path & "\CFOP.TXT" For Input As #1
        Do While Not EOF(1)
            Line Input #1, LineofText
            'Debug.Print LineofText
            Cnn.Execute LineofText
            Cnn.CommitTrans
        Loop
        Close #1
    End If
    
    Rstemp6.Close
    Set Rstemp6 = Nothing

    ' VERIFICA SE O CAMPO EXISTE SE NAO CRIA
    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'DEVOLUCAO_NFE' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "CREATE TABLE DEVOLUCAO_NFE (ID INTEGER NOT NULL, SEQUENCIA DOUBLE PRECISION NOT NULL,DATA_NF DATE,NF DOUBLE PRECISION,"
        sql = sql & " COD_FORNECEDOR DOUBLE PRECISION,TOTAL_SAIDA DOUBLE PRECISION,PRIMARY KEY (ID))"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "CREATE TABLE ITENS_DEVOLUCAO_NFE (ID INTEGER NOT NULL, SEQUENCIA DOUBLE PRECISION, CODIGO_PRODUTO DOUBLE PRECISION, QTDE DOUBLE PRECISION,"
        sql = sql & "VALOR_UNITARIO DOUBLE PRECISION,VALOR_TOTAL DOUBLE PRECISION, PRIMARY KEY (ID))"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        On Error Resume Next
        ' Indices
        sql = "CREATE INDEX ID_X ON DEVOLUCAO_NFE (ID)"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_SEQ_DEVOLUCAO_NFE "
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "SET GENERATOR GEN_SEQ_DEVOLUCAO_NFE TO 0"
        Cnn.Execute sql
        Cnn.CommitTrans
        
         'cria TRIGGER PARA AUTONUMERADOR
        sql = " CREATE TRIGGER DEVOLUCAO_NFE FOR DEVOLUCAO_NFE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_SEQ_DEVOLUCAO_NFE, 1); END  "
        Cnn.Execute sql
        Cnn.CommitTrans
        
        'ITENS_DEVOLUCAO_NFE
        sql = "CREATE INDEX ID_Y ON ITENS_DEVOLUCAO_NFE (ID)"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_SEQ_ITENS_DEVOLUCAO_NFE "
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "SET GENERATOR GEN_SEQ_ITENS_DEVOLUCAO_NFE TO 0"
        Cnn.Execute sql
        Cnn.CommitTrans
        
         'cria TRIGGER PARA AUTONUMERADOR
        sql = " CREATE TRIGGER ITENS_DEVOLUCAO_NFE FOR ITENS_DEVOLUCAO_NFE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_SEQ_ITENS_DEVOLUCAO_NFE, 1); END  "
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "SELECT * FROM CAD_MENUS WHERE MENU_CD_CODI = 136 "
        Set Rstemp6 = New ADODB.Recordset
        Rstemp6.Open sql, Cnn, 1, 2
        If Rstemp6.RecordCount = 0 Then
            sql = "INSERT INTO CAD_MENUS VALUES("
            sql = sql & "'SisAdven',"
            sql = sql & "136,"
            Menu = UCase("menu_movimentacao_nfe_Devolucao")
            sql = sql & "'" & Menu & "',"
            sql = sql & "'Movimentação - NFe Devolução')"
            frmMenu.menu_movimentacao_nfe_Devolucao.Enabled = True
            Cnn.Execute sql
            Cnn.CommitTrans
        End If
    End If
    Rstemp6.Close
    Set Rstemp6 = Nothing

'    On Error GoTo TRATA_NCM
'    'sql = "SELECT NCM FROM PRODUTO "
'    'Set Rstemp6 = New ADODB.Recordset
'    'Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    'Rstemp6.Close
'    'Set Rstemp6 = Nothing
'
'    'Screen.MousePointer = 1
'
'    Exit Function
'
'TRATA_NCM:
'    If Err.Number <> 0 Then
'        Set Rstemp6 = Nothing
'        sql = "ALTER TABLE PRODUTO ADD NCM VARCHAR(8)"
'        Cnn.Execute sql
'    End If
 
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
Public Function VerificaNulo(campo As Field) As Variant
    If IsNull(campo) Then
        If campo.Type = adVarChar Then
            VerificaNulo = ""
        Else
            VerificaNulo = 0
        End If
    Else
        VerificaNulo = campo
    End If
End Function
'====== VERIFICANDO E TIRANDO OS ACENTOS DAS PALABRAS =============================================
Function TiraAcento(Palavra)
CAcento = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
SAcento = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
texto = ""
    If Palavra <> "" Then
        For X = 1 To Len(Palavra)
            Letra = Mid(Palavra, X, 1)
            Pos_Acento = InStr(CAcento, Letra)
                If Pos_Acento > 0 Then
                    Letra = Mid(SAcento, Pos_Acento, 1)
                End If
            texto = texto & Letra
        Next
        TiraAcento = texto
    End If
End Function

Public Function VerificaAtualizacoes2()
    '05/10/2010
    On Error GoTo TRATA_ENDERECO_ENTREGA_CLIENTE
    sql = "SELECT ENDERECO_ENTREGA FROM CLIENTE "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, 1, 2
    Rstemp6.Close
    Set Rstemp6 = Nothing

TRATA_ENDERECO_ENTREGA_CLIENTE:
    If Err.Number <> 0 Then
        Set Rstemp6 = Nothing
        sql = "ALTER TABLE CLIENTE ADD ENDERECO_ENTREGA VARCHAR(60)"
        Cnn.Execute sql
        Call VerificaAtualizacoes2
    End If
    

        
'Fecha conexão para garantir transações
If Cnn.State = 1 Then
    On Error Resume Next
    Cnn.Close
    Set Cnn = Nothing
    If Conecta_Banco = False Then
        Exit Function
    End If
End If

Exit Function

End Function

Function VerificaPalavra(atributo)

Dim i
Dim ID
Dim Auxiliar
Dim Resultado

Auxiliar = Split(atributo, " ", -1, vbBinaryCompare)

For i = LBound(Auxiliar) To UBound(Auxiliar)
    Resultado = Resultado & " " & TiraAcento(Auxiliar(i))
Next

VerificaPalavra = Trim(Resultado)

End Function
Public Function NoRecord(Rs As ADODB.Recordset) As Boolean
  'exemplo uso
'    If NoRecord(Rstemp) Then
'        MsgBox "No record for the years specified found", vbCritical
'        ClearRS Rstemp
'        'ClearRS rsChartDataYear
'        Exit Sub
'    End If
    
    
    If Rs.BOF And Rs.EOF Then
        NoRecord = True
    Else
        NoRecord = False
    End If
     
End Function


Public Sub ClearRS(Rstemp_ As ADODB.Recordset)
    On Error Resume Next
    If Rstemp_.State = adStateOpen Then Rstemp_.Close
    Set Rstemp_ = Nothing
End Sub
Public Sub FormMode()
    On Error GoTo ShowErr

    With Frm_Transf_Produtos.Toolbar1
        .Buttons(2).Enabled = True
        .Buttons(4).Enabled = False
        .Buttons(5).Enabled = False
        .Buttons(6).Enabled = False
        .Buttons(7).Enabled = False
    End With
    Exit Sub
ShowErr:
    MsgBox "Error No: " & Err.Number & " (" & Err.Description & ") - FormMode - Module Mainmodule"
End Sub
Public Sub MostraErro()

    Dim Erro As Error
    
    If Cnn.Errors.Count > 0 Then
        For Each Erro In Cnn.Errors
            MsgBox Erro.Number & ": " & Erro.Description, vbCritical, "Descrição de Erro"
        Next
    Else
        MsgBox Err.Number & ": " & Err.Description, vbCritical, "Descrição de Erro"
    End If

End Sub


'------------------------------------------------------------------------
Public Sub ClearCommandParameters()
'------------------------------------------------------------------------

    Dim lngX    As Long
    
    For lngX = (mobjCmd.Parameters.Count - 1) To 0 Step -1
        mobjCmd.Parameters.Delete lngX
    Next
    
    'ou
    'Set mobjCmd = Nothing

End Sub
Public Function Retorna_Tributo_ICMS(ByVal parametro As String) As String
    Dim Trib_ICMS As String
    
    Trib_ICMS = ""

    Select Case parametro
        Case "1"
            Trib_ICMS = "II"
        Case "2"
            Trib_ICMS = "FF"
        Case "3"
            Trib_ICMS = "0700"
        Case "4"
            Trib_ICMS = "1200"
        Case "5"
            Trib_ICMS = "1800"
        Case "6"
            Trib_ICMS = "2500"
    End Select

    Retorna_Tributo_ICMS = Trib_ICMS

End Function
Public Function Tira_Acento_SAT(StrAcento)
caract = StrAcento
    For i = 1 To Len(StrAcento)
        Letra = Mid(StrAcento, i, 1)
        Select Case Letra
            Case "á", "Á", "à", "À", "ã", "Ã", "â", "Â", "â", "ä", "Ä"
                Letra = "A"
            Case "é", "É", "ê", "Ê", "Ë", "ë", "È", "è"
                Letra = "E"
            Case "í", "Í", "ï", "Ï", "Ì", "ì"
                Letra = "I"
            Case "ó", "Ó", "ô", "Ô", "õ", "Õ", "ö", "Ö", "ò", "Ò"
                Letra = "O"
            Case "ú", "Ú", "Ù", "ù", "ú", "û", "ü", "Ü", "Û"
                Letra = "U"
            Case "ç", "Ç"
                Letra = "C"
'            Case "&"
'                Letra = ""
            Case "'"
                Letra = ""
            Case "''"
                Letra = ""
            Case "ñ"
                Letra = "N"
            Case "¹"
                Letra = "1"
            Case "²"
                Letra = "2"
            Case "³"
                Letra = "3"
            Case "º"
                Letra = ""
            Case "&"
                Letra = "&amp;"
                'letra = "& = &"
            Case "<"
                Letra = "&lt;"
            Case ">"
                Letra = "&gt;"
            Case "''"
                'letra = ""
                Letra = "&quot;"
        End Select
        texto = texto & Letra
    Next
    Tira_Acento_SAT = texto
End Function


Public Function Tira_Acento(StrAcento)
caract = StrAcento
    For i = 1 To Len(StrAcento)
        Letra = Mid(StrAcento, i, 1)
        Select Case Letra
            Case "á", "Á", "à", "À", "ã", "Ã", "â", "Â", "â", "ä", "Ä"
                Letra = "A"
            Case "é", "É", "ê", "Ê", "Ë", "ë", "È", "è"
                Letra = "E"
            Case "í", "Í", "ï", "Ï", "Ì", "ì"
                Letra = "I"
            Case "ó", "Ó", "ô", "Ô", "õ", "Õ", "ö", "Ö", "ò", "Ò"
                Letra = "O"
            Case "ú", "Ú", "Ù", "ù", "ú", "û", "ü", "Ü", "Û"
                Letra = "U"
            Case "ç", "Ç"
                Letra = "C"
            Case "&"
                Letra = ""
            Case "'"
                Letra = ""
            Case "''"
                Letra = ""
            Case "ñ"
                Letra = "N"
            Case "¹"
                Letra = "1"
            Case "²"
                Letra = "2"
            Case "³"
                Letra = "3"
            Case "º"
                Letra = ""
            Case "<"
                Letra = ""
            Case ">"
                Letra = ""
            Case "''"
                Letra = ""
        End Select
        texto = texto & Letra
    Next
    Tira_Acento = texto
End Function

Public Function Conecta_Banco() As Boolean

Conecta_Banco = False

On Error GoTo Erro
        
Set Cnn = New ADODB.Connection
    
    With Cnn
        '.CursorLocation = adUseServer
        .CursorLocation = adUseClient
        .Open "File Name=" & App.Path & "\cnn_fire_Servidor.udl;"
        '''.Open "DRIVER=Firebird/InterBase(r) driver; UID=SYSDBA; PWD=masterkey;DBNAME=servidor:" & App.Path & "\arqdados.GDB"
    End With
    
    Conecta_Banco = True

Exit Function
    
Erro:
    If Err.Number <> 0 Then
        MsgBox "Impossível Abrir o Banco ! Erro: " & Err.Number & " - " & Err.Description & Chr(13) + Chr(10) & _
               Chr(13) + Chr(10) & _
               "Este Erro pode ter sido causado pelos seguintes motivos: " & Chr(13) + Chr(10) & Chr(13) + Chr(10) & _
               "1.    O Servidor pode estar desligado, ou não conectado na rede." & Chr(13) + Chr(10) & _
               "2.    O Hub (Conector da rede) pode estar desligado ou cabos de rede desligados." & Chr(13) + Chr(10) & _
               "3.    Seu micro não está conectado à rede. Verifique o Cabo de Rede (Azul), e tente reiniciar o computador." & Chr(13) + Chr(10) & _
               "4.    Se nenhuma destas possibilidades funcionarem, entre em contato com o suporte técnico do sistema." & vbNewLine & vbNewLine & _
               "5.    O sistema será finalizado...!", vbInformation, "Aviso"
            Call Fecha_Formularios
        End
    End If
Err.Clear

End Function


Public Sub Fecha_Formularios()
    Dim Form As Form
    For Each Form In Forms
       Unload Form
       Set Form = Nothing
    Next Form
End Sub


Public Function FechaRecordsets()
If Rs.State = 1 Then
    Rs.Close
    Set Rs.ActiveConnection = Nothing
End If

If Rstemp.State = 1 Then
    Rstemp.Close
    Set Rstemp.ActiveConnection = Nothing
End If

If RsTemp1.State = 1 Then
    RsTemp1.Close
    Set RsTemp1.ActiveConnection = Nothing
End If

If Rstemp2.State = 1 Then
    Rstemp2.Close
    Set Rstemp2.ActiveConnection = Nothing
End If

If Rstemp3.State = 1 Then
    Rstemp3.Close
    Set Rstemp3.ActiveConnection = Nothing
End If

If Rstemp5.State = 1 Then
    Rstemp5.Close
    Set Rstemp5.ActiveConnection = Nothing
End If

If Rstemp6.State = 1 Then
    Rstemp6.Close
    Set Rstemp6.ActiveConnection = Nothing
End If


End Function


Public Sub MenuStatus(EnabledStatus As Boolean, EnabledFile As Boolean, CheckLiskIntro As Boolean)

    mnuEdit.Enabled = EnabledStatus
    mnuSetup.Enabled = EnabledStatus
    
        mnuSales.Enabled = CheckLiskIntro
        mnuPurchasing.Enabled = CheckLiskIntro
        mnuInventory.Enabled = CheckLiskIntro
        mnuAccounting.Enabled = CheckLiskIntro
        'mnuPayroll.Enabled = CheckLiskIntro
        'mnuShowReport.Enabled = CheckLiskIntro
        mnuFinaceCharges.Enabled = CheckLiskIntro
        mnuCloseMonth.Enabled = CheckLiskIntro
    'mnuTools.Enabled = EnabledStatus
    'mnuWindow.Enabled = EnabledStatus
    'mnuHelp.Enabled = EnabledStatus
    tbToolBar.Buttons("InventoryList").Enabled = EnabledStatus
    tbToolBar.Buttons("Contacts").Enabled = EnabledStatus
    
    tbToolBar.Buttons("Save").Enabled = EnabledStatus
    tbToolBar.Buttons("stop").Enabled = EnabledStatus
    
    mnuFileClose.Enabled = EnabledFile
    mnuFileSaveAs.Enabled = EnabledFile
    mnuFileProperties.Enabled = EnabledFile
    mnuHistory.Enabled = EnabledFile
    mnuErrorLog.Enabled = EnabledFile
    'mnuFilePageSetup.Enabled = EnabledFile
    mnuFilePrintPreview.Enabled = EnabledFile
    mnuFilePrint.Enabled = EnabledFile
    
    'mnuPayrollPayEmployees.Enabled = False
    'mnuPayrollViodChecks.Enabled = False
    'mnuFileSend.Enabled = False
End Sub



Public Function VerificaAtualizacoes_old()
    Dim sql As String
    
    On Error GoTo TRATA_XIBIU

    sql = "SELECT COUNT(SEQUENCIA) FROM XIBIU "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, 1, 2
    
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    'FIREBIRD tabelas relacionadas
'    sql = " CREATE TABLE CLIENTES ("
'    sql = sql & " CLI_CODIGO INTEGER NOT NULL, CLI_TIPOPESSOA CHAR(1), CLI_NOME VARCHAR(40), CLI_RAZAOSOCIAL VARCHAR(40), "
'    sql = sql & " CLI_RG VARCHAR(10), CLI_IE VARCHAR(11), CLI_CPF VARCHAR(11), CLI_CNPJ VARCHAR(14),"
'    sql = sql & "  Primary Key(CLI_CODIGO)" & ")"
'    Cnn.Execute sql
'
'    sql = " CREATE TABLE VENDAS ("
'    sql = sql & " SEQUENCIA INTEGER NOT NULL,  CLI_CODIGO INTEGER NOT NULL, VEN_VALOR DOUBLE PRECISION, "
'    sql = sql & " VEN_DESCONTO DOUBLE PRECISION,  PRIMARY KEY(SEQUENCIA),"
'    sql = sql & " FOREIGN KEY(CLI_CODIGO) REFERENCES CLIENTES(CLI_CODIGO)" & ")"
'    Cnn.Execute sql

    
    sql = "CREATE TABLE TESTE (COD_PRODUTO double PRECISION, QTDE DOUBLE PRECISION, PRECO_UNIT double PRECISION, PRECO_TOTAL double PRECISION)"
    Cnn.Execute sql
    Call VerificaAtualizacoes
    
    'SQL = "CREATE TABLE tbCatalogue ([Id] COUNTER, [Stock Code] TEXT(10), [Account Number] TEXT(6))"
    'SQL = "CREATE TABLE XIBIU ([SEQUENCIA] COUNTER, [PEDIDO] DOUBLE, [VALOR] double, data DATETIME,[USUÁRIO] TEXT(10))"
    'SQL = "CREATE Table MyTable (MyID AutoIncrement CONSTRAINT MyIdConstraint PRIMARY KEY, MyText CHAR(50))"
    
    'remove tabela
    'SQL = "DROP Table XIBIU"
    
    'CRIA CAMPO NA TABELA
    'SQL = "ALTER TABLE SAIDAS_PRODUTO ADD COLUMN STATUS_NF TEXT(1)"
    
    On Error GoTo TRATA_REL_RANCKING_PRODUTOS_VENDEDOR

    sql = "SELECT COUNT(*) FROM REL_RANCKING_PRODUTOS_VENDEDOR "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, 1, 2
    
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    Exit Function
    
TRATA_XIBIU:

    If Err.Number <> 0 Then
        Set Rstemp6 = Nothing
        'CRIA TABELA COM INDICE
        sql = "CREATE Table XIBIU (SEQUENCIA AutoIncrement CONSTRAINT MySequencia PRIMARY KEY, PEDIDO DOUBLE, VALOR double, DATA DATETIME, TELA_CHAMOU TEXT(20), USUARIO TEXT(30), SENHA TEXT(6))"
        
        Cnn.Execute sql
        
        'Set dbSeguranca = OpenDatabase(App.Path & "\Seguranca.mdb", False, False)
        'CRIA CAMPO NA TABELA
        
        sql = "select * from cad_menus where MENU_CD_CODI = 85"
        Set Rstemp = New ADODB.Recordset
        Rstemp.Open sql, Cnn, 1, 2
        If Rstemp.RecordCount = 0 Then
            sql = "INSERT INTO Cad_Menus VALUES ('SisAdven',85,'MENU_CONSULTA_PEDIDOS_EXCLUIDOS','Consulta - Pedidos Excluidos')"
            Cnn.Execute sql
        End If
        
        Rstemp.Close
        Set Rstemp = Nothing
        dbSeguranca.Close
        Set dbSeguranca = Nothing
        Call VerificaAtualizacoes
    End If
    
TRATA_REL_RANCKING_PRODUTOS_VENDEDOR:
    Set Rstemp6 = Nothing

    sql = "CREATE TABLE REL_RANCKING_PRODUTOS_VENDEDOR ([COD_PRODUTO] double, [QTDE] DOUBLE, [PRECO_UNIT] double, [PRECO_TOTAL] double)"
    Cnn.Execute sql
    Call VerificaAtualizacoes

End Function


Public Function VerificaAtualizacoes()

    sql = "SELECT * FROM CAD_MENUS WHERE MENU_CD_CODI = 13 "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, 1, 2
    If Rstemp6.RecordCount = 0 Then
        sql = "INSERT INTO CAD_MENUS VALUES("
        sql = sql & "'SisAdven',"
        sql = sql & "13,"
        Menu = UCase("menu_cadastro_Cedente")
        sql = sql & "'" & Menu & "',"
        sql = sql & "'Cadastro - Cedente')"
        frmMenu.menu_cadastro_Cedente.Enabled = True
        Cnn.Execute sql
    End If

    Rstemp6.Close
    Set Rstemp6 = Nothing
  

    On Error GoTo TRATA_FORMA_PGTO
    sql = "SELECT CAIXA_RECEBEU FROM FORMA_PGTO "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, 1, 2
    Rstemp6.Close
    Set Rstemp6 = Nothing

TRATA_FORMA_PGTO:
    
    If Err.Number <> 0 Then
        Set Rstemp6 = Nothing
        'CRIA CAMPO NA TABELA EXISTENTE
        sql = "ALTER TABLE FORMA_PGTO ADD CAIXA_RECEBEU VARCHAR(30)  "
        Cnn.Execute sql
        Call VerificaAtualizacoes
    End If
    
    On Error GoTo TRATA_CONTA_CORRENTE_BOLETO
    sql = "SELECT * FROM CONTA_CORRENTE_BOLETO "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, 1, 2
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    On Error GoTo TRATA_CONTA_CORRENTE_BOLETO_JUROS
    sql = "SELECT PercentualJurosDiaAtraso  FROM CONTA_CORRENTE_BOLETO "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, 1, 2
    Rstemp6.Close
    Set Rstemp6 = Nothing

    On Error GoTo TRATA_BOLETOS_PAGOS
    sql = "SELECT * FROM BOLETOS_PG "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, 1, 2
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    On Error GoTo REL_TRATA_BOLETOS_PAGOS
    sql = "SELECT * FROM REL_BOLETOS_PG "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, 1, 2
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    On Error GoTo TRATA_ID_RECE_PAGA
    sql = "SELECT ID FROM RECE_PAGA "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, 1, 2
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
TRATA_CONTA_CORRENTE_BOLETO:
    If Err.Number <> 0 Then
        Set Rstemp6 = Nothing
        sql = "CREATE TABLE CONTA_CORRENTE_BOLETO (ID INTEGER NOT NULL, BANCOCEDENTE VARCHAR(3), AGENCIACEDENTE VARCHAR(20), CONTACORRENTECEDENTE VARCHAR(20),"
        sql = sql & "CODIGOCEDENTE VARCHAR(20), NOMECEDENTE VARCHAR(20), CNPJCPFCEDENTE VARCHAR(19), INICIONOSSONUMERO VARCHAR(20), FIMNOSSONUMERO VARCHAR(20), PROXIMONOSSONUMERO VARCHAR(20), ARQUIVOLICENCA VARCHAR(200), DIAS_PROTESTO INTEGER, DEMONSTRATIVO VARCHAR(200), INSTRUCAO1 VARCHAR(200), INSTRUCAO2 VARCHAR(200),INSTRUCAO3 VARCHAR(200),CAMINHO_LOGOTIPO_BOLETO_IMP VARCHAR(200))"
        Cnn.Execute sql
        
        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_BL_ID1 "
        Cnn.Execute sql
        sql = "SET GENERATOR GEN_BL_ID1 TO 0"
        Cnn.Execute sql
        
        'cria TRIGGER PARA AUTONUMERADOR
        sql = " CREATE TRIGGER CONTA_CORRENTE_BOLETO_BI FOR CONTA_CORRENTE_BOLETO ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_BL_ID1, 1); END  "
        Cnn.Execute sql
   
        'TABELA BOLETOS
        sql = "ALTER TABLE RECE_PAGA ADD PEDIDO DOUBLE PRECISION, ADD COD_CLIENTE DOUBLE PRECISION, ADD NOSSO_NUMERO VARCHAR(20) "
        Cnn.Execute sql
        Call VerificaAtualizacoes
    End If
    
TRATA_CONTA_CORRENTE_BOLETO_JUROS:
    If Err.Number <> 0 Then
        Set Rstemp6 = Nothing
        sql = "ALTER TABLE CONTA_CORRENTE_BOLETO ADD PERCENTUALJUROSDIAATRASO  DOUBLE PRECISION, ADD PERCENTUALMULTAATRASO DOUBLE PRECISION"
        Cnn.Execute sql
        Call VerificaAtualizacoes
    End If
    
TRATA_BOLETOS_PAGOS:
    
    If Err.Number <> 0 Then
        sql = "CREATE TABLE BOLETOS_PG (SEQUENCIA INTEGER NOT NULL, NOSSO_NUMERO VARCHAR(20), DATA_VENCIMENTO DATE, VLR_BOLETO DOUBLE PRECISION, DATA_PAGAMENTO DATE, VLR_PAGO DOUBLE PRECISION, PRIMARY KEY (SEQUENCIA))"
        Cnn.Execute sql

        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_BLPG_ID "
        Cnn.Execute sql
        sql = "SET GENERATOR GEN_BLPG_ID TO 0"
        Cnn.Execute sql

        'cria TRIGGER PARA AUTONUMERADOR
        sql = " CREATE TRIGGER BOLETOS_PG FOR BOLETOS_PG ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        sql = sql & "if (NEW.SEQUENCIA is NULL) then NEW.SEQUENCIA = GEN_ID(GEN_BLPG_ID, 1); END  "
        Cnn.Execute sql
        Call VerificaAtualizacoes
    End If
    
    
REL_TRATA_BOLETOS_PAGOS:
    
    If Err.Number <> 0 Then
        sql = "CREATE TABLE REL_BOLETOS_PG (NOSSO_NUMERO VARCHAR(20), DATA_VENCIMENTO DATE, VLR_BOLETO DOUBLE PRECISION, DATA_PAGAMENTO DATE, VLR_PAGO DOUBLE PRECISION )"
        Cnn.Execute sql
        Call VerificaAtualizacoes
    End If
    
      

    
TRATA_ID_RECE_PAGA:

    If Err.Number <> 0 Then
        Set Rstemp6 = Nothing
        
        sql = "ALTER TABLE RECE_PAGA ADD ID INTEGER NOT NULL "
        Cnn.Execute sql
    
        sql = " update RECE_PAGA SET ID = SEQUENCIA "
        Cnn.Execute sql
        
        DoEvents
        
        If Cnn.State = 1 Then
            On Error Resume Next
            Cnn.Close
            Set Cnn = Nothing
            If Conecta_Banco = False Then
                Exit Function
            End If
        End If
        
        'Primary Keys     SITE REFEENCIA = "http://www.firebirdsql.org/dotnetfirebird/create-a-new-database-from-an-sql-script.html"
        sql = "ALTER TABLE RECE_PAGA ADD PRIMARY KEY (ID)"
        Cnn.Execute sql
        
        ' Indices
        sql = "CREATE INDEX ID_X ON RECE_PAGA (ID)"
        Cnn.Execute sql
        
        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_RECEPG "
        Cnn.Execute sql
        
        ULTIMO_NUMERO = (Select_Max("RECE_PAGA", "SEQUENCIA") - 1)
        
        'sql = "SET GENERATOR GEN_RCPG_ID TO 0"
        sql = "SET GENERATOR GEN_RECEPG TO " & CDbl(ULTIMO_NUMERO)
        Cnn.Execute sql
        
          'cria TRIGGER PARA AUTONUMERADOR
         sql = " CREATE TRIGGER RECE_PAGA_BI FOR RECE_PAGA ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
         sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_RECEPG, 1); END  "
         Cnn.Execute sql
        Call VerificaAtualizacoes
    End If
    
    
   Exit Function
        
    'Fecha conexão para garantir transações
    If Cnn.State = 1 Then
        On Error Resume Next
        Cnn.Close
        Set Cnn = Nothing
        If Conecta_Banco = False Then
            Exit Function
        End If
    End If
    
    
    
'
'    On Error Resume Next
'    sql = "Create PROCEDURE SP_UPDATE_FORMA_PGTO1000 (NRO_PEDIDO DOUBLE PRECISION, FORMA_PGTO VARCHAR(5), STR_STATUS_SAIDA VARCHAR(1))"
'    sql = sql + " AS BEGIN "
'    sql = sql + " UPDATE SAIDAS_PRODUTO SET FORMAPGTO=:FORMA_PGTO, STATUS_SAIDA =:STR_STATUS_SAIDA WHERE SEQUENCIA=:NRO_PEDIDO; END"
'    Cnn.Execute sql
'
'    On Error GoTo trata:
'    sql = "CREATE PROCEDURE SP_PRODUTO 2("
'    sql = sql + " @CODIGO_INTERNO VARCHAR(13))"
'    sql = sql + " RETURNS ("
'    sql = sql + " OUT_DESCRICAO VARCHAR(80), OUT_PRECO DOUBLE PRECISION )"
'    sql = sql + " AS BEGIN FOR "
'    sql = sql + " Select CODIGO_INTERNO, DESCRICAO, PRECO "
'    sql = sql & " FROM PRODUTO WHERE CODIGO_INTERNO = :@CODIGO_INTERNO INTO :OUT_DESCRICAO, :OUT_PRECO "
'    sql = sql + "  DO suspend; END "
'    Cnn.Execute sql
'trata:

'''    On Error GoTo trata:
'''    sql = "CREATE PROCEDURE SP_PRODUTO ("
'''    sql = sql + " CODIGO_PRODUTO VARCHAR(13))"
'''    sql = sql + " RETURNS ("
'''    sql = sql + " OUT_CODIGO DOUBLE PRECISION, OUT_DESCRICAO VARCHAR(80), OUT_PRECO DOUBLE PRECISION )"
'''    sql = sql + " AS BEGIN FOR "
'''    sql = sql + " Select CODIGO_INTERNO, DESCRICAO, PRECO "
'''    sql = sql & " FROM PRODUTO WHERE CODIGO_INTERNO = :CODIGO_PRODUTO INTO :OUT_CODIGO, OUT_DESCRICAO, :OUT_PRECO "
'''    sql = sql + "  DO suspend; END "
'''    Cnn.Execute sql
'''trata:
'''
'''    Set mobjCmd = New ADODB.Command
'''    Set mobjCmd.ActiveConnection = Cnn
'''
'''    **********
'''    Call ClearCommandParameters
'''
'''    mobjCmd.CommandType = adCmdStoredProc
'''
'''    Codigo = 1026
'''     IN-parameters
'''    mobjCmd.Parameters.Append mobjCmd.CreateParameter("CODIGO_PRODUTO", adVarChar, adParamInput, 14, Codigo)
'''
'''    OUT -Parameters
'''    mobjCmd.Parameters.Append mobjCmd.CreateParameter("OUT_CODIGO", adDouble, adParamOutput) 'RETORNA_PARAMETRO DO CAMPO
'''    mobjCmd.Parameters.Append mobjCmd.CreateParameter("OUT_DESCRICAO", adBSTR, adParamOutput)
'''    mobjCmd.Parameters.Append mobjCmd.CreateParameter("OUT_PRECO", adDouble, adParamOutput)
'''    mobjCmd.Parameters.Append mobjCmd.CreateParameter("@VLR_TOT_CUST", adDouble, adParamInput, 8, FormatNumber(RsTemp1!TOT_ITEN_CUSTO, 3))
'''    mobjCmd.Parameters.Append mobjCmd.CreateParameter("@VLR_TOT_VEND", adDouble, adParamInput, 8, FormatNumber(RsTemp1!TOT_ITEN_VENDA, 3))
'''    mobjCmd.Parameters.Append mobjCmd.CreateParameter("@PERC_LUCRO", adDouble, adParamInput, 8, PERC_LUCRO_ITEM)
'''    mobjCmd.Parameters.Append mobjCmd.CreateParameter("@PERC_PARTICIP_PROD", adDouble, adParamInput, 8, FormatNumber(PARTICIPACAO_PRODUTO, 2))
'''    mobjCmd.CommandText = "SP_PRODUTO"
'''    mobjCmd.Execute
'''
'''    strOutputParam = mobjCmd.Parameters("OUT_CODIGO") 'RETORNA PARAMETRO  - adParamOutput
'''    Descricao = mobjCmd.Parameters("OUT_DESCRICAO") 'RETORNA PARAMETRO  - adParamOutput
'''    PRECO = mobjCmd.Parameters("OUT_PRECO") 'RETORNA PARAMETRO  - adParamOutput
'''    strOutputParam = mobjCmd.Parameters(0).Value
'''    strOutputParam = mobjCmd.Parameters(1).Value
'''    strOutputParam = mobjCmd.Parameters(2).Value
'''    strOutputParam = mobjCmd.Parameters(3).Value
  

    
    'FIREBIRD tabelas relacionadas
'    sql = " CREATE TABLE CLIENTES ("
'    sql = sql & " CLI_CODIGO INTEGER NOT NULL, CLI_TIPOPESSOA CHAR(1), CLI_NOME VARCHAR(40), CLI_RAZAOSOCIAL VARCHAR(40), "
'    sql = sql & " CLI_RG VARCHAR(10), CLI_IE VARCHAR(11), CLI_CPF VARCHAR(11), CLI_CNPJ VARCHAR(14),"
'    sql = sql & "  Primary Key(CLI_CODIGO)" & ")"
'    Cnn.Execute sql
'
'    sql = " CREATE TABLE VENDAS ("
'    sql = sql & " SEQUENCIA INTEGER NOT NULL,  CLI_CODIGO INTEGER NOT NULL, VEN_VALOR DOUBLE PRECISION, "
'    sql = sql & " VEN_DESCONTO DOUBLE PRECISION,  PRIMARY KEY(SEQUENCIA),"
'    sql = sql & " FOREIGN KEY(CLI_CODIGO) REFERENCES CLIENTES(CLI_CODIGO)" & ")"
'    Cnn.Execute sql

    'CRIA TABELA FIREBIRD
'    sql = "CREATE TABLE TESTE (COD_PRODUTO double PRECISION, QTDE DOUBLE PRECISION, PRECO_UNIT double PRECISION, PRECO_TOTAL double PRECISION)"
'    Cnn.Execute sql
'    Call VerificaAtualizacoes

'/* Incluindo um campo */
'ALTER TABLE NOMETABELA ADD NOVOCAMPO5 TIPO;
'
'/* Excluindo um campo */
'ALTER TABLE NOMETABELA DROP NOVOCAMPO5;
'
'/* Incluindo e excluindo ao mesmo tempo
'ALTER TABLE NOMETABELA DROP NOVOCAMPO5, ADD NOVOCAMPO6 TIPO
'
'/* Alterando o nome de um campo */
'ALTER TABLE NOMETABELA ALTER CAMPO5 TO CAMPO6;
'
'/* Adicionando uma chave primária */
'ALTER TABLE NOMETABELA ALTER PRIMARY KEY (CAMPO1);
'
'/* Adicionando uma chave estrangeira */
'ALTER TABLE NOMETABELA ALTER FOREIGN KEY (CAMPO1) REFERENCES TABELAESTRANGEIRA (CAMPOCHAVE);Alguns exemplos:
'
'
'/* Adicionando uma chave estrangeira */
'CREATE TABLE CLIENTE (
'CODIGO INTEGER NOT NULL,
'NOME VARCHAR(40) NOT NULL,
'TIPO INTEGER NOT NULL,
'ENDERECO VARCHAR(70),
'CIDADE VARCHAR(40),
'UF CHAR(2) DEFAULT 'BA',
'OBSERVACAO BLOB SUB_TYPE 1,
'DATANASCIMENTO DATE,
'DATACADASTRO DATE,
'Primary Key(Codigo)
');
    
   
    'remove tabela
    'SQL = "DROP Table XIBIU"
    
'    'CRIA CAMPO NA TABELA EXISTENTE
'    sql = "ALTER TABLE PRODUTO ADD COLUMN ULTIMA_VENDA DATE, ADD COLUMN ULTIMA_COMPRA DATE "

'    On Error GoTo TRATA_PRODUTO
'
'    sql = "SELECT ULTIMA_VENDA FROM PRODUTO "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    On Error GoTo trata_Cad_Lojas
'    sql = "SELECT * FROM CAD_LOJAS "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    On Error GoTo TRANSF_PROD
'    sql = "SELECT * FROM TRANSF_PROD "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'
'    On Error GoTo ITENS_TRANSF_PROD
'    sql = "SELECT * FROM ITENS_TRANSF_PROD "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    On Error GoTo TRATA_TABELA_COMPRAS
'    sql = "SELECT * FROM COMPRA_PRODUTO "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    On Error GoTo TRATA_REL_RANCKING_PRODUTOS_VENDEDOR
'
'    sql = "SELECT COUNT(*) FROM REL_RANCKING_PRODUTOS_VENDED"
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    On Error GoTo TRATA_PRODUTOS_ATACADO
'    sql = "SELECT COUNT(PRECO_ATACADO) FROM PRODUTO "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    On Error GoTo TRATA_VIEW_PRODUTO_DESCRIC
'    sql = "SELECT ULTIMA_VENDA FROM VIEW_PRODUTO_DESCRIC "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    On Error GoTo TRATA_VIEW_PRODUTO_DESCRIC
'    sql = "SELECT COUNT(PRECO_ATACADO) FROM VIEW_PRODUTO_DESCRIC "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    On Error GoTo TRATA_PRODUTO_ALIQUOTA
'    sql = "SELECT ALIQUOTA_ECF FROM PRODUTO "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    On Error GoTo VIEW_ESTOQUE_NEG
'    sql = "SELECT * FROM VIEW_ESTOQUE_NEG "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    On Error GoTo VIEW_SO_ESTOQUE_NEG
'    sql = "SELECT * FROM VIEW_SO_ESTOQUE_NEG "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    sql = "SELECT * FROM CAD_MENUS WHERE MENU_CD_CODI = 150 "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'    If Rstemp6.RecordCount = 0 Then
'        sql = "INSERT INTO CAD_MENUS VALUES("
'        sql = sql & "'SisAdven',"
'        sql = sql & "150,"
'        sql = sql & "'MENU_MOVIMENTACAO_SAIDA_ATACADO',"
'        sql = sql & "'Movimentação - Emissao de Pedidos Atacado')"
'        Cnn.Execute sql
'    End If
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    sql = "SELECT * FROM CAD_MENUS WHERE MENU_CD_CODI = 160 "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'    If Rstemp6.RecordCount = 0 Then
'        sql = "INSERT INTO CAD_MENUS VALUES("
'        sql = sql & "'SisAdven',"
'        sql = sql & "160,"
'        Menu = UCase("menu_movimentacao_Transferencia_Prod")
'        sql = sql & "'" & Menu & "',"
'        sql = sql & "'Movimentação - Transferência de Mercadorias Lojas')"
'        Cnn.Execute sql
'    End If
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    On Error GoTo REL_ITENS_TRANSF_PROD
'    sql = "SELECT * FROM REL_ITENS_TRANSF_PROD "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    Exit Function
'
'TRATA_PRODUTO:
'
'    If Err.Number <> 0 Then
'        Set Rstemp6 = Nothing
'        'CRIA CAMPO NA TABELA EXISTENTE
'        sql = "ALTER TABLE PRODUTO ADD ULTIMA_VENDA DATE, ADD ULTIMA_COMPRA DATE "
'         'sql = "ALTER TABLE PRODUTO ADD COLUMN ULTIMA_VENDA DATE "
'         'sql = "ALTER TABLE PRODUTO DROP ULTIMA_VENDA DATE" 'adicionar campo
'         'sql = "ALTER TABLE PRODUTO DROP ULTIMA_VENDA "     'excluir campo
'        Cnn.Execute sql
'        Call VerificaAtualizacoes
'    End If
'
'
'trata_Cad_Lojas:
'If Err.Number <> 0 Then
'    Set Rstemp6 = Nothing
'    sql = " CREATE TABLE CAD_LOJAS (SEQUENCIA DOUBLE PRECISION NOT NULL, DESCRICAO VARCHAR(30), PRIMARY KEY(SEQUENCIA))"
'    Cnn.Execute sql
'    Call VerificaAtualizacoes
'End If
'
'ITENS_TRANSF_PROD:
'If Err.Number <> 0 Then
'    Set Rstemp6 = Nothing
'    sql = " CREATE TABLE ITENS_TRANSF_PROD (SEQUENCIA DOUBLE PRECISION NOT NULL, CODIGO_PRODUTO  DOUBLE PRECISION, QTDE DOUBLE PRECISION , PRECO_CUSTO DOUBLE PRECISION )"
'    Cnn.Execute sql
'    Call VerificaAtualizacoes
'End If
'
'TRANSF_PROD:
'If Err.Number <> 0 Then
'    Set Rstemp6 = Nothing
'    sql = " CREATE TABLE TRANSF_PROD (SEQUENCIA DOUBLE PRECISION NOT NULL, DATA DATE, HORA VARCHAR(5), HISTORICO VARCHAR(60), COD_LOJA DOUBLE PRECISION , TOTAL_PEDIDO DOUBLE PRECISION, PRIMARY KEY(SEQUENCIA))"
'    Cnn.Execute sql
'    Call VerificaAtualizacoes
'End If
'
'TRATA_PRODUTOS_ATACADO:
'
'    If Err.Number <> 0 Then
'        Set Rstemp6 = Nothing
'        'CRIA CAMPOS NA TABELA PRODUTOS
'        sql = "ALTER TABLE PRODUTO ADD PRECO_ATACADO double PRECISION, ADD PRECO_MINIMO_ATACADO double PRECISION "
'        Cnn.Execute sql
'        Call VerificaAtualizacoes
'    End If
'
'TRATA_REL_RANCKING_PRODUTOS_VENDEDOR:
'    If Err.Number <> 0 Then
'        Set Rstemp6 = Nothing
'        sql = "CREATE TABLE REL_RANCKING_PRODUTOS_VENDEDOR ([COD_PRODUTO] double, [QTDE] DOUBLE, [PRECO_UNIT] double, [PRECO_TOTAL] double)"
'        Cnn.Execute sql
'        Call VerificaAtualizacoes
'    End If
'
'TRATA_VIEW_PRODUTO_DESCRIC:
'    If Err.Number <> 0 Then
'        Set Rstemp6 = Nothing
'        sql = "RECREATE VIEW VIEW_PRODUTO_DESCRIC (CODIGO,CODIGO_INTERNO,DESCRICAO,PRECO,UNIDADE,SALDO_EM_ESTOQUE, "
'        sql = sql & "ULTIMA_VENDA,ULTIMA_COMPRA, DATA_CAD_ALT, PRECO_ATACADO, PRECO_MINIMO_ATACADO) AS Select A.CODIGO,A.CODIGO_INTERNO,A.DESCRICAO,A.PRECO,A.UNIDADE, "
'        sql = sql & "B.SALDO_EM_ESTOQUE,A.ULTIMA_VENDA,A.ULTIMA_COMPRA,A.DATA_CAD_ALT, A.PRECO_ATACADO, A.PRECO_MINIMO_ATACADO FROM PRODUTO A, ESTOQUE B "
'        sql = sql & "WHERE A.Codigo = B.CODIGO_PRODUTO ORDER BY A.Descricao ASC "
'        Cnn.Execute sql
'        Call VerificaAtualizacoes
'    End If
'
'TRATA_VIEW_PRODUTO_DESCRIC_2:
''    If Err.Number <> 0 Then
''        Set Rstemp6 = Nothing
''        sql = "RECREATE VIEW VIEW_PRODUTO_DESCRIC (CODIGO,CODIGO_INTERNO,DESCRICAO,PRECO,UNIDADE,SALDO_EM_ESTOQUE, "
''        sql = sql & "ULTIMA_VENDA,ULTIMA_COMPRA,DATA_CAD_ALT) AS Select A.CODIGO,A.CODIGO_INTERNO,A.DESCRICAO,A.PRECO,A.UNIDADE, "
''        sql = sql & "B.SALDO_EM_ESTOQUE,A.ULTIMA_VENDA,A.ULTIMA_COMPRA,DATA_CAD_ALT FROM PRODUTO A, ESTOQUE B "
''        sql = sql & "WHERE A.Codigo = B.CODIGO_PRODUTO ORDER BY A.Descricao ASC "
''        Cnn.Execute sql
''        Call VerificaAtualizacoes
''    End If
'
'TRATA_TABELA_COMPRAS:
'    If Err.Number <> 0 Then
'        Set Rstemp6 = Nothing
'        sql = " CREATE TABLE COMPRA_PRODUTO (SEQUENCIA DOUBLE PRECISION NOT NULL, DATA_PED DATE,"
'        sql = sql & " CODIGO_EMPRESA  DOUBLE PRECISION,  CODIGO_FORNECEDOR   DOUBLE PRECISION,"
'        sql = sql & "OBS VARCHAR(100), VALOR_TOTAL DOUBLE PRECISION, PRIMARY KEY(SEQUENCIA))"
'        Cnn.Execute sql
'
''        sql = " CREATE TABLE RELCOMPRAS (SEQUENCIA DOUBLE PRECISION NOT NULL, CODIGO_PRODUTO  DOUBLE PRECISION, "
''        sql = sql & "  QTDE VARCHAR(15), VALOR_UNITARIO DOUBLE PRECISION, VALOR_TOTAL DOUBLE PRECISION "
''        Cnn.Execute sql
'
''        sql = " CREATE TABLE ITENS_COMPRA (SEQUENCIA DOUBLE PRECISION, CODIGO_PRODUTO DOUBLE PRECISION, "
''        sql = sql & "   QTDE  DOUBLE PRECISION, VALOR_UNITARIO DOUBLE PRECISION, VALOR_TOTAL DOUBLE PRECISION )"
''        Cnn.Execute sql
'        Call VerificaAtualizacoes
'    End If
'
'TRATA_PRODUTO_ALIQUOTA:
'
'    If Err.Number <> 0 Then
'        Set Rstemp6 = Nothing
'        sql = "ALTER TABLE PRODUTO ADD ALIQUOTA_ECF DOUBLE PRECISION "
'        Cnn.Execute sql
'        sql = "UPDATE PRODUTO SET ALIQUOTA_ECF = 5 "
'        Cnn.Execute sql
'        Call VerificaAtualizacoes
'    End If
'
'VIEW_ESTOQUE_NEG:
'
'    If Err.Number <> 0 Then
'        Set Rstemp6 = Nothing
'        sql = "CREATE VIEW VIEW_ESTOQUE_NEG (CODIGO_INTERNO,PRODUTO,SALDO_EM_ESTOQUE,MARCA,GRUPO) AS "
'        sql = sql & "SELECT PRODUTO.CODIGO_INTERNO,PRODUTO.DESCRICAO AS PRODUTO,ESTOQUE.SALDO_EM_ESTOQUE,"
'        sql = sql & " MARCAS.DESCRICAO AS MARCA,GRUPO.DESCRICAO AS GRUPO FROM ((PRODUTO INNER JOIN ESTOQUE ON PRODUTO.CODIGO = ESTOQUE.CODIGO_PRODUTO)"
'        sql = sql & " INNER JOIN MARCAS ON PRODUTO.MARCA = MARCAS.CODIGO)"
'        sql = sql & " INNER JOIN GRUPO ON PRODUTO.GRUPO = GRUPO.CODIGO"
'        sql = sql & " WHERE ESTOQUE.SALDO_EM_ESTOQUE <=0"
'        Cnn.Execute sql
'        Call VerificaAtualizacoes
'    End If
'
'VIEW_SO_ESTOQUE_NEG:
'    If Err.Number <> 0 Then
'        Set Rstemp6 = Nothing
'        sql = "CREATE VIEW VIEW_SO_ESTOQUE_NEG (CODIGO_INTERNO,NOME_PRODUTO,ULTIMA_VENDA,ULTIMA_COMPRA,SALDO_EM_ESTOQUE,QTD_MINIMA,COMPRAR,GRUPO) AS "
'
'        sql = sql & "Select B.CODIGO_INTERNO, B.DESCRICAO AS NOME_PRODUTO, B.ULTIMA_VENDA, B.ULTIMA_COMPRA, A.SALDO_EM_ESTOQUE, B.QTD_MINIMA,"
'        sql = sql & " (B.QTD_MINIMA - A.SALDO_EM_ESTOQUE) AS COMPRAR, G.DESCRICAO AS GRUPO FROM ESTOQUE A, PRODUTO B, "
'        sql = sql & " GRUPO G WHERE A.CODIGO_PRODUTO = B.CODIGO  "
'        sql = sql & " AND G.CODIGO = B.GRUPO "
'        sql = sql & " AND (A.SALDO_EM_ESTOQUE < B.QTD_MINIMA)  and (A.SALDO_EM_ESTOQUE >= 0)"
'        sql = sql & " ORDER BY NOME_PRODUTO "
'        Cnn.Execute sql
'        Call VerificaAtualizacoes
'    End If
'
'
'REL_ITENS_TRANSF_PROD:
'If Err.Number <> 0 Then
'    Set Rstemp6 = Nothing
'    sql = " CREATE TABLE REL_ITENS_TRANSF_PROD (SEQUENCIA DOUBLE PRECISION NOT NULL, CODIGO_PRODUTO DOUBLE PRECISION, QTDE DOUBLE PRECISION)"
'    Cnn.Execute sql
'    Call VerificaAtualizacoes
'End If
   

End Function


Public Function VerificaPermissaoExcluiPedido(ByVal Codigo_usuario As String)
  lbl_NmUsuario.Caption = frm_CdI_Usuario.txt_Nome
    
    If tipo = "I" Then
        sql = "select * from Cad_Menus order by MENU_DS_SISTEMA,MENU_CD_CODI"
        Set Rstemp = New ADODB.Recordset
        Rstemp.Open sql, Cnn, 1, 2
        With Rstemp
            .MoveLast
            .MoveFirst
            Spr_Menu.MaxRows = .RecordCount
            For X = 1 To .RecordCount
                Spr_Menu.Row = X
                Spr_Menu.Col = 2
                Spr_Menu.Text = !MENU_DS_SISTEMA
                Spr_Menu.Col = 3
                Spr_Menu.Text = Space(4 - Len(Format(!MENU_CD_CODI, "####"))) _
                                            & Format(!MENU_CD_CODI, "####") & _
                                            " - " & !MENU_DS_NOME_MOSTRA
                .MoveNext
            Next
        End With
    Else
        sql = "select * from Cad_Menus order by MENU_DS_SISTEMA,MENU_CD_CODI"
        Set Rstemp = New ADODB.Recordset
        Rstemp.Open sql, Cnn, 1, 2
        With Rstemp
            .MoveLast
            .MoveFirst
            Spr_Menu.MaxRows = .RecordCount
            For X = 1 To .RecordCount
                Spr_Menu.Row = X
                Spr_Menu.Col = 2
                Spr_Menu.Text = !MENU_DS_SISTEMA
                Spr_Menu.Col = 3
                Spr_Menu.Text = Space(4 - Len(Format(!MENU_CD_CODI, "####"))) _
                                            & Format(!MENU_CD_CODI, "####") & _
                                            " - " & !MENU_DS_NOME_MOSTRA
                sql = ""
                sql = "select * from Cad_Opcoes_Usuario_Acesso "
                sql = sql & " where OPAC_CD_CODI = " & frm_CdI_Usuario.lbl_NrCodigo
                sql = sql & " and MENU_CD_CODI = " & !MENU_CD_CODI
                sql = sql & " and MENU_DS_SISTEMA = '" & !MENU_DS_SISTEMA & "'"
                Set RsTemp1 = New ADODB.Recordset
                RsTemp1.Open sql, Cnn, 1, 2
                If RsTemp1.RecordCount <> 0 Then
                    Spr_Menu.Col = 1
                    Spr_Menu.Value = 1
                    Spr_Menu.Col = 4
                    Spr_Menu.Text = "S"
                End If
                RsTemp1.Close
                .MoveNext
            Next
        End With
    End If
    Rstemp.Close
    Set Rstemp = Nothing
End Function

Public Function VolumeSerialNumber(ByVal RootPath As String) As String
    
    Dim VolLabel As String
    Dim VolSize As Long
    Dim Serial As Long
    Dim MaxLen As Long
    Dim Flags As Long
    Dim Name As String
    Dim NameSize As Long
    Dim s As String

    If GetVolumeSerialNumber(RootPath, VolLabel, VolSize, Serial, MaxLen, Flags, Name, NameSize) Then
        'Create an 8 character string
        s = Format(Hex(Serial), "00000000")
        'Adds the '-' between the first 4 characters and the last 4 characters
        VolumeSerialNumber = Left(s, 4) + "-" + Right(s, 4)
    Else
        'If the call to API function fails the function returns a zero serial number
        VolumeSerialNumber = "0000-0000"
    End If

End Function

Sub WriteToErrorLog(sData As String, sFormNome As String, sRotina As String, sErro As String, intErroNumero As Integer, intNroPedido As Double)

'sData = Format(sData, "dd/mm/yyyy")
On Local Error Resume Next

Dim FileFree As Integer

FileFree = FreeFile
Open App.Path & "\ErrosPedidosLog.Txt" For Append As #FileFree
    Print #FileFree, sData, sFormNome, sRotina, sErro, intErroNumero, intNroPedido
Close #FileFree


End Sub

Sub ErrosGeraisLog(sData As String, sFormNome As String, sRotina As String, sErro As String, intErroNumero As Integer)

'sData = Format(sData, "dd/mm/yyyy")
On Local Error Resume Next

Dim FileFree As Integer

FileFree = FreeFile
Open App.Path & "\ErrosGeraisLog.Txt" For Append As #FileFree
    Print #FileFree, sData, sFormNome, sRotina, sErro, intErroNumero
Close #FileFree

End Sub


Public Sub ExportToHTML(ByVal TitleOfHTML As String, _
                        Optional TitleFont As String = "Tahoma", _
                        Optional HeaderFont As String = "Tahoma", _
                        Optional TitleFontSize As Byte = 5, _
                        Optional HeaderFontSize As Byte = 3, _
                        Optional TableBorder As Integer = 0, _
                        Optional CellPadding As Integer = 0, _
                        Optional CellSpacing As Integer = 5, _
                        Optional hexBodyBackground As String = "FFFFFF", _
                        Optional hexTitleBackground As String = "800000", _
                        Optional hexTitleForeground As String = "FFFFFF", _
                        Optional hexHeaderBackground As String = "FFFFEF", _
                        Optional hexHeaderForeground As String = "111111", _
                        Optional hexRecordsForeground As String = "111111", _
                        Optional hexTableBackground = "FFFFEF", _
                        Optional hexTableForeground = "111111", _
                        Optional hexBorderColor As String = "111111")

Dim TotalRecords As Long, i As Integer, NumberOfFields As Integer
Dim ErrorOccured As Boolean
Const Quote As String = """"

    On Error GoTo hell

'    With Progress
'        .Min = 0
'        .Max = ADODBRecordset.RecordCount
'        .Value = 0
'    End With

    Open ExportFilePath For Output Access Write As #1

    With ADODBRecordset
        .MoveFirst
        NumberOfFields = .Fields.Count - 1

        Print #1, "<HTML><HEAD><TITLE>" & TitleOfHTML & "</TITLE></HEAD>"
        Print #1, "<meta name=""GENERATEDBY"" content="" [HME] ADO Recordset Export Class "">"
        Print #1, "<meta name=""GENERATEDINFO"" content = "" www.elvista.cjb.net "">"
        Print #1, "<BODY BGCOLOR= " & Quote & hexBodyBackground & Quote & " Text = " & Quote & hexRecordsForeground & Quote & ">"
        Print #1, "<TABLE BORDER= " & Quote & TableBorder & Quote & " CellPadding = " & Quote & CellPadding & Quote & " CellSpacing = " & Quote & CellSpacing & Quote & " BODERCOLOR = " & hexBorderColor & " BGCOLOR = " & Quote & hexTableBackground & Quote & " Width = " & Quote & "100%" & Quote & ">"
        Print #1, "<TR><TD WIDTH=""100%"" COLSPAN=" & Quote & NumberOfFields + 1 & Quote & " BGCOLOR=" & Quote & hexTitleBackground & Quote & ">"
        Print #1, "<FONT COLOR = " & Quote & hexTitleForeground & Quote & "FACE=" & TitleFont & " SIZE=" & Quote & TitleFontSize & Quote & "><B>" & TitleOfHTML & "</B></FONT></TD></TR>"

        Print #1,
        Print #1, "<!-- Database Headers are are listed below -->"
        Print #1,

        Print #1, "     <TR>"        'First, add the Usual HTML Tags ^^^
        For i = 0 To NumberOfFields  'Now, add the titles to the file
            Print #1, "          <TD BGCOLOR=" & hexHeaderBackground & "><B>"
            Print #1, "          <FONT COLOR=" & hexTableForeground & Quote & " FACE=" & Quote & HeaderFont & Quote & " SIZE=" & Quote & HeaderFontSize & Quote & ">" & .Fields(i).Name & "</FONT></B></TD>"
        Next i
        Print #1, "     </TR>"

        Print #1,
       ' Print #1, "<!-- Database Records are are listed below -->"
        Print #1,

        Do While Not .EOF
            Print #1, "  <TR>"  'Add database records in HTML Format
            For i = 0 To NumberOfFields
                Print #1, "    <TD>" & .Fields(i) & "</TD>"
            Next i
            Print #1, "  </TR>"
            'Progress.Value = Progress.Value + 1
            .MoveNext
            
        Loop

    End With

    Print #1, "</TABLE></BODY></HTML>" 'Complete and close the HTML file
    Close #1



Exit Sub

hell:

    If Err.Number = 0 Then
        Resume Next
        ErrorOccured = True
    End If

End Sub

Sub DrawBarcode(ByVal bc_string As String, obj As Control)
    
    Dim xpos!, Y1!, Y2!, dw%, Th!, tw, new_string$
    
    'define barcode patterns
    Dim bc(90) As String
    bc(1) = "1 1221"            'pre-amble
    bc(2) = "1 1221"            'post-amble
    bc(48) = "11 221"           'digits
    bc(49) = "21 112"
    bc(50) = "12 112"
    bc(51) = "22 111"
    bc(52) = "11 212"
    bc(53) = "21 211"
    bc(54) = "12 211"
    bc(55) = "11 122"
    bc(56) = "21 121"
    bc(57) = "12 121"
                                'capital letters
    bc(65) = "211 12"           'A
    bc(66) = "121 12"           'B
    bc(67) = "221 11"           'C
    bc(68) = "112 12"           'D
    bc(69) = "212 11"           'E
    bc(70) = "122 11"           'F
    bc(71) = "111 22"           'G
    bc(72) = "211 21"           'H
    bc(73) = "121 21"           'I
    bc(74) = "112 21"           'J
    bc(75) = "2111 2"           'K
    bc(76) = "1211 2"           'L
    bc(77) = "2211 1"           'M
    bc(78) = "1121 2"           'N
    bc(79) = "2121 1"           'O
    bc(80) = "1221 1"           'P
    bc(81) = "1112 2"           'Q
    bc(82) = "2112 1"           'R
    bc(83) = "1212 1"           'S
    bc(84) = "1122 1"           'T
    bc(85) = "2 1112"           'U
    bc(86) = "1 2112"           'V
    bc(87) = "2 2111"           'W
    bc(88) = "1 1212"           'X
    bc(89) = "2 1211"           'Y
    bc(90) = "1 2211"           'Z
                                'Misc
    bc(32) = "1 2121"           'space
    bc(35) = ""                 '# cannot do!
    bc(36) = "1 1 1 11"         '$
    bc(37) = "11 1 1 1"         '%
    bc(43) = "1 11 1 1"         '+
    bc(45) = "1 1122"           '-
    bc(47) = "1 1 11 1"         '/
    bc(46) = "2 1121"           '.
    bc(64) = ""                 '@ cannot do!
    bc(65) = "1 1221"           '*
    
    
    
    bc_string = UCase(bc_string)
    
    
    'dimensions
    obj.ScaleMode = 3                               'pixels
    obj.Cls
    obj.Picture = Nothing
    dw = CInt(obj.ScaleHeight / 40)                 'space between bars
    If dw < 1 Then dw = 1
    'Debug.Print dw
    Th = obj.TextHeight(bc_string)                  'text height
    tw = obj.TextWidth(bc_string)                   'text width
    new_string = Chr$(1) & bc_string & Chr$(2)      'add pre-amble, post-amble
    
    Y1 = obj.ScaleTop
    Y2 = obj.ScaleTop + obj.ScaleHeight - 1.5 * Th
    obj.Width = 1.1 * Len(new_string) * (15 * dw) * obj.Width / obj.ScaleWidth
    
    
    'draw each character in barcode string
    xpos = obj.ScaleLeft
    For n = 1 To Len(new_string)
        c = Asc(Mid$(new_string, n, 1))
        If c > 90 Then c = 0
        bc_pattern$ = bc(c)
        
        'draw each bar
        For i = 1 To Len(bc_pattern$)
            Select Case Mid$(bc_pattern$, i, 1)
                Case " "
                    'space
                    obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                    xpos = xpos + dw
                    
                Case "1"
                    'space
                    obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                    xpos = xpos + dw
                    'line
                    obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &H0&, BF
                    xpos = xpos + dw
                
                Case "2"
                    'space
                    obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                    xpos = xpos + dw
                    'wide line
                    obj.Line (xpos, Y1)-(xpos + 2 * dw, Y2), &H0&, BF
                    xpos = xpos + 2 * dw
            End Select
        Next
    Next
    
    '1 more space
    obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
    xpos = xpos + dw
    
    'final size and text
    obj.Width = (xpos + dw) * obj.Width / obj.ScaleWidth
    obj.CurrentX = (obj.ScaleWidth - tw) / 2
    obj.CurrentY = Y2 + 0.25 * Th
    'obj.Print bc_string
    
    'copy to clipboard
    obj.Picture = obj.Image
    Clipboard.Clear
    Clipboard.SetData obj.Image, 2



End Sub

Public Function Extenso(nvalor)
'Valida Argumento
If IsNull(nvalor) Or nvalor <= 0 Or nvalor > 9999999.99 Then
   Exit Function
End If

'Variáveis
Dim nContador, nTamanho As Integer
Dim cValor, cParte, cFinal As String
ReDim aGrupo(4), aTexto(4) As String
'Matrizes de extensos (Parciais)
ReDim aUnid(19) As String
aUnid(1) = "um ": aUnid(2) = "dois ": aUnid(3) = "tres "
aUnid(4) = "quatro ": aUnid(5) = "cinco ": aUnid(6) = "seis "
aUnid(7) = "sete ": aUnid(8) = "oito ": aUnid(9) = "nove "
aUnid(10) = "dez ": aUnid(11) = "onze ": aUnid(12) = "doze "
aUnid(13) = "treze ": aUnid(14) = "quatorze ": aUnid(15) = "quinze "
aUnid(16) = "dezesseis ": aUnid(17) = "dezessete ": aUnid(18) = "dezoito "
aUnid(19) = "dezenove "

ReDim aDezena(9) As String
aDezena(1) = "dez ": aDezena(2) = "vinte ": aDezena(3) = "trinta "
aDezena(4) = "quarenta ": aDezena(5) = "cinquenta "
aDezena(6) = "sessenta ": aDezena(7) = "setenta ": aDezena(8) = "oitenta "
aDezena(9) = "noventa "

ReDim aCentena(9) As String
aCentena(1) = "cento ": aCentena(2) = "duzentos "
aCentena(3) = "trezentos ": aCentena(4) = "quatrocentos "
aCentena(5) = "quinhentos ": aCentena(6) = "seiscentos "
aCentena(7) = "setecentos ": aCentena(8) = "oitocentos "
aCentena(9) = "novecentos "

'Separa valor em grupos
cValor = Format$(nvalor, "0000000000.00")
aGrupo(1) = Mid$(cValor, 2, 3)
aGrupo(2) = Mid$(cValor, 5, 3)
aGrupo(3) = Mid$(cValor, 8, 3)
aGrupo(4) = "0" + Mid$(cValor, 12, 2)

'Calcula cada grupo
For nContador = 1 To 4
cParte = aGrupo(nContador)
nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
If nTamanho = 3 Then
If Right$(cParte, 2) <> "00" Then
aTexto(nContador) = aTexto(nContador) + aCentena(Left(cParte, 1)) + "e "
nTamanho = 2
Else
aTexto(nContador) = aTexto(nContador) + IIf(Left$(cParte, 1) = "1", "cem ", aCentena(Left(cParte, 1)))
End If
End If
If nTamanho = 2 Then
If Val(Right(cParte, 2)) < 20 Then
aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 2))
Else
aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
If Right$(cParte, 1) <> "0" Then
aTexto(nContador) = aTexto(nContador) + "e "
nTamanho = 1
End If
End If
End If
If nTamanho = 1 Then
aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 1))
End If
Next

'Final
If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
cFinal = aTexto(4) + IIf(Val(aGrupo(4)) = 1, "centavo", "centavos")
Else
cFinal = ""
cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, aTexto(1) + IIf(Val(aGrupo(1)) > 1, "milhões ", "milhão "), "")
If Val(aGrupo(2) + aGrupo(3)) = 0 Then
cFinal = cFinal + "de "
Else
cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, aTexto(2) + "mil ", "")
End If
cFinal = cFinal + aTexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "real ", "reais ")
cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "E " + aTexto(4) + IIf(Val(aGrupo(4)) = 1, "centavo", "centavos"), "")
End If
Extenso = UCase$(cFinal)

End Function


Function ValidaCartaoCredito(CCNumber As String) As Boolean
  Dim Counter As Integer, TmpInt As Integer
  Dim Answer As Integer
  'Dim IsEven As Integer
  Counter = 1
  TmpInt = 0
    While Counter <= Len(CCNumber)
        If IsNumeric(Len(CCNumber)) Then
            TmpInt = Val(Mid$(CCNumber, Counter, 1))
            If Not IsNumeric(Counter) Then
                TmpInt = TmpInt * 2
                If TmpInt > 9 Then TmpInt = TmpInt - 9
            End If
            Answer = Answer + TmpInt
            Counter = Counter + 1
        Else
            TmpInt = Val(Mid$(CCNumber, Counter, 1))
            If IsNumeric(Counter) Then
                TmpInt = TmpInt * 2
                If TmpInt > 9 Then TmpInt = TmpInt - 9
            End If
            Answer = Answer + TmpInt
            Counter = Counter + 1
        End If
    Wend
    Answer = Answer Mod 10
    If Answer = 0 Then ValidaCartaoCredito = True
End Function
Public Sub PrintCenter(PrintString$)
   'print the string in the center of the page
   Printer.CurrentX = (Printer.ScaleWidth / 2) - ((Printer.FontSize * _
         (Printer.TextWidth(PrintString$) / 8.28)) / 2)
   'where the 8.28 is the PC
   'default font size   (where the width of the letters comnes from)
   Printer.Print PrintString$
End Sub

Public Sub AddIconToMenu()
    Dim v_lMenuHnd    As Long
    Dim v_lSubMenuHnd As Long
    Dim v_lMenuCnt    As Long
    Dim v_lSubMenuCnt As Long
    Dim v_lSubMenuID  As Long

    v_lMenuHnd = GetMenu(frmMenu.hwnd)
    v_lMenuCnt = GetMenuItemCount(lMenuHnd)

    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 0)
    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, _
                            frmMenu.iml_Menu.ListImages(1).Picture, frmMenu.iml_Menu.ListImages(1).Picture)
        
'    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
'    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 1)
'    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frmMenu.iml_Menu.ListImages(2).Picture, frmMenu.iml_Menu.ListImages(2).Picture)

'    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
'    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 2)
'    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frmMenu.iml_Menu.ListImages(3).Picture, frmMenu.iml_Menu.ListImages(3).Picture)
'
'    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
'    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 3)
'    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frmMenu.iml_Menu.ListImages(4).Picture, frmMenu.iml_Menu.ListImages(4).Picture)
'
'    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
'    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 4)
'    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frmMenu.iml_Menu.ListImages(5).Picture, frmMenu.iml_Menu.ListImages(5).Picture)
'
'    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
'    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 5)
'    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frmMenu.iml_Menu.ListImages(6).Picture, frmMenu.iml_Menu.ListImages(6).Picture)
'
'    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
'    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 6)
'    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frmMenu.iml_Menu.ListImages(7).Picture, frmMenu.iml_Menu.ListImages(7).Picture)
'
    'menu movimentação
'    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 2)
'    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 0)
'    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BYCOMMAND, frmMenu.iml_Menu.ListImages(1).Picture, frmMenu.iml_Menu.ListImages(1).Picture)
    
'    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 2)
'    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 2)
'    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BYCOMMAND, frmMenu.iml_Menu.ListImages(2).Picture, frmMenu.iml_Menu.ListImages(2).Picture)
'
'    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 2)
'    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 2)
'    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frmMenu.iml_Menu.ListImages(10).Picture, frmMenu.iml_Menu.ListImages(11).Picture)
End Sub


' CONVERT STRING EM NUMERICO
Function Convert_Numeric(p As String, IsMoney As Boolean) As String
    If p <> "" Then
        If IsNumeric(p) Then
            If IsMoney Then
                Convert_Numeric = Format(p, "#,###,###,##0.00")
            Else
                Convert_Numeric = Format(p, "0.00")
            End If
        Else
            MsgBox "Valor numérico invalido."
            
        End If
    Else
        Convert_Numeric = "0"
    End If
End Function

Public Sub UnloadAllForms(Optional sFormName As String = "")
    Dim Form As Form
    For Each Form In Forms
        If Form.Name <> sFormName Then
            Unload Form
            Set Form = Nothing
        End If
    Next Form
End Sub


Public Sub SpoolFile(sFile As String, PrnName As String, Optional AppName As String = "")
   Dim hPrn As Long
   Dim Buffer() As Byte
   Dim hFile As Integer
   Dim Written As Long
   Dim di As DOC_INFO_1
   Dim i As Long
   Const BufSize As Long = &H4000
   '
   ' Extract filename from passed spec, and build job name.
   ' Fill remainder of DOC_INFO_1 structure.
   '
    If InStr(sFile, "\") Then
       For i = Len(sFile) To 1 Step -1
          If Mid(sFile, i, 1) = "\" Then Exit For
          di.pDocName = Mid(sFile, i, 1) & di.pDocName
       Next i
    Else
       di.pDocName = sFile
    End If
    If Len(AppName) Then
       di.pDocName = AppName & ": " & di.pDocName
    End If
    di.pOutputFile = vbNullString
    di.pDatatype = "RAW"

  ' Call OpenPrinter(PrnName, hPrn, vbNullString)
  ' Call StartDocPrinter(hPrn, 1, di)
  ' Call StartPagePrinter(hPrn)
  ' Call EndPagePrinter(hPrn)
   'Call EndDocPrinter(hPrn)
   'Call ClosePrinter(hPrn)
End Sub


Public Sub SelecionaImpressoraAtiva(Lst As ComboBox)
   Dim sRet As String
   Dim nRet As Integer
   Dim i As Integer
   '
   ' Look for default printer in WIN.INI
   '
   sRet = Space(255)
   nRet = GetProfileString("Windows", ByVal "device", "", _
                           sRet, Len(sRet))
   '
   ' Truncate default printer name.
   '
   If nRet Then
      sRet = UCase(Left(sRet, InStr(sRet, ",") - 1))
      '
      ' Cycle list looking for matching entry.
      '
      For i = 0 To Lst.ListCount
         If Left(UCase(Lst.List(i)), Len(sRet)) = sRet Then
            '
            ' Found it. Set index and bail.
            '
            Lst.ListIndex = i
            Exit For
         End If
      Next i
   End If
End Sub

Public Sub RemoveMenus(ByVal ObjForm As Form)
   Dim hMenu As Long
   ' Get the form's system menu handle.
   hMenu = GetSystemMenu(ObjForm.hwnd, False)
   DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

Function Pad_Str(str As String, val_to_pad As String, strlength As Integer, Right As Boolean) As String
    Dim s1 As String
    s1 = ""
    For i = 1 To strlength - Len(str) Step 1
        s1 = s1 & val_to_pad
    Next i
    If Right Then
        Pad_Str = str & s1
    Else
        Pad_Str = s1 & str
    End If
End Function

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


Public Function FormatProduto(DescriProduto As String)
Dim NroDescriProduto      As String
Dim cont        As Integer

NroDescriProduto = ""

For cont = 1 To Len(DescriProduto)
    If Mid(DescriProduto, cont, 1) <> "-" Then
        NroDescriProduto = NroDescriProduto & Mid(DescriProduto, cont, 1)
    End If
Next
DescriProduto = NroDescriProduto

If DescriProduto <> "" And Len(DescriProduto) <= 60 Then
    'Produto = Right(DescriProduto, Len(DescriProduto) - 4) & "-" & Right(DescriProduto, 4)
    PRODUTO = Left(DescriProduto, Len(DescriProduto) + 10) & "-" & Len(DescriProduto) + 4
End If

End Function
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
Public Function ValidaCPF(CPF As String) As Boolean
   Dim EVAR1 As Integer
   Dim evar2 As Integer
   Dim F As Integer
   If Len(Trim(CPF)) <> 11 Then
      ValidaCPF = False
      Exit Function
   End If
   EVAR1 = 0
   For F = 1 To 9
      EVAR1 = EVAR1 + Val(Mid(CPF, F, 1)) * (11 - F)
   Next F
   evar2 = 11 - (EVAR1 - (Int(EVAR1 / 11) * 11))
   If evar2 = 10 Or evar2 = 11 Then evar2 = 0
   If evar2 <> Val(Mid(CPF, 10, 1)) Then
      ValidaCPF = False
      Exit Function
   End If
   EVAR1 = 0
   For F = 1 To 10
       EVAR1 = EVAR1 + Val(Mid(CPF, F, 1)) * (12 - F)
   Next F
   evar2 = 11 - (EVAR1 - (Int(EVAR1 / 11) * 11))
   If evar2 = 10 Or evar2 = 11 Then evar2 = 0
   If evar2 <> Val(Mid(CPF, 11, 1)) Then
      ValidaCPF = False
      Exit Function
  End If
  ValidaCPF = True
End Function

Public Function ConsisteTeclaValorNumerico(pValor As String, pTecla As Integer) As Integer
   If Not IsNumeric(Chr(pTecla)) Then
      If pTecla = 46 Or pTecla = 44 Then
         pTecla = 44
         For i = 1 To Len(pValor)
            If Mid(pValor, i, 1) = "," Then
               pTecla = 0
               Exit For
            End If
         Next
      Else
         If pTecla <> 8 Then pTecla = 0
      End If
   End If
   ConsisteTeclaValorNumerico = pTecla
End Function

Public Function ConsisteTeclaValorNumericoParaTrocaProdutos(pValor As String, pTecla As Integer) As Integer
    If Not IsNumeric(Chr(pTecla)) Then
        If pTecla = 46 Or pTecla = 44 Then
            pTecla = 44
            For i = 1 To Len(pValor)
                If Mid(pValor, i, 1) = "," Then
                    pTecla = 0
                    Exit For
                End If
            Next
        ElseIf pTecla = 45 Then
            pTecla = 45
            For i = 1 To Len(pValor)
                If Mid(pValor, i, 1) = "-" Then
                    pTecla = 0
                    Exit For
                End If
            Next
        Else
            If pTecla <> 8 Then pTecla = 0
        End If
    End If
    ConsisteTeclaValorNumericoParaTrocaProdutos = pTecla
End Function
Function ValidaCGC(CGC As String) As Integer
        Dim Retorno, a, j, i, d1, d2
        If Len(CGC) = 8 And Val(CGC) > 0 Then
           a = 0
           j = 0
           d1 = 0
           For i = 1 To 7
               a = Val(Mid(CGC, i, 1))
               If (i Mod 2) <> 0 Then
                  a = a * 2
               End If
               If a > 9 Then
                  j = j + Int(a / 10) + (a Mod 10)
               Else
                  j = j + a
               End If
           Next i
           d1 = IIf((j Mod 10) <> 0, 10 - (j Mod 10), 0)
           If d1 = Val(Mid(CGC, 8, 1)) Then
              ValidaCGC = True
           Else
              ValidaCGC = False
           End If
        Else
           If Len(CGC) = 14 And Val(CGC) > 0 Then
              a = 0
              i = 0
              d1 = 0
              d2 = 0
              j = 5
              For i = 1 To 12 Step 1
                  a = a + (Val(Mid(CGC, i, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next i
              a = a Mod 11
              d1 = IIf(a > 1, 11 - a, 0)
              a = 0
              i = 0
              j = 6
              For i = 1 To 13 Step 1
                  a = a + (Val(Mid(CGC, i, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next i
              a = a Mod 11
              d2 = IIf(a > 1, 11 - a, 0)
              If (d1 = Val(Mid(CGC, 13, 1)) And d2 = Val(Mid(CGC, 14, 1))) Then
                 ValidaCGC = True
              Else
                 ValidaCGC = False
              End If
           Else
              ValidaCGC = False
           End If
        End If
End Function

' Carrega Nome colunas list view
 Public Function Carrega_Nome_colunas_LIST_DETALHES(ByVal LIST_DETALHES As ListView) As Boolean
    Dim ListviewHeader  As MSComctlLib.ColumnHeader
    Dim lvListItems As MSComctlLib.ListItem
    
    'Clear the Listview Control
    LIST_DETALHES.ListItems.Clear
    
    Set ListviewHeader = Nothing
    
    Set ListviewHeader = LIST_DETALHES.ColumnHeaders.Add(, "C1", "Cod.Movto.", 1000, lvwColumnLeft)
    Set ListviewHeader = LIST_DETALHES.ColumnHeaders.Add(, "C2", "Vencimento", 1100, lvwColumnLeft)
    Set ListviewHeader = LIST_DETALHES.ColumnHeaders.Add(, "C3", "Dias", 600, lvwColumnLeft)
    Set ListviewHeader = LIST_DETALHES.ColumnHeaders.Add(, "C4", "Valor Bruto", 1100, lvwColumnRight)
    Set ListviewHeader = LIST_DETALHES.ColumnHeaders.Add(, "C5", "Banco", 800, lvwColumnLeft)
    Set ListviewHeader = LIST_DETALHES.ColumnHeaders.Add(, "C6", "Agencia", 800, lvwColumnLeft)
    Set ListviewHeader = LIST_DETALHES.ColumnHeaders.Add(, "C7", "Conta", 1000, lvwColumnLeft)
    Set ListviewHeader = LIST_DETALHES.ColumnHeaders.Add(, "C8", "Nro.Cheque", 1100, lvwColumnLeft)
    Set ListviewHeader = LIST_DETALHES.ColumnHeaders.Add(, "C9", "Vlr. Líquido", 1100, lvwColumnRight)
    Set ListviewHeader = LIST_DETALHES.ColumnHeaders.Add(, "C10", "Nome", 2500, lvwColumnLeft)
    Set ListviewHeader = LIST_DETALHES.ColumnHeaders.Add(, "C11", "CPF/CNPJ", 1500, lvwColumnLeft)
    
    
    LIST_DETALHES.View = lvwReport
    Carrega_Nome_colunas_LIST_DETALHES = True
End Function


Public Sub LimpaControles(frmAtivo As Form)
Dim MyCtrls As Object
    
    If frmAtivo.Name = "frmCdPgto" Then
        For Each MyCtrls In frmAtivo.Controls
            If TypeOf MyCtrls Is TextBox Then
                MyCtrls.Text = "0.00"
                MyCtrls.Enabled = False
            End If
        Next MyCtrls
        Exit Sub
    End If
    
    For Each MyCtrls In frmAtivo.Controls
         If TypeOf MyCtrls Is TextBox Then
                MyCtrls.Text = ""
         ElseIf TypeOf MyCtrls Is ComboBox Then
                MyCtrls.ListIndex = -1
         End If
    Next MyCtrls
    
End Sub

Public Sub HablilitaLimpaControles(frmAtivo As Form)
Dim MyCtrls As Object
    
    For Each MyCtrls In frmAtivo.Controls
         If TypeOf MyCtrls Is TextBox Then
             MyCtrls.Text = ""
             MyCtrls.Enabled = True
         End If
    Next MyCtrls

End Sub


Public Sub SetEnabledControles(frmAtivo As Form)
Dim MyCtrls As Object
    
    For Each MyCtrls In frmAtivo.Controls
        If TypeOf MyCtrls Is TextBox Then
           'MyCtrls.Text = ""
            MyCtrls.Enabled = False
        ElseIf TypeOf MyCtrls Is ComboBox Then
            MyCtrls.Enabled = False
        ElseIf TypeOf MyCtrls Is Toolbar Then
            MyCtrls.Enabled = False
        ElseIf TypeOf MyCtrls Is ListBox Then
            MyCtrls.Enabled = False
         ElseIf TypeOf MyCtrls Is Frame Then
            If MyCtrls.Name <> "FmeTpoImpressora" Then
                MyCtrls.Enabled = False
            End If
        ElseIf TypeOf MyCtrls Is CommandButton Then
            If MyCtrls.Name <> "cmd_Imprimir" And MyCtrls <> "cmd_Cancelar" Then
                MyCtrls.Enabled = False
            End If
        End If
    Next
    
End Sub

Public Sub SetDesabledControles(frmAtivo As Form)
Dim MyCtrls As Object
    
    For Each MyCtrls In frmAtivo.Controls
        If TypeOf MyCtrls Is TextBox Then
            MyCtrls.Enabled = True
        ElseIf TypeOf MyCtrls Is ComboBox Then
            MyCtrls.Enabled = True
        ElseIf TypeOf MyCtrls Is Toolbar Then
            MyCtrls.Enabled = True
        ElseIf TypeOf MyCtrls Is ListBox Then
            MyCtrls.Enabled = True
        ElseIf TypeOf MyCtrls Is Frame Then
            MyCtrls.Enabled = True
        End If
    Next
    
End Sub
Public Function SemFormatoCPF_CNPJ(CPF_CNPJ As String)
Dim cont        As Integer

cont = 1

For cont = 1 To Len(CPF_CNPJ)
    If InStr("0123456789", Mid(CPF_CNPJ, cont, 1)) > 0 Then
        SemFormatoCPF_CNPJ = SemFormatoCPF_CNPJ & Mid(CPF_CNPJ, cont, 1)
    End If
Next

End Function

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

Function Troca_Virg_Zero(campo As Variant)
    Dim xcont As Long
   ' For xcont = 1 To Len(campo)
   '     If InStr(1, campo, ",") <> 0 Then
   '        campo = Mid(campo, 1, InStr(1, campo, ",") - 1) & "." & Mid(campo, InStr(1, campo, ",") + 1, Len(campo))
   '     End If
   ' Next
    'Troca_Virg_Zero = campo
    
    Troca_Virg_Zero = Replace(campo, ",", ".")
End Function


'Public Function ListaArquivosRetorno(strPath As String, Optional Extention As String) As Boolean
Public Function ListaArquivosRetorno() As Boolean
    ListaArquivosRetorno = False
    
    strPath = App.Path & "\Retorno\"
    
    Dim File As String
    If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    If Trim$(Extention) = "" Then
        Extention = "*.*"
    ElseIf Left$(Extention, 2) <> "*." Then
        Extention = "*." & Extention
    End If
   
    File = Dir$(strPath & Extention)
   
    Do While Len(File)
        File = Dir$
        ListaArquivosRetorno = True
        Exit Do
    Loop
End Function
Public Sub Main()

IP_Servidor = ""

    'IP_Servidor = ReadIniFile("c:\Sisadven.ini", "IP", "", "")

    Versao_Software = App.Major & "." & App.Minor

   ''' StrNomeMaquina = GetIPHostName()
   ''' IP_Servidor = GetIPAddress()
     'IP_Servidor = IP_Servidor + "\Arqdados.GDB"
    
    '''IP_Servidor_Relatorios = ReadIniFile("c:\Sisadven.ini", "IP", "", "")
    
    On Error GoTo Erro
            
    Set Cnn = New ADODB.Connection
    
    With Cnn
        .CursorLocation = adUseClient
        '.Open "DRIVER=Firebird/InterBase(r) driver; UID=SYSDBA; PWD=masterkey;DBNAME=servidor:" & App.Path & "\arqdados.GDB"
        .Open "File Name=" & App.Path & "\cnn_fire_Servidor.udl;"
        '.Open "Provider=IBOLE.Provider.v4;Password=masterkey;Persist Security Info=True;Data Source=servidor:c:\Sistema SisAdven\ARQDADOS.GDB;Mode=ReadWrite|Share Deny None"
    End With

    Call ChecaRegistro
        
Exit Sub
    
Erro:
        If Err Then
            MsgBox "Impossível Abrir o Banco ! Erro: " & Err.Number & " - " & Err.Description & Chr(13) + Chr(10) & _
                   Chr(13) + Chr(10) & _
                   "Este Erro pode ter sido causado pelos seguintes motivos: " & Chr(13) + Chr(10) & Chr(13) + Chr(10) & _
                   "1.    O Servidor pode estar desligado, ou não conectado na rede." & Chr(13) + Chr(10) & _
                   "2.    O Hub (Conector da rede) pode estar desligado ou cabos de rede desligados." & Chr(13) + Chr(10) & _
                   "3.    Seu micro não está conectado à rede. Verifique o Cabo de Rede (Azul), e tente reiniciar o computador." & Chr(13) + Chr(10) & _
                   "4.    Se nenhuma destas possibilidades funcionarem, entre em contato com o suporte técnico do sistema.", vbInformation, "Aviso"
            'Unload frmMenu
            Call Fecha_Formularios
            End
        Else
            MsgBox Err.Number & " descrição" & Err.Description
            Unload frmMenu
            End
        End If
Err.Clear

 End Sub


Public Sub CompactJetDatabase(Location As String, Optional BackupOriginal As Boolean = True)


'''Dim strBackupFile As String
'''Dim strTempFile As String
'''
''''Check the database exists
'''If Len(Dir(Location)) Then
'''
'''    ' If a backup is required, do it!
'''    If BackupOriginal = True Then
'''        strBackupFile = GetTemporaryPath & "backup.mdb"
'''        If Len(Dir(strBackupFile)) Then Kill strBackupFile
'''        FileCopy Location, strBackupFile
'''    End If
'''
'''    ' Create temporary filename
'''    strTempFile = GetTemporaryPath & "temp.mdb"
'''    If Len(Dir(strTempFile)) Then Kill strTempFile
'''
'''    ' Do the compacting via DBEngine
'''    DBEngine.CompactDatabase Location, strTempFile
'''
'''    ' Remove the original database file
'''    Kill Location
'''
'''    ' Copy the temporary now-compressed
'''    ' database file back to the original
'''    ' location
'''    FileCopy strTempFile, Location
'''
'''    ' Delete the temporary file
'''    Kill strTempFile
'''
'''Else
'''
'''End If
'''
'''Exit Sub
'''
'''CompactErr:
'''    Err.Clear

    
Dim strBackupFile As String
Dim strTempFile As String

'Check the database exists
If Len(Dir(Location)) Then

    ' If a backup is required, do it!
    If BackupOriginal = True Then
        strBackupFile = GetTemporaryPath & "backup.mdb"
        If Len(Dir(strBackupFile)) Then Kill strBackupFile
        FileCopy Location, strBackupFile
    End If

    ' Create temporary filename
    strTempFile = GetTemporaryPath & "temp.mdb"
    If Len(Dir(strTempFile)) Then Kill strTempFile

    ' Do the compacting via DBEngine
    DBEngine.CompactDatabase Location, strTempFile, , , ";pwd=;"

    ' Remove the original database file
    Kill Location

    ' Copy the temporary now-compressed
    ' database file back to the original
    ' location
    FileCopy strTempFile, Location

    ' Delete the temporary file
    Kill strTempFile
    Kill strBackupFile

Else

End If

CompactErr:
    
    Exit Sub


End Sub

Function CompactarRepararDatabase(DatabasePath As String, _
Optional Password As String, Optional TempFile As String = "c:\temp.mdb")

'se a versão DAO for anterior a 3.6 , então devemos usar o método RepairDatabase
'se a versao DAO for a 3.6 ou superior basta usar a função CompactDatabase
If DBEngine.Version < "3.6" Then DBEngine.RepairDatabase DatabasePath

'se nao informou um arquivo temporario usa "c:\temp.mdb"
If TempFile = "" Then TempFile = "c:\temp.mdb"

'apaga o arquivo temp se existir
If Dir(TempFile) <> "" Then Kill TempFile

'formata a senha no formato ";pwd=PASSWORD" se a mesma existir
If Password <> "" Then Password = ";pwd=" & Password

'compacta a base criando um novo banco de dados
DBEngine.CompactDatabase DatabasePath, TempFile, , , Password

'apaga o primeiro banco de dados
Kill DatabasePath

'move a base compactada para a origem
FileCopy TempFile, DatabasePath

'apaga o arquivo temporario
Kill TempFile

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



Public Function IsValidaEmail(ByVal endEmail As String) As Boolean
   IsValidaEmail = endEmail Like "*@[A-Z,a-z,0-9]*.*"
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


Sub SelText(object As Control)
    
    With object
        .SelStart = 0
        .SelLength = Len(object)
    End With

End Sub

Function Resolucao(pixellargura As Long, pixelaltura As Long) As Boolean

    Dim TwipsL As Long
    Dim TwipsA As Long

    ' converte pixels para twips
    TwipsL = pixellargura * 15
    TwipsA = pixelaltura * 15

    ' verifica comparando com a resolução atual
    If TwipsL <> Screen.Width Then   'Width = largura
        Resolucao = False
    Else
        If TwipsA <> Screen.Height Then 'Height = altura
    Else
        Resolucao = True
        End If
    End If
End Function

Public Sub CorFoco(obj As Object)

obj.ForeColor = &HFF

End Sub



Function GetInitVar(lpAppName As String, lpKeyName As String, lpDefault As Variant) As String
    Dim lpReturnedString As String
    Dim nSize As Integer
    Dim lpFilename As String
    Dim nStringSize As Integer
    
    lpReturnedString = Space(81)
    nSize = 81
    lpFilename = winDir$ + "Cardbrad.INI"
    nStringSize = GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFilename)
    GetInitVar = Left$(lpReturnedString, nStringSize)
End Function

Function WriteIniFile(ByVal sIniFileName As String, ByVal sSection As String, ByVal sItem As String, ByVal sText As String) As Boolean
   Dim i As Integer
   On Error GoTo sWriteIniFileError

   i = WritePrivateProfileString(sSection, sItem, sText, sIniFileName)
   WriteIniFile = True

   Exit Function
sWriteIniFileError:
    WriteIniFile = False
End Function


Function Executa_Storead_Procedure(Tabela As String) As Integer
'    Dim objCmd As New ADODB.Command
'    Dim objRS As New ADODB.Recordset
'    Set Banco = New cBanco
'    Banco.AbrirBanco
'
'    With objCmd
'        .ActiveConnection = Banco.Conexao
'        .CommandType = adCmdStoredProc
'        .CommandText = "minhaProcedure"
'        .Parameters.Refresh
'        .Parameters(0) = UCase$(Left$(tabela, 3))
'        .Prepared = True
'    End With
'
'    Set objRS = objCmd.Execute
'    PegaCodigo = objRS(0).Value
'
'    Set objRS = Nothing
'    objCmd.Prepared = False
'    Set objCmd = Nothing
'
'    Banco.FecharBanco
'    Set Banco = Nothing
End Function
Function ReadIniFile(ByVal sIniFileName As String, ByVal sSection As String, ByVal sItem As String, ByVal sDefault As String) As String
   Dim iRetAmount As Integer   'the amount of characters returned
   Dim sTemp As String

   sTemp = String$(400, 0)   'fill with nulls
   iRetAmount = GetPrivateProfileString(sSection, sItem, sDefault, sTemp, 400, sIniFileName)
   sTemp = Left$(sTemp, iRetAmount)
   ReadIniFile = sTemp
End Function
'
Public Sub CorPerdeFoco(obj As Object)

obj.ForeColor = &H800000

End Sub

Function ChecaRegistro()

Load FrmRegistro

With FrmRegistro.ActiveLock1
    'Se o usuário é registrado exibe o formulário principal
    Situacao_Registro = .RegisteredUser
    Dias_Uso_Sistema = .UsedDays
    FrmRegistro.lblQtdDias.Caption = Dias_Uso_Sistema
    
    If .RegisteredUser Then
        '''frmAcesso.Show 1
        Unload FrmRegistro
        frmMenu.Show
        frmAcesso.Show 1, frmMenu
    Else
        'Se o usuário não esta registrado, verifica se foi alterada a data do sistema
        If .LastRunDate > Now Then
            MsgBox "Atenção! Foi detectado Violação na Data do Sistema !..." & vbNewLine & vbNewLine & "Corrija a Data do Sistema para Continuar...!", vbCritical, "Aviso"
            Call Fecha_Formularios
            End
        End If
            
    Select Case .UsedDays
        Case 85
            MsgBox "Faltam 5 (Cinco) Dias para o Sistema Expirar...!" & Chr(13) & "Entre em contato com Novavia Soluções em Informática" & Chr(13) & Chr(13) & "Fone : " & "(11) 5548-3890", 32, "Aviso de Registro"
        Case 86
            MsgBox "Faltam 4 (Quatro) Dias para o Sistema Expirar...!" & Chr(13) & "Entre em contato com Novavia Soluções em Informática " & Chr(13) & "Fone : " & "(11) 5548-3890", 32, "Aviso de Registro"
        Case 87
            MsgBox "Faltam 3 (Três) Dias para o Sistema Expirar...!" & Chr(13) & "Entre em contato com Novavia Soluções em Informática " & Chr(13) & "Fone : " & "(11) 5548-3890", 32, "Aviso de Registro"
        Case 88
            MsgBox "Faltam 2 (Dois) Dias para o Sistema Expirar...!" & Chr(13) & "Entre em contato com Novavia Soluções em Informática " & Chr(13) & "Fone : " & "(11) 5548-3890", 32, "Aviso de Registro"
        Case 89
            MsgBox "Falta 1 (Um) Dia para o Sistema Expirar...!" & Chr(13) & "Entre em contato com Novavia Soluções em Informática " & Chr(13) & "Fone : " & "(11) 5548-3890", 32, "Aviso de Registro"
        Case 90
            MsgBox "Hoje é o Último Dia de funcionamento do Sistema...! " & vbNewLine & "Para prosseguir," & Chr(13) & "Entre em contato com Novavia Soluções em Informática " & Chr(13) & "Fone : " & "(11) 5548-3890" & Chr(13) & "Você irá receber uma senha para se registrar ao Sistema.", 32, "Aviso de Registro"
    End Select
              
    If Dias_Uso_Sistema > 90 Then
        MsgBox "Desculpe, seu período de utilização terminou...!" & vbCrLf & "Entre em contato com Novavia Soluções em Informática " & Chr(13) & "Fone : " & "(11) 5548-3890", vbInformation, "Aviso de Registro"
        FrmRegistro.CmdAvaliacao.Enabled = False
        FrmRegistro.cmdSair.Visible = True
       '''FrmRegistro.Show 1
        frmMenu.Show
        FrmRegistro.Show 1, frmMenu
    Else
        '''frmAcesso.Show 1
        frmMenu.Show
        frmAcesso.Show 1, frmMenu
    End If
End If
End With

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
Public Function VerificaRegistro(Tabela As String) As Boolean

    On Error GoTo Trata_Erro
    
    sql = ""
    sql = sql & "SELECT count(*) FROM " & "" & Tabela
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open sql, Cnn, 1, 2
    If Rstemp(0) = 5 And Situacao_Registro = False Then
        MsgBox "Desculpe, a Versão Demo só é possível Adicionar até (5) cinco Registros...!" & Chr(13) & Chr(13) & "Registre-se e use o sistema na sua totalidade." & Chr(13) & "Entre em contato com Novavia Soluções em Informática " & Chr(13) & "Fone : " & "(11) 5548-3890" & Chr(13) & "Email : " & "arlindo.jr@ig.com.br", vbInformation, "Aviso"
        VerificaRegistro = False
        Rstemp.Close
        Exit Function
    End If
    
    Rstemp.Close
    VerificaRegistro = True
Trata_Erro:
   
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
