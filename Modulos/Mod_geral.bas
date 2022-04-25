Attribute VB_Name = "Mod_geral"
'****************************************************
'variaveis de conexao
Public Cnn As New ADODB.Connection
Public cmd As New ADODB.Command
'
'*************************
'variaveis para recordsets
Public Rstemp       As New ADODB.Recordset
Public RsTemp1      As New ADODB.Recordset
Public Rstemp2      As New ADODB.Recordset
Public Rstemp3      As New ADODB.Recordset
Public RsTemp4      As New ADODB.Recordset
Public Rstemp5      As New ADODB.Recordset
Public Rstemp6      As New ADODB.Recordset

Public Rs           As New ADODB.Recordset
'
'variaveis pra controle de registro
Global Situacao_Registro As String
Global Dias_Uso_Sistema As Integer
Global ConsultaProd_Ped As Integer
Global flagConsultaPedProd As Boolean
Global gcEmpresa As String
Global gcEndereco As String

Public gTransacao As Boolean
Public sql  As String
Public tmpSQL As String
'
Public gMensagem As String
Public strSql  As String
Public strSql1 As String
Public strSql2 As String
Public strSql3 As String
Public strPesqProdProv As Boolean
Public strFormaPgto As String

Global sysNomeAcesso As String

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
      "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
      lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As _
      String, ByVal nSize As Long, ByVal lpFilename As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
      "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
      lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
'
'*** Fabio Reinert - 10/2017 - Inclusão de variaveis para autocompletar o combobox - Inicio
'
#If Win32 Then
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, _
         ByVal wParam As Long, lParam As Any) As Long
#Else
    Declare Function SendMessage Lib "User" _
        (ByVal hWnd As Integer, ByVal wMsg As Integer, _
         ByVal wParam As Integer, lParam As Any) As Long
#End If
'
'*** Fabio Reinert - 10/2017 - Inclusão de variaveis para autocompletar o combobox - Fim
'

'*************************************************************************************
'*** Fabio Reinert ( Alemao) 06/2017 - Inclusão de captura de IP do cliente - Inicio *
'*************************************************************************************
'
Public STR_IP_COMPUTADOR As String

Public Function BuscaIP() As String
Dim NIC As Variant
Dim NICs As Object

sysNomeAcesso = "MASTER"

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

Sub SelText(object As Control)
    
    With object
        .SelStart = 0
        .SelLength = Len(object)
    End With

End Sub
''
'Public Sub CorFoco(obj As Object)
'    obj.ForeColor = &HFF
'End Sub
''
'Public Sub CorPerdeFoco(obj As Object)
'    obj.ForeColor = &H800000
'End Sub
'
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
'
Function ValidaCGC(CGC As String) As Integer
        Dim Retorno, a, j, I, d1, d2
        If Len(CGC) = 8 And Val(CGC) > 0 Then
           a = 0
           j = 0
           d1 = 0
           For I = 1 To 7
               a = Val(Mid(CGC, I, 1))
               If (I Mod 2) <> 0 Then
                  a = a * 2
               End If
               If a > 9 Then
                  j = j + Int(a / 10) + (a Mod 10)
               Else
                  j = j + a
               End If
           Next I
           d1 = IIf((j Mod 10) <> 0, 10 - (j Mod 10), 0)
           If d1 = Val(Mid(CGC, 8, 1)) Then
              ValidaCGC = True
           Else
              ValidaCGC = False
           End If
        Else
           If Len(CGC) = 14 And Val(CGC) > 0 Then
              a = 0
              I = 0
              d1 = 0
              d2 = 0
              j = 5
              For I = 1 To 12 Step 1
                  a = a + (Val(Mid(CGC, I, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next I
              a = a Mod 11
              d1 = IIf(a > 1, 11 - a, 0)
              a = 0
              I = 0
              j = 6
              For I = 1 To 13 Step 1
                  a = a + (Val(Mid(CGC, I, 1)) * j)
                  j = IIf(j > 2, j - 1, 9)
              Next I
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
'


