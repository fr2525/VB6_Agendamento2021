VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMenu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SisAdven - Sistema de Controle Administrativo"
   ClientHeight    =   7395
   ClientLeft      =   5820
   ClientTop       =   4935
   ClientWidth     =   10935
   Icon            =   "frmMenu.frx":0000
   Picture         =   "frmMenu.frx":08CA
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList iml_Menu 
      Left            =   2400
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":138A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   1429
      ButtonWidth     =   1931
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Re&login"
            Object.ToolTipText     =   "Troca Usuário"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Clientes"
            Object.ToolTipText     =   "Cadastro de Clientes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Fornecedores"
            Object.ToolTipText     =   "Cadastro de Fornecedores"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Produtos"
            Object.ToolTipText     =   "Cadastro de Produtos"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Entradas"
            Object.ToolTipText     =   "Entrada de Produtos"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Pedidos"
            Object.ToolTipText     =   "Emissão de Pedidos/Orçamentos"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "1"
                  Text            =   "Registro"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Object.ToolTipText     =   "Manual de Ajuda"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Agenda"
            Key             =   "g"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Saír"
            Object.ToolTipText     =   "Sair do Sistema"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":14936
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":14C50
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":14F6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":15844
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1611E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":169F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":176D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":193DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":196F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":19B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1FA12
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":20AA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2640
      Top             =   3960
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   3960
   End
   Begin Crystal.CrystalReport Relatorios 
      Left            =   960
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog SelecPrint 
      Left            =   1560
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   7095
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   1323
            MinWidth        =   1323
            Text            =   "Empresa:"
            TextSave        =   "Empresa:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   2206
            MinWidth        =   2206
            Picture         =   "frmMenu.frx":21B7E
            TextSave        =   "23/10/2017"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            Picture         =   "frmMenu.frx":21D58
            TextSave        =   "17:41"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   1940
            MinWidth        =   1940
            Picture         =   "frmMenu.frx":21F32
            Text            =   "USUÁRIO:"
            TextSave        =   "USUÁRIO:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10583
            MinWidth        =   10583
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu menu_relogin 
      Caption         =   "Re&Login"
   End
   Begin VB.Menu menu_caixa 
      Caption         =   "Caixa"
      Visible         =   0   'False
   End
   Begin VB.Menu menu_Cadastro 
      Caption         =   "&Cadastro"
      Enabled         =   0   'False
      Begin VB.Menu menu_cadastro_Cedente 
         Caption         =   "Cedente Boleto"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_cadastro_cliente 
         Caption         =   "&Clientes"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_cadastro_Empresa 
         Caption         =   "&Empresa"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_cadastro_fornecedor 
         Caption         =   "&Fornecedores"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_cadastro_Lojas 
         Caption         =   "&Lojas"
      End
      Begin VB.Menu menu_cadastro_representante 
         Caption         =   "&Vendedores"
         Enabled         =   0   'False
      End
      Begin VB.Menu traco_06 
         Caption         =   "-"
      End
      Begin VB.Menu menu_cadastro_grupo 
         Caption         =   "&Grupos"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_cadastro_marca 
         Caption         =   "&Marcas"
         Enabled         =   0   'False
         Shortcut        =   {F1}
      End
      Begin VB.Menu menu_cadastro_produtos 
         Caption         =   "&Produtos"
         Enabled         =   0   'False
         Shortcut        =   {F2}
      End
      Begin VB.Menu traco_01 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menu_cadastro_transp 
         Caption         =   "&Transportadora"
      End
   End
   Begin VB.Menu menu_Consulta 
      Caption         =   "Co&nsultas"
      Begin VB.Menu menu_consulta_cadastro 
         Caption         =   "&Cadastros"
         Begin VB.Menu menu_consulta_cliente 
            Caption         =   "&Clientes"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu ecfe1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_consulta_CFe 
         Caption         =   "&CFe Fiscais Eletrônicos &Emitidos por Período"
      End
      Begin VB.Menu eepppacfe 
         Caption         =   "-"
      End
      Begin VB.Menu menu_consulta_Nfe 
         Caption         =   "&Nota Fiscal Eletrônica por Nº"
      End
      Begin VB.Menu menu_consulta_Nfe_Emitidas 
         Caption         =   "Notas Fiscais Eletrônicas &Emitidas por Período"
      End
      Begin VB.Menu seeeppa 
         Caption         =   "-"
      End
      Begin VB.Menu menu_ConsultaPedidos 
         Caption         =   "&Pedidos "
         Enabled         =   0   'False
      End
      Begin VB.Menu MENU_CONSULTA_PEDIDOS_EXCLUIDOS 
         Caption         =   "Pedidos &Excluidos"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu menu_movimentacao 
      Caption         =   "&Movimentação"
      Enabled         =   0   'False
      Begin VB.Menu menu_movimentacao_caixa_recebimento 
         Caption         =   "Caixa Recebimento"
         Enabled         =   0   'False
      End
      Begin VB.Menu seeppp 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Movimentacao_Compras 
         Caption         =   "&Compras"
         Shortcut        =   {F3}
      End
      Begin VB.Menu menu_movimentacao_entrada 
         Caption         =   "E&ntradas"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
      Begin VB.Menu traco1500 
         Caption         =   "-"
      End
      Begin VB.Menu menu_movimentacao_nfe_Devolucao 
         Caption         =   "Nfe - Devolução"
         Enabled         =   0   'False
      End
      Begin VB.Menu traco1600 
         Caption         =   "-"
      End
      Begin VB.Menu menu_movimentacao_orcamento 
         Caption         =   "Orçamento"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu traco_14 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menu_movimentacao_saida 
         Caption         =   "&Pedidos"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu menu_movimentacao_saida_Atacado 
         Caption         =   "Pedidos &Atacado"
         Enabled         =   0   'False
         Shortcut        =   {F6}
      End
      Begin VB.Menu epaaaa 
         Caption         =   "-"
      End
      Begin VB.Menu menu_movimentacao_Transferencia_Prod 
         Caption         =   "Transferência de Produtos"
         Shortcut        =   {F7}
      End
      Begin VB.Menu seepaa 
         Caption         =   "-"
      End
      Begin VB.Menu MENU_MOVIMENTACAO_LOG 
         Caption         =   "&Log"
      End
   End
   Begin VB.Menu menu_financeiro 
      Caption         =   "&Financeiro"
      Enabled         =   0   'False
      Begin VB.Menu menu_Financeiro_CadCartao 
         Caption         =   "Cadastro de Cartões"
      End
      Begin VB.Menu menu_Financeiro_Boletos 
         Caption         =   "Boletos"
         Begin VB.Menu menu_Financeiro_Boletos_Retorno 
            Caption         =   "Retorno"
         End
         Begin VB.Menu menu_Financeiro_Boletos_Remessa 
            Caption         =   "Remessa"
         End
      End
      Begin VB.Menu menu_Cheques 
         Caption         =   "&Cheques"
      End
      Begin VB.Menu menu_financeiro_ctapag 
         Caption         =   "Contas a &Pagar"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_financeiro_ctarec 
         Caption         =   "Contas a &Receber"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_financeiro_ctpend 
         Caption         =   "Contas P&endentes"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_financeiro_fluxo 
         Caption         =   "&Fluxo de Caixa"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu menu_Relatorios 
      Caption         =   "&Relatórios"
      Enabled         =   0   'False
      Begin VB.Menu menu_relatorios_cadastros 
         Caption         =   "&Cadastros"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_relatorios_financeiro 
         Caption         =   "&Financeiro"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_relatorios_estoque 
         Caption         =   "&Estoque"
         Enabled         =   0   'False
         Begin VB.Menu menu_relatorios_Contagem_Estoque_sem_Qtde_Estoque 
            Caption         =   "&Contagem Estoque sem Qtde de Estoque"
         End
         Begin VB.Menu menu_relatorios_Contagem_Estoque_com_Qtde_Estoque 
            Caption         =   "Contagem Estoque com Qtde Estoque"
         End
         Begin VB.Menu menu_relatorios_Contagem_estoque_por_Grupo 
            Caption         =   "Contagem Estoque por Grupo"
         End
         Begin VB.Menu EEEPP 
            Caption         =   "-"
         End
         Begin VB.Menu menu_relatorios_estoque_minimo 
            Caption         =   "&Mínimo"
         End
         Begin VB.Menu menu_relatorios_estoque_minimo_compras 
            Caption         =   "&Mínimo somente com estoque > = 0"
         End
         Begin VB.Menu menu_relatorios_estoque_prod_falta 
            Caption         =   "&Produtos em Falta "
         End
         Begin VB.Menu mnu_relatorios_estoque_produtos_fornecedor_grupo 
            Caption         =   "Produtos por &Fornececedor e Grupo"
         End
         Begin VB.Menu ESP 
            Caption         =   "-"
         End
         Begin VB.Menu menu_relatorios_estoque_saldo 
            Caption         =   "Saldo Geral Custo/Venda"
         End
         Begin VB.Menu menu_relatorios_estoque_saldo_Analitico 
            Caption         =   "Saldo Geral Custo/Venda Analítico"
         End
         Begin VB.Menu menu_relatorios_estoque_saldo_por_Grupo 
            Caption         =   "Saldo Geral por Grupo"
         End
      End
      Begin VB.Menu mnu_ETQ_Prood 
         Caption         =   "E&tiquetas Produtos"
      End
      Begin VB.Menu menu_relatorios_comissao 
         Caption         =   "C&omissão"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_relatorios_vendas 
         Caption         =   "&Vendas"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_relatorios_compras 
         Caption         =   "Co&mpras"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_relatorios_mdireta 
         Caption         =   "&Mala Direta Clientes"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu menu_utilitarios 
      Caption         =   "&Utilitários"
      Begin VB.Menu menu_Utilitarios_Backup 
         Caption         =   "&Backup"
         Enabled         =   0   'False
         Shortcut        =   ^B
      End
      Begin VB.Menu menu_utilitarios_config 
         Caption         =   "Con&figurações"
         Enabled         =   0   'False
      End
      Begin VB.Menu MENU_UTILITARIOS_LIBERA_SENHA_MENSAL 
         Caption         =   "Libera Senha Mensal"
      End
      Begin VB.Menu menu_Utilitarios_Restaurar_Banco 
         Caption         =   "&Restaurar Banco de Dados"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu menu_utilitarios_manutarq 
         Caption         =   "&Manutenção Arquivos"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu SEPA_RESTAURAR 
         Caption         =   "-"
      End
      Begin VB.Menu menu_utilitarios_Informativo 
         Caption         =   "&Informativo"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu_utilitarios_Imp_Fiscal 
         Caption         =   "Impressora Fiscal"
         Enabled         =   0   'False
         Begin VB.Menu menu_utilitarios_Imp_Leitura_X 
            Caption         =   "Leitura &X"
         End
         Begin VB.Menu menu_utilitarios_Imp_Reducao_Z 
            Caption         =   "Leitura Redução &Z"
         End
         Begin VB.Menu sepa_l 
            Caption         =   "-"
         End
         Begin VB.Menu menu_utilitarios_Imp_Cancela_Cupom 
            Caption         =   "Cancela Último Cupom Fiscal &Emitido"
         End
         Begin VB.Menu menu_Sangria_Gaveta 
            Caption         =   "&Sangria de Gaveta"
         End
         Begin VB.Menu menu_Suprimento_Gaveta 
            Caption         =   "S&uprimento de Gaveta"
         End
         Begin VB.Menu menu_Cancela_Cupom_Pendente 
            Caption         =   "Cancela Cupom &Pendente"
         End
         Begin VB.Menu SEPA_R 
            Caption         =   "-"
         End
         Begin VB.Menu menu_utilitarios_Imp_Leitura_Memoria_Fiscal_Data 
            Caption         =   "&Leitura Memoria Fiscal por Data"
         End
         Begin VB.Menu menu_utilitarios_Imp_Leitura_Memoria_Fiscal_Reducao 
            Caption         =   "Leitura Memoria Fiscal por &Redução Z"
         End
         Begin VB.Menu menu_Retorno_Aliquotas 
            Caption         =   "Retorno de &Aliquotas"
         End
         Begin VB.Menu menu_Totalizadores_Parciais 
            Caption         =   "&Totalizadores Parciais"
         End
         Begin VB.Menu sseee 
            Caption         =   "-"
         End
         Begin VB.Menu menu_Horario_Verao 
            Caption         =   "Programa Horário de &Verão"
         End
         Begin VB.Menu aaae 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu menu_Cadastro_Formas_Pgto 
            Caption         =   "Cadastrar &Formas Pagamento"
            Visible         =   0   'False
         End
         Begin VB.Menu menu_Inserir_Formas_Pgto 
            Caption         =   "&Inserir Formas Pagamento"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu MENU_UTILITARIOS_REGISTRO 
         Caption         =   "Registr&o"
         Enabled         =   0   'False
      End
      Begin VB.Menu traco_03 
         Caption         =   "-"
      End
      Begin VB.Menu menu_utilitarios_trocasenha 
         Caption         =   "&Troca Senha"
      End
      Begin VB.Menu traco_04 
         Caption         =   "-"
      End
      Begin VB.Menu menu_utilitarios_cteacesso 
         Caption         =   "&Controle de Acesso"
         Enabled         =   0   'False
      End
      Begin VB.Menu traco_05 
         Caption         =   "-"
      End
      Begin VB.Menu menu_utilitarios_calculadora 
         Caption         =   "Ca&lculadora"
      End
      Begin VB.Menu menu_utilitarios_edttexto 
         Caption         =   "&Editor de Texto"
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu menu_Excluir_Dados 
         Caption         =   "Limpar Registros Banco de &Dados"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_agenda 
         Caption         =   "Agenda "
         Enabled         =   0   'False
         Shortcut        =   {F8}
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menu_Sobre 
      Caption         =   "So&bre"
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
   End
   Begin VB.Menu menu_sair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim verifica_abertura As Integer
Dim str As String
Dim X As Integer

Private Sub ExcluirDados()

On Error GoTo trataErroExclusao

    If MsgBox("Esta operação irá apagar toda movimentação de vendas do banco de dados, ficando somente os DADOS do cadastro de Clientes, Fornecedores, Produtos e Estoque." & vbNewLine & vbNewLine & "ESTA OPERAÇÃO É IRREVERSSÍVEL." & vbNewLine & vbNewLine & "Deseja Realmente Efetuar esta Operação ?", vbYesNo + vbDefaultButton2 + vbOKOnly + vbExclamation, "Atenção") = vbNo Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
        
    gTransacao = True
    Cnn.BeginTrans
      
'    SQL = "update Estoque set "
'    SQL = SQL & "SALDO_EM_ESTOQUE = '0" & "'"
'    CNN.Execute SQL

    sql = "DELETE  FROM CTAS_PENDENTE "
    Cnn.Execute sql
    
'    sql = "DELETE  FROM GERA_COD_PED "
'    Cnn.Execute sql
'
'    sql = "DELETE  FROM GERA_COD_ORC "
'    Cnn.Execute sql
    
    sql = "DELETE  FROM ITENS_ORCAMENTO "
    Cnn.Execute sql

    sql = "DELETE  FROM orcamentos "
    Cnn.Execute sql
    
    sql = "DELETE  FROM itens_saida "
    Cnn.Execute sql

    sql = "DELETE  FROM saidas_produto "
    Cnn.Execute sql
    
    sql = "DELETE  FROM entrada_Produto "
    Cnn.Execute sql
    
    sql = "DELETE  FROM itens_entrada "
    Cnn.Execute sql
    
    sql = "DELETE  FROM COMPRA_Produto "
    Cnn.Execute sql
    
    sql = "DELETE  FROM ITENS_COMPRA"
    Cnn.Execute sql
    
    sql = "DELETE  FROM DEVOLUCAO_NFE"
    Cnn.Execute sql
    
    sql = "DELETE  FROM ITENS_DEVOLUCAO_NFE"
    Cnn.Execute sql
    
    sql = "DELETE  FROM ETIQ_GONDOLA "
    Cnn.Execute sql
    
    sql = "DELETE  FROM rece_Paga "
    Cnn.Execute sql
    
    sql = "DELETE  FROM INVENTARIO "
    Cnn.Execute sql
    
    sql = "DELETE  FROM INVENTARIO_GRUPO "
    Cnn.Execute sql
    
    sql = "DELETE  FROM ITENS_TRANSF_PROD"
    Cnn.Execute sql
    
    sql = "DELETE  FROM TRANSF_PROD"
    Cnn.Execute sql
    
    sql = "DELETE  FROM LOG"
    Cnn.Execute sql
    
    sql = "DELETE  FROM forma_Pgto "
    Cnn.Execute sql
    
    sql = "DELETE  FROM cheques "
    Cnn.Execute sql
    
    sql = "DELETE  FROM rel_comissao "
    Cnn.Execute sql
    
    sql = "DELETE  FROM rel_estoque_minimo "
    Cnn.Execute sql

    sql = "DELETE  FROM rel_etq_prod "
    Cnn.Execute sql
    
    sql = "DELETE  FROM REL_FATURAMENTO "
    Cnn.Execute sql
    
    sql = "DELETE  FROM relPedidos "
    Cnn.Execute sql
    
    sql = "DELETE  FROM rel_financ "
    Cnn.Execute sql
    
    sql = "DELETE  FROM rel_cheques "
    Cnn.Execute sql
            
    sql = "DELETE  FROM REL_VDACPA "
    Cnn.Execute sql
    
    sql = "DELETE  FROM rel_lucro_prod "
    Cnn.Execute sql
    
    sql = "DELETE  FROM boletos_PG "
    Cnn.Execute sql
    
    sql = "DELETE  FROM XIBIU "
    Cnn.Execute sql
    
    sql = "DELETE FROM NFE "
    Cnn.Execute sql
    
    sql = "DELETE FROM LOTE_NFE "
    Cnn.Execute sql
    
    sql = "DELETE FROM CFE "
    Cnn.Execute sql
    
    sql = "DELETE FROM  CAD_USUARIOS WHERE USUA_CD_USUARIO > 1 "
    Cnn.Execute sql
    
    sql = "DELETE FROM  CAD_OPCOES_USUARIO_ACESSO "
    Cnn.Execute sql
    
    sql = "DELETE FROM ARQ_ESTOQUE "
    Cnn.Execute sql
'
'**** Alemão - 09/2017 - Acrescimo das tabelas novas do contas a receber  - Inicio
    sql = "DELETE FROM TAB_RECCARTOES "
    Cnn.Execute sql
    
    sql = "DELETE FROM TAB_RECAVISTA "
    Cnn.Execute sql
'**** Alemão - 09/2017 - Acrescimo das tabelas novas do contas a receber  - Fim
'
    'ZERAR GENERATOR
    Cnn.Execute "SET GENERATOR GEN_SEQ_PED TO 0"
    Cnn.Execute "SET GENERATOR GEN_SEQ_ORC TO 0"
    Cnn.Execute "SET GENERATOR GEN_CFE_ID1 TO 0"
    Cnn.Execute "SET GENERATOR G$_SEQUENCIA TO 0"
    Cnn.Execute "SET GENERATOR GEN_SEQ_PED TO 0"
    Cnn.Execute "SET GENERATOR GEN_BLPG_ID TO 0"
    Cnn.Execute "SET GENERATOR GEN_BL_ID1 TO 0"
    Cnn.Execute "SET GENERATOR GEN_INVENTARIO_GR_ID TO 0"
    Cnn.Execute "SET GENERATOR GEN_INVENTARIO_ID TO 0"
    Cnn.Execute "SET GENERATOR GEN_LOT_ID1 TO 0"
    Cnn.Execute "SET GENERATOR GEN_NFE_ID1 TO 0"
    Cnn.Execute "SET GENERATOR GEN_RECEPG TO 0"
    Cnn.Execute "SET GENERATOR GEN_SEQ_DEVOLUCAO_NFE TO 0"
    Cnn.Execute "SET GENERATOR GEN_SEQ_ITENS_DEVOLUCAO_NFE TO 0"
    Cnn.Execute "SET GENERATOR GEN_SEQ_LOG TO 0"
    Cnn.Execute "SET GENERATOR GEN_SEQ_ORC TO 0"
    Cnn.Execute "SET GENERATOR GEN_SEQ_ARQ_ESTOQUE TO 0"
    'Cnn.Execute "SET GENERATOR GEN_SEQ_CAD_DESPESAS TO 0"
    
    'zerar todos GENERATOR
    'Cnn.Execute " delete from rdb$database   where exists(select * from RDB$Relations)"
    
    Cnn.CommitTrans
    DoEvents
    
    gTransacao = False
            
    Screen.MousePointer = 1
    MsgBox "Operação Realizada com Sucesso.", vbInformation, "Aviso"
        
 Exit Sub
 
trataErroExclusao:
 If gTransacao = True Then Cnn.RollbackTrans
 
 If Err.Number <> 0 Then
    MsgBox "Ocorreu um Erro Nesta Operação.." & "Nº Erro  " & Err.Number & vbNewLine & "Descrição:  " & Err.Description, vbCritical, "Aviso"
    Err.Clear
    Screen.MousePointer = 1
 End If
 
        
End Sub


Private Sub MDIForm_Activate()


'    sql = "create or alter procedure SP_RETORNA_ALIQ_INTER ( "
'    sql = sql & "VUF_ORIGEM char(2),VUF_DESTINO char(2),VORIG char(1))"
'    sql = sql & "returns (NALIQ numeric(15,4))"
'    sql = sql & "AS BEGIN  IF (:vorig IN (1,2,3,8)) THEN naliq = '4.00';"
'    sql = sql & " ELSE  BEGIN  IF (:vuf_origem IN ('RS','SC','PR','SP','MG','RJ') AND"
'    sql = sql & ":vuf_destino IN "
'    sql = sql & "('ES','AC','AM','RO','RR','PA','AP','TO','MA','PI','CE','BA','SE','AL','PE','PB','RN','GO','MT','MS','DF')) "
'    sql = sql & " THEN naliq = '7.00';"
'    sql = sql & " ELSE naliq = '12.00';"
'    sql = sql & " END SUSPEND; END "
'    Cnn.Execute sql
'    Cnn.CommitTrans



'        Call ClearCommandParameters

'        Set cmd = Nothing
'        Set cmd = New ADODB.Command
'        cmd.ActiveConnection = Cnn
'        cmd.CommandType = adCmdStoredProc
'        cmd.CommandText = "SP_RETORNA_ALIQ_INTER"
'
'        UF_ORIGEM = "SP"
'        UF_DESTINO = "MG"
'
'        'IN-parameters
'        cmd.Parameters.Append cmd.CreateParameter("VUF_ORIGEM", adVarChar, adParamInput, 2, UF_ORIGEM)
'        cmd.Parameters.Append cmd.CreateParameter("VUF_DESTINO", adVarChar, adParamInput, 2, UF_DESTINO)
'        cmd.Parameters.Append cmd.CreateParameter("VORIG", adVarChar, adParamInput, 1, "4")
'
'        'OUT -Parameters
'        cmd.Parameters.Append cmd.CreateParameter("NALIQ", adBSTR, adParamReturnValue) 'RETORNA_PARAMETRO DO CAMPO
'        cmd.Execute
'        'RETORNA OS PARAMETROS
'        str_VUF_ORIGEM = cmd.Parameters("VUF_ORIGEM") 'RETORNA PARAMETRO  - adParamOutput
'        str_VUF_DESTINO = cmd.Parameters("VUF_DESTINO") 'RETORNA PARAMETRO  - adParamOutput
'        srt_NALIQ = cmd.Parameters("NALIQ") 'RETORNA PARAMETRO  - adParamOutput
'        'ou
'        Aliquota = cmd!NALIQ
'        'ou outra forma de retorno pelo index
'        strOutputParam0 = cmd.Parameters(0).Value
'        strOutputParam0 = cmd.Parameters(1).Value
'        'https://www.codeproject.com/Articles/15222/How-to-Use-Stored-Procedures-in-VB
        
'        Set cmd.ActiveConnection = cn
'        cmd.CommandType = adCmdStoredProc
'        cmd.CommandText = "TESTEVCT"
'        cmd.Parameters.Append cmd.CreateParameter("VAL", adVarChar, adParamInput, 1, "2")
'        cmd.Parameters.Append cmd.CreateParameter("res1", adInteger, adParamReturnValue)
'        cmd.Parameters.Append cmd.CreateParameter("res2", adInteger, adParamReturnValue)
'        cmd.Execute
'        Debug.Print cmd("res1"), cmd("res2")
        
        

'        'IN-parameters
'        cmd.Parameters.Append cmd.CreateParameter("STR_CODIGO", adDouble, adParamInput, , CodCliente)
'        cmd.Parameters.Append cmd.CreateParameter("STR_DATA", adBSTR, adParamInput, , Format(Date, "MM/DD/yyyy"))
'        cmd.Parameters.Append cmd.CreateParameter("STR_DESCRICAO", adBSTR, adParamInput, , "Recebimento Entrada nº " & Format(txtCodSeq.Text, "0000"))
'        cmd.Parameters.Append cmd.CreateParameter("STR_VALOR", adBSTR, adParamInput, , Troca_Virg_Zero(Format(lblTotPedido.Caption, "0.00")))
'        cmd.Parameters.Append cmd.CreateParameter("STR_DATA_BAIXA", adBSTR, adParamInput, , Format(Date, "MM/DD/yyyy"))
'        cmd.Parameters.Append cmd.CreateParameter("STR_VALOR_BAIXA", adBSTR, adParamInput, , Troca_Virg_Zero(Format(lblTotPedido.Caption, "0.00")))
'        cmd.Parameters.Append cmd.CreateParameter("STR_TIPO_MOVIMENTACAO", adBSTR, adParamInput, , "R")
'        cmd.Parameters.Append cmd.CreateParameter("STR_TP_FAVORECIDO", adBSTR, adParamInput, , "C")
'        cmd.Parameters.Append cmd.CreateParameter("STR_CODIGO", adDouble, adParamInput, , CCur(NovoCodigo))
'        cmd.CommandText = "SP_INSERT_RECE_PAGA"
'        cmd.Execute

'        sql = "CREATE PROCEDURE EXEC_SQL2 (WTIPO INTEGER,WTABELA VARCHAR(15),WDDL VARCHAR(100), WWHERE VARCHAR(100))"
'        sql = sql + " AS BEGIN     "
'        sql = sql & " IF (WTIPO = 0) THEN  EXECUTE STATEMENT 'INSERT INTO ' || WTABELA ||' '|| WDDL||' '||WWHERE; "
'        sql = sql & " IF (WTIPO = 1) THEN  EXECUTE STATEMENT 'DELETE FROM ' || WTABELA ||' '|| WWHERE;"
'        sql = sql & " IF (WTIPO = 2) THEN  EXECUTE STATEMENT 'UPDATE ' || WTABELA || ' SET ' || WDDL ||' '|| WWHERE;"
'        sql = sql & " EXIT; END "
'        Cnn.Execute sql
'
'        'Call Conecta_Banco
'
'        'ClearCommandParameters
'        Set cmd = Nothing
'        cmd.ActiveConnection = Cnn
'        cmd.CommandType = adCmdStoredProc
'
'        sql = "Create PROCEDURE SP_UPDATE_FORMA_PGTO1000 (NRO_PEDIDO DOUBLE PRECISION, FORMA_PGTO VARCHAR(5), STR_STATUS_SAIDA VARCHAR(1))"
'        sql = sql + " AS BEGIN "
'        sql = sql + " UPDATE SAIDAS_PRODUTO SET FORMAPGTO=:FORMA_PGTO, STATUS_SAIDA =:STR_STATUS_SAIDA WHERE SEQUENCIA=:NRO_PEDIDO; END"
'        Cnn.Execute sql
'
'        Set cmd = Nothing
'        cmd.ActiveConnection = Cnn
'        cmd.CommandType = adCmdStoredProc
'
'        cmd.Parameters.Append cmd.CreateParameter("WTIPO", adInteger, adParamInput, , 0) 'INSERT
'        cmd.Parameters.Append cmd.CreateParameter("WTABELA", adBSTR, adParamInput, , "MARCAS")
'        cmd.Parameters.Append cmd.CreateParameter("WDDL", adBSTR, adParamInput, , "(CODIGO,DESCRICAO)")
'        cmd.Parameters.Append cmd.CreateParameter("WWHERE", adBSTR, adParamInput, , "VALUES(100,'MARCASQL')")
'        cmd.CommandText = "EXEC_SQL"
'        cmd.Execute
'
'        Set cmd = Nothing
'        cmd.ActiveConnection = Cnn
'
'        cmd.CommandType = adCmdStoredProc
'        cmd.Parameters.Append cmd.CreateParameter("WTIPO", adInteger, adParamInput, , 1) 'DELETE
'        cmd.Parameters.Append cmd.CreateParameter("WTABELA", adBSTR, adParamInput, , "GRUPO")
'        cmd.Parameters.Append cmd.CreateParameter("WWHERE", adBSTR, adParamInput, , "CODIGO = 9")
'        cmd.CommandText = "EXEC_SQL"
'        cmd.Execute
'
'        Set cmd = Nothing
'        cmd.ActiveConnection = Cnn
'        cmd.CommandType = adCmdStoredProc
'
'        cmd.Parameters.Append cmd.CreateParameter("WTIPO", adInteger, adParamInput, , 2) 'update
'        cmd.Parameters.Append cmd.CreateParameter("WTABELA", adBSTR, adParamInput, , "MARCAS")
'        cmd.Parameters.Append cmd.CreateParameter("WDDL", adBSTR, adParamInput, , "DESCRICAO='ARLINDO'")
'        cmd.Parameters.Append cmd.CreateParameter("WWHERE", adBSTR, adParamInput, , "WHERE CODIGO='1'")
'        cmd.CommandText = "EXEC_SQL"
'
'        cmd.Execute
'
'        'Cnn.CommitTrans
   
    Call AddIconToMenu

End Sub


Private Sub MENU_UTILITARIOS_LIBERA_SENHA_MENSAL_Click()
Call sLibera
End Sub

Private Sub menu_cadastro_Cedente_Click()
frm_Conta_Corrente_Cedente_Boleto.Show 1
End Sub


Private Sub menu_cadastro_Lojas_Click()
Frm_Cad_Lojas.Show 1
End Sub

Private Sub menu_consulta_CFe_Click()
Me.Timer1.Enabled = False
Me.Timer2.Enabled = False
Frm_CFe_Emitidas.Show 1
End Sub

Private Sub menu_consulta_Nfe_Click()
With Frm_NFe
    .FLAG_CONSULTA_NFE = True
    .cmd_Gravar.Enabled = False
    .Show 1
End With
End Sub

Private Sub menu_consulta_Nfe_Emitidas_Click()
Frm_NFe_Emitidas.Show 1
End Sub


Private Sub menu_Financeiro_Boletos_Remessa_Click()
frm_Arquivos_Remessa.Show 1
End Sub

Private Sub menu_Financeiro_Boletos_Retorno_Click()
Set frm_Arquivos_Retorno = Nothing
'frm_Arquivos_Retorno.Show 1
If ListaArquivosRetorno = True Then
    Set frm_Arquivos_Retorno = Nothing
    With frm_Arquivos_Retorno
        .ProcessaArquivos
    End With
Else
    MsgBox "Não foi encontrado nenhum Arquivo de retorno para Processamento na Pasta Retorno ...!" & vbNewLine & vbNewLine & _
    "Os Arquivos de retorno devem estar no SERVIDOR na pasta C:\Sistema SisAdven\Retorno", vbInformation, "Aviso"
End If

End Sub

Private Sub menu_movimentacao_emissao_nota_Fiscal_Click()
Frm_NFe.Show 1
End Sub

Private Sub menu_movimentacao_caixa_recebimento_Click()
frmCdPgto.Show 1
End Sub

Private Sub MENU_MOVIMENTACAO_LOG_Click()
If sysAcesso <> 1 Then
    MsgBox "Você não tem permissão de acesso a esse módulo...!", vbInformation, "Aviso"
    Exit Sub
End If

FrmLog.Show 1
End Sub

Private Sub menu_movimentacao_nfe_Devolucao_Click()
frmMenu.Timer1.Enabled = False
frmMenu.Timer2.Enabled = False
FRM_Devolucao.Show 1
End Sub

Private Sub menu_movimentacao_saida_Atacado_Click()
    'Chamar rotina de checagem segurança
    If fValida_No_Pedido() Then
        frmSaidas_Atacado.Show 1
    Else
        Call Fecha_Formularios
        Call MDIForm.Main
    End If
End Sub


Private Sub menu_movimentacao_Transferencia_Prod_Click()
    Frm_Transf_Produtos.Show 1
End Sub

Private Sub menu_relatorios_Contagem_Estoque_com_Qtde_Estoque_Click()
    With Relatorios
        .Reset
        .WindowShowZoomCtl = True
        '.WindowControlBox = False
        .PageZoom (100)
        .WindowShowExportBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowCloseBtn = True
        .WindowShowGroupTree = False
        .WindowState = crptMaximized
        .PageZoom (100)
        .WindowTitle = Me.Caption
        '.SelectionFormula = "{INVENTARIO.DATA} >= Date(" & dataDe & ") and {INVENTARIO.DATA} <= Date(" & dataate & ")"
        '.SelectionFormula = "{representante.codigo} = " & cboInicial.ItemData(cboInicial.ListIndex)
        .ReportFileName = App.Path & "\REL_CONT_ESTOQUE_COM_QTDE.rpt"
        '.Connect = "DSN=cnn_firebird;UID=SYSDBA;PWD=masterkey"
        '.SelectionFormula = "{INVENTARIO.estoque} < 0 "
        .WindowTitle = frmRelVenda.Caption
        '.Formulas(0) = "PERI_DE = '" & Me.TXT_DATA.Text & "'"
        '.Formulas(1) = "PERI_ATE = '" & TXT_DATA.Text & "'"
        .Formulas(2) = "subTITULO = 'Relatório Inventário Contagem de Estoque'"
    
        .RetrieveDataFiles
        .Action = 1
    End With
End Sub

Private Sub menu_relatorios_Contagem_estoque_por_Grupo_Click()
frm_Lista_Prod_por_Grupo.Show 1
End Sub

Private Sub menu_relatorios_Contagem_Estoque_sem_Qtde_Estoque_Click()
    With Relatorios
        .Reset
        .WindowShowZoomCtl = True
        '.WindowControlBox = False
        .PageZoom (100)
        .WindowShowExportBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowCloseBtn = True
        .WindowShowGroupTree = False
        .WindowState = crptMaximized
        .PageZoom (100)
        .WindowTitle = Me.Caption
        '.SelectionFormula = "{INVENTARIO.DATA} >= Date(" & dataDe & ") and {INVENTARIO.DATA} <= Date(" & dataate & ")"
        '.SelectionFormula = "{representante.codigo} = " & cboInicial.ItemData(cboInicial.ListIndex)
        .ReportFileName = App.Path & "\REL_CONT_ESTOQUE.rpt"
        '.Connect = "DSN=cnn_firebird;UID=SYSDBA;PWD=masterkey"
        '.SelectionFormula = "{INVENTARIO.estoque} < 0 "
        .WindowTitle = frmRelVenda.Caption
        '.Formulas(0) = "PERI_DE = '" & Me.TXT_DATA.Text & "'"
        '.Formulas(1) = "PERI_ATE = '" & TXT_DATA.Text & "'"
        .Formulas(2) = "SUBTITULO = 'Relatório Inventário Contagem de Estoque'"
    
        .RetrieveDataFiles
        .Action = 1
    End With
End Sub

Private Sub menu_relatorios_estoque_minimo_compras_Click()
    frmMenu.MousePointer = 11

    
    On Error GoTo SaiImp
   '' SelecPrint.Action = 5
    
    Relatorios.Reset
    Relatorios.Destination = crptToWindow
    Relatorios.WindowState = crptMaximized
    Relatorios.ReportFileName = App.Path & "\Rel_so_estoque_neg.rpt"
    Relatorios.WindowTitle = "Relatório de Compras Estoque Mínimo"
    
    Relatorios.Action = 1
    Relatorios.PageZoom (100)
    frmMenu.MousePointer = 0
    
    Exit Sub

SaiImp:
    If Err.Number = 32755 Then
        Err.Clear
        Screen.MousePointer = 1
        Exit Sub
    ElseIf Err.Number = 20525 Then
        Err.Clear
       ' frmConfiguraBase.Show 1
        Screen.MousePointer = 1
        Exit Sub
    Else
        MsgBox "Ocorreu um erro: " & Err.Description & "Nro: " & Err.Number
    End If
End Sub

Private Sub menu_relatorios_estoque_prod_falta_Click()

    On Error GoTo SaiImp
    'SelecPrint.Action = 5
    
    Relatorios.Reset
    Relatorios.Destination = crptToWindow
    Relatorios.WindowState = crptMaximized
    'Relatorios.ReportFileName = "c:\Sistema SisAdven\Rel_estoque_neg.rpt"
    'Relatorios.ReportFileName = "c:\Sistema SisAdven\Rel_estoque_neg.rpt"
    Relatorios.ReportFileName = App.Path & "\Rel_estoque_neg.rpt"
    Relatorios.WindowTitle = "Relatório de Estoque Mínimo"
    
    Relatorios.Action = 1
    Relatorios.PageZoom (100)

Exit Sub

SaiImp:

    If Err.Number = 32755 Then
        Err.Clear
        Screen.MousePointer = 1
        Exit Sub
    End If


End Sub
Private Sub Verifica_Status_ImpFiscal()
 

    LocalRetorno = LeParametrosIni("Sistema", "Retorno")
    If LocalRetorno = "-2" Then
        LocalRetorno = "0" 'devolve o retorno na variavel
    Else
        LocalRetorno = Left(LocalRetorno, 1)
    End If
    
    Retorno = Bematech_FI_AbrePortaSerial()
    
    'gravar arquivo ini RETORNO AbrePortaSerial
    'WriteIniFile App.Path & "\SisAdven.ini", "Ret_AbrePortaSerial", "", Retorno
    WriteIniFile "C:\SisAdven.ini", "Ret_AbrePortaSerial", "", Retorno
    
   ' Call VerificaRetornoImpressora("", "", "BemaFI32")
    If Retorno = -4 Or Retorno = -5 Then
      '  frmMenu.menu_utilitarios_Imp_Fiscal.Enabled = False
    End If
End Sub
Public Sub CarregaNomeEmpresa()
    sql = "select * from empresa "
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Cnn, 1, 2
    
    If Rs.RecordCount = 0 Then Exit Sub
    Rs.MoveLast
    Rs.MoveFirst
    While Not Rs.EOF
        NomeEmpresa = IIf(IsNull(Rs("RazaoSocial_Empresa")), "", Rs("RazaoSocial_Empresa"))
        EnderecoEmpresa = IIf(IsNull(Rs("Endereco_Empresa")), "", Rs("Endereco_Empresa"))
        CGC_EMPRESA = IIf(IsNull(Rs("Cgc_Cpf")), "", Rs("Cgc_Cpf"))
        CEP_EMPRESA = IIf(IsNull(Rs("Cep_Empresa")), "", Rs("Cep_Empresa"))
        Fone1Empresa = IIf(IsNull(Rs("fone1_Empresa")), "", Rs("fone1_Empresa"))
        Fone2Empresa = IIf(IsNull(Rs("fone2_Empresa")), "", Rs("fone2_Empresa"))
        emailEmpresa = IIf(IsNull(Rs("E_MAIL_EMPRESA")), "", Rs("E_MAIL_EMPRESA"))
        Rs.MoveNext
    Wend
    status.Panels(2).Text = Trim(UCase(NomeEmpresa))
    
    Rs.Close
    Set Rs = Nothing
End Sub

Private Sub MDIForm_Load()

    Call RemoveMenus(Me)

    NomeSistema = "SisAdven"
   ' verifica_abertura = App.PrevInstance
    If App.PrevInstance = True Then
        Dim Form As Form
        For Each Form In Forms
           MsgBox "O Sistema já se Encontra Aberto...", vbInformation, "Aviso"
           Unload Form
           Set Form = Nothing
        Next Form
        End
    End If
    
    If Val(Acesso_OK) = 0 Then
        frmMenu.Enabled = False
      ' frmAcesso.Show 1
    End If
    
    Call Main
    
    STR_IP_COMPUTADOR = BuscaIP()
    
    '*** Fabio Reinert - 04/2017 - Nova Checagem de segurança - Inicio
    Call sValidaCliente
    '*** Fabio Reinert - 04/2017 - Nova Checagem de segurança - Fim
    
    Call CarregaNomeEmpresa
    
  '  Retorno = Bematech_FI_AbrePortaSerial()
    
    status.Panels(6).Text = UCase(NomeUsuario)
    NomeComputador = Environ("ComputerName")
    
'*** Fabio Reinert - 07/2017 0 inclusão do botão da agenda caso a empresa seja Petshop - Inicio
    Dim sTipoEmpresa As String
  
    sTipoEmpresa = ReadIniFile(App.Path & "\SisAdven.ini", "TIPO_EMPRESA", "", "")
    If sTipoEmpresa = "PETSHOP" Then
        Toolbar1.Buttons(12).Visible = True
        Toolbar1.Buttons(12).Enabled = True
        Toolbar1.Buttons(13).Visible = True
        Toolbar1.Buttons(13).Enabled = True
        mnu_agenda.Visible = True
        mnu_agenda.Enabled = True
    End If

'*** Fabio Reinert - 07/2017 0 inclusão do botão da agenda caso a empresa seja Petshop - Fim
    '
    '*** Fabio Reinert - 09/2017 Chamada a sub para atualização de tabelas e campos de tabelas - Inicio
    '
    Call sVerificaAtualizacoes
    '
    '*** Fabio Reinert - 09/2017 Chamada a sub para atualização de tabelas e campos de tabelas - Fim
    '*
    If sysAcesso = 1 Then
        DoEvents
        'FRM_TASK.Show
        '**** Fabio Reinert - 10/2017 - Colocada leitura da linha de comando do projeto para efeito de testes - Inicio
        If Len(Command$) = 0 Then
            With FRM_TASK
                .CarregaInformativo
                .Show 1
            End With
        Else
            Dim sArgs() As String
            Dim sPrimeiro, sSegundo As String
            sArgs = Split(Command$, ",")
            sPrimeiro = sArgs(0)
            
            
        End If
        'With FRM_TASK
        '    .CarregaInformativo
        '    .Show 1
        'End With
        '****
        '**** Fabio Reinert - 10/2017 - Colocada leitura da linha de comando do projeto para efeito de testes - Fim
        If UCase(NomeComputador) = "SERVIDOR" Then
            sql = "SELECT * FROM CAD_MENUS WHERE MENU_CD_CODI = 72 "
            Set Rstemp6 = New ADODB.Recordset
            Rstemp6.Open sql, Cnn, 1, 2
            If Rstemp6.RecordCount = 0 Then
                sql = "INSERT INTO CAD_MENUS VALUES("
                sql = sql & "'SisAdven',"
                sql = sql & "72,"
                Menu = UCase("txtQtdeEstoque")
                sql = sql & "'" & Menu & "',"
                sql = sql & "'Cadastro - Produtos - Alterar Qtde Estoque')"
                frmMenu.menu_cadastro_Cedente.Enabled = True
                Cnn.Execute sql
            End If
            Rstemp6.Close
            Set Rstemp6 = Nothing
            
            sql = "SELECT count(*) FROM CAD_MENUS WHERE MENU_CD_CODI = 135 "
            Set Rstemp6 = New ADODB.Recordset
            Rstemp6.Open sql, Cnn, 1, 2
            If Rstemp6(0) = 0 Then
                sql = "INSERT INTO CAD_MENUS VALUES("
                sql = sql & "'SisAdven',"
                sql = sql & "135,"
                sql = sql & "'MENU_MOVIMENTACAO_CAIXA_RECEBIMENTO',"
                sql = sql & "'Movimentação - Caixa Recebimento')"
                Cnn.Execute sql
                Cnn.CommitTrans
                frmMenu.menu_movimentacao_caixa_recebimento.Enabled = True
            End If
            'processa arquivos de retorno
            If ListaArquivosRetorno = True Then
                With frm_Arquivos_Retorno
                    .ProcessaArquivos
                End With
            End If
        End If
    End If
    
    Call DoAboutTxt
    
'    If Resolucao(800, 600) = False Then
'       MsgBox "O Programa fica melhor Visível na Resolução de Vídeo," & vbNewLine & "(800 x 600) High color (16 bits)", 64, "Aviso"
'    End If
    
    If ReadIniFile(App.Path & "\SisAdven.ini", "Cursor", "Chk", "0") Then
        flagCursorCodigo = True
    Else
        flagCursorCodigo = False
    End If
    
    If ReadIniFile(App.Path & "\SisAdven.ini", "Qtde", "Chk", "0") Then
        flagQtde1 = True
    Else
        flagQtde1 = False
    End If
    
    Mensagem_Final_Cupom = ReadIniFile(App.Path & "\SisAdven.ini", "Men_Promoc", "", "")
    Mensagem_Final_Cupom = TiraAcento(UCase(Mid(Mensagem_Final_Cupom, 1, 492)))
    
    retImpFiscal = ReadIniFile("C:\SisAdven.ini", "Imp_Fisc_Sel", "Uso", "0")
    If retImpFiscal = 1 Then    'BEMATECH
        Call Verifica_Status_ImpFiscal
        flagImpFiscalSelecionada = True
        'cmbImpFiscal.Text = "Bematech - Mp20FI-II"
        retImpFiscal = 1
         frmMenu.menu_utilitarios_Imp_Fiscal.Enabled = True
    ElseIf retImpFiscal = 5 Then
        flagImpFiscalSelecionada = True 'SAT
        'frmMenu.menu_utilitarios_Imp_Fiscal.Enabled = True
       ' Me.menu_utilitarios_Imp_Leitura_Memoria_Fiscal_Data.Enabled = False
       ' Me.menu_utilitarios_Imp_Leitura_Memoria_Fiscal_Reducao.Enabled = False
        Me.menu_Suprimento_Gaveta.Enabled = False
        Me.menu_Sangria_Gaveta.Enabled = False
        Me.menu_Totalizadores_Parciais.Enabled = False
       ' menu_utilitarios_Imp_Cancela_Cupom.Enabled = False
        menu_Sangria_Gaveta.Enabled = False
        menu_Suprimento_Gaveta.Enabled = False
        menu_utilitarios_Imp_Leitura_Memoria_Fiscal_Data.Enabled = False
        menu_utilitarios_Imp_Leitura_Memoria_Fiscal_Reducao.Enabled = False
        menu_Totalizadores_Parciais.Enabled = False
        menu_Horario_Verao.Enabled = True
        menu_Cadastro_Formas_Pgto.Enabled = False
        frmMenu.menu_utilitarios_Imp_Fiscal.Enabled = False
    Else
        retImpFiscal = 0
        flagImpFiscalSelecionada = False
        'cmbImpFiscal.Text = "Nenhuma"
    End If
    
    If ReadIniFile(App.Path & "\SisAdven.ini", "EstoqueNeg", "Chk", "0") Then
        flagComEstoque = True
    Else
        flagComEstoque = False
    End If
    
    If ReadIniFile(App.Path & "\SisAdven.ini", "AltVenda", "Chk", "0") Then
        flagAltPrVenda = True
    Else
        flagAltPrVenda = False
    End If

    If ReadIniFile("c:\SisAdven.ini", "Desconto", "Chk", "0") = 0 Then
        flagDescPedOrc = True
    Else
        flagDescPedOrc = False
    End If
    
    If Situacao_Registro = True Then
        MENU_UTILITARIOS_REGISTRO.Visible = False
    End If
    
    If ReadIniFile("c:\SisAdven.ini", "Gaveta_Bematec", "Chk", "0") = 1 Then
        flag_Gaveta_Bematec = True
        flag_Gaveta_Elgin = False
    ElseIf ReadIniFile("c:\SisAdven.ini", "Gaveta_Elgin", "Chk", "0") = 1 Then
        flag_Gaveta_Elgin = True
        flag_Gaveta_Bematec = False
    Else
        flag_Gaveta_Bematec = False
        flag_Gaveta_Elgin = False
    End If
    
    If ReadIniFile(App.Path & "\SisAdven.ini", "Orcamentos", "Chk", "1") Then
        flagEmitir_Orcamentos = True
    Else
        flagEmitir_Orcamentos = False
    End If
    
'''    Call VerificaAtualizacoes

    NomeComputador = Environ("ComputerName")
    If UCase(NomeComputador) = "SERVIDOR" Then
        Call VerificaAtualizacoes_CST
    End If
'
'    'TRATA TABELA SAIDAS_PRODUTO CAMPO PDV
'    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'SAIDAS_PRODUTO' AND rdb$field_name = 'PDV' "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) = 0 Then
'        sql = "ALTER TABLE SAIDAS_PRODUTO ADD CAIXA varchar(30), ADD PDV varchar(30)"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        ' Indices
'        sql = "CREATE INDEX ID_X_SEQUENCIA ON ITENS_SAIDA (SEQUENCIA)"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'    End If
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
''
'    'TRATA TABELA SAIDAS_PRODUTO CAMPO PDV
'    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'COD_BAR'"
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) = 0 Then
'        sql = "CREATE TABLE COD_BAR (COD_BAR varchar(10))"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        sql = "UPDATE COD_BAR SET COD_BAR = '987654'"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'    End If
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    'TRATA TABELA ETIQ_GONDOLA
'    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'ETIQ_GONDOLA' "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) = 0 Then
'        sql = "CREATE TABLE ETIQ_GONDOLA (CODIGO DOUBLE PRECISION,PRIMARY KEY (CODIGO))"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'    End If
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    'CFE - SAT
'    'sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'CFE' AND rdb$field_name = 'NRO_CAIXA' "
'    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'CFE'"
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) = 0 Then
'        sql = "CREATE TABLE CFE ("
'        sql = sql & " ID  INTEGER NOT NULL,"
'        sql = sql & " NRO_PEDIDO_CFE        BLOB SUB_TYPE 1 SEGMENT SIZE 3000 CHARACTER SET WIN1252 COLLATE WIN1252,"
'        sql = sql & " EMISSAO_CFE           DATE,"
'        sql = sql & " NRO_CFE               INTEGER,"
'        sql = sql & " SESSAO_CFE            VARCHAR(10),"
'        sql = sql & " CHAVE_ACESSO_CFE      VARCHAR(200),"
'        sql = sql & " STATUS_RETORNO_CFE    VARCHAR(200),"
'        sql = sql & " XML_CFE               BLOB SUB_TYPE 1 SEGMENT SIZE 16000 CHARACTER SET WIN1252,"
'        sql = sql & " CAMINHO_XML_CFE       BLOB SUB_TYPE 1 SEGMENT SIZE 3000 CHARACTER SET WIN1252 COLLATE WIN1252,"
'        sql = sql & " MODELO_CFE            VARCHAR(4),"
'        sql = sql & " SERIE_SAT_CFE         VARCHAR(30),"
'        sql = sql & " CAIXA_CFE             VARCHAR(50),"
'        sql = sql & " PDV_CFE               VARCHAR(50),"
'        sql = sql & " NRO_CAIXA_CFE         VARCHAR(15),"
'        sql = sql & " CPF_CLIENTE_CFE       VARCHAR(14),"
'        sql = sql & " VALOR_TRIBUTOS_CFE    DOUBLE PRECISION,"
'        sql = sql & " BASE_ICMS_CFE         DOUBLE PRECISION,"
'        sql = sql & " VALOR_ICMS_CFE        DOUBLE PRECISION,"
'        sql = sql & " TOTAL_BRUTO_CFE       NUMERIC(12,2),"
'        sql = sql & " TOTAL_DESCONTO_CFE    NUMERIC(12,2),"
'        sql = sql & " TOTAL_ACRESCIMO_CFE   NUMERIC(12,2),"
'        sql = sql & " TOTAL_CFE             NUMERIC(12,2),"
'        sql = sql & " VALOR_PAGO_CFE        NUMERIC(12,2),"
'        sql = sql & " TROCO_CFE             NUMERIC(12,2),"
'        sql = sql & " CANCELADO             Char (1))"
'
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        sql = "ALTER TABLE CFE ADD CONSTRAINT PK_CFE PRIMARY KEY (ID)"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        'cria GENERATOR
'        sql = "CREATE GENERATOR GEN_CFE_ID1 "
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        sql = "SET GENERATOR GEN_CFE_ID1 TO 0"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        'cria TRIGGER PARA AUTONUMERADOR
'        sql = " CREATE TRIGGER CFE_BI FOR CFE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
'        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_CFE_ID1, 1); END  "
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'    End If
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    'IBTP
'    Call Verifica_IBPT
'
'
'    'ALTERAÇÃO CADASTRO DE PRODUTOS EM 08/09/2015
'    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'PRODUTO' AND rdb$field_name = 'CFOP' "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) = 0 Then
'            sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'PRODUTO' AND rdb$field_name = 'INATIVO' "
'            Set Rstemp5 = New ADODB.Recordset
'            Rstemp5.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'            If Rstemp5(0) = 0 Then
'                sql = "ALTER TABLE PRODUTO ADD CFOP varchar(4), ADD INATIVO char(1)"
'                Cnn.Execute sql
'                Cnn.CommitTrans
'            Else
'                sql = "DROP VIEW VIEW_ESTOQUE_NEG "
'                Cnn.Execute sql
'                Cnn.CommitTrans
'
'                sql = "DROP VIEW VIEW_SO_ESTOQUE_NEG "
'                Cnn.Execute sql
'                Cnn.CommitTrans
'
'                sql = "DROP VIEW VIEW_PRODUTO_DESCRIC "
'                Cnn.Execute sql
'                Cnn.CommitTrans
'
'                sql = "ALTER TABLE PRODUTO DROP INATIVO "
'                Cnn.Execute sql
'                Cnn.CommitTrans
'
'                'sql = "ALTER TABLE PRODUTO DROP INATIVO "
'                sql = "ALTER TABLE PRODUTO ADD CFOP varchar(4), ADD INATIVO char(1)"
'                Cnn.Execute sql
'                Cnn.CommitTrans
'
'                sql = " CREATE VIEW VIEW_ESTOQUE_NEG(CODIGO_INTERNO,PRODUTO,SALDO_EM_ESTOQUE,MARCA,GRUPO,INATIVO) AS "
'                sql = sql & "SELECT PRODUTO.CODIGO_INTERNO,PRODUTO.DESCRICAO AS PRODUTO,ESTOQUE.SALDO_EM_ESTOQUE, MARCAS.DESCRICAO AS MARCA,GRUPO.DESCRICAO AS GRUPO,PRODUTO.INATIVO  "
'                sql = sql & " FROM ((PRODUTO INNER JOIN ESTOQUE ON PRODUTO.CODIGO = ESTOQUE.CODIGO_PRODUTO) INNER JOIN MARCAS ON PRODUTO.MARCA = MARCAS.CODIGO) "
'                sql = sql & " INNER JOIN GRUPO ON PRODUTO.GRUPO = GRUPO.CODIGO WHERE ESTOQUE.SALDO_EM_ESTOQUE <=0 "
'                Cnn.Execute sql
'                Cnn.CommitTrans
'
'                sql = "CREATE VIEW VIEW_SO_ESTOQUE_NEG(CODIGO_INTERNO,NOME_PRODUTO,ULTIMA_VENDA,ULTIMA_COMPRA,SALDO_EM_ESTOQUE,QTD_MINIMA,COMPRAR,GRUPO,INATIVO) AS "
'                sql = sql & " SELECT B.CODIGO_INTERNO, B.DESCRICAO AS NOME_PRODUTO, B.ULTIMA_VENDA, B.ULTIMA_COMPRA, A.SALDO_EM_ESTOQUE, B.QTD_MINIMA,"
'                sql = sql & " (B.QTD_MINIMA - A.SALDO_EM_ESTOQUE) AS COMPRAR, G.DESCRICAO AS GRUPO,B.INATIVO FROM ESTOQUE A, PRODUTO B,  GRUPO G WHERE A.CODIGO_PRODUTO = B.CODIGO"
'                sql = sql & " AND G.CODIGO = B.GRUPO  AND (A.SALDO_EM_ESTOQUE < B.QTD_MINIMA)  and (A.SALDO_EM_ESTOQUE >= 0) ORDER BY NOME_PRODUTO"
'                Cnn.Execute sql
'                Cnn.CommitTrans
'
'                sql = "CREATE VIEW VIEW_PRODUTO_DESCRIC(CODIGO,CODIGO_INTERNO,DESCRICAO,PRECO,UNIDADE,SALDO_EM_ESTOQUE,ULTIMA_VENDA,ULTIMA_COMPRA,DATA_CAD_ALT,"
'                sql = sql & " PRECO_ATACADO,"
'                sql = sql & " PRECO_MINIMO_ATACADO,"
'                sql = sql & " INATIVO) AS"
'                sql = sql & " Select A.CODIGO,A.CODIGO_INTERNO,A.DESCRICAO,A.PRECO,A.UNIDADE, B.SALDO_EM_ESTOQUE,A.ULTIMA_VENDA,A.ULTIMA_COMPRA,A.DATA_CAD_ALT, A.PRECO_ATACADO, A.PRECO_MINIMO_ATACADO, A.INATIVO FROM PRODUTO A, ESTOQUE B WHERE A.Codigo = B.CODIGO_PRODUTO ORDER BY A.Descricao ASC"
'                Cnn.Execute sql
'                Cnn.CommitTrans
'            End If
'            Rstemp5.Close
'
'            sql = "UPDATE PRODUTO SET CFOP = '5405'"
'            Cnn.Execute sql
'            Cnn.CommitTrans
'
'    End If
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'
'
'    ' VERIFICA SE O CAMPO EXISTE SE NAO CRIA
'    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'ARQ_ESTOQUE' "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) = 0 Then
'         sql = "CREATE TABLE ARQ_ESTOQUE (ID_ESTOQUE INTEGER NOT NULL, ID_PRODUTO DOUBLE PRECISION NOT NULL, ID_FORNECEDOR DOUBLE PRECISION, "
'         sql = sql & " DOCUMENTO DOUBLE PRECISION, DATA TIMESTAMP, SALDO_ANTERIOR DOUBLE PRECISION, ENTRADA DOUBLE PRECISION,SAIDA DOUBLE PRECISION,"
'         sql = sql & " SALDO_AJUSTADO DOUBLE PRECISION,SALDO_ATUAL DOUBLE PRECISION,SALDO_BONIFIC DOUBLE PRECISION, PRECO_CUSTO DOUBLE PRECISION, PRECO_VENDA DOUBLE PRECISION,"
'         sql = sql & " ENTRADA_BONIF DOUBLE PRECISION,SAIDA_BONIFIC DOUBLE PRECISION, TRANSFERENCIA DOUBLE PRECISION, QUEBRA DOUBLE PRECISION, "
'         sql = sql & " JUSTIFICATIVA VARCHAR(100)CHARACTER SET WIN1252,PRIMARY KEY (ID_ESTOQUE))"
'         Cnn.Execute sql
'         Cnn.CommitTrans
'
'         'sql = "DROP TABLE ARQ_ESTOQUE"
'         'Cnn.Execute sql
'         'Cnn.CommitTrans
'
'         ' Indices
'         sql = "CREATE INDEX X_ID_ ON ARQ_ESTOQUE (ID_ESTOQUE)"
'         Cnn.Execute sql
'         Cnn.CommitTrans
'
'         'cria GENERATOR
'         sql = "CREATE GENERATOR GEN_SEQ_ARQ_ESTOQUE "
'         Cnn.Execute sql
'         Cnn.CommitTrans
'
'         sql = "SET GENERATOR GEN_SEQ_ARQ_ESTOQUE TO 0"
'         Cnn.Execute sql
'         Cnn.CommitTrans
'
'          'cria TRIGGER PARA AUTONUMERADOR
'         sql = " CREATE TRIGGER ARQ_ESTOQUE FOR ARQ_ESTOQUE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
'         sql = sql & "if (NEW.ID_ESTOQUE is NULL) then NEW.ID_ESTOQUE = GEN_ID(GEN_SEQ_ARQ_ESTOQUE, 1); END  "
'         Cnn.Execute sql
'         Cnn.CommitTrans
'
'    End If
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'PRODUTO_FORNECEDOR' "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) = 0 Then
'        sql = "CREATE TABLE PRODUTO_FORNECEDOR ("
'        sql = sql & "    COD_FORNECEDOR  VARCHAR(50) NOT NULL,"
'        sql = sql & "    COD_ENTRADA     VARCHAR(30) NOT NULL,"
'        sql = sql & "    COD_INTERNO     VARCHAR(30) NOT NULL,"
'        sql = sql & "    CODIGO          DOUBLE PRECISION NOT NULL)"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'    End If
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    'select entrada e saida
'    sql = "Select Sum(Saida) as Total_Saida, Sum(Entrada) as Total_Entrada FROM "
'    sql = sql & "(select Cast(0 as numeric(15,3)) as Saida, qtde as Entrada FROM itens_entrada "
'    sql = sql & " Union All"
'    sql = sql & " select sum(qtde) as Saida, Cast(0 as numeric(15,3)) as Entrada from itens_saida) as TMP"
''    Set Rstemp6 = New ADODB.Recordset
''    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
''    If Rstemp6(0) > 0 Then
''
''    End If
''    Rstemp6.Close
''    Set Rstemp6 = Nothing
'
'    'ou mais completo
'    'sql = "Select Codigo, Tipo, Data, Saida, Entrada, MotSai, MotEntrada, pagSaida From "
'    'sql = sql & " (select codEntrada as Codigo  , Cast('E'as char(1)) as Tipo , datEntrada as Data, Cast(0 as numeric(15,3)) as Saida,"
'    'sql = sql & " vlrEntrada as Entrada, Cast(0 as integer) as MotSai, motEntrada as MotEntrada, Cast(0 as integer) as pagSaida From tbEntrada"
'    'sql = sql & " Union All "
'    'sql = sql & " select codSaida as Codigo, Cast('E'as char(1)) as Tipo, datSaida as Data, vlrSaida as Saida, Cast(0 as numeric(15,3)) as Entrada,"
'    'sql = sql & " motSaida as MotSai, Cast(0 as integer) as MotEntrada, pagSaida as pagSaida from tbSaida) as TMP"
'
'    sql = ""
'
'    sql = "SELECT * FROM CAD_MENUS WHERE MENU_CD_CODI = 74 "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, 1, 2
'    If Rstemp6.RecordCount = 0 Then
'        sql = "INSERT INTO CAD_MENUS VALUES("
'        sql = sql & "'SisAdven',"
'        sql = sql & "74,"
'        Menu = UCase("ALTERA_PRECOS_PRODUTOS")
'        sql = sql & "'" & Menu & "',"
'        sql = sql & "'Cadastro - Produtos - Alterar Preços')"
'        Cnn.Execute sql
'    End If
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    '*** Fabio Reinert (Alemão) - 07/2017 - Verificação de arquivo CEST.TXT e se existir cria nova TAB_CEST - Inicio
'    '***
'    'CEST     --->  Primeiro verificar se existe o arquivo CEST.TXT
'    If Len(Dir(App.Path & "\CEST.TXT")) > 0 Then
'        strSql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'TAB_CEST' "
'        Set Rstemp6 = New ADODB.Recordset
'        Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'        If Rstemp6(0) = 0 Then
'            strSql = "CREATE TABLE TAB_CEST (CEST VARCHAR(7) NOT NULL,NCM VARCHAR(8),DESCRICAO VARCHAR(512) ) "
'            Cnn.Execute strSql
'            Cnn.CommitTrans
'        Else
'            sql = "DELETE FROM TAB_CEST "
'            Cnn.Execute strSql
'            Cnn.CommitTrans
'        End If
'        Call sPopula_Tab_Cest   'Sub que popula a tab_cest com o conteúdo do arquivo texto CEST.TXT
'        Kill App.Path & "\CEST.TXT"
'          '***
'    Else  '*** Perguntar se não tiver o arquivo e a tabela não existir o que fazer?
'          '***
'        strSql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'TAB_CEST' "
'        Set Rstemp6 = New ADODB.Recordset
'        Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'        If Rstemp6(0) = 0 Then
'            strSql = "CREATE TABLE TAB_CEST (CEST VARCHAR(7) NOT NULL,NCM VARCHAR(8),DESCRICAO VARCHAR(512) ) "
'            Cnn.Execute strSql
'            Cnn.CommitTrans
'        End If
'    End If
'    If Rstemp6.State = adStateOpen Then
'        Rstemp6.Close
'    End If
'    Set Rstemp6 = Nothing
'    '*** Fabio Reinert (Alemão) - 07/2017 - Verificação de arquivo CEST.TXT e se existir cria nova TAB_CEST - Inicio
'    '***
'
'
'    'CEST
''    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'TAB_CEST' "
''    Set Rstemp6 = New ADODB.Recordset
''    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
''    If Rstemp6(0) = 0 Then
''        sql = "CREATE TABLE TAB_CEST (CEST VARCHAR(7) NOT NULL,NCM VARCHAR(8),DESCRICAO VARCHAR(512) ) "
''        Cnn.Execute sql
''        Cnn.CommitTrans
''    End If
''    Rstemp6.Close
''    Set Rstemp6 = Nothing
''
''    Dim LineofText As String
''
''    Dim Linha As String
''    Dim Separa() As String
''
''    X = 0
''    i = 0
''    sql = ""
'
'
''Open App.Path & "\CEST.txt" For Input As #2
''    Do While Not EOF(1)
''        'pega a linha do TXT
''        Line Input #2, Linha
''        'separa o texto antes de gravar
''        Separa() = Split(Linha, ";")
''        'grava o registro na tabela
''        'TBCliente1(0) = Separa(0)
''        Texto = Separa(0)
''
''        Line Input #2, Linha
''        texto1 = Separa(1)
''        texto2 = Separa(2)
''        texto3 = Separa(3)
''        texto4 = Separa(4)
''        texto5 = Separa(5)
''    Loop
''Close #2
'
'
''
''    Dim Cnn_a_Importar As New ADODB.Connection
''    Set Cnn_a_Importar = New ADODB.Connection
''
''    With Cnn_a_Importar
''        .CursorLocation = adUseClient
''            'SqlServer 2000
''            .Open "Provider=SQLOLEDB.1;Password=123;Persist Security Info=True;User ID=BILL;Initial Catalog=ArqdadosPDV;Data Source=servidor"
''
''           'firebird
''             '.Open "Provider=IBOLE.Provider.v4;Persist Security Info=False;Data Source=GUSTAVO:c:\Sistema SisAdven\ARQDADOS.GDB"
''            '.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=bill2;Initial Catalog=SimplesFIL;Data Source=SERVIDOR\MSSQLSERVER_R2"
''
''            'firebird
''            '.Open "Provider=IBOLE.Provider.v4;Persist Security Info=False;Data Source=servidor:c:\Sistema SisAdven\gilda.FDB"
''
''    End With
''
''
''    sql = "SELECT * FROM TAB_CEST ORDER BY NCM "
''    Set Rstemp = New ADODB.Recordset
''    Rstemp.Open sql, Cnn_a_Importar, 1, 2
''    If Rstemp.RecordCount > 0 Then
''        Rstemp.MoveLast
''        Rstemp.MoveFirst
''        While Not Rstemp.EOF
''            sql = "INSERT INTO TAB_CEST VALUES ("
''            sql = sql & "'" & Rstemp(0) & "',"
''            sql = sql & "'" & Rstemp(1) & "',"
''            sql = sql & "'" & Rstemp(2) & "')"
''            Cnn.Execute sql
''            Cnn.CommitTrans
''            Rstemp.MoveNext
''        Wend
''
''    End If
''    Rstemp.Close
'
'
''    Cnn.Execute "delete from tab_cest"
''    Cnn.CommitTrans
''
''    'importa cest
''    Dim Separa() As String
''
''    LineofText = ""
''    sql = ""
''    Open App.Path & "\TAB_EST_FIREBIRD.sql" For Input As #1
''    'Open App.Path & "\CEST_old.txt" For Input As #1
''    Do While Not EOF(1)
''    '            sql = ""
''    '            For X = 1 To 2 'necessário por conta da quebra de linha
''    '                Line Input #1, LineofText
''    '                'Debug.Print LineofText
''    '                If X = 1 Then
''    '                    sql = LineofText
''    '                Else
''    '                    sql = sql & LineofText
''    '                End If
''    '            Next X
''    '            'Debug.Print sql
''    '            Cnn.Execute sql
''    '            Cnn.CommitTrans
''        Line Input #1, LineofText
''        If Len(LineofText) > 0 Then
''            Cnn.Execute LineofText
''            Cnn.CommitTrans
''        Else
''            aaa = LineofText
''        End If
''    Loop
''    Close #1
''
''
''    MsgBox "Tabela CEST criada com sucesso...!", vbInformation, "Aviso"
'
'    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'PRODUTO' AND  rdb$field_name='CEST'"
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) = 0 Then
'        sql = "ALTER TABLE PRODUTO ADD CEST VARCHAR(7)"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'    End If
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'REL_RANKING_VENDAS_VENDEDOR '"
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) = 0 Then
'        sql = "CREATE TABLE REL_RANKING_VENDAS_VENDEDOR (COD_VEND double PRECISION, RAZAO_SOCIAL VARCHAR(60), QTDE DOUBLE PRECISION, TOTAL double PRECISION)"
'        Cnn.Execute sql
'    End If
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    'CURVA ABC
'    sql = "SELECT  P.CODIGO_INTERNO, P.DESCRICAO, SUM(I.QTDE) AS"
'    sql = sql & " SUBTOTAL, SUM(I.QTDE * I.VALOR_UNITARIO) / SUM(I.VALOR_TOTAL) * 100 AS"
'    sql = sql & " CURVA_ABC, V.Data_NF,TOTAL_SAIDA FROM ITENS_SAIDA I INNER JOIN SAIDAS_PRODUTO V ON I.SEQUENCIA = V.SEQUENCIA"
'    sql = sql & " INNER JOIN PRODUTO P ON I.CODIGO_PRODUTO = P.CODIGO"
'    'sql = sql & " WHERE  C.DATA_NF BETWEEN  '" & Format(mskDe.Text, "MM/DD/yyyy") & "'"
'    sql = sql & " WHERE V.Data_NF BETWEEN '10/11/2016' AND '10/11/2016'"
'    sql = sql & " GROUP BY V.Data_NF, P.CODIGO_INTERNO, P.DESCRICAO, V.TOTAL_SAIDA"
'    sql = sql & " ORDER BY SUM(I.QTDE * I.VALOR_UNITARIO) / SUM(I.VALOR_total) * 100 DESC"
'    'Set Rstemp = New ADODB.Recordset
'    'Rstemp.Open sql, Cnn, 1, 2
'    'If Rstemp.RecordCount <> 0 Then
'    'End If
'
'    'CURVA ABC CLIENTES
'    sql = "CREATE VIEW CURVA_ABC_CLIENTES(DATA,CLIENTE,CONTATO,CIDADE,UF,FONE,ULT_VENDA,TOTAL)"
'    sql = sql & " AS select  v.DATA_NF as data, c.codigo||' - '||c.RAZAO_SOCIAL as nome_cliente, C.CONTATO,            "
'    sql = sql & "c.CIDADE_END_PRINCIPAL AS CIDADE,"
'    sql = sql & "c.UF_END_PRINCIPAL AS UF,"
'    sql = sql & "c.FONE1 AS FONE,"
'    sql = sql & "max(v.DATA_NF) as ULT_VENDA,"
'    sql = sql & "sum(v.TOTAL_SAIDA) as TOTAL"
'    sql = sql & "from saidas_produto v"
'    sql = sql & "inner join cliente c on (c.codigo = v.CODIGO_CLIENTE)"
'    sql = sql & "group by 1,2,3,4,5,6"
'    'Set Rstemp = New ADODB.Recordset
'    'Rstemp.Open sql, Cnn, 1, 2
'    'If Rstemp.RecordCount <> 0 Then
'    'End If
'
'
'    'TRATA TABELA SAIDAS_PRODUTO CAMPO PDV
'    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'TRANSPORTADORA' AND rdb$field_name = 'DDD_TELEFONE' "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) = 0 Then
'        sql = "DROP TABLE TRANSPORTADORA"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        sql = "CREATE TABLE TRANSPORTADORA (CODIGO DOUBLE PRECISION NOT NULL,RAZAO_SOCIAL varchar (60),  CEP_PRINCIPAL varchar (9),"
'        sql = sql & "ENDERECO_PRINCIPAL varchar (60),NRO_END_PRINCIPAL varchar (10),COMPL_END_PRINCIPAL varchar (30),BAIRRO_END_PRINCIPAL varchar (45),"
'        sql = sql & "CIDADE_END_PRINCIPAL varchar (60),UF_END_PRINCIPAL varchar (2),CNPJ varchar (18),INSC_ESTADUAL varchar (30),SITE varchar (40),"
'        sql = sql & "EMAIL varchar (60),NOME varchar (40),DEPTO varchar (30),EMAIL_CONTATO varchar (60),DDD_TELEFONE DOUBLE PRECISION,"
'        sql = sql & "TELEFONE varchar (30), DDD_FAX DOUBLE PRECISION,   FAX varchar (30)) "
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        'Indices
'        '''sql = "CREATE INDEX ID_X_SEQUENCIA ON ITENS_SAIDA (SEQUENCIA)"
'        sql = "ALTER TABLE TRANSPORTADORA ADD CONSTRAINT I101 PRIMARY KEY (CODIGO);"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        sql = "CREATE INDEX PK_RAZAO_SOCIAL_ ON TRANSPORTADORA (RAZAO_SOCIAL);"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        sql = "INSERT INTO TRANSPORTADORA (CODIGO,RAZAO_SOCIAL) VALUES (1,'NOSSO CARRO')"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'    End If
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'
''*** Fabio Reinert ( Alemão ) - 08/2017 - Alteração da tabela FORMAS - Novo conteudo das formas de pagto. - Inicio
'    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'FORMAS' "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) = 0 Then
'
'        sql = "CREATE TABLE FORMAS (CODIGO DOUBLE PRECISION NOT NULL ,DESCRICAO varchar (60) )"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        'cria GENERATOR
'        sql = "CREATE GENERATOR GEN_SEQ_FORMAS "
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        sql = "SET GENERATOR GEN_SEQ_FORMAS TO 0"
'        Cnn.Execute sql
'        Cnn.CommitTrans
'
'        'cria TRIGGER PARA AUTONUMERADOR
'        sql = " CREATE TRIGGER FORMAS_BI FOR FORMAS ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
'        sql = sql & "if (NEW.CODIGO is NULL) then NEW.CODIGO = GEN_ID(GEN_SEQ_FORMAS, 1); END  "
'        Cnn.Execute sql
'        Cnn.CommitTrans
'    End If
'
'    sql = "DELETE FROM FORMAS"
'    Cnn.Execute sql
'    Cnn.CommitTrans
'
'    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Dinheiro')"
'    Cnn.Execute sql
'    Cnn.CommitTrans
'    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Cartão de Débito')"
'    Cnn.Execute sql
'    Cnn.CommitTrans
'    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Cartão de Crédito')"
'    Cnn.Execute sql
'    Cnn.CommitTrans
'    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Pendente')"
'    Cnn.Execute sql
'    Cnn.CommitTrans
'    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Cheques')"
'    Cnn.Execute sql
'    Cnn.CommitTrans
'    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('DOC/TED')"
'    Cnn.Execute sql
'    Cnn.CommitTrans
'    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Boleto')"
'    Cnn.Execute sql
'    Cnn.CommitTrans
'    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Cartão da Loja')"
'    Cnn.Execute sql
'    Cnn.CommitTrans
'
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
''*

End Sub

'*** Fabio Reinert ( Alemão ) - 07/2017 - Criação da tabela TAB_CEST Através de arquivo texto - Inicio
'***
'*** SUB QUE LE O ARQUIVO TEXTO cest.txt e popula a TAB_CEST
Private Sub sPopula_Tab_Cest()
    Open App.Path & "\CEST.txt" For Input As #1
    Do While Not EOF(1)
         strSql = ""
         For X = 1 To 2 'necessário por conta da quebra de linha
             Line Input #1, LineofText
             If EOF(1) Then
                Exit For
             End If
             'Debug.Print LineofText
             If X = 1 Then
                strSql = LineofText
             Else
                strSql = strSql & LineofText
             End If
         Next X
         If Not EOF(1) Then
            Debug.Print strSql
            Cnn.Execute strSql
            Cnn.CommitTrans
        End If
        'Line Input #1, LineofText
        'If Len(LineofText) > 0 Then
'            Cnn.Execute LineofText
'            Cnn.CommitTrans
        'Else
        '    aaa = LineofText
        'End If
    Loop
    Close #1

    'MsgBox "Tabela CEST criada com sucesso...!", vbInformation, "Aviso"

End Sub
'***
'*** Fabio Reinert ( Alemão ) - 07/2017 - Criação da tabela TAB_CEST Através de arquivo texto - Fim
'***

Private Sub Verifica_IBPT()

    Dim LineofText As String
    
    Dim Linha As String
    Dim Separa() As String

    
    'sql = "DROP TABLE ALIQUOTAS_IBPT "
    'Cnn.Execute sql
    'Cnn.CommitTrans
    
    'tabela IBPT
    sql = "SELECT  COUNT(*) FROM RDB$RELATIONS WHERE  RDB$RELATION_NAME = 'ALIQUOTAS_IBPT'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        'sql = "CREATE TABLE ALIQUOTAS_IBPT(CODIGO  DOUBLE PRECISION,  ALIQ_NAC  VARCHAR(15),  EX  VARCHAR(15),"
        'sql = sql & "TABELA DOUBLE PRECISION,ALIQ_IMP VARCHAR(15))"
        'Cnn.Execute sql
        'Cnn.CommitTrans
        
        Cnn.Execute "DROP TABLE ALIQUOTAS_IBPT"
        Cnn.CommitTrans
        
        LineofText = ""
        sql = ""
        Open App.Path & "\IBPT.sql" For Input As #1
        Do While Not EOF(1)
'            sql = ""
'            For X = 1 To 2 'necessário por conta da quebra de linha
'                Line Input #1, LineofText
'                'Debug.Print LineofText
'                If X = 1 Then
'                    sql = LineofText
'                Else
'                    sql = sql & LineofText
'                End If
'            Next X
'            'Debug.Print sql
'            Cnn.Execute sql
'            Cnn.CommitTrans
            Line Input #1, LineofText
            If Len(LineofText) > 0 Then
                Cnn.Execute LineofText
                Cnn.CommitTrans
            Else
                aaa = LineofText
            End If
        Loop
        Close #1
        
        sql = "CREATE INDEX ALIQUOTAS_IBPT_IDX1 ON ALIQUOTAS_IBPT (CODIGO)"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        GoTo continua
    Else

continua:
        Screen.MousePointer = 11
        
        If Dir(App.Path & "\IBPT.CSV") <> "" Then
            'File exists
            sql = "DELETE FROM ALIQUOTAS_IBPT "
            Cnn.Execute sql
            Cnn.CommitTrans
            
            i = 1
            
            Open App.Path & "\IBPT.CSV" For Input As #1
            Do While Not EOF(1)
                NCM = ""
                EX = ""
                Tabela = ""
                ALIQ_NAC = ""
                ALIQ_IMP = ""
                
                'pega a linha do TXT
                Line Input #1, Linha
                'separa o texto antes de gravar
                Separa() = Split(Linha, ";")
                'grava o registro na tabela
                If i > 1 Then
                    NCM = Separa(0)
                    EX = Separa(1)
                    Tabela = Separa(2)
                    ALIQ_NAC = Separa(4)
                    ALIQ_IMP = Separa(5)
                    sql = "INSERT INTO ALIQUOTAS_IBPT (CODIGO, EX, TABELA, ALIQ_NAC, ALIQ_IMP)"
                    sql = sql & " VALUES ( "
                    sql = sql & "'" & NCM & "',"
                    sql = sql & "NULL,"
                    sql = sql & "'" & Tabela & "',"
                    sql = sql & "'" & ALIQ_NAC & "',"
                    sql = sql & "'" & ALIQ_IMP & "')"
                    Cnn.Execute sql
                    Cnn.CommitTrans
                End If
                i = 1 + 1
            Loop
            'fecha o arquivo texto
            Close #1
            
            'remove arquivo
            Kill App.Path & "\IBPT.CSV"
            Screen.MousePointer = 1
        End If
        Screen.MousePointer = 1
    End If
    
    Rstemp6.Close
    Set Rstemp6 = Nothing

    
    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'PRODUTO' AND  rdb$field_name='NCM'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "ALTER TABLE PRODUTO ADD NCM VARCHAR(8)"
        Cnn.Execute sql
        Cnn.CommitTrans
    End If

    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    LineofText = ""
    
    sql = "SELECT  COUNT(*) FROM RDB$RELATIONS WHERE  RDB$RELATION_NAME = 'MUNICIPIOS'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
'        sql = "CREATE TABLE MUNICIPIOS (CUF INTEGER,UF VARCHAR(2),XUF VARCHAR(120) CHARACTER SET WIN1252, CMUN  VARCHAR(7) , XMUN  VARCHAR(120) CHARACTER SET WIN1252, primary key(cmun));"
'        Cnn.Execute sql
        Open App.Path & "\municipios-insert-firebird.sql" For Input As #1
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

    'VERIFICA SE O CAMPO EXISTE SE NAO CRIA
    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'PRODUTO' AND  rdb$field_name='ALIQUOTA_ECF'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "ALTER TABLE PRODUTO ADD ALIQUOTA_ECF DOUBLE PRECISION "
        Cnn.Execute sql
        
        sql = "UPDATE PRODUTO SET ALIQUOTA_ECF = 5 "
        Cnn.Execute sql
        Cnn.CommitTrans
    End If
        
    
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    'verifica se stored procedure ou tabela existe e cria
    sql = "select count(*) from RDB$RELATION_FIELDS where RDB$RELATION_NAME = 'VIEW_LISTA_PROD' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "CREATE VIEW VIEW_LISTA_PROD (CODIGO,CODIGO_INTERNO,DESCRICAO,VLRCUSTO,PRECO,UNIDADE,SALDO_EM_ESTOQUE, "
        sql = sql & "MARCA,ULTIMA_VENDA,ULTIMA_COMPRA) AS "
        sql = sql & " select A.CODIGO,A.CODIGO_INTERNO,A.DESCRICAO,A.VLRCUSTO, A.PRECO,A.UNIDADE,B.SALDO_EM_ESTOQUE,M.DESCRICAO AS MARCA, A.ULTIMA_VENDA,"
        sql = sql & " A.ULTIMA_COMPRA FROM PRODUTO A, ESTOQUE B, MARCAS M  WHERE A.CODIGO = B.CODIGO_PRODUTO"
        sql = sql & " AND M.CODIGO = A.MARCA  ORDER BY A.DESCRICAO ASC "
        Cnn.Execute sql
        
        Cnn.Close
        Call Conecta_Banco
    End If
    
   'VERIFICA SE O CAMPO EXISTE SE NAO CRIA
    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'EMPRESA' AND  rdb$field_name='NOME_FANTASIA'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "ALTER TABLE EMPRESA ADD NOME_FANTASIA VARCHAR(30), ADD INSC_ESTADUAL VARCHAR(19), ADD NRO_ENDERECO VARCHAR(9)"
        Cnn.Execute sql
        Cnn.CommitTrans
    End If
    
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    'verifica se stored procedure ou tabela existe e cria
    sql = "select count(*) from RDB$RELATION_FIELDS where RDB$RELATION_NAME = 'LOTE_NFE' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        
        sql = "CREATE TABLE LOTE_NFE (ID INTEGER NOT NULL, LOTE DOUBLE PRECISION, NRO_RECIBO VARCHAR(20), PRIMARY KEY (ID))"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_LOT_ID1 "
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "SET GENERATOR GEN_LOT_ID1 TO 0"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        'cria TRIGGER PARA AUTONUMERADOR
        sql = " CREATE TRIGGER LOTE_NFE_BI FOR LOTE_NFE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_LOT_ID1, 1); END  "
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "INSERT INTO LOTE_NFE (LOTE)"
        sql = sql & "values (1)"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        'TABELA NFE
        '''sql = "CREATE TABLE NFE (ID INTEGER NOT NULL, CHAVE_NFE VARCHAR(50), NRO_LOTE VARCHAR(100), NRO_RECIBO VARCHAR(20), NRO_PROTOCOLO VARCHAR(20), NRO_PEDIDO DOUBLE PRECISION, NRO_NF DOUBLE PRECISION, NRO_CANCELAMENTO_NF VARCHAR(20), STATUS VARCHAR(20), PRIMARY KEY (ID))"
        sql = "CREATE TABLE NFE (ID INTEGER NOT NULL, CHAVE_NFE VARCHAR(50), NRO_LOTE VARCHAR(100), NRO_RECIBO VARCHAR(20), NRO_PROTOCOLO VARCHAR(20), NRO_PEDIDO BLOB SUB_TYPE 1, NRO_NF DOUBLE PRECISION, NRO_CANCELAMENTO_NF VARCHAR(20), STATUS VARCHAR(20), PRIMARY KEY (ID))"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_NFE_ID1 "
        Cnn.Execute sql
        Cnn.CommitTrans
        sql = "SET GENERATOR GEN_NFE_ID1 TO 0"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        'cria TRIGGER PARA AUTONUMERADOR
        sql = " CREATE TRIGGER NFE_BI FOR NFE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_NFE_ID1, 1); END  "
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "CREATE TABLE DEVOLUCAO_NFE (ID INTEGER NOT NULL, SEQUENCIA DOUBLE PRECISION NOT NULL,DATA_NF DATE,NF DOUBLE PRECISION,"
        sql = sql & " COD_FORNECEDOR DOUBLE PRECISION,TOTAL_SAIDA DOUBLE PRECISION,PRIMARY KEY (ID))"
        Cnn.Execute sql
        
        sql = "CREATE TABLE ITENS_DEVOLUCAO_NFE (ID INTEGER NOT NULL, SEQUENCIA DOUBLE PRECISION, CODIGO_PRODUTO DOUBLE PRECISION, QTDE DOUBLE PRECISION,"
        sql = sql & "VALOR_UNITARIO DOUBLE PRECISION,VALOR_TOTAL DOUBLE PRECISION, PRIMARY KEY (ID))"
        Cnn.Execute sql
        
        'sql = "ALTER TABLE ITENS_DEVOLUCAO_NFE ADD FOREIGN KEY (SEQUENCIA) REFERENCES DEVOLUCAO_NFE (SEQUENCIA)"
        'Cnn.Execute sql
        
        
        ''''''''''''''''''''
         'Primary Keys     SITE REFEENCIA = "http://www.firebirdsql.org/dotnetfirebird/create-a-new-database-from-an-sql-script.html"
        'sql = "ALTER TABLE DEVOLUCAO_NFE ADD PRIMARY KEY (ID)"
        'Cnn.Execute sql
        
        ' Indices
        sql = "CREATE INDEX ID_X ON DEVOLUCAO_NFE (ID)"
        Cnn.Execute sql
        
        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_SEQ_DEVOLUCAO_NFE "
        Cnn.Execute sql
        
        sql = "SET GENERATOR GEN_SEQ_DEVOLUCAO_NFE TO 0"
        Cnn.Execute sql
        
         'cria TRIGGER PARA AUTONUMERADOR
        sql = " CREATE TRIGGER DEVOLUCAO_NFE FOR DEVOLUCAO_NFE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_SEQ_DEVOLUCAO_NFE, 1); END  "
        Cnn.Execute sql
        
        
        'ITENS_DEVOLUCAO_NFE
        sql = "CREATE INDEX ID_Y ON ITENS_DEVOLUCAO_NFE (ID)"
        Cnn.Execute sql
        
        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_SEQ_ITENS_DEVOLUCAO_NFE "
        Cnn.Execute sql
        
        sql = "SET GENERATOR GEN_SEQ_ITENS_DEVOLUCAO_NFE TO 0"
        Cnn.Execute sql
        
         'cria TRIGGER PARA AUTONUMERADOR
        sql = " CREATE TRIGGER ITENS_DEVOLUCAO_NFE FOR ITENS_DEVOLUCAO_NFE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_SEQ_ITENS_DEVOLUCAO_NFE, 1); END  "
        Cnn.Execute sql
        
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
        End If
        
        sql = "ALTER TABLE NFE ADD DATA_EMISSAO DATE "
        Cnn.Execute sql
        
        sql = "ALTER TABLE NFE ADD TOTAL_NF DOUBLE PRECISION "
        Cnn.Execute sql
        
        Cnn.CommitTrans
        
    End If
    
    Rstemp6.Close
    Set Rstemp6 = Nothing

    Exit Sub
 
End Sub



Public Sub DoAboutTxt()
'cria e abre arquivo.cdc
Open App.Path & "\SisAdVen-Credito.cdc" For Output As #1
Print #1, "SisAdVen " & Versao_Software & " Mod. Fiscal"
Print #1,
Print #1, "Autor"
Print #1, "Arlindo J.R da Silva"
Print #1,
If Situacao_Registro = True Then
    Print #1, "Licenciado"
    Print #1, UserName
Else
    Print #1, "Sistema Versão Avaliação"
    Print #1, UserName
End If
Print #1, "Novavia Soluções em Informática"
Print #1,
Print #1, "e-Mail"
Print #1, "comercial@novavia.com.br"
Print #1,
Print #1, "Fone: (11)5548-3890"
Print #1,
Close #1
End Sub


Private Sub MDIForm_Resize()
'''
'''On Error GoTo trata
'''    If sysAcesso = 1 Then
'''        FRM_TASK.Show
'''    End If
'''
'''
'''
'''Exit Sub
'''
'''trata:

If flag_Gaveta_Bematec = True Then
    Screen.MousePointer = 11
    MsgBox "Por favor, clique no botão OK e aguarde comunicação com a gaveta ", vbInformation, "Aviso"
    Retorno = Bematech_FI_AcionaGaveta()
    Call VerificaRetornoImpressora("", "", "Acionamento da Gaveta")
    Screen.MousePointer = 1
ElseIf flag_Gaveta_Elgin = True Then
    Screen.MousePointer = 11
    MsgBox "Por favor, clique no botão OK e aguarde comunicação com a gaveta ", vbInformation, "Aviso"
    'Retorno = Elgin_AcionaGaveta()
    'Call TrataRetorno(Retorno)
    Screen.MousePointer = 1
End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    menu_sair_Click
End Sub

Private Sub menu_cadastro_cliente_Click()
    menu_cadastro_cliente.Enabled = False
    frmCliente.Show 1
    
End Sub

Private Sub menu_cadastro_Empresa_Click()
    menu_cadastro_Empresa.Enabled = False
    Frm_Cad_Empresa.Show 1
End Sub

Private Sub menu_cadastro_fornecedor_Click()
    menu_cadastro_fornecedor.Enabled = False
    frmFornecedor.Show 1

End Sub

Private Sub menu_cadastro_grupo_Click()
    menu_cadastro_grupo.Enabled = False
    frmGrupo.Show 1

End Sub

Private Sub menu_cadastro_marca_Click()
    menu_cadastro_marca.Enabled = False
    frmMarca.Show 1

End Sub

Private Sub menu_cadastro_produtos_Click()
    menu_cadastro_produtos.Enabled = False
    frmProduto.Show 1
End Sub

Private Sub menu_cadastro_representante_Click()
    menu_cadastro_representante.Enabled = False
    frmRepresentante.Show 1

End Sub

Private Sub menu_cadastro_transp_Click()
    menu_cadastro_transp.Enabled = False
    frmTransportadora.Show
'MsgBox "Módulo em desenvolvimento, por favor aguarde...!", vbInformation, "Aviso"

End Sub

Private Sub menu_Agenda_Click()
    'Dim prog
    'prog = Shell(App.Path & "\petAgenda.exe " & sysAcesso & "," & sysSenha, vbNormalFocus)
    frmAgenda.Show vbModal
End Sub


Private Sub menu_Cancela_Cupom_Pendente_Click()
If MsgBox("Deseja Realmente Cancelar Cupom Pendente...?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
    Screen.MousePointer = 11
    Retorno = Bematech_FI_CancelaItemAnterior()
    'Função que analisa o retorno da impressora
    Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
End If
Screen.MousePointer = 1
End Sub

Private Sub menu_Cheques_Click()
Consultas = "S"
FrmCheques.Show 1
End Sub

Private Sub menu_consulta_cliente_Click()
    Consultas = "S"
    frmCliente.Show 1

End Sub

Private Sub menu_consulta_condpgto_Click()
    Consultas = "S"
    frmCdPgto.Show 1

End Sub

Private Sub menu_consulta_estoque_Click()
    Consultas = "S"
    frmEstoque.Show
End Sub

Private Sub menu_consulta_fornecedor_Click()
    Consultas = "S"
    frmFornecedor.Show 1

End Sub

Private Sub MENU_CONSULTA_PEDIDOS_EXCLUIDOS_Click()
frm_Pedidos_Excluidos.Show 1
End Sub

Private Sub menu_ConsultaPedidos_Click()
FrmConsultaPedidos.Show 1
End Sub

Private Sub menu_Excluir_Dados_Click()
If MsgBox("Esta operação irá apagar toda movimentação de vendas do banco de dados," & vbNewLine & "ficando somente os DADOS do cadastro de Clientes, Fornecedores, Produtos e Estoque." & vbNewLine & vbNewLine & "Deseja Realmente Efetuar esta Operação...?", vbYesNo + vbDefaultButton2 + vbOKOnly + vbExclamation, "Atenção") = vbNo Then
    Exit Sub
Else
    Call ExcluirDados
End If
End Sub

Private Sub menu_financeiro_CadCartao_Click()
    'menu_financeiro_ctapag.Enabled = False
    'tipo_financeiro = "P"
    'frmRecPag.Caption = "Contas a Pagar"
    FrmCadCartao.Show 1
End Sub

Private Sub menu_financeiro_ctapag_Click()
    menu_financeiro_ctapag.Enabled = False
    tipo_financeiro = "P"
    frmRecPag.Caption = "Contas a Pagar"
    frmRecPag.Show 1
End Sub

Private Sub menu_financeiro_ctarec_Click()

    menu_financeiro_ctarec.Enabled = False
    tipo_financeiro = "R"
    '*** Fabio Reinert - 09/2017 - Alteração do formulário do contas a receber - inicio
    'frmRecPag.Caption = "Contas a Receber"
    'frmRecPag.Show 1
    frmAReceber.Show vbModal
    '*** Fabio Reinert - 09/2017 - Alteração do formulário do contas a receber - inicio
End Sub

Private Sub menu_financeiro_ctpend_Click()
    menu_financeiro_ctpend.Enabled = False
    frmCtPend.Show 1

End Sub

Private Sub menu_financeiro_fluxo_Click()
    menu_financeiro_fluxo.Enabled = False
    frmFluxoRP.Show 1

End Sub

Private Sub menu_consulta_grupo_Click()
    Consultas = "S"
    frmGrupo.Show 1

End Sub

Private Sub menu_consulta_marca_Click()
    Consultas = "S"
    frmMarca.Show 1

End Sub

Private Sub menu_consulta_produto_Click()
    Consultas = "S"
    frmProduto.Show 1

End Sub

Private Sub menu_consulta_representantes_Click()
    Consultas = "S"
    frmRepresentante.Show

End Sub

Private Sub menu_consulta_transp_Click()
    Consultas = "S"
    frmTransportadora.Show

End Sub

Private Sub menu_Horario_Verao_Click()

If MsgBox("Executar Horário de Verão...?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
    Screen.MousePointer = 11
    Retorno = Bematech_FI_ProgramaHorarioVerao()
    'Função que analisa o retorno da impressora
    Call VerificaRetornoImpressora("", "", "Programação do Horário de Verão")
End If
Screen.MousePointer = 1
End Sub

Private Sub Menu_Movimentacao_Compras_Click()
FRM_Compras.Show 1
End Sub

Private Sub menu_movimentacao_entrada_Click()
frmMenu.Timer1.Enabled = False
frmMenu.Timer2.Enabled = False
menu_movimentacao_entrada.Enabled = False
frmEntradas.Show 1

End Sub

Private Sub menu_movimentacao_estoque_Click()
    'menu_movimentacao_estoque.Enabled = False
    'frmEstoque.Show

End Sub

Private Sub menu_movimentacao_orcamento_Click()
'    menu_movimentacao_orcamento.Enabled = False
   'frmOrcamento.Show
    MsgBox "Em Manutenção, por favor Aguarde.", vbInformation, "Aviso"
End Sub

Private Sub menu_movimentacao_saida_Click()
    menu_movimentacao_saida.Enabled = False
    'Chamar rotina de checagem segurança
    If fValida_No_Pedido() Then
        frmSaidas.Show 1
    'Else
    '    Call Fecha_Formularios
    '    Call MDIForm.Main
    End If
End Sub

Private Sub menu_relatorios_cadastros_Click()
    menu_relatorios_cadastros.Enabled = False
    frmRelCad.Show 1

End Sub

Private Sub menu_relatorios_comissao_Click()
    menu_relatorios_comissao.Enabled = False
    frmRelComissao.Show 1

End Sub

Private Sub menu_relatorios_compras_Click()
    menu_relatorios_compras.Enabled = False
    frmRelCompra.Show 1

End Sub

Private Sub menu_relatorios_estoque_saldo_Analitico_Click()
    frmMenu.MousePointer = 11
    On Error GoTo SaiImp
   ' SelecPrint.Action = 5
    Relatorios.ReportFileName = App.Path & "\relanalitdestoque.rpt"
    Relatorios.WindowTitle = "Relatório analítico de Produtos/Estoque"
    Relatorios.Action = 1
    Screen.MousePointer = 1

SaiImp:

    If Err.Number = 32755 Then
        Err.Clear
        Screen.MousePointer = 1
        Exit Sub
    End If
End Sub

Private Sub menu_relatorios_estoque_saldo_Click()
    frmMenu.MousePointer = 11
    On Error GoTo SaiImp
   ' SelecPrint.Action = 5
    Relatorios.ReportFileName = App.Path & "\relsdestoque.rpt"
    Relatorios.WindowTitle = "Relatório de Saldo em Estoque"
'    Relatorios.PrinterDriver = "winspool"
'    Relatorios.PrinterPort = "LPT1"
'    Relatorios.PrinterName = "\\arlindo\Xerox WorkCentre XK Series"
    Relatorios.Action = 1
    Screen.MousePointer = 1

SaiImp:

    If Err.Number = 32755 Then
        Err.Clear
        Screen.MousePointer = 1
        Exit Sub
    End If
End Sub

Private Sub menu_relatorios_estoque_minimo_Click()
    frmMenu.MousePointer = 11

   
    sql = "Delete from REL_ESTOQUE_MINIMO"
    Cnn.Execute sql
    
    sql = "Insert into REL_ESTOQUE_MINIMO "
    sql = sql & "Select A.CODIGO_PRODUTO FROM ESTOQUE A, PRODUTO B "
    sql = sql & "WHERE A.CODIGO_PRODUTO = B.CODIGO "
    sql = sql & "AND (A.SALDO_EM_ESTOQUE < B.QTD_MINIMA) "
    Cnn.Execute sql
    
    'Select A.CODIGO_PRODUTO, (B.QTD_MINIMA - A.SALDO_EM_ESTOQUE) as Comprar FROM ESTOQUE A, PRODUTO B WHERE A.CODIGO_PRODUTO = B.CODIGO
    'AND (A.SALDO_EM_ESTOQUE < B.QTD_MINIMA)  and (A.SALDO_EM_ESTOQUE > 0)
    
    On Error GoTo SaiImp
   '' SelecPrint.Action = 5
    
    Relatorios.Reset
    Relatorios.Destination = crptToWindow
    Relatorios.WindowState = crptMaximized
    Relatorios.ReportFileName = App.Path & "\relminestoque.rpt"
    Relatorios.WindowTitle = "Relatório de Estoque Mínimo"
    
    Relatorios.Action = 1
    Relatorios.PageZoom (100)
    frmMenu.MousePointer = 0
    
    Exit Sub

SaiImp:
    If Err.Number = 32755 Then
        Err.Clear
        Screen.MousePointer = 1
        Exit Sub
    ElseIf Err.Number = 20525 Then
        Err.Clear
       ' frmConfiguraBase.Show 1
        Screen.MousePointer = 1
        Exit Sub
    Else
        MsgBox "Ocorreu um erro: " & Err.Description & "Nro: " & Err.Number
    End If

End Sub

Private Sub menu_relatorios_estoque_saldo_por_Grupo_Click()
Flag_Destoque_Por_Grupo = True
frm_Lista_Prod_por_Grupo.Show 1
End Sub

Private Sub menu_relatorios_financeiro_Click()
    menu_relatorios_financeiro.Enabled = False
    frmMenu.Timer1.Enabled = False
    frmRelFin.Show 1

End Sub

Private Sub menu_relatorios_mdireta_Click()
   ' SQL = ""
   ' SQL = "Select * from CONFIGURACOES"
   'Set Rstemp = New ADODB.Recordset
   'Rstemp.Open sql, Cnn, 1, 2
   ' If rstemp.RecordCount <> 0 Then
       ' If Not IsNull(rstemp!END_IMP_DIVERSOS) Then
       '     Relatorios.PrinterPort = rstemp!END_IMP_DIVERSOS
       ' Else
       '     MsgBox "Não existe Impressora definida para Impressão de Relatórios Diversos.", vbInformation, Me.Caption
       '     Exit Sub
       ' End If
    'Else
    '    MsgBox "Não existe Impressora definida para Impressão de Relatórios Diversos.", vbInformation, Me.Caption
    '    Exit Sub
    'End If
   ' rstemp.Close
    
    frmMenu.MousePointer = 11
    On Error GoTo SaiImp
   ' SelecPrint.Action = 5
    Relatorios.ReportFileName = App.Path & "\relet6180.rpt"
    Relatorios.Action = 1
    frmMenu.MousePointer = 0

SaiImp:
    If Err.Number = 32755 Then
        Err.Clear
        Screen.MousePointer = 1
        Exit Sub
    End If
End Sub

Private Sub menu_relatorios_vendas_Click()
    menu_relatorios_vendas.Enabled = False
    frmMenu.Timer1.Enabled = False
    frmMenu.Timer2.Enabled = False
    frmRelVenda.Show 1

End Sub

Private Sub menu_relogin_Click()
  Dim rsMenu As Recordset

     '   If FRM_TASK.Visible = True Then
            FRM_TASK.cmd_fechar_Click
     '   End If
    
        flag_Relogin = True
        
        Call Conecta_Banco
        
        sql = " select * from Cad_Menus WHERE MENU_DS_SISTEMA = '" & NomeSistema & "'"
        Set Rstemp = New ADODB.Recordset
        Rstemp.Open sql, Cnn, 1, 2
        With Rstemp
            If .RecordCount = 0 Then
                rsMenu.Close
                MsgBox "Tabela de Menus esta vazia.", vbInformation, "Aviso"
                Exit Sub
            End If
            .MoveLast
            .MoveFirst
            Do While Not .EOF
                campotab = "MENU_DS_NOME_SISTEMA"
                For Each campo In .Fields
                    If UCase(campo.Name) = campotab Then
                        For Each campo2 In frmMenu.Controls
                            If UCase(campo.Value) = UCase(campo2.Name) Then
                                campo2.Enabled = False
'                                If UCase(campo.Value) = "MENU_CADASTRO_PRODUTOS" Then
'                                    frmMenu.Toolbar1.Buttons(1).Enabled = False
'                                    frmMenu.Toolbar1.Buttons(2).Enabled = False
'                                    frmMenu.Toolbar1.Buttons(3).Enabled = False
'                                    frmMenu.Toolbar1.Buttons(4).Enabled = False
'                                End If
                                frmMenu.Toolbar1.Buttons(3).Enabled = False
                                frmMenu.Toolbar1.Buttons(4).Enabled = False
                                frmMenu.Toolbar1.Buttons(5).Enabled = False
                                frmMenu.Toolbar1.Buttons(7).Enabled = False
                                frmMenu.Toolbar1.Buttons(8).Enabled = False
                                frmMenu.status.Panels(6).Text = "Ausente"
                                Exit For
                            End If
                        Next
                    End If
                Next
                .MoveNext
            Loop
            .Close
        End With
        
        frmAcesso.Show 1
End Sub

Private Sub menu_Retorno_Aliquotas_Click()

'*************************************************************
'*
'*  Obs.: Nessas funções de retorno de informações da
'*  impressora você tem a opção de escolher se o retorno
'*  virá na própria variável ou se será gravado no arquivo
'*  retorno.txt no diretório especificado no arquivo ini.
'*
'*  IMPORTANTE: Veja o tópico "Arquivo de Configuração
'*  BemaFi32.ini" na documentação da Dll para maiores
'*  informações
'*
'************************************************************

Dim Aliquotas As String

    If MsgBox("Executar Consulta das Aliquotas Cadastradas...?", vbYesNo + vbDefaultButton2 + vbQuestion, "Responda-me") = vbYes Then
        Screen.MousePointer = 11
        If (LocalRetorno = "1") Then 'Grava retorno em arquivo
            Aliquotas = Space(1)
        Else
            Aliquotas = Space(79)
        End If
        Retorno = Bematech_FI_RetornoAliquotas(Aliquotas)
        Call VerificaRetornoImpressora("Alíquotas Cadastradas: ", Aliquotas, "Informações da Impressora")
    End If
    Screen.MousePointer = 1
End Sub

Private Sub menu_sair_Click()
On Error GoTo TrataErro
    
    Call FechaRecordsets
    
    Kill App.Path & "\SisAdVen-Credito.cdc"
    
    If Cnn.State = 1 Then
        Cnn.Close
        Set Cnn = Nothing
    End If

    End
    
Exit Sub

TrataErro:
    Err.Clear
    End
End Sub

Private Sub menu_Sangria_Gaveta_Click()
frmSangria.Show 1
End Sub

Private Sub menu_Sobre_Click()
Frm_Sobre.Show 1
End Sub

Private Sub menu_Suprimento_Gaveta_Click()
frmSuprimento.Show 1
End Sub

Private Sub menu_Totalizadores_Parciais_Click()
'*************************************************************
'*
'*  Obs.: Nessas funções de retorno de informações da
'*  impressora você tem a opção de escolher se o retorno
'*  virá na própria variável ou se será gravado no arquivo
'*  retorno.txt no diretório especificado no arquivo ini.
'*
'*  IMPORTANTE: Veja o tópico "Arquivo de Configuração
'*  BemaFi32.ini" na documentação da Dll para maiores
'*  informações
'*
'************************************************************

    Dim Totalizadores As String
    
    If MsgBox("Executar Consulta dos Totalizadores Parciais...?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        Screen.MousePointer = 11
        If (LocalRetorno = "1") Then 'Grava retorno em arquivo
            Totalizadores = Space(1)
        Else
            Totalizadores = Space(445)
        End If
        Retorno = Bematech_FI_VerificaTotalizadoresParciais(Totalizadores)
        Call VerificaRetornoImpressora("Totalizadores Parciais: ", Totalizadores, "Informações da Impressora")
    End If
    Screen.MousePointer = 1
End Sub

Private Sub menu_Utilitarios_Backup_Click()
FRM_BACK_UP.Show 1
End Sub

Private Sub menu_utilitarios_calculadora_Click()
On Error GoTo Err
Shell "calc.exe", vbNormalFocus

Exit Sub

Err:
    MsgBox "Você não tem a calculadora instalada em seu computador.", vbExclamation, "Aviso"

End Sub

Private Sub menu_utilitarios_config_Click()
'    menu_utilitarios_config.Enabled = False

    frmMenu.Timer1.Enabled = False
    FrmConfiguracoes.Show 1

End Sub

Private Sub menu_utilitarios_cteacesso_Click()
    menu_utilitarios_cteacesso.Enabled = False
    frm_CdI_Usuario.Show 1

End Sub

Private Sub menu_utilitarios_edttexto_Click()
On Error GoTo Err

Shell "notepad.exe", vbNormalFocus

Exit Sub

Err:
    MsgBox "Você não tem o Bloco de Notas instalado em seu computador..", vbExclamation, "Aviso"

End Sub

Private Sub menu_utilitarios_Imp_Cancela_Cupom_Click()
If MsgBox("Deseja Realmente Cancelar Último Cupom Fiscal Emitido...?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
   Screen.MousePointer = 11
    Retorno = Bematech_FI_CancelaCupom()
    'Função que analisa o retorno da impressora
    Call VerificaRetornoImpressora("", "", "Cancelamento de Cupom Fiscal")
End If
Screen.MousePointer = 1
End Sub

Private Sub menu_utilitarios_Imp_Leitura_Memoria_Fiscal_Data_Click()
    frmLeituraMemoriaData.Caption = "Leitura da Memória Fiscal por Data"
    Funcao = 1
    frmLeituraMemoriaData.Show 1
    Funcao = 0
End Sub


Private Sub menu_utilitarios_Imp_Leitura_Memoria_Fiscal_Reducao_Click()
Funcao = 1
frmLeituraMemoriaReducao.Show 1
End Sub

Private Sub menu_utilitarios_Imp_Leitura_X_Click()

    If MsgBox("Deseja Realmente executar Leitura X...?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        Screen.MousePointer = 11
        Retorno = Bematech_FI_LeituraX()
        Call VerificaRetornoImpressora("", "", "Leitura X")
    End If
    
    Screen.MousePointer = 1

End Sub

Private Sub menu_utilitarios_Imp_Reducao_Z_Click()
Dim strData As String
Dim strHora As String

    strData = Format(Date, "dd/mm/yyyy")
    strshora = Format(Time, "hh:mm:ss")

    'Os parâmetros opcionais são para alterar
    'a hora da impressora em até + ou - 5 min.
    'para isso deve-se passar os parâmetros "Data" e "Hora"

    If MsgBox("Esta Operação irá executar a Redução Z." & vbNewLine & "Isto significa que nenhum Cupom Fiscal," & vbNewLine & "será mais Emitido no dia de Hoje...Deseja Realmente Gerar Redução Z ? ", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        Screen.MousePointer = 11
        Retorno = Bematech_FI_ReducaoZ("", "")
        Call VerificaRetornoImpressora("", "", "Redução Z")
    End If
    
    Screen.MousePointer = 1
End Sub

Private Sub menu_utilitarios_Informativo_Click()
If sysAcesso = 1 Then
    With FRM_TASK
        .CarregaInformativo
        .Show 1
    End With
End If

End Sub



Private Sub MENU_UTILITARIOS_REGISTRO_Click()
frmChamou = True
FrmRegistro.Show 1
End Sub

Private Sub menu_Utilitarios_Restaurar_Banco_Click()
FRM_RESTORE.Show 1
End Sub

Private Sub menu_utilitarios_trocasenha_Click()
    menu_utilitarios_trocasenha.Enabled = False
    frmTrocaSenha.Show 1

End Sub

Private Sub mnu_Backup_Click()
    frm_Backup.Show 1
End Sub

Private Sub menu_caixa_Click()
    frmCdPgto.Show 1
End Sub

Private Sub mnu_agenda_Click()
    Call menu_Agenda_Click
End Sub

Private Sub mnu_ETQ_Prood_Click()
FrmRel_Etiq_Prod.Show 1
End Sub

Private Sub mnu_help_Click()
On Error GoTo trata
    
Dim strfile As String
Dim objHelp As vbhelp
Set objHelp = New vbhelp
strfile = App.Path & "\SisAdvenHELP.chm"
Call objHelp.Show(strfile, "")
Set objHelp = Nothing

Exit Sub

trata:
    MsgBox "Arquivo de Help não encontrado...!", vbInformation, "Aviso"
Err.Clear
End Sub

Private Sub mnu_relatorios_estoque_produtos_fornecedor_grupo_Click()
    frm_Rel_Produtos_Grupos_Fornec.Show 1
End Sub

Private Sub Status_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   status.Panels(2).ToolTipText = " " & Format(Now, "LONG DATE") & " "
   status.Panels(3).ToolTipText = " " & Format(Now, "LONG TIME") & " "
End Sub

Private Sub Timer1_Timer()
    str = "SisAdven " & Versao_Software & " - Sistema de controle administrativo"
    Timer1.Interval = 150
    X = X + 1
    'Status.Panels(2).Text = Left(Str, x)
    Me.Caption = Left(str, X)
    If X = Len(str) Then
        X = 1
        Timer1.Interval = 3000
        Timer1.Enabled = False
        Timer2.Enabled = True
    End If
'    '*** Fabio Reinert - 10/2017 - Teste para ver se funciona com a tecla de atalho - Inicio
'    If GetAsyncKeyState(vbKeyF9) < 0 Then
'        frmAgenda.Show vbModal
'    End If
'    '*** Fabio Reinert - 10/2017 - Teste para ver se funciona com a tecla de atalho - Fim
End Sub

Private Sub Timer2_Timer()
   ' If X = 18 Then
   '     Timer2.Interval = 3000
        str = "www.novavia.com.br"
   ' Else
   '     str = "Sistema de controle administrativo"
   ' End If
    Timer2.Interval = 150
    X = X + 1
    'Status.Panels(2).Text = Left(Str, x)
    Me.Caption = Left(str, X)
    If X = Len(str) Then
       X = 1
        Timer2.Interval = 3000
        Timer2.Enabled = False
        Timer1.Enabled = True
        str = "SisAdven - Sistema de controle administrativo"
    End If
   
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        menu_relogin_Click
    Case 3
        menu_cadastro_cliente_Click
    Case 4
        menu_cadastro_fornecedor_Click
    Case 5
        menu_cadastro_produtos_Click
    Case 7
        menu_movimentacao_entrada_Click
    Case 8
        menu_movimentacao_saida_Click
    Case 10
        mnu_help_Click
'**** Fabio Reinert - 07/2017 - Alterado para comtemplar os botoes da agenda também - Inicio
    Case 12
        menu_Agenda_Click
    Case 14
        menu_sair_Click
'**** Fabio Reinert - 07/2017 - Alterado para comtemplar os botoes da agenda também - Fim
End Select

End Sub
'***
'
Private Sub sVerificaAtualizacoes()
'*
'****************************************************************
'**** Fabio Reinert - Criação da sub sVerificaAtualizacoes   ****
'****************************************************************
'*
'*****************************************************************************************
'*** Fabio Reinert - Inclusão do campo TELEFONE CELULAR na tabela clientes - Inicio   ****
'*****************************************************************************************
'
    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS " & _
          " WHERE RDB$RELATION_NAME = 'CLIENTE' AND  rdb$field_name='CELULAR'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "ALTER TABLE CLIENTE ADD CELULAR VARCHAR(10)"
        Cnn.Execute sql
        Cnn.CommitTrans
    End If
    Rstemp6.Close
    Set Rstemp6 = Nothing
'*****************************************************************************************
'*** Fabio Reinert - Inclusão do campo TELEFONE CELULAR na tabela clientes - Fim
'*****************************************************************************************
'
'***********************************************************************************************
'*** Fabio Reinert - Inclusão dos campos TIPO CLIENTE (PENDENTE) e                          ****
'***                  PAGTODIA p/ pendentes, na tabela clientes - Inicio                    ****
'***********************************************************************************************
'
    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS " & _
          " WHERE RDB$RELATION_NAME = 'CLIENTE' AND  rdb$field_name='PRAZOPEND'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "ALTER TABLE CLIENTE ADD PRAZOPEND VARCHAR(3)"
        Cnn.Execute sql
        Cnn.CommitTrans
    End If
    Rstemp6.Close
    Set Rstemp6 = Nothing
'
    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS " & _
          " WHERE RDB$RELATION_NAME = 'CLIENTE' AND  rdb$field_name='PAGTODIA'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "ALTER TABLE CLIENTE ADD PAGTODIA VARCHAR(2)"
        Cnn.Execute sql
        Cnn.CommitTrans
    End If
    Rstemp6.Close
    Set Rstemp6 = Nothing
'*******************************************************************************************
'*** Fabio Reinert - Inclusão dos campos TIPO CLIENTE (PENDENTE) e                      ****
'***                  PAGTODIA p/ pendentes, na tabela clientes - FIM                   ****
'*******************************************************************************************
'
    'TRATA TABELA SAIDAS_PRODUTO CAMPO PDV
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'SAIDAS_PRODUTO' AND rdb$field_name = 'PDV' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "ALTER TABLE SAIDAS_PRODUTO ADD CAIXA varchar(30), ADD PDV varchar(30)"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        ' Indices
        sql = "CREATE INDEX ID_X_SEQUENCIA ON ITENS_SAIDA (SEQUENCIA)"
        Cnn.Execute sql
        Cnn.CommitTrans
    End If
    
    Rstemp6.Close
    Set Rstemp6 = Nothing
'
    'TRATA TABELA SAIDAS_PRODUTO CAMPO PDV
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'COD_BAR'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "CREATE TABLE COD_BAR (COD_BAR varchar(10))"
        Cnn.Execute sql
        Cnn.CommitTrans

        sql = "UPDATE COD_BAR SET COD_BAR = '987654'"
        Cnn.Execute sql
        Cnn.CommitTrans
    End If

    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    'TRATA TABELA ETIQ_GONDOLA
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'ETIQ_GONDOLA' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "CREATE TABLE ETIQ_GONDOLA (CODIGO DOUBLE PRECISION,PRIMARY KEY (CODIGO))"
        Cnn.Execute sql
        Cnn.CommitTrans
    End If
    
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    'CFE - SAT
    'sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'CFE' AND rdb$field_name = 'NRO_CAIXA' "
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'CFE'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "CREATE TABLE CFE ("
        sql = sql & " ID  INTEGER NOT NULL,"
        sql = sql & " NRO_PEDIDO_CFE        BLOB SUB_TYPE 1 SEGMENT SIZE 3000 CHARACTER SET WIN1252 COLLATE WIN1252,"
        sql = sql & " EMISSAO_CFE           DATE,"
        sql = sql & " NRO_CFE               INTEGER,"
        sql = sql & " SESSAO_CFE            VARCHAR(10),"
        sql = sql & " CHAVE_ACESSO_CFE      VARCHAR(200),"
        sql = sql & " STATUS_RETORNO_CFE    VARCHAR(200),"
        sql = sql & " XML_CFE               BLOB SUB_TYPE 1 SEGMENT SIZE 16000 CHARACTER SET WIN1252,"
        sql = sql & " CAMINHO_XML_CFE       BLOB SUB_TYPE 1 SEGMENT SIZE 3000 CHARACTER SET WIN1252 COLLATE WIN1252,"
        sql = sql & " MODELO_CFE            VARCHAR(4),"
        sql = sql & " SERIE_SAT_CFE         VARCHAR(30),"
        sql = sql & " CAIXA_CFE             VARCHAR(50),"
        sql = sql & " PDV_CFE               VARCHAR(50),"
        sql = sql & " NRO_CAIXA_CFE         VARCHAR(15),"
        sql = sql & " CPF_CLIENTE_CFE       VARCHAR(14),"
        sql = sql & " VALOR_TRIBUTOS_CFE    DOUBLE PRECISION,"
        sql = sql & " BASE_ICMS_CFE         DOUBLE PRECISION,"
        sql = sql & " VALOR_ICMS_CFE        DOUBLE PRECISION,"
        sql = sql & " TOTAL_BRUTO_CFE       NUMERIC(12,2),"
        sql = sql & " TOTAL_DESCONTO_CFE    NUMERIC(12,2),"
        sql = sql & " TOTAL_ACRESCIMO_CFE   NUMERIC(12,2),"
        sql = sql & " TOTAL_CFE             NUMERIC(12,2),"
        sql = sql & " VALOR_PAGO_CFE        NUMERIC(12,2),"
        sql = sql & " TROCO_CFE             NUMERIC(12,2),"
        sql = sql & " CANCELADO             Char (1))"
        
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "ALTER TABLE CFE ADD CONSTRAINT PK_CFE PRIMARY KEY (ID)"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_CFE_ID1 "
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "SET GENERATOR GEN_CFE_ID1 TO 0"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        'cria TRIGGER PARA AUTONUMERADOR
        sql = " CREATE TRIGGER CFE_BI FOR CFE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_CFE_ID1, 1); END  "
        Cnn.Execute sql
        Cnn.CommitTrans
        
    End If
       
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    'IBTP
    Call Verifica_IBPT
    
    
    'ALTERAÇÃO CADASTRO DE PRODUTOS EM 08/09/2015
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'PRODUTO' AND rdb$field_name = 'CFOP' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
            sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'PRODUTO' AND rdb$field_name = 'INATIVO' "
            Set Rstemp5 = New ADODB.Recordset
            Rstemp5.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
            If Rstemp5(0) = 0 Then
                sql = "ALTER TABLE PRODUTO ADD CFOP varchar(4), ADD INATIVO char(1)"
                Cnn.Execute sql
                Cnn.CommitTrans
            Else
                sql = "DROP VIEW VIEW_ESTOQUE_NEG "
                Cnn.Execute sql
                Cnn.CommitTrans
                
                sql = "DROP VIEW VIEW_SO_ESTOQUE_NEG "
                Cnn.Execute sql
                Cnn.CommitTrans
                
                sql = "DROP VIEW VIEW_PRODUTO_DESCRIC "
                Cnn.Execute sql
                Cnn.CommitTrans
               
                sql = "ALTER TABLE PRODUTO DROP INATIVO "
                Cnn.Execute sql
                Cnn.CommitTrans
                
                'sql = "ALTER TABLE PRODUTO DROP INATIVO "
                sql = "ALTER TABLE PRODUTO ADD CFOP varchar(4), ADD INATIVO char(1)"
                Cnn.Execute sql
                Cnn.CommitTrans
                
                sql = " CREATE VIEW VIEW_ESTOQUE_NEG(CODIGO_INTERNO,PRODUTO,SALDO_EM_ESTOQUE,MARCA,GRUPO,INATIVO) AS "
                sql = sql & "SELECT PRODUTO.CODIGO_INTERNO,PRODUTO.DESCRICAO AS PRODUTO,ESTOQUE.SALDO_EM_ESTOQUE, MARCAS.DESCRICAO AS MARCA,GRUPO.DESCRICAO AS GRUPO,PRODUTO.INATIVO  "
                sql = sql & " FROM ((PRODUTO INNER JOIN ESTOQUE ON PRODUTO.CODIGO = ESTOQUE.CODIGO_PRODUTO) INNER JOIN MARCAS ON PRODUTO.MARCA = MARCAS.CODIGO) "
                sql = sql & " INNER JOIN GRUPO ON PRODUTO.GRUPO = GRUPO.CODIGO WHERE ESTOQUE.SALDO_EM_ESTOQUE <=0 "
                Cnn.Execute sql
                Cnn.CommitTrans
                                
                sql = "CREATE VIEW VIEW_SO_ESTOQUE_NEG(CODIGO_INTERNO,NOME_PRODUTO,ULTIMA_VENDA,ULTIMA_COMPRA,SALDO_EM_ESTOQUE,QTD_MINIMA,COMPRAR,GRUPO,INATIVO) AS "
                sql = sql & " SELECT B.CODIGO_INTERNO, B.DESCRICAO AS NOME_PRODUTO, B.ULTIMA_VENDA, B.ULTIMA_COMPRA, A.SALDO_EM_ESTOQUE, B.QTD_MINIMA,"
                sql = sql & " (B.QTD_MINIMA - A.SALDO_EM_ESTOQUE) AS COMPRAR, G.DESCRICAO AS GRUPO,B.INATIVO FROM ESTOQUE A, PRODUTO B,  GRUPO G WHERE A.CODIGO_PRODUTO = B.CODIGO"
                sql = sql & " AND G.CODIGO = B.GRUPO  AND (A.SALDO_EM_ESTOQUE < B.QTD_MINIMA)  and (A.SALDO_EM_ESTOQUE >= 0) ORDER BY NOME_PRODUTO"
                Cnn.Execute sql
                Cnn.CommitTrans
                
                sql = "CREATE VIEW VIEW_PRODUTO_DESCRIC(CODIGO,CODIGO_INTERNO,DESCRICAO,PRECO,UNIDADE,SALDO_EM_ESTOQUE,ULTIMA_VENDA,ULTIMA_COMPRA,DATA_CAD_ALT,"
                sql = sql & " PRECO_ATACADO,"
                sql = sql & " PRECO_MINIMO_ATACADO,"
                sql = sql & " INATIVO) AS"
                sql = sql & " Select A.CODIGO,A.CODIGO_INTERNO,A.DESCRICAO,A.PRECO,A.UNIDADE, B.SALDO_EM_ESTOQUE,A.ULTIMA_VENDA,A.ULTIMA_COMPRA,A.DATA_CAD_ALT, A.PRECO_ATACADO, A.PRECO_MINIMO_ATACADO, A.INATIVO FROM PRODUTO A, ESTOQUE B WHERE A.Codigo = B.CODIGO_PRODUTO ORDER BY A.Descricao ASC"
                Cnn.Execute sql
                Cnn.CommitTrans
            End If
            Rstemp5.Close
                
            sql = "UPDATE PRODUTO SET CFOP = '5405'"
            Cnn.Execute sql
            Cnn.CommitTrans

    End If
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    
    
    ' VERIFICA SE O CAMPO EXISTE SE NAO CRIA
    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'ARQ_ESTOQUE' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
         sql = "CREATE TABLE ARQ_ESTOQUE (ID_ESTOQUE INTEGER NOT NULL, ID_PRODUTO DOUBLE PRECISION NOT NULL, ID_FORNECEDOR DOUBLE PRECISION, "
         sql = sql & " DOCUMENTO DOUBLE PRECISION, DATA TIMESTAMP, SALDO_ANTERIOR DOUBLE PRECISION, ENTRADA DOUBLE PRECISION,SAIDA DOUBLE PRECISION,"
         sql = sql & " SALDO_AJUSTADO DOUBLE PRECISION,SALDO_ATUAL DOUBLE PRECISION,SALDO_BONIFIC DOUBLE PRECISION, PRECO_CUSTO DOUBLE PRECISION, PRECO_VENDA DOUBLE PRECISION,"
         sql = sql & " ENTRADA_BONIF DOUBLE PRECISION,SAIDA_BONIFIC DOUBLE PRECISION, TRANSFERENCIA DOUBLE PRECISION, QUEBRA DOUBLE PRECISION, "
         sql = sql & " JUSTIFICATIVA VARCHAR(100)CHARACTER SET WIN1252,PRIMARY KEY (ID_ESTOQUE))"
         Cnn.Execute sql
         Cnn.CommitTrans
         
         'sql = "DROP TABLE ARQ_ESTOQUE"
         'Cnn.Execute sql
         'Cnn.CommitTrans
        
         ' Indices
         sql = "CREATE INDEX X_ID_ ON ARQ_ESTOQUE (ID_ESTOQUE)"
         Cnn.Execute sql
         Cnn.CommitTrans
         
         'cria GENERATOR
         sql = "CREATE GENERATOR GEN_SEQ_ARQ_ESTOQUE "
         Cnn.Execute sql
         Cnn.CommitTrans
         
         sql = "SET GENERATOR GEN_SEQ_ARQ_ESTOQUE TO 0"
         Cnn.Execute sql
         Cnn.CommitTrans
         
          'cria TRIGGER PARA AUTONUMERADOR
         sql = " CREATE TRIGGER ARQ_ESTOQUE FOR ARQ_ESTOQUE ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
         sql = sql & "if (NEW.ID_ESTOQUE is NULL) then NEW.ID_ESTOQUE = GEN_ID(GEN_SEQ_ARQ_ESTOQUE, 1); END  "
         Cnn.Execute sql
         Cnn.CommitTrans
    
    End If
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'PRODUTO_FORNECEDOR' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "CREATE TABLE PRODUTO_FORNECEDOR ("
        sql = sql & "    COD_FORNECEDOR  VARCHAR(50) NOT NULL,"
        sql = sql & "    COD_ENTRADA     VARCHAR(30) NOT NULL,"
        sql = sql & "    COD_INTERNO     VARCHAR(30) NOT NULL,"
        sql = sql & "    CODIGO          DOUBLE PRECISION NOT NULL)"
        Cnn.Execute sql
        Cnn.CommitTrans
    End If
    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    'select entrada e saida
    sql = "Select Sum(Saida) as Total_Saida, Sum(Entrada) as Total_Entrada FROM "
    sql = sql & "(select Cast(0 as numeric(15,3)) as Saida, qtde as Entrada FROM itens_entrada "
    sql = sql & " Union All"
    sql = sql & " select sum(qtde) as Saida, Cast(0 as numeric(15,3)) as Entrada from itens_saida) as TMP"
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) > 0 Then
'
'    End If
'    Rstemp6.Close
'    Set Rstemp6 = Nothing

    'ou mais completo
    'sql = "Select Codigo, Tipo, Data, Saida, Entrada, MotSai, MotEntrada, pagSaida From "
    'sql = sql & " (select codEntrada as Codigo  , Cast('E'as char(1)) as Tipo , datEntrada as Data, Cast(0 as numeric(15,3)) as Saida,"
    'sql = sql & " vlrEntrada as Entrada, Cast(0 as integer) as MotSai, motEntrada as MotEntrada, Cast(0 as integer) as pagSaida From tbEntrada"
    'sql = sql & " Union All "
    'sql = sql & " select codSaida as Codigo, Cast('E'as char(1)) as Tipo, datSaida as Data, vlrSaida as Saida, Cast(0 as numeric(15,3)) as Entrada,"
    'sql = sql & " motSaida as MotSai, Cast(0 as integer) as MotEntrada, pagSaida as pagSaida from tbSaida) as TMP"
    
    sql = ""
    
    sql = "SELECT * FROM CAD_MENUS WHERE MENU_CD_CODI = 74 "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, 1, 2
    If Rstemp6.RecordCount = 0 Then
        sql = "INSERT INTO CAD_MENUS VALUES("
        sql = sql & "'SisAdven',"
        sql = sql & "74,"
        Menu = UCase("ALTERA_PRECOS_PRODUTOS")
        sql = sql & "'" & Menu & "',"
        sql = sql & "'Cadastro - Produtos - Alterar Preços')"
        Cnn.Execute sql
    End If

    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    '*** Fabio Reinert (Alemão) - 07/2017 - Verificação de arquivo CEST.TXT e se existir cria nova TAB_CEST - Inicio
    '***
    'CEST     --->  Primeiro verificar se existe o arquivo CEST.TXT
    If Len(Dir(App.Path & "\CEST.TXT")) > 0 Then
        strSql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'TAB_CEST' "
        Set Rstemp6 = New ADODB.Recordset
        Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Rstemp6(0) = 0 Then
            strSql = "CREATE TABLE TAB_CEST (CEST VARCHAR(7) NOT NULL,NCM VARCHAR(8),DESCRICAO VARCHAR(512) ) "
            Cnn.Execute strSql
            Cnn.CommitTrans
        Else
            sql = "DELETE FROM TAB_CEST "
            Cnn.Execute strSql
            Cnn.CommitTrans
        End If
        Call sPopula_Tab_Cest   'Sub que popula a tab_cest com o conteúdo do arquivo texto CEST.TXT
        Kill App.Path & "\CEST.TXT"
          '***
    Else  '*** Perguntar se não tiver o arquivo e a tabela não existir o que fazer?
          '***
        strSql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'TAB_CEST' "
        Set Rstemp6 = New ADODB.Recordset
        Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
        If Rstemp6(0) = 0 Then
            strSql = "CREATE TABLE TAB_CEST (CEST VARCHAR(7) NOT NULL,NCM VARCHAR(8),DESCRICAO VARCHAR(512) ) "
            Cnn.Execute strSql
            Cnn.CommitTrans
        End If
    End If
    If Rstemp6.State = adStateOpen Then
        Rstemp6.Close
    End If
    Set Rstemp6 = Nothing
    '*** Fabio Reinert (Alemão) - 07/2017 - Verificação de arquivo CEST.TXT e se existir cria nova TAB_CEST - Inicio
    '***
    
    'CEST
'    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'TAB_CEST' "
'    Set Rstemp6 = New ADODB.Recordset
'    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
'    If Rstemp6(0) = 0 Then
'        sql = "CREATE TABLE TAB_CEST (CEST VARCHAR(7) NOT NULL,NCM VARCHAR(8),DESCRICAO VARCHAR(512) ) "
'        Cnn.Execute sql
'        Cnn.CommitTrans
'    End If
'    Rstemp6.Close
'    Set Rstemp6 = Nothing
'
'    Dim LineofText As String
'
'    Dim Linha As String
'    Dim Separa() As String
'
'    X = 0
'    i = 0
'    sql = ""
    
    
'Open App.Path & "\CEST.txt" For Input As #2
'    Do While Not EOF(1)
'        'pega a linha do TXT
'        Line Input #2, Linha
'        'separa o texto antes de gravar
'        Separa() = Split(Linha, ";")
'        'grava o registro na tabela
'        'TBCliente1(0) = Separa(0)
'        Texto = Separa(0)
'
'        Line Input #2, Linha
'        texto1 = Separa(1)
'        texto2 = Separa(2)
'        texto3 = Separa(3)
'        texto4 = Separa(4)
'        texto5 = Separa(5)
'    Loop
'Close #2
    
    
'
'    Dim Cnn_a_Importar As New ADODB.Connection
'    Set Cnn_a_Importar = New ADODB.Connection
'
'    With Cnn_a_Importar
'        .CursorLocation = adUseClient
'            'SqlServer 2000
'            .Open "Provider=SQLOLEDB.1;Password=123;Persist Security Info=True;User ID=BILL;Initial Catalog=ArqdadosPDV;Data Source=servidor"
'
'           'firebird
'             '.Open "Provider=IBOLE.Provider.v4;Persist Security Info=False;Data Source=GUSTAVO:c:\Sistema SisAdven\ARQDADOS.GDB"
'            '.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=bill2;Initial Catalog=SimplesFIL;Data Source=SERVIDOR\MSSQLSERVER_R2"
'
'            'firebird
'            '.Open "Provider=IBOLE.Provider.v4;Persist Security Info=False;Data Source=servidor:c:\Sistema SisAdven\gilda.FDB"
'
'    End With
'
'
'    sql = "SELECT * FROM TAB_CEST ORDER BY NCM "
'    Set Rstemp = New ADODB.Recordset
'    Rstemp.Open sql, Cnn_a_Importar, 1, 2
'    If Rstemp.RecordCount > 0 Then
'        Rstemp.MoveLast
'        Rstemp.MoveFirst
'        While Not Rstemp.EOF
'            sql = "INSERT INTO TAB_CEST VALUES ("
'            sql = sql & "'" & Rstemp(0) & "',"
'            sql = sql & "'" & Rstemp(1) & "',"
'            sql = sql & "'" & Rstemp(2) & "')"
'            Cnn.Execute sql
'            Cnn.CommitTrans
'            Rstemp.MoveNext
'        Wend
'
'    End If
'    Rstemp.Close
    
    
'    Cnn.Execute "delete from tab_cest"
'    Cnn.CommitTrans
'
'    'importa cest
'    Dim Separa() As String
'
'    LineofText = ""
'    sql = ""
'    Open App.Path & "\TAB_EST_FIREBIRD.sql" For Input As #1
'    'Open App.Path & "\CEST_old.txt" For Input As #1
'    Do While Not EOF(1)
'    '            sql = ""
'    '            For X = 1 To 2 'necessário por conta da quebra de linha
'    '                Line Input #1, LineofText
'    '                'Debug.Print LineofText
'    '                If X = 1 Then
'    '                    sql = LineofText
'    '                Else
'    '                    sql = sql & LineofText
'    '                End If
'    '            Next X
'    '            'Debug.Print sql
'    '            Cnn.Execute sql
'    '            Cnn.CommitTrans
'        Line Input #1, LineofText
'        If Len(LineofText) > 0 Then
'            Cnn.Execute LineofText
'            Cnn.CommitTrans
'        Else
'            aaa = LineofText
'        End If
'    Loop
'    Close #1
'
'
'    MsgBox "Tabela CEST criada com sucesso...!", vbInformation, "Aviso"

    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'PRODUTO' AND  rdb$field_name='CEST'"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "ALTER TABLE PRODUTO ADD CEST VARCHAR(7)"
        Cnn.Execute sql
        Cnn.CommitTrans
    End If

    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    sql = "select COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'REL_RANKING_VENDAS_VENDEDOR '"
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "CREATE TABLE REL_RANKING_VENDAS_VENDEDOR (COD_VEND double PRECISION, RAZAO_SOCIAL VARCHAR(60), QTDE DOUBLE PRECISION, TOTAL double PRECISION)"
        Cnn.Execute sql
    End If

    Rstemp6.Close
    Set Rstemp6 = Nothing
    
    'CURVA ABC
    sql = "SELECT  P.CODIGO_INTERNO, P.DESCRICAO, SUM(I.QTDE) AS"
    sql = sql & " SUBTOTAL, SUM(I.QTDE * I.VALOR_UNITARIO) / SUM(I.VALOR_TOTAL) * 100 AS"
    sql = sql & " CURVA_ABC, V.Data_NF,TOTAL_SAIDA FROM ITENS_SAIDA I INNER JOIN SAIDAS_PRODUTO V ON I.SEQUENCIA = V.SEQUENCIA"
    sql = sql & " INNER JOIN PRODUTO P ON I.CODIGO_PRODUTO = P.CODIGO"
    'sql = sql & " WHERE  C.DATA_NF BETWEEN  '" & Format(mskDe.Text, "MM/DD/yyyy") & "'"
    sql = sql & " WHERE V.Data_NF BETWEEN '10/11/2016' AND '10/11/2016'"
    sql = sql & " GROUP BY V.Data_NF, P.CODIGO_INTERNO, P.DESCRICAO, V.TOTAL_SAIDA"
    sql = sql & " ORDER BY SUM(I.QTDE * I.VALOR_UNITARIO) / SUM(I.VALOR_total) * 100 DESC"
    'Set Rstemp = New ADODB.Recordset
    'Rstemp.Open sql, Cnn, 1, 2
    'If Rstemp.RecordCount <> 0 Then
    'End If
    
    'CURVA ABC CLIENTES
    sql = "CREATE VIEW CURVA_ABC_CLIENTES(DATA,CLIENTE,CONTATO,CIDADE,UF,FONE,ULT_VENDA,TOTAL)"
    sql = sql & " AS select  v.DATA_NF as data, c.codigo||' - '||c.RAZAO_SOCIAL as nome_cliente, C.CONTATO,            "
    sql = sql & "c.CIDADE_END_PRINCIPAL AS CIDADE,"
    sql = sql & "c.UF_END_PRINCIPAL AS UF,"
    sql = sql & "c.FONE1 AS FONE,"
    sql = sql & "max(v.DATA_NF) as ULT_VENDA,"
    sql = sql & "sum(v.TOTAL_SAIDA) as TOTAL"
    sql = sql & "from saidas_produto v"
    sql = sql & "inner join cliente c on (c.codigo = v.CODIGO_CLIENTE)"
    sql = sql & "group by 1,2,3,4,5,6"
    'Set Rstemp = New ADODB.Recordset
    'Rstemp.Open sql, Cnn, 1, 2
    'If Rstemp.RecordCount <> 0 Then
    'End If

    
    'TRATA TABELA SAIDAS_PRODUTO CAMPO PDV
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'TRANSPORTADORA' AND rdb$field_name = 'DDD_TELEFONE' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        sql = "DROP TABLE TRANSPORTADORA"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "CREATE TABLE TRANSPORTADORA (CODIGO DOUBLE PRECISION NOT NULL,RAZAO_SOCIAL varchar (60),  CEP_PRINCIPAL varchar (9),"
        sql = sql & "ENDERECO_PRINCIPAL varchar (60),NRO_END_PRINCIPAL varchar (10),COMPL_END_PRINCIPAL varchar (30),BAIRRO_END_PRINCIPAL varchar (45),"
        sql = sql & "CIDADE_END_PRINCIPAL varchar (60),UF_END_PRINCIPAL varchar (2),CNPJ varchar (18),INSC_ESTADUAL varchar (30),SITE varchar (40),"
        sql = sql & "EMAIL varchar (60),NOME varchar (40),DEPTO varchar (30),EMAIL_CONTATO varchar (60),DDD_TELEFONE DOUBLE PRECISION,"
        sql = sql & "TELEFONE varchar (30), DDD_FAX DOUBLE PRECISION,   FAX varchar (30)) "
        Cnn.Execute sql
        Cnn.CommitTrans
        
        'Indices
        '''sql = "CREATE INDEX ID_X_SEQUENCIA ON ITENS_SAIDA (SEQUENCIA)"
        sql = "ALTER TABLE TRANSPORTADORA ADD CONSTRAINT I101 PRIMARY KEY (CODIGO);"
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "CREATE INDEX PK_RAZAO_SOCIAL_ ON TRANSPORTADORA (RAZAO_SOCIAL);"
        Cnn.Execute sql
        Cnn.CommitTrans
                
        sql = "INSERT INTO TRANSPORTADORA (CODIGO,RAZAO_SOCIAL) VALUES (1,'NOSSO CARRO')"
        Cnn.Execute sql
        Cnn.CommitTrans
    End If
    
    Rstemp6.Close
    Set Rstemp6 = Nothing


'*** Fabio Reinert ( Alemão ) - 08/2017 - Alteração da tabela FORMAS - Novo conteudo das formas de pagto. - Inicio
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'FORMAS' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
            
        sql = "CREATE TABLE FORMAS (CODIGO DOUBLE PRECISION NOT NULL ,DESCRICAO varchar (60) )"
        Cnn.Execute sql
        Cnn.CommitTrans
         
        'cria GENERATOR
        sql = "CREATE GENERATOR GEN_SEQ_FORMAS "
        Cnn.Execute sql
        Cnn.CommitTrans
        
        sql = "SET GENERATOR GEN_SEQ_FORMAS TO 0"
        Cnn.Execute sql
        Cnn.CommitTrans
         
        'cria TRIGGER PARA AUTONUMERADOR
        sql = " CREATE TRIGGER FORMAS_BI FOR FORMAS ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
        sql = sql & "if (NEW.CODIGO is NULL) then NEW.CODIGO = GEN_ID(GEN_SEQ_FORMAS, 1); END  "
        Cnn.Execute sql
        Cnn.CommitTrans
    End If
        
    sql = "DELETE FROM FORMAS"
    Cnn.Execute sql
    Cnn.CommitTrans
            
    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Dinheiro')"
    Cnn.Execute sql
    Cnn.CommitTrans
    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Cartão de Débito')"
    Cnn.Execute sql
    Cnn.CommitTrans
    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Cartão de Crédito')"
    Cnn.Execute sql
    Cnn.CommitTrans
    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Pendente')"
    Cnn.Execute sql
    Cnn.CommitTrans
    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Cheques')"
    Cnn.Execute sql
    Cnn.CommitTrans
    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('DOC/TED')"
    Cnn.Execute sql
    Cnn.CommitTrans
    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Boleto')"
    Cnn.Execute sql
    Cnn.CommitTrans
    sql = "INSERT INTO FORMAS (DESCRICAO) VALUES ('Cartão da Loja')"
    Cnn.Execute sql
    Cnn.CommitTrans
    
    Rstemp6.Close
    Set Rstemp6 = Nothing

'*** Fabio Reinert ( Alemão ) - 08/2017 - Alteração da tabela FORMAS - Novo conteudo das formas de pagto. - Fim
    
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
                                            " CODCLIENTE integer NOT NULL," & _
                                            " TIPO_CARTAO CHAR(1), " & _
                                            " COD_CARTAO integer NOT NULL, " & _
                                            " DT_EMISSAO DATE, " & _
                                            " VALOR_ORIG NUMERIC(12,2) , " & _
                                            " VALOR_CORRIG NUMERIC(12,2) , " & _
                                            " DT_VENCTO DATE," & _
                                            " DT_BAIXA DATE," & _
                                            " OPERADOR VARCHAR(10), " & _
                                            " DT_ATUALIZA DATE) "

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
                                            " DT_ATUALIZA DATE )"

        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
'
'******************************************************************************
'**** Fabio Reinert (Alemão) - 08/2017 - Alteração nas tabelas já existentes  *
'****                           Inclusão do campo DT_EMISSAO   - INICIO       *
'******************************************************************************
'
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'CHEQUES' AND rdb$field_name = 'DT_EMISSAO' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE CHEQUES ADD dt_emissao date"
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
    
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'CHEQUES' AND rdb$field_name = 'DT_BAIXA' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE CHEQUES ADD dt_baixa date "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
    
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'CHEQUES' AND rdb$field_name = 'OPERADOR' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE CHEQUES add operador varchar(10) "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
    
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'CHEQUES' AND rdb$field_name = 'DT_ATUALIZA' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE CHEQUES ADD dt_atualiza date "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
        
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'CTAS_PENDENTE' AND rdb$field_name = 'VALOR_BAIXADO' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE CTAS_PENDENTE ADD valor_baixado numeric(12,2) "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
    
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'CTAS_PENDENTE' AND rdb$field_name = 'OPERADOR' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE CTAS_PENDENTE ADD operador varchar(10) "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
    
    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'CTAS_PENDENTE' AND rdb$field_name = 'DT_ATUALIZA' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE CTAS_PENDENTE ADD dt_atualiza date "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
'
'**********************************************************************************
'**** Fabio Reinert (Alemão) - 08/2017 - Alteração n tabela CHEQUES já existente  *
'****                           Inclusão do campo DT_EMISSAO         - FIM        *
'**********************************************************************************
'*
'*** Fabio Reinert ( Alemão ) - 09/2017 - Criação de campos novos - Inicio
'*
'*** Criação do campo Valor da mensalidade no cadastro de cartoes

    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'TAB_CARTOES' AND rdb$field_name = 'VLR_MENSA' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE TAB_CARTOES ADD vlr_mensa numeric(12,2) "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
    
'*** Criação do campo dia do vencimento da mensalidade no cadastro de cartoes

    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'TAB_CARTOES' AND rdb$field_name = 'DIA_VENC' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE TAB_CARTOES ADD dia_venc integer "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
    
'*** Criação do campo data do vencimento na tabela de contas pendentes (FIADO)

    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'CTAS_PENDENTE' AND rdb$field_name = 'DTA_VENCTO' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE CTAS_PENDENTE ADD DTA_VENCTO DATE "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If

'*** Criação do campo Operador na tabela de contas pendentes (FIADO)

    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'CTAS_PENDENTE' AND rdb$field_name = 'OPERADOR' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE CTAS_PENDENTE ADD OPERADOR VARCHAR(10) "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If

'*** Criação do campo dta_atualiza na tabela de contas pendentes (FIADO)

    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'CTAS_PENDENTE' AND rdb$field_name = 'DT_ATUALIZA' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE CTAS_PENDENTE ADD dt_atualiza date "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If

'*** Criação do campo pedido na tabela de boletos (BOLETOS_PG) - inicialmente ficarão sem data e valor baixado

    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'BOLETOS_PG' AND rdb$field_name = 'PEDIDO' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE BOLETOS_PG ADD PEDIDO INTEGER "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If

'*** Criação do campo Operador na tabela de boletos (BOLETOS_PG)

    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'BOLETOS_PG' AND rdb$field_name = 'OPERADOR' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE BOLETOS_PG ADD OPERADOR VARCHAR(10) "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If

'*** Criação do campo dta_atualiza na tabela de boletos (BOLETOS_PG)

    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'BOLETOS_PG' AND rdb$field_name = 'DT_ATUALIZA' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE BOLETOS_PG ADD dt_atualiza date "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
'*** Criação de novos campos na tabela RECE_PAGA
'*** Criação do campo Operador na tabela rece_paga

    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'RECE_PAGA' AND rdb$field_name = 'OPERADOR' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE RECE_PAGA ADD OPERADOR VARCHAR(10) "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If

'*** Criação do campo dta_atualiza na tabela de rece_paga

    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'RECE_PAGA' AND rdb$field_name = 'DT_ATUALIZA' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE RECE_PAGA ADD dt_atualiza date "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If
    

'*** Criação do campo forma de pagamento na tabela de atendimentos da Agenda Pet Shop

    sql = "SELECT COUNT(*) rdb$field_name from RDB$RELATION_FIELDS WHERE "
    sql = sql & " RDB$RELATION_NAME = 'TAB_ATENDIMENTOS' AND rdb$field_name = 'FORMAPAGTO' "
    Set Rstemp6 = New ADODB.Recordset
    Rstemp6.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
    If Rstemp6(0) = 0 Then
        strSql = "ALTER TABLE TAB_ATENDIMENTOS ADD FORMAPAGTO varchar(50) "
        Cnn.Execute strSql
        Cnn.CommitTrans
    End If


'*
'*** Fabio Reinert ( Alemão ) - 09/2017 - Criação de campos novos - Fim
'*
    
    i = ReadIniFile(App.Path & "\SisAdven.ini", "Desc_prod_cupom", "Chk", "0")
    If i = 0 Then
        flag_Desc_prod_cupom = False
    Else
        flag_Desc_prod_cupom = True
    End If

End Sub
