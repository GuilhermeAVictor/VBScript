Sub AutoTester_CustomConfig()
    'Equipe Quality Assurance - Célula 5
    '***********************************************************************
    '*  Sub‑rotina : AutoTester_CustomConfig
    '*  Finalidade : Validar conectividade com o banco Quality Assurance e iniciar testes
    '***********************************************************************
    
    Dim Resposta
    Resposta = MsgBox( _
        "ATENÇÃO: Este processo pode ser lento." & vbCrLf & vbCrLf & _
        "O teste automático irá realizar uma varredura completa nas telas do domínio atual, verificando propriedades e objetos conforme as regras do AutoTester." & vbCrLf & vbCrLf & _
        "Deseja realmente iniciar o teste automático agora?", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Iniciar Teste Automático do Domínio")

    If Resposta = vbNo Then Exit Sub

    '======================================================================
    ' 1) Testa conectividade VPN via ping nos dois IPs definidos
    '======================================================================
    Dim ConexaoVPN
    ConexaoVPN = False

    If TestarPing("192.168.20.12") Or TestarPing("192.168.20.15") Then
        ConexaoVPN = True
		GerarCSV = False
    End If

    '======================================================================
    ' 2) Se conexão falhar, permite fallback para CSV
    '======================================================================
    If Not ConexaoVPN Then
        Dim GeraCSV
        GeraCSV = MsgBox( _
            "Não foi possível estabelecer conexão com os servidores de Quality Assurance (192.168.20.12 / 192.168.20.15)." & vbCrLf & _
            "Deseja gerar relatório local em formato csv?", _
            vbYesNo + vbQuestion + vbDefaultButton2, "Conexão com VPN falhou")

        If GeraCSV = vbYes Then
		MsgBox "Atenção: O relatório gerado não irá inserir as informações no banco de dados da equipe de Quality Assurance. Quando possível, realize a verificação em conjunto com a equipe.", vbInformation, "Aviso"
            GerarCSV = True
        Else
            MsgBox "Execução cancelada. Sem conexão com o banco e sem autorização para gerar CSV.", vbCritical
            Exit Sub
        End If
    Else
        '==================================================================
        ' 3) Se conexão VPN disponível, valida conexão com o banco
        '==================================================================
        Dim connTest
        Set connTest = ConectarBancoQA()

        If connTest Is Nothing Then
            MsgBox _
                "====================================" & vbCrLf & _
                "Falha ao conectar ao banco de dados do AutoTester." & vbCrLf & vbCrLf & _
                "Caso o problema persista, entre em contato com a equipe de Quality Assurance." & vbCrLf & _
                "===================================="
            Exit Sub
        Else
            connTest.Close
            Set connTest = Nothing
        End If
    End If

    Main
End Sub

'***********************************************************************
'*  Seção : Declaração de Variáveis Globais
'*----------------------------------------------------------------------
'*  Finalidade :
'*     • Criar dicionários que armazenam os resultados das verificações
'*       (Excel, TXT, Banco de Dados) e caches auxiliares.
'*     • Definir o caminho onde serão gravados logs e relatórios.
'*
'*  Dicionários criados:
'*     ‑ DadosExcel        : (Objeto, Tipo, Problema)  → Código/Aviso
'*     ‑ DadosTxt          : Índice incremental        → Linha de log
'*     ‑ DadosBancoDeDados : DBServer/PathName         → Coleção de objetos
'*     ‑ ListaObjetosLib   : Chave genérica            → Objeto de biblioteca
'*     ‑ TiposRegistrados  : TypeName                  → Boolean (já verificado)
'*
'*  Propriedades do Objeto:
'*     ‑ VerificarBancosDriversCustomizadores  (Boolean) - Propriedade oculta
'*     ‑ DebugMode              (Boolean) - Propriedade oculta
'*     ‑ PathNameTelas          (String)
'*     ‑ PastaParaSalvarLogs    (String)
'*     ‑ Empreendimento         (String)
'*     ‑ Projeto                (String)
'*     ‑ Localidade             (String)
'*     ‑ ResponsavelQA          (String) - Caso a propriedade esteja vazia, irá preencher com "AutoTester"
'***********************************************************************

Dim DadosExcel, DadosTxt, DadosBancoDeDados, ListaObjetosLib, TiposRegistrados, CaminhoPrj, TiposUnicos

'-- Instanciação dos dicionários -----------------------------------------------
Set DadosExcel = CreateObject("Scripting.Dictionary")
Set DadosTxt = CreateObject("Scripting.Dictionary")
Set DadosBancoDeDados = CreateObject("Scripting.Dictionary")
Set ListaObjetosLib = CreateObject("Scripting.Dictionary")
Set TiposRegistrados = CreateObject("Scripting.Dictionary")
Set TiposUnicos = CreateObject("Scripting.Dictionary")
'--------------------------------------------------------------------------------

'-- Definição do diretório de saída ---------------------------------------------
If PastaParaSalvarLogs <> "" Then
    CaminhoPrj = PastaParaSalvarLogs ' Caminho informado externamente
Else
    CaminhoPrj = CreateObject("WScript.Shell").CurrentDirectory
End If

'--------------------------------------------------------------------------------
'***********************************************************************
'*  Sub‑rotina : Main
'*----------------------------------------------------------------------
'*  Finalidade :
'*     1) Obter e verificar telas indicadas em PathNameTelas.
'*     2) Verificar demais objetos de domínio (DataServer, Folder…).
'*     3) Verificar Servidores de Alarme e Campos de Usuário.
'*     4) (Opcional) Verificar configurações de Hist / Historian.
'*     5) Gerar relatórios TXT e Excel com os resultados.
'***********************************************************************
Sub Main()

    Dim telaArray, tempoInicio, tempoFim, tempoGasto, tempoFormatado
	tempoInicio = Timer

    '------------------------------------------------------------------
    ' 1) Obter lista de telas para verificação
    '------------------------------------------------------------------
    telaArray = SplitTelas(PathNameTelas)

    '------------------------------------------------------------------
    ' 2) Verificar as telas
    '------------------------------------------------------------------
    VerificarTelas telaArray

    If VerificarBancosDriversCustomizadores = True Then

    '------------------------------------------------------------------
    ' 3) Verificar demais objetos do domínio (exceto telas)
    '------------------------------------------------------------------
	For Each Objeto In Application.ListFiles()
   		If TypeName(Objeto) <> "Screen" Then
       	VerificarPropriedadesObjetoBase Objeto
    	End If
	Next

    '------------------------------------------------------------------
    ' 4) Verificar Servidores de Alarme e Campos de Usuário
    '------------------------------------------------------------------
	For Each objServidor In Application.ListFiles("DB.AlarmServer")
    	VerificarUnicidadeObjetoTipo objServidor, "DB.AlarmServer", "Banco de dados", 0, False
    	VerificarCamposUsuariosServidorAlarmes objServidor
	Next
    End If

    '------------------------------------------------------------------
    ' 5) Geração de relatórios
    '------------------------------------------------------------------
    If Not DebugMode Then

        If GerarLogErrosScript Then
            If Not GerarRelatorioTxt(DadosTxt, CaminhoPrj) Then
                MsgBox _
						"====================================" & vbCrLf & _
						"Falha ao gerar o relatório TXT." & vbCrLf & _
						"Verifique se o diretório está acessível e se não há arquivos bloqueando a escrita.", _
						"===================================="
            End If
        End If

        If GerarCSV Then
            If Not GerarRelatorioExcel(DadosExcel, CaminhoPrj) Then
            	MsgBox _
						"====================================" & vbCrLf & _
						"Falha ao gerar o relatório Excel." & vbCrLf & _
						"Verifique se o Microsoft Excel está instalado corretamente e se o arquivo não está aberto.", _
						"===================================="
            End If
        Else
            Dim connDB
            Set connDB = ConectarBancoQA()
            
            If connDB Is Nothing Then
                MsgBox _
						"====================================" & vbCrLf & _
						"Falha ao conectar ao banco de dados Quality Assurance." & vbCrLf & _
						"Certifique-se de que a VPN está ativa e que as credenciais estão corretas." & vbCrLf & _
						"Contate a equipe de Quality Assurance caso o problema persista.", _
						"===================================="
            Else
                InserirInconsistenciasBanco DadosExcel, connDB
                connDB.Close
                Set connDB = Nothing

                tempoFim = Timer
                tempoGasto = Round(tempoFim - tempoInicio, 2)
				If tempoGasto >= 60 Then
    				tempoFormatado = Int(tempoGasto / 60) & " min " & Round(tempoGasto Mod 60) & " s"
				Else
    				tempoFormatado = Round(tempoGasto, 2) & " s"
				End If
				MsgBox _
    				"Inconsistências registradas com sucesso no banco de dados do Quality Assurance." & vbCrLf & vbCrLf & _
    				"────────────  Informações Gerais  ────────────" & vbCrLf & _
    				"Número de telas verificadas: " & Item("ContagemTelas").Value & vbCrLf & _
    				"Número de objetos de tela verificados: " & Item("ContagemObjetosTelas").Value & vbCrLf & _
    				"Número de objetos de base verificados: " & Item("ContagemObjetosBase").Value & vbCrLf & _
    				"Tempo gasto: " & tempoFormatado & vbCrLf & vbCrLf & _
    				"Acesse http://app.qa.devolutivas para visualizar as inconsistências geradas.", _
        			vbInformation, "AutoTester - Resultado"
            End If
        End If
    End If

	'------------------------------------------------------------------
	' 6) Limpeza de objetos
	'------------------------------------------------------------------
	Set DadosExcel = Nothing
	Set DadosTxt = Nothing
	Set DadosBancoDeDados = Nothing
	Set ListaObjetosLib = Nothing
	Set TiposRegistrados = Nothing
	Item("ContagemTelas").Value = 0
	Item("ContagemObjetosTelas").Value = 0
	Item("ContagemObjetosBase").Value = 0
End Sub

'***********************************************************************
'*  Sub‑rotina : VerificarTelas
'*----------------------------------------------------------------------
'*  Finalidade :
'*     • Se a propriedade PathNameTelas contiver nomes de telas, verifica
'*       apenas essas telas.
'*     • Caso contrário, percorre e verifica todas as telas do domínio.
'*
'*  Parâmetros :
'*     ‑ telaArray (Variant) : Array de strings com os PathNames das
'*                             telas a serem inspecionadas. Se vazio
'*                             (UBound < 0), todas as telas são analisadas.
'***********************************************************************
Sub VerificarTelas(telaArray)

    Dim Objeto

    If UBound(telaArray) >= 0 Then

        '--------------------------------------------------------------
        ' Há telas específicas definidas; verifica somente essas
        '--------------------------------------------------------------
        For Each Objeto In Application.ListFiles("Screen")
            If IsTelaNaLista(Objeto.PathName, telaArray) Then
                VerificarPropriedadesObjetoTela Objeto
				Item("ContagemTelas").Value = Item("ContagemTelas").Value + 1
            End If
        Next
    Else

        '--------------------------------------------------------------
        ' Nenhuma tela específica indicada; verifica todas as telas
        '--------------------------------------------------------------
        For Each Objeto In Application.ListFiles("Screen")
            VerificarPropriedadesObjetoTela Objeto
			Item("ContagemTelas").Value = Item("ContagemTelas").Value + 1
        Next

    End If

End Sub

'***********************************************************************
'*  Função : SplitTelas
'*----------------------------------------------------------------------
'*  Finalidade :
'*     Receber a string PathNameTelas e dividi‑la em um array, usando
'*     “/” como delimitador.  Se PathNameTelas estiver vazia, devolve um
'*     array vazio.
'***********************************************************************
Function SplitTelas(PathNameTelas)

    If Len(Trim(PathNameTelas)) > 0 Then
        SplitTelas = Split(PathNameTelas, "/")
    Else
        SplitTelas = Array()' Retorna array vazio
    End If

End Function

'***********************************************************************
'*  Função : IsTelaNaLista
'*----------------------------------------------------------------------
'*  Finalidade :
'*     Verificar se o PathName de uma tela está contido em telaArray.
'*     A comparação é case-insensitive e baseada apenas em "/".
'*
'*  Parâmetros :
'*     ‑ PathName  (String)  : Path completo da tela analisada.
'*     ‑ telaArray (Variant) : Array de PathNames (strings) de telas
'*                              que devem ser verificadas.
'***********************************************************************
Function IsTelaNaLista(PathName, telaArray)

    Dim tela

    For Each tela In telaArray
        tela = Trim(tela)
        If tela <> "" And StrComp(PathName, tela, vbTextCompare) = 0 Then
            IsTelaNaLista = True
            Exit Function
        End If
    Next

    IsTelaNaLista = False

End Function

'***********************************************************************
'*  Função : VerificarPropriedadesObjetoBase
'*----------------------------------------------------------------------
'*  Finalidade :
'*     Verificar o tipo do objeto recebido e acionar as rotinas de
'*     validação correspondentes.  Para objetos contêineres
'*     (Screen, DataServer, Folder, DrawGroup) a verificação é
'*     recursiva, percorrendo todos os seus filhos.
'***********************************************************************
Function VerificarPropriedadesObjetoBase(Obj)

    Dim TipoObjeto, child
    TipoObjeto = TypeName(Obj)

    Select Case TipoObjeto

    '=================================================================
    ' Objetos contêineres  →  verificação recursiva
    '=================================================================
Case "DataServer", "DataFolder", "Folder", "Screen", "DrawGroup"
    For Each child In Obj
        VerificarPropriedadesObjetoBase child
		Item("ContagemObjetosBase").Value = Item("ContagemObjetosBase").Value + 1
    Next
	  '=================================================================
	  ' Fluxo de dados
	  '================================================================= 
Case "frCustomAppConfig"
	VerificarBancoDeDados Obj, TipoObjeto, "AppDBServerPathName", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, TipoObjeto, "DataServer", 1, "Fluxo de dados", 1, False
	VerificarBibliotecasPorTypeName Obj, TipoObjeto, 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "ww_Parameters"
	VerificarBancoDeDados Obj, TipoObjeto, "DBServer", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, TipoObjeto, "DataServer", 1, "Fluxo de dados", 1, False
	VerificarBibliotecasPorTypeName Obj, TipoObjeto, 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "DatabaseTags_Parameters"
	VerificarBancoDeDados Obj, TipoObjeto, "DBServerPathName", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, TipoObjeto, "DataServer", 1, "Fluxo de dados", 1, False
	VerificarBibliotecasPorTypeName Obj, TipoObjeto, 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "cmdscr_CustomCommandScreen"
	VerificarBancoDeDados Obj, TipoObjeto, "DBServerPathName", "Fluxo de dados", 0, False
	'-----------------------------------------------------------------------------
Case "patm_CmdBoxXmlCreator"
	VerificarPropriedadeVazia Obj, TipoObjeto, "ConfigPower", "Fluxo de dados", 0, False
	'-----------------------------------------------------------------------------
Case "patm_DeviceNote"
	VerificarPropriedadeVazia Obj, TipoObjeto, "AlarmSource", "Fluxo de dados", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "NoteDatabaseControl", "Fluxo de dados", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject", "Fluxo de dados", 1, False
	VerificarBibliotecasPorTypeName Obj, TipoObjeto, 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "patm_NoteDatabaseControl"
	VerificarBancoDeDados Obj, TipoObjeto, "DBServer", "Fluxo de dados", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "GroupCanAddModifyNote", "Opera  o", 1, "Fluxo de dados", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Level", "2=[EquipeSCADA, Instrutor]/3=[Supervis o]/4=[EquipeDeTestes]/5=[Opera  o]", 1, "Fluxo de dados", 0, False
	VerificarTipoPai Obj, TipoObjeto, "DataServer", 1, "Fluxo de dados", 1, False
	VerificarBibliotecasPorTypeName Obj, TipoObjeto, 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "patm_xoAlarmHistConfig"
	VerificarBancoDeDados Obj, TipoObjeto, "MainDBServerPathName", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, TipoObjeto, "DataServer", 1, "Fluxo de dados", 1, False
	VerificarBibliotecasPorTypeName Obj, TipoObjeto, 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "dtRedundancyConfig"
	VerificarPropriedadeVazia Obj, TipoObjeto, "NameOfServerToBeStopped", "Fluxo de dados", 1, False
	VerificarTipoPai Obj, TipoObjeto, "DataServer", 1, "Fluxo de dados", 1, False
	VerificarBibliotecasPorTypeName Obj, TipoObjeto, 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "patm_CommandLogger"
	VerificarPropriedadeVazia Obj, TipoObjeto, "PowerConfigObj", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, TipoObjeto, "DataServer", 1, "Fluxo de dados", 1, False
	VerificarBibliotecasPorTypeName Obj, TipoObjeto, 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "hpXMLGenerateStruct"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Log_BancoDeDados", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, TipoObjeto, "DataServer", 1, "Fluxo de dados", 1, False
	VerificarBibliotecasPorTypeName Obj, TipoObjeto, 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "gtwFrozenMeasurements"
	VerificarPropriedadeVazia Obj, TipoObjeto, "DateTag", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, TipoObjeto, "DataServer", 1, "Fluxo de dados", 1, False
	VerificarBibliotecasPorTypeName Obj, TipoObjeto, 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "aainfo_NoteController"
	VerificarPropriedadeVazia Obj, TipoObjeto, "DBServerPathName", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, TipoObjeto, "DataServer", 1, "Fluxo de dados", 1, False
	VerificarBibliotecasPorTypeName Obj, TipoObjeto, 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "xoExecuteScheduler"
	VerificarPropriedadeVazia Obj, TipoObjeto, "aActivateCommandsGroup", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "dteEndEvent", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "dteEndRepeatDate", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "dteNextEndEvent", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "dteNextStartEvent", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "objCommand", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "strSchedulerName", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "UserField01", "Fluxo de dados", 0, False
	'-----------------------------------------------------------------------------
Case "manut_ImportMeasAndCmdList"
	VerificarObjetoInternoIndevido Obj, TipoObjeto, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "xots_StandardStudioSettings"
	VerificarObjetoInternoIndevido Obj, TipoObjeto, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "xots_ConvertAqDriversIntoVbScri"
	VerificarObjetoInternoIndevido Obj, TipoObjeto, "Fluxo de dados", 1, False
	  '=================================================================
	  ' Dominio
	  '================================================================= 
	'-----------------------------------------------------------------------------
Case "AlarmServer"
	VerificarPropriedadeVazia Obj, TipoObjeto, "DataSource", "Dominio", 0, False
	'-----------------------------------------------------------------------------
Case "DBServer"
	VerificarPropriedadeValor Obj, TipoObjeto, "SourceType", 2, 0, "Dominio", 0, False
	'-----------------------------------------------------------------------------
Case "Hist"
	VerificarPropriedadeVazia Obj, TipoObjeto, "DBServer", "Dominio", 1, False
	VerificarBancoDeDados Obj, TipoObjeto, "DBServer", "Dominio", 0, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "EnableDiscard", False, "Dominio", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "DiscardInterval", 1, 1, "Dominio", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "DiscardTimeUnit", 2, 1, "Dominio", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "VerificationInterval", 1, 1, "Dominio", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "VerificationUnit", 2, 1, "Dominio", 0, False
	End If
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "EnableBackupTable", False, "Dominio", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "BackupDiscardInterval", 12, 1, "Dominio", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "BackupDiscardTimeUnit", 2, 1, "Dominio", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "Historian"
	VerificarPropriedadeVazia Obj, TipoObjeto, "DBServer", "Dominio", 1, False
	VerificarBancoDeDados Obj, TipoObjeto, "DBServer", "Dominio", 0, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "EnableDiscard", False, "Dominio", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "DiscardInterval", 1, 1, "Dominio", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "DiscardTimeUnit", 2, 1, "Dominio", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "VerificationInterval", 1, 1, "Dominio", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "VerificationUnit", 2, 1, "Dominio", 0, False
	End If
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "EnableBackupTable", False, "Dominio", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "BackupDiscardInterval", 12, 1, "Dominio", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "BackupDiscardTimeUnit", 2, 1, "Dominio", 0, False
	End If
	  '=================================================================
	  ' Drivers
	  '================================================================= 
	'-----------------------------------------------------------------------------
Case "IODriver"
	VerificarPropriedadeVazia Obj, TipoObjeto, "DriverLocation", "Drivers", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "WriteSyncMode", 2, 0, "Drivers", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ExposeToOpc", 3, 0, "Drivers", 0, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "EnableReadGrouping ", True, "Telas", 1, False
	ContarObjetosDoTipo Obj, TipoObjeto, "IOTag", 0,1,1, "Drivers", 1, False
	  '=================================================================
	  ' Biblioteca
	  '================================================================= 
	'-----------------------------------------------------------------------------
Case "WaterConfig"
	VerificarPropriedadeVazia Obj, TipoObjeto, "ModelFile", "Biblioteca", 0, False
	'-----------------------------------------------------------------------------
Case "WaterDistributionNetwork"
	VerificarPropriedadeVazia Obj, TipoObjeto, "City", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Company", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "CompanyAcronym", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Contract", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Name", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Neighborhood", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Organization", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Region", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "State", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "StateAcronym", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Note", "Biblioteca", 0, False
	Dim containerTypes
	containerTypes = Array("DataFolder", "DrawGroup", "DataServer", "WaterDistributionNetwork")
	If HasChildOfType(Obj, "WaterStationData", containerTypes) Then
	Dim arrUserFields
	arrUserFields = Array("DadosDaPlanta", "Mapa3D")
	VerificarUserFields Obj, arrUserFields, "WaterDistributionNetwork", 1
	End If
Case Else
    RegistrarTipoSemPropriedade TipoObjeto
End Select
End Function

'***********************************************************************
'*  Função : VerificarPropriedadesObjetoTela
'*----------------------------------------------------------------------
'*  Finalidade :
'*     Verificar o tipo do objeto recebido e acionar as rotinas de
'*     validação correspondentes.  Para objetos contêineres
'*     (Screen, DataServer, Folder, DrawGroup) a verificação é
'*     recursiva, percorrendo todos os seus filhos.
'***********************************************************************
Function VerificarPropriedadesObjetoTela(Obj)

    Dim TipoObjeto, child, Area
    TipoObjeto = TypeName(Obj)
	Area = "Telas"

    '------------------------------------------------------
    ' Ignorar objetos de linha se VerificarLinhas = True
    '------------------------------------------------------
    If Not EnergizacaoLinhasBarras Then
        Select Case TipoObjeto
            Case "archLineHorizontal", "archLineVertical", "pwa_LineHoriz", "pwa_Barra", "pwa_Barra2", "pwa_Barra2Vert", "pwa_LineVert", "Uhe_LineHoriz", "Uhe_LineVert", "pwa_Carga", "pwa_Bateria", _
				"pwa_Conexao", "pwa_Jumper", "pwa_Terra", "pwa_Terra2", "uhe_EarthSwitch", "uhe_ExcitationTransformer"
                Exit Function ' Apenas ignora este objeto atual
        End Select
    End If

Select Case TipoObjeto
	'=================================================================
	' Objetos contêineres  →  verificação recursiva
	'=================================================================

Case "DataServer", "DataFolder", "Folder", "Screen", "DrawGroup"

    For Each child In Obj
        VerificarPropriedadesObjetoTela child
		Item("ContagemObjetosTelas").Value = Item("ContagemObjetosTelas").Value + 1
    Next
	  '=================================================================
	  ' Telas
	  '================================================================= 
	'-----------------------------------------------------------------------------
Case "DrawString"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Value", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "TextColor", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "E3Query"
	VerificarPropriedadeVazia Obj, TipoObjeto, "DataSource", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Disjuntor"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "PositionMeas", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "DeviceNote", False, "Telas", 1, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "HideLines", True, "Telas", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 0, 1, "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "pwa_DisjuntorP"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "PositionMeas", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "DeviceNote", False, "Telas", 1, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "HideLines", True, "Telas", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 0, 1, "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "pwa_DisjuntorPP"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "PositionMeas", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "DeviceNote", False, "Telas", 1, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "HideLines", True, "Telas", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 0, 1, "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "pwa_Seccionadora"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "PositionMeas", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "DeviceNote", False, "Telas", 1, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "HideLines", True, "Telas", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 0, 1, "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "pwa_BarraAlarme"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "AnalogMeas", False, "Telas", 0, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "AlarmSource", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "AnalogMeas", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ValorMaximo", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ValorMinimo", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Barra"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Barra2"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Barra2Vert"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Bateria"
	VerificarPropriedadeValor Obj, TipoObjeto, "Energizado", True, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 15790320, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 5263440, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_BotaoAbreTela"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Config_Zoom", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Config_TelaQuadroPatName", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Config_Descricao", "Desccri  o", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Gerador"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "GenEstado", "SourceObject", "NOTEMPTY", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_GeradorG"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "PotenciaMedia", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "PotenciaMedia", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "PotenciaMaximaNominal", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Gerador", True, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_HomeButton"
	VerificarPropriedadeVazia Obj, TipoObjeto, "ScreenOrFramePathName", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "ScreenDescription", "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Description", "Alarmes", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_GrupoVSL"
	VerificarPropriedadeVazia Obj, TipoObjeto, "PositionMeasObject", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "AnalogMeas", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoAlarme01"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descricao", "XXX", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoAlarme05"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descricao", "XXX", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoAlarme10"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descricao", "XXX", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoAnalogica"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SourceObject", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SPShow", "SPTag", True, "Telas", 1, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "SourceObject", ".Value", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "ShowUE", True, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoAnalogicaG"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SourceObject", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SPShow", "SPTag", True, "Telas", 1, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "SourceObject", ".Value", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "ShowUE", True, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoPot"
	VerificarPropriedadeVazia Obj, TipoObjeto, "PotenciaMedida", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "PotenciaMedia", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "HabilitaSetpoint", "SetPointPotencia", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "PotenciaMaximaNominal", 0, 1, "Telas", 0, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoPotP"
	VerificarPropriedadeVazia Obj, TipoObjeto, "PotenciaMedida", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "PotenciaMedia", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "HabilitaSetpoint", "SetPointPotencia", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "PotenciaMaximaNominal", 100, 1, "Telas", 0, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoPotG"
	VerificarPropriedadeVazia Obj, TipoObjeto, "PotenciaMedida", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "PotenciaMedia", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "HabilitaSetpoint", "SetPointPotencia", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "PotenciaMaximaNominal", 100, 1, "Telas", 0, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Inversor"
	VerificarPropriedadeValor Obj, TipoObjeto, "Energizado", True, 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_LineHoriz"
	VerificarPropriedadeValor Obj, TipoObjeto, "Energizado", True, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 15790320, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 5263440, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_LineVert"
	VerificarPropriedadeValor Obj, TipoObjeto, "Energizado", True, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 15790320, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 5263440, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_AutoTrafo"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SourceObject", "DeviceNote", "NOTEMPTY", "Telas", 1, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "HideLine", True, "Telas", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal2", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal3", 0, 1, "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "pwa_Capacitor"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Carga"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SourceObject", "DeviceNote", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Conexao"
	VerificarPropriedadeValor Obj, TipoObjeto, "CorObjeto", 4605520, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Jumper"
	VerificarPropriedadeValor Obj, TipoObjeto, "CorObjeto", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Retificador"
	VerificarPropriedadeValor Obj, TipoObjeto, "Energizado", True, 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_Terra"
	VerificarPropriedadeValor Obj, TipoObjeto, "CorTerrra", 16777215, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Terra2"
	VerificarPropriedadeValor Obj, TipoObjeto, "CorTerrra", 16777215, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Reactor"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Relig"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "NaoSupervisionado", "PositionMeas", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "UseNotes", "DeviceNote", True, "Telas", 1, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "HideLines", True, "Telas", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 0, 1, "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "pwa_Sensor"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "BorderColor", 255, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_VentForc"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Unsupervised", "Measure", False, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_TapV"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "CmdDown", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "CmdUp", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MaxLimit", 8, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MinLimit", 2, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoPotRea"
	VerificarPropriedadeVazia Obj, TipoObjeto, "PotRea", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "PotRea", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SpShow", "SetPointPotencia", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MaxPotReaPos", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MinPotReaPos", -100, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_ReguladorTensao"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SourceObject", "DeviceNote", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "UseNotes", "TAPMeas", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal2", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Menu"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "SpecialScreens", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "ScreenArg", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_TrafoSA"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "UseNotes", "DeviceNote", True, "Telas", 1, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "HideLine", True, "Telas", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal2", 0, 1, "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "pwa_Trafo3Type01"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "TAPSPShow", "TAPSPTag", True, "Telas", 1, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "HideLine", True, "Telas", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal2", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal3", 0, 1, "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "pwa_Trafo3_P"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "TAPSPShow", "TAPSPTag", True, "Telas", 1, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "HideLine", True, "Telas", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal2", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal3", 0, 1, "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "pwa_Trafo3"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "TAPSPShow", "TAPSPTag", True, "Telas", 1, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "HideLine", True, "Telas", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal2", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal3", 0, 1, "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "pwa_Trafo2Term"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "TAPSPShow", "TAPSPTag", True, "Telas", 1, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "HideLine", True, "Telas", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal2", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal3", 0, 1, "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "pwa_Trafo2"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "TAPSPShow", "TAPSPTag", True, "Telas", 1, False
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "HideLine", True, "Telas", 2, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal2", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOnTerminal3", 0, 1, "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "gx_AbnormalityIndicator"
	VerificarPropriedadeValor Obj, TipoObjeto, "Measurement01Active", False, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Measurement01Desc", "XXX", 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "SideToGrowing", 1, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_Analogic"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Show", False, 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_ButtonOpenCommandScreen"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Description", "descri  o", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_Counter"
	VerificarPropriedadeValor Obj, TipoObjeto, "Value", 0, 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_CtrlDigital"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Active", False, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr", "Desc", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_CtrlDigital1Op"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr", "Desc", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_CtrlDigital2Op"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames1", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag1", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr1", "Desc1", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames2", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr2", "Desc2", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_CtrlDigital3Op"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames1", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag1", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr1", "Desc1", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames2", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr2", "Desc2", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames3", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag3", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr3", "Desc3", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_CtrlDigital4Op"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames1", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag1", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr1", "Desc1", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames2", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr2", "Desc2", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames3", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag3", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr3", "Desc3", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames4", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag4", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr4", "Desc4", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_InfoAnalogic"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SPShow", "SPTag", True, "Telas", 1, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "ShowUE", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_InfoAnalogic2"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SPShow", "SPTag", True, "Telas", 1, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "ShowUE", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_InfoDoughnutChart"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SPShow", "SPTag", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "NominalValue", 100, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_InfoSetpoint"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SPTag", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "ShowUE", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_Menu"
	'Existe evento OnClick
	'-----------------------------------------------------------------------------
Case "gx_Notes"
	VerificarPropriedadeVazia Obj, TipoObjeto, "DeviceNote", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_OpenTabularArea1"
	VerificarPropriedadeValor Obj, TipoObjeto, "Descricao", "XXX", 1, "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Areas", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "ScreenZoom", "ScreenPathNames ", "NOTEMPTY", "Telas", 1, False
	'Talvez nova fun  o se propriedade auxiliar estiver preenchidad verficar se a propriedade principal est  dfierente de 0
	'-----------------------------------------------------------------------------
Case "gx_QualityIcon"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measurement", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart03"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart04"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart05"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart06"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart07"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart08"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas08", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart08_2Z"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas06_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas06_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas07_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas07_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas08_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas08_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Title01", "Quente", 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Title02", "Frio", 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart09"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas08", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas09", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas09MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas09MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart10"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas08", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas09", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas09MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas09MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas10", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas10MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas10MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart10_2Z"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas06_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas06_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas07_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas07_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas08_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas08_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas09_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas09_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas09MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas09MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas10_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas10_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas10MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas10MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Title01", "Quente", 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Title02", "Frio", 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart12"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas08", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas09", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas09MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas09MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas10", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas10MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas10MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas11", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas11MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas11MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas12", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas12MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas12MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart16"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas08", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas09", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas09MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas09MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas10", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas10MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas10MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas11", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas11MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas11MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas12", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas12MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas12MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas13", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas13MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas13MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas14", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas14MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas14MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas15", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas15MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas15MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas16", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas16MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas16MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart20"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas08", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas09", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas09MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas09MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas10", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas10MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas10MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas11", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas11MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas11MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas12", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas12MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas12MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas13", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas13MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas13MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas14", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas14MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas14MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas15", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas15MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas15MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas16", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas16MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas16MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas17", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas17MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas17MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas18", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas18MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas18MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas19", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas19MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas19MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Meas20", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas20MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Meas20MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_AbnormalityIndicator"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measurement01Active", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Measurement01Desc", "XXX", 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "SideToGrowing", 1, 1, "Telas", 0, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_AirCompressor"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Unsupervised", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Unsupervised", "CompressorOff", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Unsupervised", "CompressorOn", False, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AirOilTank"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MaxLimit", 3, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MinLimit", 0, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AlarmBar"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ValorMaximo", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ValorMinimo", 0, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AnalogBar"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MaxValue", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MinValue", 0, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AnalogBar5Limits"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MaxValue", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MinValue", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Limit01", 50, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Limit02", 60, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Limit03", 70, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Limit04", 80, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Limit05", 90, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AnalogBar5LimitsH"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MaxValue", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MinValue", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Limit01", 50, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Limit02", 60, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Limit03", 70, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Limit04", 80, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Limit05", 90, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AnalogBarHor"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MaxValue", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MinValue", 0, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AnalogBarP"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MaxValue", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MinValue", 0, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_BielaHidraulica"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Flambada", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_BielaMecanica"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Flambada", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Block"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Block_Tag", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "BlockArea", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "OpenCommandSelectMenu", "CommandPathNames", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Bomb"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Unsupervised", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "OpenCommandSelectMenu", "CommandPathNames", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Unsupervised", "DeviceNote", False, "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "BombOff", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "BombOn", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Bomb2"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Unsupervised", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "OpenCommandSelectMenu", "CommandPathNames", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Unsupervised", "DeviceNote", False, "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "BombOff", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "BombOn", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_BrakeAlert"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SourceObject", "DeviceNote", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "BrakeTag", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_BulbTurbine"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Distributor_Tag", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_BusBar"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ObjectColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_Buzzer"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Playing", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Caixa"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Energized", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Chart"
	VerificarPropriedadeVazia Obj, TipoObjeto, "PenData1", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ObjectColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_ChartP"
	VerificarPropriedadeVazia Obj, TipoObjeto, "PenData1", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ObjectColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_Command"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommandPathNames", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr", "Desccri  o", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_CommandButton"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommandPathNames", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Description", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_Conduto"
	VerificarPropriedadeVazia Obj, TipoObjeto, "ComAgua", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_ControlGate"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Unsupervised", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "StateOff_Tag", "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "StateOn_Tag", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlDigital"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Active", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr", "Desc", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlDigital1Op"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr", "Desc", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlDigital2Op"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames1", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag1", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr1", "Desc1", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames2", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr2", "Desc2", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlDigital3Op"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames1", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag1", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr1", "Desc1", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames2", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr2", "Desc2", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames3", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag3", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr3", "Desc3", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlDigital4Op"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames1", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag1", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr1", "Desc1", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames2", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag2", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr2", "Desc2", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames3", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag3", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr3", "Desc3", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames4", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Tag4", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr4", "Desc4", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlDigitalOp"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "CommandPathNames", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Active", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descr", "Desc", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlPulse"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CmdDecrement", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "CmdIncrement", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Device"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_DieselGenerator"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SourceObject", "DeviceNote", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Estado_Tag", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descricao", "GDE", 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_Direction"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "AnalogMeasure", True, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "ObjectColor", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_EarthSwitch"
	VerificarPropriedadeValor Obj, TipoObjeto, "CorTerrra", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "ObjectColor", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_ExcitationTransformer"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Off", "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "On", "Telas", 0, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "OpenCommandSelectMenu", "CommandPathNames", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Filter"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_Fan"
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Unsupervised", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "SourceObject", "DeviceNote", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "FanOff", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "FanOn", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "OpenCommandSelectMenu", "CommandPathNames", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_FilterSelfCleaning"
	VerificarPropriedadeValor Obj, TipoObjeto, "FilterOn", False, 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_FrancisTurbine"
	VerificarPropriedadeValor Obj, TipoObjeto, "Energized", False, 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_InfoAlarm01", "uhe_InfoAlarm03", "uhe_InfoAlarm05", "uhe_InfoAlarm10"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject01", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Descricao", "XXX", 1, "Telas", 1, False
	If VerificarPropriedadeVazia (Obj, TipoObjeto, "SourceObject01", "Telas", 1, True) Then
	VerificarPropriedadeVazia Obj, TipoObjeto, "ScreenPathNames", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "OpeningMode", 0, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "ScreenZoom", "Telas", 1, False
	If VerificarPropriedadeValor (Obj, TipoObjeto, "OpeningMode", 4, 1, "Telas", 0, True) Then
	VerificarPropriedadeVazia Obj, TipoObjeto, "CustomScriptOpeningMode", "Telas", 1, False
	End If
	End If
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_InfoAnalogic", "uhe_InfoAnalogic2"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	If VerificarPropriedadeVazia (Obj, TipoObjeto, "Measure", "Telas", 1, True) Then
	VerificarPropriedadeVazia Obj, TipoObjeto, "AlarmSource", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "ShowUE", True, "Telas", 1, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	End If
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "SPShow", True, "Telas", 1, True) Then
	VerificarPropriedadeVazia Obj, TipoObjeto, "SPTag", "Telas", 1, False
	End If
	If VerificarPropriedadeVazia (Obj, TipoObjeto, "SPTag", "Telas", 1, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "SPShow", False, 1, "Telas", 1, False
	End If
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_InfoPotP"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	If VerificarPropriedadeVazia (Obj, TipoObjeto, "Measure", "Telas", 1, True) Then
	VerificarPropriedadeVazia Obj, TipoObjeto, "AlarmSource", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "ShowUE", True, "Telas", 1, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	End If
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "SPShow", True, "Telas", 1, True) Then
	VerificarPropriedadeVazia Obj, TipoObjeto, "SPTag", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "SetPointPotencia", "Telas", 1, False
	End If
	If VerificarPropriedadeVazia (Obj, TipoObjeto, "SPTag", "Telas", 1, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "SPShow", False, 1, "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "SetPointPotencia", "Telas", 1, False
	End If
	VerificarPropriedadeValor Obj, TipoObjeto, "NominalValue", 100, 1, "Telas", 0, False
	VerificarObjetoDesatualizado Obj, TipoObjeto, "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_InfoPotRea"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	If VerificarPropriedadeVazia (Obj, TipoObjeto, "Measure", "Telas", 1, True) Then
	VerificarPropriedadeVazia Obj, TipoObjeto, "AlarmSource", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "ShowUE", True, "Telas", 1, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject", "Telas", 1, False
	End If
	If VerificarPropriedadeHabilitada (Obj, TipoObjeto, "SPShow", True, "Telas", 1, True) Then
	VerificarPropriedadeVazia Obj, TipoObjeto, "SPTag", "Telas", 1, False
	End If
	If VerificarPropriedadeVazia (Obj, TipoObjeto, "SPTag", "Telas", 1, True) Then
	VerificarPropriedadeValor Obj, TipoObjeto, "SPShow", False, 1, "Telas", 1, False
	End If
	VerificarPropriedadeValor Obj, TipoObjeto, "ValueVisible", True, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MinPos", -100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "MaxPos", 100, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_InfoSetpoint"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SPTag", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "ShowUE", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Format", 0.00, 0, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_LineHoriz"
	VerificarPropriedadeValor Obj, TipoObjeto, "Energizado", True, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 16448250, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 5263440, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_LineVert"
	VerificarPropriedadeValor Obj, TipoObjeto, "Energizado", True, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOff", 16448250, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "CorOn", 5263440, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_Lock"
	If VerificarPropriedadeValor (Obj, TipoObjeto, "Visible", True, 1, "Telas", 1, True) Then
	VerificarAssociacaoBase Obj, TipoObjeto, "ClosedState", "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "uhe_PresDiferencial"
	VerificarPropriedadeValor Obj, TipoObjeto, "ActiveState", True, 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_PressureSwitch"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Measure", "Telas", 1, False
	If VerificarPropriedadeVazia (Obj, TipoObjeto, "Measure", "Telas", 1, True) Then
	VerificarPropriedadeVazia Obj, TipoObjeto, "AlarmSource", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "ShowUE", True, "Telas", 1, False
	VerificarPropriedadeTextoProibido Obj, TipoObjeto, "Measure", ".Value", "Telas", 1, False
	End If
	'-----------------------------------------------------------------------------
Case "uhe_Rectifier"
	If VerificarPropriedadeValor (Obj, TipoObjeto, "Enabled", True, 1, "Telas", 1, True) Then
	VerificarAssociacaoBase Obj, TipoObjeto, "Energizado", "Telas", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "XCPump"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "iconElectricity"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "iconComFail"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "xcLabel"
	VerificarPropriedadeVazia Obj, TipoObjeto, "Caption", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "xcEtiqueta_Manut"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CorObjeto", "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "EtiquetaVisivel", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "xcEtiqueta"
	VerificarPropriedadeVazia Obj, TipoObjeto, "AvisoVisivel", "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "EventoVisivel", "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "FonteObjeto", "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "ForaVisivel", "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "PathNote", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Visible", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "xcWaterTank"
	VerificarPropriedadeVazia Obj, TipoObjeto, "objSource", "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "objWaterDistribution", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Visible", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "XCDistribution"
	VerificarPropriedadeVazia Obj, TipoObjeto, "SourceObject", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "XCArrow"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Visible", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archAeroGenerator"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archCloud"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archNuclearPlant"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archServerRackmountMultiple"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archSolarPanel"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archSurveillanceCamera"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archWifi"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archDatabase"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archFirewall"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archPCH"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archUHE"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archVideoWall"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Enabled", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archSwitch"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archServerDesktop"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archRouter"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archServerRackmountSingle"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Description_Text", "Server name", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archViewer"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archElectricalMeter"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Description_Text", "IED NAME", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archGPSAntenna"
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archLineHorizontal"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Visible", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "FlhCom", "NOTEMPTY", "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "BorderColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archLineVertical"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Visible", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "FlhCom", "NOTEMPTY", "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "BorderColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archChannelPanel", "archChannelPanelP", "archChannelPanelPP", "archChannelPanelG"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Visible", True, "Telas", 1, False
		    Dim tipoPortas, i, qtdEsperada, propFailure
		    tipoPortas = GetPropriedade(Obj, "Type")
		    Select Case CStr(tipoPortas)
		        Case "3": qtdEsperada = 4
		        Case "2": qtdEsperada = 8
		        Case "1": qtdEsperada = 20
		        Case "0": qtdEsperada = 24
		        Case Else
		            qtdEsperada = 0 ' nao definido ou invalido
		    End Select
		    If qtdEsperada > 0 Then
		        For i = 1 To qtdEsperada
		            propFailure = "FailureState" & Right("0" & i, 2)
		            Call VerificarPropriedadeValor(Obj, TipoObjeto, propFailure, False, 1, "Telas", 1, False)
		        Next
		    End If
	'-----------------------------------------------------------------------------
Case "archLed"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Visible", True, "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "FailedState", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archModuloIO"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Visible", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "Failure", "NOTEMPTY", "Telas", 0, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "Text", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archRTU"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Visible", True, "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Description_Text", "RTU NAME", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "IP_Show", "IP_Text", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archIED"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Visible", True, "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "Enabled", "SourceObject", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Description_Text", "IED NAME", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, TipoObjeto, "IP_Show", "IP_Text", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archInfo"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Visible", True, "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "FailedState", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Description_Text", "Painel - Eqp - Port", 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ObjectColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archInfoLine"
	VerificarPropriedadeHabilitada Obj, TipoObjeto, "Visible", True, "Telas", 1, False
	VerificarPropriedadeVazia Obj, TipoObjeto, "FailedState", "Telas", 1, False
	VerificarPropriedadeValor Obj, TipoObjeto, "Description_Text", "Painel - Eqp - Port", 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, TipoObjeto, "ObjectColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case Else
    RegistrarTipoSemPropriedade TipoObjeto
End Select
End Function

'***********************************************************************
'*  Grupo : Funções de Acesso a Propriedades
'*----------------------------------------------------------------------

'-----------------------------------------------------------------------
'*  Função : GetPropriedade
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Retorna a propriedade do objeto com prioridade para Link.
'*     Se não houver link, retorna o valor direto.
'*     Se houver aspas no retorno do link, remove.
'-----------------------------------------------------------------------
Function GetPropriedade(Obj, PropName)

    Dim resultadoLink, resultadoValor
    resultadoLink  = GetPropertyLink(Obj, PropName)

    If Trim(resultadoLink) <> "" Then
        ' Se vier entre aspas, remove (ex: "100" → 100)
        If Left(resultadoLink, 1) = """" And Right(resultadoLink, 1) = """" Then
            resultadoLink = Mid(resultadoLink, 2, Len(resultadoLink) - 2)
        End If
        GetPropriedade = resultadoLink
    Else
        resultadoValor = GetPropertyValue(Obj, PropName)
        GetPropriedade = resultadoValor
    End If

End Function

'-----------------------------------------------------------------------
'*  Função : GetPropertyLink
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Tenta retornar a associação (link) de uma propriedade.
'*     Se não existir ou ocorrer erro, retorna vazio.
'-----------------------------------------------------------------------
Function GetPropertyLink(Obj, PropName)
    On Error Resume Next
    Dim tmp
    tmp = Obj.Links.Item(Trim(PropName)).Source
    If Err.Number <> 0 Then tmp = "" : Err.Clear
    On Error GoTo 0
    GetPropertyLink = tmp
End Function

'-----------------------------------------------------------------------
'*  Função : GetPropertyValue
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Retorna o valor direto de uma propriedade (via Eval).
'*     Se houver erro ou propriedade inexistente, retorna vazio.
'-----------------------------------------------------------------------
Function GetPropertyValue(Obj, PropName)
    On Error Resume Next
    Dim tmp
    tmp = Eval("Obj." & Trim(PropName))
    If Err.Number <> 0 Then tmp = "" : Err.Clear
    On Error GoTo 0
    GetPropertyValue = CStr(tmp)
End Function

'***********************************************************************
'*  Função : VerificarPropriedadeVazia
'*----------------------------------------------------------------------
'*  Finalidade :
'*     • Ler uma propriedade de um objeto e verificar se está vazia 
'*       ("" depois de Trim) ou se ocorreu erro de leitura.
'*     • Caso a propriedade esteja vazia, registrar a inconsistência no
'*       relatório Excel ou banco de dados.
'*
'*  Parâmetros :
'*     ‑ Obj            : Objeto a ser analisado.
'*     ‑ TipoObjeto     : TipoObjeto do objeto.
'*     ‑ Propriedade    : Nome da propriedade a ser verificada.
'*     ‑ AreaErro       : Área onde o erro foi encontrado.
'*     ‑ TipoErro       : Tipo de erro (0 = Aviso, 1 = Erro, 2 = Revisar).
'*     ‑ CondicaoRetorno: Se True, retorna True se a propriedade estiver vazia, caso contrário retorna False.
'***********************************************************************
Function VerificarPropriedadeVazia(Obj, TipoObjeto, Propriedade, AreaErro, TipoErro, CondicaoRetorno)

    On Error Resume Next

    Dim ValorLeitura, Mensagem
    ValorLeitura = GetPropriedade(Obj, Propriedade)

    If Trim(ValorLeitura) = "" Then
        If CondicaoRetorno = True Then
            VerificarPropriedadeVazia = True
            Exit Function
        End If 

        Mensagem = TipoObjeto & " com a propriedade " & Propriedade & " vazia."
        If GerarCSV Then
            AdicionarErroExcel DadosExcel, Obj.PathName, CStr(TipoErro), Mensagem, AreaErro, TipoObjeto
        Else
            AdicionarErroBanco DadosExcel, Obj.PathName, CStr(TipoErro), Mensagem, TipoObjeto, AreaErro
        End If
    Else
        If CondicaoRetorno = True Then VerificarPropriedadeVazia = False
    End If

    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

'*******************************************************************************
'*  Função : ContarObjetosDoTipo
'*-------------------------------------------------------------------------------
'*  Finalidade :
'*     Contabiliza objetos irmãos ou filhos de um objeto base, com filtros
'*     opcionais por tipo (TypeName) e lógica condicional para geração de log.
'*
'*  Parâmetros :
'*     - Obj                : Objeto base para contagem.
'*     - TipoObjeto         : Nome lógico (exibição) do objeto analisado.
'*     - TipoDesejado       : String com o TypeName a ser filtrado (ex: "IOTag").
'*     - DeveConterNao      : 
'*         0 → Espera que existam objetos do tipo (gera log se não houver).
'*         1 → Espera que NÃO existam objetos do tipo (gera log se houver).
'*     - TipoIgual          :
'*         0 → Conta todos os objetos (sem filtro de tipo).
'*         1 → Conta apenas objetos do tipo TipoDesejado.
'*     - ContagemIrmaoFilho :
'*         0 → Conta os irmãos (Parent).
'*         1 → Conta os filhos (Children).
'*     - AreaErro           : Categoria da verificação (ex: "Drivers", "Tags").
'*     - TipoErro           : Severidade (0=Aviso, 1=Erro, 2=Revisar).
'*     - CondicaoRetorno    : 
'*         True → Apenas retorna a contagem sem registrar erros.
'*         False → Executa registro de erro conforme regras.
'*
'*  Retorno :
'*     Integer → Número de objetos encontrados de acordo com os critérios.
'*******************************************************************************
Function ContarObjetosDoTipo(Obj, TipoObjeto, TipoDesejado, DeveConterNao, TipoIgual, ContagemIrmaoFilho, AreaErro, TipoErro, CondicaoRetorno)
    On Error Resume Next

    Dim contador, childObj, ParentObj, PathName, Msg
    contador = 0
    PathName = Obj.PathName

    ' Contagem baseada em irmãos ou filhos
    If ContagemIrmaoFilho = 0 Then
        Set ParentObj = Obj.Parent
        If Not ParentObj Is Nothing Then
            For Each childObj In ParentObj
                If TipoIgual = 0 Or (TipoIgual = 1 And TypeName(childObj) = TipoDesejado) Then
                    contador = contador + 1
                End If
            Next
        End If
    ElseIf ContagemIrmaoFilho = 1 Then
        For Each childObj In Obj
            If TipoIgual = 0 Or (TipoIgual = 1 And TypeName(childObj) = TipoDesejado) Then
                contador = contador + 1
            End If
        Next
    End If

    ' Apenas retorna se modo CondicaoRetorno estiver ativo
    If CondicaoRetorno Then
        ContarObjetosDoTipo = contador
        Exit Function
    End If

    ' Validação com base em DeveConterNao
    If (DeveConterNao = 0 And contador = 0) Or (DeveConterNao = 1 And contador > 0) Then
        If DeveConterNao = 0 Then
            Msg = "Nenhum objeto do tipo '" & TipoDesejado & "' encontrado "
            If ContagemIrmaoFilho = 0 Then
                Msg = Msg & "entre os irmãos de '" & PathName & "'."
            Else
                Msg = Msg & "entre os filhos de '" & PathName & "'."
            End If
        Else
            Msg = "Objeto do tipo '" & TipoDesejado & "' encontrado "
            If ContagemIrmaoFilho = 0 Then
                Msg = Msg & "entre os irmãos de '" & PathName & "', quando não deveria existir."
            Else
                Msg = Msg & "entre os filhos de '" & PathName & "', quando não deveria existir."
            End If
        End If

        If GerarCSV Then
            AdicionarErroExcel DadosExcel, PathName, CStr(TipoErro), Msg, AreaErro, TipoObjeto
        Else
            AdicionarErroBanco DadosExcel, PathName, CStr(TipoErro), Msg, TipoObjeto, AreaErro
        End If
    End If

    ContarObjetosDoTipo = contador

    On Error GoTo 0
End Function

'***********************************************************************
'*  Função : VerificarAssociacaoBase
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Verifica se a propriedade está associada a um objeto existente
'*     no domínio, validando o caminho retornado por GetPropertyLink.
'*
'*  Parâmetros :
'*     ‑ Obj          : Objeto a ser verificado.
'*     ‑ TipoObjeto   : TypeName ou rótulo do objeto.
'*     ‑ Propriedade  : Nome da propriedade a ser verificada.
'*     ‑ AreaErro     : Área onde a verificação está ocorrendo.
'*     ‑ TipoErro     : Severidade (0=Aviso, 1=Erro, 2=Revisar).
'*     ‑ CondicaoRetorno: Boolean que indica se a função deve retornar
'***********************************************************************
Function VerificarAssociacaoBase(Obj, TipoObjeto, Propriedade, AreaErro, TipoErro, CondicaoRetorno)
    On Error Resume Next

    Dim PathName, Link, Valor, ObjAssoc, Msg
    PathName = Obj.PathName

    Link  = GetPropertyLink(Obj, Propriedade)
    Valor = GetPropertyValue(Obj, Propriedade)

    ' Caso 1: Link preenchido e objeto não existe
    If Trim(Link) <> "" Then
        Set ObjAssoc = Application.GetObject(Link)
        If ObjAssoc Is Nothing Then
            If CondicaoRetorno Then
                VerificarAssociacaoBase = True
                Exit Function
            End If
            Msg = TipoObjeto & " com propriedade '" & Propriedade & _
                  "' associada ao link '" & Link & "', que não existe no domínio."
            If GerarCSV Then
                AdicionarErroExcel DadosExcel, PathName, CStr(TipoErro), Msg, AreaErro, TipoObjeto
            Else
                AdicionarErroBanco DadosExcel, PathName, CStr(TipoErro), Msg, TipoObjeto, AreaErro
            End If
        End If

    ' Caso 2: Valor direto e inválido
    ElseIf Not IsNumeric(Valor) And Not IsBoolean(Valor) And Trim(Valor) <> "" Then
        Set ObjAssoc = Application.GetObject(Valor)
        If ObjAssoc Is Nothing Then
            If CondicaoRetorno Then
                VerificarAssociacaoBase = True
                Exit Function
            End If
            Msg = TipoObjeto & " com valor '" & Valor & _
                  "' na propriedade '" & Propriedade & "', que não corresponde a nenhum objeto válido."
            If GerarCSV Then
                AdicionarErroExcel DadosExcel, PathName, CStr(TipoErro), Msg, AreaErro, TipoObjeto
            Else
                AdicionarErroBanco DadosExcel, PathName, CStr(TipoErro), Msg, TipoObjeto, AreaErro
            End If
        End If
    End If

    If CondicaoRetorno Then VerificarAssociacaoBase = False

    On Error GoTo 0
End Function

'***********************************************************************
'*  Função : HasChildOfType
'*----------------------------------------------------------------------
'*  Finalidade :
'*     Percorrer recursivamente a hierarquia de “Obj” e indicar se existe
'*     pelo menos um objeto cujo TypeName seja “TargetType”.
'*     A recursão só acontece para objetos cujo TypeName conste no array
'*     “ContainerTypes” (ex.: Folder, Screen, DrawGroup…).
'*
'*  Parâmetros :
'*     ‑ Obj            : Objeto raiz a partir do qual será feita a busca.
'*     ‑ TargetType     : String com o TypeName desejado (ex.: "WaterStationData").
'*     ‑ ContainerTypes : Array de strings contendo os tipos que podem
'*                         possuir filhos a serem varridos.
'*
'*  Retorno :
'*     Boolean → True  se encontrar “TargetType” em qualquer nível.
'*                False caso contrário.
'***********************************************************************
Function HasChildOfType(Obj, TargetType, ContainerTypes)

    On Error Resume Next

    Dim currentType
    currentType = TypeName(Obj)

    '------------------------------------------------------------------
    ' 1) Se o objeto atual já for do tipo procurado, retorna True
    '------------------------------------------------------------------
    If currentType = TargetType Then
        HasChildOfType = True
        Exit Function
    End If

    '------------------------------------------------------------------
    ' 2) Caso o objeto seja um “container”, percorre seus filhos
    '------------------------------------------------------------------
    Dim cType
    For Each cType In ContainerTypes

        If currentType = cType Then

            Dim childObj
            For Each childObj In Obj
                If HasChildOfType(childObj, TargetType, ContainerTypes) Then
                    HasChildOfType = True
                    Exit Function
                End If
            Next

            Exit For ' Já percorremos este container
        End If

    Next

    '------------------------------------------------------------------
    ' 3) Não encontrado
    '------------------------------------------------------------------
    HasChildOfType = False
    On Error GoTo 0

End Function

'***********************************************************************
'*  Função : VerificarUserFields
'*----------------------------------------------------------------------
'*  Finalidade :
'*     Para cada nome de userfield em “arrFields”, verificar se:
'*        a) o campo existe em Obj.UserFields;
'*        b) o valor não é vazio (Trim = "").
'*     Se qualquer condição falhar, registrar a inconsistência no
'*     relatório Excel ou banco de dados, conforme GerarCSV.
'*
'*  Parâmetros :
'*     ‑ Obj           : Objeto que contém a coleção UserFields.
'*     ‑ arrFields()   : Array de strings com os nomes a verificar.
'*     ‑ NomeObjeto    : Texto que aparecerá na coluna “Tipo” do relatório.
'*     ‑ AreaErro      : Área onde o erro foi encontrado.
'*     ‑ Classificacao : Código de severidade (0 = Aviso, 1 = Erro…).
'* 	   - CondicaoRetorno: Boolean que indica se a função deve retornar True
'***********************************************************************
Function VerificarUserFields(Obj, arrFields, NomeObjeto, AreaErro, Classificacao, CondicaoRetorno)

    On Error Resume Next

    Dim fieldName, fieldValue, mensagem

    For Each fieldName In arrFields

        fieldValue = Obj.UserFields.Item(fieldName)

        If Err.Number <> 0 Then
            If CondicaoRetorno = True Then
                VerificarUserFields = False
                Exit Function
            End If

            mensagem = NomeObjeto & " sem userfield '" & fieldName & "' (inexistente)."
            If GerarCSV Then
                AdicionarErroExcel DadosExcel, Obj.PathName, CStr(Classificacao), mensagem, AreaErro, NomeObjeto
            Else
                AdicionarErroBanco DadosExcel, Obj.PathName, CStr(Classificacao), mensagem, NomeObjeto, AreaErro
            End If
            Err.Clear

        ElseIf Trim(CStr(fieldValue)) = "" Then
            If CondicaoRetorno = True Then
                VerificarUserFields = False
                Exit Function
            End If

            mensagem = NomeObjeto & " userfield '" & fieldName & "' vazio."
            If GerarCSV Then
                AdicionarErroExcel DadosExcel, Obj.PathName, CStr(Classificacao), mensagem, AreaErro, NomeObjeto
            Else
                AdicionarErroBanco DadosExcel, Obj.PathName, CStr(Classificacao), mensagem, NomeObjeto, AreaErro
            End If
        End If

    Next

    If CondicaoRetorno = True Then
        VerificarUserFields = True
    End If

    On Error GoTo 0

End Function

'***********************************************************************
'*  Função : VerificarBibliotecasPorTypeName
'*----------------------------------------------------------------------
'*  Finalidade :
'*     Verificar as bibliotecas associadas a objetos irmãos em um mesmo
'*     DataServer e registrar inconsistências conforme o tipo de verificação.
'*
'*  Parâmetros :
'*     ‑ Obj              : Objeto base da análise.
'*     ‑ TipoObjeto       : Nome lógico do tipo de objeto.
'*     ‑ TipoCondicao     : Define a lógica da verificação:
'*                            0 = Erro se houver objetos de bibliotecas distintas.
'*                            1 = Erro se houver múltiplos objetos da mesma biblioteca.
'*                            2 = Retorna o nome da biblioteca do objeto (sem log).
'*     ‑ AreaErro         : Nome lógico da área para registro.
'*     ‑ TipoErro         : Severidade do erro (0=Aviso, 1=Erro, 2=Revisar).
'*     ‑ CondicaoRetorno  : Se True, retorna o resultado lógico ou texto, sem log.
'***********************************************************************
Function VerificarBibliotecasPorTypeName(Obj, TipoObjeto, TipoCondicao, AreaErro, TipoErro, CondicaoRetorno)
    On Error Resume Next

    Dim ParentObj, ObjFilho, tipo, lib, libAtual
    Dim libsEncontradas, PathName, Msg, countIgual

    Set ParentObj = Obj.Parent
    PathName = Obj.PathName
    libAtual = BibliotecaDoObjeto(TypeName(Obj))
    Set libsEncontradas = CreateObject("Scripting.Dictionary")
    countIgual = 0

    '--- Varrer objetos irmãos ---
    For Each ObjFilho In ParentObj
        tipo = TypeName(ObjFilho)
        lib = BibliotecaDoObjeto(tipo)

        If lib <> "Desconhecida" Then
            If Not libsEncontradas.Exists(lib) Then
                libsEncontradas.Add lib, True
            End If
        End If

        If lib = libAtual Then
            countIgual = countIgual + 1
        End If
    Next

    '--- Retorno lógico sem log ---
    If CondicaoRetorno Then
        Select Case TipoCondicao
            Case 0
                VerificarBibliotecasPorTypeName = (libsEncontradas.Count > 1)
            Case 1
                VerificarBibliotecasPorTypeName = (countIgual > 1)
            Case 2
                VerificarBibliotecasPorTypeName = libAtual
            Case Else
                VerificarBibliotecasPorTypeName = False
        End Select
        Exit Function
    End If

    '--- Registro: múltiplas bibliotecas ---
    If TipoCondicao = 0 And libsEncontradas.Count > 1 Then
        Msg = "A hierarquia '" & ParentObj.Name & "' possui objetos de múltiplas bibliotecas: "
        For Each lib In libsEncontradas.Keys
            Msg = Msg & lib & ", "
        Next
        Msg = Left(Msg, Len(Msg) - 2)

        If GerarCSV Then
            AdicionarErroExcel DadosExcel, PathName, CStr(TipoErro), Msg, AreaErro, TipoObjeto
        Else
            AdicionarErroBanco DadosExcel, PathName, CStr(TipoErro), Msg, TipoObjeto, AreaErro
        End If

    '--- Registro: duplicidade de biblioteca ---
    ElseIf TipoCondicao = 1 And countIgual > 1 Then
        Msg = "A hierarquia '" & ParentObj.Name & "' possui múltiplos objetos da biblioteca '" & libAtual & "' (além de '" & Obj.Name & "')."

        If GerarCSV Then
            AdicionarErroExcel DadosExcel, PathName, CStr(TipoErro), Msg, AreaErro, TipoObjeto
        Else
            AdicionarErroBanco DadosExcel, PathName, CStr(TipoErro), Msg, TipoObjeto, AreaErro
        End If
    End If

    On Error GoTo 0
End Function

'***********************************************************************
'*  Função : RegistrarTipoSemPropriedade
'*----------------------------------------------------------------------
'*  Finalidade :
'*     Registrar no log TXT cada TypeName que não possui regras de
'*     verificação específicas.  Cada tipo é gravado apenas uma vez,
'*     evitando linhas duplicadas no arquivo.
'*
'*  Parâmetros :
'*     ‑ TipoObjeto : String com o TypeName do objeto não tratado.
'*
'*  Retorno :
'*     Boolean → True  se o tipo foi registrado nesta chamada.
'*                False se já havia sido registrado anteriormente.
'***********************************************************************
Function RegistrarTipoSemPropriedade(TipoObjeto)

    On Error Resume Next

    '--------------------------------------------------------------
    ' Verifica se o tipo já foi registrado
    '--------------------------------------------------------------
    If Not TiposRegistrados.Exists(TipoObjeto) Then

        TiposRegistrados.Add TipoObjeto, True ' Marca como já visto

        AdicionarErroTxt DadosTxt, _
            "Objetos do tipo: ", _
            TipoObjeto, _
            "não são tratado ou não possuem propriedades a ser verificadas. "

        RegistrarTipoSemPropriedade = True ' Registrado agora
    Else
        RegistrarTipoSemPropriedade = False ' Já existia
    End If

    On Error GoTo 0

End Function

'***********************************************************************
'*  Função : VerificarPropriedadeCondicional
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Validar que a propriedade "Propriedade2" não esteja vazia quando a
'*     condição definida em "Propriedade1" for satisfeita.
'*
'*  Parâmetros :
'*     ‑ Obj            : Objeto a ser analisado.
'*     ‑ TipoObjeto       : TypeName do objeto.
'*     ‑ Propriedade1   : Nome da propriedade-condição.
'*     ‑ Propriedade2   : Nome da propriedade a ser verificada.
'*     ‑ TextoAuxiliar  : Valor que satisfaz a condição ou "NOTEMPTY".
'*     ‑ AreaErro       : Área onde o erro foi encontrado.
'*     ‑ TipoErro       : Tipo de erro (0 = Aviso, 1 = Erro, 2 = Revisar).
'*     ‑ CondicaoRetorno : Se True, retorna True se a condição for satisfeita e a propriedade estiver vazia.: 
'***********************************************************************
Function VerificarPropriedadeCondicional(Obj, TipoObjeto, PropriedadeCondicional, PropriedadeAlvo, TextoAuxiliar, AreaErro, TipoErro, CondicaoRetorno)
    On Error Resume Next

    Dim ValorCondicional, ValorAlvo, Mensagem
    ValorCondicional = GetPropriedade(Obj, PropriedadeCondicional)
    ValorAlvo = GetPropriedade(Obj, PropriedadeAlvo)

    ' Condição NOTEMPTY
    If UCase(CStr(TextoAuxiliar)) = "NOTEMPTY" Then
        If LenB(Trim(ValorAlvo)) = 0 Then
            If CondicaoRetorno = True Then
                VerificarPropriedadeCondicional = True
                Exit Function
            Else
                Mensagem = TipoObjeto & " com a propriedade " & PropriedadeAlvo & " vazia (condição NOTEMPTY)."
                If GerarCSV Then
                    AdicionarErroExcel DadosExcel, Obj.PathName, CStr(TipoErro), Mensagem, AreaErro, TipoObjeto
                Else
                    AdicionarErroBanco DadosExcel, Obj.PathName, CStr(TipoErro), Mensagem, TipoObjeto, AreaErro
                End If
            End If
        ElseIf CondicaoRetorno = True Then
            VerificarPropriedadeCondicional = False
        End If

    ' Condição explícita
    Else
        If CStr(ValorCondicional) = CStr(TextoAuxiliar) And LenB(Trim(ValorAlvo)) = 0 Then
            If CondicaoRetorno = True Then
                VerificarPropriedadeCondicional = True
                Exit Function
            Else
                Mensagem = TipoObjeto & " com a propriedade " & PropriedadeAlvo & _
                           " vazia enquanto " & PropriedadeCondicional & " está igual a " & CStr(TextoAuxiliar) & "."
                If GerarCSV Then
                    AdicionarErroExcel DadosExcel, Obj.PathName, CStr(TipoErro), Mensagem, AreaErro, TipoObjeto
                Else
                    AdicionarErroBanco DadosExcel, Obj.PathName, CStr(TipoErro), Mensagem, TipoObjeto, AreaErro
                End If
            End If
        ElseIf CondicaoRetorno = True Then
            VerificarPropriedadeCondicional = False
        End If
    End If

    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

'*********************************************************************** 
'*  Função : VerificarUnicidadeObjetoTipo
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     • Garante que apenas um objeto por tipo (ex.: DB.AlarmServer) exista.
'*     • Caso mais de um objeto de mesmo tipo seja detectado, registra erro.
'*  Parâmetros :
'*     Obj        → Objeto que está sendo analisado
'*     TipoObjeto → Nome do tipo do objeto (ex.: "DB.AlarmServer")
'*     AreaErro   → Rótulo de onde o erro foi detectado (ex.: "Banco de dados")
'*     TipoErro   → Severidade: 0=Aviso, 1=Erro, etc.
'***********************************************************************
Function VerificarUnicidadeObjetoTipo(Obj, TipoObjeto, AreaErro, TipoErro, ModoRetorno)
    On Error Resume Next

    Dim mensagem, achouDuplicidade
    achouDuplicidade = False

    If Not IsObject(Item("TiposUnicos")) Then
        Set Item("TiposUnicos") = CreateObject("Scripting.Dictionary")
    End If

    If Item("TiposUnicos").Exists(TipoObjeto) Then
        mensagem = "Já existe outro objeto do tipo '" & TipoObjeto & "' no domínio. Mais de um objeto deste tipo pode gerar conflitos de operação."
        achouDuplicidade = True

        If Not ModoRetorno Then
            If GerarCSV Then
                AdicionarErroExcel DadosExcel, Obj.PathName, CStr(TipoErro), mensagem, AreaErro, TipoObjeto
            Else
                AdicionarErroBanco DadosExcel, Obj.PathName, CStr(TipoErro), mensagem, TipoObjeto, AreaErro
            End If
        End If
    Else
        Item("TiposUnicos").Add TipoObjeto, Obj.PathName
    End If

    VerificarUnicidadeObjetoTipo = achouDuplicidade
    On Error GoTo 0
End Function

'***********************************************************************
'*  Sub‑rotina : VerificarCamposUsuariosServidorAlarmes
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     • Validar os campos de usuário (UserFields) configurado no objeto DB.AlarmServer.
'*     • Registrar inconsistências no relatório Excel ou banco.
'***********************************************************************
Sub VerificarCamposUsuariosServidorAlarmes(objServidor)

    On Error Resume Next

    Dim colUserFields, qtdeCampos, i
    Dim campoAtual, objUserFieldsPath, mensagem

    objUserFieldsPath = objServidor.PathName & ".UserFieldsCount"

    Dim objServidorUserFields
    Set objServidorUserFields = Application.GetObject(objUserFieldsPath)

    If objServidorUserFields Is Nothing Then
        mensagem = "Não foi possível obter referência ao Servidor de Alarmes (verifique se está rodando)."
        If GerarCSV Then
            AdicionarErroExcel DadosExcel, objUserFieldsPath, "1", mensagem, "Banco de dados", "DB.AlarmServer"
        Else
            AdicionarErroBanco DadosExcel, objUserFieldsPath, "1", mensagem, "DB.AlarmServer", "Banco de dados"
        End If
        Exit Sub
    End If

    Set colUserFields = objServidorUserFields.UserFields

    If colUserFields Is Nothing Then
        mensagem = "Não existe coleção cadastrada de campos de usuário no Servidor de Alarmes."
        If GerarCSV Then
            AdicionarErroExcel DadosExcel, objUserFieldsPath, "1", mensagem, "Banco de dados", "DB.AlarmServer"
        Else
            AdicionarErroBanco DadosExcel, objUserFieldsPath, "1", mensagem, "DB.AlarmServer", "Banco de dados"
        End If
        Exit Sub
    End If

    qtdeCampos = colUserFields.Count

    Dim camposExistentes
    Set camposExistentes = CreateObject("Scripting.Dictionary")

    Dim obrigatorios, opcionais, descontinuados

    obrigatorios = Array("SignalName", "SignalCaption", "AOR1", "AOR2", "AOR3", _
        "Categories", "DeviceType", "Hierarchy1", "Hierarchy2", _
        "Hierarchy3", "Hierarchies", "Screens", "Note", "FooterAlarmAreaID")

    opcionais = Array("SignalCaption2", "SignalCaption3", "ContainerGroup", _
        "Company", "Message2", "Message3", "Flags")

    descontinuados = Array("CriticalAlarm")

    For i = 1 To qtdeCampos
        campoAtual = colUserFields.Item(i).Name
        camposExistentes.Add campoAtual, True
    Next

    For Each campoAtual In obrigatorios
        If Not camposExistentes.Exists(campoAtual) Then
            mensagem = "Campo de usuário obrigatório faltando no Servidor de Alarmes: " & campoAtual
            If GerarCSV Then
                AdicionarErroExcel DadosExcel, objUserFieldsPath, "1", mensagem, "Banco de dados", "DB.AlarmServer"
            Else
                AdicionarErroBanco DadosExcel, objUserFieldsPath, "1", mensagem, "DB.AlarmServer", "Banco de dados"
            End If
        End If
    Next

    For Each campoAtual In descontinuados
        If camposExistentes.Exists(campoAtual) Then
            mensagem = "Campo de usuário não utilizado encontrado no Servidor de Alarmes (deve ser removido): " & campoAtual
            If GerarCSV Then
                AdicionarErroExcel DadosExcel, objUserFieldsPath, "0", mensagem, "Banco de dados", "DB.AlarmServer"
            Else
                AdicionarErroBanco DadosExcel, objUserFieldsPath, "0", mensagem, "DB.AlarmServer", "Banco de dados"
            End If
        End If
    Next

    Dim encontrado
    For Each campoAtual In camposExistentes.Keys
        encontrado = False

        If UBound(Filter(obrigatorios, campoAtual)) >= 0 Or _
                UBound(Filter(opcionais, campoAtual)) >= 0 Or _
                UBound(Filter(descontinuados, campoAtual)) >= 0 Then
            encontrado = True
        End If

        If Not encontrado Then
            mensagem = "Campo de usuário não previsto cadastrado no Servidor de Alarmes: " & campoAtual
            If GerarCSV Then
                AdicionarErroExcel DadosExcel, objUserFieldsPath, "0", mensagem, "Banco de dados", "DB.AlarmServer"
            Else
                AdicionarErroBanco DadosExcel, objUserFieldsPath, "0", mensagem, "DB.AlarmServer", "Banco de dados"
            End If
        End If
    Next

    On Error GoTo 0

End Sub

'***********************************************************************
'*  Função : VerificarBancoDeDados
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     • Ler o caminho do DBServer indicado em "Propriedade1".
'*     • Detectar reuso indevido do mesmo banco de dados por múltiplos
'*       objetos: o primeiro objeto que utilizar o BD é registrado em
'*       DadosBancoDeDados; se outro objeto apontar para o mesmo BD,
'*       gera-se log no Excel ou banco de dados.
'*
'*  Parâmetros :
'*     ‑ Obj            : Objeto a ser analisado.
'*     ‑ TipoObjeto     : TypeName do objeto.
'*     ‑ Propriedade    : Nome da propriedade que contém o DBServer.
'*     ‑ AreaErro       : Área onde o erro foi encontrado.
'*     ‑ TipoErro       : Tipo de erro (0 = Aviso, 1 = Erro, 2 = Revisar).
'* 	   - CondicaoRetorno: Se True, retorna True se o banco de dados for único, caso contrário retorna False.
'***********************************************************************
Function VerificarBancoDeDados(Obj, TipoObjeto, Propriedade, AreaErro, TipoErro, CondicaoRetorno)

    On Error Resume Next

    Dim ValorBD, PathName, Mensagem
    ValorBD = GetPropriedade(Obj, Propriedade)
    PathName = Obj.PathName

    '--------------------------------------------------------------
    ' Regra: se houver valor, verificar duplicidade
    '--------------------------------------------------------------
    If Trim(ValorBD) <> "" Then
        If Not DadosBancoDeDados.Exists(ValorBD) Then
            DadosBancoDeDados.Add ValorBD, PathName ' Primeiro uso
        Else
            If CondicaoRetorno = True Then
                VerificarBancoDeDados = False
                Exit Function
            End If

            Mensagem = TipoObjeto & " compartilhando o banco de dados '" & ValorBD & "' com o objeto " & DadosBancoDeDados(ValorBD) & "."
            If GerarCSV Then
                AdicionarErroExcel DadosExcel, PathName, CStr(TipoErro), Mensagem, AreaErro, TipoObjeto
            Else
                AdicionarErroBanco DadosExcel, PathName, CStr(TipoErro), Mensagem, TipoObjeto, AreaErro
            End If
        End If
    ElseIf CondicaoRetorno = True Then
        VerificarBancoDeDados = True
        Exit Function
    End If

    If Err.Number <> 0 Then
        AdicionarErroTxt DadosTxt, "VerificarBancoDeDados", Obj, _
            "Erro ao acessar " & Propriedade & " em " & TipoObjeto
        Err.Clear
    End If

    On Error GoTo 0

End Function

'***********************************************************************
'*  Função : VerificarPropriedadeHabilitada
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Garantir que a propriedade booleana/flag esteja no valor esperado
'*     (Enabled, Visible, ExposeToOpc, etc.). Se divergir, registrar no
'*     relatório Excel ou banco de dados.
'*
'*  Parâmetros :
'*     ‑ Obj            : Objeto a ser analisado.
'*     ‑ TipoObjeto     : TypeName do objeto.
'*     ‑ Propriedade    : Nome da propriedade a ser verificada.
'*     ‑ TextoAuxiliar  : Valor esperado para a propriedade (True/False ou número).
'*     ‑ AreaErro       : Área onde o erro foi encontrado (ex.: "Drivers").
'*     ‑ TipoErro       : Tipo de erro (0 = Aviso, 1 = Erro, 2 = Revisar).
'*     ‑ CondicaoRetorno: Se True, retorna True/False sem registrar.
'***********************************************************************
Function VerificarPropriedadeHabilitada(Obj, TipoObjeto, Propriedade, TextoAuxiliar, AreaErro, TipoErro, CondicaoRetorno)

    On Error Resume Next

    Dim ValorAtual, PathName, Mensagem
    ValorAtual = GetPropriedade(Obj, Propriedade)
    PathName = Obj.PathName

    If CStr(ValorAtual) <> CStr(TextoAuxiliar) Then
        Mensagem = TipoObjeto & " com a propriedade " & Propriedade & " diferente do esperado (Esperado: " & TextoAuxiliar & "; Atual: " & ValorAtual & ")."

        If CondicaoRetorno = True Then
            VerificarPropriedadeHabilitada = True
            Exit Function
        End If

        If GerarCSV Then
            AdicionarErroExcel DadosExcel, PathName, CStr(TipoErro), Mensagem, AreaErro, TipoObjeto
        Else
            AdicionarErroBanco DadosExcel, PathName, CStr(TipoErro), Mensagem, TipoObjeto, AreaErro
        End If
    Else
        If CondicaoRetorno = True Then
            VerificarPropriedadeHabilitada = False
        End If
    End If

    If Err.Number <> 0 Then
        AdicionarErroTxt DadosTxt, "VerificarPropriedadeHabilitada", Obj, _
            "Erro ao acessar " & Propriedade & " em " & TipoObjeto
        Err.Clear
    End If

    On Error GoTo 0

End Function

'***********************************************************************
'*  Função : VerificarPropriedadeValor
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Comparar o valor atual de uma propriedade com um valor esperado,
'*     gerando log conforme o modo de comparação.
'*
'*  Parâmetros :
'*     ‑ Obj            : Objeto a ser analisado.
'*     ‑ TipoObjeto     : TypeName do objeto.
'*     ‑ Propriedade    : Nome da propriedade a ser verificada.
'*     ‑ TextoAuxiliar  : Valor esperado para comparação.
'*     ‑ MetodoAuxiliar : Modo de comparação: 
'*                        0 = "igual" → log se "diferente" do esperado
'*                        1 = "diferente" → log se "igual" ao esperado
'*     ‑ AreaErro       : Área onde o erro foi encontrado.
'*     ‑ TipoErro       : Tipo de erro (0 = Aviso, 1 = Erro, 2 = Revisar).
'*     ‑ CondicaoRetorno: Se True, retorna True/False sem registrar.
'***********************************************************************
Function VerificarPropriedadeValor(Obj, TipoObjeto, Propriedade, TextoAuxiliar, MetodoAuxiliar, AreaErro, TipoErro, CondicaoRetorno)

    On Error Resume Next

    Dim ValorAtual, ValorAtualStr, ValorEsperadoStr, Mensagem, PathName
    ValorAtual = GetPropriedade(Obj, Propriedade)
    ValorAtualStr = CStr(ValorAtual)
    ValorEsperadoStr = CStr(TextoAuxiliar)
    PathName = Obj.PathName

    Select Case MetodoAuxiliar

        Case 0 ' comparação "igual" → log se for diferente
            If ValorAtualStr <> ValorEsperadoStr Then
                If CondicaoRetorno = True Then
                    VerificarPropriedadeValor = False
                    Exit Function
                End If

                Mensagem = TipoObjeto & " com a propriedade " & Propriedade & _
                           " diferente do valor esperado. (Esperado: " & ValorEsperadoStr & "; Atual: " & ValorAtualStr & ")"
                If GerarCSV Then
                    AdicionarErroExcel DadosExcel, PathName, CStr(TipoErro), Mensagem, AreaErro, TipoObjeto
                Else
                    AdicionarErroBanco DadosExcel, PathName, CStr(TipoErro), Mensagem, TipoObjeto, AreaErro
                End If
            ElseIf CondicaoRetorno = True Then
                VerificarPropriedadeValor = True
                Exit Function
            End If

        Case 1 ' comparação "diferente" → log se for igual
            If ValorAtualStr = ValorEsperadoStr Then
                If CondicaoRetorno = True Then
                    VerificarPropriedadeValor = False
                    Exit Function
                End If

                Mensagem = TipoObjeto & " com a propriedade " & Propriedade & _
                           " igual ao valor: " & ValorAtualStr & ", sendo o valor nativo do objeto."
                If GerarCSV Then
                    AdicionarErroExcel DadosExcel, PathName, CStr(TipoErro), Mensagem, AreaErro, TipoObjeto
                Else
                    AdicionarErroBanco DadosExcel, PathName, CStr(TipoErro), Mensagem, TipoObjeto, AreaErro
                End If
            ElseIf CondicaoRetorno = True Then
                VerificarPropriedadeValor = True
                Exit Function
            End If

        Case Else
            AdicionarErroTxt DadosTxt, "VerificarPropriedadeValor", Obj, _
                "MetodoAuxiliar inválido (" & MetodoAuxiliar & ") para propriedade " & Propriedade

    End Select

    If Err.Number <> 0 Then
        AdicionarErroTxt DadosTxt, "VerificarPropriedadeValor", Obj, _
            "Erro ao acessar " & Propriedade & " em " & TipoObjeto
        Err.Clear
    End If

    On Error GoTo 0
End Function

'*********************************************************************** 
'*  Função : VerificarObjetoDesatualizado
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Registrar no Excel que um objeto de biblioteca antiga deve ser
'*     substituído por outro de biblioteca atual.
'*
'*  Parâmetros :
'*     ‑ Obj            : Objeto a ser analisado.
'*     ‑ TipoObjeto     : TypeName do objeto.
'*     ‑ TextoAuxiliar  : Nome da biblioteca recomendada.
'*     ‑ AreaErro       : Área onde o erro foi encontrado.
'*     ‑ TipoErro       : Tipo de erro (0 = Aviso, 1 = Erro, 2 = Revisar).
'*     ‑ CondicaoRetorno: Se True, retorna True/False sem registrar.
'***********************************************************************
Function VerificarObjetoDesatualizado(Obj, TipoObjeto, TextoAuxiliar, AreaErro, TipoErro, CondicaoRetorno)

    On Error Resume Next

    Dim Mensagem, CaminhoObjeto
    CaminhoObjeto = Obj.PathName
    Mensagem = "O objeto " & TipoObjeto & " é obsoleto e deve ser substituído pela biblioteca " & TextoAuxiliar & "."

    If CondicaoRetorno = True Then
        VerificarObjetoDesatualizado = True
        Exit Function
    End If

    If GerarCSV Then
        AdicionarErroExcel DadosExcel, CaminhoObjeto, CStr(TipoErro), Mensagem, AreaErro, TipoObjeto
    Else
        AdicionarErroBanco DadosExcel, CaminhoObjeto, CStr(TipoErro), Mensagem, TipoObjeto, AreaErro
    End If

    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

End Function

'********************************************************************************
' Nome: VerificarPropriedadeTextoProibido
' Objetivo: Verificar se a propriedade (via link ou valor) contém um texto proibido.
'
' Parâmetros:
'   Obj            -> Objeto a verificar (ex.: pwa_Disjuntor, pwa_BarraAlarme)
'   TipoObjeto     -> Rótulo para o log (ex.: "pwa_Disjuntor")
'   Propriedade    -> Nome da propriedade (ex.: "SourceObject")
'   TextoProibido  -> Texto que não deve aparecer (ex.: ".Value")
'   Area           -> Área onde o erro foi encontrado (ex: "Telas", "Biblioteca")
'   Classificacao  -> Código de severidade no Excel (0=Aviso, 1=Erro, etc.)
'   CondicaoRetorno->  Se True, retorna True/False sem registrar.
'********************************************************************************
Function VerificarPropriedadeTextoProibido(Obj, TipoObjeto, Propriedade, TextoProibido, Area, Classificacao, CondicaoRetorno)
    On Error Resume Next

    Dim ValorAtual, mensagem
    ValorAtual = GetPropriedade(Obj, Propriedade)

    If InStr(1, ValorAtual, TextoProibido, vbTextCompare) > 0 Then
        If CondicaoRetorno = True Then
            VerificarPropriedadeTextoProibido = False
            Exit Function
        End If

        mensagem = "A propriedade " & Propriedade & " não deve conter o texto '" & TextoProibido & "'. " & _
                   "(Atual: " & ValorAtual & ")"

        If GerarCSV Then
            AdicionarErroExcel DadosExcel, Obj.PathName, CStr(Classificacao), mensagem, Area, TipoObjeto
        Else
            AdicionarErroBanco DadosExcel, Obj.PathName, CStr(Classificacao), mensagem, TipoObjeto, Area
        End If
    ElseIf CondicaoRetorno = True Then
        VerificarPropriedadeTextoProibido = True
        Exit Function
    End If

    If Err.Number <> 0 Then
        AdicionarErroTxt DadosTxt, "VerificarPropriedadeTextoProibido", Obj, _
            "Erro ao acessar " & Propriedade & " em " & TipoObjeto
        Err.Clear
    End If

    On Error GoTo 0
End Function

'***********************************************************************
'*  Função : VerificarObjetoInternoIndevido
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Registrar que determinado objeto é destinado a uso interno e
'*     não deve estar presente em uma aplicação entregue ao cliente.
'*
'*  Parâmetros :
'*     ‑ Obj            : Objeto a ser analisado.
'*     ‑ TipoObjeto     : TypeName do objeto.
'*     ‑ AreaErro       : Área onde o erro foi detectado (ex: Telas, Biblioteca...).
'*     ‑ TipoErro       : Severidade (0 = Aviso, 1 = Erro, 2 = Revisar).
'*     ‑ CondicaoRetorno: Se True, retorna True/False sem registrar.
'***********************************************************************
Function VerificarObjetoInternoIndevido(Obj, TipoObjeto, AreaErro, TipoErro, CondicaoRetorno)
    On Error Resume Next

    Dim mensagem, caminhoObjeto
    caminhoObjeto = Obj.PathName
    mensagem = "O objeto " & TipoObjeto & " é de uso interno e não deve estar presente na aplicação final entregue ao cliente."

    If CondicaoRetorno = True Then
        VerificarObjetoInternoIndevido = True  ' Encontrou objeto indevido → retorna True
        Exit Function
    End If

    ' Registro se não estiver no modo de retorno lógico
    If GerarCSV Then
        AdicionarErroExcel DadosExcel, caminhoObjeto, CStr(TipoErro), mensagem, AreaErro, TipoObjeto
    Else
        AdicionarErroBanco DadosExcel, caminhoObjeto, CStr(TipoErro), mensagem, TipoObjeto, AreaErro
    End If

    VerificarObjetoInternoIndevido = False  ' Registro feito → mas não está no modo de retorno lógico
    On Error GoTo 0
End Function


'***********************************************************************
'*  Função : VerificarTipoPai
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Verificar se o pai de determinado objeto é ou não do tipo desejado,
'*     e registrar uma inconsistência caso a regra de posicionamento não
'*     esteja sendo respeitada.
'*
'*  Parâmetros :
'*     ‑ Obj            : Objeto a ser verificado.
'*     ‑ TipoObjeto       : TypeName do objeto.
'*     ‑ TipoEsperado   : Tipo que o pai do objeto deveria (ou não) ser.
'*     ‑ Regra          : Define a regra de verificação:
'*                        0 = O pai NÃO pode ser do tipo indicado.
'*                        1 = O pai DEVE ser do tipo indicado.
'*     ‑ AreaErro       : Área onde o erro foi identificado.
'*     ‑ TipoErro       : Severidade do erro (0 = Aviso, 1 = Erro, 2 = Revisar).
'*     ‑ CondicaoRetorno: Se True, retorna True/False sem registrar.
'***********************************************************************
Function VerificarTipoPai(Obj, TipoObjeto, TipoEsperado, Regra, AreaErro, TipoErro, CondicaoRetorno)
    On Error Resume Next

    Dim Pai, TipoDoPai, Msg, PathName
    Set Pai = Obj.Parent
    PathName = Obj.PathName

    If Pai Is Nothing Then
        TipoDoPai = "[pai nulo]"
    ElseIf IsObject(Pai) Then
        TipoDoPai = TypeName(Pai)
        If TipoDoPai = "" Then TipoDoPai = "[tipo indefinido]"
    Else
        TipoDoPai = "[tipo inválido]"
    End If

    '-----------------------------------------------------
    ' Regra 0: Pai não deve ser do tipo informado
    '-----------------------------------------------------
    If Regra = 0 Then
        If StrComp(TipoDoPai, TipoEsperado, vbTextCompare) = 0 Then
            If CondicaoRetorno Then
                VerificarTipoPai = True
                Exit Function
            End If
            Msg = TipoObjeto & " está localizado dentro de um objeto do tipo '" & TipoEsperado & "', o que não é permitido."
            If GerarCSV Then
                AdicionarErroExcel DadosExcel, PathName, CStr(TipoErro), Msg, AreaErro, TipoObjeto
            Else
                AdicionarErroBanco DadosExcel, PathName, CStr(TipoErro), Msg, TipoObjeto, AreaErro
            End If
        End If

    '-----------------------------------------------------
    ' Regra 1: Pai deve ser do tipo informado
    '-----------------------------------------------------
    ElseIf Regra = 1 Then
        If StrComp(TipoDoPai, TipoEsperado, vbTextCompare) <> 0 Then
            If CondicaoRetorno Then
                VerificarTipoPai = True
                Exit Function
            End If
            Msg = TipoObjeto & " deveria estar contido em um objeto do tipo '" & TipoEsperado & "', mas está dentro de '" & TipoDoPai & "'."
            If GerarCSV Then
                AdicionarErroExcel DadosExcel, PathName, CStr(TipoErro), Msg, AreaErro, TipoObjeto
            Else
                AdicionarErroBanco DadosExcel, PathName, CStr(TipoErro), Msg, TipoObjeto, AreaErro
            End If
        End If
    End If

    If CondicaoRetorno Then VerificarTipoPai = False
    On Error GoTo 0
End Function

'***********************************************************************
'*  Função : AdicionarErroTxt
'*----------------------------------------------------------------------
'*  Finalidade :
'*     Inserir linha de erro no dicionário DadosTxt para posterior
'*     geração de log TXT.  Cada entrada recebe uma chave numérica
'*     incremental (string).
'*
'*  Parâmetros :
'*     ‑ DadosTxt      : Dicionário global onde as linhas são gravadas.
'*     ‑ NomeSub       : Nome da sub/função que originou o erro.
'*     ‑ Obj           : Objeto ou string associada ao erro.
'*     ‑ DescricaoErro : Texto explicativo.
'***********************************************************************
Function AdicionarErroTxt(DadosTxt, NomeSub, Obj, DescricaoErro)

    On Error Resume Next

    '--------------------------------------------------------------
    ' 1) Gera chave única incremental
    '--------------------------------------------------------------
    Dim LinhaTxt, keyTxt
    LinhaTxt = DadosTxt.Count + 1
    keyTxt = CStr(LinhaTxt)

    While DadosTxt.Exists(keyTxt)
        LinhaTxt = LinhaTxt + 1
        keyTxt = CStr(LinhaTxt)
    Wend

    '--------------------------------------------------------------
    ' 2) Valida dicionário
    '--------------------------------------------------------------
    If Not IsObject(DadosTxt) Then
        MsgBox "Erro: O dicionário DadosTxt não foi inicializado.", vbCritical
        Exit Function
    End If

    '--------------------------------------------------------------
    ' 3) Obtém PathName do objeto (se for objeto)
    '--------------------------------------------------------------
    Dim ObjPath
    If IsObject(Obj) Then
        ObjPath = Obj.PathName
        If Err.Number <> 0 Then
            ObjPath = "[Sem PathName]"
            Err.Clear
        End If
    Else
        ObjPath = Obj ' Se já for string, usa diretamente
    End If

    '--------------------------------------------------------------
    ' 4) Monta e grava mensagem
    '--------------------------------------------------------------
    Dim MensagemErro
    MensagemErro = "Erro na Sub " & NomeSub & "/" & ObjPath & ": " & DescricaoErro

    DadosTxt.Add keyTxt, MensagemErro

    On Error GoTo 0

End Function

'***********************************************************************
'*  Função : GerarRelatorioTxt
'*----------------------------------------------------------------------
'*  Finalidade :
'*     Criar um arquivo TXT contendo todas as mensagens armazenadas no
'*     dicionário DadosTxt.  Ao final, pergunta ao usuário se deseja
'*     abrir o arquivo recém‑gerado.
'*
'*  Parâmetros :
'*     ‑ DadosTxt  : Dicionário onde cada item é uma linha de log.
'*     ‑ CaminhoPrj: Pasta onde o arquivo será salvo.
'*
'*  Retorno :
'*     Boolean → True  se o arquivo foi gerado com sucesso.
'*                False caso não haja dados ou ocorra falha.
'***********************************************************************
Function GerarRelatorioTxt(DadosTxt, CaminhoPrj)

    On Error Resume Next

    '--------------------------------------------------------------
    ' 1) Valida se há conteúdo
    '--------------------------------------------------------------
    If DadosTxt.Count = 0 Then
        MsgBox "Nenhum dado disponível para gerar o relatório TXT.", vbExclamation
        GerarRelatorioTxt = False
        Exit Function
    End If

    '--------------------------------------------------------------
    ' 2) Define nome do arquivo
    '--------------------------------------------------------------
    Dim NomeTxt
    NomeTxt = CaminhoPrj & "\Log_" & _
        Replace(Replace(Date() & "_" & Time(), ":", "_"), "/", "_") & ".txt"

    '--------------------------------------------------------------
    ' 3) Cria arquivo e grava linhas
    '--------------------------------------------------------------
    Dim FSO, ArquivoTxt, Linha
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ArquivoTxt = FSO.CreateTextFile(NomeTxt, True)

    For Each Linha In DadosTxt
        If Trim(DadosTxt.Item(Linha)) <> "" Then
            ArquivoTxt.WriteLine DadosTxt.Item(Linha)
        End If
    Next
    ArquivoTxt.Close

    '--------------------------------------------------------------
    ' 4) Pergunta se deve abrir o arquivo
    '--------------------------------------------------------------
    Dim Resposta, ShellObj
    Resposta = MsgBox("Foram gerados logs de erro de código. Deseja abrir o arquivo?", _
        vbYesNo + vbQuestion, "AutoTester")
    If Resposta = vbYes Then
        Set ShellObj = CreateObject("WScript.Shell")
        ShellObj.Run """" & NomeTxt & """"
        Set ShellObj = Nothing
    End If

    GerarRelatorioTxt = True
    On Error GoTo 0

End Function

'***********************************************************************
'*  Função : AdicionarErroExcel
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Adicionar uma entrada ao dicionário DadosExcel no formato:
'*        key → "PathName/TipoErro/Descricao/TypeName/Area"
'*     onde:
'*        • PathName   = Caminho do objeto
'*        • TipoErro   = Código numérico (0,1,2…)
'*        • Descricao  = Texto explicativo do problema
'*        • TypeName   = Tipo do objeto
'*        • Area       = Área onde ocorreu a inconsistência
'*
'*  Parâmetros :
'*     ‑ DadosExcel       : Dicionário global para o Excel.
'*     ‑ CaminhoObjeto    : PathName do objeto ou texto livre.
'*     ‑ ClassificacaoCode: "0","1","2"… (Aviso, Erro, Revisar…).
'*     ‑ Mensagem         : Descrição do problema.
'*     ‑ Area             : Área da inconsistência (ex.: "Drivers").
'*     ‑ TipoObjeto         : Nome do tipo de objeto.
'***********************************************************************
Function AdicionarErroExcel(DadosExcel, CaminhoObjeto, ClassificacaoCode, Mensagem, Area, TipoObjeto)

    On Error Resume Next

    If Not IsObject(DadosExcel) Then
        MsgBox "Erro: O dicionário DadosExcel não foi inicializado.", vbCritical
        Exit Function
    End If

    Dim LinhaExcel, keyExcel
    LinhaExcel = DadosExcel.Count + 1
    keyExcel = CStr(LinhaExcel)

    ' Evita colisões de chave
    While DadosExcel.Exists(keyExcel)
        LinhaExcel = LinhaExcel + 1
        keyExcel = CStr(LinhaExcel)
    Wend

    If Len(Trim(CaminhoObjeto)) > 0 And Len(Trim(ClassificacaoCode)) > 0 And Len(Trim(Mensagem)) > 0 Then
        DadosExcel.Add keyExcel, _
            CaminhoObjeto & "/" & _
            ClassificacaoCode & "/" & _
            Mensagem & "/" & _
            TipoObjeto & "/" & _
            Area
    Else
        MsgBox "Erro: Valores inválidos ao adicionar ao Excel:" & vbCrLf & _
            "CaminhoObjeto: " & CaminhoObjeto & vbCrLf & _
            "ClassificacaoCode: " & ClassificacaoCode & vbCrLf & _
            "Mensagem: " & Mensagem & vbCrLf & _
            "TypeName: " & TipoObjeto & vbCrLf & _
            "Area: " & Area, vbCritical
    End If

    On Error GoTo 0

End Function

'***********************************************************************
'*  Função : GerarRelatorioExcel
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Exportar o conteúdo de DadosExcel para um arquivo .xlsx com os
'*     seguintes campos: Empreendimento, Projeto, Localidade,
'*     ResponsavelQA, PathName, TypeName, Categoria, Area, Descricao.
'*
'*  Parâmetros :
'*     ‑ DadosExcel : Dicionário onde cada item segue o padrão
'*                    "PathName/TipoErro/Descricao/TypeName/Area".
'*     ‑ CaminhoPrj : Pasta onde o arquivo será salvo.
'*
'*  Retorno :
'*     Boolean → True  se o arquivo foi salvo com sucesso.
'*                False se não houver dados ou ocorrer erro.
'***********************************************************************
Function GerarRelatorioExcel(DadosExcel, CaminhoPrj)

    On Error Resume Next

    '--------------------------------------------------------------
    ' 1) Verifica se há dados
    '--------------------------------------------------------------
    If DadosExcel.Count = 0 Then
        MsgBox "Nenhum dado disponível para gerar o relatório Excel.", vbExclamation
        GerarRelatorioExcel = False
        Exit Function
    End If

    '--------------------------------------------------------------
    ' 2) Define nome do arquivo
    '--------------------------------------------------------------
    Dim NomeExcel
    NomeExcel = CaminhoPrj & "\RelatorioTester_" & _
        Replace(Replace(Date() & "_" & Time(), ":", "_"), "/", "_") & ".xlsx"

    '--------------------------------------------------------------
    ' 3) Cria planilha e cabeçalho
    '--------------------------------------------------------------
    Dim objExcel, objWorkBook, sheet, Linha, campos
    Set objExcel = CreateObject("EXCEL.APPLICATION")
    Set objWorkBook = objExcel.Workbooks.Add
    Set sheet = objWorkBook.Sheets(1)

    sheet.Cells(1, 1).Value = "Empreendimento"
    sheet.Cells(1, 2).Value = "Projeto"
    sheet.Cells(1, 3).Value = "Localidade"
    sheet.Cells(1, 4).Value = "ResponsavelQA"
    sheet.Cells(1, 5).Value = "PathName"
    sheet.Cells(1, 6).Value = "TypeName"
    sheet.Cells(1, 7).Value = "Categoria"
    sheet.Cells(1, 8).Value = "Area"
    sheet.Cells(1, 9).Value = "Descricao"
    sheet.Rows(1).Font.Bold = True

    '--------------------------------------------------------------
    ' 4) Preenche os dados
    '--------------------------------------------------------------
    Dim linhaIndex, classificationText
    linhaIndex = 2

    For Each Linha In DadosExcel

        campos = Split(DadosExcel.Item(Linha), "/")
        If UBound(campos) >= 4 Then

            Select Case campos(1)
                Case "0" : classificationText = "Aviso"
                Case "1" : classificationText = "Erro"
                Case "2" : classificationText = "Revisar"
                Case Else : classificationText = "Desconhecido"
            End Select

            sheet.Cells(linhaIndex, 1).Value = Empreendimento
            sheet.Cells(linhaIndex, 2).Value = Projeto

            If Trim(Localidade) = "" Then
                sheet.Cells(linhaIndex, 3).Value = Null
            Else
                sheet.Cells(linhaIndex, 3).Value = Localidade
            End If

            If Trim(ResponsavelQA) = "" Then
                sheet.Cells(linhaIndex, 4).Value = "AutoTester"
            Else
                sheet.Cells(linhaIndex, 4).Value = ResponsavelQA
            End If

            sheet.Cells(linhaIndex, 5).Value = campos(0) ' PathName
            sheet.Cells(linhaIndex, 6).Value = campos(3) ' TypeName
            sheet.Cells(linhaIndex, 7).Value = classificationText ' Categoria
            sheet.Cells(linhaIndex, 8).Value = campos(4) ' Area
            sheet.Cells(linhaIndex, 9).Value = campos(2) ' Descricao

            linhaIndex = linhaIndex + 1
        End If
    Next

    '--------------------------------------------------------------
    ' 5) Salva e encerra o Excel
    '--------------------------------------------------------------
    objWorkBook.SaveAs NomeExcel
    objWorkBook.Close False
    objExcel.Quit

    Set sheet = Nothing
    Set objWorkBook = Nothing
    Set objExcel = Nothing

    '--------------------------------------------------------------
    ' 6) Pergunta ao usuário se deseja abrir o arquivo
    '--------------------------------------------------------------
    Dim Resposta, ShellObj
    Resposta = MsgBox("Foram gerados logs de correção. Deseja abrir o arquivo?", _
        vbYesNo + vbQuestion, "AutoTester")
    If Resposta = vbYes Then
        Set ShellObj = CreateObject("WScript.Shell")
        ShellObj.Run """" & NomeExcel & """"
        Set ShellObj = Nothing
    End If

    GerarRelatorioExcel = True
    On Error GoTo 0

End Function

Function TestarPing(ip)
    Dim shell, resultado
    Set shell = CreateObject("WScript.Shell")
    resultado = shell.Run("ping -n 1 -w 1000 " & ip, 0, True) ' o terceiro argumento True força espera
    TestarPing = (resultado = 0)
    Set shell = Nothing
End Function

Function ConectarBancoQA()
    On Error Resume Next

    'MsgBox "DEBUG: Entrou na função ConectarBancoQA", vbInformation

    Dim conn, strConexao
    Set conn = CreateObject("ADODB.Connection")

    Dim servidor, banco, usuario, senha
    servidor = "192.168.20.12"
    banco = "QA"
    usuario = "sa"
    senha = "1234"

    strConexao = "Provider=SQLOLEDB;Data Source=" & servidor & ";Initial Catalog=" & banco & ";User ID=" & usuario & ";Password=" & senha & ";"

    conn.Open strConexao

    If conn.State = 1 Then
        'MsgBox "DEBUG: Conexão bem-sucedida em ConectarBancoQA", vbInformation
        Set ConectarBancoQA = conn
    Else
        'MsgBox "DEBUG: Falha na conexão em ConectarBancoQA", vbCritical
        Set ConectarBancoQA = Nothing
    End If

    On Error GoTo 0
End Function

'***********************************************************************
'*  Função : AdicionarErroBanco
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Adiciona uma inconsistência ao dicionário DadosExcel, formatada
'*     no padrão esperado para inserção posterior no banco de dados.
'*
'*     Cada entrada será estruturada como:
'*        key   → número sequencial (1, 2, 3…)
'*        value → "PathName/TipoErro/Descricao/TypeName/Area"
'*
'*  Parâmetros :
'*     ‑ DadosExcel : Dicionário global de inconsistências.
'*     ‑ PathName   : Caminho do objeto (Obj.PathName).
'*     ‑ TipoErro   : Classificação numérica: "0", "1", "2", etc.
'*     ‑ Descricao  : Texto descrevendo o problema.
'*     ‑ TipoObjeto   : Tipo do objeto.
'*     ‑ Area       : Área da inconsistência (ex.: "Drivers", "Telas").
'***********************************************************************
Function AdicionarErroBanco(DadosExcel, PathName, TipoErro, Descricao, TipoObjeto, Area)
    On Error Resume Next

    ' MsgBox "DEBUG: Entrou em AdicionarErroBanco" & vbCrLf & _
    '        "PathName: " & PathName & vbCrLf & _
    '        "TipoErro: " & TipoErro & vbCrLf & _
    '        "Descricao: " & Descricao & vbCrLf & _
    '        "TypeName: " & TipoObjeto & vbCrLf & _
    '        "Area: " & Area, vbInformation

    Dim LinhaBanco, keyBanco
    LinhaBanco = DadosExcel.Count + 1
    keyBanco = CStr(LinhaBanco)

    While DadosExcel.Exists(keyBanco)
        LinhaBanco = LinhaBanco + 1
        keyBanco = CStr(LinhaBanco)
    Wend

    If Not IsObject(DadosExcel) Then
        'MsgBox "Erro: O dicionário DadosExcel não foi inicializado.", vbCritical
        Exit Function
    End If

    If Len(Trim(PathName)) > 0 And Len(Trim(TipoErro)) > 0 And Len(Trim(Descricao)) > 0 Then
        DadosExcel.Add keyBanco, PathName & "/" & TipoErro & "/" & Descricao & "/" & TipoObjeto & "/" & Area
    Else
        'MsgBox "Erro ao adicionar erro ao banco (valores ausentes).", vbCritical
    End If

    On Error GoTo 0
End Function

'***********************************************************************
'*  Função : InserirInconsistenciasBanco
'*---------------------------------------------------------------------- 
'*  Finalidade :
'*     Exportar as inconsistências do dicionário DadosExcel para o
'*     banco de dados SQL conforme a estrutura da tabela AutoTester.
'*
'*  Parâmetros :
'*     ‑ DadosExcel : Dicionário com dados formatados como:
'*                    "PathName/TipoErro/Descricao/TypeName/Area"
'*     ‑ conn       : Objeto de conexão com o banco de dados.
'***********************************************************************
Function InserirInconsistenciasBanco(DadosExcel, conn)
    On Error Resume Next

    'MsgBox "DEBUG: Entrou em InserirInconsistenciasBanco - Total de registros: " & DadosExcel.Count, vbInformation

    Dim linha, campos, loteSQL
    Dim PathName, TipoErro, Descricao, TipoObjeto, Area, Categoria
    Dim LocalidadeFinal, insertCount
    insertCount = 0
    loteSQL = ""

    For Each linha In DadosExcel
	'MsgBox "DEBUG: Processando chave: " & linha, vbInformation

        campos = Split(DadosExcel.Item(linha), "/")
        If UBound(campos) >= 4 Then

            PathName = campos(0)
            TipoErro = campos(1)
            Descricao = campos(2)
            TipoObjeto = campos(3)
            Area = campos(4)

            Select Case TipoErro
                Case "0": Categoria = "Aviso"
                Case "1": Categoria = "Erro"
                Case "2": Categoria = "Revisar"
                Case Else: Categoria = "Desconhecido"
            End Select

            If Trim(Localidade) = "" Then
                LocalidadeFinal = "NULL"
            Else
                LocalidadeFinal = "'" & Replace(Localidade, "'", "''") & "'"
            End If

            Dim SQLLinha
            SQLLinha = "INSERT INTO AutoTester " & _
                       "(Empreendimento, Projeto, Localidade, Responsavel_QA, PathName, TypeName, Categoria, Area, Descricao) " & _
                       "VALUES (" & _
                       "'" & Replace(Empreendimento, "'", "''") & "', " & _
                       "'" & Replace(Projeto, "'", "''") & "', " & _
                       LocalidadeFinal & ", " & _
                       "'" & Replace(IIf(ResponsavelQA = "", "AutoTester", ResponsavelQA), "'", "''") & "', " & _
                       "'" & Replace(PathName, "'", "''") & "', " & _
                       "'" & Replace(TipoObjeto, "'", "''") & "', " & _
                       "'" & Replace(Categoria, "'", "''") & "', " & _
                       "'" & Replace(Area, "'", "''") & "', " & _
                       "'" & Replace(Descricao, "'", "''") & "');"

            loteSQL = loteSQL & vbCrLf & SQLLinha
            insertCount = insertCount + 1

            ' A cada 100 registros, executa o lote
            If insertCount >= 100 Then
                'MsgBox "DEBUG: Enviando lote de 100 registros ao banco", vbInformation
                conn.Execute loteSQL
                loteSQL = ""
                insertCount = 0
            End If
        End If
    Next

    ' Executa o restante se sobrou menos de 100
    If insertCount > 0 And loteSQL <> "" Then
        'MsgBox "DEBUG: Enviando lote final ao banco", vbInformation
        conn.Execute loteSQL
    End If

    On Error GoTo 0
End Function

Function BibliotecaDoObjeto(tipo)
    Select Case tipo
        case "gx_StatusIndicator", "gx_QualityIcon", "gx_OpenTabularArea", "gx_Notes", "gx_Note_Tag", "gx_Note_Out", "gx_Note_Fork", "gx_Menu", "gx_LocalTeam", "gx_InfoSetpoint", "gx_InfoDoughnutChart", "gx_InfoAnalogic2", "gx_InfoAnalogic", "gx_InfoAlarm10", "gx_InfoAlarm05", "gx_InfoAlarm03", "gx_InfoAlarm01", "gx_DiscreteIndicator", "gx_CtrlPulse", "gx_CtrlDigital4Op", "gx_CtrlDigital3Op", "gx_CtrlDigital2Op", "gx_CtrlDigital1Op", "gx_CtrlDigital", "gx_Counter", "gx_ButtonOpenCommandScreen", "gx_Analogic", "gx_AbnormalityIndicator", "gx_RadarChartInfo", "gx_RadarChart20", "gx_RadarChart16", "gx_RadarChart12", "gx_RadarChart10_2Z", "gx_RadarChart10", "gx_RadarChart09", "gx_RadarChart08_2Z", "gx_RadarChart08", "gx_RadarChart07", "gx_RadarChart06", "gx_RadarChart05", "gx_RadarChart04", "gx_RadarChart03"
            BibliotecaDoObjeto = "generic_automalogica"
        Case "pwa_GrupoVSL", "pwa_HomeButton", "pwa_VentForc", "pwa_Terra2", "pwa_Terra", "pwa_TapV", "pwa_Sensor", "pwa_Retificador", "pwa_Relig", "pwa_ReguladorTensao", "pwa_Reactor", "pwa_Menu", "pwa_Jumper", "pwa_Inversor", "pwa_InfoPotRea", "pwa_InfoAnalogicaG", "pwa_InfoAnalogica", "pwa_InfoAlarme10", "pwa_InfoAlarme05", "pwa_InfoAlarme01", "pwa_GeradorG", "pwa_Gerador", "pwa_Conexao", "pwa_Carga", "pwa_Capacitor", "pwa_BotaoAbreTela", "pwa_Bateria", "pwa_BarraAlarme", "pwa_Barra2Vert", "pwa_Barra2", "pwa_Barra", "pwa_AutoTrafo", "pwa_InfoPotP", "pwa_InfoPotG", "pwa_InfoPot", "pwa_LineVert", "pwa_LineHoriz", "pwa_TrafoSA", "pwa_Trafo3Type01", "pwa_Trafo3_P", "pwa_Trafo3", "pwa_Trafo2Term", "pwa_Trafo2", "pwa_Seccionadora", "pwa_DisjuntorPP", "pwa_DisjuntorP", "pwa_Disjuntor"
            BibliotecaDoObjeto = "poweratm_xc"
        Case "archAeroGenerator" ,"archChannelPanelPP" ,"archChannelPanelP" ,"archChannelPanel" ,"archChannelPanelG" ,"archCloud" ,"archDatabase" ,"archElectricalMeter" ,"archFirewall" ,"archGPSAntenna" ,"archIED" ,"archInfo" ,"archInfoLine" ,"archLineHorizontal" ,"archLineVertical" ,"archModuloIO" ,"archNuclearPlant" ,"archPCH" ,"archRouter" ,"archRTU" ,"archServerDesktop" ,"archWifi" ,"archViewer" ,"archVideoWall" ,"archUHE" ,"archSwitch" ,"archSurveillanceCamera" ,"archSubtitle2" ,"archSubtitle1" ,"archSolarPanel" ,"archServerRackmountSingle" ,"archServerRackmountMultiple"
            BibliotecaDoObjeto = "Architecture_xc"
        Case "uhe_Filter", "uhe_FanIndicator", "uhe_Fan", "uhe_ExcitationTransformer", "uhe_Etiqueta_Manut", "uhe_Etiqueta", "uhe_EqpMenuCreator", "uhe_EarthSwitch", "uhe_DiscreteIndicator", "uhe_Direction", "uhe_DieselGenerator", "uhe_Device", "uhe_CtrlPulse", "uhe_CtrlDigitalOp", "uhe_CtrlDigital4Op", "uhe_CtrlDigital3Op", "uhe_CtrlDigital2Op", "uhe_CtrlDigital", "uhe_ControlGate", "uhe_Conduto", "uhe_CommandIcon", "uhe_CommandButton", "uhe_Command", "uhe_ChartP", "uhe_Chart", "uhe_Caixa", "uhe_Buzzer", "uhe_BusBar", "uhe_BulbTurbine", "uhe_BrakeAlert", "uhe_Bomb2", "uhe_Bomb", "uhe_Block", "uhe_BielaMecanica", "uhe_BielaHidraulica", "uhe_AnalogIndicator", "uhe_AnalogBarP", "uhe_AnalogBarHor", "uhe_AnalogBar5LimitsH", "uhe_AnalogBar5Limits", "uhe_AnalogBar", "uhe_AlarmBar", "uhe_AirOilTank", "uhe_AirCompressor", "uhe_AbnormalityIndicator", "uhe_StartStopSquenceTab", "uhe_ValveManual", "uhe_ValveDistributing", "uhe_ValveByPass", "uhe_ValveButterfly", "uhe_Valve4Ways", "uhe_Valve3Ways", "uhe_Valve", "uhe_UgHeatResistence", "uhe_TrendArrow", "uhe_StatusIndicator", "uhe_StartStopSequenceTxt", "uhe_StartStopButton", "uhe_StartStop", "uhe_SpillwayP", "uhe_Spillway", "uhe_Sapata", "uhe_Rectifier", "uhe_PressureSwitch", "uhe_PresDiferencial", "uhe_PowerFlowDirection", "uhe_OilTank", "uhe_NivelDrenagem", "uhe_MeasBar", "uhe_Lock", "uhe_LineVert", "uhe_LineHoriz", "uhe_KaplanTurbine", "uhe_Inverter", "uhe_InfoSetpoint", "uhe_InfoPotRea", "uhe_InfoPotP", "uhe_InfoAnalogic2", "uhe_InfoAnalogic", "uhe_InfoAlarm10", "uhe_InfoAlarm05", "uhe_InfoAlarm03", "uhe_InfoAlarm01", "uhe_GeneratorPP", "uhe_GeneratorP", "uhe_Generator", "uhe_FrancisTurbine", "uhe_FilterSelfCleaning", "uhe_ZoomObject", "uhe_ZoomFrame", "uhe_WaterTank", "uhe_ValveSolenoid"
            BibliotecaDoObjeto = "uhe_automalogica"
        Case "frCustomAlarmAndEventConfig" ,"fCustomAppConfig" ,"frCustomGmanobConfig" ,"frCustomNotesConfig" ,"frDBMonitor" ,"frFindApplicationFrames" ,"frFooterButton" ,"frFooterMaintenance" ,"frFooterRoot" ,"frFooterSeparator" ,"frFooterTab" ,"frHDMonitor" ,"frInsertFocusScriptOnScreens" ,"frLibVersion" ,"frLibVersionXO" ,"frRetentiveValue" ,"frThemeColorProperty" ,"frThemeConfig" ,"frThemeFileProperty" ,"frThemeFolder" ,"frThemeMaintenance" ,"frThemeRoot" ,"frTimer" ,"frTimerFolder" ,"frTimerRoot" ,"frTrend" ,"frXMLController",  "fr_TimerObject", "fr_TimerLine", "fr_SetpointSuggestionLine", "fr_SetpointSuggestion", "fr_SetpointIncludeSuggestion", "fr_SetpointDevice", "fr_SetpointBarGraph", "fr_Setpoint", "fr_SelectMenuLine", "fr_RadarChartInfo", "fr_RadarChart8_2Z", "fr_RadarChart8", "fr_RadarChart16", "fr_RadarChart10_2Z", "fr_RadarChart10", "fr_RadarChart05", "fr_RadarChart04", "fr_RadarChart03", "fr_OpenScreenButton", "fr_ObjectChecker", "fr_LockG", "fr_InfoAlarm", "fr_Gmanob", "fr_FooterProcessor", "fr_FooterButton", "fr_FocusScreenObjectFrame", "fr_FocusScreenObject", "fr_FindObjInScreen", "fr_E3Browser", "fr_E3Alarm", "fr_AreaOfResponsibilityLine", "fr_Analogic"
            BibliotecaDoObjeto = "Frameautomalogica"
        Case "ufv_TrackersTitulo" ,"ufv_TrackersTexto" ,"ufv_TrackersRetCmd" ,"ufv_Solarimetric" ,"ufv_PositionSummary" ,"ufv_OpenPopup" ,"ufv_Map_Frame" ,"ufv_Map_Trackers" ,"ufv_Map_MouseArea" ,"ufv_Map_MenuInterface" ,"ufv_Map_Balloon" ,"ufv_InverterDetails" ,"ufv_Inverter" ,"ufv_InfoSetpoint" ,"ufv_InfoPotP" ,"ufv_InfoPotG" ,"ufv_InfoAnalogicV" ,"ufv_InfoAnalogicH" ,"ufv_InfoAlarm10" ,"ufv_InfoAlarm05" ,"ufv_InfoAlarm01" ,"ufv_Device" ,"ufv_CommunInfosInverter" ,"ufv_BarChart" ,"ufv_WeatherStation"
            BibliotecaDoObjeto = "ufv_automalogica"
        Case "xcSetpointBarGraph" ,"xcSetpoint" ,"xcSetaUnifilar" ,"xcRetificador" ,"xcResumoUsinaVW" ,"xcReset" ,"xcPresDiferencial" ,"xcPotenciaGerada" ,"xcNotaOperacional" ,"xcNivelDrenagem" ,"xcMsgCOG" ,"xcMenuNotaOperacional" ,"xcMancalLNAT" ,"xcLockG" ,"xcLinhaVert" ,"xcLinhaHor" ,"xcLinhaEncaixe" ,"xcInibicaoAlarme" ,"xcInfoAnalogica" ,"xcInfoAdicLinha" ,"xcConfigColors" ,"xcFiltroAlarmeRodape" ,"xcIncluirSugestaoSP" ,"xcInc" ,"xcGeradorDiesel" ,"xcGerador" ,"xcFreio" ,"xcDisjuntor" ,"xcCtrlPotRea" ,"xcCtrlPartidaParada" ,"xcCtrlParadaEmerg" ,"xcCtrlDigitalOp" ,"xcCtrlDigital4Op" ,"xcCtrlDigital3Op" ,"xcCtrlDigital2Op" ,"xcControleHorario" ,"xcContator" ,"xcComunicacaoStatusUsina" ,"xcComunicacaoStatusCLP" ,"xcFiltro" ,"xcEstadoEstavel" ,"xcEquipeLocal" ,"xcEmManutencao" ,"xcDropList" ,"xcCtrlAnalogica" ,"xcControleMesaLinhaDeConfig" ,"xcComportaGrande" ,"xcComporta" ,"xcCommSubMenu" ,"xcCommandIcon" ,"xcCommandCustomIconMultipleOps" ,"xcCommandCustomIcon" ,"xcVentilador" ,"xcValvulaDistribuidora" ,"xcValvulaByPass" ,"xcValvulaBorboleta" ,"xcValvula4Vias" ,"xcValvula" ,"xcTurbina" ,"xcTrendArrow" ,"xcTrafoSA" ,"xctrafo3enrol" ,"xctrafo2enrol" ,"xcTextoSequenciaPartida" ,"xcTextBox" ,"xcTerra" ,"xcTelaAlarmes" ,"xcStatusMenuComando" ,"xcSetpointSugestao" ,"xcTelaTimeLine" ,"xcTimeLineAlarmObj_Txt4Char" ,"xcTimeLineAlarmObj" ,"vguBarraAlarme" ,"xcBarraSelecaoSequenciaPartida" ,"vguRepComporta" ,"vguPotenciaGerada" ,"vguLinhaVert" ,"vguBarraHor" ,"vguBarraAlarmePeq" ,"xcCaixaEspiral" ,"vguSetaUnifilar" ,"vguCtrlAnalogica" ,"vguCompensadorSincrono" ,"xcCommand" ,"xcChaveFusivel" ,"xcChave2posicoes" ,"xcChave" ,"xcCaixaPelton" ,"xcBomba" ,"xcBateria" ,"xcBarraHor" ,"xcBarraAlarme" ,"xcAlertaFreio" ,"xcAccessObjs" ,"xcAbreTelaAlarmes" ,"vguLinhaHor" ,"vguInfoAnalogica" ,"vguGeradorDiesel" ,"vguGerador" ,"vguEmManutencao" ,"vguDisjuntor"
            BibliotecaDoObjeto = "controles"
        Case "ww_CorrigeLetrasMaiusculasSQL", "ww_CreateMonitorTags", "ww_EngineForLengthyOperations", "ww_LogTrackingEvent", "ww_Monitor", "ww_Parameters", "ww_RetentiveEngine", "ww_SuppressAlarmsOrEvents", "ww_XMLTransfer_QUARENTENA", "ww_CorrigeLetrasMaiusculasSQL", "ww_CreateMonitorTags", "ww_EngineForLengthyOperations", "ww_LogTrackingEvent", "ww_Monitor", "ww_RetentiveEngine", "ww_SuppressAlarmsOrEvents", "ww_XMLTransfer_QUARENTENA", "ww_CommandIcon", "ww_ExplorerLine", "ww_ExplorerLine2", "ww_ExplorerLine2Properties", "ww_Filter", "ww_OpenTabularArea", "ww_ServerInterface", "ww_Sticky", "ww_TabularAreaEngine", "ww_TabularGeralEngine", "ww_TrendMeasNotification", "ww_TypeButton", "ww_ViewerMainFunctions", "ww_CommandIcon", "ww_ExplorerLine", "ww_ExplorerLine2", "ww_ExplorerLine2Properties", "ww_Filter", "ww_OpenTabularArea", "ww_ServerInterface", "ww_Sticky", "ww_TabularAreaEngine", "ww_TabularGeralEngine", "ww_TrendMeasNotification", "ww_TypeButton", "ww_ViewerMainFunctions"
            BibliotecaDoObjeto = "watchwindow"
        Case "DatabaseTags_Parameters", "DatabaseTags_TagMonitor", "DatabaseTags_Version"
            BibliotecaDoObjeto = "databasetags"
        Case "cmdscr_CustomCommandScreen", "cmdScrLibVersion", "patm_CommandCreateInterlockNote", "patm_CommandLogger", "patm_MaintenanceCleanNoteConfig", "patm_xcList", "patm_xcListLine", "patm_xcNoteControlAdd", "cmdscr_Checkbox", "cmdscr_DiscreteArea", "cmdscr_DiscreteLine", "cmdscr_InfoBalloon", "cmdscr_NoteHistProcessor", "cmdscr_SearchTextbox", "patm_AnalogChangePage", "patm_cCommandButton", "patm_cCommandAnalog", "patm_CommandButtonWStatus", "patm_CommandScreenCreator", "patm_CommandScreenProcessor", "patm_cPopupCommand", "patm_cPopupCommandExt", "patm_DeviceMenu", "patm_NoteAlertLine", "patm_ShowMessage", "patm_SimpleSearchTextbox"
            BibliotecaDoObjeto = "poweratm_commandscreen"
        Case "cmdScrXoLibVersion", "cmdScrXoLibVersionXO", "patm_DeviceNote", "patm_NoteControl", "patm_NoteDatabaseControl", "patm_NoteDeviceMonitoring"
            BibliotecaDoObjeto = "poweratm_commandscreen_xo"
        Case "ae_ComboBox", "patm_AlarmScreenEngine", "patm_DropListControl", "patm_HistScreenEngine", "patm_InfoBalloon", "patm_openAlarmHist", "patm_SearchList_AE", "patm_SearchTextbox", "aeLibVersion", "aeLibVersionXO", "patm_AlarmsEventsDBCreator", "patm_oracleConnection", "patm_xoAlarmHistConfig", "patm_xoDatabaseStatus", "patm_xoLogTrackingEvent"
              BibliotecaDoObjeto = "poweratm_commandscreen_xo"
        Case "gtwAckAlarms", "gtwAntiBouncing", "gtwCommand", "gtwFrozenMeasurements", "gtwSetpoint", "gtwTwoCommands", "gtwVersion", "gtwWriteExWithDelay"
              BibliotecaDoObjeto = "gateway"
        Case "aainfo_Note", "aainfo_NoteController", "aainfoXoLibVersion", "aainfo_NoteSource"
              BibliotecaDoObjeto = "advancedalarminformation_xo"
        Case "xcListViewHPxml", "xcMonthScheduling", "xcSchedulingGroup", "xcWeeklyScheduling", "xcAddCommandsSG", "xcBreadcrumb", "xcConfigProperties", "xcConfigPropertiesValues", "xcDateTimeDisplay", "xcDayMonth", "xcEvent", "xcLinkBreadcrumb", "xcMiniListCommands", "xcMiniListCommandsSingle", "xcTemplateListCommands", "xcTemplateListCommandsSingle", "xfBreadcrumb", "xoCalendarScheduling", "xoCheckLicenseCode", "xoExecuteScheduling"
              BibliotecaDoObjeto = "Scheduling"
        Case Else
            BibliotecaDoObjeto = "Desconhecida"
    End Select
End Function

Sub Fim()
End Sub