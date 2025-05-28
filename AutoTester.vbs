Sub AutoTester_CustomConfig()
    'Equipe QA - Célula 5
    '***********************************************************************
    '*  Sub‑rotina : AutoTester_CustomConfig
    '*  Finalidade : Solicitar confirmação ao usuário e, se aprovada,
    '*               disparar o teste automático do domínio.
    '*----------------------------------------------------------------------
    '*  Fluxo:
    '*     1. Exibe MsgBox perguntando se o usuário deseja iniciar o teste.
    '*     2. Se a resposta for vbNo, encerra a execução.
    '*     3. Caso contrário, chama a rotina Main que contém as verificações.
    '***********************************************************************
    Dim Resposta
	Resposta = MsgBox( _
    "ATENÇÃO: Este processo pode ser lento." & vbCrLf & vbCrLf & _
    "O teste automático irá realizar uma varredura completa nas telas do domínio atual, verificando propriedades e objetos conforme as regras do AutoTester." & vbCrLf & vbCrLf & _
    "Deseja realmente iniciar o teste automático agora?", _
    vbYesNo + vbQuestion + vbDefaultButton2, _
    "Iniciar Teste Automático do Domínio")

If Resposta = vbNo Then
    Exit Sub
End If

    Dim connTest
    Set connTest = ConectarBancoQA()

If connTest Is Nothing Then
    MsgBox _
        "====================================" & vbCrLf & _
        "Falha ao conectar ao banco de dados do AutoTester." & vbCrLf & vbCrLf & _
        "Verifique se a VPN da Automa está conectada corretamente." & vbCrLf & _
        "Caso o problema persista, entre em contato com a equipe de QA." & vbCrLf & _
        "===================================="
    Exit Sub
Else
    connTest.Close
    Set connTest = Nothing
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
						"Falha ao conectar ao banco de dados QA." & vbCrLf & _
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
    				"Inconsistências registradas com sucesso no banco de dados do QA." & vbCrLf & vbCrLf & _
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
	VerificarBancoDeDados Obj, "frCustomAppConfig", "AppDBServerPathName", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, "frCustomAppConfig", "DataServer", 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "ww_Parameters"
	VerificarBancoDeDados Obj, "ww_Parameters", "DBServer", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, "ww_Parameters", "DataServer", 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "DatabaseTags_Parameters"
	VerificarBancoDeDados Obj, "DatabaseTags_Parameters", "DBServerPathName", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, "DatabaseTags_Parameters", "DataServer", 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "cmdscr_CustomCommandScreen"
	VerificarBancoDeDados Obj, "cmdscr_CustomCommandScreen", "DBServerPathName", "Fluxo de dados", 0, False
	'-----------------------------------------------------------------------------
Case "patm_CmdBoxXmlCreator"
	VerificarPropriedadeVazia Obj, "patm_CmdBoxXmlCreator", "ConfigPower", "Fluxo de dados", 0, False
	'-----------------------------------------------------------------------------
Case "patm_DeviceNote"
	VerificarPropriedadeVazia Obj, "patm_DeviceNote", "AlarmSource", "Fluxo de dados", 1, False
	VerificarPropriedadeVazia Obj, "patm_DeviceNote", "NoteDatabaseControl", "Fluxo de dados", 1, False
	VerificarPropriedadeVazia Obj, "patm_DeviceNote", "SourceObject", "Fluxo de dados", 1, False
	VerificarTipoPai Obj, "patm_DeviceNote", "DataServer", 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "patm_NoteDatabaseControl"
	VerificarBancoDeDados Obj, "patm_NoteDatabaseControl", "DBServer", "Fluxo de dados", 0, False
	VerificarPropriedadeValor Obj, "patm_NoteDatabaseControl", "GroupCanAddModifyNote", "Operação", 1, "Fluxo de dados", 0, False
	VerificarPropriedadeValor Obj, "patm_NoteDatabaseControl", "Level", "2=[EquipeSCADA, Instrutor]/3=[Supervisão]/4=[EquipeDeTestes]/5=[Operação]", 1, "Fluxo de dados", 0, False
	VerificarTipoPai Obj, "patm_NoteDatabaseControl", "DataServer", 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "patm_xoAlarmHistConfig"
	VerificarBancoDeDados Obj, "patm_xoAlarmHistConfig", "MainDBServerPathName", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, "patm_xoAlarmHistConfig", "DataServer", 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "dtRedundancyConfig"
	VerificarPropriedadeVazia Obj, "dtRedundancyConfig", "NameOfServerToBeStopped", "Fluxo de dados", 1, False
	VerificarTipoPai Obj, "dtRedundancyConfig", "DataServer", 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "patm_CommandLogger"
	VerificarPropriedadeVazia Obj, "patm_CommandLogger", "PowerConfigObj", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, "patm_CommandLogger", "DataServer", 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "hpXMLGenerateStruct"
	VerificarPropriedadeVazia Obj, "hpXMLGenerateStruct", "Log_BancoDeDados", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, "hpXMLGenerateStruct", "DataServer", 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "gtwFrozenMeasurements"
	VerificarPropriedadeVazia Obj, "gtwFrozenMeasurements", "DateTag", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, "gtwFrozenMeasurements", "DataServer", 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "aainfo_NoteController"
	VerificarPropriedadeVazia Obj, "aainfo_NoteController", "DBServerPathName", "Fluxo de dados", 0, False
	VerificarTipoPai Obj, "aainfo_NoteController", "DataServer", 0, "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "xoExecuteScheduler"
	VerificarPropriedadeVazia Obj, "xoExecuteScheduler", "aActivateCommandsGroup", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, "xoExecuteScheduler", "dteEndEvent", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, "xoExecuteScheduler", "dteEndRepeatDate", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, "xoExecuteScheduler", "dteNextEndEvent", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, "xoExecuteScheduler", "dteNextStartEvent", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, "xoExecuteScheduler", "objCommand", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, "xoExecuteScheduler", "strSchedulerName", "Fluxo de dados", 0, False
	VerificarPropriedadeVazia Obj, "xoExecuteScheduler", "UserField01", "Fluxo de dados", 0, False
	'-----------------------------------------------------------------------------
Case "manut_ImportMeasAndCmdList"
	VerificarObjetoInternoIndevido Obj, "manut_ImportMeasAndCmdList", "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "xots_StandardStudioSettings"
	VerificarObjetoInternoIndevido Obj, "xots_StandardStudioSettings", "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
Case "xots_ConvertAqDriversIntoVbScri"
	VerificarObjetoInternoIndevido Obj, "xots_ConvertAqDriversIntoVbScri", "Fluxo de dados", 1, False
	'-----------------------------------------------------------------------------
	  '=================================================================
	  ' Dominio
	  '================================================================= 
Case "AlarmServer"
	VerificarPropriedadeVazia Obj, "AlarmServer", "DataSource", "Domínio", 0, False
	'-----------------------------------------------------------------------------
Case "DBServer"
	VerificarPropriedadeValor Obj, "DBServer", "SourceType", 2, 0, "Domínio", 0, False
	'-----------------------------------------------------------------------------
Case "Hist"
	VerificarPropriedadeVazia Obj, "Hist", "DBServer", "Domínio", 1, False
	VerificarBancoDeDados Obj, "Hist", "DBServer", "Domínio", 0, False
	If VerificarPropriedadeHabilitada(Obj, "Hist", "EnableDiscard", False, "Domínio", 2, True) Then
	VerificarPropriedadeValor Obj, "Hist", "DiscardInterval", 1, 1, "Domínio", 0, False
	VerificarPropriedadeValor Obj, "Hist", "DiscardTimeUnit", 2, 1, "Domínio", 0, False
	VerificarPropriedadeValor Obj, "Hist", "VerificationInterval", 1, 1, "Domínio", 0, False
	VerificarPropriedadeValor Obj, "Hist", "VerificationUnit", 2, 1, "Domínio", 0, False
	End If
	If VerificarPropriedadeHabilitada(Obj, "Hist", "EnableBackupTable", True, "Domínio", 2, True) Then
	VerificarPropriedadeValor Obj, "Hist", "BackupDiscardInterval", 12, 1, "Domínio", 0, False
	VerificarPropriedadeValor Obj, "Hist", "BackupDiscardTimeUnit", 2, 1, "Domínio", 0, False
	End If
	'-----------------------------------------------------------------------------
Case "Historian"
	VerificarPropriedadeVazia Obj, "Historian", "DBServer", "Domínio", 1, False
	VerificarBancoDeDados Obj, "Historian", "DBServer", "Domínio", 0, False
	If VerificarPropriedadeHabilitada(Obj, "Historian", "EnableDiscard", False, "Domínio", 2, True) Then
	VerificarPropriedadeValor Obj, "Historian", "DiscardInterval", 1, 1, "Domínio", 0, False
	VerificarPropriedadeValor Obj, "Historian", "DiscardTimeUnit", 2, 1, "Domínio", 0, False
	VerificarPropriedadeValor Obj, "Historian", "VerificationInterval", 1, 1, "Domínio", 0, False
	VerificarPropriedadeValor Obj, "Historian", "VerificationUnit", 2, 1, "Domínio", 0, False
	End If
	If VerificarPropriedadeHabilitada(Obj, "Historian", "EnableBackupTable", True, "Domínio", 2, True) Then
	VerificarPropriedadeValor Obj, "Historian", "BackupDiscardInterval", 12, 1, "Domínio", 0, False
	VerificarPropriedadeValor Obj, "Historian", "BackupDiscardTimeUnit", 2, 1, "Domínio", 0, False
	End If
	'-----------------------------------------------------------------------------
	  '=================================================================
	  ' Drivers
	  '================================================================= 
Case "IODriver"
	VerificarPropriedadeVazia Obj, "IODriver", "DriverLocation", "Drivers", 0, False
	VerificarPropriedadeValor Obj, "IODriver", "WriteSyncMode", 2, 0, "Drivers", 0, False
	VerificarPropriedadeValor Obj, "IODriver", "ExposeToOpc", 3, 0, "Drivers", 0, False
	            ' Verificação adicional: contagem de IOTags associadas ao driver
	             Dim qtdeIOTags, mensagem
	            qtdeIOTags = ContarObjetosDoTipo(Obj, "IOTag")
	                If qtdeIOTags <= 1 Then
	                    mensagem = "IODriver com quantidade insuficiente de IOTags. (" & qtdeIOTags & " encontrada(s))"
	                    If GerarCSV Then
	                        Call AdicionarErroExcel(DadosExcel, Obj.PathName, "0", mensagem, "Drivers", "IODriver")
	                    Else
	                        Call AdicionarErroBanco(DadosExcel, Obj.PathName, "0", mensagem, "IODriver", "Drivers")
	                    End If
	                End If
	'-----------------------------------------------------------------------------
	  '=================================================================
	  ' Biblioteca
	  '================================================================= 
Case "WaterConfig"
	VerificarPropriedadeVazia Obj, "WaterConfig", "ModelFile", "Biblioteca", 0, False
	'-----------------------------------------------------------------------------
Case "WaterDistributionNetwork"
	VerificarPropriedadeVazia Obj, "WaterDistributionNetwork", "City", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, "WaterDistributionNetwork", "Company", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, "WaterDistributionNetwork", "CompanyAcronym", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, "WaterDistributionNetwork", "Contract", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, "WaterDistributionNetwork", "Name", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, "WaterDistributionNetwork", "Neighborhood", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, "WaterDistributionNetwork", "Organization", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, "WaterDistributionNetwork", "Region", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, "WaterDistributionNetwork", "State", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, "WaterDistributionNetwork", "StateAcronym", "Biblioteca", 0, False
	VerificarPropriedadeVazia Obj, "WaterDistributionNetwork", "Note", "Biblioteca", 0, False
	Dim containerTypes
	containerTypes = Array("DataFolder", "DrawGroup", "DataServer", "WaterDistributionNetwork")
	If HasChildOfType(Obj, "WaterStationData", containerTypes) Then
	Dim arrUserFields
	arrUserFields = Array("DadosDaPlanta", "Mapa3D")
	VerificarUserFields Obj, arrUserFields, "WaterDistributionNetwork", 1
	End If
	'-----------------------------------------------------------------------------

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
    If EnergizacaoLinhasBarras Then
        Select Case TipoObjeto
            Case "archLineHorizontal", "archLineVertical", "pwa_LineHoriz", "pwa_Barra", "pwa_Barra2", "pwa_Barra2Vert", "pwa_LineVert", "Uhe_LineHoriz", "Uhe_LineVert", "pwa_Carga"
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
Case "DrawString"
	VerificarPropriedadeVazia Obj, "DrawString", "Value", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "DrawString", "TextColor", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "E3Query"
	VerificarPropriedadeVazia Obj, "E3Query", "DataSource", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Disjuntor"
	VerificarPropriedadeCondicional Obj, "pwa_Disjuntor", "NaoSupervisionado", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Disjuntor", "NaoSupervisionado", "PositionMeas", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Disjuntor", "NaoSupervisionado", "DeviceNote", False, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Disjuntor", "DeviceNote", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Disjuntor", "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Disjuntor", "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_DisjuntorP"
	VerificarPropriedadeCondicional Obj, "pwa_DisjuntorP", "NaoSupervisionado", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_DisjuntorP", "NaoSupervisionado", "PositionMeas", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_DisjuntorP", "NaoSupervisionado", "DeviceNote", False, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_DisjuntorP", "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_DisjuntorP", "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_DisjuntorPP"
	VerificarPropriedadeCondicional Obj, "pwa_DisjuntorPP", "NaoSupervisionado", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_DisjuntorPP", "NaoSupervisionado", "PositionMeas", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_DisjuntorPP", "NaoSupervisionado", "DeviceNote", False, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_DisjuntorPP", "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_DisjuntorPP", "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Seccionadora"
	VerificarPropriedadeCondicional Obj, "pwa_Seccionadora", "NaoSupervisionado", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Seccionadora", "NaoSupervisionado", "PositionMeas", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Seccionadora", "NaoSupervisionado", "DeviceNote", False, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Seccionadora", "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Seccionadora", "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_BarraAlarme"
	VerificarPropriedadeCondicional Obj, "pwa_BarraAlarme", "NaoSupervisionado", "AnalogMeas", False, "Telas", 0, False
	VerificarPropriedadeCondicional Obj, "pwa_BarraAlarme", "NaoSupervisionado", "AlarmSource", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_BarraAlarme", "AnalogMeas", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_BarraAlarme", "ValorMaximo", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_BarraAlarme", "ValorMinimo", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Barra"
	VerificarPropriedadeCondicional Obj, "pwa_Barra", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Barra", "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Barra", "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Barra2"
	VerificarPropriedadeCondicional Obj, "pwa_Barra2", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Barra2", "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Barra2", "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Barra2Vert"
	VerificarPropriedadeCondicional Obj, "pwa_Barra2Vert", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Barra2Vert", "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Barra2Vert", "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Bateria"
	VerificarPropriedadeValor Obj, "pwa_Bateria", "Energizado", True, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Bateria", "CorOff", 15790320, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Bateria", "CorOn", 5263440, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_BotaoAbreTela"
	VerificarPropriedadeVazia Obj, "pwa_BotaoAbreTela", "Config_Zoom", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "pwa_BotaoAbreTela", "Config_TelaQuadroPatName", "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_BotaoAbreTela", "Config_Descricao", "Desccrição", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, "pwa_BotaoAbreTela", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Gerador"
	VerificarPropriedadeCondicional Obj, "pwa_Gerador", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Gerador", "GenEstado", "SourceObject", "NOTEMPTY", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_GeradorG"
	VerificarPropriedadeCondicional Obj, "pwa_GeradorG", "PotenciaMedia", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "pwa_GeradorG", "PotenciaMedia", "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_GeradorG", "PotenciaMaximaNominal", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_GeradorG", "Gerador", True, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_HomeButton"
	VerificarPropriedadeVazia Obj, "pwa_HomeButton", "ScreenOrFramePathName", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "pwa_HomeButton", "ScreenDescription", "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_HomeButton", "Description", "Alarmes", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_GrupoVSL"
	VerificarPropriedadeVazia Obj, "pwa_GrupoVSL", "PositionMeasObject", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "pwa_GrupoVSL", "AnalogMeas", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoAlarme01"
	VerificarPropriedadeVazia Obj, "pwa_InfoAlarme01", "SourceObject01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_InfoAlarme01", "Descricao", "XXX", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, "pwa_InfoAlarme01", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoAlarme05"
	VerificarPropriedadeVazia Obj, "pwa_InfoAlarme05", "SourceObject01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_InfoAlarme05", "Descricao", "XXX", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, "pwa_InfoAlarme05", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoAlarme10"
	VerificarPropriedadeVazia Obj, "pwa_InfoAlarme10", "SourceObject01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_InfoAlarme10", "Descricao", "XXX", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, "pwa_InfoAlarme10", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoAnalogica"
	VerificarPropriedadeVazia Obj, "pwa_InfoAnalogica", "SourceObject", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_InfoAnalogica", "SourceObject", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_InfoAnalogica", "SPShow", "SPTag", True, "Telas", 1, False
	VerificarPropriedadeTextoProibido Obj, "pwa_InfoAnalogica", "SourceObject", ".Value", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, "pwa_InfoAnalogica", "ShowUE", True, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, "pwa_InfoAnalogica", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoAnalogicaG"
	VerificarPropriedadeVazia Obj, "pwa_InfoAnalogicaG", "SourceObject", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_InfoAnalogicaG", "SourceObject", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_InfoAnalogicaG", "SPShow", "SPTag", True, "Telas", 1, False
	VerificarPropriedadeTextoProibido Obj, "pwa_InfoAnalogicaG", "SourceObject", ".Value", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, "pwa_InfoAnalogicaG", "ShowUE", True, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, "pwa_InfoAnalogicaG", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoPot"
	VerificarPropriedadeVazia Obj, "pwa_InfoPot", "PotenciaMedida", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_InfoPot", "PotenciaMedia", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_InfoPot", "HabilitaSetpoint", "SetPointPotencia", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_InfoPot", "PotenciaMaximaNominal", 0, 1, "Telas", 0, False
	VerificarObjetoDesatualizado Obj, "pwa_InfoPot", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoPotP"
	VerificarPropriedadeVazia Obj, "pwa_InfoPotP", "PotenciaMedida", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_InfoPotP", "PotenciaMedia", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_InfoPotP", "HabilitaSetpoint", "SetPointPotencia", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_InfoPotP", "PotenciaMaximaNominal", 100, 1, "Telas", 0, False
	VerificarObjetoDesatualizado Obj, "pwa_InfoPotP", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoPotG"
	VerificarPropriedadeVazia Obj, "pwa_InfoPotG", "PotenciaMedida", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_InfoPotG", "PotenciaMedia", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_InfoPotG", "HabilitaSetpoint", "SetPointPotencia", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_InfoPotG", "PotenciaMaximaNominal", 100, 1, "Telas", 0, False
	VerificarObjetoDesatualizado Obj, "pwa_InfoPotG", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Inversor"
	VerificarPropriedadeValor Obj, "pwa_Inversor", "Energizado", True, 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_LineHoriz"
	VerificarPropriedadeValor Obj, "pwa_LineHoriz", "Energizado", True, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_LineHoriz", "CorOff", 15790320, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_LineHoriz", "CorOn", 5263440, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_LineVert"
	VerificarPropriedadeValor Obj, "pwa_LineVert", "Energizado", True, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_LineVert", "CorOff", 15790320, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_LineVert", "CorOn", 5263440, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_AutoTrafo"
	VerificarPropriedadeCondicional Obj, "pwa_AutoTrafo", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_AutoTrafo", "SourceObject", "DeviceNote", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_AutoTrafo", "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_AutoTrafo", "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_AutoTrafo", "CorOnTerminal2", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_AutoTrafo", "CorOnTerminal3", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Capacitor"
	VerificarPropriedadeCondicional Obj, "pwa_Capacitor", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Capacitor", "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Capacitor", "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Carga"
	VerificarPropriedadeCondicional Obj, "pwa_Carga", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Carga", "SourceObject", "DeviceNote", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Carga", "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Carga", "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Conexao"
	VerificarPropriedadeValor Obj, "pwa_Conexao", "CorObjeto", 4605520, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Jumper"
	VerificarPropriedadeValor Obj, "pwa_Jumper", "CorObjeto", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Retificador"
	VerificarPropriedadeValor Obj, "pwa_Retificador", "Energizado", True, 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_Terra"
	VerificarPropriedadeValor Obj, "pwa_Terra", "CorTerrra", 16777215, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Terra2"
	VerificarPropriedadeValor Obj, "pwa_Terra2", "CorTerrra", 16777215, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Reactor"
	VerificarPropriedadeCondicional Obj, "pwa_Reactor", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Reactor", "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Reactor", "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Relig"
	VerificarPropriedadeCondicional Obj, "pwa_Relig", "NaoSupervisionado", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Relig", "NaoSupervisionado", "PositionMeas", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Relig", "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Relig", "CorOff", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Relig", "CorOn", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Sensor"
	VerificarPropriedadeCondicional Obj, "pwa_Sensor", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Sensor", "BorderColor", 255, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_VentForc"
	VerificarPropriedadeCondicional Obj, "pwa_VentForc", "Unsupervised", "Measure", False, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_TapV"
	VerificarPropriedadeVazia Obj, "pwa_TapV", "Measure", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "pwa_TapV", "CmdDown", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "pwa_TapV", "CmdUp", "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_TapV", "MaxLimit", 8, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_TapV", "MinLimit", 2, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_InfoPotRea"
	VerificarPropriedadeVazia Obj, "pwa_InfoPotRea", "PotRea", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_InfoPotRea", "PotRea", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_InfoPotRea", "SpShow", "SetPointPotencia", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_InfoPotRea", "MaxPotReaPos", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_InfoPotRea", "MinPotReaPos", -100, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_ReguladorTensao"
	VerificarPropriedadeCondicional Obj, "pwa_ReguladorTensao", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_ReguladorTensao", "SourceObject", "DeviceNote", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_ReguladorTensao", "UseNotes", "TAPMeas", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_ReguladorTensao", "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_ReguladorTensao", "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_ReguladorTensao", "CorOnTerminal2", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Menu"
	VerificarPropriedadeVazia Obj, "pwa_Menu", "SourceObject", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "pwa_Menu", "SpecialScreens", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "pwa_Menu", "ScreenArg", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "pwa_TrafoSA"
	VerificarPropriedadeCondicional Obj, "pwa_TrafoSA", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_TrafoSA", "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_TrafoSA", "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_TrafoSA", "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_TrafoSA", "CorOnTerminal2", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Trafo3Type01"
	VerificarPropriedadeCondicional Obj, "pwa_Trafo3Type01", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Trafo3Type01", "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Trafo3Type01", "TAPSPShow", "TAPSPTag", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Trafo3Type01", "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo3Type01", "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo3Type01", "CorOnTerminal2", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo3Type01", "CorOnTerminal3", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Trafo3_P"
	VerificarPropriedadeCondicional Obj, "pwa_Trafo3_P", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Trafo3_P", "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Trafo3_P", "TAPSPShow", "TAPSPTag", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Trafo3_P", "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo3_P", "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo3_P", "CorOnTerminal2", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo3_P", "CorOnTerminal3", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Trafo3"
	VerificarPropriedadeCondicional Obj, "pwa_Trafo3", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Trafo3", "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Trafo3", "TAPSPShow", "TAPSPTag", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Trafo3", "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo3", "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo3", "CorOnTerminal2", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo3", "CorOnTerminal3", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Trafo2Term"
	VerificarPropriedadeCondicional Obj, "pwa_Trafo2Term", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Trafo2Term", "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Trafo2Term", "TAPSPShow", "TAPSPTag", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Trafo2Term", "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo2Term", "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo2Term", "CorOnTerminal2", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo2Term", "CorOnTerminal3", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "pwa_Trafo2"
	VerificarPropriedadeCondicional Obj, "pwa_Trafo2", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Trafo2", "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "pwa_Trafo2", "TAPSPShow", "TAPSPTag", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "pwa_Trafo2", "CorOff", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo2", "CorOnTerminal1", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo2", "CorOnTerminal2", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "pwa_Trafo2", "CorOnTerminal3", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_AbnormalityIndicator"
	VerificarPropriedadeValor Obj, "gx_AbnormalityIndicator", "Measurement01Active", False, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_AbnormalityIndicator", "Measurement01Desc", "XXX", 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_AbnormalityIndicator", "SideToGrowing", 1, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_Analogic"
	VerificarPropriedadeVazia Obj, "gx_Analogic", "Measure", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_Analogic", "Show", False, 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_ButtonOpenCommandScreen"
	VerificarPropriedadeVazia Obj, "gx_ButtonOpenCommandScreen", "SourceObject", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_ButtonOpenCommandScreen", "Description", "descrição", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_Counter"
	VerificarPropriedadeValor Obj, "gx_Counter", "Value", 0, 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_CtrlDigital"
	VerificarPropriedadeCondicional Obj, "gx_CtrlDigital", "Enabled", "CommandPathName", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_CtrlDigital", "Active", False, 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_CtrlDigital", "Descr", "Desc", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_CtrlDigital1Op"
	VerificarPropriedadeCondicional Obj, "gx_CtrlDigital1Op", "Enabled", "CommandPathName", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_CtrlDigital1Op", "Tag", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_CtrlDigital1Op", "Descr", "Desc", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_CtrlDigital2Op"
	VerificarPropriedadeCondicional Obj, "gx_CtrlDigital2Op", "Enabled", "CommandPathName1", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_CtrlDigital2Op", "Tag1", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_CtrlDigital2Op", "Descr1", "Desc1", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_CtrlDigital2Op", "Enabled", "CommandPathName2", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_CtrlDigital2Op", "Tag2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_CtrlDigital2Op", "Descr2", "Desc2", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_CtrlDigital3Op"
	VerificarPropriedadeCondicional Obj, "gx_CtrlDigital3Op", "Enabled", "CommandPathName1", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_CtrlDigital3Op", "Tag1", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_CtrlDigital3Op", "Descr1", "Desc1", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_CtrlDigital3Op", "Enabled", "CommandPathName2", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_CtrlDigital3Op", "Tag2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_CtrlDigital3Op", "Descr2", "Desc2", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_CtrlDigital3Op", "Enabled", "CommandPathName3", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_CtrlDigital3Op", "Tag3", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_CtrlDigital3Op", "Descr3", "Desc3", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_CtrlDigital4Op"
	VerificarPropriedadeCondicional Obj, "gx_CtrlDigital4Op", "Enabled", "CommandPathName1", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_CtrlDigital4Op", "Tag1", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_CtrlDigital4Op", "Descr1", "Desc1", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_CtrlDigital4Op", "Enabled", "CommandPathName2", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_CtrlDigital4Op", "Tag2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_CtrlDigital4Op", "Descr2", "Desc2", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_CtrlDigital4Op", "Enabled", "CommandPathName3", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_CtrlDigital4Op", "Tag3", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_CtrlDigital4Op", "Descr3", "Desc3", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_CtrlDigital4Op", "Enabled", "CommandPathName4", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_CtrlDigital4Op", "Tag4", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_CtrlDigital4Op", "Descr4", "Desc4", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_InfoAnalogic"
	VerificarPropriedadeVazia Obj, "gx_InfoAnalogic", "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_InfoAnalogic", "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_InfoAnalogic", "SPShow", "SPTag", True, "Telas", 1, False
	VerificarPropriedadeTextoProibido Obj, "gx_InfoAnalogic", "Measure", ".Value", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, "gx_InfoAnalogic", "ShowUE", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_InfoAnalogic2"
	VerificarPropriedadeVazia Obj, "gx_InfoAnalogic2", "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_InfoAnalogic2", "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_InfoAnalogic2", "SPShow", "SPTag", True, "Telas", 1, False
	VerificarPropriedadeTextoProibido Obj, "gx_InfoAnalogic2", "Measure", ".Value", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, "gx_InfoAnalogic2", "ShowUE", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_InfoDoughnutChart"
	VerificarPropriedadeVazia Obj, "gx_InfoDoughnutChart", "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_InfoDoughnutChart", "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_InfoDoughnutChart", "SPShow", "SPTag", True, "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_InfoDoughnutChart", "NominalValue", 100, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_InfoSetpoint"
	VerificarPropriedadeVazia Obj, "gx_InfoSetpoint", "SPTag", "Telas", 1, False
	VerificarPropriedadeHabilitada Obj, "gx_InfoSetpoint", "ShowUE", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_Menu"
	'Existe evento OnClick
	'-----------------------------------------------------------------------------
Case "gx_Notes"
	VerificarPropriedadeVazia Obj, "gx_Notes", "DeviceNote", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_OpenTabularArea1"
	VerificarPropriedadeValor Obj, "gx_OpenTabularArea1", "Descricao", "XXX", 1, "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_OpenTabularArea1", "Areas", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "gx_OpenTabularArea1", "ScreenZoom", "ScreenPathNames ", "NOTEMPTY", "Telas", 1, False
	'Talvez nova função se propriedade auxiliar estiver preenchidad verficar se a propriedade principal está dfierente de 0
	'-----------------------------------------------------------------------------
Case "gx_QualityIcon"
	VerificarPropriedadeVazia Obj, "gx_QualityIcon", "Measurement", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart03"
	VerificarPropriedadeVazia Obj, "gx_RadarChart03", "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart03", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart03", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart03", "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart03", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart03", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart03", "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart03", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart03", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart03", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart03", "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart04"
	VerificarPropriedadeVazia Obj, "gx_RadarChart04", "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart04", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart04", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart04", "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart04", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart04", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart04", "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart04", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart04", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart04", "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart04", "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart04", "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart04", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart04", "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart05"
	VerificarPropriedadeVazia Obj, "gx_RadarChart05", "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart05", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart05", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart05", "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart05", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart05", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart05", "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart05", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart05", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart05", "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart05", "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart05", "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart05", "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart05", "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart05", "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart05", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart05", "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart06"
	VerificarPropriedadeVazia Obj, "gx_RadarChart06", "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart06", "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart06", "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart06", "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart06", "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart06", "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart06", "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart07"
	VerificarPropriedadeVazia Obj, "gx_RadarChart07", "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart07", "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart07", "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart07", "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart07", "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart07", "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart07", "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart07", "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart08"
	VerificarPropriedadeVazia Obj, "gx_RadarChart08", "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08", "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08", "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08", "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08", "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08", "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08", "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08", "Meas08", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08", "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart08_2Z"
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas01_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas01_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas02_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas02_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas03_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas03_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas04_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas04_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas05_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas05_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas06_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas06_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas07_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas07_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas08_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart08_2Z", "Meas08_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "ZoneMinLim", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Title01", "Quente", 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart08_2Z", "Title02", "Frio", 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart09"
	VerificarPropriedadeVazia Obj, "gx_RadarChart09", "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart09", "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart09", "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart09", "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart09", "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart09", "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart09", "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart09", "Meas08", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart09", "Meas09", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas09MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "Meas09MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart09", "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart10"
	VerificarPropriedadeVazia Obj, "gx_RadarChart10", "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10", "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10", "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10", "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10", "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10", "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10", "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10", "Meas08", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10", "Meas09", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas09MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas09MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10", "Meas10", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas10MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "Meas10MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10", "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart10_2Z"
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas01_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas01_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas02_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas02_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas03_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas03_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas04_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas04_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas05_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas05_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas06_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas06_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas07_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas07_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas08_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas08_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas09_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas09_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas09MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas09MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas10_z1", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart10_2Z", "Meas10_z2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas10MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Meas10MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "ZoneMinLim", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Title01", "Quente", 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart10_2Z", "Title02", "Frio", 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart12"
	VerificarPropriedadeVazia Obj, "gx_RadarChart12", "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart12", "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart12", "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart12", "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart12", "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart12", "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart12", "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart12", "Meas08", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart12", "Meas09", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas09MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas09MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart12", "Meas10", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas10MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas10MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart12", "Meas11", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas11MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas11MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart12", "Meas12", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas12MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "Meas12MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart12", "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart16"
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas08", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas09", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas09MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas09MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas10", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas10MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas10MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas11", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas11MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas11MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas12", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas12MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas12MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas13", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas13MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas13MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas14", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas14MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas14MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas15", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas15MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas15MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart16", "Meas16", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas16MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "Meas16MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart16", "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "gx_RadarChart20"
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas01", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas01MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas01MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas02", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas02MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas02MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas03", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas03MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas03MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas04", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas04MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas04MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas05", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas05MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas05MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas06", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas06MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas06MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas07", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas07MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas07MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas08", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas08MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas08MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas09", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas09MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas09MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas10", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas10MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas10MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas11", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas11MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas11MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas12", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas12MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas12MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas13", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas13MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas13MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas14", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas14MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas14MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas15", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas15MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas15MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas16", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas16MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas16MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas17", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas17MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas17MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas18", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas18MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas18MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas19", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas19MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas19MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "gx_RadarChart20", "Meas20", "Telas", 1, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas20MaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "Meas20MinLim", 15, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "ZoneMaxLim", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "gx_RadarChart20", "ZoneMinLim", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_AbnormalityIndicator"
	VerificarPropriedadeVazia Obj, "uhe_AbnormalityIndicator", "Measurement01Active", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_AbnormalityIndicator", "Measurement01Desc", "XXX", 1, "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_AbnormalityIndicator", "SideToGrowing", 1, 1, "Telas", 0, False
	VerificarObjetoDesatualizado Obj, "uhe_AbnormalityIndicator", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_AirCompressor"
	VerificarPropriedadeCondicional Obj, "uhe_AirCompressor", "Unsupervised", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_AirCompressor", "Unsupervised", "CompressorOff", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_AirCompressor", "Unsupervised", "CompressorOn", False, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AirOilTank"
	VerificarPropriedadeVazia Obj, "uhe_AirOilTank", "Measure", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_AirOilTank", "MaxLimit", 3, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AirOilTank", "MinLimit", 0, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, "uhe_AirOilTank", "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AlarmBar"
	VerificarPropriedadeVazia Obj, "uhe_AlarmBar", "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_AlarmBar", "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_AlarmBar", "ValorMaximo", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AlarmBar", "ValorMinimo", 0, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, "uhe_AlarmBar", "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AnalogBar"
	VerificarPropriedadeVazia Obj, "uhe_AnalogBar", "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_AnalogBar", "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar", "MaxValue", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar", "MinValue", 0, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, "uhe_AnalogBar", "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AnalogBar5Limits"
	VerificarPropriedadeVazia Obj, "uhe_AnalogBar5Limits", "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_AnalogBar5Limits", "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5Limits", "MaxValue", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5Limits", "MinValue", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5Limits", "Limit01", 50, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5Limits", "Limit02", 60, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5Limits", "Limit03", 70, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5Limits", "Limit04", 80, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5Limits", "Limit05", 90, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, "uhe_AnalogBar5Limits", "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AnalogBar5LimitsH"
	VerificarPropriedadeVazia Obj, "uhe_AnalogBar5LimitsH", "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_AnalogBar5LimitsH", "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5LimitsH", "MaxValue", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5LimitsH", "MinValue", 0, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5LimitsH", "Limit01", 50, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5LimitsH", "Limit02", 60, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5LimitsH", "Limit03", 70, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5LimitsH", "Limit04", 80, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBar5LimitsH", "Limit05", 90, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, "uhe_AnalogBar5LimitsH", "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AnalogBarHor"
	VerificarPropriedadeVazia Obj, "uhe_AnalogBarHor", "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_AnalogBarHor", "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBarHor", "MaxValue", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBarHor", "MinValue", 0, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, "uhe_AnalogBarHor", "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_AnalogBarP"
	VerificarPropriedadeVazia Obj, "uhe_AnalogBarP", "Measure", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_AnalogBarP", "Measure", "AlarmSource", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBarP", "MaxValue", 100, 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "uhe_AnalogBarP", "MinValue", 0, 1, "Telas", 0, False
	VerificarPropriedadeTextoProibido Obj, "uhe_AnalogBarP", "Measure", ".Value", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_BielaHidraulica"
	VerificarPropriedadeVazia Obj, "uhe_BielaHidraulica", "Flambada", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_BielaMecanica"
	VerificarPropriedadeVazia Obj, "uhe_BielaMecanica", "Flambada", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Block"
	VerificarPropriedadeVazia Obj, "uhe_Block", "Block_Tag", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_Block", "BlockArea", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_Block", "OpenCommandSelectMenu", "CommandPathNames", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Bomb"
	VerificarPropriedadeCondicional Obj, "uhe_Bomb", "Unsupervised", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_Bomb", "OpenCommandSelectMenu", "CommandPathNames", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_Bomb", "Unsupervised", "DeviceNote", False, "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_Bomb", "BombOff", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_Bomb", "BombOn", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Bomb2"
	VerificarPropriedadeCondicional Obj, "uhe_Bomb2", "Unsupervised", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_Bomb2", "OpenCommandSelectMenu", "CommandPathNames", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_Bomb2", "Unsupervised", "DeviceNote", False, "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_Bomb2", "BombOff", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_Bomb2", "BombOn", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_BrakeAlert"
	VerificarPropriedadeCondicional Obj, "uhe_BrakeAlert", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_BrakeAlert", "SourceObject", "DeviceNote", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_BrakeAlert", "BrakeTag", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_BulbTurbine"
	VerificarPropriedadeVazia Obj, "uhe_BulbTurbine", "Distributor_Tag", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_BusBar"
	VerificarPropriedadeVazia Obj, "uhe_BusBar", "SourceObject", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_BusBar", "ObjectColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_Buzzer"
	VerificarPropriedadeCondicional Obj, "uhe_Buzzer", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_Buzzer", "Playing", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Caixa"
	VerificarPropriedadeVazia Obj, "uhe_Caixa", "Energized", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Chart"
	VerificarPropriedadeVazia Obj, "uhe_Chart", "PenData1", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_Chart", "ObjectColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_ChartP"
	VerificarPropriedadeVazia Obj, "uhe_ChartP", "PenData1", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_ChartP", "ObjectColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_Command"
	VerificarPropriedadeVazia Obj, "uhe_Command", "CommandPathNames", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_Command", "Descr", "Desccrição", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_CommandButton"
	VerificarPropriedadeVazia Obj, "uhe_CommandButton", "CommandPathNames", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CommandButton", "Description", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CommandButton", "Measure", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_Conduto"
	VerificarPropriedadeVazia Obj, "uhe_Conduto", "ComAgua", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_ControlGate"
	VerificarPropriedadeCondicional Obj, "uhe_ControlGate", "Unsupervised", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_ControlGate", "StateOff_Tag", "Telas", 0, False
	VerificarPropriedadeVazia Obj, "uhe_ControlGate", "StateOn_Tag", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlDigital"
	VerificarPropriedadeCondicional Obj, "uhe_CtrlDigital", "Enabled", "CommandPathName", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlDigital", "Active", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_CtrlDigital", "Descr", "Desc", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, "uhe_CtrlDigital", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlDigital1Op"
	VerificarPropriedadeCondicional Obj, "uhe_CtrlDigital1Op", "Enabled", "CommandPathName", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlDigital1Op", "Tag", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_CtrlDigital1Op", "Descr", "Desc", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, "uhe_CtrlDigital1Op", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlDigital2Op"
	VerificarPropriedadeCondicional Obj, "uhe_CtrlDigital2Op", "Enabled", "CommandPathName1", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlDigital2Op", "Tag1", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_CtrlDigital2Op", "Descr1", "Desc1", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_CtrlDigital2Op", "Enabled", "CommandPathName2", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlDigital2Op", "Tag2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_CtrlDigital2Op", "Descr2", "Desc2", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, "uhe_CtrlDigital2Op", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlDigital3Op"
	VerificarPropriedadeCondicional Obj, "uhe_CtrlDigital3Op", "Enabled", "CommandPathName1", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlDigital3Op", "Tag1", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_CtrlDigital3Op", "Descr1", "Desc1", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_CtrlDigital3Op", "Enabled", "CommandPathName2", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlDigital3Op", "Tag2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_CtrlDigital3Op", "Descr2", "Desc2", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_CtrlDigital3Op", "Enabled", "CommandPathName3", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlDigital3Op", "Tag3", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_CtrlDigital3Op", "Descr3", "Desc3", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, "uhe_CtrlDigital3Op", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlDigital4Op"
	VerificarPropriedadeCondicional Obj, "uhe_CtrlDigital4Op", "Enabled", "CommandPathName1", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlDigital4Op", "Tag1", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_CtrlDigital4Op", "Descr1", "Desc1", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_CtrlDigital4Op", "Enabled", "CommandPathName2", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlDigital4Op", "Tag2", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_CtrlDigital4Op", "Descr2", "Desc2", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_CtrlDigital4Op", "Enabled", "CommandPathName3", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlDigital4Op", "Tag3", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_CtrlDigital4Op", "Descr3", "Desc3", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_CtrlDigital4Op", "Enabled", "CommandPathName4", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlDigital4Op", "Tag4", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_CtrlDigital4Op", "Descr4", "Desc4", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, "uhe_CtrlDigital4Op", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlDigitalOp"
	VerificarPropriedadeCondicional Obj, "uhe_CtrlDigitalOp", "Enabled", "CommandPathName", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlDigitalOp", "Active", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_CtrlDigitalOp", "Descr", "Desc", 1, "Telas", 1, False
	VerificarObjetoDesatualizado Obj, "uhe_CtrlDigitalOp", "generic_automalogica", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_CtrlPulse"
	VerificarPropriedadeVazia Obj, "uhe_CtrlPulse", "CmdDecrement", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_CtrlPulse", "CmdIncrement", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Device"
	VerificarPropriedadeVazia Obj, "uhe_Device", "SourceObject", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_DieselGenerator"
	VerificarPropriedadeCondicional Obj, "uhe_DieselGenerator", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_DieselGenerator", "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_DieselGenerator", "SourceObject", "DeviceNote", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_DieselGenerator", "Estado_Tag", "Telas", 1, False
	VerificarPropriedadeValor Obj, "uhe_DieselGenerator", "Descricao", "GDE", 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_Direction"
	VerificarPropriedadeCondicional Obj, "uhe_Direction", "Enabled", "AnalogMeasure", False, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "uhe_Direction", "ObjectColor", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_EarthSwitch"
	VerificarPropriedadeValor Obj, "uhe_EarthSwitch", "CorTerrra", 16777215, 1, "Telas", 0, False
	VerificarPropriedadeVazia Obj, "uhe_EarthSwitch", "ObjectColor", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_ExcitationTransformer"
	VerificarPropriedadeCondicional Obj, "uhe_ExcitationTransformer", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_ExcitationTransformer", "Off", "Telas", 0, False
	VerificarPropriedadeVazia Obj, "uhe_ExcitationTransformer", "On", "Telas", 0, False
	VerificarPropriedadeCondicional Obj, "uhe_ExcitationTransformer", "OpenCommandSelectMenu", "CommandPathNames", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "uhe_Filter"
	VerificarPropriedadeHabilitada Obj, "uhe_Filter", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "uhe_Fan"
	VerificarPropriedadeCondicional Obj, "uhe_Fan", "Unsupervised", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_Fan", "UseNotes", "DeviceNote", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_Fan", "SourceObject", "DeviceNote", "NOTEMPTY", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_Fan", "FanOff", "Telas", 1, False
	VerificarPropriedadeVazia Obj, "uhe_Fan", "FanOn", "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "uhe_Fan", "OpenCommandSelectMenu", "CommandPathNames", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "XCPump"
	VerificarPropriedadeVazia Obj, "XCPump", "SourceObject", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "iconElectricity"
	VerificarPropriedadeVazia Obj, "iconElectricity", "SourceObject", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "iconComFail"
	VerificarPropriedadeVazia Obj, "iconComFail", "SourceObject", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "xcLabel"
	VerificarPropriedadeVazia Obj, "xcLabel", "Caption", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "xcEtiqueta_Manut"
	VerificarPropriedadeVazia Obj, "xcEtiqueta_Manut", "CorObjeto", "Telas", 0, False
	VerificarPropriedadeVazia Obj, "xcEtiqueta_Manut", "EtiquetaVisivel", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "xcEtiqueta"
	VerificarPropriedadeVazia Obj, "xcEtiqueta", "AvisoVisivel", "Telas", 0, False
	VerificarPropriedadeVazia Obj, "xcEtiqueta", "EventoVisivel", "Telas", 0, False
	VerificarPropriedadeVazia Obj, "xcEtiqueta", "FonteObjeto", "Telas", 0, False
	VerificarPropriedadeVazia Obj, "xcEtiqueta", "ForaVisivel", "Telas", 0, False
	VerificarPropriedadeVazia Obj, "xcEtiqueta", "PathNote", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, "xcEtiqueta", "Visible", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "xcWaterTank"
	VerificarPropriedadeVazia Obj, "xcWaterTank", "objSource", "Telas", 0, False
	VerificarPropriedadeVazia Obj, "xcWaterTank", "objWaterDistribution", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, "xcWaterTank", "Visible", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "XCDistribution"
	VerificarPropriedadeVazia Obj, "XCDistribution", "SourceObject", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "XCArrow"
	VerificarPropriedadeHabilitada Obj, "XCArrow", "Visible", True, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archAeroGenerator"
	VerificarPropriedadeVazia Obj, "archAeroGenerator", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, "archAeroGenerator", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archCloud"
	VerificarPropriedadeVazia Obj, "archCloud", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, "archCloud", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archNuclearPlant"
	VerificarPropriedadeVazia Obj, "archNuclearPlant", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, "archNuclearPlant", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archServerRackmountMultiple"
	VerificarPropriedadeVazia Obj, "archServerRackmountMultiple", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, "archServerRackmountMultiple", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archSolarPanel"
	VerificarPropriedadeVazia Obj, "archSolarPanel", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, "archSolarPanel", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archSurveillanceCamera"
	VerificarPropriedadeVazia Obj, "archSurveillanceCamera", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, "archSurveillanceCamera", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archWifi"
	VerificarPropriedadeVazia Obj, "archWifi", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeHabilitada Obj, "archWifi", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archDatabase"
	VerificarPropriedadeHabilitada Obj, "archDatabase", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archFirewall"
	VerificarPropriedadeHabilitada Obj, "archFirewall", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archPCH"
	VerificarPropriedadeHabilitada Obj, "archPCH", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archUHE"
	VerificarPropriedadeHabilitada Obj, "archUHE", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archVideoWall"
	VerificarPropriedadeHabilitada Obj, "archVideoWall", "Enabledd", False, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archSwitch"
	VerificarPropriedadeVazia Obj, "archSwitch", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeCondicional Obj, "archSwitch", "Enabled", "SourceObject", False, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archServerDesktop"
	VerificarPropriedadeVazia Obj, "archServerDesktop", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeCondicional Obj, "archServerDesktop", "Enabled", "SourceObject", False, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archRouter"
	VerificarPropriedadeVazia Obj, "archRouter", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeCondicional Obj, "archRouter", "Enabled", "SourceObject", False, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archServerRackmountSingle"
	VerificarPropriedadeVazia Obj, "archServerRackmountSingle", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeValor Obj, "archServerRackmountSingle", "Description_Text", "Server name", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archViewer"
	VerificarPropriedadeVazia Obj, "archViewer", "CommunicationFailure", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archElectricalMeter"
	VerificarPropriedadeVazia Obj, "archElectricalMeter", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeValor Obj, "archElectricalMeter", "Description_Text", "IED NAME", 1, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archGPSAntenna"
	VerificarPropriedadeVazia Obj, "archGPSAntenna", "CommunicationFailure", "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archLineHorizontal"
	VerificarPropriedadeHabilitada Obj, "archLineHorizontal", "Visible", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "archLineHorizontal", "Enabled", "FlhCom", "NOTEMPTY", "Telas", 0, False
	VerificarPropriedadeValor Obj, "archLineHorizontal", "BorderColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archLineVertical"
	VerificarPropriedadeHabilitada Obj, "archLineVertical", "Visible", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "archLineVertical", "Enabled", "FlhCom", "NOTEMPTY", "Telas", 0, False
	VerificarPropriedadeValor Obj, "archLineVertical", "BorderColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archChannelPanel"
	VerificarPropriedadeHabilitada Obj, "archChannelPanel", "Visible", True, "Telas", 1, False
	    Dim tipoPortas, i, qtdEsperada, propFailure
	    tipoPortas = GetPropriedade(Obj, "Type")
	    Select Case CStr(tipoPortas)
	        Case "3": qtdEsperada = 4
	        Case "2": qtdEsperada = 8
	        Case "1": qtdEsperada = 20
	        Case "0": qtdEsperada = 24
	        Case Else
	            qtdEsperada = 0 ' não definido ou inválido
	    End Select
	    If qtdEsperada > 0 Then
	        For i = 1 To qtdEsperada
	            propFailure = "FailureState" & Right("0" & i, 2)
	            Call VerificarPropriedadeVazia(Obj, TipoObjeto, propFailure, 1, Area, 0)
	        Next
	    End If
	'-----------------------------------------------------------------------------
Case "archChannelPanelG"
	VerificarPropriedadeHabilitada Obj, "archChannelPanelG", "Visible", True, "Telas", 1, False
	    Dim tipoPortasG, iG, qtdEsperadaG, propFailureG
	    tipoPortasG = GetPropriedade(Obj, "Type")
	    Select Case CStr(tipoPortasG)
	        Case "3": qtdEsperadaG = 4
	        Case "2": qtdEsperadaG = 8
	        Case "1": qtdEsperadaG = 20
	        Case "0": qtdEsperadaG = 24
	        Case Else
	            qtdEsperadaG = 0 ' não definido ou inválido
	    End Select
	    If qtdEsperadaG > 0 Then
	        For iG = 1 To qtdEsperadaG
	            propFailureG = "FailureState" & Right("0" & iG, 2)
	            Call VerificarPropriedadeVazia(Obj, TipoObjeto, propFailureG, 1, Area, 0)
	        Next
	    End If
	'-----------------------------------------------------------------------------
Case "archChannelPanelP"
	VerificarPropriedadeHabilitada Obj, "archChannelPanelP", "Visible", True, "Telas", 1, False
	    Dim tipoPortasP, iP, qtdEsperadaP, propFailureP
	    tipoPortasP = GetPropriedade(Obj, "Type")
	    Select Case CStr(tipoPortasP)
	        Case "3": qtdEsperadaP = 4
	        Case "2": qtdEsperadaP = 8
	        Case "1": qtdEsperadaP = 20
	        Case "0": qtdEsperadaP = 24
	        Case Else
	            qtdEsperadaP = 0 ' não definido ou inválido
	    End Select
	    If qtdEsperadaP > 0 Then
	        For iP = 1 To qtdEsperadaP
	            propFailureP = "FailureState" & Right("0" & iP, 2)
	            Call VerificarPropriedadeVazia(Obj, TipoObjeto, propFailureP, 1, Area, 0)
	        Next
	    End If
	'-----------------------------------------------------------------------------
Case "archChannelPanelPP"
	VerificarPropriedadeHabilitada Obj, "archChannelPanelPP", "Visible", True, "Telas", 1, False
	    Dim tiPPoPortasPP, iPP, qtdEsPPeradaPP, propFailurePP
	    tiPPoPortasPP = GetPropriedade(Obj, "Type")
	    Select Case CStr(tiPPoPortasPP)
	        Case "3": qtdEsPPeradaPP = 4
	        Case "2": qtdEsPPeradaPP = 8
	        Case "1": qtdEsPPeradaPP = 20
	        Case "0": qtdEsPPeradaPP = 24
	        Case Else
	            qtdEsPPeradaPP = 0 ' não definido ou inválido
	    End Select
	    If qtdEsPPeradaPP > 0 Then
	        For iPP = 1 To qtdEsPPeradaPP
	            propFailurePP = "FailureState" & Right("0" & iPP, 2)
	            Call VerificarPropriedadeVazia(Obj, TipoObjeto, propFailurePP, 1, Area, 0)
	        Next
	    End If
	'-----------------------------------------------------------------------------
Case "archLed"
	VerificarPropriedadeHabilitada Obj, "archLed", "Visible", True, "Telas", 1, False
	VerificarPropriedadeVazia Obj, "archLed", "FailedState", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archModuloIO"
	VerificarPropriedadeHabilitada Obj, "archModuloIO", "Visible", True, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "archModuloIO", "Enabled", "Failure", "NOTEMPTY", "Telas", 0, False
	VerificarPropriedadeVazia Obj, "archModuloIO", "Text", "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archRTU"
	VerificarPropriedadeHabilitada Obj, "archRTU", "Visible", True, "Telas", 1, False
	VerificarPropriedadeVazia Obj, "archRTU", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeCondicional Obj, "archRTU", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeValor Obj, "archRTU", "Description_Text", "RTU NAME", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "archRTU", "IP_Show", "IP_Text", False, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archIED"
	VerificarPropriedadeHabilitada Obj, "archIED", "Visible", True, "Telas", 1, False
	VerificarPropriedadeVazia Obj, "archIED", "CommunicationFailure", "Telas", 0, False
	VerificarPropriedadeCondicional Obj, "archIED", "Enabled", "SourceObject", False, "Telas", 1, False
	VerificarPropriedadeValor Obj, "archIED", "Description_Text", "IED NAME", 1, "Telas", 1, False
	VerificarPropriedadeCondicional Obj, "archIED", "IP_Show", "IP_Text", False, "Telas", 1, False
	'-----------------------------------------------------------------------------
Case "archInfo"
	VerificarPropriedadeHabilitada Obj, "archInfo", "Visible", True, "Telas", 1, False
	VerificarPropriedadeVazia Obj, "archInfo", "FailedState", "Telas", 1, False
	VerificarPropriedadeValor Obj, "archInfo", "Description_Text", "Painel - Eqp - Port", 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "archInfo", "ObjectColor", 0, 1, "Telas", 0, False
	'-----------------------------------------------------------------------------
Case "archInfoLine"
	VerificarPropriedadeHabilitada Obj, "archInfoLine", "Visible", True, "Telas", 1, False
	VerificarPropriedadeVazia Obj, "archInfoLine", "FailedState", "Telas", 1, False
	VerificarPropriedadeValor Obj, "archInfoLine", "Description_Text", "Painel - Eqp - Port", 1, "Telas", 0, False
	VerificarPropriedadeValor Obj, "archInfoLine", "ObjectColor", 0, 1, "Telas", 0, False
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

'***********************************************************************
'*  Função : ContarObjetosDoTipo
'*----------------------------------------------------------------------
'*  Finalidade :
'*     Percorrer recursivamente a hierarquia de um objeto pai e contar
'*     quantas instâncias de um tipo específico existem (TypeName).
'*
'*  Parâmetros :
'*     ‑ Obj          : Objeto raiz (ex.: um IODriver).
'*     ‑ TipoDesejado : String com o TypeName a ser contado (ex.: "IOTag").
'*
'*  Retorno :
'*     Integer → Quantidade total de objetos do tipo solicitado.
'***********************************************************************
Function ContarObjetosDoTipo(Obj, TipoDesejado)

    On Error Resume Next

    Dim contador, childObj
    contador = 0

    '------------------------------------------------------------------
    ' 1) Conta o próprio objeto, se for do tipo desejado
    '------------------------------------------------------------------
    If TypeName(Obj) = TipoDesejado Then
        contador = contador + 1
    End If

    '------------------------------------------------------------------
    ' 2) Percorre todos os filhos recursivamente
    '------------------------------------------------------------------
    For Each childObj In Obj
        contador = contador + ContarObjetosDoTipo(childObj, TipoDesejado)
    Next

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

Function ConectarBancoQA()
    On Error Resume Next

    'MsgBox "DEBUG: Entrou na função ConectarBancoQA", vbInformation

    Dim conn, strConexao
    Set conn = CreateObject("ADODB.Connection")

    Dim servidor, banco, usuario, senha
    servidor = "GUILHERMEAMARAL\SQLEXPRESS"
    banco = "Teste_QA"
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

Sub Fim()
End Sub