Sub AutoTester_CustomConfig()
'***********************************************************************
'*
'*  Nome:           AutoTester_CustomConfig
'*  Objetivo:       Iniciar o teste automático do domínio, solicitando
'*                  confirmação ao usuário, e então rodar verificações
'*                  em telas, objetos (DataServer/DataFolder) e
'*                  historiadores (Hist/Historian).
'*
'***********************************************************************
    Dim Resposta
    Resposta = MsgBox("Tem certeza que deseja iniciar o teste automático do domínio?", vbYesNo + vbQuestion, "Iniciar teste de domínio?")
    If Resposta = vbNo Then
        Exit Sub
    End If
    
    Main()
End Sub

'***********************************************************************
'*
'*  Declaração de Variáveis Globais
'*
'***********************************************************************
Dim DadosExcel, DadosTxt, DadosBancoDeDados, ListaObjetosLib
Dim CaminhoPrj

' Variáveis/flags presumidas:
'   VerificarBancosCustom (Boolean)
'   DebugMode (Boolean)
'   PathNameTelas (String)
' Declaradas em outro local ou definidas no ambiente.

Set DadosExcel = CreateObject("Scripting.Dictionary")
Set DadosTxt = CreateObject("Scripting.Dictionary")
Set DadosBancoDeDados = CreateObject("Scripting.Dictionary")
Set ListaObjetosLib = CreateObject("Scripting.Dictionary")

If PastaParaSalvarLogs <> "" Then
    CaminhoPrj = PastaParaSalvarLogs
Else
    CaminhoPrj = CreateObject("WScript.Shell").CurrentDirectory
End If

'***********************************************************************
'*
'*  Sub: Main
'*  Rotina Principal de verificação
'*
'***********************************************************************
Sub Main()
    Dim telaArray

    ' 1) Obter lista de telas para verificar (via PathNameTelas)
    telaArray = SplitTelas(PathNameTelas)

    ' 2) Verificar as telas
    VerificarTelas telaArray

    ' 3) Verificar demais objetos do domínio (exceto telas)
    VerificarObjetosDominio

    ' 4) Se for para verificar bancos/Hist, enumerar Hist e Historian
    If VerificarBancosCustom = True Then
        Dim histObj
        ' Lista "Hist"
        For Each histObj In Application.ListFiles("Hist")
            VerificarBancoDoHist histObj
        Next
        ' Lista "Historian"
        For Each histObj In Application.ListFiles("Historian")
            VerificarBancoDoHist histObj
        Next
    End If

    ' 5) Gera relatórios
    If Not DebugMode Then
    	If GerarLogErrosScript Then
        	If Not GerarRelatorioTxt(DadosTxt, CaminhoPrj) Then
            	MsgBox "Falha ao gerar o relatório TXT.", vbCritical
        	End If
    	End If
	
    	If Not GerarRelatorioExcel(DadosExcel, CaminhoPrj) Then
        	MsgBox "Falha ao gerar o relatório Excel.", vbCritical
    	End If
	End If
    MsgBox "Fim"
End Sub

'***********************************************************************
'*
'*  Sub:          VerificarTelas
'*  Objetivo:     Se PathNameTelas tiver nomes de tela,
'*                verifica apenas aquelas; caso contrário, verifica
'*                todas as telas do domínio.
'*
'***********************************************************************
Sub VerificarTelas(telaArray)
    Dim Objeto
    If UBound(telaArray) >= 0 Then
        ' Se houver telas definidas, verifica só essas
        For Each Objeto In Application.ListFiles("Screen")
            If IsTelaNaLista(Objeto.PathName, telaArray) Then
                VerificarPropriedadesObjeto Objeto
            End If
        Next
    Else
        ' Caso não tenha telas específicas, verifica todas
        For Each Objeto In Application.ListFiles("Screen")
            VerificarPropriedadesObjeto Objeto
        Next
    End If
End Sub

'***********************************************************************
'*
'*  Sub:          VerificarObjetosDominio
'*  Objetivo:     Percorrer TODOS os objetos top-level do domínio,
'*                e verificar DataServer, DataFolder, etc.
'*                Se for tela, pula (pois já foi verificada).
'*
'***********************************************************************
Sub VerificarObjetosDominio()
    Dim Obj, tipoObj
    For Each Obj In Application.ListFiles()
        tipoObj = TypeName(Obj)
        If StrComp(tipoObj, "Screen", vbTextCompare) <> 0 Then
            VerificarPropriedadesObjeto Obj
        End If
    Next
End Sub

'***********************************************************************
'*
'*  Função:       SplitTelas
'*  Objetivo:     Receber PathNameTelas, dividi-la em array ("/").
'*
'***********************************************************************
Function SplitTelas(PathNameTelas)
    If Len(Trim(PathNameTelas)) > 0 Then
        SplitTelas = Split(PathNameTelas, "/")
    Else
        SplitTelas = Array()
    End If
End Function

'***********************************************************************
'*
'*  Função:       IsTelaNaLista
'*  Objetivo:     Verificar se o PathName de uma tela está em telaArray
'*
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
'*
'*  Função:       VerificarBancoDoHist
'*  Objetivo:     Recebe um objeto, verifica se é "Hist" ou "Historian",
'*                e caso sim, valida DBServer, duplicidade, etc.
'*
'***********************************************************************
Function VerificarBancoDoHist(Obj)
    If (TypeName(Obj) = "Hist") Or (TypeName(Obj) = "Historian") Then
        ' Exemplo: DBServer deve ser lido via GetPropertyValue
        VerificarPropriedadeVazia Obj, "DBServer", 1, "Hist", 1
        VerificarBancoDeDados Obj, "DBServer", 1, "Hist", 0
    End If
End Function

'***********************************************************************
'*
'*  Função:       VerificarPropriedadesObjeto
'*  Objetivo:     Verifica o tipo do objeto e chama as funções de verificação
'*                necessárias. Cada verificação indica qual método usar.
'*
'***********************************************************************
Function VerificarPropriedadesObjeto(Obj)
    Dim TipoObjeto, child
    TipoObjeto = TypeName(Obj)
	
    Select Case TipoObjeto
    
        Case "DataServer", "DataFolder", "Screen", "DrawGroup"
            ' Recursão
            For Each child In Obj
                VerificarPropriedadesObjeto child
            Next

        '-----------------------------------------------------------------------------
        Case "frCustomAppConfig"
            VerificarBancoDeDados Obj, "AppDBServerPathName", 1, "frCustomAppConfig", 0
			
        '-----------------------------------------------------------------------------
        Case "ww_Parameters"
            VerificarBancoDeDados Obj, "DBServer", 1, "ww_Parameters", 0
			
        '-----------------------------------------------------------------------------
        Case "DatabaseTags_Parameters"
            VerificarBancoDeDados Obj, "DBServerPathName", 1, "DatabaseTags_Parameters", 0
			
        '-----------------------------------------------------------------------------
        Case "cmdscr_CustomCommandScreen"
            VerificarBancoDeDados Obj, "DBServerPathName", 1, "cmdscr_CustomCommandScreen", 0
			
        '-----------------------------------------------------------------------------
        Case "patm_NoteDatabaseControl"
            VerificarBancoDeDados Obj, "DBServer", 1, "patm_NoteDatabaseControl", 0
			
        '-----------------------------------------------------------------------------
        Case "patm_xoAlarmHistConfig"
            VerificarBancoDeDados Obj, "MainDBServerPathName", 1, "patm_xoAlarmHistConfig", 0

        '-----------------------------------------------------------------------------
        Case "pwa_Disjuntor"
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "SourceObject", 0, "Disjuntor", 1
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "PositionMeas", 0, "Disjuntor", 1
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "DeviceNote", 0, "Disjuntor", 1

            VerificarPropriedadeVazia Obj, "CorOff", 0, "Disjuntor", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "Disjuntor", 1

        '-----------------------------------------------------------------------------
        Case "pwa_DisjuntorP"
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "SourceObject", 0, "Disjuntor", 1
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "PositionMeas", 0, "Disjuntor", 1
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "DeviceNote", 0, "Disjuntor", 1

            VerificarPropriedadeVazia Obj, "CorOff", 0, "Disjuntor", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "Disjuntor", 1

        '-----------------------------------------------------------------------------
        Case "pwa_DisjuntorPP"
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "SourceObject", 0, "Disjuntor", 1
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "PositionMeas", 0, "Disjuntor", 1
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "DeviceNote", 0, "Disjuntor", 1

            VerificarPropriedadeVazia Obj, "CorOff", 0, "Disjuntor", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "Disjuntor", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Seccionadora"
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "SourceObject", 0, "Seccionadora", 1
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "PositionMeas", 0, "Seccionadora", 1
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "DeviceNote", 0, "Seccionadora", 1

            VerificarPropriedadeVazia Obj, "CorOff", 0, "Seccionadora", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "Seccionadora", 1

        '-----------------------------------------------------------------------------
        Case "pwa_BarraAlarme"
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "AnalogMeas", 0, "BarraAlarme", 1
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "AlarmSource", 0, "BarraAlarme", 1
            VerificarPropriedadeCondicional Obj, "AnalogMeas", 1, "NOTEMPTY", "AlarmSource", 0, "BarraAlarme", 1

        '-----------------------------------------------------------------------------
        Case "pwa_LineHoriz"
            VerificarPropriedadeVazia Obj, "Energizado", 0, "LineHoriz", 2
            VerificarPropriedadeVazia Obj, "CorOff", 0, "LineHoriz", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "LineHoriz", 1

        '-----------------------------------------------------------------------------
        Case "pwa_LineVert"
            VerificarPropriedadeVazia Obj, "Energizado", 0, "LineVert", 2
            VerificarPropriedadeVazia Obj, "CorOff", 0, "LineVert", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "LineVert", 1

        '-----------------------------------------------------------------------------
        Case "pwa_InfoPot"
            VerificarPropriedadeCondicional Obj, "PotenciaMedia", 0, "NOTEMPTY", "AlarmSource", 0, "InfoPot", 1
            VerificarPropriedadeCondicional Obj, "HabilitaSetpoint", 1, False, "SetPointPotencia", 1, "InfoPot", 1
            VerificarPropriedadeVazia Obj, "PotenciaMedia", 0, "InfoPot", 1
            VerificarPropriedadeValor Obj, "PotenciaMaximaNominal", 1, "InfoPot", 0, 1, 1

        '-----------------------------------------------------------------------------
        Case "pwa_InfoPotG"
            VerificarPropriedadeCondicional Obj, "PotenciaMedia", 0, "NOTEMPTY", "AlarmSource", 0, "InfoPotG", 1
            VerificarPropriedadeCondicional Obj, "HabilitaSetpoint", 1, False, "SetPointPotencia", 1, "InfoPotG", 1
            VerificarPropriedadeVazia Obj, "PotenciaMedia", 0, "InfoPotG", 1
            VerificarPropriedadeValor Obj, "PotenciaMaximaNominal", 100, "InfoPotG", 0, 1, 1
            VerificarObjetoDesatualizado Obj, "InfoPotG", "generic_automalogica"

        '-----------------------------------------------------------------------------
        Case "pwa_InfoPotP"
            VerificarPropriedadeCondicional Obj, "PotenciaMedia", 0, "NOTEMPTY", "AlarmSource", 0, "InfoPotP", 1
            VerificarPropriedadeCondicional Obj, "HabilitaSetpoint", 1, False, "SetPointPotencia", 1, "InfoPotP", 1
            VerificarPropriedadeVazia Obj, "PotenciaMedia", 0, "InfoPotP", 1
            VerificarPropriedadeValor Obj, "PotenciaMaximaNominal", 100, "InfoPotP", 0, 1, 1
            VerificarObjetoDesatualizado Obj, "InfoPotP", "generic_automalogica"

        '-----------------------------------------------------------------------------
        Case "pwa_AutoTrafo"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "AutoTrafo", 1
            VerificarPropriedadeCondicional Obj, "SourceObject", 1, False, "DeviceNote", 0, "AutoTrafo", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "AutoTrafo", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal1", 0, "AutoTrafo", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal2", 0, "AutoTrafo", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal3", 0, "AutoTrafo", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Barra"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Barra", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Barra", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "Barra", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Barra2"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Barra2", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Barra2", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "Barra2", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Barra2Vert"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Barra2Vert", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Barra2Vert", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "Barra2Vert", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Bateria"
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Bateria", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "Bateria", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Capacitor"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Capacitor", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Capacitor", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "Capacitor", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Carga"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Carga", 1
            VerificarPropriedadeCondicional Obj, "SourceObject", 1, False, "DeviceNote", 0, "Carga", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Carga", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "Carga", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Conexao"
            VerificarPropriedadeVazia Obj, "CorObjeto", 0, "Conexao", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Gerador"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Gerador", 1
            VerificarPropriedadeCondicional Obj, "SourceObject", 1, "NOTEMPTY", "GenEstado", 0, "Gerador", 1

        '-----------------------------------------------------------------------------
        Case "pwa_GeradorG"
            VerificarPropriedadeCondicional Obj, "PotenciaMedia", 0, "NOTEMPTY", "AlarmSource", 0, "GeradorG", 1
            VerificarPropriedadeVazia Obj, "PotenciaMedia", 0, "GeradorG", 1
            VerificarPropriedadeValor Obj, "PotenciaMaximaNominal", 0, "GeradorG", 0, 1, 1

        '-----------------------------------------------------------------------------
        Case "pwa_InfoAlarme01"
            VerificarPropriedadeVazia Obj, "SourceObject01", 0, "InfoAlarme01", 1
            VerificarPropriedadeValor Obj, "Descricao", "XXX", "InfoAlarme01", 0, 1, 1
            VerificarObjetoDesatualizado Obj, "InfoAlarme01", "generic_automalogica"

        '-----------------------------------------------------------------------------
        Case "pwa_InfoAlarme05"
            VerificarPropriedadeVazia Obj, "SourceObject01", 0, "InfoAlarme05", 1
            VerificarPropriedadeValor Obj, "Descricao", "XXX", "InfoAlarme05", 0, 1, 1
            VerificarObjetoDesatualizado Obj, "InfoAlarme05", "generic_automalogica"

        '-----------------------------------------------------------------------------
        Case "pwa_InfoAlarme10"
            VerificarPropriedadeVazia Obj, "SourceObject01", 0, "InfoAlarme10", 1
            VerificarPropriedadeValor Obj, "Descricao", "XXX", "InfoAlarme10", 0, 1, 1
            VerificarObjetoDesatualizado Obj, "InfoAlarme10", "generic_automalogica"

        '-----------------------------------------------------------------------------
        Case "pwa_InfoAnalogica"
            VerificarPropriedadeCondicional Obj, "SourceObject", 0, "NOTEMPTY", "AlarmSource", 0, "InfoAnalogica", 1
            VerificarPropriedadeCondicional Obj, "SpShow", 1, False, "SPTag", 0, "InfoAnalogica", 1
            VerificarObjetoDesatualizado Obj, "InfoAnalogica", "generic_automalogica"
            VerificarPropriedadeTextoProibido Obj, "SourceObject", 0, "InfoAnalogica", ".Value", 1

        '-----------------------------------------------------------------------------
        Case "pwa_InfoAnalogicaG"
            VerificarPropriedadeCondicional Obj, "SourceObject", 0, "NOTEMPTY", "AlarmSource", 0, "InfoAnalogicaG", 1
            VerificarObjetoDesatualizado Obj, "InfoAnalogica", "generic_automalogica"
            VerificarPropriedadeTextoProibido Obj, "SourceObject", 0, "InfoAnalogica", ".Value", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Inversor"
            VerificarPropriedadeVazia Obj, "Energizado", 0, "Inversor", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Jumper"
            VerificarPropriedadeVazia Obj, "CorObjeto", 0, "Jumper", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Retificador"
            VerificarPropriedadeVazia Obj, "CorObjeto", 0, "Retificador", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Terra"
            VerificarPropriedadeVazia Obj, "CorTerra", 0, "Terra", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Terra2"
            VerificarPropriedadeVazia Obj, "CorTerra", 0, "Terra2", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Reactor"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Reactor", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Reactor", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "Reactor", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Relig"
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "SourceObject", 0, "Relig", 1
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "PositionMeas", 0, "Relig", 1
            VerificarPropriedadeCondicional Obj, "NaoSupervisionado", 1, False, "DeviceNote", 0, "Relig", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Relig", 1
            VerificarPropriedadeVazia Obj, "CorOn", 0, "Relig", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Sensor"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Sensor", 1
            VerificarPropriedadeVazia Obj, "BorderColor", 0, "Sensor", 1

        '-----------------------------------------------------------------------------
        Case "pwa_VentForc"
            VerificarPropriedadeCondicional Obj, "Unsupervised", 1, False, "Measure", 0, "VentForc", 1

        '-----------------------------------------------------------------------------
        Case "pwa_TapV"
            VerificarPropriedadeVazia Obj, "Measure", 0, "TapV", 1
            VerificarPropriedadeVazia Obj, "CmdDown", 0, "TapV", 1
            VerificarPropriedadeVazia Obj, "CmdUp", 0, "TapV", 1
            VerificarPropriedadeValor Obj, "MaxLimit", 8, "TapV", 0, 1, 1
            VerificarPropriedadeValor Obj, "MinLimit", 2, "TapV", 0, 1, 1

        '-----------------------------------------------------------------------------
        Case "pwa_InfoPotRea"
            VerificarPropriedadeVazia Obj, "PotRea", 0, "InfoPotRea", 1
            VerificarPropriedadeCondicional Obj, "PotRea", 0, "NOTEMPTY", "AlarmSource", 1, "InfoPotRea", 1
            VerificarPropriedadeCondicional Obj, "SpShow", 1, False, "SetPointPotencia", 0, "InfoPotRea", 1
            VerificarPropriedadeValor Obj, "MaxPotReaPos", 100, "InfoPotRea", 0, 1, 1
            VerificarPropriedadeValor Obj, "MinPotReaPos", -100, "InfoPotRea", 0, 1, 1

        '-----------------------------------------------------------------------------
        Case "pwa_ReguladorTensao"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "ReguladorTensao", 1
            VerificarPropriedadeCondicional Obj, "SourceObject", 0, "NOTEMPTY", "DeviceNote", 0, "ReguladorTensao", 1
            VerificarPropriedadeCondicional Obj, "MostraTAP", 1, False, "TAPMeas", 0, "ReguladorTensao", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "ReguladorTensao", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal1", 0, "ReguladorTensao", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal2", 0, "ReguladorTensao", 1

        '-----------------------------------------------------------------------------
        Case "pwa_BotaoAbreTela"
            VerificarPropriedadeVazia Obj, "Config_Zoom", 1, "BotaoAbreTela", 1
            VerificarPropriedadeVazia Obj, "Config_TelaQuadroPatName", 1, "BotaoAbreTela", 1
            VerificarPropriedadeValor Obj, "Config_Descricao", "Desccrição", "BotaoAbreTela", 1, 1, 1
            VerificarObjetoDesatualizado Obj, "BotaoAbreTela", "generic_automalogica"

        '-----------------------------------------------------------------------------
        Case "pwa_Menu"
            VerificarPropriedadeVazia Obj, "ObjectColor", 0, "Menu", 1
            VerificarPropriedadeVazia Obj, "Hierarchy1", 1, "Menu", 1
            VerificarPropriedadeVazia Obj, "SpecificTabularArea", 1, "Menu", 1

        '-----------------------------------------------------------------------------
        Case "pwa_TrafoSA"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "TrafoSA", 1
            VerificarPropriedadeCondicional Obj, "SourceObject", 0, "NOTEMPTY", "DeviceNote", 0, "TrafoSA", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "TrafoSA", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal1", 0, "TrafoSA", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal2", 0, "TrafoSA", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Trafo3Type01"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Trafo3Type01", 1
            VerificarPropriedadeCondicional Obj, "SourceObject", 0, "NOTEMPTY", "DeviceNote", 0, "Trafo3Type01", 1
            VerificarPropriedadeCondicional Obj, "TAPSPShow", 1, False, "TAPSPTag", 0, "Trafo3Type01", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Trafo3Type01", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal1", 0, "Trafo3Type01", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal2", 0, "Trafo3Type01", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal3", 0, "Trafo3Type01", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Trafo3_P"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Trafo3_P", 1
            VerificarPropriedadeCondicional Obj, "SourceObject", 0, "NOTEMPTY", "DeviceNote", 0, "Trafo3_P", 1
            VerificarPropriedadeCondicional Obj, "TAPSPShow", 1, False, "TAPSPTag", 0, "Trafo3_P", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Trafo3_P", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal1", 0, "Trafo3_P", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal2", 0, "Trafo3_P", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal3", 0, "Trafo3_P", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Trafo3"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Trafo3", 1
            VerificarPropriedadeCondicional Obj, "SourceObject", 0, "NOTEMPTY", "DeviceNote", 0, "Trafo3", 1
            VerificarPropriedadeCondicional Obj, "TAPSPShow", 1, False, "TAPSPTag", 0, "Trafo3", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Trafo3", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal1", 0, "Trafo3", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal2", 0, "Trafo3", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal3", 0, "Trafo3", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Trafo2Term"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Trafo2Term", 1
            VerificarPropriedadeCondicional Obj, "SourceObject", 0, "NOTEMPTY", "DeviceNote", 0, "Trafo2Term", 1
            VerificarPropriedadeCondicional Obj, "TAPSPShow", 1, False, "TAPSPTag", 0, "Trafo2Term", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Trafo2Term", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal1", 0, "Trafo2Term", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal2", 0, "Trafo2Term", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal3", 0, "Trafo2Term", 1

        '-----------------------------------------------------------------------------
        Case "pwa_Trafo2"
            VerificarPropriedadeCondicional Obj, "Enable", 1, False, "SourceObject", 0, "Trafo2", 1
            VerificarPropriedadeCondicional Obj, "SourceObject", 0, "NOTEMPTY", "DeviceNote", 0, "Trafo2", 1
            VerificarPropriedadeCondicional Obj, "TAPSPShow", 1, False, "TAPSPTag", 0, "Trafo2", 1
            VerificarPropriedadeVazia Obj, "CorOff", 0, "Trafo2", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal1", 0, "Trafo2", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal2", 0, "Trafo2", 1
            VerificarPropriedadeVazia Obj, "CorOnTerminal3", 0, "Trafo2", 1

        '-----------------------------------------------------------------------------
        Case "gx_AbnormalityIndicator"
            VerificarPropriedadeVazia Obj, "Measurement01Active", 0, "AbnormalityIndicator", 1
            VerificarPropriedadeValor Obj, "Measurement01Desc", "XXX", "AbnormalityIndicator", 0, 1, 1

        '-----------------------------------------------------------------------------
        Case "gx_Analogic"
            VerificarPropriedadeVazia Obj, "Measure", 0, "Analogic", 1

        '-----------------------------------------------------------------------------
        Case "gx_ButtonOpenCommandScreen"
            VerificarPropriedadeVazia Obj, "SourceObject", 0, "ButtonOpenCommandScreen", 1

        '-----------------------------------------------------------------------------
        Case "gx_Counter"
            VerificarPropriedadeVazia Obj, "Value", 0, "Counter", 1

        '-----------------------------------------------------------------------------
        Case "gx_CtrlDigital"
            VerificarPropriedadeCondicional Obj, "CommandPathName", 0, "NOTEMPTY", "Active", 0, "CtrlDigital", 1
            VerificarPropriedadeVazia Obj, "Active", 0, "CtrlDigital", 1
            VerificarPropriedadeValor Obj, "Descr", "Desc", "CtrlDigital", 1, 1, 1

        '-----------------------------------------------------------------------------
        Case "gx_CtrlDigital1Op"
            VerificarPropriedadeCondicional Obj, "CommandPathName", 0, "NOTEMPTY", "Active", 0, "CtrlDigital1Op", 1
            VerificarPropriedadeVazia Obj, "Tag", 0, "CtrlDigital1Op", 1
            VerificarPropriedadeValor Obj, "Descr", "Desc", "CtrlDigital1Op", 1, 1, 1

        '-----------------------------------------------------------------------------
        Case "gx_CtrlDigital2Op"
            VerificarPropriedadeCondicional Obj, "CommandPathName", 0, "NOTEMPTY", "Active", 0, "CtrlDigital2Op", 1
            VerificarPropriedadeVazia Obj, "Tag", 0, "CtrlDigital2Op", 1
            VerificarPropriedadeValor Obj, "Descr", "Desc", "CtrlDigital2Op", 1, 1, 1

        '-----------------------------------------------------------------------------
        Case "gx_CtrlDigital1Op"
            VerificarPropriedadeCondicional Obj, "CommandPathName", 0, "NOTEMPTY", "Active", 0, "CtrlDigital", 1
            VerificarPropriedadeVazia Obj, "Tag", 0, "CtrlDigital", 1
            VerificarPropriedadeValor Obj, "Descr", "Desc", "CtrlDigital", 1, 1, 1

        '-----------------------------------------------------------------------------
        Case "gx_CtrlDigital1Op"
            VerificarPropriedadeCondicional Obj, "CommandPathName", 0, "NOTEMPTY", "Active", 0, "CtrlDigital", 1
            VerificarPropriedadeVazia Obj, "Tag", 0, "CtrlDigital", 1
            VerificarPropriedadeValor Obj, "Descr", "Desc", "CtrlDigital", 1, 1, 1
        
        '-----------------------------------------------------------------------------
        Case "XCPump"
			VerificarPropriedadeVazia Obj, "SourceObject", 1, "Pump", 0
			' Se sourceobject vazio, enable=false
		'-----------------------------------------------------------------------------
        Case "iconElectricity"
			VerificarPropriedadeVazia Obj, "SourceObject", 1, "iconElectricity", 0
			
		'-----------------------------------------------------------------------------
		Case "iconComFail"
			VerificarPropriedadeVazia Obj, "SourceObject", 1, "iconComFail", 0
			
		'-----------------------------------------------------------------------------
		Case "xcLabel"
			VerificarPropriedadeVazia Obj, "Caption", 1, "Label", 1
			
		'-----------------------------------------------------------------------------
		Case "DrawString"
			VerificarPropriedadeVazia Obj, "Value", 1, "DrawString", 1
			VerificarPropriedadeVazia Obj, "TextColor", 0, "DrawString", 0
			
		'-----------------------------------------------------------------------------
		Case "xcEtiqueta_Manut"
			VerificarPropriedadeVazia Obj, "CorObjeto", 0, "Etiqueta_Manut", 0
			VerificarPropriedadeVazia Obj, "EtiquetaVisivel", 0, "Etiqueta_Manut", 0
			
		'-----------------------------------------------------------------------------
		Case "xcEtiqueta"
			VerificarPropriedadeVazia Obj, "AvisoVisivel", 0, "Etiqueta", 0
			VerificarPropriedadeVazia Obj, "EventoVisivel", 0, "Etiqueta", 0
			VerificarPropriedadeVazia Obj, "FonteObjeto", 0, "Etiqueta", 0
			VerificarPropriedadeVazia Obj, "ForaVisivel", 0, "Etiqueta", 0
			VerificarPropriedadeVazia Obj, "PathNote", 0, "Etiqueta", 0
			VerificarPropriedadeVazia Obj, "Visible", 0, "Etiqueta", 0
			
		'-----------------------------------------------------------------------------
		Case "xcWaterTank"
			VerificarPropriedadeVazia Obj, "objSource", 1, "WaterTank", 0
			VerificarPropriedadeVazia Obj, "objWaterDistribution", 1, "WaterTank", 0
			
		'-----------------------------------------------------------------------------
		Case "xcRetArea"
			'AdicionarErroTxt DadosTxt, "VerificarPropriedadesObjeto", Obj, "Objeto sem propriedades cadastradas para verificar: " & TypeName(Obj)
			
		'-----------------------------------------------------------------------------
		Case "XCArrow"
			VerificarPropriedadeVazia Obj, "Visible", 0, "Arrow", 0
			
		'-----------------------------------------------------------------------------
		Case "XCVerticalPipe"
			'AdicionarErroTxt DadosTxt, "VerificarPropriedadesObjeto", Obj, "Objeto sem propriedades cadastradas para verificar: " & TypeName(Obj)
			
		'-----------------------------------------------------------------------------
		Case "XCHorizontalPipe"
			'AdicionarErroTxt DadosTxt, "VerificarPropriedadesObjeto", Obj, "Objeto sem propriedades cadastradas para verificar: " & TypeName(Obj)
			
		'-----------------------------------------------------------------------------
		Case "XCDistribution"
			VerificarPropriedadeVazia Obj, "SourceObject", 0, "Distribution", 0
			
		'-----------------------------------------------------------------------------
		Case "XCSewage_Plant"
			'AdicionarErroTxt DadosTxt, "VerificarPropriedadesObjeto", Obj, "Objeto sem propriedades cadastradas para verificar: " & TypeName(Obj)
			
		'-----------------------------------------------------------------------------
		Case "IODriver"
			VerificarPropriedadeVazia Obj, "DriverLocation", 1, "IODriver", 0
			VerificarPropriedadeValor Obj, "WriteSyncMode", 1, "IODriver", 2, 1, 0
			
		'-----------------------------------------------------------------------------
		Case "patm_DeviceNote"
			VerificarPropriedadeVazia Obj, "AlarmSource", 1, "patm_DeviceNote", 1
			VerificarPropriedadeVazia Obj, "NoteDatabaseControl", 1, "patm_DeviceNote", 1
			VerificarPropriedadeVazia Obj, "SourceObject", 1, "patm_DeviceNote", 1
			VerificarPropriedadeVazia Obj, "NotePropertyValue", 1, "patm_DeviceNote", 1
			
		'-----------------------------------------------------------------------------
		Case "WaterDistributionNetwork"
			VerificarPropriedadeVazia Obj, "City", 1, "WaterDistributionNetwork", 0
			VerificarPropriedadeVazia Obj, "Company", 1, "WaterDistributionNetwork", 0
			VerificarPropriedadeVazia Obj, "CompanyAcronym", 1, "WaterDistributionNetwork", 0
			VerificarPropriedadeVazia Obj, "Contract", 1, "WaterDistributionNetwork", 0
			VerificarPropriedadeVazia Obj, "Name", 1, "WaterDistributionNetwork", 0
			VerificarPropriedadeVazia Obj, "Neighborhood", 1, "WaterDistributionNetwork", 0
			VerificarPropriedadeVazia Obj, "Organization", 1, "WaterDistributionNetwork", 0
			VerificarPropriedadeVazia Obj, "Region", 1, "WaterDistributionNetwork", 0
			VerificarPropriedadeVazia Obj, "State", 1, "WaterDistributionNetwork", 0
			VerificarPropriedadeVazia Obj, "StateAcronym", 1, "WaterDistributionNetwork", 0
			VerificarPropriedadeVazia Obj, "Note", 0, "WaterDistributionNetwork", 0

		'-----------------------------------------------------------------------------
		Case "Frame"

		'-----------------------------------------------------------------------------
		Case "InternalTag"
			VerificarPropriedadeVazia Obj, "Value", 0, "InternalTag", 0
			
			
		'-----------------------------------------------------------------------------
		Case "xoExecuteScheduler"
			VerificarPropriedadeVazia Obj, "aActivateCommandsGroup", 1, "xoExecuteScheduler", 0
			VerificarPropriedadeVazia Obj, "dteEndEvent", 1, "xoExecuteScheduler", 0
			VerificarPropriedadeVazia Obj, "dteEndRepeatDate", 1, "xoExecuteScheduler", 0
			VerificarPropriedadeVazia Obj, "dteNextEndEvent", 1, "xoExecuteScheduler", 0
			VerificarPropriedadeVazia Obj, "dteNextStartEvent", 1, "xoExecuteScheduler", 0
			VerificarPropriedadeVazia Obj, "objCommand", 1, "xoExecuteScheduler", 0
			VerificarPropriedadeVazia Obj, "strSchedulerName", 1, "xoExecuteScheduler", 0
			VerificarPropriedadeVazia Obj, "UserField01", 1, "xoExecuteScheduler", 0
			
		'-----------------------------------------------------------------------------
		Case "ext_ConfigTabularWater"
			
		'-----------------------------------------------------------------------------
		Case "manut_CreateNoteObjects"
			
		'-----------------------------------------------------------------------------
		Case "DBServer"
			VerificarPropriedadeValor Obj, "SourceType", 1, "DBServer", 2, 1, 0
			
		'-----------------------------------------------------------------------------
		Case "patm_XmlObj"
			
		'-----------------------------------------------------------------------------
		Case "AutoTester","PreencherPropriedade"
			
		'-----------------------------------------------------------------------------
		Case "frFooterRoot"
			
		'-----------------------------------------------------------------------------
		Case "Viewer"
			
		'-----------------------------------------------------------------------------
		Case "WaterConfig"
			VerificarPropriedadeVazia Obj, "ModelFile", 1, "WaterConfig", 0
			
		'-----------------------------------------------------------------------------
		Case "patm_CmdBoxXmlCreator"
			VerificarPropriedadeVazia Obj, "ConfigPower", 1, "CmdBoxXmlCreator", 0
			
		'-----------------------------------------------------------------------------
		Case "patm_CommandLogger"
			VerificarPropriedadeVazia Obj, "PowerConfigObj", 1, "CommandLogger", 0
			
		'-----------------------------------------------------------------------------
		Case "frCustomAlarmAndEventConfig"
			
		'-----------------------------------------------------------------------------
		Case "frThemeRoot"
			
		'-----------------------------------------------------------------------------
		Case "DemoTag"
			
		'-----------------------------------------------------------------------------
		Case "hpTheme01"
			
		'-----------------------------------------------------------------------------
		Case "hpXMLGenerateStruct"
			VerificarPropriedadeVazia Obj, "Log_BancoDeDados", 1, "BancoDados", 0
			
		'-----------------------------------------------------------------------------
		Case "hpXMLCatalog"
			
		'-----------------------------------------------------------------------------
		Case "xo_xmlDocStringInsert", "xo_1_Addr_Scan"
			
		'-----------------------------------------------------------------------------
		Case "hpXMLFilterVX"
			
		'-----------------------------------------------------------------------------
		Case "~hpImportXMLModel"
			
		'-----------------------------------------------------------------------------
		Case "hplog", "hpLogEvent"
			
		'-----------------------------------------------------------------------------
		Case "hpMultiMonitorConfig_DEPRECATED"
			
		'-----------------------------------------------------------------------------
		Case "hpPopupTemplate"
			
		'-----------------------------------------------------------------------------
		Case "~hpThemePublisher"
			
		'-----------------------------------------------------------------------------
		Case "~hpColorPalette"
			
		'-----------------------------------------------------------------------------
		Case "~hpBehaviorGroup"
			
		'-----------------------------------------------------------------------------
		Case "hpTheme"
			
		'-----------------------------------------------------------------------------
		Case "hpTranslatorController"
			
		'-----------------------------------------------------------------------------
		Case "gtwFrozenMeasurements"
			VerificarPropriedadeVazia Obj, "DateTag", 1, "gtwFrozenMeasurements", 0
			
		'-----------------------------------------------------------------------------
		Case "AlarmServer"
			VerificarPropriedadeVazia Obj, "DataSource", 1, "AlarmServer", 0
			
		'-----------------------------------------------------------------------------
		Case "xoFalhaOPC"
			
		'-----------------------------------------------------------------------------
		Case "E3Query"
			VerificarPropriedadeVazia Obj, "DataSource", 1, "E3Query", 0
			VerificarPropriedadeValor Obj, "QueryType", 1, "E3Query", 0, 1, 0
			
		'-----------------------------------------------------------------------------
		Case "CounterTag"
			
		'-----------------------------------------------------------------------------
		Case "aainfo_NoteController"
			VerificarPropriedadeVazia Obj, "DBServerPathName", 1, "NoteController", 0
			
		'-----------------------------------------------------------------------------
		Case "aainfoXcLibVersion"
			
		'-----------------------------------------------------------------------------
		Case "manut_ConfigShelveProperties"
			
		'-----------------------------------------------------------------------------
        Case Else
            ' Caso não haja tratamento específico
            AdicionarErroTxt DadosTxt, "VerificarPropriedadesObjeto", Obj, _
                "Tipo de objeto não tratado: " & TypeName(Obj)

    End Select
End Function

'***********************************************************************
'*
'*  Funções de Acesso às Propriedades
'*  -> 0 => GetPropertyLink
'*  -> 1 => GetPropertyValue
'*
'***********************************************************************
Function GetPropriedade(Obj, PropName, Metodo)
    If Metodo = 0 Then
        ' Usa link
        GetPropriedade = GetPropertyLink(Obj, PropName)
    Else
        ' Usa valor
        GetPropriedade = GetPropertyValue(Obj, PropName)
    End If
End Function

Function GetPropertyLink(Obj, PropName)
    On Error Resume Next
    Dim tmpValue
    tmpValue = Obj.Links.Item(PropName).Source
    If Err.Number <> 0 Then
        tmpValue = ""  ' Consideramos vazio se não existe link ou gerou erro
        Err.Clear
    End If
    On Error GoTo 0

    GetPropertyLink = tmpValue
End Function

Function GetPropertyValue(Obj, PropName)
    ' Tenta ler via Eval, assumindo que seja uma property do objeto
    On Error Resume Next
    Dim tmpValue
    tmpValue = Eval("Obj." & PropName)
    If Err.Number <> 0 Then
        tmpValue = ""   ' Se der erro, consideramos vazio
        Err.Clear
    End If
    On Error GoTo 0

    GetPropertyValue = CStr(tmpValue)
End Function

'***********************************************************************
'*
'*  Funções de verificação de propriedades
'*  Ajustamos para receber também o “metodo”
'*  -> 0 => Link
'*  -> 1 => Value
'*
'***********************************************************************
Function VerificarPropriedadeVazia(Obj, Propriedade, Metodo, NomeObjeto, Classificacao)
    On Error Resume Next

    Dim ValorLeitura
    ValorLeitura = GetPropriedade(Obj, Propriedade, Metodo)

    If Trim(ValorLeitura) = "" Then
        ' Está efetivamente vazio (ou erro) => registramos
        AdicionarErroExcel DadosExcel, Obj.PathName, CStr(Classificacao), _
                           NomeObjeto & " com " & Propriedade & " vazia"
    End If
End Function

Function VerificarPropriedadeCondicional(Obj, PropCond, MetodoCond, ValorEsperado, _
                                         PropVerif, MetodoVerif, NomeObjeto, TipoProblema)
    On Error Resume Next

    Dim ValorCondicional, ValorAVerificar
    ValorCondicional = GetPropriedade(Obj, PropCond, MetodoCond)
    ValorAVerificar  = GetPropriedade(Obj, PropVerif, MetodoVerif)

    If ValorEsperado = "NOTEMPTY" Then
        ' Se a condição é "NOTEMPTY", significa que precisamos
        ' verificar se ValorCondicional não está vazio.
        If Trim(ValorCondicional) <> "" Then
            ' Se a propriedade a verificar está vazia, registramos erro.
            If Trim(ValorAVerificar) = "" Then
                AdicionarErroExcel DadosExcel, Obj.PathName, CStr(TipoProblema), _
                    NomeObjeto & " com " & PropVerif & _
                    " vazia enquanto " & PropCond & " está preenchida"
            End If
        End If
    ElseIf CStr(ValorCondicional) = CStr(ValorEsperado) Then
        ' Caso comum: se ValorCondicional == ValorEsperado, então checamos se ValorAVerificar está vazio
        If Trim(ValorAVerificar) = "" Then
            AdicionarErroExcel DadosExcel, Obj.PathName, CStr(TipoProblema), _
                NomeObjeto & " com " & PropVerif & _
                " vazia enquanto " & PropCond & " está " & ValorEsperado
        End If
    End If

    ' Se ocorreu erro no acesso às propriedades
    If Err.Number <> 0 Then
        AdicionarErroTxt DadosTxt, "VerificarPropriedadeCondicional", Obj, _
            "Erro ao acessar " & PropCond & " ou " & PropVerif
        Err.Clear
    End If

    On Error GoTo 0
End Function

Function VerificarBancoDeDados(Obj, CampoBD, MetodoBD, NomeObjeto, Classificacao)
    On Error Resume Next

    Dim ValorBD
    ValorBD = GetPropriedade(Obj, CampoBD, MetodoBD)

    If Trim(ValorBD) <> "" Then
        If Not DadosBancoDeDados.Exists(ValorBD) Then
            DadosBancoDeDados.Add ValorBD, Obj.PathName
        Else
            AdicionarErroExcel DadosExcel, Obj.PathName, CStr(Classificacao), _
                NomeObjeto & " compartilhando BD '" & ValorBD & "' com " & DadosBancoDeDados(ValorBD)
        End If
    End If

    If Err.Number <> 0 Then
        AdicionarErroTxt DadosTxt, "VerificarBancoDeDados", Obj, _
            "Erro ao acessar " & CampoBD & " em " & NomeObjeto
        Err.Clear
    End If
End Function

Function VerificarPropriedadeHabilitada(Obj, Propriedade, MetodoProp, NomeObjeto, Esperado, Classificacao)
    On Error Resume Next

    Dim ValorAtual
    ValorAtual = GetPropriedade(Obj, Propriedade, MetodoProp)

    If CStr(ValorAtual) <> CStr(Esperado) Then
        AdicionarErroExcel DadosExcel, Obj.PathName, CStr(Classificacao), _
            NomeObjeto & " com " & Propriedade & " diferente do esperado (Esperado: " & Esperado & ", Atual: " & ValorAtual & ")"
    End If

    If Err.Number <> 0 Then
        AdicionarErroTxt DadosTxt, "VerificarPropriedadeHabilitada", Obj, _
            "Erro ao acessar " & Propriedade & " em " & NomeObjeto
        Err.Clear
    End If
End Function

Function VerificarPropriedadeValor(Obj, Propriedade, MetodoProp, NomeObjeto, ValorEsperado, Classificacao, ModoComparacao)
    ' ModoComparacao:
    '   0 => "igual"
    '   1 => "diferente"
    
    On Error Resume Next
    
    Dim ValorAtual, ValorAtualStr, ValorEsperadoStr
    ValorAtual = GetPropriedade(Obj, Propriedade, MetodoProp)  ' Lê via link ou property, conforme MetodoProp
    
    ' Converte ambos para string para comparação "básica"
    ValorAtualStr      = CStr(ValorAtual)
    ValorEsperadoStr   = CStr(ValorEsperado)
    
    Select Case ModoComparacao
        Case 0 ' "igual"
            ' Se ValorAtual for diferente do esperado, gera log
            If ValorAtualStr <> ValorEsperadoStr Then
                AdicionarErroExcel DadosExcel, Obj.PathName, Classificacao, _
                    NomeObjeto & " com " & Propriedade & " diferente do valor esperado: " & _
                    "(Esperado: " & ValorEsperadoStr & ", Atual: " & ValorAtualStr & ")"
            End If
        
        Case 1 ' "diferente"
            ' Se ValorAtual for igual ao esperado, gera log
            If ValorAtualStr = ValorEsperadoStr Then
                AdicionarErroExcel DadosExcel, Obj.PathName, Classificacao, _
                    NomeObjeto & " com " & Propriedade & " igual ao valor que deveria ser diferente: " & _
                    "(Valor: " & ValorAtualStr & ")"
            End If
        
        Case Else
            ' Qualquer outro valor de ModoComparacao não é reconhecido
            AdicionarErroTxt DadosTxt, "VerificarPropriedadeValor", Obj, _
                "ModoComparacao inválido: " & ModoComparacao & " para propriedade " & Propriedade
    End Select

    If Err.Number <> 0 Then
        AdicionarErroTxt DadosTxt, "VerificarPropriedadeValor", Obj, _
            "Erro ao acessar " & Propriedade & " em " & NomeObjeto
        Err.Clear
    End If
    
    On Error GoTo 0
End Function

'********************************************************************************
' Nome: VerificarPropriedadeTextoProibido
' Objetivo: Verificar se a propriedade (via link ou valor) contém um texto proibido.
'
' Parâmetros:
'   Obj            -> Objeto a verificar (ex.: pwa_Disjuntor, pwa_BarraAlarme)
'   Propriedade    -> Nome da propriedade (ex.: "SourceObject")
'   MetodoProp     -> 0 => Link (GetPropertyLink), 1 => Valor (GetPropertyValue)
'   NomeObjeto     -> Rótulo para o log (ex.: "pwa_Disjuntor")
'   TextoProibido  -> Texto que não deve aparecer (ex.: ".Value")
'   Classificacao  -> Código de severidade no Excel (0=Aviso, 1=Erro, etc.)
'********************************************************************************
Function VerificarPropriedadeTextoProibido(Obj, Propriedade, MetodoProp, _
                                           NomeObjeto, TextoProibido, Classificacao)
    On Error Resume Next

    ' 1) Ler a propriedade via Link ou Valor (Eval)
    Dim ValorAtual
    ValorAtual = GetPropriedade(Obj, Propriedade, MetodoProp)

    ' 2) Se conter o TextoProibido, registramos erro/aviso
    If InStr(1, ValorAtual, TextoProibido, vbTextCompare) > 0 Then
        ' Exemplo de mensagem: "SourceObject usa '.Value'"
        Dim mensagem
        mensagem = "A propriedade " & Propriedade & " não deve conter '" & TextoProibido & "'. " & _
                   "(Atual: " & ValorAtual & ")"

        AdicionarErroExcel DadosExcel, Obj.PathName, CStr(Classificacao), _
            NomeObjeto & " -> " & mensagem
    End If

    ' 3) Capturar eventuais erros no acesso
    If Err.Number <> 0 Then
        AdicionarErroTxt DadosTxt, "VerificarPropriedadeTextoProibido", Obj, _
            "Erro ao acessar " & Propriedade & " em " & NomeObjeto
        Err.Clear
    End If
    On Error GoTo 0
End Function


Function VerificarObjetoDesatualizado(Obj, NomeAntigo, NovaBiblioteca)
    ' Exemplo de mensagem:
    ' "O objeto pwa_Gerador é obsoleto e deve ser substituído por generic_automalogica."

    Dim Mensagem
    Mensagem = "O objeto " & NomeAntigo & _
               " é obsoleto e deve ser substituído pela biblioteca " & NovaBiblioteca & "."
    AdicionarErroExcel DadosExcel, Obj.PathName, "1", Mensagem
End Function

'***********************************************************************
'*
'*  Função:       AdicionarErroTxt
'*  Objetivo:     Adicionar texto de erro ao dicionário DadosTxt
'*                para posterior geração de log TXT.
'*
'***********************************************************************
Function AdicionarErroTxt(DadosTxt, NomeSub, Obj, DescricaoErro)
    On Error Resume Next

    Dim LinhaTxt
    LinhaTxt = DadosTxt.Count + 1

    Dim keyTxt
    keyTxt = CStr(LinhaTxt)
    While DadosTxt.Exists(keyTxt)
        LinhaTxt = LinhaTxt + 1
        keyTxt = CStr(LinhaTxt)
    Wend

    If Not IsObject(DadosTxt) Then
        MsgBox "Erro: O dicionário DadosTxt não foi inicializado.", vbCritical
        Exit Function
    End If

    Dim MensagemErro
    MensagemErro = "Erro na Sub " & NomeSub & "/" & Obj.PathName & ": " & DescricaoErro
    DadosTxt.Add keyTxt, MensagemErro
End Function

'***********************************************************************
'*
'*  Função:       GerarRelatorioTxt
'*  Objetivo:     Gera um arquivo TXT com base no dicionário DadosTxt
'*
'***********************************************************************
Function GerarRelatorioTxt(DadosTxt, CaminhoPrj)
    On Error GoTo 0

    If Not DadosTxt.Exists(CStr(1)) Then
        MsgBox "Nenhum dado disponível para gerar o relatório TXT.", vbExclamation
        GerarRelatorioTxt = False
        Exit Function
    End If

    Dim NomeTxt
    NomeTxt = CaminhoPrj & "\Log_" & Replace(Replace(Date() & "_" & Time(), ":", "_"), "/", "_") & ".txt"

    Dim FSO, ArquivoTxt, Linha
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ArquivoTxt = FSO.CreateTextFile(NomeTxt, True)

    For Each Linha In DadosTxt
        ArquivoTxt.WriteLine DadosTxt.Item(Linha)
    Next
    ArquivoTxt.Close

    Dim Resposta, ShellObj
    Resposta = MsgBox("Foram gerados logs de erro de código. Deseja abrir o arquivo?", vbYesNo + vbQuestion, "AutomaTester")
    If Resposta = vbYes Then
        Set ShellObj = CreateObject("WScript.Shell")
        ShellObj.Run """" & NomeTxt & """"
        Set ShellObj = Nothing
    End If

    GerarRelatorioTxt = True
    Exit Function
End Function

'***********************************************************************
'*
'*  Função:       AdicionarErroExcel
'*  Objetivo:     Adicionar uma linha de erro ao dicionário DadosExcel
'*
'***********************************************************************
Function AdicionarErroExcel(DadosExcel, CaminhoObjeto, ClassificacaoCode, Mensagem)
    On Error Resume Next

    Dim LinhaExcel
    LinhaExcel = DadosExcel.Count + 1

    Dim keyExcel
    keyExcel = CStr(LinhaExcel)
    While DadosExcel.Exists(keyExcel)
        LinhaExcel = LinhaExcel + 1
        keyExcel = CStr(LinhaExcel)
    Wend

    If Not IsObject(DadosExcel) Then
        MsgBox "Erro: O dicionário DadosExcel não foi inicializado.", vbCritical
        Exit Function
    End If

    If Len(Trim(CaminhoObjeto)) > 0 And Len(Trim(ClassificacaoCode)) > 0 And Len(Trim(Mensagem)) > 0 Then
        DadosExcel.Add keyExcel, CaminhoObjeto & "/" & ClassificacaoCode & "/" & Mensagem
    Else
        MsgBox "Erro: Valores inválidos ao adicionar ao Excel:" & vbCrLf & _
               "CaminhoObjeto: " & CaminhoObjeto & vbCrLf & _
               "ClassificacaoCode: " & ClassificacaoCode & vbCrLf & _
               "Mensagem: " & Mensagem, vbCritical
    End If
    On Error GoTo 0
End Function

'***********************************************************************
'*
'*  Função:       GerarRelatorioExcel
'*  Objetivo:     Gera um arquivo Excel com base no dicionário DadosExcel
'*                com colunas: Objeto / Tipo / Problema
'*
'***********************************************************************
Function GerarRelatorioExcel(DadosExcel, CaminhoPrj)
    On Error GoTo 0

    If DadosExcel.Count = 0 Then
        MsgBox "Nenhum dado disponível para gerar o relatório Excel.", vbExclamation
        GerarRelatorioExcel = False
        Exit Function
    End If

    Dim NomeExcel
    NomeExcel = CaminhoPrj & "\RelatorioTester_" & Replace(Replace(Date() & "_" & Time(), ":", "_"), "/", "_") & ".xlsx"

    Dim objExcel, objWorkBook, sheet, Linha
    Set objExcel = CreateObject("EXCEL.APPLICATION")
    Set objWorkBook = objExcel.Workbooks.Add
    Set sheet = objWorkBook.Sheets(1)

    sheet.Cells(1, 1).Value = "Objeto"
    sheet.Cells(1, 2).Value = "Tipo"
    sheet.Cells(1, 3).Value = "Problema"

    sheet.Cells(1, 1).Font.Bold = True
    sheet.Cells(1, 2).Font.Bold = True
    sheet.Cells(1, 3).Font.Bold = True

    Dim celulas
    For Each Linha In DadosExcel
        celulas = Split(DadosExcel.Item(Linha), "/")
        
        If UBound(celulas) >= 2 Then
            Dim classificationCode, classificationText
            classificationCode = celulas(1)

            Select Case classificationCode
                Case "0"
                    classificationText = "Aviso"
                Case "1"
                    classificationText = "Erro"
                Case "2"
                    classificationText = "Revisar"
                Case Else
                    classificationText = "Desconhecido"
            End Select

            sheet.Cells(CInt(Linha) + 1, 1).Value = celulas(0)
            sheet.Cells(CInt(Linha) + 1, 2).Value = classificationText
            sheet.Cells(CInt(Linha) + 1, 3).Value = celulas(2)
        End If
    Next

    objWorkBook.SaveAs NomeExcel
    objWorkBook.Close
    objExcel.Quit
    Set objWorkBook = Nothing
    Set objExcel = Nothing

    Dim Resposta, ShellObj
    Resposta = MsgBox("Foram gerados logs de correção. Deseja abrir o arquivo?", vbYesNo + vbQuestion, "AutomaTester")
    If Resposta = vbYes Then
        Set ShellObj = CreateObject("WScript.Shell")
        ShellObj.Run """" & NomeExcel & """"
        Set ShellObj = Nothing
    End If

    GerarRelatorioExcel = True
    Exit Function
End Function

Sub Fim()
End Sub