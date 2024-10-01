Sub AutoTester_CustomConfig()
    ' Script de verificações automáticas no domínio
    Dim Resposta
    Resposta = MsgBox("Tem certeza que deseja iniciar o teste automático do domínio?", vbYesNo + vbQuestion, "Iniciar teste de domínio?")
    If Resposta = vbNo Then
        Exit Sub
    End If
    Main()
End Sub

' Global Variables
Dim DadosExcel, DadosTxt, DadosBancoDeDados, ListaObjetosLib
Dim nomeExcel, nomeTxt, Linha, LinhaTxt
Dim CaminhoPrj

' Initialize Variables
Set DadosExcel = CreateObject("Scripting.Dictionary")
Set DadosTxt = CreateObject("Scripting.Dictionary")
Set DadosBancoDeDados = CreateObject("Scripting.Dictionary")
Set ListaObjetosLib = CreateObject("Scripting.Dictionary")

nomeExcel = Replace(Replace(Date() & "_" & Time(), ":", "_"), "/", "_")
nomeTxt = nomeExcel

Linha = 1
LinhaTxt = 1

' Get Project Path
If PastaParaSalvarLogs <> "" Then
    CaminhoPrj = PastaParaSalvarLogs
Else
    CaminhoPrj = CreateObject("WScript.Shell").CurrentDirectory
End If

Sub Main()
    Dim telaArray, ScreenObj
    telaArray = SplitTelas(PathNameTelas)
    
    If UBound(telaArray) >= 0 Then
        ' Filtrar por telas específicas
        For Each ScreenObj In Application.ListFiles("Screen")
            If IsTelaNaLista(ScreenObj.PathName, telaArray) Then
                VerificarTela ScreenObj
            End If
        Next
    Else
        ' Verificar todas as telas
        For Each ScreenObj In Application.ListFiles("Screen")
            VerificarTela ScreenObj
        Next
    End If

    ' Verificação de uso dos bancos de dados e historiadores
    If VerificarBancosCustom Then
        Dim tiposArray
        tiposArray = Array("DataServer", "Hist")
        VerificarObjetos tiposArray
    End If

    ' Gerar relatórios
    GerarRelatorios DebugMode
    MsgBox "Fim"
End Sub

'----------------- Funções Auxiliares -----------------

' Função para dividir telas ou retornar array vazio
Function SplitTelas(PathNameTelas)
    If Len(Trim(PathNameTelas)) > 0 Then
        SplitTelas = Split(PathNameTelas, "/")
    Else
        SplitTelas = Array()
    End If
End Function

' Função para verificar se uma tela está na lista
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

' Função para verificar objetos (DataServers e Historiadores)
Sub VerificarObjetos(tipos)
    Dim tipo, obj
    For Each tipo In tipos
        For Each obj In Application.ListFiles(tipo)
            If tipo = "DataServer" Then
                FiltrarXObjectsDominio obj
            ElseIf tipo = "Hist" Then
                VerificarHistoriadores obj
            End If
        Next
    Next
End Sub

Sub GerarRelatorios(DebugMode)
    If DebugMode Then
            GerarRelatorioExcel
            GerarRelatorioTxt
    Else
            GerarRelatorioExcel
    End If
End Sub

'----------------- Chamado das Verificações -----------------

Sub VerificarTela(ParentObj)
    Dim Eletricos, Mecanicos, Energizacao, Objeto
    Eletricos = Array("Disjuntor", "Seccionadora", "Trafo", "Gerador", "Chave", "Switch")
    Mecanicos = Array("Bomb", "Valve", "Brake")
    Energizacao = Array("Disjuntor", "Seccionadora", "Line")
    
    If Not UsandoLibControle Then
        InfoAlarmGenericLib ParentObj
        InfoAnalogicaGenericLib ParentObj
    End If
    
    InfoAlarmSourceObject ParentObj
    VerificarCaptionTela ParentObj
    VerificarBotaoAbreTela ParentObj
    InfoAlarmComValue ParentObj
    InfoAnalogicaSemSourceObject ParentObj
    VerificarSPShowInfoAnalogic ParentObj
    CorBackgroundTela ParentObj
    VerificaAlarmBar ParentObj
    InfoAlarmDivergeDescricao ParentObj
    VerificarMesmoObjetoLibsDiferentes
    
    For Each Objeto In Eletricos
        ClassificarLibEletricos ParentObj, Objeto
    Next
    For Each Objeto In Mecanicos
        ClassificarLibMecanicos ParentObj, Objeto
    Next
    For Each Objeto In Energizacao
        VerificarCorEnergizacao ParentObj, Objeto
    Next
End Sub

' Sub para filtrar objetos no domínio
Sub FiltrarXObjectsDominio(DataServer)
    Dim Object
    For Each Object In DataServer
        Select Case TypeName(Object)
            Case "DataServer", "DataFolder"
                FiltrarXObjectsDominio Object
            Case "frCustomAppConfig"
                VerificarBancoDeDados Object.AppDBServerPathName, Object.PathName, Object.Name
            Case "ww_Parameters"
                VerificarBancoDeDados Object.DBServer, Object.PathName, Object.Name
            Case "DatabaseTags_Parameters"
                VerificarBancoDeDados Object.StorageMethod, Object.PathName, Object.Name
            Case "patm_CmdBoxXmlCreator"
                VerificarBancoDeDados Object.DBServerPathName, Object.PathName, Object.Name
            Case "patm_NoteDatabaseControl"
                VerificarBancoDeDados Object.DBServer, Object.PathName, Object.Name
            Case "patm_xoAlarmHistConfig"
                VerificarBancoDeDados Object.MainDBServerPathName, Object.PathName, Object.Name
        End Select
    Next
End Sub

' Sub para verificar historiadores
Sub VerificarHistoriadores(Historiadores)
    Dim Hist
    For Each Hist In Historiadores
        Select Case TypeName(Hist)
            Case "DataFolder"
                VerificarHistoriadores Hist
            Case "Hist"
                VerificarHist Hist.DBServer, Hist.PathName, Hist.Name
        End Select
    Next
End Sub
'----------------- Funções de Verificação das telas -----------------

Sub ClassificarLibEletricos(Tela, Objeto)
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then 'Faz o código procurar dentro de grupos
            ClassificarLibEletricos Obj, Objeto
        End If
        If InStr(1, TypeNameObj, Objeto, 1) > 0 Then
            If Left(TypeName(Obj), 2) = "xc" Then
                Lib = Left(TypeName(Obj), 2)
            ElseIf Left(TypeName(Obj), 4) = "arch" Then
                Lib = Left(TypeName(Obj), 4)
            ElseIf TypeName(Obj) = "Switch" Then
                Lib = "powercontrols"
            Else
                On Error Resume Next
                Lib = Left(TypeName(Obj), InStr(1, TypeName(Obj), "_", 1) - 1)
                If Err.Number <> 0 Then
                    AdicionarTxt DadosTxt, LinhaTxt, "Definindo Lib", Obj, Err.Description
                    Err.Clear
                End If
            End If
            Select Case True
                Case(Lib = "xc") And (Objeto = "Disjuntor")
                    ObjetoMecanicoSupervisaoXC Tela, Obj, Objeto
                Case(Lib = "xc") And (Objeto = "Gerador")
                    'Verificar o que deve ser linkado com gerador
                Case(Lib = "xc") And (Objeto = "Trafo")
                    'Nada a colocar aqui, a menos que um link em energizado seja obrigatório
                Case(Lib = "xc") And (Objeto = "NotaOperacional")
                    ObjetoLibXCNotaOperacional Tela, Obj, Objeto
                Case(Lib = "xc") And (Objeto = "Switch")
                    'Nada pra fazer por enquanto
                Case(Lib = "xc") And (Objeto = "Chave")
                    ObjetoEletricoSemSourceObject Tela, Obj, Objeto
                Case(Lib = "arch") And (Objeto = "Switch")
                    'Nada pra fazer por enquanto
                Case(Lib = "pwa") And (Objeto = "Disjuntor")
                    ObjetoEletricoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos eletricos que abrem tela de comando
                    ObjetoEletricoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos eletricos sem tela mas com DeviceNote
                    ObjetoEletricoSemSourceObject Tela, Obj, Objeto
                Case(Lib = "pwa") And (Objeto = "Switch")
                    ObjetoEletricoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos eletricos que abrem tela de comando
                    ObjetoEletricoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos eletricos sem tela mas com DeviceNote
                    ObjetoEletricoSemSourceObject Tela, Obj, Objeto
                Case(Lib = "pwa") And (Objeto = "Seccionadora")
                    ObjetoEletricoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos eletricos que abrem tela de comando
                    ObjetoEletricoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos eletricos sem tela mas com DeviceNote
                    ObjetoEletricoSemSourceObject Tela, Obj, Objeto
                Case(Lib = "pwa") And (Objeto = "Trafo")
                    ObjetoEletricoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos eletricos que abrem tela de comando
                    ObjetoEletricoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos eletricos sem tela mas com DeviceNote
                Case(Lib = "pwa") And (Objeto = "Gerador")
                    ObjetoEletricoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos eletricos que abrem tela de comando
                    ObjetoEletricoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos eletricos sem tela mas com DeviceNote
                Case(Lib = "powercontrols") And (Objeto = "Switch")
                    'Nada pra fazer por enquanto
                Case Else
                    MsgBox "Lib: " & Lib & " " & Objeto & " Não cadastrada como elétrico, consulte equipe de testers"
            End Select
        End If
    Next
End Sub

Sub ClassificarLibMecanicos(Tela, Objeto)
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then 'Faz o código procurar dentro de grupos
            ClassificarLibMecanicos Obj, Objeto
        End If
        If InStr(1, TypeNameObj, Objeto, 1) > 0 Then
            If Left(TypeName(Obj), 2) = "xc" Then
                Lib = Left(TypeName(Obj), 2)
            Else
                Lib = Left(TypeName(Obj), InStr(1, TypeName(Obj), "_", 1) - 1)
            End If
            Select Case True
                Case(Lib = "xc") And (Objeto = "Bomb")
                    ObjetoBombaSupervisionadaXC Tela, Obj, Objeto
                Case(Lib = "uhe") And (Objeto = "Valve")
                    ObjetoMecanicoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos mecanicos que abrem tela de comando
                    ObjetoMecanicoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos mecanicos sem tela mas com DeviceNote
                    ObjetoMecanicoSemSourceObject Tela, Obj, Objeto ' Verifica se tem objetos mecanicos sem SourceObject preenchido e supervisionado
                    ConferirLinkObjetosMecanicos Tela, Obj, Objeto ' Verifica os equipamentos mecânicos que são supervisionados estão linkados
                Case(Lib = "uhe") And (Objeto = "Bomb")
                    ObjetoMecanicoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos mecanicos que abrem tela de comando
                    ObjetoMecanicoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos mecanicos sem tela mas com DeviceNote
                    ObjetoMecanicoSemSourceObject Tela, Obj, Objeto ' Verifica se tem objetos mecanicos sem SourceObject preenchido e supervisionado
                    ConferirLinkObjetosMecanicos Tela, Obj, Objeto ' Verifica os equipamentos mecânicos que são supervisionados estão linkados
                Case(Lib = "uhe") And (Objeto = "Brake")
                    ObjetoMecanicoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos mecanicos que abrem tela de comando
                    ObjetoMecanicoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos mecanicos sem tela mas com DeviceNote
                    ObjetoMecanicoSemSourceObject Tela, Obj, Objeto ' Verifica se tem objetos mecanicos sem SourceObject preenchido e supervisionado
                    ConferirLinkObjetosMecanicos Tela, Obj, Objeto ' Verifica os equipamentos mecânicos que são supervisionados estão linkados
                Case Else
                    MsgBox "Lib: " & Lib & " " & Objeto & " Não cadastrada como mecânico, consulte equipe de testers"
            End Select
        End If
    Next
End Sub

Sub VerificarCorEnergizacao(Tela, Objeto)
    On Error Resume Next
    Set ObjetosIgnorados = CreateObject("Scripting.Dictionary")
    ObjetosIgnorados.Add "archLineVertical", Empty
    ObjetosIgnorados.Add "archLineHorizontal", Empty
    ObjetosIgnorados.Add "DrawLine", Empty
    ObjetosIgnorados.Add "DisjuntoresERAC", Empty
    ObjetosIgnorados.Add "EquivLine", Empty
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then
            VerificarCorEnergizacao Obj, Objeto
        ElseIf InStr(1, TypeNameObj, Objeto, 1) > 0 And ( Not ObjetosIgnorados.Exists(TypeNameObj)) Then
            On Error Resume Next
            If (Obj.Links.Item("CorOn").Source = "") Then
                AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", "Propriedade CorOn está vazia"
            End If
            If (Obj.Links.Item("CorOff").Source = "") Then
                AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", "Propriedade CorOff está vazia"
            End If
        End If
        On Error GoTo 0
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
            ListaObjetosLib.Add TypeNameObj, Empty
        End If
        If Err.Number <> 0 Then
            AdicionarTxt DadosTxt, LinhaTxt, "VerificarCorEnergizacao", Obj, Err.Description & " Se este erro se repetir no excel ele pode ser ignorado"
        End If
    Next
End Sub

' Sub para verificar a descrição do alarme
Sub InfoAlarmDivergeDescricao(Tela)
    On Error Resume Next
    Dim Obj, TypeNameObj, sourceObject, descricao, areaAlarme
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If TypeNameObj = "DrawGroup" Then
            InfoAlarmDivergeDescricao Obj
        End If
        If Left(TypeName(Obj), 2) = "xc" And InStr(1, TypeNameObj, "InfoAlarme") > 0 Then
            areaAlarme = Obj.AreaAlarme
            descricao = Obj.Descricao
            If Len(Trim(areaAlarme)) <> 0 Then
                If InStr(1, areaAlarme, descricao, vbTextCompare) = 0 Then
                    AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", "Descrição do objeto contém um texto diferente da associação"
                End If
            End If
        ElseIf InStr(1, TypeNameObj, "InfoAlarme") > 0 Then
            sourceObject = Obj.SourceObject01
            descricao = Obj.Descricao
            If Len(Trim(sourceObject)) <> 0 Then
                If InStr(1, sourceObject, descricao, vbTextCompare) = 0 Then
                    AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", "Descrição do objeto contém um texto diferente da associação"
                End If
            End If
        End If
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
            ListaObjetosLib.Add TypeNameObj, Empty
        End If
        If Err.Number <> 0 Then
            AdicionarTxt DadosTxt, LinhaTxt, "InfoAlarmDivergeDescricao", Obj, Err.Description
            Err.Clear
        End If
    Next
    On Error GoTo 0
End Sub

' Sub para verificar barras de alarme
Sub VerificaAlarmBar(Tela)
    On Error Resume Next
    Dim ObjetosIgnorados, Obj, TypeNameObj
    Set ObjetosIgnorados = CreateObject("Scripting.Dictionary")
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If TypeNameObj = "DrawGroup" Then
            VerificaAlarmBar Obj
        ElseIf InStr(1, TypeNameObj, "AlarmBar") > 0 And Not ObjetosIgnorados.Exists(TypeNameObj) Then
            If Obj.NaoSupervisionado = False And Obj.Measure = "" Then
                AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Propriedade Measure está vazia"
            ElseIf Obj.NaoSupervisionado = True And Obj.Measure <> "" Then
                AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", "Propriedade Measure está preenchida mesmo com objeto não supervisionado"
            End If
        End If
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
            ListaObjetosLib.Add TypeNameObj, Empty
        End If
        If Err.Number <> 0 Then
            AdicionarTxt DadosTxt, LinhaTxt, "VerificaAlarmBar", Obj, Err.Description
            Err.Clear
        End If
    Next
    On Error GoTo 0
End Sub

' Sub para verificar InfoAlarms com SourceObject01 vazio
Sub InfoAlarmSourceObject(Tela)
    On Error Resume Next
    Dim Obj, TypeNameObj, Lib
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If TypeNameObj = "DrawGroup" Then
            InfoAlarmSourceObject Obj
        End If
        If Left(TypeName(Obj), 2) = "xc" Then
            Lib = Left(TypeName(Obj), 2)
        Else
            Lib = ""
        End If
        If Lib = "xc" And InStr(1, TypeNameObj, "InfoAlarme") > 0 Then
            If Obj.AreaAlarme = "" Then
                AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Sem AreaAlarme"
            End If
        ElseIf InStr(1, TypeNameObj, "InfoAlarme") > 0 Then
            If Obj.SourceObject01 = "" Then
                AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Propriedade SourceObject01 em branco"
            End If
        End If
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
            ListaObjetosLib.Add TypeNameObj, Empty
        End If
        If Err.Number <> 0 Then
            AdicionarTxt DadosTxt, LinhaTxt, "InfoAlarmSourceObject", Obj, Err.Description
            Err.Clear
        End If
    Next
    On Error GoTo 0
End Sub

' Sub para verificar InfoAlarms usando lib antiga
Sub InfoAlarmGenericLib(Tela)
    On Error Resume Next
    Dim Obj, TypeNameObj
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If TypeNameObj = "DrawGroup" Then
            InfoAlarmGenericLib Obj
        End If
        If InStr(1, TypeNameObj, "InfoAlarme") > 0 Then
            If Left(TypeNameObj, InStr(1, TypeNameObj, "_")) <> "gx" Then
                AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", "Objeto com a lib de InfoAlarm antiga, recomenda-se usar a generic"
            End If
        End If
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
            ListaObjetosLib.Add TypeNameObj, Empty
        End If
        If Err.Number <> 0 Then
            AdicionarTxt DadosTxt, LinhaTxt, "InfoAlarmGenericLib", Obj, Err.Description
            Err.Clear
        End If
    Next
    On Error GoTo 0
End Sub

' Sub para verificar InfoAlarms com .Value no SourceObject
Sub InfoAlarmComValue(Tela)
    On Error Resume Next
    Dim Obj, TypeNameObj, Lib, i, SourceObjectxx
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If TypeNameObj = "DrawGroup" Then
            InfoAlarmComValue Obj
        End If
        If Left(TypeName(Obj), 2) = "xc" Then
            Lib = Left(TypeName(Obj), 2)
        Else
            Lib = ""
        End If
        If Lib = "xc" And InStr(1, TypeNameObj, "InfoAlarme") > 0 Then
            ' Não verifica para lib xc
        ElseIf InStr(1, TypeNameObj, "InfoAlarme") > 0 Then
            For i = 1 To CInt(Right(TypeNameObj, 2))
                If i < 10 Then
                    i = "0" & CStr(i)
                End If
                Execute "SourceObjectxx = Obj.SourceObject" & CStr(i)
                If InStr(1, SourceObjectxx, ".Value") > 0 Then
                    AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Objeto com .Value no SourceObject"
                End If
            Next
        End If
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
            ListaObjetosLib.Add TypeNameObj, Empty
        End If
        If Err.Number <> 0 Then
            AdicionarTxt DadosTxt, LinhaTxt, "InfoAlarmComValue", Obj, Err.Description
            Err.Clear
        End If
    Next
    On Error GoTo 0
End Sub

' Sub para verificar InfoAnalogics sem SourceObject
Sub InfoAnalogicaSemSourceObject(Tela)
    On Error Resume Next
    Dim Obj, TypeNameObj, Lib
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If TypeNameObj = "DrawGroup" Then
            InfoAnalogicaSemSourceObject Obj
        End If
        If Left(TypeName(Obj), 2) = "xc" Then
            Lib = Left(TypeName(Obj), 2)
        Else
            Lib = ""
        End If
        If Lib = "xc" And InStr(1, TypeNameObj, "InfoAnalogica") > 0 Then
            If Obj.ValueTag = "" Then
                AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "InfoAnalogica sem ValueTag"
            End If
        ElseIf InStr(1, TypeNameObj, "InfoAnalogica") > 0 Then
            If Obj.SourceObject = "" Then
                AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "InfoAnalogica sem SourceObject"
            End If
        End If
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
            ListaObjetosLib.Add TypeNameObj, Empty
        End If
        If Err.Number <> 0 Then
            AdicionarTxt DadosTxt, LinhaTxt, "InfoAnalogicaSemSourceObject", Obj, Err.Description
            Err.Clear
        End If
    Next
    On Error GoTo 0
End Sub

' Sub para verificar InfoAnalogics usando lib antiga
Sub InfoAnalogicaGenericLib(Tela)
    On Error Resume Next
    Dim Obj, TypeNameObj
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If TypeNameObj = "DrawGroup" Then
            InfoAnalogicaGenericLib Obj
        End If
        If InStr(1, TypeNameObj, "InfoAnalogica") > 0 Then
            If Left(TypeNameObj, InStr(1, TypeNameObj, "_")) <> "gx" Then
                AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", "Objeto com a lib de InfoAnalogica antiga, recomenda-se usar a generic"
            End If
        End If
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
            ListaObjetosLib.Add TypeNameObj, Empty
        End If
        If Err.Number <> 0 Then
            AdicionarTxt DadosTxt, LinhaTxt, "InfoAnalogicaGenericLib", Obj, Err.Description
            Err.Clear
        End If
    Next
    On Error GoTo 0
End Sub

' Sub para verificar InfoAnalogics com setpoint
Sub VerificarSPShowInfoAnalogic(Tela)
    On Error Resume Next
    Dim Obj, TypeNameObj, Lib, SPShow
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If TypeNameObj = "DrawGroup" Then
            VerificarSPShowInfoAnalogic Obj
        End If
        If Left(TypeName(Obj), 2) = "xc" Then
            Lib = Left(TypeName(Obj), 2)
        Else
            Lib = ""
        End If
        If Lib = "xc" And InStr(1, TypeNameObj, "InfoAnalogic") > 0 Then
            ' Não verifica para lib xc
        ElseIf InStr(1, TypeNameObj, "InfoAnalogic") > 0 Then
            If Obj.SPTag <> "" Then
                SPShow = Obj.Links.Item("SPShow").Source
                If Obj.SPShow = False And SPShow = "" Then
                    AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "InfoAnalogic possui setpoint, porém SPShow em false ou sem associação"
                End If
            End If
        End If
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
            ListaObjetosLib.Add TypeNameObj, Empty
        End If
        If Err.Number <> 0 Then
            AdicionarTxt DadosTxt, LinhaTxt, "VerificarSPShowInfoAnalogic", Obj, Err.Description
            Err.Clear
        End If
    Next
    On Error GoTo 0
End Sub

' Sub para verificar a cor de fundo da tela
Sub CorBackgroundTela(Tela)
    On Error Resume Next
    If Tela.Links.Item("BackgroundColor").Source = "" Then
        AdicionarExcel DadosExcel, Linha, Tela.PathName, "Aviso", "A cor de fundo da tela deve ser feita através de um link associado com o objeto relacionado a cores do frame dentro do viewer"
    End If
    If Err.Number <> 0 Then
        AdicionarTxt DadosTxt, LinhaTxt, "CorBackgroundTela", Tela, Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Sub para verificar o Caption da tela
Sub VerificarCaptionTela(Tela)
    On Error Resume Next
    If Tela.Caption = "Screen Title" Then
        AdicionarExcel DadosExcel, Linha, Tela.PathName, "Erro", "A propriedade Caption não foi preenchida"
    End If
    If Err.Number <> 0 Then
        AdicionarTxt DadosTxt, LinhaTxt, "VerificarCaptionTela", Tela, Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Sub para verificar se um objeto está sendo usado em libs diferentes
Sub VerificarMesmoObjetoLibsDiferentes()
    Dim ExclusiveValues, obj, celulas
    Set ExclusiveValues = CreateObject("Scripting.Dictionary")
    For Each obj In ListaObjetosLib.Keys
        If InStr(1, obj, "_") > 0 Then
            celulas = Split(obj, "_")
            If Not ExclusiveValues.Exists(celulas(1)) Then
                ExclusiveValues.Add celulas(1), celulas(0)
            Else
                AdicionarExcel DadosExcel, Linha, celulas(1), "Aviso", "O objeto está sendo utilizado através da Lib """ & celulas(0) & """ e da Lib """ & ExclusiveValues.Item(celulas(1)) & """, recomenda-se usar a mesma lib para todos os objetos desse tipo"
            End If
        End If
    Next
End Sub

Sub ObjetoEletricoSemSourceObject(Tela, Obj, ObjetoEletrico) 'Procura sourceObject em objetos elétricos
    On Error Resume Next
	If InStr(1, TypeName(Obj), "Chave", 1) > 0 Then
    	If Obj.NaoSupervisionado = False And Obj.EstadoON = "" Then
        	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", "Chave está supervisionada e com EstadoON em branco"
    	ElseIf Obj.NaoSupervisionado = False And Obj.EstadoOFF = "" Then
        	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", "Chave está supervisionada e com EstadoOFF em branco"
    	End If
	Else
    	If Obj.NaoSupervisionado = False And Obj.SourceObject = "" Then
        	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Chave está supervisionada e sem SourceObject"
    	ElseIf Obj.NaoSupervisionado = True And Obj.SourceObject <> "" Then
        	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", "Chave não está supervisionada e com SourceObject"
    	End If
	End If

    If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
        ListaObjetosLib.Add TypeName(Obj), Empty
    End If
	If Err.Number <> 0 Then
    	AdicionarTxt DadosTxt, LinhaTxt, "ObjetoEletricoSemSourceObject", Obj, Err.Description
    	Err.Clear
	End If

    On Error GoTo 0
End Sub

Sub VerificarBotaoAbreTela(Tela)
    On Error Resume Next
    
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then
            VerificaAlarmBar Obj
        ElseIf InStr(1, TypeNameObj, "BotaoAbreTela", 1) > 0 Then
            On Error Resume Next
            If Obj.Config_TelaOuQuadroPathname = "" Then
                AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Propriedade Config_TelaOuQuadroPathname está vazia, não é possível abrir a tela por esse objeto"
            End If
        End If
        On Error GoTo 0
        
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
            ListaObjetosLib.Add TypeNameObj, Empty
        End If
        
        If Err.Number <> 0 Then
            AdicionarTxt DadosTxt, LinhaTxt, "VerificarBotaoAbreTela", Obj, Err.Description
    		Err.Clear
        End If
    Next
End Sub

Sub ObjetoEletricoDeviceNoteVazio(Tela, Obj, ObjetoEletrico) ' Verifica o DeviceNote vazio em disjuntores que abrem tela de comando
    On Error Resume Next
    
	If InStr(1, TypeName(Obj), ObjetoEletrico, 1) > 0 Then
    	If TypeName(Obj) = "pwa_Trafo3Term" Then
        	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", ObjetoEletrico & " não suporta Notas Operacionais pois não possui a propriedade DeviceNote"
    	ElseIf TypeName(Obj) = "pwa_Gerador" Then
        	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", ObjetoEletrico & " não suporta Notas Operacionais pois não possui a propriedade DeviceNote"
    	ElseIf Obj.NoCommand = False And Obj.DeviceNote = "" Then
        	If Err.Number = 0 Then
            	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", ObjetoEletrico & " com DeviceNote vazio com tela de comando habilitada"
        	End If
    	End If
	End If

    
    If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
        ListaObjetosLib.Add TypeName(Obj), Empty
    End If

        If Err.Number <> 0 Then
            AdicionarTxt DadosTxt, LinhaTxt, "Erro na Sub ObjetoEletricoDeviceNoteVazio/", Obj, Err.Description
    		Err.Clear
        End If
        
    On Error GoTo 0
End Sub

Sub ObjetoEletricoSemTelaDeComandoDeviceNote(Tela, Obj, ObjetoEletrico) ' Verifica o DeviceNote vazio em disjuntores que abrem tela de comando
    On Error Resume Next
    
    If InStr(1, TypeName(Obj), ObjetoEletrico, 1) > 0 Then
        If TypeName(Obj) = "pwa_Trafo3Term" Then
            
        ElseIf TypeName(Obj) = "pwa_Gerador" Then
            
        ElseIf Obj.NoCommand = True And Obj.DeviceNote <> "" Then
            If Err.Number = 0 Then
    			AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", ObjetoEletrico & " sem tela de comando mas com DeviceNote"
            End If
        End If
    End If
    
    If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
        ListaObjetosLib.Add TypeName(Obj), Empty
    End If
    
        If Err.Number <> 0 Then
            AdicionarTxt DadosTxt, LinhaTxt, "Erro na Sub ObjetoEletricoSemTelaDeComandoDeviceNote/", Obj, Err.Description
    		Err.Clear
        End If 
    
    On Error GoTo 0
End Sub

Sub ObjetoMecanicoDeviceNoteVazio(Tela, Obj, ObjetoMecanico) ' Verifica o DeviceNote vazio em disjuntores que abrem tela de comando
    On Error Resume Next
    Set ObjetosIgnorados = CreateObject("Scripting.Dictionary")
    ObjetosIgnorados.Add "uhe_ValveButterfly", Empty
    ObjetosIgnorados.Add "uhe_ValveDistributing", Empty
    ObjetosIgnorados.Add "uhe_Valve3Ways", Empty
    ObjetosIgnorados.Add "uhe_Valve4Ways", Empty
    If InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 Then
        If Not ObjetosIgnorados.Exists(TypeName(Obj)) Then
            If Obj.DeviceNote = "" And Obj.UseNotes = True Then
                If Err.Number = 0 Then
    				AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", ObjetoMecanico & " com DeviceNote vazio mas UseNotes True"
				End If

            End If
        End If
    End If
    
    If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
        ListaObjetosLib.Add TypeName(Obj), Empty
    End If
    
    If Err.Number <> 0 Then
        AdicionarTxt DadosTxt, LinhaTxt, "Erro na Sub ObjetoMecanicoDeviceNoteVazio/", Obj, Err.Description
    	Err.Clear
    End If
    
    On Error GoTo 0
End Sub

Sub ObjetoMecanicoSemTelaDeComandoDeviceNote(Tela, Obj, ObjetoMecanico) ' Verifica o DeviceNote vazio em disjuntores que abrem tela de comando
    On Error Resume Next
    Set ObjetosIgnorados = CreateObject("Scripting.Dictionary")
    ObjetosIgnorados.Add "uhe_ValveButterfly", Empty
    ObjetosIgnorados.Add "uhe_ValveDistributing", Empty
    ObjetosIgnorados.Add "uhe_Valve3Ways", Empty
    ObjetosIgnorados.Add "uhe_Valve4Ways", Empty
    
    If Not ObjetosIgnorados.Exists(TypeName(Obj)) Then
        If InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 Then
            If Obj.DeviceNote <> "" And Obj.UseNotes = False Then
                If Err.Number = 0 Then
    				AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", ObjetoMecanico & " com DeviceNote mas UseNotes False"
				End If
            End If
        End If
    End If
    
    If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
        ListaObjetosLib.Add TypeName(Obj), Empty
    End If

    If Err.Number <> 0 Then
        AdicionarTxt DadosTxt, LinhaTxt, "Erro na Sub ObjetoMecanicoSemTelaDeComandoDeviceNote/", Obj, Err.Description
    	Err.Clear
    End If
    
    On Error GoTo 0
End Sub

Sub ObjetoMecanicoSemSourceObject(Tela, Obj, ObjetoMecanico) ' Verifica se objetos mecanicos supervisionados possuem SourceObject
    On Error Resume Next
    Set ObjetosIgnorados = CreateObject("Scripting.Dictionary")
    ObjetosIgnorados.Add "uhe_ValveDistributing", Empty
    
	If Not ObjetosIgnorados.Exists(TypeName(Obj)) Then
    	If InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 And TypeName(Obj) = "uhe_ValveButterfly" Then
        	If Obj.SourceObject = "" And Obj.NaoSupervisionada = False Then
            	If Err.Number = 0 Then
                	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", ObjetoMecanico & " supervisionada mas sem SourceObject"
            	End If
        	ElseIf Obj.SourceObject <> "" And Obj.NaoSupervisionada = True Then
            	If Err.Number = 0 Then
                	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", ObjetoMecanico & " não supervisionada mas com SourceObject"
    	        End If
    	    End If
    	ElseIf InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 And TypeName(Obj) = "uhe_BrakeAlert" Then
        	If Obj.SourceObject = "" Then
            	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", ObjetoMecanico & " sem SourceObject"
        	End If
    	ElseIf InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 Then
        	If Obj.SourceObject = "" And Obj.Unsupervised = False Then
            	If Err.Number = 0 Then
                	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", ObjetoMecanico & " supervisionada mas sem SourceObject"
            	End If
        	ElseIf Obj.SourceObject <> "" And Obj.Unsupervised = True Then
            	If Err.Number = 0 Then
                	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", ObjetoMecanico & " não supervisionada mas com SourceObject"
            	End If
        	End If
    	End If
	End If

    If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
        ListaObjetosLib.Add TypeName(Obj), Empty
    End If

    If Err.Number <> 0 Then
        AdicionarTxt DadosTxt, LinhaTxt, "Erro na Sub ObjetoMecanicoSemSourceObject/", Obj, Err.Description
    	Err.Clear
    End If
    
    On Error GoTo 0
End Sub

Sub ConferirLinkObjetosMecanicos(Tela, Obj, ObjetoMecanico) 'Verifica os equipamentos mecânicos que são supervisionados estão linkados
    On Error Resume Next
    Set ObjetosIgnorados = CreateObject("Scripting.Dictionary")
    ObjetosIgnorados.Add "uhe_ValveDistributing", Empty
    ObjetosIgnorados.Add "uhe_ValveButterfly", Empty
    ObjetosIgnorados.Add "uhe_Valve3Ways", Empty
    ObjetosIgnorados.Add "uhe_Valve4Ways", Empty
    
	If InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 Then
    	If ObjetoMecanico = "Bomb" Then
        	If (Obj.Unsupervised = False) Then
            	If (Obj.Links.Item("BombOn").Source = "") Then
                	If Err.Number <> 0 Then
                    	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Bomba supervisionada faltando link em BombOn"
                	End If
            	End If
            	If (Obj.Links.Item("BombOff").Source = "") Then
                	If Err.Number <> 0 Then
                    	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Bomba supervisionada faltando link em BombOff"
                	End If
            	End If
        	End If
    	ElseIf ObjetoMecanico = "Valve" And (Not ObjetosIgnorados.Exists(TypeName(Obj))) Then
        	If (Obj.Unsupervised = False) Then
            	If (Obj.Links.Item("Open").Source = "") Then
                	If Err.Number <> 0 Then
                    	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Válvula supervisionada faltando link em Open"
                	End If
            	End If
            	If (Obj.Links.Item("Close").Source = "") Then
                	If Err.Number <> 0 Then
                    	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Válvula supervisionada faltando link em Close"
                	End If
            	End If
        	End If
    	End If
	End If
    
    If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
        ListaObjetosLib.Add TypeName(Obj), Empty
    End If

    If Err.Number <> 0 Then
        AdicionarTxt DadosTxt, LinhaTxt, "Erro na Sub ConferirLinkObjetosMecanicos/", Obj, Err.Description
    	Err.Clear
    End If
    
    On Error GoTo 0
End Sub

Sub ObjetoLibXCNotaOperacional(Tela, Obj, Objeto)
    On Error Resume Next
    
	If InStr(1, TypeName(Obj), "NotaOperacional", 1) > 0 Then
    	If Obj.SourceObject = "" Then
        	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", Objeto & " sem SourceObject"
    	End If
	End If

    
    If Not ListaObjetosLib.Exists("xc_" & Objeto) Then
        ListaObjetosLib.Add "xc_" & Objeto, Empty
    End If
    
    If Err.Number <> 0 Then
        AdicionarTxt DadosTxt, LinhaTxt, "Erro na Sub ObjetoLibXCNotaOperacional/", Obj, Err.Description
    	Err.Clear
    End If
    
    On Error GoTo 0
End Sub

Sub ObjetoMecanicoSupervisaoXC(Tela, Obj, ObjetoMecanico) ' Verifica o DeviceNote vazio em disjuntores que abrem tela de comando
    On Error Resume Next
    Set ObjetosIgnorados = CreateObject("Scripting.Dictionary")
    ObjetosIgnorados.Add "xc_DisjuntoresERAC", Empty
	If Not ObjetosIgnorados.Exists(TypeName(Obj)) Then
    	If InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 Then
        	On Error Resume Next
        	If Obj.NaoSupervisionado = False And Obj.Estado = "" Then
            	If Err.Number <> 0 Then
                	If Obj.Fonte = "" Then
                    	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Propriedade Fonte está vazia"
                	End If
                	If Obj.Command = "" Then
                    	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Propriedade Command está vazia"
                	End If
                	Err.Clear
            	Else
                	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", ObjetoMecanico & " supervisionado sem link de estado"
            	End If
        	End If
        	If Obj.NaoSupervisionado = False And Obj.Cmd = "" Then
            	If Err.Number = 0 Then
                	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", ObjetoMecanico & " supervisionado sem link de comando"
            	End If
        	ElseIf Obj.NaoSupervisionado = True And (Obj.Cmd <> "" Or Obj.Estado <> "") Then
            	If Err.Number = 0 Then
                	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", ObjetoMecanico & " não supervisionado com link de estado ou comando"
            	End If
        	End If
    	End If
	End If

    
    If Not ListaObjetosLib.Exists("xc_" & ObjetoMecanico) Then
        ListaObjetosLib.Add "xc_" & ObjetoMecanico, Empty
    End If

    If Err.Number <> 0 Then
        AdicionarTxt DadosTxt, LinhaTxt, "Erro na Sub ObjetoMecanicoSupervisaoXC/", Obj, Err.Description
    	Err.Clear
    End If

    On Error GoTo 0
End Sub

Sub ObjetoBombaSupervisionadaXC(Tela, Obj, ObjetoMecanico) ' Verifica se o objeto é da biblioteca "xc_" e corresponde ao objeto mecânico
    On Error Resume Next
    	If Left(TypeName(Obj), 3) = "xc_" And InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 Then
    	' Verifica se o objeto não é supervisionado e não possui estado definido
    	If Obj.NaoSupervisionado = False And Obj.Estado = "" Then
        	' Verifica possíveis erros ao acessar as propriedades Fonte e Command
        	If Obj.Fonte = "" Then
            	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Propriedade Fonte está vazia"
        	End If
        	If Obj.Command = "" Then
            	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Propriedade Command está vazia"
        	End If
        	If Err.Number = 0 Then
            	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", ObjetoMecanico & " supervisionado sem link de estado"
        	End If
        	Err.Clear
    	End If
    	
		If Obj.NaoSupervisionado = False And Obj.Cmd = "" Then ' Verifica se o objeto supervisionado não tem comando
    		AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", ObjetoMecanico & " supervisionado sem link de comando"
		ElseIf Obj.NaoSupervisionado = True And (Obj.Cmd <> "" Or Obj.Estado <> "") Then
    		AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", ObjetoMecanico & " não supervisionado com link de estado ou comando"
		End If
		Err.Clear

		ElseIf Left(TypeName(Obj), 2) = "xc" And InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 Then ' Verifica se o objeto não é supervisionado e não possui estado definido
    		If Obj.NaoSupervisionado = False And Obj.Estado = "" Then
        	' Verifica possíveis erros ao acessar as propriedades Cmd e Estado
        	If Obj.Cmd = "" Then
            		AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Propriedade Cmd está vazia"
       		End If
       		If Obj.Estado = "" Then
            	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", "Propriedade Estado está vazia"
        	End If
        	' Caso não tenha erros, indica que o objeto está supervisionado sem link de estado
        	If Err.Number = 0 Then
            	AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", ObjetoMecanico & " supervisionado sem link de estado"
        	End If
        	Err.Clear
    		End If

		If Obj.NaoSupervisionado = False And Obj.Cmd = "" Then ' Verifica se o objeto supervisionado não tem comando
    		AdicionarExcel DadosExcel, Linha, Obj.PathName, "Aviso", ObjetoMecanico & " supervisionado sem link de comando"
		ElseIf Obj.NaoSupervisionado = True And (Obj.Cmd <> "" Or Obj.Estado <> "") Then
    	' Verifica se o objeto não supervisionado contém link de estado ou comando, o que não deveria
    		AdicionarExcel DadosExcel, Linha, Obj.PathName, "Erro", ObjetoMecanico & " não supervisionado com link de estado ou comando"
		End If
		Err.Clear
	End If
   	 ' Verifica se o objeto mecânico existe na lista de objetos já processados
    	If Not ListaObjetosLib.Exists("xc_" & ObjetoMecanico) Then
        ListaObjetosLib.Add "xc_" & ObjetoMecanico, Empty
    	End If
    
    If Err.Number <> 0 Then
        AdicionarTxt DadosTxt, LinhaTxt, "Erro na Sub ObjetoBombaSupervisionadaXC/", Obj, Err.Description
    	Err.Clear
    End If
    On Error GoTo 0
End Sub

'----------------- Funções de Verificação dos Bancos -----------------

Sub VerificarBancoDeDados(DBServerPathName, ObjectPathName, ObjectName)
    On Error Resume Next
    If Not DadosBancoDeDados.Exists(DBServerPathName) Then
        DadosBancoDeDados.Add DBServerPathName, ObjectPathName
    Else
        AdicionarExcel DadosExcel, Linha, ObjectPathName, "Aviso", "O customizador do " & ObjectName & " não possui um banco de dados exclusivo e compartilha o " & DBServerPathName & " com o objeto " & DadosBancoDeDados(DBServerPathName)
    End If

    If Err.Number <> 0 Then
        AdicionarTxt DadosTxt, LinhaTxt, "Erro na Sub VerificarBancoDeDados/", Obj, Err.Description
    	Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub VerificarHist(DBServerPathName, ObjectPathName, ObjectName)
    On Error Resume Next

    ' Verifica se o banco de dados já foi adicionado, se não, adiciona
    If Not DadosBancoDeDados.Exists(DBServerPathName) Then
        DadosBancoDeDados.Add DBServerPathName, ObjectPathName
    Else
        ' Usar a função AdicionarExcel para adicionar um aviso
        AdicionarExcel DadosExcel, Linha, ObjectPathName, "Aviso", "O historiador " & ObjectName & " não possui um banco de dados exclusivo e compartilha o " & DBServerPathName & " com o objeto " & DadosBancoDeDados(DBServerPathName)
    End If

    ' Captura de erros e uso da função AdicionarTxt
    If Err.Number <> 0 Then
        AdicionarTxt DadosTxt, LinhaTxt, "VerificarHist", ObjectPathName, Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub


Sub VerificarHistoriadores(Historiadores)
    For Each Hist In Historiadores
        On Error Resume Next
        Select Case TypeName(Hist)
            Case "DataFolder"
                VerificarHistoriadores Hist
            Case "Hist"
                VerificarHist Hist.DBServer, Hist.PathName, Hist.Name
        End Select
        
    If Err.Number <> 0 Then
        AdicionarTxt DadosTxt, LinhaTxt, "Erro na Sub VerificarHistoriadores/", Obj, Err.Description
    	Err.Clear
    End If
    Next
    On Error GoTo 0
End Sub

'----------------- Funções de Relatório -----------------

' Função para adicionar informações ao Excel
Sub AdicionarExcel(DadosExcel, ByRef Linha, CaminhoObjeto, Tipo, Mensagem)
    If IsObject(DadosExcel) Then
        DadosExcel.Add CStr(Linha), CaminhoObjeto & "/" & Tipo & "/" & Mensagem
        Linha = Linha + 1
    Else
        MsgBox "Erro: O dicionário DadosExcel não foi inicializado."
    End If
End Sub

' Função para adicionar informações ao arquivo TXT
Sub AdicionarTxt(DadosTxt, ByRef LinhaTxt, NomeSub, Obj, DescricaoErro)
    If IsObject(DadosTxt) Then
        DadosTxt.Add CStr(LinhaTxt), "Erro na Sub " & NomeSub & "/" & Obj.PathName & ": " & DescricaoErro
        LinhaTxt = LinhaTxt + 1
    Else
        MsgBox "Erro: O dicionário DadosTxt não foi inicializado."
    End If
End Sub

' Sub para gerar relatório Excel
Sub GerarRelatorioExcel()
    On Error Resume Next
    If DadosExcel.Exists(CStr(2)) Then
        Dim objExcel, objWorkBook, sheet, obj
        Set objExcel = CreateObject("EXCEL.APPLICATION")
        Set objWorkBook = objExcel.Workbooks.Add
        Set sheet = objWorkBook.Sheets(1) ' Ajustado para garantir que a planilha seja acessada corretamente

        ' Definindo o cabeçalho
        sheet.Cells(1, 1).Value = "Objeto"
        sheet.Cells(1, 2).Value = "Tipo"
        sheet.Cells(1, 3).Value = "Problema"

        nomeExcel = CaminhoPrj & "\RelatorioTester_" & nomeExcel & ".xlsx"

        ' Verificando se o dicionário contém dados
        For Each obj In DadosExcel
            Dim celulas
            celulas = Split(DadosExcel.Item(obj), "/")
            sheet.Cells(CInt(obj) + 1, 1).Value = celulas(0)
            sheet.Cells(CInt(obj) + 1, 2).Value = celulas(1)
            sheet.Cells(CInt(obj) + 1, 3).Value = celulas(2)
        Next

        ' Salvando o arquivo Excel
        objWorkBook.SaveAs nomeExcel
        objWorkBook.Close
        objExcel.Quit
        Set objWorkBook = Nothing
        Set objExcel = Nothing

        ' Pergunta se o usuário quer abrir o arquivo gerado
        Dim Resposta
        Resposta = MsgBox("Foram gerados logs de correção, deseja abrir o arquivo?", vbYesNo + vbQuestion, "AutomaTester")
        If Resposta = vbYes Then
            Dim shell
            Set shell = CreateObject("WScript.Shell")
            shell.Run """" & nomeExcel & """"
            Set shell = Nothing
        End If
    Else
        MsgBox "Nenhum dado disponível para gerar o relatório Excel.", vbExclamation
    End If
    On Error GoTo 0

    ' Tratamento de erro
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro na criação do log de erros do projeto, por favor confira o caminho definido para salvar o arquivo"
        Err.Clear
    End If
End Sub

' Sub para gerar relatório TXT
Sub GerarRelatorioTxt()
    On Error Resume Next
    If DadosTxt.Exists(CStr(1)) Then
        Dim aux, aux1, obj, Resposta, shell
        Set aux = CreateObject("Scripting.FileSystemObject")
        nomeTxt = CaminhoPrj & "\Log_" & nomeTxt & ".txt"
        Set aux1 = aux.CreateTextFile(nomeTxt, True)

        ' Verificando se o dicionário contém dados
        For Each obj In DadosTxt
            aux1.WriteLine DadosTxt.Item(obj)
        Next

        aux1.Close

        ' Pergunta se o usuário quer abrir o arquivo gerado
        Resposta = MsgBox("Foram gerados logs de erro de código, deseja abrir o arquivo?", vbYesNo + vbQuestion, "AutomaTester")
        If Resposta = vbYes Then
            Set shell = CreateObject("WScript.Shell")
            shell.Run """" & nomeTxt & """"
            Set shell = Nothing
        End If
    Else
        MsgBox "Nenhum dado disponível para gerar o relatório TXT.", vbExclamation
    End If
    On Error GoTo 0

    ' Tratamento de erro
    If Err.Number <> 0 Then
        MsgBox "Ocorreu um erro na criação do log de erros do script, por favor confira o caminho definido para salvar o arquivo"
        Err.Clear
    End If
End Sub

Sub Fim()
End Sub