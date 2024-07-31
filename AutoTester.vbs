Sub AutoTester_CustomConfig()
'Script que faz uma verificação automática no domínio
Resposta = MsgBox("Tem certeza que deseja iniciar o teste automático do domínio?", 0+4+32, "Iniciar teste de domínio?")

If Resposta = 7 Then
	Exit Sub
End If

    Main()
End Sub

' Configuração para nomes de arquivos
nomeExcel = Replace(Replace(Date() & "_" & Time(), ":", "_"), "/", "_")
nomeTxt = nomeExcel

' Contador das linhas que foram preenchidas
Linha = 2
LinhaTxt = 1

' Criação do dicionário para o Excel
Dim DadosExcel
Set DadosExcel = CreateObject("Scripting.Dictionary")

' Criação dos dicionários para o Txt
Dim DadosTxt
Set DadosTxt = CreateObject("Scripting.Dictionary")

' Criação de dicionário para armazenar os BancosDeDadaos
Dim DadosBancoDeDados
Set DadosBancoDeDados = CreateObject("Scripting.Dictionary")

' Criação do dicionário para verificar o mesmo objeto sendo utilizado em libs diferentes
Set ListaObjetosLib = CreateObject("Scripting.Dictionary")

' Obter o caminho do projeto
Dim CaminhoPrj
If PastaParaSalvarLogs <> ""  Then
	CaminhoPrj = PastaParaSalvarLogs
Else
	CaminhoPrj = CreateObject("WScript.Shell").CurrentDirectory
End If

Sub Main()
    ' Coletar todas as telas disponíveis no projeto
    For Each ScreenObj in Application.ListFiles("Screen")
        VerificarTela ScreenObj
    Next 
	
	ListarXObjectsDominio
	
    ' Gerar relatórios
    If GerarLogErrosScript And Not DebugMode Then
    	VerificarMesmoObjetoLibsDiferentes()
    	GerarRelatorioExcel
    	GerarRelatorioTxt
    ElseIf Not GerarLogErrosScript And Not DebugMode Then
    	VerificarMesmoObjetoLibsDiferentes()
    	GerarRelatorioExcel
    End If
    
    MsgBox "Fim"

End Sub    

Sub VerificarTela(ParentObj)
    Eletricos = Array("Disjuntor", "Seccionadora", "Trafo", "Gerador", "Chave")
    Mecanicos = Array("Bomb", "Valve", "Brake")
    
    InfoAlarmSourceObject ParentObj ' Verifica se o SourceObject01 está preenchido dos InfoAlarmes
    If Not UsandoLibControle Then
    	InfoAlarmGenericLib ParentObj ' Verifica se os InfoAlarmes estão utilizando a lib Generic
    	InfoAnalogicaGenericLib ParentObj ' Verifica se os InfoAnalogica estão utilizando a lib Generic
    End If
    InfoAlarmComValue ParentObj ' Verifica se os SourceObjectXX estão preenchidos incorretamente com .Value
    InfoAnalogicaSemSourceObject ParentObj ' Verifica se os SourceObject da InfoAnalogica estão preenchidos
    VerificarInfoAnalogic ParentObj ' Verifica os objetos gx_InfoAnalogic
    VerificarPwaLineVert ParentObj ' Verificar pwa_LineVert está com a prorpiedade CorOn vazia
    VerificarPwaLineHoriz ParentObj ' Verificar pwa_LineHoriz está com a prorpiedade CorOn vazia
    CorBackgroundTela ParentObj ' Verifica se o background da tela está linkado com a cor do frame
    
    For Each Objeto in Eletricos
    	ClassificarLibEletricos ParentObj, Objeto
    Next
    
    For Each Objeto in Mecanicos
    	ClassificarLibMecanicos ParentObj, Objeto
    Next
	
End Sub

Sub GerarRelatorioExcel()
	On Error Resume Next
    If DadosExcel.Exists(Cstr(2)) Then
        Dim objExcel, objWorkBook
        Set objExcel = CreateObject("EXCEL.APPLICATION")
        Set objWorkBook = objExcel.Workbooks.add
        Set sheet = objWorkBook.Sheets("Planilha1")
        sheet.Cells(1, 1) = "Objeto"
        sheet.Cells(1, 2) = "Tipo"
        sheet.Cells(1, 3) = "Problema"
        nomeExcel = CaminhoPrj & "\RelatorioTester_" & nomeExcel & ".xlsx"
        
        For each obj in DadosExcel
            celulas = Split(DadosExcel.Item(obj), "/")
            sheet.Cells(CInt(obj), 1) = celulas(0)
            sheet.Cells(CInt(obj), 2) = celulas(1)
            sheet.Cells(CInt(obj), 3) = celulas(2)
        Next

        objWorkBook.SaveAs nomeExcel
        objWorkBook.Close
        objExcel.Quit
        Set objWorkBook = Nothing
        Set objExcel = Nothing
        Resposta = (MsgBox("Foram gerados logs de correção, deseja abrir o arquivo?", vbYesNo + vbQuestion + vbDefaultButton1, "AutomaTester"))
        If (Resposta = vbYes) Then
            Set shell = CreateObject("WScript.Shell")
            shell.Run """" & nomeExcel & """"
            Set shell = Nothing
        End If
    End If

    On Error GoTo 0
     If Err.Number <> 0 Then
            MsgBox "Ocorreu um erro na criação do log de erros do projeto, por favor confira o caminho definido para salvar o arquivo"
            Err.Clear
        End If

End Sub

Sub GerarRelatorioTxt()
	On Error Resume Next
    If (DadosTxt.Exists(Cstr(1))) Then
        ' Configuração da criação do log
        Set aux = CreateObject("Scripting.FileSystemObject")
        nomeTxt = CaminhoPrj & "\Log_" & nomeTxt & ".txt"
        Set aux1 = aux.CreateTextFile(nomeTxt, True)
        For each obj in DadosTxt
            aux1.WriteLine DadosTxt.Item(obj)
        Next
        aux1.Close
        Resposta = (MsgBox("Foram gerados logs de erro de código, deseja abrir o arquivo?", vbYesNo + vbQuestion + vbDefaultButton1, "AutomaTester"))
        If (Resposta = vbYes) Then
            Set shell = CreateObject("WScript.Shell")
            shell.Run """" & nomeTxt & """"
            Set shell = Nothing
        End If
    End If
    On Error GoTo 0
    If Err.Number <> 0 Then
            MsgBox "Ocorreu um erro na criação do log de erros do script, por favor confira o caminho definido para salvar o arquivo"
            Err.Clear
        End If
End Sub


Sub VerificarMesmoObjetoLibsDiferentes()
    Set ExclusiveValues = CreateObject("Scripting.Dictionary")
    
    For each obj in ListaObjetosLib.Keys
    	If InStr(1, obj, "_", 1) > 0 Then
    		celulas = Split(obj, "_")
    		If Not ExclusiveValues.Exists(celulas(1)) Then
				ExclusiveValues.Add celulas(1), celulas(0)
			Else
    			DadosExcel.Add CStr(Linha), celulas(1) & "/" & "Aviso" & "/" & "O objeto está sendo utilizado através da Lib " & """" & celulas(0) & """" & " e da Lib " & """" & ExclusiveValues.Item(celulas(1)) & """" & " recomenda-se usar a mesma lib para todos os objetos desse tipo"
        		Linha = Linha + 1
        	End If
        End If
    Next
    
End Sub


Sub ClassificarLibEletricos (Tela, Objeto)
	For Each Obj in Tela
		TypeNameObj = TypeName(Obj)
		If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then'Faz o código procurar dentro de grupos
			ClassificarLibEletricos Obj,Objeto
		End If
	If InStr(1, TypeNameObj, Objeto, 1) > 0 Then
		If Left(TypeName(Obj),2) = "xc" Then
			Lib = Left(TypeName(Obj),2)	
		Else
			Lib = Left(TypeName(Obj),InStr(1, TypeName(Obj), "_", 1)-1)
		End If
		Select Case True
		
			Case (Lib = "xc") And (Objeto = "Disjuntor")
				ObjetoMecanicoSupervisaoXC Tela, Obj, Objeto
			Case (Lib = "xc") And (Objeto = "Gerador")
				'Verificar o que deve ser linkado com gerador
			Case (Lib="xc") And (Objeto = "Trafo")
				'Nada a colocar aqui, a menos que um link em energizado seja obrigatório
			Case (Lib = "xc") And (Objeto = "NotaOperacional")
				ObjetoLibXCNotaOperacional Tela, Obj, Objeto
			Case (Lib = "xc") And (Objeto = "Chave")
				ObjetoEletricoSemSourceObject Tela, Obj, Objeto
            Case (Lib = "pwa") And (Objeto = "Disjuntor")
                ObjetoEletricoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos eletricos que abrem tela de comando
      		    ObjetoEletricoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos eletricos sem tela mas com DeviceNote
      		Case (Lib = "pwa") And (Objeto = "Seccionadora")
                ObjetoEletricoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos eletricos que abrem tela de comando
      		    ObjetoEletricoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos eletricos sem tela mas com DeviceNote
      		Case (Lib = "pwa") And (Objeto = "Trafo")
				ObjetoEletricoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos eletricos que abrem tela de comando
      			ObjetoEletricoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos eletricos sem tela mas com DeviceNote
      		 Case (Lib = "pwa") And (Objeto = "Gerador")
      		   ObjetoEletricoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos eletricos que abrem tela de comando
      		   ObjetoEletricoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos eletricos sem tela mas com DeviceNote
Case Else
               MsgBox "Lib: " & Lib & " " & Objeto & " Não cadastrada como elétrico, consulte equipe de testers"
         End select
	End If
	Next
End Sub

Sub ClassificarLibMecanicos (Tela, Objeto)
	For Each Obj in Tela
		TypeNameObj = TypeName(Obj)
		If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then'Faz o código procurar dentro de grupos
			ClassificarLibMecanicos Obj,Objeto
		End If
	If InStr(1, TypeNameObj, Objeto, 1) > 0 Then
		If Left(TypeName(Obj),2) = "xc" Then
			Lib = Left(TypeName(Obj),2)	
		Else
			Lib = Left(TypeName(Obj),InStr(1, TypeName(Obj), "_", 1)-1)
		End If
		Select Case True
			Case (Lib = "xc") And (Objeto = "Bomb")
				ObjetoBombaSupervisionadaXC Tela, Obj, Objeto
            Case (Lib = "uhe") And (Objeto = "Valve")
				ObjetoMecanicoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos mecanicos que abrem tela de comando
        		ObjetoMecanicoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos mecanicos sem tela mas com DeviceNote
        		ObjetoMecanicoSemSourceObject Tela, Obj, Objeto ' Verifica se tem objetos mecanicos sem SourceObject preenchido e supervisionado
        		ConferirLinkObjetosMecanicos Tela, Obj, Objeto ' Verifica os equipamentos mecânicos que são supervisionados estão linkados
        	Case (Lib = "uhe") And (Objeto = "Bomb")
				ObjetoMecanicoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos mecanicos que abrem tela de comando
        		ObjetoMecanicoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos mecanicos sem tela mas com DeviceNote
        		ObjetoMecanicoSemSourceObject Tela, Obj, Objeto ' Verifica se tem objetos mecanicos sem SourceObject preenchido e supervisionado
        		ConferirLinkObjetosMecanicos Tela, Obj, Objeto ' Verifica os equipamentos mecânicos que são supervisionados estão linkados
        	Case (Lib = "uhe") And (Objeto = "Brake")
				ObjetoMecanicoDeviceNoteVazio Tela, Obj, Objeto ' Verifica o DeviceNote vazio em objetos mecanicos que abrem tela de comando
        		ObjetoMecanicoSemTelaDeComandoDeviceNote Tela, Obj, Objeto ' Verifica se tem objetos mecanicos sem tela mas com DeviceNote
        		ObjetoMecanicoSemSourceObject Tela, Obj, Objeto ' Verifica se tem objetos mecanicos sem SourceObject preenchido e supervisionado
        		ConferirLinkObjetosMecanicos Tela, Obj, Objeto ' Verifica os equipamentos mecânicos que são supervisionados estão linkados
        	Case Else
               MsgBox "Lib: " & Lib & " " & Objeto & " Não cadastrada como mecânico, consulte equipe de testers"
         End select
	End If
	Next
End Sub


Sub VerificarPwaLineVert(Tela)
    On Error Resume Next
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then
            VerificarPwaLineVert Obj
        ElseIf InStr(1, TypeNameObj, "LineVert", 1) > 0 Then
            On Error Resume Next  
            If (Obj.Links.Item("CorOn").Source = "") Then
                DadosExcel.Add CStr(Linha), Obj.PathName & "/" & "Aviso" & "/" & "Propriedade CorOn está vazia"
                Linha = Linha + 1
            End If
        End If
        On Error GoTo 0
        
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
			ListaObjetosLib.Add TypeNameObj, Empty
		End If	
        
        If Err.Number <> 0 Then
            DadosTxt.Add CStr(LinhaTxt), "Erro na Sub VerificarPwaLineVert/" & Obj.PathName & ": " & Err.Description
            LinhaTxt = LinhaTxt + 1
            Err.Clear
        End If
    Next
End Sub

Sub VerificarPwaLineHoriz(Tela)
    On Error Resume Next
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then
            VerificarPwaLineHoriz Obj
        ElseIf InStr(1, TypeNameObj, "LineHoriz", 1) > 0 Then
            On Error Resume Next  
            If (Obj.Links.Item("CorOn").Source = "") Then
                DadosExcel.Add CStr(Linha), Obj.PathName & "/" & "Aviso" & "/" & "Propriedade CorOn está vazia"
                Linha = Linha + 1
            End If
        End If
        On Error GoTo 0
        
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
			ListaObjetosLib.Add TypeNameObj, Empty
		End If	
        
        If Err.Number <> 0 Then
            DadosTxt.Add CStr(LinhaTxt), "Erro na Sub VerificarPwaLineHoriz/" & Obj.PathName & ": " & Err.Description
            LinhaTxt = LinhaTxt + 1
            Err.Clear
        End If
    Next
End Sub

Sub VerificarInfoAnalogic(Tela)
    On Error Resume Next
    For Each Obj In Tela
        TypeNameObj = TypeName(Obj)
        If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then
            VerificarInfoAnalogic Obj 
        End If
        
        If Left(TypeName(Obj),2) = "xc" Then
			Lib = Left(TypeName(Obj),2)	
		Else
			Lib = ""
		End If
        
        If Lib = "xc" And InStr(1, TypeNameObj, "InfoAnalogic", 1) > 0 Then
        	'Não verifica nada pois não possui propriedade de setpoint, só está aqui para evitar erros com a lib xc
        ElseIf InStr(1, TypeNameObj, "InfoAnalogic", 1) > 0 Then ' Verificar se o objeto é do tipo InfoAnalogic
        	On Error Resume Next
            If Obj.SPTag <> "" Then ' Verificar se a propriedade SPTag não está vazia
                On Error Resume Next
                SPShow = Obj.Links.Item("SPShow").Source
                If (Obj.SPShow = False and SPShow = "") Then' Verificar se SPShow é False ou se a associação está vazia
                    DadosExcel.Add CStr(Linha), Obj.PathName & "/" & "Erro" & "/" & "InfoAnalogic possui setpoint, porem SPShow em false ou sem associação"
                    Linha = Linha + 1
                End If
            End If
            On Error GoTo 0
        End If
        
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
			ListaObjetosLib.Add TypeNameObj, Empty
		End If	
        
        If Err.Number <> 0 Then
            DadosTxt.Add CStr(LinhaTxt), "Erro na Sub VerificarInfoAnalogic/" & Obj.PathName & ": " & Err.Description
            LinhaTxt = LinhaTxt + 1
            Err.Clear
        End If
    Next
    On Error GoTo 0
End Sub

Sub InfoAlarmSourceObject(Tela)'Verifica se os InfoAlarmes estão com o SourceObject01 preenchido
'On Error Resume Next
    For Each Obj in Tela
        TypeNameObj = TypeName(Obj)
        If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then'Faz o código procurar dentro de grupos
            InfoAlarmSourceObject(Obj)
        End If
        
        If Left(TypeName(Obj),2) = "xc" Then
			Lib = Left(TypeName(Obj),2)	
		Else
			Lib = ""
		End If
        
        If Lib = "xc" And InStr(1, TypeNameObj, "InfoAlarme", 1) > 0 Then
        	If (Obj.AreaAlarme = "") Then
                DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "Sem AreaAlarme"
                Linha = Linha + 1
            End If
        ElseIf InStr(1, TypeNameObj, "InfoAlarme", 1) > 0 Then
            If (Obj.SourceObject01 = "") Then
                DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "Propriedade SourceObjecto01 em branco"
                Linha = Linha + 1
            End If
        End If
        
        If Not ListaObjetosLib.Exists(TypeNameObj) Then
			ListaObjetosLib.Add TypeNameObj, Empty
		End If	
        
        If Err.Number <> 0 Then
            DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub InfoAlarmSourceObject/" & Obj.PathName & ": " & Err.Description
            LinhaTxt = LinhaTxt + 1
            Err.Clear
        End If
    Next
On Error Goto 0
End Sub

Sub InfoAlarmGenericLib(Tela)'Verifica se os InfoAlarmes estão sendo utilizados com a lib nova
On Error Resume Next
	For Each Obj in Tela
		TypeNameObj = TypeName(Obj)
		If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then'Faz o código procurar dentro de grupos
			InfoAlarmGenericLib(Obj)
		End If
		If InStr(1, TypeNameObj, "InfoAlarme", 1) > 0 Then
			If (Left(TypeNameObj,InStr(1, TypeNameObj, "_", 1)) <> "gx") Then
				DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Aviso" & "/" & "Objeto com a lib de InfoAlarm antiga, recomenda-se usar a generic"
				Linha = Linha + 1
			End If
		End If
		
		If Not ListaObjetosLib.Exists(TypeNameObj) Then
			ListaObjetosLib.Add TypeNameObj, Empty
		End If	
		
		If Err.Number <> 0 Then
			DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub InfoAlarmGenericLib/" & Obj.PathName & ": " & Err.Description
  			LinhaTxt = LinhaTxt + 1
  			Err.Clear
		End If
	Next
On Error Goto 0
End Sub

Sub InfoAlarmComValue(Tela)'Verifica se os InfoAlarmes estão com .value no SourceObject
On Error Resume Next
	For Each Obj in Tela
		TypeNameObj = TypeName(Obj)
		If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then'Faz o código procurar dentro de grupos
			InfoAlarmComValue(Obj)
		End If
		
		If Left(TypeName(Obj),2) = "xc" Then
			Lib = Left(TypeName(Obj),2)	
		Else
			Lib = ""
		End If
        
        If Lib = "xc" And InStr(1, TypeNameObj, "InfoAlarme", 1) > 0 Then
			'Não existe o que verificar, essa linha é só para não gerar erros com essa lib xc.
		ElseIf InStr(1, TypeNameObj, "InfoAlarme", 1) > 0 Then
			For i=1 to Cint(Right(TypenameObj,2))
				If i < 10 Then
					i = "0" & Cstr(i)
				End If
				Execute "SourceObjectxx = Obj.SourceObject" & Cstr(i)
				If (InStr(1,SourceObjectxx,".Value",1) > 0) Then
					DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "Objeto com .Value no SourceObject"
					Linha = Linha + 1
				End If
			Next
		End If
		
		If Not ListaObjetosLib.Exists(TypeNameObj) Then
			ListaObjetosLib.Add TypeNameObj, Empty
		End If	
		
		If Err.Number <> 0 Then
			DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub InfoAlarmComValue/" & Obj.PathName & ": " & Err.Description
  			LinhaTxt = LinhaTxt + 1
			Err.Clear
		End If
	Next
On Error Goto 0
End Sub

Sub ObjetoEletricoSemSourceObject(Tela,Obj, ObjetoEletrico) ' Verifica o DeviceNote vazio em disjuntores que abrem tela de comando
On Error Resume Next

	If InStr(1, TypeName(Obj), "Chave", 1) > 0 Then
		If Obj.NaoSupervisionado = False And Obj.EstadoON = "" Then
			DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Aviso" & "/" & "Chave está supervisionada e com EstadoON em branco"
			Linha = Linha + 1
		ElseIf Obj.NaoSupervisionado = False And Obj.EstadoOFF = "" Then
			DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Aviso" & "/" & "Chave está supervisionada e com EstadoOFF em branco"
			Linha = Linha + 1
		End If
	'Por enquanto está sem configuração para procurar sourceoobject em objetos elétricos, apenas essa lib (xc) pede	
	End If
	
	If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
		ListaObjetosLib.Add TypeName(Obj), Empty
	End If
 	
 	If Err.Number <> 0 Then
		DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub ObjetoEletricoSemSourceObject/" & Obj.PathName & ": " & Err.Description
  		LinhaTxt = LinhaTxt + 1
		Err.Clear
	End If

	On Error Goto 0
End Sub


Sub ObjetoEletricoDeviceNoteVazio(Tela, Obj, ObjetoEletrico) ' Verifica o DeviceNote vazio em disjuntores que abrem tela de comando
On Error Resume Next

	If InStr(1, TypeName(Obj), ObjetoEletrico, 1) > 0 Then
		If TypeName(Obj) = "pwa_Trafo3Term" Then
			DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Aviso" & "/" & ObjetoEletrico & " não suporta Notas Operacionais pois não possui a propriedade DeviceNote"
			Linha = Linha + 1
		ElseIf TypeName(Obj) = "pwa_Gerador" Then
			DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Aviso" & "/" & ObjetoEletrico & " não suporta Notas Operacionais pois não possui a propriedade DeviceNote"
			Linha = Linha + 1
		ElseIf Obj.NoCommand = False And Obj.DeviceNote = "" Then
			If Err.Number = 0 Then
				DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & ObjetoEletrico & " com DeviceNote vazio com tela de comando habilitada"
				Linha = Linha + 1
			End If
		End If
	End If
	
	If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
		ListaObjetosLib.Add TypeName(Obj), Empty
	End If		
 	
 	If Err.Number <> 0 Then
		DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub ObjetoEletricoDeviceNoteVazio/" & Obj.PathName & ": " & Err.Description
  		LinhaTxt = LinhaTxt + 1
		Err.Clear
	End If

	On Error Goto 0
End Sub

Sub ObjetoEletricoSemTelaDeComandoDeviceNote(Tela,Obj, ObjetoEletrico) ' Verifica o DeviceNote vazio em disjuntores que abrem tela de comando
On Error Resume Next

	If InStr(1, TypeName(Obj), ObjetoEletrico, 1) > 0 Then
		If TypeName(Obj) = "pwa_Trafo3Term" Then
		
		ElseIf TypeName(Obj) = "pwa_Gerador" Then				
		
		ElseIf Obj.NoCommand = True And Obj.DeviceNote <> "" Then
			If Err.Number = 0 Then
				DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Aviso" & "/" & ObjetoEletrico & " sem tela de comando mas com DeviceNote"
				Linha = Linha + 1
			End If
		End If
	End If
	
	If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
		ListaObjetosLib.Add TypeName(Obj), Empty
	End If	
	
	If Err.Number <> 0 Then
		DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub ObjetoEletricoSemTelaDeComandoDeviceNote/" & Obj.PathName & ": " & Err.Description
  		LinhaTxt = LinhaTxt + 1
		Err.Clear
	End If

	On Error Goto 0
End Sub

Sub ObjetoMecanicoDeviceNoteVazio(Tela, Obj, ObjetoMecanico) ' Verifica o DeviceNote vazio em disjuntores que abrem tela de comando
On Error Resume Next
	If InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 Then
		If TypeName(Obj) <> "uhe_ValveButterfly" And TypeName(Obj) <> "uhe_ValveDistributing" And TypeName(Obj) <> "uhe_Valve3Ways" And TypeName(Obj) <> "uhe_Valve4Ways" Then
			If Obj.DeviceNote = "" And Obj.UseNotes = True Then	
				If Err.Number = 0 Then
					DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & ObjetoMecanico & " com DeviceNote vazio mas UseNotes True"
					Linha = Linha + 1
				End If
			End If
		End If
	End If
		
	If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
		ListaObjetosLib.Add TypeName(Obj), Empty
	End If	
		
	If Err.Number <> 0 Then
		DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub ObjetoMecanicoDeviceNoteVazio/" & Obj.PathName & ": " & Err.Description
  		LinhaTxt = LinhaTxt + 1
 		Err.Clear
	End If
	
On Error Goto 0
End Sub

Sub ObjetoMecanicoSemTelaDeComandoDeviceNote(Tela, Obj, ObjetoMecanico) ' Verifica o DeviceNote vazio em disjuntores que abrem tela de comando
On Error Resume Next

	If TypeName(Obj) <> "uhe_ValveButterfly" And TypeName(Obj) <> "uhe_ValveDistributing" And TypeName(Obj) <> "uhe_Valve3Ways" And TypeName(Obj) <> "uhe_Valve4Ways" Then
		If InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 Then
			If Obj.DeviceNote <> "" And Obj.UseNotes = False Then
				If Err.Number = 0 Then
					DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Aviso" & "/" & ObjetoMecanico & " com DeviceNote mas UseNotes False"
					Linha = Linha + 1
				End If
			End If
		End If
	End If
	
	If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
		ListaObjetosLib.Add TypeName(Obj), Empty
	End If	
	
	If Err.Number <> 0 Then
		DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub ObjetoMecanicoSemTelaDeComandoDeviceNote/" & Obj.PathName & ": " & Err.Description
  		LinhaTxt = LinhaTxt + 1
  		Err.Clear
	End If

On Error Goto 0
End Sub

Sub ObjetoMecanicoSemSourceObject(Tela, Obj, ObjetoMecanico) ' Verifica se objetos mecanicos supervisionados possuem SourceObject
On Error Resume Next

	If TypeName(Obj) <> "uhe_ValveDistributing" Then
		If InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 And TypeName(Obj) = "uhe_ValveButterfly" Then
			If Obj.SourceObject = "" And Obj.NaoSupervisionada = False Then
				If Err.Number = 0 Then
					DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & ObjetoMecanico & " supervisionada mas sem SourceObject"
					Linha = Linha + 1
				End If
			End If
		ElseIf InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 And TypeName(Obj) = "uhe_BrakeAlert"Then
			If Obj.SourceObject = "" Then
				DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & ObjetoMecanico & " sem SourceObject"
				Linha = Linha + 1
			End If
		ElseIf InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 Then
			If Obj.SourceObject = "" And Obj.Unsupervised = False Then
				If Err.Number = 0 Then
					DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & ObjetoMecanico & " supervisionada mas sem SourceObject"
					Linha = Linha + 1
				End If
			End If	
		End If
	End If
	
	If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
		ListaObjetosLib.Add TypeName(Obj), Empty
	End If	
	
	If Err.Number <> 0 Then
		DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub ObjetoMecanicoSemSourceObject/" & Obj.PathName & ": " & Err.Description
  		LinhaTxt = LinhaTxt + 1
  		Err.Clear
	End If

On Error Goto 0
End Sub

Sub InfoAnalogicaSemSourceObject(Tela) ' Verifica o DeviceNote vazio em disjuntores que abrem tela de comando
On Error Resume Next
	For Each Obj in Tela
		TypeNameObj = TypeName(Obj)
		If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then'Faz o código procurar dentro de grupos
			InfoAnalogicaSemSourceObject Obj
		End If
		
		If Left(TypeName(Obj),2) = "xc" Then
			Lib = Left(TypeName(Obj),2)	
		Else
			Lib = ""
		End If
        
        If Lib = "xc" And InStr(1, TypeNameObj, "InfoAnalogica", 1) > 0 Then
			If Obj.ValueTag = "" Then	
				If Err.Number = 0 Then
					DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "InfoAnalogica sem ValueTag"
					Linha = Linha + 1
				End If
			End If
		ElseIf InStr(1, TypeNameObj, "InfoAnalogica", 1) > 0 Then
			If Obj.SourceObject = "" Then	
				If Err.Number = 0 Then
					DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "InfoAnalogica sem SourceObject"
					Linha = Linha + 1
				End If
			End If
		End If
		
		If Not ListaObjetosLib.Exists(TypeNameObj) Then
			ListaObjetosLib.Add TypeNameObj, Empty
		End If	
		
	If Err.Number <> 0 Then
		DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub InfoAnalogicaSemSourceObject/" & Obj.PathName & ": " & Err.Description
  		LinhaTxt = LinhaTxt + 1
 		Err.Clear
	End If
	Next
On Error Goto 0
End Sub

Sub InfoAnalogicaGenericLib(Tela)'Verifica se os InfoAnalogicas estão sendo utilizados com a lib nova

On Error Resume Next
	For Each Obj in Tela
		TypeNameObj = TypeName(Obj)
		If StrComp(TypeNameObj, "DrawGroup", 1) = 0 Then'Faz o código procurar dentro de grupos
			InfoAnalogicaGenericLib Obj
		End If
		If InStr(1, TypeNameObj, "InfoAnalogica", 1) > 0 Then
			If (Left(TypeNameObj,InStr(1, TypeNameObj, "_", 1)) <> "gx") Then
				DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Aviso" & "/" & "Objeto com a lib de InfoAnalogica antiga, recomenda-se usar a generic"
				Linha = Linha + 1
			End If
		End If
		
		If Not ListaObjetosLib.Exists(TypeNameObj) Then
			ListaObjetosLib.Add TypeNameObj, Empty
		End If	
		
		If Err.Number <> 0 Then
			DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub InfoAnalogicaGenericLib/" & Obj.PathName & ": " & Err.Description
  			LinhaTxt = LinhaTxt + 1
  			Err.Clear
		End If
	Next
On Error Goto 0
End Sub

Sub ConferirLinkObjetosMecanicos(Tela, Obj, ObjetoMecanico)'Verifica os equipamentos mecânicos que são supervisionados estão linkados
On Error Resume Next

	If InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 Then
			If ObjetoMecanico = "Bomb" Then
				If (Obj.Unsupervised = False) Then
					If (Obj.Links.Item("BombOn").Source = "") Then
						If Err.Number <> 0 Then
							DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "Bomba supervisionada faltando link em BombOn"
							Linha = Linha + 1
						Else
							DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "Bomba supervisionada faltando link em BombOn"
							Linha = Linha + 1
						End If
					End If
					If (Obj.Links.Item("BombOff").Source = "") Then
						If Err.Number <> 0 Then
							DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "Bomba supervisionada faltando link em BombOff"
							Linha = Linha + 1
						Else
							DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "Bomba supervisionada faltando link em BombOff"
							Linha = Linha + 1
						End If
					End If
				End If
			ElseIf ObjetoMecanico = "Valve" And TypeName(Obj) <> "uhe_ValveDistributing" And TypeName(Obj) <> "uhe_ValveButterfly" And TypeName(Obj) <> "uhe_Valve3Ways" Then
				If (Obj.Unsupervised = False) Then
					If (Obj.Links.Item("Open").Source = "") Then
						If Err.Number <> 0 Then
							DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "Válvula supervisionada faltando link em Open"
							Linha = Linha + 1
						Else
							DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "Válvula supervisionada faltando link em Open"
							Linha = Linha + 1
						End If
					End If
					If (Obj.Links.Item("Close").Source = "") Then
						If Err.Number <> 0 Then
							DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "Válvula supervisionada faltando link em Close"
							Linha = Linha + 1
						Else
							DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & "Válvula supervisionada faltando link em Close"
							Linha = Linha + 1
						End If
					End If
				End If
			End If
		
	End If
	
	If Not ListaObjetosLib.Exists(TypeName(Obj)) Then
		ListaObjetosLib.Add TypeName(Obj), Empty
	End If	
	
	If Err.Number <> 0 Then
		DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub ConferirLinkObjetosMecanicos/" & Obj.PathName & ": " & Err.Description & "/Se existir o mesmo objeto no excel com um erro de link este erro é porque o link estava em branco e este erro pode ser ignorado"
  		LinhaTxt = LinhaTxt + 1
  		Err.Clear
	End If

On Error Goto 0
End Sub

Sub CorBackgroundTela(Tela)'Verifica se as telas estão linkadas com o background do frame
On Error Resume Next
	
	If Tela.Links.Item("BackgroundColor").Source = "" Then
		DadosExcel.Add Cstr(Linha), Tela.PathName & "/" & "Aviso" & "/" & "A cor de fundo da tela deve ser feita através de um link associado com o objeto relacionado a cores do frame dentro do viewer"
		Linha = Linha + 1
	End If
	
	If Err.Number <> 0 Then
		DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub CorBackgroundTela/" & Tela.PathName & ": " & Err.Description & "/Se existir o mesmo objeto no excel com um erro de link este erro é porque o link estava em branco e este erro pode ser ignorado"
  		LinhaTxt = LinhaTxt + 1
  		Err.Clear
	End If
	
On Error Goto 0
End Sub

Sub ObjetoLibXCNotaOperacional(Tela, Obj, Objeto)
On Error Resume Next

	If InStr(1, TypeName(Obj), "NotaOperacional", 1) > 0 Then
		If Obj.SourceObject = "" Then
			DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Aviso" & "/" & Objeto & " sem SourceObject"
			Linha = Linha + 1
		End If
	End If
	
	If Not ListaObjetosLib.Exists("xc_" & Objeto) Then
		ListaObjetosLib.Add "xc_" & Objeto, Empty
	End If
	
	If Err.Number <> 0 Then
		DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub ObjetoLibXCNotaOperacional/" & Obj.PathName & ": " & Err.Description
  		LinhaTxt = LinhaTxt + 1
  		Err.Clear
	End If

On Error Goto 0
End Sub

Sub ObjetoMecanicoSupervisaoXC(Tela, Obj, ObjetoMecanico) ' Verifica o DeviceNote vazio em disjuntores que abrem tela de comando
On Error Resume Next

	If InStr(1, TypeName(Obj), ObjetoMecanico, 1) > 0 Then
		If Obj.NaoSupervisionado = False And Obj.Estado = "" Then	
			DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & ObjetoMecanico & " supervisionado sem link de estado"
			Linha = Linha + 1
		End If
		If Obj.NaoSupervisionado = False And Obj.Cmd = "" Then	
			DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Aviso" & "/" & ObjetoMecanico & " supervisionado sem link de comando"
			Linha = Linha + 1
		ElseIf Obj.NaoSupervisionado = True And (Obj.Cmd <> "" Or Obj.Estado <> "") Then	
			DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & ObjetoMecanico & " não supervisionado com link de estado ou comando"
			Linha = Linha + 1
		End If
	End If
	If Not ListaObjetosLib.Exists("xc_" & ObjetoMecanico) Then
		ListaObjetosLib.Add "xc_" & ObjetoMecanico, Empty
	End If	
	If Err.Number <> 0 Then
		DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub ObjetoMecanicoSupervisaoXC/" & Obj.PathName & ": " & Err.Description
  		LinhaTxt = LinhaTxt + 1
 		Err.Clear
	End If

On Error Goto 0
End Sub

Sub ObjetoBombaSupervisionadaXC(Tela, Obj, ObjetoMecanico)
On Error Resume Next
	If Obj.NaoSupervisionada = False And Obj.Estado = "" Then	
		DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & ObjetoMecanico & " supervisionado sem link de estado"
		Linha = Linha + 1
	End If
	If Obj.NaoSupervisionada = False And Obj.Cmd = "" Then	
		DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Aviso" & "/" & ObjetoMecanico & " supervisionado sem link de comando"
		Linha = Linha + 1
	ElseIf Obj.NaoSupervisionada = True And (Obj.Cmd <> "" Or Obj.Estado <> "") Then	
		DadosExcel.Add Cstr(Linha), Obj.PathName & "/" & "Erro" & "/" & ObjetoMecanico & " não supervisionado com link de estado ou comando"
		Linha = Linha + 1
	End If

	If Not ListaObjetosLib.Exists("xc_" & ObjetoMecanico) Then
		ListaObjetosLib.Add "xc_" & ObjetoMecanico, Empty
	End If
	
	If Err.Number <> 0 Then
		DadosTxt.Add Cstr(LinhaTxt), "Erro na Sub ObjetoBombaSupervisionadaXC/" & Obj.PathName & ": " & Err.Description
  		LinhaTxt = LinhaTxt + 1
 		Err.Clear
	End If

On Error Goto 0
End Sub

Sub ListarXObjectsDominio()
    Set DataServer = Application.ListFiles("DataServer")
    FiltrarXObjectsDominio DataServer

    ' Verifica os historiadores
    Set Historiadores = Application.ListFiles("Hist")
    VerificarHistoriadores Historiadores
     
End Sub

Sub FiltrarXObjectsDominio(DataServer)
    For Each Object in DataServer
        Select Case TypeName(Object)
        Case "DataServer"
            FiltrarXObjectsDominio Object
        Case "DataFolder"
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

Sub VerificarBancoDeDados(DBServerPathName, ObjectPathName, ObjectName)
    If Not DadosBancoDeDados.Exists(DBServerPathName) Then
        DadosBancoDeDados.Add DBServerPathName, ObjectPathName
    Else
        DadosExcel.Add CStr(Linha), ObjectPathName & "/" & "Aviso" & "/" & "O customizador do " & ObjectName & " não possui um banco de dados exclusivo e compartilha o " & DBServerPathName & " com o objeto " & DadosBancoDeDados(DBServerPathName)
        Linha = Linha + 1
    End If
End Sub

Sub VerificarHist(DBServerPathName, ObjectPathName, ObjectName)
    If Not DadosBancoDeDados.Exists(DBServerPathName) Then
        DadosBancoDeDados.Add DBServerPathName, ObjectPathName
    Else
        DadosExcel.Add CStr(Linha), ObjectPathName & "/" & "Aviso" & "/" & "O historiador " & ObjectName & " não possui um banco de dados exclusivo e compartilha o " & DBServerPathName & " com o objeto " & DadosBancoDeDados(DBServerPathName)
        Linha = Linha + 1
    End If
End Sub

Sub VerificarHistoriadores(Historiadores)
    For Each Hist In Historiadores
        Select Case TypeName(Hist)
        Case "DataFolder"
            VerificarHistoriadores Hist
        Case "Hist"
            VerificarHist Hist.DBServer, Hist.PathName, Hist.Name
        End Select
    Next
End Sub

Sub Fim()
End Sub