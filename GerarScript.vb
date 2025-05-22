' Chamado da Função GerarScript
' =GerarScript(B3; C3; D3; E3; F3; G3; H3; I3; J3; K3; L3; M3; N3; O3; P3; Q3; R3; S3; T3; U3; V3; W3; X3; Y3; Z3; AA3; AB3; AC3)

Function GerarScript(TypeName As String, Propriedade1 As String, Propriedade2 As String, TextoAux As String, MetodoAux As String, _
    VerificarPropriedadeVazia As String, VerificarPropriedadeHabilitada As String, _
    VerificarPropriedadeCondicional As String, VerificarObjetoDesatualizado As String, VerificarPropriedadeValor As String, _
    VerificarPropriedadeTextoProibido As String, VerificarBancoDeDados As String, ContarObjetosDoTipo As String, _
    VerificarUserFields As String, VerificarCamposUsuariosServidorAlarmes As String, VerificarObjetoInternoIndevido As String, _
    VerificarTipoPai As String, VerificarAssociacaoBase As String, AreaDominio As String, AreaDrivers As String, AreaBancoDados As String, _
    AreaBiblioteca As String, AreaFluxoDados As String, AreaEstruturaPastas As String, _
    AreaTelas As String, Aviso As String, erro As String, Revisar As String) As String

    On Error GoTo erro

    Dim funcao As String
    Dim area As String
    Dim nivel As String

    ' Definir função com base em prioridade
    If VerificarPropriedadeVazia <> "" Then
        funcao = "VerificarPropriedadeVazia"
    ElseIf VerificarPropriedadeHabilitada <> "" Then
        funcao = "VerificarPropriedadeHabilitada"
    ElseIf VerificarPropriedadeCondicional <> "" Then
        funcao = "VerificarPropriedadeCondicional"
    ElseIf VerificarObjetoDesatualizado <> "" Then
        funcao = "VerificarObjetoDesatualizado"
    ElseIf VerificarPropriedadeValor <> "" Then
        funcao = "VerificarPropriedadeValor"
    ElseIf VerificarPropriedadeTextoProibido <> "" Then
        funcao = "VerificarPropriedadeTextoProibido"
    ElseIf VerificarBancoDeDados <> "" Then
        funcao = "VerificarBancoDeDados"
    ElseIf ContarObjetosDoTipo <> "" Then
        funcao = "ContarObjetosDoTipo"
    ElseIf VerificarUserFields <> "" Then
        funcao = "VerificarUserFields"
    ElseIf VerificarCamposUsuariosServidorAlarmes <> "" Then
        funcao = "VerificarCamposUsuariosServidorAlarmes"
    ElseIf VerificarObjetoInternoIndevido <> "" Then
        funcao = "VerificarObjetoInternoIndevido"
    ElseIf VerificarTipoPai <> "" Then
        funcao = "VerificarTipoPai"
    ElseIf VerificarAssociacaoBase <> "" Then
        funcao = "VerificarAssociacaoBase"
    Else
        GerarScript = "'Objeto não possui verificações"
        Exit Function
    End If

    ' Definir área
    If AreaDominio <> "" Then
        area = "Domínio"
    ElseIf AreaDrivers <> "" Then
        area = "Drivers"
    ElseIf AreaBancoDados <> "" Then
        area = "Banco de dados"
    ElseIf AreaBiblioteca <> "" Then
        area = "Biblioteca"
    ElseIf AreaFluxoDados <> "" Then
        area = "Fluxo de dados"
    ElseIf AreaEstruturaPastas <> "" Then
        area = "Estrutura de pastas"
    ElseIf AreaTelas <> "" Then
        area = "Telas"
    Else
        area = "Área não definida"
    End If

    ' Definir nível
    If Aviso <> "" Then
        nivel = "0"
    ElseIf erro <> "" Then
        nivel = "1"
    ElseIf Revisar <> "" Then
        nivel = "2"
    Else
        nivel = "Nível não definido"
    End If

    ' Montar script
    Select Case funcao
        Case "VerificarPropriedadeVazia"
            GerarScript = funcao & " Obj, """ & TypeName & """, """ & Propriedade1 & """, """ & area & """, " & nivel
        Case "VerificarPropriedadeHabilitada"
            GerarScript = funcao & " Obj, """ & TypeName & """, """ & Propriedade1 & """, " & TextoAux & ", """ & area & """, " & nivel
        Case "VerificarPropriedadeCondicional"
            GerarScript = funcao & " Obj, """ & TypeName & """, """ & Propriedade1 & """, """ & Propriedade2 & """, " & TextoAux & ", """ & area & """, " & nivel
        Case "VerificarObjetoDesatualizado"
            GerarScript = funcao & " Obj, """ & TypeName & """, " & TextoAux & ", """ & area & """, " & nivel
        Case "VerificarObjetoInternoIndevido", "ContarObjetosDoTipo", "VerificarUserFields", "VerificarCamposUsuariosServidorAlarmes"
            GerarScript = funcao & " Obj, """ & TypeName & """, """ & area & """, " & nivel
        Case "VerificarPropriedadeValor"
            GerarScript = funcao & " Obj, """ & TypeName & """, """ & Propriedade1 & """, " & TextoAux & ", " & CStr(MetodoAux) & ", """ & area & """, " & nivel
        Case "VerificarPropriedadeTextoProibido"
            GerarScript = funcao & " Obj, """ & TypeName & """, """ & Propriedade1 & """, " & TextoAux & ", """ & area & """, " & nivel
        Case "VerificarBancoDeDados", "VerificarAssociacaoBase"
            GerarScript = funcao & " Obj, """ & TypeName & """, """ & Propriedade1 & """, """ & area & """, " & nivel
        Case "VerificarTipoPai"
            GerarScript = funcao & " Obj, """ & TypeName & """, """ & TextoAux & """, " & CStr(MetodoAux) & ", """ & area & """, " & nivel
        Case Else
            GerarScript = "'Objeto não possui verificações"
    End Select
    Exit Function

erro:
    GerarScript = "#ERRO VBA: " & Err.Description
End Function