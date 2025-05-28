Function GerarScript(TypeName As String, Propriedade1 As String, Propriedade2 As String, TextoAux As String, MetodoAux As String, DefinirBlocoLaco As String, _
    VerificarPropriedadeVazia As String, VerificarPropriedadeHabilitada As String, _
    VerificarPropriedadeCondicional As String, VerificarObjetoDesatualizado As String, VerificarPropriedadeValor As String, _
    VerificarPropriedadeTextoProibido As String, VerificarBancoDeDados As String, ContarObjetosDoTipo As String, _
    VerificarUserFields As String, VerificarCamposUsuariosServidorAlarmes As String, VerificarObjetoInternoIndevido As String, _
    VerificarTipoPai As String, VerificarAssociacaoBase As String, AreaDominio As String, AreaDrivers As String, AreaBancoDados As String, _
    AreaBiblioteca As String, AreaFluxoDados As String, AreaEstruturaPastas As String, _
    AreaTelas As String, RelAviso As String, RelErro As String, RelRevisar As String, Observacao As String) As String

    On Error GoTo erro

    Dim funcao As String
    Dim area As String
    Dim nivel As String
    Dim linhaFinal As String

' === BLOCO ESPECIAL BASEADO NA COLUNA G ===
Select Case DefinirBlocoLaco
    Case "Iniciar área", "Finalizar área"
        GerarScript = "  '=================================================================" & vbCrLf & _
                      "  ' " & Observacao & vbCrLf & _
                      "  '================================================================= "
        Exit Function

    Case "Else", "Finalizar laço(End If)"
        Dim instrucao As String
        If DefinirBlocoLaco = "Finalizar laço(End If)" Then
            instrucao = "End If"
        Else
            instrucao = "Else"
        End If

        If Observacao <> "" Then
            GerarScript = instrucao & " '" & Observacao
        Else
            GerarScript = instrucao
        End If
        Exit Function
End Select

    ' === DEFINIR FUNÇÃO ===
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
    If Observacao <> "" Then
        GerarScript = Observacao
    Else
        GerarScript = "'Objeto não possui verificações"
    End If
    Exit Function

    End If

    ' === DEFINIR ÁREA ===
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

    ' === DEFINIR NÍVEL ===
    If RelAviso <> "" Then
        nivel = "0"
    ElseIf RelErro <> "" Then
        nivel = "1"
    ElseIf RelRevisar <> "" Then
        nivel = "2"
    Else
        nivel = "Nível não definido"
    End If

    ' === MONTAGEM DO SCRIPT ===
    Dim flagFinal As String

    If DefinirBlocoLaco = "Inicar laço(If ... Then)" Or DefinirBlocoLaco = "Condicioonar laço(ElseIf ... Then)" Then
        flagFinal = "True"
    ElseIf DefinirBlocoLaco = "" Then
        flagFinal = "False"
    Else
        flagFinal = "" ' Para casos que não precisam de complemento
    End If

    Select Case funcao
        Case "VerificarPropriedadeVazia"
            linhaFinal = funcao & " Obj, """ & TypeName & """, """ & Propriedade1 & """, """ & area & """, " & nivel
        Case "VerificarPropriedadeHabilitada"
            linhaFinal = funcao & " Obj, """ & TypeName & """, """ & Propriedade1 & """, " & TextoAux & ", """ & area & """, " & nivel
        Case "VerificarPropriedadeCondicional"
            linhaFinal = funcao & " Obj, """ & TypeName & """, """ & Propriedade1 & """, """ & Propriedade2 & """, " & TextoAux & ", """ & area & """, " & nivel
        Case "VerificarObjetoDesatualizado"
            linhaFinal = funcao & " Obj, """ & TypeName & """, " & TextoAux & ", """ & area & """, " & nivel
        Case "VerificarObjetoInternoIndevido", "ContarObjetosDoTipo", "VerificarUserFields", "VerificarCamposUsuariosServidorAlarmes"
            linhaFinal = funcao & " Obj, """ & TypeName & """, """ & area & """, " & nivel
        Case "VerificarPropriedadeValor"
            linhaFinal = funcao & " Obj, """ & TypeName & """, """ & Propriedade1 & """, " & TextoAux & ", " & MetodoAux & ", """ & area & """, " & nivel
        Case "VerificarPropriedadeTextoProibido"
            linhaFinal = funcao & " Obj, """ & TypeName & """, """ & Propriedade1 & """, " & TextoAux & ", """ & area & """, " & nivel
        Case "VerificarBancoDeDados", "VerificarAssociacaoBase"
            linhaFinal = funcao & " Obj, """ & TypeName & """, """ & Propriedade1 & """, """ & area & """, " & nivel
        Case "VerificarTipoPai"
            linhaFinal = funcao & " Obj, """ & TypeName & """, """ & TextoAux & """, " & MetodoAux & ", """ & area & """, " & nivel
        Case Else
            linhaFinal = "'Objeto não possui verificações"
    End Select

    ' Finalizar
    If flagFinal <> "" Then
        linhaFinal = linhaFinal & ", " & flagFinal
    End If

If DefinirBlocoLaco = "Inicar laço(If ... Then)" Then
    GerarScript = "If " & funcao & "(" & Mid(linhaFinal, Len(funcao) + 2) & ") Then"
ElseIf DefinirBlocoLaco = "Condicioonar laço(ElseIf ... Then)" Then
    GerarScript = "Else If " & funcao & "(" & Mid(linhaFinal, Len(funcao) + 2) & ") Then"

    Else
        GerarScript = linhaFinal
    End If

    Exit Function

erro:
    GerarScript = "#ERRO VBA: " & Err.Description
End Function