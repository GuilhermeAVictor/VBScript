Sub ExportarCasesParaTXT()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim linha As Long
    Dim tipoAtual As String, tipoAnterior As String
    Dim output As String
    Dim caminhoTXT As String
    Dim fs As Object, arquivo As Object

    ' Define a planilha com os dados (ajuste conforme necessário)
    Set ws = ThisWorkbook.Sheets("Objetos") ' Altere para o nome correto da aba

    ' Última linha com dados na coluna B (TypeName)
    ultimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    tipoAnterior = ""

    For linha = 3 To ultimaLinha
        tipoAtual = ws.Cells(linha, "B").Value ' Coluna B = TypeName

        If tipoAtual <> "" Then
            ' Se mudou o TypeName, começa novo bloco Case
            If tipoAtual <> tipoAnterior Then
                If tipoAnterior <> "" Then
                    output = output & vbTab & "'-----------------------------------------------------------------------------" & vbCrLf
                End If

                output = output & "Case """ & tipoAtual & """" & vbCrLf
                tipoAnterior = tipoAtual
            End If

            ' Adiciona o script (em coluna AD - ajuste se necessário)
            output = output & vbTab & ws.Cells(linha, "AD").Value & vbCrLf
        End If
    Next linha

    ' Fecha último bloco
    If tipoAnterior <> "" Then
        output = output & vbTab & "'-----------------------------------------------------------------------------" & vbCrLf
    End If

    ' Caminho temporário do arquivo .txt
    caminhoTXT = Environ("TEMP") & "\ScriptsGerados.txt"

    ' Criar arquivo e salvar conteúdo
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set arquivo = fs.CreateTextFile(caminhoTXT, True)
    arquivo.Write output
    arquivo.Close

    ' Abrir no Bloco de Notas
    Shell "notepad.exe " & caminhoTXT, vbNormalFocus

    MsgBox "Arquivo gerado com sucesso!", vbInformation
End Sub