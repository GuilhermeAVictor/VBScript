Sub ExportarCasesParaTXT()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim linha As Long
    Dim tipoAtual As String, tipoAnterior As String
    Dim output As String
    Dim caminhoTXT As String
    Dim fs As Object, arquivo As Object
    Dim valorScript As Variant
    Dim linhas() As String
    Dim i As Integer

    ' Define a planilha com os dados
    Set ws = ThisWorkbook.Sheets("Objetos") ' Altere se necessário

    ' Última linha com dados na coluna B (TypeName)
    ultimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    tipoAnterior = ""

    For linha = 3 To ultimaLinha
        tipoAtual = ws.Cells(linha, "B").Value ' Coluna B = TypeName

        If tipoAtual <> "" Then
            ' Novo bloco Case
            If tipoAtual <> tipoAnterior Then
                If tipoAnterior <> "" Then
                    output = output & vbTab & "'-----------------------------------------------------------------------------" & vbCrLf
                End If

                output = output & "Case """ & tipoAtual & """" & vbCrLf
                tipoAnterior = tipoAtual
            End If

            ' --- Lê célula AE com tratamento ---
            valorScript = ws.Cells(linha, "AE").Value

            If Not IsError(valorScript) Then
    linhas = Split(Replace(CStr(valorScript), vbCrLf, vbLf), vbLf)

    For i = 0 To UBound(linhas)
If Len(linhas(i)) > 0 Then
    output = output & vbTab & linhas(i) & vbCrLf
End If

    Next i
End If
        End If
    Next linha

    ' Fecha último bloco Case
    If tipoAnterior <> "" Then
        output = output & vbTab & "'-----------------------------------------------------------------------------" & vbCrLf
    End If

    ' Caminho do .txt temporário
    caminhoTXT = Environ("TEMP") & "\ScriptsGerados.txt"

    ' Cria e escreve no arquivo
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set arquivo = fs.CreateTextFile(caminhoTXT, True)
    arquivo.Write output
    arquivo.Close

    ' Abre no bloco de notas
    Shell "notepad.exe " & caminhoTXT, vbNormalFocus

    MsgBox "Arquivo gerado com sucesso!", vbInformation
End Sub