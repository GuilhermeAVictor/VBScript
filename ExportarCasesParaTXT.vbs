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
    Set ws = ThisWorkbook.Sheets("Objetos")

    ' Última linha com dados na coluna AE
    ultimaLinha = ws.Cells(ws.Rows.Count, "AE").End(xlUp).Row

    tipoAnterior = ""

    For linha = 3 To ultimaLinha
        valorScript = ws.Cells(linha, "AE").Value

        If Not IsError(valorScript) Then
            If Len(CStr(valorScript)) > 0 Then
                tipoAtual = ws.Cells(linha, "B").Value

                ' Montar linha do Case considerando "/"
                Dim tipos() As String
                tipos = Split(tipoAtual, "/")
                Dim caseLinha As String
                caseLinha = "Case "

                For i = 0 To UBound(tipos)
                    If i > 0 Then caseLinha = caseLinha & ", "
                    caseLinha = caseLinha & """" & Trim(tipos(i)) & """"
                Next i

                If tipoAtual <> "" And tipoAtual <> tipoAnterior Then
                    If tipoAnterior <> "" Then
                        output = output & vbTab & "'-----------------------------------------------------------------------------" & vbCrLf
                    End If
                    output = output & caseLinha & vbCrLf
                    tipoAnterior = tipoAtual
                End If

                ' Quebra de linhas para script
                linhas = Split(Replace(CStr(valorScript), vbCrLf, vbLf), vbLf)
                For i = 0 To UBound(linhas)
                    If Len(linhas(i)) > 0 Then
                        output = output & vbTab & linhas(i) & vbCrLf
                    End If
                Next i
            End If
        End If
    Next linha

    If tipoAnterior <> "" Then
        output = output & vbTab & "'-----------------------------------------------------------------------------" & vbCrLf
    End If

    caminhoTXT = Environ("TEMP") & "\ScriptsGerados.txt"

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set arquivo = fs.CreateTextFile(caminhoTXT, True)
    arquivo.Write output
    arquivo.Close

    Shell "notepad.exe " & caminhoTXT, vbNormalFocus

    MsgBox "Arquivo gerado com sucesso!", vbInformation
End Sub
