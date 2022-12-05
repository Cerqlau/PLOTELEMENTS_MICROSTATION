Attribute VB_Name = "Start"
Private handler As eLocate 'necess�rio definir o manipulador de evento desta forma para evitar sua desintegra��o ap�s execu��o da sub
Global Path As String

Sub main()
'------------------------------------------------------------------------------------------------------------------------------------------
'                                                           INFORMA��ES
'------------------------------------------------------------------------------------------------------------------------------------------

'C�digo para plotar e manipular elementos no microstation a partir de uma base de dados no excel
'criado por Lauro Cerqueira

Dim excelApp As Excel.Application 'Cria��o de aplica��o do excel
Set excelApp = CreateObject("Excel.Application") 'instancia de um objeto da aplica��o excel
Dim text As String
Dim Elemento0 As Element 'Elemento
Dim Elemento1 As Element 'Elemento
Dim Elemento2 As Element 'Elemento
Dim elemento_enum As ElementEnumerator
Dim Check As Integer

text = "                                INFORMA��O" _
        & "" & Chr(13) & "--------------------------------------------------------------------------" _
        & "" & Chr(13) & "Esta aplica��o foi criada com intuido de facilitar a plotagem de elementos no Microstation." _
        & "" & Chr(13) & "Para sua exceu��o deve-se possuir uma base de dados em planilha no excel que contenha a coluna A" _
        & " preenchida com as coordenadas Leste, B com as coordenadas Norte e C com profundidade." _
        & "" & Chr(13) & "--------------------------------------------------------------------------" _
        & "" & Chr(13) & "                          INSTRU��ES PARA USO" _
        & "" & Chr(13) & "--------------------------------------------------------------------------" _
        & "" & Chr(13) & "1 - Previamente Plote a primeiro elemento e texto de profundidade;" _
        & "" & Chr(13) & "2 - Ajuste cor, escala e camada;" _
        & "" & Chr(13) & "3 - Carregue a planilha base de dados;" _
        & "" & Chr(13) & "4 - Selecione o elemento e o texto;" _
        & "" & Chr(13) & "6 - Be Happy !!!   =]" _
        & "" & Chr(13) & "--------------------------------------------------------------------------" _
        & "" & Chr(13) & "Created By Lauro Cerqueira"
MsgBox text, vbQuestion 'Texto de informa��es iniciais


'------------------------------------------------------------------------------------------------------------------------------------------
'                                                           MANIPULA��O DA PLANILA DE DADOS
'------------------------------------------------------------------------------------------------------------------------------------------


'Solicita escolha de base de dados
Caminho:

'Encontrar o arquivo
Set Objeto = CreateObject("scripting.filesystemobject")
Path = excelApp.GetOpenFilename(filefilter:="xlsx File, *.*")

'Verifica escolha do usu�rio e se o arquivo � v�lido
If Path = "" Or Path = "False" Then
    GoTo Sair
Else
    If Right(Path, 4) = "xlsx" Then
       GoTo Inicio
    Else
        MsgBox "Tipo de arquivo incorreto", vbCritical
        GoTo Caminho
    End If
End If

Sair:
MsgBox "Deixando aplica��o", vbOKOnly
Exit Sub

Inicio:

'Chamada da clasee criada elocate e que utiliza as primitiva IlocateCommandEvents para intera��o de sele��o com o usu�rio
Set handler = New eLocate
    CommandState.StartLocate handler
End Sub
