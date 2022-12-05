Attribute VB_Name = "Texto"
Public Sub Profundidade(Elementos As Element)

'------------------------------------------------------------------------------------------------------------------------------------------
'                                                           INFORMAÇÕES
'------------------------------------------------------------------------------------------------------------------------------------------

'Código para plotar e manipular elementos no microstation a partir de uma base de dados no excel
'criado por Lauro Cerqueira

'------------------------------------------------------------------------------------------------------------------------------------------
'                                                           DECLARAÇÃO DE VARIÁVEIS
'------------------------------------------------------------------------------------------------------------------------------------------
    
'Tratamento de Erro
On Error GoTo Error

'Declaração de variáveis globais para leitura do database na planilha Excel
Dim excelApp As Excel.Application 'Criação de aplicação do excel
Dim mySheet As Worksheet 'Criação de um objeto planilha
Dim X As Long 'coordenadas East
Dim Y As Long 'coordenadas North
Dim Array_X() As Long, Array_Y() As Long, Array_Z() As Long 'Arrays para armazenamento dos pontos X,Y,Z
Dim Array_NX() As Long, Array_NY() As Long, Array_NZ() As Long 'Arrays para armazenamento dos pontos da profundiade X,Y,Z
Dim Qt_X As Integer, Qt_Y As Integer, Qt_Z As Integer 'quantidade de pontos X,Y,Z
Set excelApp = CreateObject("Excel.Application") 'instancia de um objeto da aplicação excel

'Declaração Variáveis Globais para manipulação do elemento  no Microstation
Dim elemento As Element 'Elemento
Dim Copia As TextElement 'Copiado
Dim elemento_enum As ElementEnumerator
Dim Origem As Point3d 'elemento coordenadas X,Y,Z Origem
Dim Destino As Point3d 'elemento coordenadas X,Y,Z Destino
      

'------------------------------------------------------------------------------------------------------------------------------------------
'                                                           MANIPULAÇÃO DA PLANILA DE DADOS
'------------------------------------------------------------------------------------------------------------------------------------------
Inicio:

'Atribuir True, a janela do Excel aparecerá
'excelApp.Visible = True
excelApp.Visible = False

excelApp.Workbooks.Open (Path) 'utiliza a aplicação excel para abrir o arquivo no caminho desejado
Set mySheet = excelApp.ActiveWorkbook.Worksheets("Sheet1") 'instancia da planilha específica
mySheet.Activate 'ativa a planilha específica para que as ações sejam executadas nela

'Verificação da quantidade de pontos por linha
Qt_X = mySheet.Cells(1, 1).End(xlDown).Row 'Comando que simula Ctrl+shift+seta para baixo no excel
Qt_Y = mySheet.Cells(1, 2).End(xlDown).Row
Qt_Z = mySheet.Cells(1, 3).End(xlDown).Row

'Redimensionamento de arrays
ReDim Array_X(Qt_X - 1)
ReDim Array_Y(Qt_Y - 1)
ReDim Array_Z(Qt_Z - 1)
ReDim Array_NX(Qt_X - 1)
ReDim Array_NY(Qt_Y - 1)
ReDim Array_NZ(Qt_Z - 1)

'Captura de cordenadas X
For i = 0 To (Qt_X - 1)
    Array_X(i) = mySheet.Cells(i + 1, 1).Value
Next i

'Captura de cordenadas Y
For i = 0 To (Qt_Y - 1)
    Array_Y(i) = mySheet.Cells(i + 1, 2).Value
Next i

'Captura de profundiade Z
For i = 0 To (Qt_Z - 1)
    Array_Z(i) = mySheet.Cells(i + 1, 3).Value
Next i

'Termina o Excel
excelApp.Quit
    
  

'------------------------------------------------------------------------------------------------------------------------------------------
'                                                          MANIPULAÇÃO DO OBJETO SELECIONADO NO MICROSTATION
'------------------------------------------------------------------------------------------------------------------------------------------
             
    Set elemento = Elementos.AsTextElement
        
        If elemento.IsTextElement Then
            
            'Capturando coordenada da seleção
            Origem.X = Round(Elementos.AsTextElement.AsTextElement.Origin.X, 0)
            Origem.Y = Round(Elementos.AsTextElement.AsTextElement.Origin.Y, 0)
            Origem.Z = Round(Elementos.AsTextElement.AsTextElement.Origin.Z, 0)
            
            'caclula o offset da do texto para o objeto batimetria
            Dif_X = Array_X(0) - Origem.X
            Dif_Y = Array_Y(0) - Origem.Y
            
            For i = 1 To (Qt_X - 1)
            
                'Cálculo coordenada destino texto
                Destino.Y = Array_Y(i) - Dif_Y
                Destino.X = Array_X(i) - Dif_X
                Destino.Z = 0
                
                'array de me para as coordenadas dos textos
                Array_NY(i) = Destino.Y
                Array_NX(i) = Destino.X
                Array_NZ(i) = Destino.Z
                
               'cópia do elemento
               Set Copia = ActiveModelReference.CopyElement(elemento)
               
               'edição do elemento texto
               With Copia.AsTextElement
               .Origin = Destino
               .text = Str(Array_Z(i))
               .Rewrite
               End With
            Next i
     End If
    
'Limpando os comandos de seleção
CadInputQueue.SendReset
Exit Sub


Error:
CadInputQueue.SendReset
MsgBox "Erros inexperado contate o desenvolvedor", vbCritical, "Error"
Exit Sub

Sair:

End Sub

