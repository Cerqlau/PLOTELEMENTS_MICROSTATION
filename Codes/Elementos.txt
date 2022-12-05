Attribute VB_Name = "Elementos"
Public Sub Copia_elementos(Elementos As Element)
    
    
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
    Dim Qt_X As Integer, Qt_Y As Integer, Qt_Z As Integer 'quantidade de pontos X,Y,Z
    Set excelApp = CreateObject("Excel.Application") 'instancia de um objeto da aplicação excel
    
    'Declaração Variáveis Globais para manipulação do elemento  no Microstation
    Dim elemento As Element 'Elemento
    Dim Copia As Element
    Dim elemento_enum As ElementEnumerator
    Dim Origem As Point3d 'elemento coordenadas X,Y,Z Origem
    Dim Destino As Point3d 'elemento coordenadas X,Y,Z Destino
    Dim mytext As TextElement 'Elemento Texto
    Dim textpt As Point3d 'Coordenada Elemento Texto
    Dim rotMatrix As Matrix3d 'Matriz 3d
    
'------------------------------------------------------------------------------------------------------------------------------------------
'                                                           MANIPULAÇÃO DA PLANILA DE DADOS
'------------------------------------------------------------------------------------------------------------------------------------------
    
    'Se atribuir True, a janela do Excel aparecerá
    'excelApp.Visible = True
    excelApp.Visible = False
    
    excelApp.Workbooks.Open (Path)
    'Debug.Print excelApp.ActiveWorkbook.Name
    Set mySheet = excelApp.ActiveWorkbook.Worksheets("Sheet1") 'instancia da planilha específica
    mySheet.Activate
    
    'Verificação da quantidade de pontos por linha
    Qt_X = mySheet.Cells(1, 1).End(xlDown).Row 'Comando que simula Ctrl+shift+seta para baixo no excel
    Qt_Y = mySheet.Cells(1, 2).End(xlDown).Row
    Qt_Z = mySheet.Cells(1, 3).End(xlDown).Row
    
    'Redimensionamento de arrays
    ReDim Array_X(Qt_X - 1)
    ReDim Array_Y(Qt_Y - 1)
    ReDim Array_Z(Qt_Z - 1)
    
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
    
    'MANIPULAÇÃO DO OBJETO SELECIONADO
        
    Set elemento = Elementos.AsCellElement
        
        If elemento.IsCellElement Then
                
            'Capturando coordenada da seleção
            
            Origem.X = elemento.AsCellElement.Origin.X
            Origem.Y = elemento.AsCellElement.Origin.Y
            Origem.Z = elemento.AsCellElement.Origin.Z
            
            ActiveModelReference.UnselectAllElements 'limpa a seleção
            ActiveModelReference.SelectElement elemento, True 'seleciona somente o objeto principal
            CadInputQueue.SendCommand "copy" 'Copiando o elemento para coodenada destino
            CadInputQueue.SendDataPoint Origem 'set de coordenada origem
            
            For i = 1 To (Qt_X - 1)
                'Inserindo coordenada destino e realizando a cópia
                Destino.X = Array_X(i)
                Destino.Y = Array_Y(i)
                Destino.Z = 0
                CadInputQueue.SendDataPoint Destino
            Next i
            CadInputQueue.SendReset
        End If
    
Saida:
'Limpando os comandos de seleção
CadInputQueue.SendReset
Exit Sub
    
Error:
MsgBox "Erros inexperado contate o desenvolvedor", vbCritical, "Error"
Exit Sub
    
End Sub
 

