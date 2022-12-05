Attribute VB_Name = "Menu"
Public Sub Inicializar()

'Tratamento de Erro
On Error GoTo Error

Dim excelApp As Excel.Application 'Criação de aplicação do excel
Set excelApp = CreateObject("Excel.Application") 'instancia de um objeto da aplicação excel
Dim text As String
Dim Elemento0 As Element 'Elemento
Dim Elemento1 As Element 'Elemento
Dim Elemento2 As Element 'Elemento
Dim elemento_enum As ElementEnumerator
Dim Check As Integer

Inicio:

'------------------------------------------------------------------------------------------------------------------------------------------
'                                                          MANIPULAÇÃO DO OBJETO SELECIONADO NO MICROSTATION
'------------------------------------------------------------------------------------------------------------------------------------------
 
Check = 0
  
Set elemento_enum = ActiveModelReference.GetSelectedElements
    
    While elemento_enum.MoveNext
         
        Set Elemento0 = elemento_enum.Current
        
            If Elemento0.IsTextElement Then
                Set Elemento1 = elemento_enum.Current
               If Check = 0 Then
                  Check = 1
               Else
                  Check = 3
               End If
            ElseIf Elemento0.IsCellElement Then
                 Set Elemento2 = elemento_enum.Current
                If Check = 0 Then
                  Check = 2
               Else
                  Check = 3
              End If
            End If
    Wend
    

Select Case Check
Case 0
    MsgBox "Nenhum elemento selecioando" & Chr(13) & "Reinicie a aplicação e selecione os elementos", vbCritical
    Exit Sub
Case 1
    Profundidade Elemento1.AsTextElement
    MsgBox "Somente o elemento texto foi plotado", vbCritical
    Exit Sub
Case 2
    Copia_elementos Elemento2.AsCellElement
    MsgBox "Somente o elemento célula foi plotado", vbCritical
    Exit Sub
Case 3
    Copia_elementos Elemento2
    Profundidade Elemento1
    MsgBox "Ambos elementos foram plotados sucesso", vbOKOnly
    Exit Sub
End Select

Error:
CadInputQueue.SendReset
MsgBox "Erros inexperado contate o desenvolvedor", vbCritical, "Error"
Exit Sub

Sair:
MsgBox "Deixando aplicação", vbOKOnly
End Sub




