VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eLocate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ILocateCommandEvents
Private SelElement As Element
Private SelElement_enum As ElementEnumerator
Private cont As Integer
Private Sub ILocateCommandEvents_Accept(ByVal Element As Element, Point As Point3d, ByVal View As View)
'Evento respons?vel por interagir ap?s dois clicks do mouse
'Vazio para que n?o haja a??es durante o evento


End Sub
Private Sub ILocateCommandEvents_Cleanup()
'Evento respons?vel por limpar vari?veis de classe
'Vazio para que n?o haja a??es durante o evento
End Sub
Private Sub ILocateCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
'Vazio para que n?o haja a??es durante o evento
End Sub
Private Sub ILocateCommandEvents_LocateFailed()
'Evento respons?vel por interagir com usu?rio quando a sele??o efetuada por este ? realizada em ?rea que n?o possuem objetos
'Vazio para que n?o haja a??es durante o evento

End Sub
Private Sub ILocateCommandEvents_LocateFilter(ByVal Element As Element, Point As Point3d, Accepted As Boolean)
'Evento respons?vel por interagir com usu?rio ap?s a sele??o de elementos

Dim i As Integer
i = 0
ActiveModelReference.SelectElement Element, True
Set SelElement_enum = ActiveModelReference.GetSelectedElements
'informa??es na aba de log do Microstation
ShowCommand "Selecionar"
ShowPrompt "Click com o bot?o direito duas vezes para aceitar a sele??o"
Accepet = False 'Op??o false permite que a sele??o seja efetuada sem a necessidade de duplo click
cont = cont + 1 'contador para limitar a quantidade de objetos selecionados

If cont = 2 Then
    verifica_elementos SelElement_enum
    Inicializar
End If


End Sub

Private Sub ILocateCommandEvents_LocateReset()
'Evento respons?vel por interagir com usu?rio quando o comando reset (bot?o direito do mouse) ? ativado pelo usu?rio

If cont = 0 Then

    MsgBox "Nenhum elemento selecionado", vbOKOnly
    CommandState.StartDefaultCommand
    Verifica_Selecao = True
    Exit Sub

End If

verifica_elementos SelElement_enum

End Sub

Private Sub ILocateCommandEvents_Start()
'Evetento iniciado automaticamento ao iniciar a classe

MsgBox "Selecione texto e elemento", vbOKOnly
ShowCommand "Selecionar"
ShowPrompt "Selecione os elementos a serem localizados"
cont = 0
End Sub

Private Sub verifica_elementos(SelElement_enum_ As ElementEnumerator)
'Fun??o respons?vel por realizar a verifica??o de quantidade e tipo de elementos selecionados
Dim elementtype As String
Dim SelElement_
Dim i As Integer
i = 0

While SelElement_enum_.MoveNext
    Set SelElement_ = SelElement_enum_.Current
    i = i + 1
    If SelElement_.IsCellElement Then
        elementtype = elementtype & "C?lula" & ", "
    ElseIf SelElement_.IsTextElement Then
       elementtype = elementtype & "Texto" & ", "
    Else
       elementtype = elementtype & "Outro" & ", "
    End If
Wend




MsgBox "Elementos selecionados: " & i & Chr(13) & "Tipo de elemento: " & elementtype, vbOKOnly
CommandState.StartDefaultCommand
Verifica_Selecao = True

End Sub
