VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClasseBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim oAutomation     As New CUIAutomation
    Dim janela          As UIAutomationClient.IUIAutomationElement
    Dim Entrada         As UIAutomationClient.IUIAutomationElement
    Dim oPattern        As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern

Public Function localizarAplicativoDoBanco()
    Set janela = WalkEnabledElements(oAutomation, oAutomation.GetRootElement, "Alex Diogo - Bank")
End Function

Public Function clicarNoBotaoIniciar()
    Dim btnIniciar      As UIAutomationClient.IUIAutomationElement
    Dim btnIniciarClick As UIAutomationClient.IUIAutomationInvokePattern
    
    'Localizar o botão Iniciar
    Set btnIniciar = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "button_Iniciar", "AutoID"))
    
    'Clicar no botão Iniciar
    Set btnIniciarClick = btnIniciar.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
    btnIniciarClick.Invoke
End Function

Public Function selecionarDebitoOuCredito(ByVal opcao As String)
    Dim tipo            As UIAutomationClient.IUIAutomationElement
    Dim rBtnOpcao       As UIAutomationClient.IUIAutomationElement
    Dim rBtnOpcaoClick  As UIAutomationClient.IUIAutomationInvokePattern
    
    Set tipo = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "Tipo", "Name"))
    
    If opcao = "Crédito" Then
        Set rBtnOpcao = tipo.FindFirst(TreeScope_Children, PropCondition(oAutomation, "radioButton_Credito", "AutoID"))
    Else
        Set rBtnOpcao = tipo.FindFirst(TreeScope_Children, PropCondition(oAutomation, "radioButton_Debito", "AutoID"))
    End If
    
        Set rBtnOpcaoClick = rBtnOpcao.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
        rBtnOpcaoClick.Invoke
End Function

Public Function localizarOGrupoEntrada()
    Set Entrada = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "Entrada", "Name"))
End Function

Public Function informarADescricao(ByVal descricao As String)
    Dim txtDescricao    As UIAutomationClient.IUIAutomationElement
        
    'Localizar o campo para a Descrição
    Set txtDescricao = Entrada.FindFirst(TreeScope_Children, PropCondition(oAutomation, "textBox_Descricao", "AutoID"))

    'Informar a descrição
    Set oPattern = txtDescricao.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
    oPattern.SetValue (descricao)
    
End Function

Public Function informarOValor(ByVal valor As String)
    Dim txtValor        As UIAutomationClient.IUIAutomationElement
    
    'Localizar o campo para o Valor
    Set txtValor = Entrada.FindFirst(TreeScope_Children, PropCondition(oAutomation, "textBox_Valor", "AutoID"))

    'Informar o Valor
    Set oPattern = txtValor.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
    oPattern.SetValue (valor)
    
End Function

Public Function informarAData(ByVal data As String)
    Dim txtData         As UIAutomationClient.IUIAutomationElement
    
    'Localizar o campo para a Data
    Set txtData = Entrada.FindFirst(TreeScope_Children, PropCondition(oAutomation, "textBox_Data", "AutoID"))

    'Informar a Data
    Set oPattern = txtData.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
    oPattern.SetValue (data)
    
End Function

Public Function clicarNoBotaoGravar()
    Dim BtnGravar       As UIAutomationClient.IUIAutomationElement
    Dim BtnGravarClick  As UIAutomationClient.IUIAutomationInvokePattern
    
    'Localizar o botão Gravar
    Set BtnGravar = Entrada.FindFirst(TreeScope_Children, PropCondition(oAutomation, "button_Gravar", "AutoID"))
    
    'Clicar no botão Gravar
    Set BtnGravarClick = BtnGravar.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
    BtnGravarClick.Invoke
    
End Function

Public Function pegarOValorDoSaldoDoBanco()
    'Localizar o campo Saldo
    Dim lblSaldo As UIAutomationClient.IUIAutomationElement
    Set lblSaldo = janela.FindFirst(TreeScope_Children, PropCondition(oAutomation, "label_SALDO", "AutoID"))
    
    'Pegar o valor atual
    Set oPattern = lblSaldo.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
    
    pegarOValorDoSaldoDoBanco = oPattern.CurrentName
End Function

