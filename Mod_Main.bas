Attribute VB_Name = "Mod_Main"
Option Explicit

Sub Main()
    Dim i           As Integer
    Dim banco       As New ClasseBanco
    Dim saldoSite   As String
    Dim saldoBanco  As String
    Dim tabela()
    
    Let tabela = pegarTabelaDoSite
    Let saldoSite = pegarOValorDoSaldoDoSite
    
    Call abrirAplicativo
    
    Application.Wait (Now + TimeValue("0:00:02"))
    
    With banco
        .localizarAplicativoDoBanco
        .clicarNoBotaoIniciar
    End With
    
    
    For i = 2 To UBound(tabela)
        With banco
            .selecionarDebitoOuCredito (tabela(i, 3))
            .localizarOGrupoEntrada
            .informarADescricao (tabela(i, 2))
            .informarOValor (tabela(i, 4))
            .informarAData (tabela(i, 5))
            .clicarNoBotaoGravar
        End With
    Next i
    
    saldoBanco = banco.pegarOValorDoSaldoDoBanco
   
   MsgBox "Saldo do Site: " & saldoSite & vbNewLine & _
          "Saldo do Banco: " & saldoBanco, vbInformation, "Saldos"
End Sub
