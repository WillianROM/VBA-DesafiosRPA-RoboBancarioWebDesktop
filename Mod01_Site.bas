Attribute VB_Name = "Mod01_Site"
Option Explicit
Dim driver As New ChromeDriver

Function pegarTabelaDoSite()

    With driver
        .Start
        .Window.Maximize
        .Get "https://desafiosrpa.com.br/extrato.html"
    End With
        
        Dim data(): data = driver.FindElementById("example").AsTable.data
        
        pegarTabelaDoSite = data
        
End Function

Function pegarOValorDoSaldoDoSite()
    pegarOValorDoSaldoDoSite = driver.FindElementByCss("div.table-responsive b").Text
End Function
