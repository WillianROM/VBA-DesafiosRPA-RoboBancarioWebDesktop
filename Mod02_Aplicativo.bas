Attribute VB_Name = "Mod02_Aplicativo"
Option Explicit

Sub abrirAplicativo()
    Dim objShell    As Object
    Dim caminho     As String

    Set objShell = CreateObject("Shell.Application")
 
    caminho = ThisWorkbook.Path & "\Banco.exe"
    
    objShell.Open (caminho)
    
End Sub

