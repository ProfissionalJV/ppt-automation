Attribute VB_Name = "AutomationDemo"
Option Explicit

' Projeto demonstrativo – VBA
' Código original desenvolvido em ambiente institucional
' Versão adaptada para fins de portfólio

Sub ExecutarAutomacao()
    Call PrepararAmbiente
    Call ProcessarDados
    Call GerarSaida
End Sub

Private Sub PrepararAmbiente()
    Debug.Print "Ambiente preparado."
End Sub

Private Sub ProcessarDados()
    Dim i As Integer
    For i = 1 To 5
        Debug.Print "Processando registro " & i
    Next i
End Sub

Private Sub GerarSaida()
    Debug.Print "Saída gerada com sucesso."
End Sub
