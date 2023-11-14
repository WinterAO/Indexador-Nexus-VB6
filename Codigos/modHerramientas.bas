Attribute VB_Name = "modHerramientas"
Option Explicit

Public Sub BuscarGrhLibres(ByVal libres As Integer, ByRef grhMin As Long, ByRef grhMax As Long)

    On Error Resume Next

    Dim i      As Long
    Dim Conta  As Long

    If IsNumeric(libres) = False Then Exit Sub

    grhMin = 0
    grhMax = 0

    For i = 1 To grhCount

        If GrhData(i).NumFrames = 0 Then
            Conta = Conta + 1

            If Conta = libres Then
                grhMin = i - (Conta - 1)
                grhMax = i
                Exit Sub

            End If

        ElseIf Conta > 0 Then
            Conta = 0

        End If

    Next

End Sub
