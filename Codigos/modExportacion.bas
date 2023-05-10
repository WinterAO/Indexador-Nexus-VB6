Attribute VB_Name = "modExportacion"
Option Explicit

Public Sub DIR_INDEXADOR()

    If LenB(Dir(DirExport, vbDirectory)) = 0 Then
        Call MkDir(DirExport)

    End If

    If LenB(Dir(DirIndex, vbDirectory)) = 0 Then
        Call MkDir(DirIndex)

    End If

    If LenB(Dir(DirCliente & "\Graficos\0.png", vbArchive)) <> 0 Then
        Call Kill(DirCliente & "\Graficos\0.png")
        DoEvents

    End If

End Sub

Public Sub ExportarCabezas()
    On Error Resume Next

    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim K As Integer

    Dim Datos As String

    frmMain.GRHt.Text = "Exportando..."
    DoEvents
    Call DIR_INDEXADOR

    Datos = "[INIT]" & vbCrLf & "NumHeads=" & Numheads & vbCrLf & vbCrLf

    For i = 1 To Numheads

        If HeadData(i).Head(1).GrhIndex > 0 Then
            Datos = Datos & "[HEAD" & (i) & "]" & vbCrLf

            For n = 1 To 4
                Datos = Datos & "Head" & (n) & "=" & HeadData(i).Head(n).GrhIndex & vbCrLf & IIf(n = 1, Chr(9) & " ' arriba", "") & IIf(n = 2, Chr(9) & " ' derecha", "") & IIf(n = 3, Chr(9) & " ' abajo", "") & IIf(n = 4, Chr(9) & " ' izq", "") & vbCrLf
            Next
            Datos = Datos & vbCrLf

        End If

    Next

    frmMain.GRHt.Text = "Guardando...Cabezas.ini"
    DoEvents

    Open (DirExport & "\Cabezas.ini") For Binary Access Write As #1
    Put #1, , Datos
    Close #1

    frmMain.GRHt.Text = "Exportado...Cabezas.ini"
End Sub

Public Sub ExportarCascos()
    On Error Resume Next

    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim K As Integer

    Dim Datos As String

    frmMain.GRHt.Text = "Exportando..."
    DoEvents
    Call DIR_INDEXADOR

    Datos = "[INIT]" & vbCrLf & "NumHelmets=" & NumCascos & vbCrLf & vbCrLf

    For i = 1 To NumCascos

        If CascoAnimData(i).Head(1).GrhIndex > 0 Then
            Datos = Datos & "[HELMET" & (i) & "]" & vbCrLf

            For n = 1 To 4
                Datos = Datos & "Helmet" & n & "=" & CascoAnimData(i).Head(n).GrhIndex & vbCrLf & IIf(n = 1, Chr(9) & " ' arriba", "") & IIf(n = 2, Chr(9) & " ' derecha", "") & IIf(n = 3, Chr(9) & " ' abajo", "") & IIf(n = 4, Chr(9) & " ' izq", "") & vbCrLf
            Next
            Datos = Datos & vbCrLf

        End If

    Next

    frmMain.GRHt.Text = "Guardando...Cascos.ini"
    DoEvents

    Open (DirExport & "\Cascos.ini") For Binary Access Write As #1
    Put #1, , Datos
    Close #1

    DoEvents
    frmMain.GRHt.Text = "Exportado...Cascos.ini"
End Sub

Public Sub ExportarCuerpos()
    On Error Resume Next

    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim K As Integer

    Dim Datos As String

    frmMain.GRHt.Text = "Exportando..."
    DoEvents
    Call DIR_INDEXADOR

    Datos = "[INIT]" & vbCrLf & "NumBodies=" & NumCuerpos & vbCrLf & vbCrLf

    For i = 1 To NumCuerpos
        Datos = Datos & "[BODY" & (i) & "]" & vbCrLf

        For n = 1 To 4
            Datos = Datos & "WALK" & (n) & "=" & BodyData(i).Walk(n).GrhIndex & vbCrLf & IIf(n = 1, Chr(9) & " ' arriba", "") & IIf(n = 2, Chr(9) & " ' derecha", "") & IIf(n = 3, Chr(9) & " ' abajo", "") & IIf(n = 4, Chr(9) & " ' izq", "") & vbCrLf
        Next
        Datos = Datos & "HeadOffsetX=" & BodyData(i).HeadOffset.x & vbCrLf & "HeadOffsetY=" & BodyData(i).HeadOffset.y & vbCrLf & vbCrLf
    Next

    frmMain.GRHt.Text = "Guardando...Cuerpos.ini"
    DoEvents

    Open (DirExport & "\Cuerpos.ini") For Binary Access Write As #1
    Put #1, , Datos
    Close #1

    frmMain.GRHt.Text = "Exportado...Cuerpos.ini"
End Sub

Public Sub ExportarFxs()
    On Error Resume Next

    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim K As Integer

    Dim Datos As String

    frmMain.GRHt.Text = "Exportando..."
    DoEvents
    Call DIR_INDEXADOR

    Datos = "[INIT]" & vbCrLf & "NumFxs=" & NumFxs & vbCrLf & vbCrLf

    For i = 1 To NumFxs

        If FxData(i).Animacion > 0 Then
            Datos = Datos & "[FX" & (i) & "]" & vbCrLf
            Datos = Datos & "Animacion=" & FxData(i).Animacion & vbCrLf & "OffsetX=" & FxData(i).OffsetX & vbCrLf & "OffsetY=" & FxData(i).OffsetY & vbCrLf & vbCrLf

        End If

    Next

    frmMain.GRHt.Text = "Guardando...FXs.ini"
    DoEvents

    Open (DirExport & "\FXs.ini") For Binary Access Write As #1
    Put #1, , Datos
    Close #1

    DoEvents

    frmMain.GRHt.Text = "Exportado...FXs.ini"
End Sub

Public Sub ExportarGraficos()
    Rem On Error Resume Next
    Dim i As Long
    Dim j As Integer
    Dim n As Integer
    Dim K As Integer

    Dim Datos$

    Ocupado = True
    Play = False
    frmMain.GRHt.Text = "Exportando..."
    Call DIR_INDEXADOR
    DoEvents

    If LenB(Dir(DirExport & "\Graficos.ini", vbArchive)) = 1 Then
        Call Kill(DirExport & "\Graficos.ini")

    End If

    n = FreeFile
    Open DirExport & "\Graficos.ini" For Binary Access Write As n
    Put n, , "[INIT]" & vbCrLf & "NumGrh=" & grhCount & vbCrLf & vbCrLf
    K = 0

    Put n, , "[Graphics]" & vbCrLf

    For i = 1 To grhCount
        K = K + 1

        If K > 100 Then
            frmMain.GRHt.Text = "Exportando..." & i & " de " & grhCount
            DoEvents
            K = 0

        End If

        If GrhData(i).NumFrames > 0 Then
            Datos$ = ""

            If GrhData(i).NumFrames = 1 Then
                Datos$ = "1-" & CStr(GrhData(i).FileNum) & "-" & CStr(GrhData(i).sX) & "-" & CStr(GrhData(i).sY) & "-" & CStr(GrhData(i).pixelWidth) & "-" & CStr(GrhData(i).pixelHeight)
            
            Else
                Datos$ = CStr(GrhData(i).NumFrames)

                For j = 1 To GrhData(i).NumFrames
                    Datos$ = Datos$ & "-" & CStr(GrhData(i).Frames(j))
                Next

                Dim speed As Double

                speed = GrhData(i).NumFrames / 0.018
                Datos$ = Datos$ & "-" & CStr(speed)

            End If

            If Len(Datos$) > 0 Then
                Put n, , "Grh" & CStr(i) & "=" & Datos$ & vbCrLf

            End If

        End If

    Next
    Close #n
    frmMain.GRHt.Text = "Exportado...Graficos.ini"
    Ocupado = False
End Sub
