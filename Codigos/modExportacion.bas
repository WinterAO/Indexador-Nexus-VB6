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

    Call AddtoRichTextBox(frmMain.RichConsola, "Exportando...", 0, 162, 232)
    DoEvents
    Call DIR_INDEXADOR

    Datos = "[INIT]" & vbCrLf & "NumHeads=" & NumHeads & vbCrLf & vbCrLf

    For i = 1 To NumHeads

        If HeadData(i).Head(1).GrhIndex > 0 Then
            Datos = Datos & "[HEAD" & (i) & "]" & vbCrLf

            For n = 1 To 4
                Datos = Datos & "Head" & (n) & "=" & HeadData(i).Head(n).GrhIndex & vbCrLf & IIf(n = 1, Chr(9) & " ' arriba", "") & IIf(n = 2, Chr(9) & " ' derecha", "") & IIf(n = 3, Chr(9) & " ' abajo", "") & IIf(n = 4, Chr(9) & " ' izq", "") & vbCrLf
            Next
            Datos = Datos & vbCrLf

        End If

    Next

    Call AddtoRichTextBox(frmMain.RichConsola, "Guardando Head.ini", 0, 162, 232)
    DoEvents

    Open (DirExport & "\Head.ini") For Binary Access Write As #1
    Put #1, , Datos
    Close #1

    Call AddtoRichTextBox(frmMain.RichConsola, "Head.ini Exportado", 0, 255, 0)
End Sub

Public Sub ExportarCascos()
    On Error Resume Next

    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim K As Integer

    Dim Datos As String

    Call AddtoRichTextBox(frmMain.RichConsola, "Exportando...", 0, 162, 232)
    DoEvents
    Call DIR_INDEXADOR

    Datos = "[INIT]" & vbCrLf & "NumCascos=" & NumCascos & vbCrLf & vbCrLf

    For i = 1 To NumCascos

        If CascoAnimData(i).Head(1).GrhIndex > 0 Then
            Datos = Datos & "[HELMET" & (i) & "]" & vbCrLf

            For n = 1 To 4
                Datos = Datos & "Helmet" & n & "=" & CascoAnimData(i).Head(n).GrhIndex & vbCrLf & IIf(n = 1, Chr(9) & " ' arriba", "") & IIf(n = 2, Chr(9) & " ' derecha", "") & IIf(n = 3, Chr(9) & " ' abajo", "") & IIf(n = 4, Chr(9) & " ' izq", "") & vbCrLf
            Next
            Datos = Datos & vbCrLf

        End If

    Next

    Call AddtoRichTextBox(frmMain.RichConsola, "Guardando Helmet.ini", 0, 162, 232)
    DoEvents

    Open (DirExport & "\Helmet.ini") For Binary Access Write As #1
    Put #1, , Datos
    Close #1

    DoEvents
    Call AddtoRichTextBox(frmMain.RichConsola, "Helmet.ini Exportado", 0, 255, 0)
End Sub

Public Sub ExportarCuerpos()
    On Error Resume Next

    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim K As Integer

    Dim Datos As String

    Call AddtoRichTextBox(frmMain.RichConsola, "Exportando...", 0, 162, 232)
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

    Call AddtoRichTextBox(frmMain.RichConsola, "Guardando Cuerpos.ini", 0, 162, 232)
    DoEvents

    Open (DirExport & "\Cuerpos.ini") For Binary Access Write As #1
    Put #1, , Datos
    Close #1

    Call AddtoRichTextBox(frmMain.RichConsola, "Cuerpos.ini Exportado", 0, 255, 0)
End Sub

Public Sub ExportarFxs()
    On Error Resume Next

    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim K As Integer

    Dim Datos As String

    Call AddtoRichTextBox(frmMain.RichConsola, "Exportando...", 0, 162, 232)
    DoEvents
    Call DIR_INDEXADOR

    Datos = "[INIT]" & vbCrLf & "NumFxs=" & NumFxs & vbCrLf & vbCrLf

    For i = 1 To NumFxs

        If FxData(i).Animacion > 0 Then
            Datos = Datos & "[FX" & (i) & "]" & vbCrLf
            Datos = Datos & "Animacion=" & FxData(i).Animacion & vbCrLf & "OffsetX=" & FxData(i).OffsetX & vbCrLf & "OffsetY=" & FxData(i).OffsetY & vbCrLf & vbCrLf

        End If

    Next

    Call AddtoRichTextBox(frmMain.RichConsola, "Guardando fxs.ini", 0, 162, 232)
    DoEvents

    Open (DirExport & "\FXs.ini") For Binary Access Write As #1
    Put #1, , Datos
    Close #1

    DoEvents

    Call AddtoRichTextBox(frmMain.RichConsola, "FXs.ini Exportado", 0, 255, 0)
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
    Call AddtoRichTextBox(frmMain.RichConsola, "Exportando...", 0, 162, 232)
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
            Call AddtoRichTextBox(frmMain.RichConsola, "Exportando " & i & " de " & grhCount, 0, 162, 232)
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
    Call AddtoRichTextBox(frmMain.RichConsola, "Graficos.ini Exportado", 0, 255, 0)
    Ocupado = False
End Sub

Public Sub ExportarParticulas()

    On Error Resume Next

    Dim i     As Integer

    Dim j     As Integer

    Dim n     As Integer

    Dim K     As Integer

    Dim Datos As String

    Call AddtoRichTextBox(frmMain.RichConsola, "Exportando...", 0, 162, 232)
    DoEvents
    Call DIR_INDEXADOR

    Datos = "[INIT]" & vbCrLf & "Total=" & TotalStreams & vbCrLf & vbCrLf

    For i = 1 To TotalStreams

        With StreamData(i)
        
            Datos = Datos & "[" & i & "]" & vbCrLf
            Datos = Datos & "Name=" & .name & vbCrLf
            Datos = Datos & "NumOfParticles=" & .NumOfParticles & vbCrLf
            Datos = Datos & "X1=" & .x1 & vbCrLf
            Datos = Datos & "Y1=" & .y2 & vbCrLf
            Datos = Datos & "X2=" & .x2 & vbCrLf
            Datos = Datos & "Y2=" & .y2 & vbCrLf
            Datos = Datos & "Angle=" & .angle & vbCrLf
            Datos = Datos & "VecX1=" & .vecx1 & vbCrLf
            Datos = Datos & "VecX2=" & .vecx2 & vbCrLf
            Datos = Datos & "VecY1=" & .vecy1 & vbCrLf
            Datos = Datos & "VecY2=" & .vecy2 & vbCrLf
            Datos = Datos & "Life1=" & .life1 & vbCrLf
            Datos = Datos & "Life2=" & .life2 & vbCrLf
            Datos = Datos & "Friction=" & .friction & vbCrLf
            Datos = Datos & "Spin=" & .spin & vbCrLf
            Datos = Datos & "Spin_SpeedL=" & .spin_speedL & vbCrLf
            Datos = Datos & "Spin_SpeedH=" & .spin_speedH & vbCrLf
            Datos = Datos & "Grav_Strength=" & .grav_strength & vbCrLf
            Datos = Datos & "Bounce_Strength=" & .bounce_strength & vbCrLf
            Datos = Datos & "AlphaBlend=" & .alphaBlend & vbCrLf
            Datos = Datos & "Gravity=" & .gravity & vbCrLf
            Datos = Datos & "XMove=" & .XMove & vbCrLf
            Datos = Datos & "YMove=" & .YMove & vbCrLf
            Datos = Datos & "move_x1=" & .move_x1 & vbCrLf
            Datos = Datos & "move_x2=" & .move_x2 & vbCrLf
            Datos = Datos & "move_y1=" & .move_y1 & vbCrLf
            Datos = Datos & "move_y2=" & .move_y2 & vbCrLf
            Datos = Datos & "life_counter=" & .life_counter & vbCrLf
            Datos = Datos & "Speed=" & .speed & vbCrLf
            Datos = Datos & "NumGrhs=" & .NumGrhs & vbCrLf
            For j = 1 To .NumGrhs
                Datos = Datos & "GrhList=" & .grh_list(j) & ","
            Next j
            Datos = Datos & vbCrLf
            
            For j = 1 To 4
                Datos = Datos & "ColorSet1=" & .colortint(j).R & .colortint(j).G & "," & .colortint(j).B & vbCrLf & vbCrLf
            Next j
            
        End With

    Next

    Call AddtoRichTextBox(frmMain.RichConsola, "Guardando Particulas.ini", 0, 162, 232)
    DoEvents

    Open (DirExport & "\particulas.ini") For Binary Access Write As #1
    Put #1, , Datos
    Close #1

    DoEvents

    Call AddtoRichTextBox(frmMain.RichConsola, "Particulas.ini Exportado", 0, 255, 0)

End Sub

Public Sub ExportarColores()
'*************************************
'Autor: Lorwik
'Fecha: 05/04/2021
'Descripción: Desindexa los Colores
'*************************************
On Error Resume Next
    Dim i As Integer, j, n, K As Integer
    Dim Datos As String
    
    Call AddtoRichTextBox(frmMain.RichConsola, "Exportando...", 0, 162, 232)
    DoEvents
    
    If FileExist(DirExport & "Colores.dat", vbArchive) = True Then Call Kill(DirExport & "Colores.dat")
    
    Datos = "'Permite customizar los colores de los PJs"
    Datos = Datos & "'todos los valores deben estar entre 0 y 255"
    Datos = Datos & "'los rangos van de 1 a 48 (inclusive). El 0 y el 49,50 estan reservados. Mas arriba son ignorados." & vbCrLf & vbCrLf
    
    For i = 0 To MAXCOLORES
        Datos = Datos & "[" & (i) & "]" & vbCrLf
        Datos = Datos & "R=" & ColoresPJ(i).R & vbCrLf
        Datos = Datos & "G=" & ColoresPJ(i).R & vbCrLf
        Datos = Datos & "B=" & ColoresPJ(i).R & vbCrLf & vbCrLf
    Next
    
    Call AddtoRichTextBox(frmMain.RichConsola, "Guardando Colores.dat", 0, 162, 232)
    DoEvents
    
    Open (DirExport & "Colores.dat") For Binary Access Write As #1
    Put #1, , Datos
    Close #1
    
    DoEvents
    
    Call AddtoRichTextBox(frmMain.RichConsola, "Colores.dat Exportado", 0, 255, 0)
End Sub
