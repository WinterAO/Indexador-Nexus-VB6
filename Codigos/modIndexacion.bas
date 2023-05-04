Attribute VB_Name = "modIndexacion"
Option Explicit

Public Function IndexarGraficos() As Boolean

    On Error GoTo ErrorHandler:

    Dim Grh        As Long
    Dim handle     As Integer
    Dim totalGrh   As Long
    Dim fVersion   As Long
    Dim Leer       As New clsIniReader
    Dim Datos      As String
    Dim DatoR()    As String
    Dim tF         As Integer
    Dim DatosGrh   As GrhData
    
    handle = FreeFile()

    Call Leer.Initialize(DirExport & "\graficos.ini")

    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function

    End If
    
    totalGrh = Val(Leer.GetValue("INIT", "NumGrh"))
    fVersion = Val(Leer.GetValue("INIT", "Version"))
    
    If (fVersion = 0) Then
        'MsgBox "El valor de 'Version' es invalido!", vbCritical
        'Exit Function
        fVersion = 1
    ElseIf (totalGrh > 200000 Or totalGrh <= 0) Then
        MsgBox "La valor de 'NumGrh' es invalido!", vbCritical
        Exit Function

    End If

    If LenB(Dir(DirIndex & "\graficos.ind")) <> 0 Then Call Kill(DirIndex & "\graficos.ind")
    DoEvents
    
    Open DirIndex & "\graficos.ind" For Binary Access Write As handle
    
    Seek handle, 1
    
    Put handle, , MiCabecera
    Put handle, , fVersion
    Put handle, , totalGrh
    
    For Grh = 1 To totalGrh
        DatosGrh.sX = 0
        DatosGrh.sY = 0
        DatosGrh.pixelWidth = 0
        DatosGrh.pixelHeight = 0
        DatosGrh.FileNum = 0
        DatosGrh.NumFrames = 0
        DatosGrh.speed = 0
        'Erase DatosGrh.Frames()
        
        Datos = Leer.GetValue("Graphics", "Grh" & Grh)

        If LenB(Datos) <> 0 Then
            DatoR() = Split(Datos, "-")

            If DatoR(0) > 1 Then
                Put handle, , Grh
                DatosGrh.NumFrames = Val(DatoR(0))
                Put handle, , DatosGrh.NumFrames
                ReDim DatosGrh.Frames(1 To DatosGrh.NumFrames)
                tF = 1

                While Not DatosGrh.NumFrames < tF

                    DatosGrh.Frames(tF) = Val(DatoR(tF))
                    Put handle, , DatosGrh.Frames(tF)
                    tF = tF + 1
                Wend
                
                DatosGrh.speed = Val(DatoR(tF))
                Put handle, , DatosGrh.speed
            ElseIf DatoR(0) = 1 Then
                Put handle, , Grh
                DatosGrh.NumFrames = Val(DatoR(0))
                Put handle, , DatosGrh.NumFrames
                DatosGrh.FileNum = Val(DatoR(1))
                Put handle, , DatosGrh.FileNum
                DatosGrh.pixelWidth = Val(DatoR(4))
                Put handle, , DatosGrh.pixelWidth
                DatosGrh.pixelHeight = Val(DatoR(5))
                Put handle, , DatosGrh.pixelHeight
                DatosGrh.sX = Val(DatoR(2))
                Put handle, , DatosGrh.sX
                DatosGrh.sY = Val(DatoR(3))
                Put handle, , DatosGrh.sY

            End If

        End If

    Next Grh
    
    Close handle
    
    IndexarGraficos = True

    Exit Function
    
ErrorHandler:
    Close handle
    IndexarGraficos = False

End Function

Public Function IndexarCuerpos() As Boolean

    On Error GoTo ErrorHandler:

    Dim handle     As Integer
    Dim handleW    As Integer
    Dim i          As Long
    Dim nCuerpos As Integer
    Dim MisCuerpos As tIndiceCuerpo
    Dim Leer       As New clsIniReader
    
    handle = FreeFile()
    Call Leer.Initialize(DirExport & "\personajes.ini")

    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function

    End If

    nCuerpos = Val(Leer.GetValue("INIT", "NumBodies"))

    If (NumCuerpos > 200000 Or NumCuerpos <= 0) Then
        MsgBox "La valor de 'NumBodies' es invalido!", vbCritical
        Exit Function

    End If

    If LenB(Dir(DirIndex & "\personajes.ind")) <> 0 Then Call Kill(DirIndex & "\personajes.ind")
    DoEvents

    Open DirIndex & "\personajes.ind" For Binary Access Write As handle

    Put handle, , MiCabecera
    Put handle, , nCuerpos
    
    For i = 1 To nCuerpos
        MisCuerpos.Body(1) = Val(Leer.GetValue("BODY" & i, "Walk1"))
        MisCuerpos.Body(2) = Val(Leer.GetValue("BODY" & i, "Walk2"))
        MisCuerpos.Body(3) = Val(Leer.GetValue("BODY" & i, "Walk3"))
        MisCuerpos.Body(4) = Val(Leer.GetValue("BODY" & i, "Walk4"))
        MisCuerpos.HeadOffsetX = Val(Leer.GetValue("BODY" & i, "HeadOffsetX"))
        MisCuerpos.HeadOffsetY = Val(Leer.GetValue("BODY" & i, "HeadOffsetY"))
        
        Put handle, , MisCuerpos
    Next i

    Close handle
    
    IndexarCuerpos = True
    Exit Function

ErrorHandler:
    Close handle
    IndexarCuerpos = False
    
End Function

Public Function IndexarCabezas() As Boolean

    On Error GoTo ErrorHandler:

    Dim handle     As Integer
    Dim i          As Long
    Dim nheads     As Integer
    Dim MisCabezas As tIndiceCabeza
    Dim Leer       As New clsIniReader

    handle = FreeFile()
    
    Call Leer.Initialize(DirExport & "\cabezas.ini")

    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function

    End If

    nheads = Val(Leer.GetValue("INIT", "NumHeads"))

    If (nheads > 200000 Or nheads <= 0) Then
        MsgBox "La valor de 'nHeads' es invalido!", vbCritical
        Exit Function

    End If

    If LenB(Dir(DirIndex & "\cabezas.ind")) <> 0 Then Call Kill(DirIndex & "\cabezas.ind")
    DoEvents

    Open DirIndex & "\cabezas.ind" For Binary Access Write As handle

    Put handle, , MiCabecera
    Put handle, , nheads
    
    For i = 1 To nheads
        MisCabezas.Head(1) = Val(Leer.GetValue("HEAD" & i, "Head1"))
        MisCabezas.Head(2) = Val(Leer.GetValue("HEAD" & i, "Head2"))
        MisCabezas.Head(3) = Val(Leer.GetValue("HEAD" & i, "Head3"))
        MisCabezas.Head(4) = Val(Leer.GetValue("HEAD" & i, "Head4"))
        
        Put handle, , MisCabezas
    Next i

    Close handle
    
    IndexarCabezas = True
    Exit Function

ErrorHandler:
    Close handle
    IndexarCabezas = False
    
End Function

Public Function IndexarCascos() As Boolean

    On Error GoTo ErrorHandler:

    Dim handle    As Integer
    Dim i         As Long
    Dim nHelmets  As Integer
    Dim MisCascos As tIndiceCabeza
    Dim Leer      As New clsIniReader

    handle = FreeFile()
    
    Call Leer.Initialize(DirExport & "\cascos.ini")

    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function

    End If

    nHelmets = Val(Leer.GetValue("INIT", "NumHelmets"))

    If (nHelmets > 200000 Or nHelmets <= 0) Then
        MsgBox "La valor de 'nHelmets' es invalido!", vbCritical
        Exit Function

    End If

    If LenB(Dir(DirIndex & "\cascos.ind")) <> 0 Then Call Kill(DirIndex & "\cascos.ind")
    DoEvents

    Open DirIndex & "\cascos.ind" For Binary Access Write As handle

    Put handle, , MiCabecera
    Put handle, , nHelmets
    
    For i = 1 To nHelmets
        MisCascos.Head(1) = Val(Leer.GetValue("HELMET" & i, "Helmet1"))
        MisCascos.Head(2) = Val(Leer.GetValue("HELMET" & i, "Helmet2"))
        MisCascos.Head(3) = Val(Leer.GetValue("HELMET" & i, "Helmet3"))
        MisCascos.Head(4) = Val(Leer.GetValue("HELMET" & i, "Helmet4"))
        
        Put handle, , MisCascos
    Next i

    Close handle
    
    IndexarCascos = True
    Exit Function

ErrorHandler:
    Close handle
    IndexarCascos = False
    
End Function

Public Function IndexarFXs() As Boolean

    On Error GoTo ErrorHandler:

    Dim handle  As Integer
    Dim i       As Long
    Dim NumFX   As Integer
    Dim MisFXs  As tIndiceFx
    Dim Leer    As New clsIniReader

    handle = FreeFile()
    
    Call Leer.Initialize(DirExport & "\fxs.ini")

    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function

    End If

    NumFX = Val(Leer.GetValue("INIT", "NumFxs"))

    If (NumFX > 200000 Or NumFX <= 0) Then
        MsgBox "La valor de 'NumFxs' es invalido!", vbCritical
        Exit Function

    End If

    If LenB(Dir(DirIndex & "\fxs.ind")) <> 0 Then Call Kill(DirIndex & "\fxs.ind")
    DoEvents

    Open DirIndex & "\fxs.ind" For Binary Access Write As handle
    Put handle, , MiCabecera
    Put handle, , NumFX
    
    For i = 1 To NumFX
        MisFXs.Animacion = Val(Leer.GetValue("FX" & i, "Animacion"))
        MisFXs.OffsetX = Val(Leer.GetValue("FX" & i, "OffsetX"))
        MisFXs.OffsetY = Val(Leer.GetValue("FX" & i, "OffsetY"))
        Put handle, , MisFXs
    Next i

    Close handle
    
    IndexarFXs = True
    Exit Function

ErrorHandler:
    Close handle
    IndexarFXs = False
    
End Function

Public Function IndexarArmas() As Boolean

    On Error GoTo ErrorHandler:

    Dim handle    As Integer
    Dim i         As Long
    Dim nArmas    As Integer
    Dim Weapon    As tIndiceArmas
    Dim Leer      As New clsIniReader

    handle = FreeFile()
    
    Call Leer.Initialize(DirExport & "\armas.ini")

    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function

    End If

    nArmas = Val(Leer.GetValue("INIT", "NumArmas"))

    If (nArmas > 200000 Or nArmas <= 0) Then
        MsgBox "La valor de 'nArmas' es invalido!", vbCritical
        Exit Function

    End If

    If LenB(Dir(DirIndex & "\armas.ind")) <> 0 Then Call Kill(DirIndex & "\armas.ind")
    DoEvents

    Open DirIndex & "\armas.ind" For Binary Access Write As handle

    Put handle, , MiCabecera
    Put handle, , nArmas
    
    For i = 1 To nArmas
        Weapon.Weapon(1) = Val(Leer.GetValue("ARMA" & i, "Dir1"))
        Weapon.Weapon(2) = Val(Leer.GetValue("ARMA" & i, "Dir2"))
        Weapon.Weapon(3) = Val(Leer.GetValue("ARMA" & i, "Dir3"))
        Weapon.Weapon(4) = Val(Leer.GetValue("ARMA" & i, "Dir4"))
        
        Put handle, , Weapon
    Next i

    Close handle
    
    IndexarArmas = True
    Exit Function

ErrorHandler:
    Close handle
    IndexarArmas = False
    
End Function

Public Function IndexarEscudos() As Boolean

    On Error GoTo ErrorHandler:

    Dim handle    As Integer
    Dim i         As Long
    Dim nEscudos  As Integer
    Dim Shield    As tIndiceEscudos
    Dim Leer      As New clsIniReader

    handle = FreeFile()
    
    Call Leer.Initialize(DirExport & "\escudos.ini")

    If Leer.KeyExists("INIT") = False Then
        MsgBox "Formato invalido!", vbCritical
        Exit Function

    End If

    nEscudos = Val(Leer.GetValue("INIT", "NumEscudos"))

    If (nEscudos > 200000 Or nEscudos <= 0) Then
        MsgBox "La valor de 'nEscudos' es invalido!", vbCritical
        Exit Function

    End If

    If LenB(Dir(DirIndex & "\escudos.ind")) <> 0 Then Call Kill(DirIndex & "\escudos.ind")
    DoEvents

    Open DirIndex & "\escudos.ind" For Binary Access Write As handle

    Put handle, , MiCabecera
    Put handle, , nEscudos
    
    For i = 1 To nEscudos
        Shield.Shield(1) = Val(Leer.GetValue("ESC" & i, "Dir1"))
        Shield.Shield(2) = Val(Leer.GetValue("ESC" & i, "Dir2"))
        Shield.Shield(3) = Val(Leer.GetValue("ESC" & i, "Dir3"))
        Shield.Shield(4) = Val(Leer.GetValue("ESC" & i, "Dir4"))
        
        Put handle, , Shield
    Next i

    Close handle
    
    IndexarEscudos = True
    Exit Function

ErrorHandler:
    Close handle
    IndexarEscudos = False
    
End Function
