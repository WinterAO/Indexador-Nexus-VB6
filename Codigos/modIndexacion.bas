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

Public Function IndexarfromMemory() As Boolean

    On Error GoTo ErrorHandler:

    Dim Grh    As Long

    Dim handle As Integer

    Dim frame  As Long
    
    handle = FreeFile()

    If LenB(Dir(DirIndex & "\graficos.ind")) <> 0 Then Call Kill(DirIndex & "\graficos.ind")
    DoEvents
    
    frmMain.GRHt.Text = "Indexando " & grhCount & " Grh's"
    
    Open DirIndex & "\graficos.ind" For Binary Access Write As handle
    
    Seek handle, 1

    fileVersion = fileVersion + 1

    Put handle, , fileVersion
    Put handle, , grhCount
    
    For Grh = 1 To grhCount
        
        With GrhData(Grh)

            If .NumFrames > 1 Then
                Put handle, , Grh
                
                Put handle, , .NumFrames

                For frame = 1 To .NumFrames
                    Put handle, , .Frames(frame)
                Next frame
                
                Put handle, , .speed
                
            ElseIf .NumFrames = 1 Then
                Put handle, , Grh
                Put handle, , .NumFrames
                Put handle, , .FileNum
                Put handle, , .pixelWidth
                Put handle, , .pixelHeight
                Put handle, , .sX
                Put handle, , .sY

            End If
            
        End With

    Next Grh
    
    Close handle
    
    IndexarfromMemory = True

    Exit Function
    
ErrorHandler:
    Close handle
    IndexarfromMemory = False

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

    If (nCuerpos > 200000 Or nCuerpos <= 0) Then
        MsgBox "La valor de 'NumBodies' es invalido!", vbCritical
        Exit Function

    End If

    If LenB(Dir(DirIndex & "\personajes.ind")) <> 0 Then Call Kill(DirIndex & "\personajes.ind")
    DoEvents

    Open DirIndex & "\personajes.ind" For Binary Access Write As handle

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

On Error GoTo fallo

    Dim i As Integer, j, K As Integer
    Dim nF As Integer
    Dim NumHeads As Integer
    
    DoEvents
    
    Dim LeerINI As New clsIniReader
    Call LeerINI.Initialize(DirExport & "Head.ini")
    
    NumHeads = CInt(LeerINI.GetValue("INIT", "NumHeads"))
    
    ReDim HeadsT(0 To NumHeads) As tHead
    
    For i = 1 To NumHeads
        HeadsT(i).Std = Val(LeerINI.GetValue("HEAD" & i, "Std"))
        HeadsT(i).Texture = Val(LeerINI.GetValue("HEAD" & i, "FileNum"))
        HeadsT(i).startX = Val(LeerINI.GetValue("HEAD" & i, "OffSetX"))
        HeadsT(i).startY = Val(LeerINI.GetValue("HEAD" & i, "OffSetY"))
    Next i
    
    nF = FreeFile
    Open DirIndex & "Head.ind" For Binary Access Write As #nF
    
    Put #nF, , NumHeads
    
    For i = 1 To NumHeads
        Put #nF, , HeadsT(i)
    Next
    
    DoEvents
    Close #nF
    IndexarCabezas = True
    
    Exit Function
    
fallo:
    MsgBox "Error en Cabezas.ini"
    IndexarCabezas = False
    
End Function

Public Function IndexarCascos() As Boolean

 On Error GoTo fallo

    Dim i As Integer, j, K As Integer
    Dim nF As Integer
    Dim NumCascos As Integer
    
    DoEvents
    
    Dim LeerINI As New clsIniReader
    Call LeerINI.Initialize(DirExport & "Helmet.ini")
    
    NumCascos = CInt(LeerINI.GetValue("INIT", "NumCascos"))
    
    ReDim HelmesT(0 To NumCascos) As tHead
    
    For i = 1 To NumCascos
        HelmesT(i).Std = Val(LeerINI.GetValue("CASCO" & i, "Std"))
        HelmesT(i).Texture = Val(LeerINI.GetValue("CASCO" & i, "FileNum"))
        HelmesT(i).startX = Val(LeerINI.GetValue("CASCO" & i, "OffSetX"))
        HelmesT(i).startY = Val(LeerINI.GetValue("CASCO" & i, "OffSetY"))
    Next i
    
    nF = FreeFile
    Open DirIndex & "Helmet.ind" For Binary Access Write As #nF
    
    Put #nF, , NumCascos
    
    For i = 1 To NumCascos
        Put #nF, , HelmesT(i)
    Next

    DoEvents
    Close #nF
    IndexarCascos = True
    
    Exit Function
fallo:
    MsgBox "Error en Cabezas.ini"
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

Public Function IndexarParticulas() As Boolean

    On Error GoTo ErrorHandler:

    Dim LoopC As Long
    Dim i As Integer
    Dim handle As Integer
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim ParticulasStream As Stream
    Dim TotalParticulas As Integer
    Dim Leer As New clsIniReader

    Call Leer.Initialize(DirExport & "\particulas.ini")

    TotalParticulas = Val(Leer.GetValue("INIT", "Total"))
    
    If LenB(Dir(DirIndex & "particulas.ind")) <> 0 Then Call Kill(DirIndex & "particulas.ind")
    DoEvents

    handle = FreeFile
    Open DirIndex & "particulas.ind" For Binary As handle
    
    Put handle, , TotalParticulas
    
    'fill StreamData array with info from Particles.ini
    For LoopC = 1 To TotalParticulas
        With ParticulasStream
            '.name = Leer.GetValue(Val(LoopC), "Name")
            'Put handle, , ParticulasStream.name
            
            .NumOfParticles = Leer.GetValue(Val(LoopC), "NumOfParticles")
            Put handle, , ParticulasStream.NumOfParticles
            
            .NumGrhs = Leer.GetValue(Val(LoopC), "NumGrhs")
            Put handle, , ParticulasStream.NumGrhs
            
            .id = LoopC
            Put handle, , ParticulasStream.id
            
            .x1 = Leer.GetValue(Val(LoopC), "X1")
            Put handle, , ParticulasStream.x1
            '
            .y1 = Leer.GetValue(Val(LoopC), "Y1")
            Put handle, , ParticulasStream.y1
            
            .x2 = Leer.GetValue(Val(LoopC), "X2")
            Put handle, , ParticulasStream.x2
            
            .y2 = Leer.GetValue(Val(LoopC), "Y2")
            Put handle, , ParticulasStream.y2
            
            .angle = Leer.GetValue(Val(LoopC), "Angle")
            Put handle, , ParticulasStream.angle
            
            .vecx1 = Leer.GetValue(Val(LoopC), "VecX1")
            Put handle, , ParticulasStream.vecx1
            
            .vecx2 = Leer.GetValue(Val(LoopC), "VecX2")
            Put handle, , ParticulasStream.vecx2
            
            .vecy1 = Leer.GetValue(Val(LoopC), "VecY1")
            Put handle, , ParticulasStream.vecy1
            
            .vecy2 = Leer.GetValue(Val(LoopC), "VecY2")
            Put handle, , ParticulasStream.vecy2
            
            .life1 = Leer.GetValue(Val(LoopC), "Life1")
            Put handle, , ParticulasStream.life1
            
            .life2 = Leer.GetValue(Val(LoopC), "Life2")
            Put handle, , ParticulasStream.life2
            
            .friction = Leer.GetValue(Val(LoopC), "Friction")
            Put handle, , ParticulasStream.friction
            
            .spin = Leer.GetValue(Val(LoopC), "Spin")
            Put handle, , ParticulasStream.spin
            
            .spin_speedL = Leer.GetValue(Val(LoopC), "Spin_SpeedL")
            Put handle, , ParticulasStream.spin_speedL
            
            .spin_speedH = Leer.GetValue(Val(LoopC), "Spin_SpeedH")
            Put handle, , ParticulasStream.spin_speedH
            
            .alphaBlend = Leer.GetValue(Val(LoopC), "AlphaBlend")
            Put handle, , ParticulasStream.alphaBlend
            
            .gravity = Leer.GetValue(Val(LoopC), "Gravity")
            Put handle, , ParticulasStream.gravity
            
            .grav_strength = Leer.GetValue(Val(LoopC), "Grav_Strength")
            Put handle, , ParticulasStream.grav_strength
            
            .bounce_strength = Leer.GetValue(Val(LoopC), "Bounce_Strength")
            Put handle, , ParticulasStream.bounce_strength
            
            .XMove = Leer.GetValue(Val(LoopC), "XMove")
            Put handle, , ParticulasStream.XMove
            
            .YMove = Leer.GetValue(Val(LoopC), "YMove")
            Put handle, , ParticulasStream.YMove
            
            .move_x1 = Leer.GetValue(Val(LoopC), "move_x1")
            Put handle, , ParticulasStream.move_x1
            
            .move_x2 = Leer.GetValue(Val(LoopC), "move_x2")
            Put handle, , ParticulasStream.move_x2
            
            .move_y1 = Leer.GetValue(Val(LoopC), "move_y1")
            Put handle, , ParticulasStream.move_y1
            
            .move_y2 = Leer.GetValue(Val(LoopC), "move_y2")
            Put handle, , ParticulasStream.move_y2
            
            .speed = Val(Leer.GetValue(Val(LoopC), "Speed"))
            Put handle, , ParticulasStream.speed
            
            .life_counter = Leer.GetValue(Val(LoopC), "life_counter")
            Put handle, , ParticulasStream.life_counter
            
            ReDim .grh_list(1 To .NumGrhs)
            GrhListing = Leer.GetValue(Val(LoopC), "Grh_List")
            
            For i = 1 To .NumGrhs
                .grh_list(i) = ReadField(i, GrhListing, Asc(","))
                Put handle, , ParticulasStream.grh_list(i)
            Next i
            
            .grh_list(i - 1) = .grh_list(i - 1)
            
            For ColorSet = 1 To 4
                TempSet = Leer.GetValue(Val(LoopC), "ColorSet" & ColorSet)
                .colortint(ColorSet - 1).R = ReadField(1, TempSet, Asc(","))
                .colortint(ColorSet - 1).G = ReadField(2, TempSet, Asc(","))
                .colortint(ColorSet - 1).B = ReadField(3, TempSet, Asc(","))
                Put handle, , ParticulasStream.colortint(ColorSet - 1).R
                Put handle, , ParticulasStream.colortint(ColorSet - 1).G
                Put handle, , ParticulasStream.colortint(ColorSet - 1).B
            Next ColorSet

            'Put #handle, , ParticulasStream
    
        End With
    Next LoopC

    Close handle
    
    Set Leer = Nothing
    IndexarParticulas = True
    
    Close handle
    
    Exit Function

ErrorHandler:
    Close handle
    IndexarParticulas = False
End Function

Public Function IndexarColores() As Boolean
'*************************************
'Autor: Lorwik
'Fecha: 30/08/2020
'Descripción: Guarda los colores en un archivo binario
'*************************************

    Dim n As Integer
    Dim i As Byte
    
    If CargarColores Then
    
        n = FreeFile
        Open DirIndex & "\Colores.ind" For Binary Access Write As #n
        
            For i = 0 To MAXCOLORES
                Put #n, , ColoresPJ(i).R
                Put #n, , ColoresPJ(i).G
                Put #n, , ColoresPJ(i).B
            Next i
        
        Close #n
        DoEvents
        
        IndexarColores = True
    
    Else
    
        frmMain.GRHt.Text = "Error al indexar Colores.dat. No se ha podido leer el archivo de origen."
        IndexarColores = False
        
    End If
    
    Exit Function

ErrorHandler:
    Close #n
    IndexarColores = False
End Function

Public Function IndexarGUI() As Boolean
'*************************************
'Autor: Lorwik
'Fecha: 30/08/2020
'Descripción: Guarda la GUI en un archivo binario
'*************************************

    On Error GoTo ErrorHandler:

    Dim n               As Integer
    Dim Leer            As New clsIniReader
    Dim i               As Integer
    Dim NumButtons      As Integer
    Dim NumConnectMap   As Byte

    If FileExist(DirExport & "GUI.dat", vbArchive) = True Then
        Call Leer.Initialize(DirExport & "GUI.dat")
        
        n = FreeFile
        Open DirIndex & "\GUI.ind" For Binary Access Write As #n
            
            NumButtons = Val(Leer.GetValue("INIT", "NumButtons"))
            Put #n, , NumButtons
            
            NumConnectMap = Val(Leer.GetValue("INIT", "NumMaps"))
            Put #n, , NumConnectMap
            
            'Mapas de GUI
            For i = 1 To NumConnectMap
                Put #n, , CInt(Leer.GetValue("MAPA" & i, "Map"))
                Put #n, , CInt(Leer.GetValue("MAPA" & i, "X"))
                Put #n, , CInt(Leer.GetValue("MAPA" & i, "Y"))
            Next i
            
            'Posiciones de los PJ
            For i = 1 To 10
                Put #n, , CInt(Leer.GetValue("PJPos" & i, "X"))
                Put #n, , CInt(Leer.GetValue("PJPos" & i, "Y"))
            Next i
            
            'Posiciones de los botones
            For i = 1 To NumButtons
                Put #n, , CInt(Leer.GetValue("BUTTON" & i, "X"))
                Put #n, , CInt(Leer.GetValue("BUTTON" & i, "Y"))
                Put #n, , CInt(Leer.GetValue("BUTTON" & i, "PosX"))
                Put #n, , CInt(Leer.GetValue("BUTTON" & i, "PosY"))
                Put #n, , CLng(Leer.GetValue("BUTTON" & i, "GrhNormal"))
       
            Next i
        
        Close #n
        DoEvents
        
        IndexarGUI = True
        
    Else
    
        frmMain.GRHt.Text = "Error al indexar GUID.dat. No se ha encontrado el archivo de origen."
        
        IndexarGUI = False
    
    End If
    
    Set Leer = Nothing
     
    Exit Function

ErrorHandler:
    Set Leer = Nothing
    Close #n
    IndexarGUI = False
End Function
