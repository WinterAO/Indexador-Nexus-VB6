Attribute VB_Name = "modCarga"
Option Explicit

Public DirCliente As String

Public DirExport  As String

Public DirIndex   As String

Public DirIndices As String

Public Type tSetupMods

    ' VIDEO
    byMemory    As Integer
    OverrideVertexProcess As Byte
    LimiteFPS As Boolean
    
End Type

Public ClientSetup As tSetupMods

Private Lector    As clsIniReader

'Lista de cabezas
Public Type tHead
    Std As Byte
    Texture As Integer
    startX As Integer
    startY As Integer
End Type

Public heads() As tHead
Public Cascos() As tHead

Public Type tIndiceCuerpo

    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer

End Type

Public Type tIndiceFx

    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer

End Type

Public Type tIndiceArmas

    Weapon(1 To 4) As Long

End Type

Public Type tIndiceEscudos

    Shield(1 To 4) As Long

End Type

Public NumHeads        As Integer
Public NumCuerpos      As Integer
Public NumCascos       As Integer
Public NumEscudosAnims As Integer
Public NumWeaponAnims  As Integer
Public NumFxs          As Integer
Public grhCount        As Long
Public fileVersion     As Long

Public Type RGB
    R As Long
    G As Long
    B As Long
End Type

'Constantes
Public IniPath As String

Public Const INITDIR As String = "Init\"

Public Function profilesFile() As String
    profilesFile = IniPath & INITDIR & "profiles.ini"
End Function

Public Function profileFile(ByVal tag As String) As String
    profileFile = IniPath & INITDIR & "profile-" & tag & ".ini"
End Function

Private Function autoCompletaPath(ByVal Path As String) As String
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'Descripcion: Completa y corrije un path
'*************************************************

    Path = Replace(Path, "/", "\")
    
    If Left(Path, 1) = "\" Then
        ' agrego app.path & path
        Path = App.Path & Path
    End If
    If Right(Path, 1) <> "\" Then
        ' me aseguro que el final sea con "\"
        Path = Path & "\"
    End If
    autoCompletaPath = Path
    
End Function

Public Function CargarConfiguracion() As Boolean

    On Local Error GoTo fileErr:
    
    Dim tStr As String
    Dim NewPath As String
    
    If Not FileExist(profileFile(ProfileTag), vbArchive) Then
        MsgBox "¡No se ha encontrado el archivo de perfil (" & profileFile(ProfileTag) & ") en la carpeta init!", vbOKOnly Or vbExclamation, App.Title
        End
    End If
    
    Set Lector = New clsIniReader
    Call Lector.Initialize(profileFile(ProfileTag))
    
    ' RUTAS
    DirCliente = Lector.GetValue("RUTAS", "DirClient")
    DirExport = Lector.GetValue("RUTAS", "DirExport")
    DirIndex = Lector.GetValue("RUTAS", "DirIndex")
    DirIndices = Lector.GetValue("RUTAS", "DirIndices")
    
    With ClientSetup
        ' VIDEO
        .byMemory = Val(Lector.GetValue("VIDEO", "DynamicMemory"))
        .OverrideVertexProcess = Val(Lector.GetValue("VIDEO", "VertexProcessingOverride"))
        .LimiteFPS = CBool(Lector.GetValue("VIDEO", "LimitarFPS"))
        
        'Client
        DirCliente = autoCompletaPath(Lector.GetValue("RUTAS", "DirClient"))
        
        If FileExist(DirCliente, vbDirectory) = False Or DirCliente = "\" Then
            MsgBox "El directorio del Cliente es incorrecto", vbCritical + vbOKOnly
            
            NewPath = Buscar_Carpeta("Seleccione la carpeta del cliente o donde se encuentren la carpeta de Graficos, Init, etc.", "")
            Call WriteVar(profileFile(ProfileTag), "RUTAS", "DirClient", NewPath)
            DirCliente = NewPath & "\"
        End If
        
        'Index
        DirIndex = autoCompletaPath(Lector.GetValue("RUTAS", "DirIndex"))
        
        If FileExist(DirIndex, vbDirectory) = False Or DirIndex = "\" Then
            MsgBox "El directorio de los Index es incorrecto", vbCritical + vbOKOnly
            
            NewPath = Buscar_Carpeta("Seleccione la carpeta de los index de graficos, personajes, cabezas, etc.", "")
            Call WriteVar(profileFile(ProfileTag), "RUTAS", "DirIndex", NewPath)
            DirIndex = NewPath & "\"
        End If
        
        'Export
        DirExport = autoCompletaPath(Lector.GetValue("RUTAS", "DirExport"))
        
        If FileExist(DirExport, vbDirectory) = False Or DirExport = "\" Then
            MsgBox "El directorio de la carpeta de Exportación es incorrecto", vbCritical + vbOKOnly
            
            NewPath = Buscar_Carpeta("Seleccione la carpeta de los exportados.", "")
            Call WriteVar(profileFile(ProfileTag), "RUTAS", "DirExport", NewPath)
            DirExport = NewPath & "\"
        End If
        
        'Indices
        DirIndices = autoCompletaPath(Lector.GetValue("RUTAS", "DirIndices"))
        
        If FileExist(DirIndices, vbDirectory) = False Or DirIndices = "\" Then
            MsgBox "El directorio de la carpeta de indices es incorrecto", vbCritical + vbOKOnly
            
            NewPath = Buscar_Carpeta("Seleccione la carpeta de indices.", "")
            Call WriteVar(profileFile(ProfileTag), "RUTAS", "DirIndices", NewPath)
            DirIndices = NewPath & "\"
        End If
        
    End With
    
    Set Lector = Nothing
    
    CargarConfiguracion = True
    
    Exit Function
  
fileErr:

    CargarConfiguracion = False
    
End Function

Public Function LoadGrhData() As Boolean

    On Error GoTo ErrorHandler
    
    Dim K As Long
    Dim Grh As Long
    Dim frame As Long
    Dim handle As Integer
    
    If Not FileExist(DirCliente & "\Init\graficos.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Graficos.ind."
        LoadGrhData = False
        Exit Function
    End If
    
    'Open files
    handle = FreeFile()
    Open DirCliente & "\Init\Graficos.ind" For Binary Access Read As handle
    
    With frmMain
        
        .LynxGrh.Clear
        .LynxGrh.Redraw = False
        .LynxGrh.Visible = False
        .LynxGrh.AddColumn "Grh", 0
        .LynxGrh.AddColumn "Tipo", 0
        
    End With
    
    Get handle, , fileVersion
        
    Get handle, , grhCount

    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData

    While Grh <> grhCount
    
        Get handle, , Grh
        
        frmMain.LynxGrh.AddItem Grh
        K = frmMain.LynxGrh.Rows - 1
        frmMain.LynxGrh.CellText(K, 1) = Grh
    
        With GrhData(Grh)

            If Grh <> 0 Then
               
                'Get number of frames
                Get handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
            
                'Minimapa
                .active = True
            
                ReDim .Frames(1 To .NumFrames)
            
                If .NumFrames > 1 Then

                    frmMain.LynxGrh.CellText(K, 1) = "ANIMACION"

                    For frame = 1 To .NumFrames
                        Get handle, , .Frames(frame)
                        If .Frames(frame) <= 0 Or .Frames(frame) > grhCount Then GoTo ErrorHandler
                    Next frame
                
                    Get handle, , .speed
                    If .speed <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                
                Else
                
                    frmMain.LynxGrh.CellText(K, 1) = ""
                    
                    'Read in normal GRH data
                    Get handle, , .FileNum
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , GrhData(Grh).sX
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .sY
                    If .sY < 0 Then GoTo ErrorHandler
                
                    'Compute width and height
                    .TileWidth = .pixelWidth / 32
                    .TileHeight = .pixelHeight / 32
                
                    .Frames(1) = Grh

                End If

            End If

        End With
        
    Wend
    
    Close handle
    
    frmMain.LynxGrh.Visible = True
    frmMain.LynxGrh.Redraw = True
    frmMain.LynxGrh.ColForceFit
    
    DoEvents
    
    LoadGrhData = True
    Exit Function

ErrorHandler:
    Close handle
    MsgBox "Error " & Err.Number & " durante la carga de Graficos.ind! La carga se ha detenido en GRH: " & Grh
    
    frmMain.LynxGrh.Visible = True
    frmMain.LynxGrh.Redraw = True
    frmMain.LynxGrh.ColForceFit

End Function

Public Function CargarCuerpos() As Boolean

    On Error GoTo ErrorHandler

    Dim n            As Integer

    Dim i            As Long

    Dim MisCuerpos() As tIndiceCuerpo

    frmCuerpos.ListaCuerpos.Clear
    
    If Not FileExist(DirCliente & "\Init\personajes.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Personajes.ind."
        CargarCuerpos = False
        Exit Function
    End If

    n = FreeFile()
    Open DirCliente & "\Init\Personajes.ind" For Binary Access Read As #n

    'num de cabezas
    Get #n, , NumCuerpos

    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo

    For i = 1 To NumCuerpos

        Get #n, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
            Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
            Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
            Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
                
            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY

        End If
        
        frmCuerpos.ListaCuerpos.AddItem i
    Next i

    Close #n
    
    CargarCuerpos = True
    
    Exit Function

ErrorHandler:
    Close #n
    'MsgBox "Error " & Err.Number & " durante la carga de Personajes.ind!"
    CargarCuerpos = False
    Resume
    
End Function

Public Function CargarCabezas() As Boolean
    On Error GoTo ErrorHandler:
    
    Dim n            As Integer

    Dim i            As Integer

    frmCabezas.ListaHead.Clear
    
    If Not FileExist(DirCliente & "\Init\head.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo head.ind."
        CargarCabezas = False
        Exit Function
    End If

    n = FreeFile
    Open DirCliente & "\Init\head.ind" For Binary Access Read As #n

    'num de cabezas
    Get #n, , NumHeads

    'Resize array
    ReDim heads(0 To NumHeads) As tHead
            
        For i = 1 To NumHeads
            Get #n, , heads(i).Std
            Get #n, , heads(i).Texture
            Get #n, , heads(i).startX
            Get #n, , heads(i).startY
            
            frmCabezas.ListaHead.AddItem i
        Next i

    Close #n

    CargarCabezas = True

    Exit Function

ErrorHandler:
    Close #n
    'MsgBox "Error " & Err.Number & " durante la carga de Head.ind!"
    CargarCabezas = False
    Resume
    
End Function

Public Function CargarCascos() As Boolean
    On Error GoTo ErrorHandler:
    
    Dim n          As Integer

    Dim i          As Integer

    frmCascos.ListaCascos.Clear

    If Not FileExist(DirCliente & "\Init\helmet.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo helmet.ind."
        CargarCascos = False
        Exit Function
    End If

    n = FreeFile
    Open DirCliente & "\Init\helmet.ind" For Binary Access Read As #n

    'num de cascos
    Get #n, , NumCascos

    'Resize array
    ReDim Cascos(0 To NumCascos) As tHead
    
    For i = 1 To NumCascos
        Get #n, , Cascos(i).Std
        Get #n, , Cascos(i).Texture
        Get #n, , Cascos(i).startX
        Get #n, , Cascos(i).startY
            
        frmCascos.ListaCascos.AddItem i
    Next i
         
    Close #n

    CargarCascos = True

    Exit Function

ErrorHandler:
    Close #n
    'MsgBox "Error " & Err.Number & " durante la carga de Helmet.ind!"
    CargarCascos = False
    Resume
    
End Function

Public Function CargarEscudos() As Boolean
    '*************************************
    'Autor: Lorwik
    'Fecha: ???
    'Descripción: Carga el index de Escudos
    '*************************************
    
    On Error GoTo errhandler:

    Dim n          As Integer

    Dim i          As Long
    
    frmEscudos.ListaEscudos.Clear

    If Not FileExist(DirCliente & "\Init\escudos.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Escudos.ind."
        CargarEscudos = False
        Exit Function
    End If

    n = FreeFile
    Open DirCliente & "\Init\escudos.ind" For Binary Access Read As #n

    'num de escudos
    Get #n, , NumEscudosAnims
        
    'Resize array
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    ReDim Shields(1 To NumEscudosAnims) As tIndiceEscudos
        
    For i = 1 To NumEscudosAnims
        Get #n, , Shields(i)
            
        If Shields(i).Shield(1) Then
            
            Call InitGrh(ShieldAnimData(i).ShieldWalk(1), Shields(i).Shield(1), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(2), Shields(i).Shield(2), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(3), Shields(i).Shield(3), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(4), Shields(i).Shield(4), 0)
            
            frmEscudos.ListaEscudos.AddItem i
            
        End If
            
    Next i
    
    Close #n
    
    CargarEscudos = True
    
    Exit Function

errhandler:
    Close #n
    'MsgBox "Error " & Err.Number & " durante la carga de Escudos.ind!"
    CargarEscudos = False
    Resume
    
End Function

Public Function CargarAnimArmas() As Boolean

    '*************************************
    'Autor: Lorwik
    'Fecha: ???
    'Descripción: Carga el index de Armas
    '*************************************
    On Error GoTo errhandler:

    Dim n          As Integer

    Dim i          As Long
    
    frmArmas.ListaArmas.Clear
    
    If Not FileExist(DirCliente & "\Init\armas.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Armas.ind."
        CargarAnimArmas = False
        Exit Function
    End If
    
    n = FreeFile
    Open DirCliente & "\Init\Armas.ind" For Binary Access Read As #n
    
    'num de armas
    Get #n, , NumWeaponAnims
        
    'Resize array
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    ReDim Weapons(1 To NumWeaponAnims) As tIndiceArmas
        
    For i = 1 To NumWeaponAnims
        Get #n, , Weapons(i)
            
        If Weapons(i).Weapon(1) Then
            
            Call InitGrh(WeaponAnimData(i).WeaponWalk(1), Weapons(i).Weapon(1), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(2), Weapons(i).Weapon(2), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(3), Weapons(i).Weapon(3), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(4), Weapons(i).Weapon(4), 0)
            
            frmArmas.ListaArmas.AddItem i

        End If

    Next i
    
    Close #n
    
    CargarAnimArmas = True
    
    Exit Function

errhandler:
    Close #n
    'MsgBox "Error " & Err.Number & " durante la carga de Armas.ind!"
    CargarAnimArmas = False
    Resume
    
End Function

Public Function CargarFxs() As Boolean

    '*************************************
    'Autor: Lorwik
    'Fecha: ???
    'Descripción: Carga el index de Fxs
    '*************************************
    On Error GoTo errhandler:
    
    Dim n          As Integer

    Dim i          As Long
    
    frmFxs.ListaFxs.Clear
    
    If Not FileExist(DirCliente & "\Init\fxs.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo FXs.ind."
        CargarFxs = False
        Exit Function
    End If
    
    n = FreeFile
    Open DirCliente & "\Init\FXs.ind" For Binary Access Read As #n

    'num de cabezas
    Get #n, , NumFxs
        
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
        
    For i = 1 To NumFxs

        Get #n, , FxData(i)
        frmFxs.ListaFxs.AddItem i
        
    Next i
    
    Close #n
    
    CargarFxs = True
    
    Exit Function
    
errhandler:
    Close #n
   'MsgBox "Error " & Err.Number & " durante la carga de FXs.ind!"
    CargarFxs = False
    Resume
    
End Function

Sub CargarParticulas()
    '*************************************
    'Autor: ????
    'Fecha: ????
    'Descripción: Cargar el archivo de particulas en memoria
    '*************************************

    Dim StreamFile As String

    Dim LoopC      As Long

    Dim i          As Long

    Dim GrhListing As String

    Dim TempSet    As String

    Dim ColorSet   As Long
    
    If Not FileExist(DirExport & "\Particulas.ini", vbArchive) Then
        MsgBox ("No se ha encontrado el archivo Particulas.ini en el directorio: " & DirExport)
        Exit Sub
    End If
    
    StreamFile = DirExport & "\Particulas.ini"
    TotalStreams = Val(GetVar(StreamFile, "INIT", "Total"))
    
    If TotalStreams < 1 Then Exit Sub
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream

    'fill StreamData array with info from particle.ini
    For LoopC = 1 To TotalStreams
        StreamData(LoopC).name = GetVar(StreamFile, Val(LoopC), "Name")
        StreamData(LoopC).NumOfParticles = GetVar(StreamFile, Val(LoopC), "NumOfParticles")
        StreamData(LoopC).x1 = GetVar(StreamFile, Val(LoopC), "X1")
        StreamData(LoopC).y1 = GetVar(StreamFile, Val(LoopC), "Y1")
        StreamData(LoopC).x2 = GetVar(StreamFile, Val(LoopC), "X2")
        StreamData(LoopC).y2 = GetVar(StreamFile, Val(LoopC), "Y2")
        StreamData(LoopC).angle = GetVar(StreamFile, Val(LoopC), "Angle")
        StreamData(LoopC).vecx1 = GetVar(StreamFile, Val(LoopC), "VecX1")
        StreamData(LoopC).vecx2 = GetVar(StreamFile, Val(LoopC), "VecX2")
        StreamData(LoopC).vecy1 = GetVar(StreamFile, Val(LoopC), "VecY1")
        StreamData(LoopC).vecy2 = GetVar(StreamFile, Val(LoopC), "VecY2")
        StreamData(LoopC).life1 = GetVar(StreamFile, Val(LoopC), "Life1")
        StreamData(LoopC).life2 = GetVar(StreamFile, Val(LoopC), "Life2")
        StreamData(LoopC).friction = GetVar(StreamFile, Val(LoopC), "Friction")
        StreamData(LoopC).spin = GetVar(StreamFile, Val(LoopC), "Spin")
        StreamData(LoopC).spin_speedL = GetVar(StreamFile, Val(LoopC), "Spin_SpeedL")
        StreamData(LoopC).spin_speedH = GetVar(StreamFile, Val(LoopC), "Spin_SpeedH")
        StreamData(LoopC).alphaBlend = GetVar(StreamFile, Val(LoopC), "AlphaBlend")
        StreamData(LoopC).gravity = GetVar(StreamFile, Val(LoopC), "Gravity")
        StreamData(LoopC).grav_strength = GetVar(StreamFile, Val(LoopC), "Grav_Strength")
        StreamData(LoopC).bounce_strength = GetVar(StreamFile, Val(LoopC), "Bounce_Strength")
        StreamData(LoopC).XMove = GetVar(StreamFile, Val(LoopC), "XMove")
        StreamData(LoopC).YMove = GetVar(StreamFile, Val(LoopC), "YMove")
        StreamData(LoopC).move_x1 = GetVar(StreamFile, Val(LoopC), "move_x1")
        StreamData(LoopC).move_x2 = GetVar(StreamFile, Val(LoopC), "move_x2")
        StreamData(LoopC).move_y1 = GetVar(StreamFile, Val(LoopC), "move_y1")
        StreamData(LoopC).move_y2 = GetVar(StreamFile, Val(LoopC), "move_y2")
        StreamData(LoopC).life_counter = GetVar(StreamFile, Val(LoopC), "life_counter")
        StreamData(LoopC).speed = Val(GetVar(StreamFile, Val(LoopC), "Speed"))
        StreamData(LoopC).NumGrhs = GetVar(StreamFile, Val(LoopC), "NumGrhs")
        
        ReDim StreamData(LoopC).grh_list(1 To StreamData(LoopC).NumGrhs) As Long
        GrhListing = GetVar(StreamFile, Val(LoopC), "Grh_List")
        
        For i = 1 To StreamData(LoopC).NumGrhs
            StreamData(LoopC).grh_list(i) = CLng(ReadField(str(i), GrhListing, 44))
        Next i
        
        'StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        
        For ColorSet = 1 To 4
            TempSet = GetVar(StreamFile, Val(LoopC), "ColorSet" & ColorSet)
            StreamData(LoopC).colortint(ColorSet - 1).R = ReadField(1, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).G = ReadField(2, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).B = ReadField(3, TempSet, 44)
        Next ColorSet

        frmParticleEditor.ListParticulas.AddItem LoopC & " - " & StreamData(LoopC).name
    Next LoopC

End Sub

Public Function CargarColores() As Boolean

On Error GoTo errhandler:

    If Not FileExist(DirExport & "colores.dat", vbNormal) Then Exit Function

    Dim LeerINI As New clsIniReader
    Call LeerINI.Initialize(DirExport & "colores.dat")
    
    Dim i As Long
    
    For i = 0 To MAXCOLORES '48, 49 y 50 reservados para atacables, ciudadano y criminal
        ColoresPJ(i).R = LeerINI.GetValue(CStr(i), "R")
        ColoresPJ(i).G = LeerINI.GetValue(CStr(i), "G")
        ColoresPJ(i).B = LeerINI.GetValue(CStr(i), "B")
    Next i
    
    Set LeerINI = Nothing
    
    CargarColores = True
    
errhandler:

End Function

Public Sub CargarIndices()
    '*************************************************
    'Autor: Lorwik
    'Fecha: 13/11/2023
    'Descripcion: Carga los indices
    '*************************************************

    On Error GoTo fallo

    Dim Leer As New clsIniReader

    Dim i    As Integer

    Dim K    As Long
    
    If FileExist(DirIndices & "indices.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'indices.ini'", vbCritical
        End

    End If
    
    Leer.Initialize DirIndices & "indices.ini"
    MaxSup = Leer.GetValue("INIT", "Referencias")
    
    ReDim SupData(MaxSup) As SupData
    
    With frmIndices
    
        .LynxIndices.Clear
        .LynxIndices.Redraw = False
        .LynxIndices.Visible = False

        .LynxIndices.AddColumn "Indice", 0
        .LynxIndices.AddColumn "Grh", 0
        .LynxIndices.AddColumn "Nombre", 3
    
        For i = 0 To MaxSup
            SupData(i).name = Leer.GetValue("REFERENCIA" & i, "Nombre")
            SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
            SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
            SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
            SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
            SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
        
            .LynxIndices.AddItem i
            K = .LynxIndices.Rows - 1
            .LynxIndices.CellText(K, 1) = SupData(i).Grh
            .LynxIndices.CellText(K, 2) = SupData(i).name
        Next
    
        .LynxIndices.Visible = True
        .LynxIndices.Redraw = True
        .LynxIndices.ColForceFit

    End With
    
    DoEvents
    
    Set Leer = Nothing
    
    Exit Sub
fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de \indices.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
    
End Sub
