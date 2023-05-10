Attribute VB_Name = "modCarga"
Option Explicit

Public DirCliente As String

Public DirExport  As String

Public DirIndex   As String

Public Type tSetupMods

    ' VIDEO
    byMemory    As Integer
    OverrideVertexProcess As Byte
    
End Type

Public ClientSetup As tSetupMods

Public Type tCabecera

    Desc As String * 255
    CRC As Long
    MagicWord As Long

End Type

Public MiCabecera As tCabecera

Private Lector    As clsIniReader

'Lista de cabezas
Public Type tIndiceCabeza

    Head(1 To 4) As Long

End Type

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

Public Numheads        As Integer
Public NumCuerpos      As Integer
Public NumCascos       As Integer
Public NumEscudosAnims As Integer
Public NumWeaponAnims  As Integer
Public NumFxs          As Integer
Public grhCount        As Long
Public fileVersion     As Long

Public Type RGB
    r As Long
    g As Long
    b As Long
End Type

Public Type Stream
    'name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    alphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    speed As Single
    life_counter As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
End Type

Public Sub IniciarCabecera()

    With MiCabecera
        .Desc = "NexusAO mod Argentum Online by Noland Studios."
        .CRC = Rnd * 245
        .MagicWord = Rnd * 92

    End With
    
End Sub

Public Function CargarConfiguracion() As Boolean

    On Local Error GoTo fileErr:
    
    Dim tStr As String
    
    Set Lector = New clsIniReader
    Call Lector.Initialize(App.Path & "\Config.ini")
    
    ' RUTAS
    DirCliente = Lector.GetValue("RUTAS", "DirClient")
    DirExport = Lector.GetValue("RUTAS", "DirExport")
    DirIndex = Lector.GetValue("RUTAS", "DirIndex")
    
    With ClientSetup
        ' VIDEO
        .byMemory = Lector.GetValue("VIDEO", "DynamicMemory")
        .OverrideVertexProcess = CByte(Lector.GetValue("VIDEO", "VertexProcessingOverride"))
    End With
    
    Set Lector = Nothing
    
    CargarConfiguracion = True
    
    Exit Function
  
fileErr:

    CargarConfiguracion = False
    
End Function

Public Function LoadGrhData() As Boolean

    On Error GoTo ErrorHandler

    Dim Grh        As Long

    Dim Frame      As Long

    Dim handle     As Integer

    Dim LaCabecera As tCabecera

    frmMain.Listado.Clear
    
    If Not FileExist(DirCliente & "\Scripts\graficos.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Graficos.ind."
        LoadGrhData = False
        Exit Function
    End If
    
    'Open files
    handle = FreeFile()
    Open DirCliente & "\Scripts\Graficos.ind" For Binary Access Read As handle

    Get handle, , LaCabecera
    
    Get handle, , fileVersion
        
    Get handle, , grhCount

    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData

    While Grh <> grhCount
    
        Get handle, , Grh
    
        With GrhData(Grh)

            If Grh <> 0 Then
                .Active = True
                'Get number of frames
                Get handle, , .NumFrames

                If .NumFrames <= 0 Then GoTo ErrorHandler
            
                ReDim .Frames(1 To .NumFrames)
            
                If .NumFrames > 1 Then
                    frmMain.Listado.AddItem Grh & " (ANIMACION)"

                    For Frame = 1 To .NumFrames
                        Get handle, , .Frames(Frame)

                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then GoTo ErrorHandler
                    Next Frame
                
                    Get handle, , .speed

                    If .speed <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight

                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth

                    If .pixelWidth <= 0 Then GoTo ErrorHandler

                    ' .TileWidth = GrhData(.Frames(1)).TileWidth
                    'If .TileWidth <= 0 Then GoTo ErrorHandler

                    ' .TileHeight = GrhData(.Frames(1)).TileHeight
                    'If .TileHeight <= 0 Then GoTo ErrorHandler
                
                Else
                    'Read in normal GRH data
                    frmMain.Listado.AddItem Grh
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
                    '.TileWidth = .pixelWidth / TilePixelHeight
                    '.TileHeight = .pixelHeight / TilePixelWidth
                
                    .Frames(1) = Grh

                End If

            End If

        End With
        
    Wend
    
    Close handle
    
    LoadGrhData = True
    Exit Function

ErrorHandler:
    Close handle
    MsgBox "Error " & Err.Number & " durante la carga de Graficos.ind! La carga se ha detenido en GRH: " & Grh

End Function

Public Function CargarCuerpos() As Boolean

    On Error GoTo ErrorHandler

    Dim n            As Integer

    Dim i            As Long

    Dim LaCabecera   As tCabecera

    Dim MisCuerpos() As tIndiceCuerpo

    frmMain.ListaCuerpos.Clear
    
    If Not FileExist(DirCliente & "\Scripts\personajes.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Personajes.ind."
        CargarCuerpos = False
        Exit Function
    End If

    n = FreeFile()
    Open DirCliente & "\Scripts\Personajes.ind" For Binary Access Read As #n

    'cabecera
    Get #n, , LaCabecera

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
                
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY

        End If
        
        frmMain.ListaCuerpos.AddItem i
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

    Dim LaCabecera   As tCabecera

    Dim MisCabezas() As tIndiceCabeza

    frmMain.ListaHead.Clear
    
    If Not FileExist(DirCliente & "\Scripts\cabezas.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Cabezas.ind."
        CargarCabezas = False
        Exit Function
    End If

    n = FreeFile
    Open DirCliente & "\Scripts\cabezas.ind" For Binary Access Read As #n

    'cabecera
    Get #n, , LaCabecera

    'num de cabezas
    Get #n, , Numheads

    'Resize array
    ReDim HeadData(1 To Numheads) As HeadData
    ReDim MisCabezas(1 To Numheads) As tIndiceCabeza

    For i = 1 To Numheads
        Get #n, , MisCabezas(i)
        
        If MisCabezas(i).Head(1) Then
            InitGrh HeadData(i).Head(1), MisCabezas(i).Head(1), 0
            InitGrh HeadData(i).Head(2), MisCabezas(i).Head(2), 0
            InitGrh HeadData(i).Head(3), MisCabezas(i).Head(3), 0
            InitGrh HeadData(i).Head(4), MisCabezas(i).Head(4), 0
            
            frmMain.ListaHead.AddItem i
        End If
        
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

    Dim LaCabecera As tCabecera

    frmMain.ListaCascos.Clear

    If Not FileExist(DirCliente & "\Scripts\cascos.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Cascos.ind."
        CargarCascos = False
        Exit Function
    End If

    n = FreeFile
    Open DirCliente & "\Scripts\cascos.ind" For Binary Access Read As #n

    'cabecera
    Get #n, , LaCabecera

    'num de cascos
    Get #n, , NumCascos

    'Resize array
    ReDim CascoAnimData(1 To NumCascos) As HeadData
    ReDim MisCabezas(1 To NumCascos) As tIndiceCabeza

    For i = 1 To NumCascos

        Get #n, , MisCabezas(i)
        
        If MisCabezas(i).Head(1) Then
            InitGrh CascoAnimData(i).Head(1), MisCabezas(i).Head(1), 0
            InitGrh CascoAnimData(i).Head(2), MisCabezas(i).Head(2), 0
            InitGrh CascoAnimData(i).Head(3), MisCabezas(i).Head(3), 0
            InitGrh CascoAnimData(i).Head(4), MisCabezas(i).Head(4), 0
            
            frmMain.ListaCascos.AddItem i
        End If
        
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

    Dim LaCabecera As tCabecera
    
    frmMain.ListaEscudos.Clear

    If Not FileExist(DirCliente & "\Scripts\escudos.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Escudos.ind."
        CargarEscudos = False
        Exit Function
    End If

    n = FreeFile
    Open DirCliente & "\Scripts\escudos.ind" For Binary Access Read As #n
        
    'cabecera
    Get #n, , LaCabecera

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
            
            frmMain.ListaEscudos.AddItem i
            
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

    Dim LaCabecera As tCabecera
    
    frmMain.ListaArmas.Clear
    
    If Not FileExist(DirCliente & "\Scripts\armas.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo Armas.ind."
        CargarAnimArmas = False
        Exit Function
    End If
    
    n = FreeFile
    Open DirCliente & "\Scripts\Armas.ind" For Binary Access Read As #n
        
    'cabecera
    Get #n, , MiCabecera
    
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
            
            frmMain.ListaArmas.AddItem i

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

    Dim LaCabecera As tCabecera
    
    frmMain.ListaFxs.Clear
    
    If Not FileExist(DirCliente & "\Scripts\fxs.ind", vbArchive) Then
        MsgBox "No se ha encontrado el archivo FXs.ind."
        CargarFxs = False
        Exit Function
    End If
    
    n = FreeFile
    Open DirCliente & "\Scripts\FXs.ind" For Binary Access Read As #n
        
    'cabecera
    Get #n, , LaCabecera

    'num de cabezas
    Get #n, , NumFxs
        
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
        
    For i = 1 To NumFxs

        Get #n, , FxData(i)
        frmMain.ListaFxs.AddItem i
        
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
