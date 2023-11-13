Attribute VB_Name = "mod_TileEngine"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' No matter what you do with DirectX8, you will need to start with
' the DirectX8 object. You will need to create a new instance of
' the object, using the New keyword, rather than just getting a
' pointer to it, since there's nowhere to get a pointer from yet (duh!).

Public DirectX               As New DirectX8

' The D3DX8 object contains lots of helper functions, mostly math
' to make Direct3D alot easier to use. Notice we create a new
' instance of the object using the New keyword.
Public DirectD3D8            As D3DX8

Public DirectD3D             As Direct3D8

' The Direct3DDevice8 represents our rendering device, which could
' be a hardware or a software device. The great thing is we still
' use the same object no matter what it is
Public DirectDevice          As Direct3DDevice8

' The D3DDISPLAYMODE type structure that holds
' the information about your current display adapter.
Public DispMode              As D3DDISPLAYMODE
    
' The D3DPRESENT_PARAMETERS type holds a description of the way
' in which DirectX will display it's rendering.
Public D3DWindow             As D3DPRESENT_PARAMETERS

Public SurfaceDB             As New clsTextureManager

Public SpriteBatch           As New clsBatch

Private Viewport             As D3DVIEWPORT8

Private Projection           As D3DMATRIX

Private View                 As D3DMATRIX

Public Engine_BaseSpeed      As Single

Public EngineRun             As Boolean

Public FPS                   As Long

Public FramesPerSecCounter   As Long

Public FPSLastCheck          As Long

Public Normal_RGBList(3)     As Long

Public Const DegreeToRadian  As Single = 0.01745329251994 'Pi / 180

'Tamano del la vista en Tiles
Private WindowTileWidth      As Integer

Private WindowTileHeight     As Integer

Public HalfWindowTileWidth   As Integer

Public HalfWindowTileHeight  As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer

Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime      As Single

Public timerTicksPerFrame    As Single

'Tamano de los tiles en pixels
Public TilePixelHeight       As Integer

Public TilePixelWidth        As Integer

'Posicion en un mapa
Public Type Position

    X As Long
    Y As Long

End Type

'Contiene info acerca de donde se puede encontrar un grh tamano y animacion
Public Type GrhData

    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    speed As Single
    
    mini_map_color As Long
    active As Boolean

End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh

    GrhIndex As Long
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    angle As Single

End Type

Public CurrentGrh       As Grh

Public GrhSelect(3)     As Long

'?????????Graficos???????????
Public GrhData()        As GrhData 'Guarda todos los grh

Public BodyData()       As BodyData

Public HeadData()       As HeadData

Public FxData()         As tIndiceFx

Public WeaponAnimData() As WeaponAnimData

Public ShieldAnimData() As ShieldAnimData

Public CascoAnimData()  As HeadData
'?????????????????????????

Type SupData
    name As String
    Grh As Integer
    Width As Byte
    Height As Byte
    Block As Boolean
    Capa As Byte
End Type

Public MaxSup As Integer

Public SupData() As SupData

'Tipo de las celdas del mapa
Public Type MapBlock

    Particle_Group_Index As Integer

End Type

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?
Public MapData()         As MapBlock
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100

Public Const XMinMapSize As Byte = 1

Public Const YMaxMapSize As Byte = 100

Public Const YMinMapSize As Byte = 1

'Lista de cuerpos
Public Type BodyData

    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position

End Type

'Lista de cabezas
Public Type HeadData

    Head(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de las armas
Type WeaponAnimData

    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData

    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency _
                Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter _
                Lib "kernel32" (lpPerformanceCount As Currency) As Long

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Checks to see if a tile position is in the maps bounds
    '*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function

    End If
    
    InMapBounds = True

End Function

Public Sub InitGrh(ByRef Grh As Grh, _
                   ByVal GrhIndex As Long, _
                   Optional ByVal Started As Byte = 2)
    '*****************************************************************
    'Sets up a grh. MUST be done before rendering
    '*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0

        End If

    Else

        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started

    End If
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0

    End If
    
    Grh.FrameCounter = 1
    Grh.speed = GrhData(Grh.GrhIndex).speed

End Sub

Public Sub InitTileEngine(ByVal setDisplayFormhWnd As Long, _
                          ByVal setTilePixelHeight As Integer, _
                          ByVal setTilePixelWidth As Integer, _
                          ByVal pixelsToScrollPerFrameX As Integer, _
                          pixelsToScrollPerFrameY As Integer)
    '***************************************************
    'Author: Aaron Perkins
    'Last Modification: 08/14/07
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Configures the engine to start running.
    '***************************************************

    On Error GoTo ErrorHandler:

    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    
    WindowTileHeight = Round(frmMain.MainViewPic.Height / 32, 0)
    WindowTileWidth = Round(frmMain.MainViewPic.Width / 32, 0)

    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY

    On Error GoTo 0
    
    Call LoadGraphics

    Exit Sub
    
ErrorHandler:

    Call LogError(Err.Number, Err.Description, "Mod_TileEngine.InitTileEngine")
    
    Call CloseClient
    
End Sub

Public Sub LoadGraphics()
    Call SurfaceDB.Initialize(DirectD3D8, ClientSetup.byMemory)

End Sub

Public Sub Engine_Update_FPS()
    '***************************************************
    'Author: ???
    'Last Modification: ????
    'Calculate $ Limitate (if active) FPS.
    '***************************************************

    If ClientSetup.LimiteFPS Then
        While (GetTickCount - FPSLastCheck) \ 10 < FramesPerSecCounter
            Call Sleep(5)
        Wend
    End If

    If FPSLastCheck + 1000 < timeGetTime Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 1
        FPSLastCheck = timeGetTime
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1

    End If

End Sub

Sub ShowNextFrame()

    On Error GoTo ErrorHandler:

    If EngineRun Then
        
        Call Engine_BeginScene

        Call Draw_Grh(CurrentGrh, 200, 250, 1, Normal_RGBList(), True)

        If frmParticleEditor.Visible Then Call RenderParticulas(50, 50)

        Call Engine_Update_FPS
        
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * Engine_BaseSpeed
            
        Call Engine_EndScene(MainScreenRect, 0)
        
    End If
    
ErrorHandler:

    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Call mDx8_Engine.Engine_DirectX8_Init
        Call LoadGraphics

    End If
  
End Sub

Public Sub RenderParticulas(ByVal tilex As Integer, ByVal tiley As Integer)

    Dim Y                 As Integer     'Keeps track of where on map we are

    Dim X                 As Integer     'Keeps track of where on map we are

    Dim screenminY        As Integer  'Start Y pos on current screen

    Dim screenmaxY        As Integer  'End Y pos on current screen

    Dim screenminX        As Integer  'Start X pos on current screen

    Dim screenmaxX        As Integer  'End X pos on current screen

    Dim ScreenX           As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY           As Integer  'Keeps track of where to place tile on screen

    Static OffsetCounterX As Single

    Static OffsetCounterY As Single

    Dim PixelOffsetX      As Integer, PixelOffsetY As Integer
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    PixelOffsetX = OffsetCounterX
    PixelOffsetY = OffsetCounterY

    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX

            With MapData(X, Y)

                '***********************************************
                If .Particle_Group_Index > 0 Then Particle_Group_Render .Particle_Group_Index, 200, 100
                
                If ReferenciaPJ Then

                    Dim PJ As Grh

                    Call InitGrh(PJ, 4581)
                    'Call Grh_Render(PJ, 70, 70, Normal_RGBList())
                
                End If
                
            End With

            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y

End Sub

Public Function GetElapsedTime() As Single

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim Start_Time    As Currency

    Static end_time   As Currency

    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        Call QueryPerformanceFrequency(timer_freq)

    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)

End Function

Public Sub GrhUninitialize(Grh As Grh)
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    'Resets a Grh
    '*****************************************************************

    With Grh
        
        'Copy of parameters
        .GrhIndex = 0
        .Started = False
        .Loops = 0
        
        'Set frame counters
        .FrameCounter = 0
        .speed = 0
                
    End With

End Sub

Sub Draw_GrhIndex(ByVal GrhIndex As Long, _
                  ByVal X As Integer, _
                  ByVal Y As Integer, _
                  ByVal Center As Byte, _
                  ByRef Color_List() As Long, _
                  Optional ByVal angle As Single = 0, _
                  Optional ByVal Alpha As Boolean = False)

    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth - TilePixelWidth) \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If
        
        'Draw
        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha)

    End With
    
End Sub

Sub Draw_Grh(ByRef Grh As Grh, _
             ByVal X As Integer, _
             ByVal Y As Integer, _
             ByVal Center As Byte, _
             ByRef Color_List() As Long, _
             ByVal Animate As Byte, _
             Optional ByVal Alpha As Boolean = False, _
             Optional ByVal angle As Single = 0, _
             Optional ByVal ScaleX As Single = 1!, _
             Optional ByVal ScaleY As Single = 1!)

    '*****************************************************************
    'Draws a GRH transparently to a X and Y position
    '*****************************************************************
    Dim CurrentGrhIndex As Long
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
    On Error GoTo Error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.speed) * Movement_Speed

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0

                    End If

                End If

            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)

        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth * ScaleX - TilePixelWidth) \ 2

            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight

            End If

        End If

        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha, angle, ScaleX, ScaleY)
        
    End With
    
    Exit Sub

Error:

    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        #If Desarrollo = 0 Then
            Call LogError(Err.Number, "Error in Draw_Grh, " & Err.Description, "Draw_Grh", Erl)
            MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
            Call CloseClient
        
        #Else
            Debug.Print "Error en Draw_Grh en el grh" & CurrentGrhIndex & ", " & Err.Description & ", (" & Err.Number & ")"
        #End If

    End If

End Sub

Public Sub Device_Textured_Render(ByVal X As Single, _
                                  ByVal Y As Single, _
                                  ByVal Width As Integer, _
                                  ByVal Height As Integer, _
                                  ByVal sX As Integer, _
                                  ByVal sY As Integer, _
                                  ByVal tex As Long, _
                                  ByRef Color() As Long, _
                                  Optional ByVal Alpha As Boolean = False, _
                                  Optional ByVal angle As Single = 0, _
                                  Optional ByVal ScaleX As Single = 1!, _
                                  Optional ByVal ScaleY As Single = 1!)

    Dim Texture      As Direct3DTexture8
        
    Dim TextureWidth As Long, TextureHeight As Long

    Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
        
    With SpriteBatch

        Call .SetTexture(Texture)
                    
        Call .SetAlpha(Alpha)
                
        If TextureWidth <> 0 And TextureHeight <> 0 Then
            Call .Draw(X, Y, Width * ScaleX, Height * ScaleY, Color, sX / TextureWidth, sY / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, angle)
        Else
            Call .Draw(X, Y, TextureWidth * ScaleX, TextureHeight * ScaleY, Color, , , , , angle)

        End If
                
    End With
        
End Sub

