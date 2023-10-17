Attribute VB_Name = "mDx8_Engine"
Option Explicit

Private Projection As D3DMATRIX
Private View As D3DMATRIX

Public MainScreenRect As RECT

Public ScreenWidth As Long
Public ScreenHeight As Long

Private EndTime As Long

Public Sub Engine_DirectX8_Init()
    On Error GoTo EngineHandler:

    ' Initialize all DirectX objects.
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
    
    Dim ERROR_DIRECTX_INIT As String
    ERROR_DIRECTX_INIT = "No se pudo inicializar el motor grafico. Por favor, verifique si tiene sus librerias y sus controladores actualizados."
    
    If ClientSetup.OverrideVertexProcess > 0 Then
        
        Select Case ClientSetup.OverrideVertexProcess
            
            Case 1:
                If Not Engine_Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
                    Call MsgBox(ERROR_DIRECTX_INIT)
                    End
                End If
            
            
            Case 2:
                If Not Engine_Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                    Call MsgBox(ERROR_DIRECTX_INIT)
                    End
                End If

            
            Case 3:
                If Not Engine_Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                    Call MsgBox(ERROR_DIRECTX_INIT)
                    End
                End If
        End Select
        
    Else
        'Detectamos el modo de renderizado mas compatible con tu PC.
        If Not Engine_Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not Engine_Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not Engine_Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
            
                    Call MsgBox(ERROR_DIRECTX_INIT)
                    End
                
                End If
            End If
        End If
    End If

    'Seteamos la matriz de proyeccion.
    Call D3DXMatrixOrthoOffCenterLH(Projection, 0, ScreenWidth, ScreenHeight, 0, -1#, 1#)
    Call D3DXMatrixIdentity(View)
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
    Call DirectDevice.SetTransform(D3DTS_VIEW, View)

    ' Set rendering options
    Call Engine_Init_RenderStates
    
    'Carga dinamica de texturas por defecto.
    Set SurfaceDB = New clsTextureManager
    
    'Sprite batching.
    Set SpriteBatch = New clsBatch
    Call SpriteBatch.Initialise(2000)
    
    'Inicializamos el resto de sistemas.
    frmCargando.lblstatus.Caption = "Cargando Recursos adicionales de DX8..."
    Call Engine_DirectX8_Aditional_Init
    
    EndTime = timeGetTime
    
    Exit Sub
EngineHandler:
    
    Call LogError(Err.Number, Err.Description, "mDx8_Engine.Engine_DirectX8_Init")
    
    Call CloseClient
End Sub

Private Function Engine_Init_DirectDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean

    'Establecemos cual va a ser el tamano del render.
    ScreenWidth = 6400
    ScreenHeight = 6400
    
    ' Retrieve the information about your current display adapter.
    Call DirectD3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    
        ' Fill the D3DPRESENT_PARAMETERS type, describing how DirectX should
    ' display it's renders.
    With D3DWindow
        .Windowed = True
        
        ' The swap effect determines how the graphics get from the backbuffer to the screen.
        ' D3DSWAPEFFECT_DISCARD:
        '   Means that every time the render is presented, the backbuffer
        '   image is destroyed, so everything must be rendered again.
        .SwapEffect = D3DSWAPEFFECT_DISCARD
        
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = ScreenWidth
        .BackBufferHeight = ScreenHeight
        .hDeviceWindow = frmMain.MainViewPic.hwnd
    End With
    
    If Not DirectDevice Is Nothing Then
        Set DirectDevice = Nothing
    End If
    
    ' Create the rendering device.
    ' Here we request a Hardware or Mixed rasterization.
    ' If your computer does not have this, the request may fail, so use
    ' D3DDEVTYPE_REF instead of D3DDEVTYPE_HAL if this happens. A real
    ' program would be able to detect an error and automatically switch device.
    ' We also request software vertex processing, which means the CPU has to
    Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DWindow.hDeviceWindow, D3DCREATEFLAGS, D3DWindow)
    
    'Lo pongo xq es bueno saberlo...
    Select Case D3DCREATEFLAGS
    
        Case D3DCREATE_MIXED_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: MIXED"
        
        Case D3DCREATE_HARDWARE_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: HARDWARE"
            
        Case D3DCREATE_SOFTWARE_VERTEXPROCESSING
            Debug.Print "Modo de Renderizado: SOFTWARE"
            
    End Select
    
    'Everything was successful
    Engine_Init_DirectDevice = True
    
    Exit Function
    
ErrorDevice:
    
    'Destroy the D3DDevice so it can be remade
    Set DirectDevice = Nothing

    'Return a failure
    Engine_Init_DirectDevice = False
    
End Function

Private Sub Engine_Init_RenderStates()

    'Set the render states
    With DirectDevice
    
        Call .SetVertexShader(D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
        Call .SetRenderState(D3DRS_LIGHTING, False)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)
        Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        Call .SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
        
    End With
    
End Sub

Public Sub Engine_DirectX8_End()
'***************************************************
'Author: Standelf
'Last Modification: 26/05/2010
'Destroys all DX objects
'***************************************************
On Error Resume Next
    Dim i As Byte
    
    '   Clean Particles
    'Call Particle_Group_Remove_All
    
    '   Clean Texture
    Call DirectDevice.SetTexture(0, Nothing)
    
    Set DirectD3D8 = Nothing
    Set DirectD3D = Nothing
    Set DirectX = Nothing
    Set DirectDevice = Nothing
    Set SpriteBatch = Nothing
End Sub

Public Sub Engine_DirectX8_Aditional_Init()
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/12/2010
'**************************************************************

    FPS = 101
    FramesPerSecCounter = 101
    
    Engine_BaseSpeed = 0.018
    
    With MainScreenRect
        .Bottom = frmMain.MainViewPic.ScaleHeight
        .Right = frmMain.MainViewPic.ScaleWidth
    End With
    
    If Not prgRun Then
        
        ' Seteamos algunos colores por adelantado y unica vez.
        Call Engine_Long_To_RGB_List(Normal_RGBList(), -1)
        
    End If
    
End Sub

Public Sub Engine_BeginScene(Optional ByVal Color As Long = 0)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | DD Clear & BeginScene
'***************************************************

    Call DirectDevice.BeginScene
    Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, Color, 1#, 0)
    Call SpriteBatch.Begin
    
End Sub

Public Sub Engine_EndScene(ByRef destRect As RECT, Optional ByVal hWndDest As Long = 0)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | DD EndScene & Present
'***************************************************
On Error GoTo DeviceHandler:

    Call SpriteBatch.Flush
    Call DirectDevice.EndScene
        
    If hWndDest = 0 Then
        Call DirectDevice.Present(destRect, ByVal 0&, ByVal 0&, ByVal 0&)
    
    Else
        Call DirectDevice.Present(destRect, ByVal 0, hWndDest, ByVal 0)
    
    End If
    
    Exit Sub
    
DeviceHandler:

    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Call mDx8_Engine.Engine_DirectX8_Init
        Call LoadGraphics
    End If
    
End Sub

Public Sub Engine_Long_To_RGB_List(rgb_list() As Long, long_color As Long)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 16/05/10
'Blisse-AO | Set a Long Color to a RGB List
'***************************************************
    rgb_list(0) = long_color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub
