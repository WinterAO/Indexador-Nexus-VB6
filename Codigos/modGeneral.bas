Attribute VB_Name = "modGeneral"
Option Explicit

Private lFrameTimer As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Sub Main()

    'Ruta principal
    IniPath = App.Path & "\"
    
    frmPerfil.Show
    
    Do While ModoElegido = False
        DoEvents
    Loop

    frmCargando.Show
    DoEvents

    If Not CargarConfiguracion Then
        MsgBox "No se ha podido cargar la configuración del Indexador, revisa Config.ini."
        End

    End If
    
    '##############
    ' MOTOR GRAFICO
    
    'Iniciamos el Engine de DirectX 8
    frmCargando.lblstatus.Caption = "Iniciando Motor Grafico..."
    Call mDx8_Engine.Engine_DirectX8_Init
    
    'Tile Engine
    frmCargando.lblstatus.Caption = "Cargando Tile Engine..."
    Call InitTileEngine(frmMain.hwnd, 32, 32, 8, 8)
    
    frmCargando.lblstatus.Caption = "Cargando Graficos"
    DoEvents
    Call LoadGrhData
    
    frmCargando.lblstatus.Caption = "Cargando Personajes"
    DoEvents
    Call CargarCuerpos

    frmCargando.lblstatus.Caption = "Cargando head"
    DoEvents
    Call CargarCabezas

    frmCargando.lblstatus.Caption = "Cargando Helmet"
    DoEvents
    Call CargarCascos

    frmCargando.lblstatus.Caption = "Cargando Escudos"
    DoEvents
    Call CargarEscudos

    frmCargando.lblstatus.Caption = "Cargando Armas"
    DoEvents
    Call CargarAnimArmas

    frmCargando.lblstatus.Caption = "Cargando Fxs"
    DoEvents
    Call CargarFxs
    
    frmCargando.lblstatus.Caption = "Cargando Particulas"
    DoEvents
    Call CargarParticulas
    
    frmCargando.lblstatus.Caption = "Cargando Colores"
    DoEvents
    Call CargarColores
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    Unload frmCargando
    frmMain.Show
    
    'Inicializacion de variables globales
    prgRun = True
    
    lFrameTimer = GetTickCount

    Do While prgRun

        'Solo dibujamos si la ventana no esta minimizada
        If frmMain.WindowState <> vbMinimized And frmMain.Visible Then _
            Call ShowNextFrame
        
        'FPS Counter - mostramos las FPS
        If GetTickCount - lFrameTimer >= 1000 Then _
            lFrameTimer = GetTickCount
        
        DoEvents
    Loop
    
    Call CloseClient

End Sub

''
' Nos dice si existe el archivo/directorio
'
' @param file Especifica el path
' @param FileType Especifica el tipo de archivo/directorio
' @return   Nos devuelve verdadero o falso

Public Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    '*************************************************
    'Author: Unkwown
    'Last modified: 26/05/06
    '*************************************************
    
    On Error GoTo FileExist_Err
    
    If LenB(Dir(File, FileType)) = 0 Then
        FileExist = False
    Else
        FileExist = True

    End If

    Exit Function

FileExist_Err:
    Call LogError(Err.Number, Err.Description, "modMapIO.FileExist", Erl)
    Resume Next
    
End Function

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, File
    
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
End Function

Public Sub LogError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
'**********************************************************
'Author: Jopi
'Guarda una descripcion detallada del error en Errores.log
'**********************************************************
    Dim File As Integer
        File = FreeFile
        
    Open App.Path & "\logs\Errores.log" For Append As #File
    
        Print #File, "Error: " & Numero
        Print #File, "Descripcion: " & Descripcion
        
        If LenB(Linea) <> 0 Then
            Print #File, "Linea: " & Linea
        End If
        
        Print #File, "Componente: " & Componente
        Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
        Print #File, vbNullString
        
    Close #File
    
    Debug.Print "Error: " & Numero & vbNewLine & _
                "Descripcion: " & Descripcion & vbNewLine & _
                "Componente: " & Componente & vbNewLine & _
                "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
                
End Sub

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String
    '*****************************************************************
    'Gets a field from a string
    '*****************************************************************

    Dim i         As Integer

    Dim lastPos   As Integer

    Dim CurChar   As String * 1

    Dim FieldNum  As Integer

    Dim Seperator As String

    Seperator = Chr(SepASCII)
    lastPos = 0
    FieldNum = 0

    For i = 1 To Len(Text)
        CurChar = mid(Text, i, 1)

        If CurChar = Seperator Then
            FieldNum = FieldNum + 1

            If FieldNum = Pos Then
                ReadField = mid(Text, lastPos + 1, (InStr(lastPos + 1, Text, Seperator, vbTextCompare) - 1) - (lastPos))
                Exit Function

            End If

            lastPos = i

        End If

    Next i

    FieldNum = FieldNum + 1

    If FieldNum = Pos Then
        ReadField = mid(Text, lastPos + 1)

    End If

End Function

Function Grh_GetColor(ByVal grh_index As Long) As Long

    On Error Resume Next
    
    Dim x             As Long, y As Long

    Dim file_path     As String

    Dim hdcsrc        As Long, OldObj As Long

    Dim R             As Currency, B As Currency, G As Currency

    Dim InvalidPixels As Long, Size As Long

    Dim TempColor     As tColor

    Dim tempGetPixel  As Long
    
    If grh_index = 0 Then Exit Function
    If GrhData(grh_index).NumFrames > 1 Then grh_index = GrhData(grh_index).Frames(1)
        
    ' file_path = App.Path & "\..\recursos\GRAFICOS\" & GrhData(grh_index).FileNum & ".png"
    
    If Not FileExist(App.Path & "\temp\" & GrhData(grh_index).FileNum & ".jpg", vbNormal) Then
        Call ConvertFileImage(DirCliente & "GRAFICOS\" & GrhData(grh_index).FileNum & ".png", App.Path & "\temp\" & GrhData(grh_index).FileNum & ".jpg", 100)
        file_path = App.Path & "\temp\" & GrhData(grh_index).FileNum & ".jpg"
    Else
        'Debug.Print "existia"
        file_path = App.Path & "\temp\" & GrhData(grh_index).FileNum & ".jpg"

    End If

    'Debug.Print file_path
    
    If FileExist(file_path, vbNormal) Then
        hdcsrc = CreateCompatibleDC(frmMinimapa.Picture1.hDC)
        OldObj = SelectObject(hdcsrc, LoadPicture(file_path))
        BitBlt frmMinimapa.Picture1.hDC, 0, 0, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, hdcsrc, GrhData(grh_index).sX, GrhData(grh_index).sY, vbSrcCopy
        DeleteObject SelectObject(hdcsrc, OldObj)
        DeleteDC hdcsrc
        
        DoEvents
               
        For x = 1 To GrhData(grh_index).pixelWidth
            For y = 1 To GrhData(grh_index).pixelHeight
                tempGetPixel = GetPixel(frmMinimapa.Picture1.hDC, x, y)

                If tempGetPixel = vbBlack Then
                    InvalidPixels = InvalidPixels + 1
                Else
                    TempColor = Long2RGB(tempGetPixel)
                    R = R + TempColor.R
                    G = G + TempColor.G
                    B = B + TempColor.B

                End If

                DoEvents
            Next y
        Next x
        
        If InvalidPixels > 0 Then
            Size = GrhData(grh_index).pixelWidth * GrhData(grh_index).pixelHeight - InvalidPixels
        Else
            Size = GrhData(grh_index).pixelWidth * GrhData(grh_index).pixelHeight

        End If
        
        If Size = 0 Then Size = 1
        
        Grh_GetColor = RGB(CByte(R / Size), CByte(G / Size), CByte(B / Size))
        frmMinimapa.Picture2.BackColor = Grh_GetColor

        Dim bmpguardado As Integer

        'Debug.Print GrhData(grh_index).FileNum

        If GrhData(grh_index + 1).FileNum <> GrhData(grh_index).FileNum Then
            'Debug.Print GrhData(grh_index).FileNum
            Kill App.Path & "\temp\" & GrhData(grh_index).FileNum & ".jpg"

        End If

    End If

End Function

Private Function Long2RGB(ByVal Color As Long) As tColor
    Long2RGB.R = Color And &HFF
    Long2RGB.G = (Color And &HFF00&) \ &H100&
    Long2RGB.B = (Color And &HFF0000) \ &H10000
End Function

Public Sub CloseClient()
    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 8/14/2007
    'Frees all used resources, cleans up and leaves
    '**************************************************************
    
    EngineRun = False
    
    'Stop tile engine
    Call Engine_DirectX8_End

    Set SurfaceDB = Nothing
    Erase MapData
    
    Call UnloadAllForms
    
    End
    
End Sub

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Function Buscar_Carpeta(Optional Titulo As String, _
                        Optional Path_Inicial As Variant) As String
                        
'******************************************************************
' Funcción que abre el cuadro de dialogo y retorna la ruta
'******************************************************************
  
On Local Error GoTo errFunction
      
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
      
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
      
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder( _
                            0, _
                            Titulo, _
                            0, _
                            Path_Inicial)
      
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
      
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path
  
Exit Function
'Error
errFunction:
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString
    Call RegistrarError(Err.Number, Err.Description, "Buscar_Carpeta", Erl)
  
End Function

