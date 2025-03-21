Attribute VB_Name = "modGeneral"
Option Explicit

Private lFrameTimer As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

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
    If Not CargarCuerpos Then End

    frmCargando.lblstatus.Caption = "Cargando head"
    DoEvents
    If Not CargarCabezas Then End

    frmCargando.lblstatus.Caption = "Cargando Helmet"
    DoEvents
    If Not CargarCascos Then End

    frmCargando.lblstatus.Caption = "Cargando Escudos"
    DoEvents
    If Not CargarEscudos Then End

    frmCargando.lblstatus.Caption = "Cargando Armas"
    DoEvents
    If Not CargarAnimArmas Then End

    frmCargando.lblstatus.Caption = "Cargando Fxs"
    DoEvents
    If Not CargarFxs Then End
    
    frmCargando.lblstatus.Caption = "Cargando Particulas"
    DoEvents
    Call CargarParticulas
    
    frmCargando.lblstatus.Caption = "Cargando Colores"
    DoEvents
    Call CargarColores
    
    frmCargando.lblstatus.Caption = "Cargando Indices"
    DoEvents
    Call CargarIndices
    
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
    Call RegistrarError(Err.Number, Err.Description, "modMapIO.FileExist", Erl)
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
    
    Dim X             As Long, Y As Long

    Dim file_path     As String

    Dim hdcsrc        As Long, OldObj As Long

    Dim r             As Currency, b As Currency, g As Currency

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
               
        For X = 1 To GrhData(grh_index).pixelWidth
            For Y = 1 To GrhData(grh_index).pixelHeight
                tempGetPixel = GetPixel(frmMinimapa.Picture1.hDC, X, Y)

                If tempGetPixel = vbBlack Then
                    InvalidPixels = InvalidPixels + 1
                Else
                    TempColor = Long2RGB(tempGetPixel)
                    r = r + TempColor.r
                    g = g + TempColor.g
                    b = b + TempColor.b

                End If

                DoEvents
            Next Y
        Next X
        
        If InvalidPixels > 0 Then
            Size = GrhData(grh_index).pixelWidth * GrhData(grh_index).pixelHeight - InvalidPixels
        Else
            Size = GrhData(grh_index).pixelWidth * GrhData(grh_index).pixelHeight

        End If
        
        If Size = 0 Then Size = 1
        
        Grh_GetColor = RGB(CByte(r / Size), CByte(g / Size), CByte(b / Size))
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
    Long2RGB.r = Color And &HFF
    Long2RGB.g = (Color And &HFF00&) \ &H100&
    Long2RGB.b = (Color And &HFF0000) \ &H10000
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

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                    ByVal Text As String, _
                    Optional ByVal Red As Integer = -1, _
                    Optional ByVal Green As Integer, _
                    Optional ByVal Blue As Integer, _
                    Optional ByVal bold As Boolean = False, _
                    Optional ByVal italic As Boolean = False, _
                    Optional ByVal bCrLf As Boolean = True, _
                    Optional ByVal Alignment As Byte = rtfLeft, _
                    Optional ByVal bFecha As Boolean = True)
    
'****************************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D apperance!
'****************************************************
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Martin Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'Jopi 17/08/2019 : Consola transparente.
'Jopi 17/08/2019 : Ahora podes especificar el alineamiento del texto.
'Lorwik 20/03/2024: Ahora puedes mostrar la hora en la que se imprimio el mensaje
'****************************************************

    Dim horaActual As String
    Dim hora As Integer
    Dim minutos As Integer
    
    ' Obtener la hora actual en formato de cadena de caracteres
    horaActual = Time
    
    ' Extraer la hora y los minutos
    hora = Hour(horaActual)
    minutos = Minute(horaActual)

    With RichTextBox
    
        If bFecha Then _
            Text = hora & ":" & minutos & "> " & Text
        
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        ' 0 = Left
        ' 1 = Center
        ' 2 = Right
        .SelAlignment = Alignment

        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        
        .SelText = Text

        ' Esto arregla el bug de las letras superponiendose la consola del frmMain
        If Not RichTextBox = frmMain.RichConsola Then RichTextBox.Refresh

    End With
End Sub

Public Sub recargarLynxGrh()

    On Error GoTo ErrorHandler
    
    Dim K   As Long

    Dim Grh As Long

    With frmMain
        
        .LynxGrh.Clear
        .LynxGrh.Redraw = False
        .LynxGrh.Visible = False
        .LynxGrh.AddColumn "Grh", 0
        .LynxGrh.AddColumn "Tipo", 0
        
    End With

    For Grh = 1 To grhCount
        
        frmMain.LynxGrh.AddItem Grh
        K = frmMain.LynxGrh.Rows - 1
        frmMain.LynxGrh.CellText(K, 0) = Grh
    
        With GrhData(Grh)
               
            If .NumFrames > 1 Then

                frmMain.LynxGrh.CellText(K, 1) = "ANIMACION"

            Else
                
                frmMain.LynxGrh.CellText(K, 1) = ""
                    
            End If

        End With
        
    Next Grh
    
    frmMain.LynxGrh.Visible = True
    frmMain.LynxGrh.Redraw = True
    frmMain.LynxGrh.ColForceFit
    
    DoEvents
    
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & " durante la recargaLynxGrh."
    
    frmMain.LynxGrh.Visible = True
    frmMain.LynxGrh.Redraw = True
    frmMain.LynxGrh.ColForceFit

End Sub
