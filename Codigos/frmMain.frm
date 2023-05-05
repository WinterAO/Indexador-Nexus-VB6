VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Indexador NexusAO"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   11610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   471
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   774
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox Archivos 
      Height          =   285
      Left            =   4110
      Pattern         =   "*.bmp"
      TabIndex        =   16
      Top             =   30
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox GRHt 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   450
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   6540
      Width           =   11535
   End
   Begin VB.ListBox ListaArmas 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   9480
      TabIndex        =   13
      Top             =   360
      Width           =   2085
   End
   Begin VB.ListBox ListaFxs 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   9480
      TabIndex        =   11
      Top             =   4590
      Width           =   2115
   End
   Begin VB.ListBox ListaEscudos 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   9480
      TabIndex        =   9
      Top             =   2460
      Width           =   2085
   End
   Begin VB.ListBox ListaCascos 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   7350
      TabIndex        =   7
      Top             =   4590
      Width           =   2085
   End
   Begin VB.ListBox ListaHead 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   7350
      TabIndex        =   5
      Top             =   2460
      Width           =   2085
   End
   Begin VB.PictureBox MainViewPic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   6105
      Left            =   2130
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   5205
   End
   Begin VB.ListBox ListaCuerpos 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   7380
      TabIndex        =   1
      Top             =   360
      Width           =   2085
   End
   Begin VB.ListBox Listado 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   6105
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   2025
   End
   Begin VB.Label lblArmas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Armas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   10050
      TabIndex        =   14
      Top             =   120
      Width           =   675
   End
   Begin VB.Label lblFxS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fx's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   10110
      TabIndex        =   12
      Top             =   4350
      Width           =   675
   End
   Begin VB.Label lblEscudos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escudos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   10080
      TabIndex        =   10
      Top             =   2220
      Width           =   675
   End
   Begin VB.Label lblCascos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cascos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   8040
      TabIndex        =   8
      Top             =   4350
      Width           =   675
   End
   Begin VB.Label lblHead 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cabezas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   7950
      TabIndex        =   6
      Top             =   2220
      Width           =   810
   End
   Begin VB.Label lbCuerpos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cuerpos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   7950
      TabIndex        =   3
      Top             =   120
      Width           =   795
   End
   Begin VB.Label lblGraficos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Graficos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   480
      TabIndex        =   2
      Top             =   150
      Width           =   780
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuExportar 
         Caption         =   "Exportar..."
         Begin VB.Menu mnuExportarTodo 
            Caption         =   "TODO"
         End
         Begin VB.Menu mnuExportarGraficos 
            Caption         =   "Graficos.ind"
         End
         Begin VB.Menu mnuExportarCabezas 
            Caption         =   "Cascos.ind"
         End
         Begin VB.Menu mnuExportarCuerpos 
            Caption         =   "Personajes.ind"
         End
         Begin VB.Menu mnuExportarCascos 
            Caption         =   "Cascos.ind"
         End
         Begin VB.Menu mnuExportarFxs 
            Caption         =   "FXs.ind"
         End
      End
      Begin VB.Menu mnuIndexar 
         Caption         =   "Indexar..."
         Begin VB.Menu mnuIndexarTodo 
            Caption         =   "&TODO"
         End
         Begin VB.Menu mnuIndexarGraficos 
            Caption         =   "Graficos.ind"
         End
         Begin VB.Menu mnuIndexarCabezas 
            Caption         =   "Cabezas.ind"
         End
         Begin VB.Menu mnuIndexarPersonajes 
            Caption         =   "Personajes.ind"
         End
         Begin VB.Menu mnuIndexarCascos 
            Caption         =   "Cascos.ind"
         End
         Begin VB.Menu mnuIndexarArmas 
            Caption         =   "Armas.ind"
         End
         Begin VB.Menu mnuIndexarEscudos 
            Caption         =   "Escudos.ind"
         End
         Begin VB.Menu mnuIndexarFXs 
            Caption         =   "FXs.ind"
         End
      End
      Begin VB.Menu mnuGenerarMinimapa 
         Caption         =   "Generar Minimapa"
      End
   End
   Begin VB.Menu mnuira 
      Caption         =   "&Ir a..."
      Begin VB.Menu mnuCarpetaClienteIr 
         Caption         =   "&Carpeta del cliente"
      End
      Begin VB.Menu mnuCarpetaExportacionIr 
         Caption         =   "&Carpeta de exportación"
      End
      Begin VB.Menu mnuCarpetaIndexacionIr 
         Caption         =   "&Carpeta de indexación"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuBuscarGrh 
         Caption         =   "&Buscar Grh"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuscarGrhconPNG 
         Caption         =   "&Buscar Grh con PNG"
      End
      Begin VB.Menu mnuIrASBMP 
         Caption         =   "&Buscar Siguiente PNG"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuscarDuplicados 
         Caption         =   "&Buscar Grh duplicados"
      End
      Begin VB.Menu mnuIndexPNG 
         Caption         =   "&Buscar errores de indexación"
      End
      Begin VB.Menu mnuPNGinutiles 
         Caption         =   "&Buscar PNG inutilizados"
      End
      Begin VB.Menu mnuBuscarGrhLibresConsecutivos 
         Caption         =   "&Buscar Grh Libres Consecutivos"
      End
      Begin VB.Menu mnuBuscarErrDim 
         Caption         =   "Buscar Errores de Dimenciónes"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BuscarPNG As Integer

Private Sub Form_Load()
    EngineRun = True
End Sub

Private Sub Listado_Click()
    DoEvents
    Call InitGrh(CurrentGrh, ReadField(1, Listado.Text, Asc(" ")))
End Sub

Private Sub mnuBuscarDuplicados_Click()
    Dim i     As Long
    Dim j     As Long
    Dim K     As Integer
    Dim Datos As String
    Dim DatX  As Byte
    Dim Tim   As Byte

    Me.Hide
    frmCargando.Show
    Tim = 0

    For i = 1 To grhCount

        If GrhData(i).NumFrames >= 1 Then

            For j = (i + 1) To grhCount
                Tim = Tim + 1

                If Tim >= 250 Then
                    Tim = 0
                    frmCargando.lblstatus.Caption = "Buscando duplicados " & i & " Grh"
                    DoEvents

                End If

                If GrhData(i).NumFrames = 1 Then
                    If GrhData(j).FileNum = GrhData(i).FileNum Then
                        If (GrhData(i).sX & GrhData(i).sY & GrhData(i).pixelHeight & GrhData(i).pixelWidth) = (GrhData(j).sX & GrhData(j).sY & GrhData(j).pixelHeight & GrhData(j).pixelWidth) Then
                            Datos = Datos & "Grh" & i & " esta duplicado con Grh" & j & vbCrLf

                        End If

                    End If

                Else

                    If (GrhData(i).NumFrames = GrhData(j).NumFrames) And (GrhData(i).speed = GrhData(j).speed) Then
                        DatX = 0

                        For K = 1 To GrhData(j).NumFrames

                            If GrhData(i).Frames(K) = GrhData(j).Frames(K) Then
                                DatX = DatX + 1

                            End If

                        Next

                        If DatX = GrhData(j).NumFrames Then
                            Datos = Datos & "Grh" & i & " (ANIMACION) esta duplicado con Grh" & j & " (ANIMACION)" & vbCrLf

                        End If

                    End If

                End If

            Next

        End If

    Next
    
    Unload frmCargando
    Me.Show
    frmCodigo.Caption = mnuBuscarDuplicados.Caption
    frmCodigo.Codigo.Text = Datos
    frmCodigo.Show

End Sub

Private Sub mnuIndexPNG_Click()

    Dim Datos As String
    Dim i     As Long
    Dim j     As Integer
    Dim Tim   As Byte

    Me.Hide
    frmCargando.Show
    Tim = 0

    For i = 1 To grhCount

        If GrhData(i).NumFrames > 1 Then
            Tim = Tim + 1

            If Tim >= 150 Then
                Tim = 0
                frmCargando.lblstatus.Caption = "Procesando " & i & " Grh"
                DoEvents

            End If

            For j = 1 To GrhData(i).NumFrames

                If GrhData(GrhData(i).Frames(j)).FileNum = 0 Then
                    Datos = Datos & "Grh" & i & " (ANIMACION) en Frame " & j & " - Le falta el GRH " & GrhData(i).Frames(j) & vbCrLf
                ElseIf LenB(Dir(DirCliente & "\Graficos\" & GrhData(GrhData(i).Frames(j)).FileNum & ".png", vbArchive)) = 0 Then
                    Datos = Datos & "Grh" & i & " (ANIMACION) en Frame " & j & " - Le falta el png " & GrhData(GrhData(i).Frames(j)).FileNum & " (GRH" & GrhData(i).Frames(j) & ")" & vbCrLf

                End If

            Next
        ElseIf GrhData(i).NumFrames = 1 Then
            Tim = Tim + 1

            If Tim >= 150 Then
                Tim = 0
                frmCargando.lblstatus.Caption = "Procesando " & i & " grh"
                DoEvents

            End If

            If LenB(Dir(DirCliente & "\Graficos\" & GrhData(i).FileNum & ".png", vbArchive)) = 0 Then
                Datos = Datos & "Grh" & i & " - Le falta el png " & GrhData(i).FileNum & vbCrLf

            End If

        End If

    Next
    Unload frmCargando
    Me.Show
    frmCodigo.Caption = mnuIndexPNG.Caption
    frmCodigo.Codigo.Text = Datos
    frmCodigo.Show

End Sub

Private Sub mnuPNGinutiles_Click()

    Dim i        As Long
    Dim j        As Long
    Dim Datos    As String
    Dim Encontre As Boolean
    Dim NumPNG   As String

    Dim Tim      As Byte

    Archivos.Path = DirCliente & "\Graficos\"
    DoEvents
    Me.Hide
    frmCargando.Show
    Tim = 0

    For i = 0 To Archivos.ListCount
        Encontre = False
        NumPNG = ReadField(1, Archivos.List(i), Asc("."))
        'Tim = Tim + 1
        'If Tim >= 2 Then
        '    Tim = 0
        frmCargando.lblstatus.Caption = "Buscando PNGs inutiles " & NumPNG & " PNG"
        DoEvents

        'End If
        For j = 1 To grhCount

            If IsNumeric(NumPNG) = False Then
                Encontre = True
                Exit For

            End If

            If GrhData(j).NumFrames = 1 Then
                If GrhData(j).FileNum = NumPNG Then
                    Encontre = True
                    Exit For

                End If

            End If

        Next

        If Encontre = False Then
            Datos = Datos & "El PNG " & NumPNG & " se encuentra inutilizado" & vbCrLf

        End If

    Next
    Unload frmCargando
    Me.Show
    frmCodigo.Caption = mnuPNGinutiles.Caption
    frmCodigo.Codigo.Text = Datos
    frmCodigo.Show

End Sub

Private Sub mnuBuscarGrhLibresConsecutivos_Click()

    On Error Resume Next

    Dim libres As Long
    Dim i      As Long
    Dim Conta  As Long

    libres = InputBox("Grh Libres Consecutivos")

    If IsNumeric(libres) = False Then Exit Sub

    For i = 1 To grhCount

        If GrhData(i).NumFrames = 0 Then
            Conta = Conta + 1

            If Conta = libres Then
                MsgBox "Desde Grh" & i - (Conta - 1) & " hasta Grh" & i & " se encuentran libres."
                Exit Sub

            End If

        ElseIf Conta > 0 Then
            Conta = 0

        End If

    Next
    MsgBox "No se encontraron " & libres & " GRH Libres Consecutivos"

End Sub

Private Sub mnuBuscarErrDim_Click()

    Dim i       As Long
    Dim j       As Long
    Dim Datos   As String
    Dim Tim     As Byte
    Dim tipo(1) As Integer

    Me.Hide
    frmCargando.Show
    Tim = 0

    For i = 1 To grhCount

        If GrhData(i).NumFrames > 1 Then
            Tim = Tim + 1

            If Tim >= 150 Then
                Tim = 0
                frmCargando.lblstatus.Caption = "Procesando " & i & " grh"
                DoEvents

            End If

            tipo(0) = 0
            tipo(1) = 0

            For j = 1 To GrhData(i).NumFrames

                If tipo(0) = 0 And tipo(1) = 0 Then
                    tipo(0) = GrhData(GrhData(i).Frames(j)).pixelHeight
                    tipo(1) = GrhData(GrhData(i).Frames(j)).pixelWidth
                Else

                    If tipo(0) <> GrhData(GrhData(i).Frames(j)).pixelHeight Then
                        ' diferente pxHight
                        Datos = Datos & "Grh" & i & " (ANIMACION) en Frame " & j & " - Pixel Height diferente a los demas frames. (Deberia ser " & tipo(0) & " y tiene " & GrhData(GrhData(i).Frames(j)).pixelHeight & ")" & vbCrLf

                    End If

                    If tipo(1) <> GrhData(GrhData(i).Frames(j)).pixelWidth Then
                        ' diferente pxWidth
                        Datos = Datos & "Grh" & i & " (ANIMACION) en Frame " & j & " - Pixel Width diferente a los demas frames. (Deberia ser " & tipo(1) & " y tiene " & GrhData(GrhData(i).Frames(j)).pixelWidth & ")" & vbCrLf

                    End If

                End If

            Next
        ElseIf GrhData(i).NumFrames = 1 Then

            'Tim = Tim + 1
            'If Tim >= 150 Then
            '    Tim = 0
            '    Carga.Label1.Caption = "Procesando " & i & " grh"
            ''    DoEvents
            'End If
            'If LenB(Dir(DirClien & "\Graficos\" & GrhData(i).FileNum & ".bmp", vbArchive)) = 0 Then
            '    Datos = Datos & "Grh" & i & " - Le falta el BMP " & GrhData(i).FileNum & vbCrLf
            'End If
        End If

    Next
    Unload frmCargando
    Me.Show
    frmCodigo.Caption = mnuBuscarErrDim.Caption
    frmCodigo.Codigo.Text = Datos
    frmCodigo.Show

End Sub


Private Sub mnuBuscarGrh_Click()

    On Error Resume Next

    Dim i       As Long
    Dim j       As Long
    Dim Archivo As String

    Archivo = InputBox("Ingrese el numero de GRH:")

    If IsNumeric(Archivo) = False Then Exit Sub
    If LenB(Archivo) > 0 And (Archivo < grhCount) And (Archivo > 0) Then

        For i = 1 To grhCount

            If GrhData(i).NumFrames >= 1 And i = Archivo Then
                DoEvents

                For j = 0 To 39000

                    If ReadField(1, Listado.List(j), Asc(" ")) = Archivo Then
                        MsgBox "GRH encontrado."
                        Listado.ListIndex = j
                        Exit Sub

                    End If

                Next

            End If

        Next
        MsgBox "No se encontro el GRH."
    Else
        MsgBox "Nombre de GRH invalido."

    End If

End Sub

Private Sub mnuBuscarGrhconPNG_Click()
    On Error Resume Next

    Dim i       As Long
    Dim j       As Long
    Dim Archivo As String

    BuscarPNG = 0
    mnuIrASBMP.Enabled = False
    Archivo = InputBox("Ingrese el numero de PNG:")

    If IsNumeric(Archivo) = False Then Exit Sub
    If LenB(Archivo) > 0 And (Archivo > 0) Then

        For i = 1 To grhCount

            If GrhData(i).FileNum = Archivo Then

                For j = 0 To Listado.ListCount - 1

                    If ReadField(1, Listado.List(j), Asc(" ")) = i Then
                        BuscarPNG = Archivo
                        mnuIrASBMP.Enabled = True
                        Listado.ListIndex = j
                        Exit Sub

                    End If

                Next

            End If

        Next
        MsgBox "No se encontro el PNG."
    Else
        MsgBox "Nombre de PNG invalido."

    End If
End Sub

Private Sub mnuIrABMP_Click()

    On Error Resume Next

    Dim i       As Long
    Dim j       As Long
    Dim Archivo As String

    BuscarPNG = 0
    mnuIrASBMP.Enabled = False
    Archivo = InputBox("Ingrese el numero de PNG:")

    If IsNumeric(Archivo) = False Then Exit Sub
    If LenB(Archivo) > 0 And (Archivo > 0) Then

        For i = 1 To grhCount

            If GrhData(i).FileNum = Archivo Then

                For j = 0 To Listado.ListCount - 1

                    If ReadField(1, Listado.List(j), Asc(" ")) = i Then
                        BuscarPNG = Archivo
                        mnuIrASBMP.Enabled = True
                        Listado.ListIndex = j
                        Exit Sub

                    End If

                Next

            End If

        Next
        MsgBox "No se encontro el PNG."
    Else
        MsgBox "Nombre de PNG invalido."

    End If

End Sub

Private Sub mnuCarpetaClienteIr_Click()
    On Error Resume Next

    Call ShellExecute(Me.hwnd, "Open", DirCliente, &O0, &O0, SW_NORMAL)
    
End Sub

Private Sub mnuCarpetaExportacionIr_Click()
    On Error Resume Next

    Call ShellExecute(Me.hwnd, "Open", DirExport, &O0, &O0, SW_NORMAL)
    
End Sub

Private Sub mnuCarpetaIndexacionIr_Click()
    On Error Resume Next

    Call ShellExecute(Me.hwnd, "Open", DirIndex, &O0, &O0, SW_NORMAL)
    
End Sub

Private Sub mnuExportarCabezas_Click()
    Call ExportarCabezas
End Sub

Private Sub mnuExportarCascos_Click()
    Call ExportarCascos
End Sub

Private Sub mnuExportarCuerpos_Click()
    Call ExportarCuerpos
End Sub

Private Sub mnuExportarFxs_Click()
    Call ExportarFxs
End Sub

Private Sub mnuExportarGraficos_Click()
    Call ExportarGraficos
End Sub

Private Sub mnuExportarTodo_Click()
    Call ExportarGraficos
    Call ExportarFxs
    Call ExportarCuerpos
    Call ExportarCascos
    Call ExportarCabezas
End Sub

Private Sub mnuGenerarMinimapa_Click()
    frmMinimapa.Show , frmMain
End Sub

Private Sub mnuIndexarGraficos_Click()
    DoEvents

    If IndexarGraficos Then
        MsgBox "Graficos.ind creado..."
    Else
        MsgBox "Error al crear Graficos.ind..."

    End If

End Sub

Private Sub mnuIndexarCabezas_Click()
    If IndexarCabezas Then
        MsgBox "Head.ind creado..."
    Else
        MsgBox "Error al crear Head.ind..."

    End If
End Sub

Private Sub mnuIndexarPersonajes_Click()
    DoEvents
    
    If IndexarCuerpos Then
        MsgBox "Personajes.ind creado..."
    Else
        MsgBox "Error al crear Personajes.ind..."
    End If
    
End Sub

Private Sub mnuIndexarCascos_Click()
    DoEvents
    
    If IndexarCascos Then
        MsgBox "Helmet.ind creado..."
    Else
        MsgBox "Error al crear Helmet.ind..."
    End If
End Sub

Private Sub mnuIndexarArmas_Click()
    DoEvents
    
    If IndexarArmas Then
        MsgBox "Armas.ind creado..."
    Else
        MsgBox "Error al crear Armas.ind..."
    End If
End Sub

Private Sub mnuIndexarEscudos_Click()
    DoEvents
    
    If IndexarEscudos Then
        MsgBox "Escudos.ind creado..."
    Else
        MsgBox "Error al crear Escudos.ind..."
    End If
End Sub

Private Sub mnuIndexarFXs_Click()
    DoEvents
    
    If IndexarFXs Then
        MsgBox "FXs.ind creado..."
    Else
        MsgBox "Error al crear FXs.ind..."
    End If
End Sub

Private Sub mnuIndexarTodo_Click()

    Call mnuIndexarGraficos_Click
    Call mnuIndexarPersonajes_Click
    Call mnuIndexarCabezas_Click
    Call mnuIndexarCascos_Click
    Call mnuIndexarArmas_Click
    Call mnuIndexarEscudos_Click
    Call mnuIndexarFXs_Click
    
End Sub
