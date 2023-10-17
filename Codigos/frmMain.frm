VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indexador Nexus"
   ClientHeight    =   11865
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   12360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   791
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   824
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox Listado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H0000FF00&
      Height          =   10170
      Left            =   60
      TabIndex        =   31
      Top             =   360
      Width           =   2235
   End
   Begin VB.ListBox ListaAtaques 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1395
      Left            =   9570
      TabIndex        =   29
      Top             =   10290
      Width           =   2700
   End
   Begin VB.Frame FraAtributosDel 
      BackColor       =   &H00404040&
      Caption         =   "Atributos del Frame"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   585
      Left            =   120
      TabIndex        =   16
      Top             =   10740
      Width           =   9405
      Begin VB.TextBox txtFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   5
         Left            =   8370
         TabIndex        =   28
         Top             =   220
         Width           =   645
      End
      Begin VB.TextBox txtFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   4
         Left            =   7020
         TabIndex        =   27
         Top             =   225
         Width           =   645
      End
      Begin VB.TextBox txtFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   3
         Left            =   5400
         TabIndex        =   26
         Top             =   220
         Width           =   645
      End
      Begin VB.TextBox txtFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   2
         Left            =   4230
         TabIndex        =   25
         Top             =   220
         Width           =   645
      End
      Begin VB.TextBox txtFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   1
         Left            =   2820
         TabIndex        =   24
         Top             =   225
         Width           =   645
      End
      Begin VB.TextBox txtFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   0
         Left            =   990
         TabIndex        =   23
         Top             =   220
         Width           =   645
      End
      Begin VB.Label lblAlto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alto:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   7830
         TabIndex        =   22
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lblAncho 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   6390
         TabIndex        =   21
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lblY 
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   5190
         TabIndex        =   20
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   4020
         TabIndex        =   19
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lblNFrames 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N. Frames:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblnGraficos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Graficos:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2100
         TabIndex        =   17
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.FileListBox Archivos 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   4110
      Pattern         =   "*.bmp"
      TabIndex        =   15
      Top             =   30
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox GRHt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   450
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   11340
      Width           =   9405
   End
   Begin VB.ListBox ListaArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1395
      Left            =   9570
      TabIndex        =   12
      Top             =   5340
      Width           =   2700
   End
   Begin VB.ListBox ListaFxs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1395
      Left            =   9570
      TabIndex        =   10
      Top             =   8610
      Width           =   2700
   End
   Begin VB.ListBox ListaEscudos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1395
      Left            =   9570
      TabIndex        =   8
      Top             =   6960
      Width           =   2700
   End
   Begin VB.ListBox ListaCascos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1395
      Left            =   9570
      TabIndex        =   6
      Top             =   2070
      Width           =   2700
   End
   Begin VB.ListBox ListaHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1395
      Left            =   9570
      TabIndex        =   4
      Top             =   390
      Width           =   2700
   End
   Begin VB.PictureBox MainViewPic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   10335
      Left            =   2340
      ScaleHeight     =   687
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   478
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   7200
   End
   Begin VB.ListBox ListaCuerpos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1395
      Left            =   9570
      TabIndex        =   0
      Top             =   3720
      Width           =   2700
   End
   Begin VB.Label lblAtaques 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ataques"
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
      Left            =   10410
      TabIndex        =   30
      Top             =   10020
      Width           =   675
   End
   Begin VB.Label lblArmas 
      Alignment       =   2  'Center
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
      Left            =   10320
      TabIndex        =   13
      Top             =   5070
      Width           =   675
   End
   Begin VB.Label lblFxS 
      Alignment       =   2  'Center
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
      Left            =   10410
      TabIndex        =   11
      Top             =   8340
      Width           =   675
   End
   Begin VB.Label lblEscudos 
      Alignment       =   2  'Center
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
      Left            =   10380
      TabIndex        =   9
      Top             =   6690
      Width           =   675
   End
   Begin VB.Label lblCascos 
      Alignment       =   2  'Center
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
      Left            =   10470
      TabIndex        =   7
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label lblHead 
      Alignment       =   2  'Center
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
      Left            =   10380
      TabIndex        =   5
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lbCuerpos 
      Alignment       =   2  'Center
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
      Left            =   10350
      TabIndex        =   2
      Top             =   3450
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
      TabIndex        =   1
      Top             =   120
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
         Begin VB.Menu cmdExportarParticulas 
            Caption         =   "Particulas.ind"
         End
         Begin VB.Menu mnuExportarColores 
            Caption         =   "Colores.ind"
         End
      End
      Begin VB.Menu mnuIndexar 
         Caption         =   "Indexar..."
         Begin VB.Menu mnuIndexarTodo 
            Caption         =   "&TODO"
         End
         Begin VB.Menu mnuIndexarGraficos 
            Caption         =   "Graficos.ini"
         End
         Begin VB.Menu mnuIndexarCabezas 
            Caption         =   "Cabezas.ini"
         End
         Begin VB.Menu mnuIndexarPersonajes 
            Caption         =   "Personajes.ini"
         End
         Begin VB.Menu mnuIndexarCascos 
            Caption         =   "Cascos.ini"
         End
         Begin VB.Menu mnuIndexarArmas 
            Caption         =   "Armas.ini"
         End
         Begin VB.Menu mnuIndexarEscudos 
            Caption         =   "Escudos.ini"
         End
         Begin VB.Menu mnuIndexarFXs 
            Caption         =   "FXs.ini"
         End
         Begin VB.Menu mnuIndexarParticulas 
            Caption         =   "Particulas.ini"
         End
         Begin VB.Menu mnuIndexarColores 
            Caption         =   "Colores.ini"
         End
      End
      Begin VB.Menu mnuGenerarMinimapa 
         Caption         =   "Generar Minimapa"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu menuCerrar 
         Caption         =   "Cerrar"
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
         Caption         =   "Buscar Errores de Dimensiones"
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuadaptador 
         Caption         =   "&Adaptador de Grh"
      End
   End
   Begin VB.Menu mnuParticleEditor 
      Caption         =   "Editor de Particulas"
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

Private Sub ListaArmas_Click()
    Dim nGrh As Long

    nGrh = WeaponAnimData(ListaArmas.Text).WeaponWalk(3).GrhIndex
    
    DoEvents
    Call InitGrh(CurrentGrh, nGrh)
End Sub

Private Sub ListaCascos_Click()
    Dim nGrh As Long

    nGrh = CascoAnimData(ListaCascos.Text).Head(3).GrhIndex
    
    DoEvents
    Call InitGrh(CurrentGrh, nGrh)
End Sub

Private Sub ListaCuerpos_Click()
    Dim nGrh As Long

    nGrh = BodyData(ListaCuerpos.Text).Walk(3).GrhIndex
    
    DoEvents
    Call InitGrh(CurrentGrh, nGrh)
End Sub

Private Sub Listado_Click()
    Dim nGrh As Long

    nGrh = ReadField(1, Listado.Text, Asc(" "))

    DoEvents
    Call InitGrh(CurrentGrh, nGrh)

    txtFrame(0).Text = GrhData(nGrh).NumFrames
    txtFrame(1).Text = GrhData(nGrh).FileNum
    txtFrame(2).Text = GrhData(nGrh).sX
    txtFrame(3).Text = GrhData(nGrh).sY
    txtFrame(4).Text = GrhData(nGrh).pixelHeight
    txtFrame(5).Text = GrhData(nGrh).pixelWidth
    
End Sub

Private Sub ListaEscudos_Click()
    Dim nGrh As Long

    nGrh = ShieldAnimData(ListaEscudos.Text).ShieldWalk(3).GrhIndex
    
    DoEvents
    Call InitGrh(CurrentGrh, nGrh)
End Sub

Private Sub ListaFxs_Click()
    Dim nGrh As Long

    nGrh = FxData(ListaFxs.Text).Animacion
    
    DoEvents
    Call InitGrh(CurrentGrh, nGrh)
End Sub

Private Sub ListaHead_Click()
    Dim nGrh As Long

    nGrh = HeadData(ListaHead.Text).Head(3).GrhIndex
    
    DoEvents
    Call InitGrh(CurrentGrh, nGrh)
End Sub

Private Sub menuCerrar_Click()
    Call CloseClient
End Sub

Private Sub mnuadaptador_Click()
    frmAdaptador.Show , frmMain
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

Private Sub mnuParticleEditor_Click()
    frmParticleEditor.Show , frmMain
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
    Archivo = InputBox("Ingrese el numero de grafico:")

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

                Next j

            End If

        Next i
                
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

Private Sub cmdExportarParticulas_Click()
    Call ExportarParticulas
End Sub

Private Sub mnuExportarTodo_Click()
    Call ExportarGraficos
    Call ExportarFxs
    Call ExportarCuerpos
    Call ExportarCascos
    Call ExportarCabezas
    Call ExportarParticulas
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

Private Sub mnuIndexarParticulas_Click()
    DoEvents
    
    If IndexarParticulas Then
        MsgBox "Particulas.ind creado..."
    Else
        MsgBox "Error al crear Particulas.ind..."
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
    Call mnuIndexarParticulas_Click
    
End Sub

