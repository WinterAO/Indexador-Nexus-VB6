VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indexador Nexus"
   ClientHeight    =   11625
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   19560
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
   ScaleHeight     =   775
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1304
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraTerreno 
      BackColor       =   &H00404040&
      Caption         =   "Terreno"
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
      Height          =   2655
      Left            =   9600
      TabIndex        =   36
      Top             =   2850
      Width           =   3465
      Begin Indexador_Nexus.lvButtons_H LvBIndexarTerreno 
         Height          =   525
         Left            =   2550
         TabIndex        =   54
         Top             =   1500
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   926
         Caption         =   "Indexar Terreno"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.TextBox tITDesde 
         Height          =   285
         Left            =   1140
         TabIndex        =   45
         Text            =   "0"
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox tITbmp 
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Text            =   "0"
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox tITLargo 
         Height          =   285
         Left            =   2640
         TabIndex        =   43
         Text            =   "2"
         Top             =   1110
         Width           =   735
      End
      Begin VB.TextBox tITAncho 
         Height          =   285
         Left            =   1800
         TabIndex        =   42
         Text            =   "2"
         Top             =   1110
         Width           =   735
      End
      Begin VB.TextBox tITY 
         Height          =   285
         Left            =   960
         TabIndex        =   41
         Text            =   "0"
         Top             =   1110
         Width           =   735
      End
      Begin VB.TextBox tITX 
         Height          =   285
         Left            =   120
         TabIndex        =   40
         Text            =   "0"
         Top             =   1110
         Width           =   615
      End
      Begin VB.TextBox tITNombre 
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   1740
         Width           =   2295
      End
      Begin VB.TextBox tITSolo 
         Height          =   285
         Left            =   960
         TabIndex        =   38
         Text            =   "0"
         Top             =   2130
         Width           =   855
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   2130
         Width           =   975
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grh Inicial:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1170
         TabIndex        =   53
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grafico:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Largo:"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   2640
         TabIndex        =   51
         Top             =   870
         Width           =   855
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho:"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1800
         TabIndex        =   50
         Top             =   870
         Width           =   735
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   960
         TabIndex        =   49
         Top             =   870
         Width           =   735
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   870
         Width           =   495
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Terreno"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Solo el grh:"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2130
         Width           =   855
      End
   End
   Begin Indexador_Nexus.lvButtons_H LvBNuevoGrh 
      Height          =   405
      Left            =   1020
      TabIndex        =   34
      Top             =   9990
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   714
      Caption         =   "Nuevo"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454016
   End
   Begin Indexador_Nexus.LynxGrid LynxGrh 
      Height          =   9105
      Left            =   90
      TabIndex        =   1
      Top             =   810
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   16060
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   5460819
      BackColorBkg    =   5460819
      BackColorEdit   =   14737632
      BackColorSel    =   12937777
      ForeColor       =   12632256
      ForeColorSel    =   8438015
      BackColorEvenRows=   3158064
      CustomColorFrom =   4210752
      CustomColorTo   =   8421504
      GridColor       =   14737632
      FocusRectColor  =   9895934
      GridLines       =   2
      ThemeColor      =   5
      ScrollBars      =   1
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      HotHeaderTracking=   0   'False
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
      Height          =   2025
      Left            =   9600
      TabIndex        =   3
      Top             =   810
      Width           =   3465
      Begin VB.CommandButton cmdAlto 
         Caption         =   "-"
         Height          =   195
         Index           =   1
         Left            =   3150
         TabIndex        =   33
         Top             =   1320
         Width           =   180
      End
      Begin VB.CommandButton cmdAlto 
         Caption         =   "+"
         Height          =   195
         Index           =   0
         Left            =   3150
         TabIndex        =   32
         Top             =   1110
         Width           =   180
      End
      Begin VB.CommandButton cmdAncho 
         Caption         =   "-"
         Height          =   195
         Index           =   1
         Left            =   1650
         TabIndex        =   31
         Top             =   1320
         Width           =   180
      End
      Begin VB.CommandButton cmdAncho 
         Caption         =   "+"
         Height          =   195
         Index           =   0
         Left            =   1650
         TabIndex        =   30
         Top             =   1110
         Width           =   180
      End
      Begin VB.CommandButton cmdsY 
         Caption         =   "-"
         Height          =   195
         Index           =   1
         Left            =   3150
         TabIndex        =   29
         Top             =   850
         Width           =   180
      End
      Begin VB.CommandButton cmdsY 
         Caption         =   "+"
         Height          =   195
         Index           =   0
         Left            =   3150
         TabIndex        =   28
         Top             =   650
         Width           =   180
      End
      Begin VB.CommandButton cmdSX 
         Caption         =   "-"
         Height          =   195
         Index           =   1
         Left            =   1650
         TabIndex        =   27
         Top             =   850
         Width           =   180
      End
      Begin VB.CommandButton cmdSX 
         Caption         =   "+"
         Height          =   195
         Index           =   0
         Left            =   1650
         TabIndex        =   26
         Top             =   650
         Width           =   180
      End
      Begin Indexador_Nexus.lvButtons_H LvBCambiar 
         Height          =   315
         Left            =   90
         TabIndex        =   25
         Top             =   1560
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "Cambiar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   5
         Left            =   2460
         TabIndex        =   12
         Top             =   1150
         Width           =   645
      End
      Begin VB.TextBox txtFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   14
         Top             =   1150
         Width           =   645
      End
      Begin VB.TextBox txtFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   3
         Left            =   2460
         TabIndex        =   9
         Top             =   720
         Width           =   645
      End
      Begin VB.TextBox txtFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   7
         Top             =   735
         Width           =   645
      End
      Begin VB.TextBox txtFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   1
         Left            =   2460
         TabIndex        =   13
         Top             =   300
         Width           =   645
      End
      Begin VB.TextBox txtFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   300
         Width           =   645
      End
      Begin VB.Label lblAlto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alto:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1920
         TabIndex        =   15
         Top             =   1110
         Width           =   645
      End
      Begin VB.Label lblAncho 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   330
         TabIndex        =   11
         Top             =   1140
         Width           =   645
      End
      Begin VB.Label lblY 
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2130
         TabIndex        =   10
         Top             =   780
         Width           =   645
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   750
         TabIndex        =   8
         Top             =   780
         Width           =   645
      End
      Begin VB.Label lblNFrames 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N. Frames:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblnGraficos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grafico:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1740
         TabIndex        =   6
         Top             =   390
         Width           =   570
      End
   End
   Begin VB.FileListBox Archivos 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   7740
      Pattern         =   "*.bmp"
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox GRHt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   1080
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   10440
      Width           =   9405
   End
   Begin VB.PictureBox MainViewPic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   9585
      Left            =   2340
      ScaleHeight     =   637
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   478
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   810
      Width           =   7200
   End
   Begin Indexador_Nexus.lvButtons_H LvBSelector 
      Height          =   495
      Index           =   0
      Left            =   150
      TabIndex        =   17
      ToolTipText     =   "Cuerpos"
      Top             =   150
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":10CA
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin Indexador_Nexus.lvButtons_H LvBSelector 
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   18
      ToolTipText     =   "Cabezas"
      Top             =   150
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":1310
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin Indexador_Nexus.lvButtons_H LvBSelector 
      Height          =   495
      Index           =   2
      Left            =   1290
      TabIndex        =   19
      ToolTipText     =   "Cascos"
      Top             =   150
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":16AD
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin Indexador_Nexus.lvButtons_H LvBSelector 
      Height          =   495
      Index           =   3
      Left            =   1860
      TabIndex        =   20
      ToolTipText     =   "Armas"
      Top             =   150
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":19B2
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin Indexador_Nexus.lvButtons_H LvBSelector 
      Height          =   495
      Index           =   4
      Left            =   2430
      TabIndex        =   21
      ToolTipText     =   "Escudos"
      Top             =   150
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":1EA1
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin Indexador_Nexus.lvButtons_H LvBSelector 
      Height          =   495
      Index           =   5
      Left            =   3000
      TabIndex        =   22
      ToolTipText     =   "Ataques"
      Top             =   150
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":24E4
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin Indexador_Nexus.lvButtons_H LvBSelector 
      Height          =   495
      Index           =   6
      Left            =   3570
      TabIndex        =   23
      ToolTipText     =   "Fx's"
      Top             =   150
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   873
      Caption         =   "A"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin Indexador_Nexus.lvButtons_H LvBSelector 
      Height          =   495
      Index           =   7
      Left            =   4140
      TabIndex        =   24
      ToolTipText     =   "Fx's"
      Top             =   150
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   4
      Image           =   "frmMain.frx":2B74
      ImgSize         =   32
      cBack           =   -2147483633
   End
   Begin Indexador_Nexus.lvButtons_H LvBBorrar 
      Height          =   405
      Left            =   120
      TabIndex        =   35
      Top             =   9990
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   714
      Caption         =   "Borrar"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   8421631
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuIndexMemory 
         Caption         =   "Indexar desde Memoria"
      End
      Begin VB.Menu mnuExportar 
         Caption         =   "Exportar..."
         Begin VB.Menu mnuExportarGraficos 
            Caption         =   "Graficos.ind"
         End
         Begin VB.Menu mnuExportarCabezas 
            Caption         =   "Head.ind"
         End
         Begin VB.Menu mnuExportarCascos 
            Caption         =   "Helmet.ind"
         End
         Begin VB.Menu mnuExportarCuerpos 
            Caption         =   "Personajes.ind"
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
         Begin VB.Menu mnuExportarGUI 
            Caption         =   "GUI.ind"
         End
         Begin VB.Menu mnuLine5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExportarTodo 
            Caption         =   "TODO"
         End
      End
      Begin VB.Menu mnuIndexar 
         Caption         =   "Indexar..."
         Begin VB.Menu mnuIndexarGraficos 
            Caption         =   "Graficos.ini"
         End
         Begin VB.Menu mnuIndexarCabezas 
            Caption         =   "Head.ini"
         End
         Begin VB.Menu mnuIndexarPersonajes 
            Caption         =   "Personajes.ini"
         End
         Begin VB.Menu mnuIndexarCascos 
            Caption         =   "Helmet.ini"
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
         Begin VB.Menu mnuIndexarGUI 
            Caption         =   "GUI.ini"
         End
         Begin VB.Menu mnuLine4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIndexarTodo 
            Caption         =   "&TODO"
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
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuscarGrhconPNG 
         Caption         =   "&Buscar Grh con PNG"
         Enabled         =   0   'False
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NO_GRH As Long = 0
Private BuscarPNG As Integer

Private Sub cmdAgregar_Click()
    MaxSup = MaxSup + 1
    ReDim Preserve SupData(MaxSup) As SupData

    With SupData(MaxSup)
        .Grh = Val(tITSolo.Text)
        .Block = False
        .Capa = 0
        .Width = 1
        .Height = 1
        .name = tITNombre.Text

    End With

End Sub

Private Sub cmdAlto_Click(Index As Integer)
    Select Case Index
    
        Case 0
            txtFrame(5).Text = Val(txtFrame(5).Text + 1)
        
        Case 1
            If Val(txtFrame(5).Text) > 1 Then _
                txtFrame(5).Text = Val(txtFrame(5).Text - 1)
    
    End Select
    
    GrhData(CurrentGrh.GrhIndex).pixelWidth = Val(txtFrame(5).Text)
End Sub

Private Sub cmdAncho_Click(Index As Integer)
    Select Case Index
    
        Case 0
            txtFrame(4).Text = Val(txtFrame(4).Text + 1)
        
        Case 1
            If Val(txtFrame(4).Text) > 1 Then _
                txtFrame(4).Text = Val(txtFrame(4).Text - 1)
    
    End Select
    
    GrhData(CurrentGrh.GrhIndex).pixelHeight = Val(txtFrame(4).Text)
End Sub

Private Sub cmdSX_Click(Index As Integer)
    
    Select Case Index
    
        Case 0
            txtFrame(2).Text = Val(txtFrame(2).Text + 1)
        
        Case 1
            If Val(txtFrame(2).Text) > 1 Then _
                txtFrame(2).Text = Val(txtFrame(2).Text - 1)
    
    End Select
    
    GrhData(CurrentGrh.GrhIndex).sX = Val(txtFrame(2).Text)
    
End Sub

Private Sub cmdsY_Click(Index As Integer)
    Select Case Index
    
        Case 0
            txtFrame(3).Text = Val(txtFrame(3).Text + 1)
        
        Case 1
            If Val(txtFrame(3).Text) > 1 Then _
                txtFrame(3).Text = Val(txtFrame(3).Text - 1)
    
    End Select
    
    GrhData(CurrentGrh.GrhIndex).sY = Val(txtFrame(3).Text)
End Sub

Private Sub Form_Load()
    EngineRun = True
End Sub

Private Sub LvBBorrar_Click()
    
    If Not isList Then
        MsgBox "Selecciona un grh de la lista de Grh general que desees borrar."
        Exit Sub
    End If

    If CurrentGrh.GrhIndex = NO_GRH Then
        MsgBox "No has seleccionado ningun Grh"
        Exit Sub
    End If
    
    If MsgBox("¿Seguro que quieres borrar el Grh: " & CurrentGrh.GrhIndex & "?" & vbCrLf & "Este cambio no tiene vuelta atrás.", vbOKCancel) = vbOK Then
        'Reset it
        With GrhData(CurrentGrh.GrhIndex)
            .FileNum = 0
            ReDim .Frames(0)
            .NumFrames = 0
            .pixelHeight = 0
            .pixelWidth = 0
            .speed = 0
            .sX = 0
            .sY = 0
            .TileHeight = 0
            .TileWidth = 0
        End With
        
        CurrentGrh.Started = 0
        
'        'Remove it
'        For i = 0 To grhlist.ListCount - 1
'            If Val(grhlist.List(i)) = CurrentGrh Then
'                grhlist.RemoveItem (i)
'                Exit For
'            End If
'        Next i
'
'        'Select next grh
'        If i < grhlist.ListCount Then
'            grhlist.ListIndex = i
'        Else
'            grhlist.ListIndex = grhlist.ListCount - 1
'        End If
    End If
    
End Sub

Private Sub LvBCambiar_Click()
'    'Prevent non numeric characters
'    If Not IsNumeric(txtFrame(Index).Text) Then
'        txtFrame(Index).Text = Val(txtFrame(Index).Text)
'
'    End If
'
'    'Prevent overflow
'    If Val(txtFrame(Index).Text) > &H7FFF Then
'        txtFrame(Index).Text = &H7FFF
'
'    End If
'
'    'Prevent negative values
'    If CInt(txtFrame(Index).Text) < 0 Then
'        txtFrame(Index).Text = 0
'
'    End If
    
    'Update data in memory
    If CurrentGrh.GrhIndex <> NO_GRH Then
        With GrhData(CurrentGrh.GrhIndex)

            .NumFrames = Val(txtFrame(0).Text)
            .FileNum = Val(txtFrame(1).Text)
            .sX = Val(txtFrame(2).Text)
            .sY = Val(txtFrame(3).Text)
            .pixelHeight = Val(txtFrame(4).Text)
            .pixelWidth = Val(txtFrame(5).Text)
            Debug.Print .NumFrames & " - " & .FileNum & " - " & .sX & " - " & .sY & " - " & .pixelHeight & " - " & .pixelWidth
        
        End With
    End If
End Sub

Private Sub LvBIndexarTerreno_Click()

    Dim W      As Integer

    Dim H      As Integer

    Dim x      As Integer

    Dim y      As Integer

    Dim iList  As Integer

    Dim GI     As Integer

    Dim Indice As Integer

    Dim DGRH   As Integer

    W = Val(tITAncho.Text)
    H = Val(tITLargo.Text)

    DGRH = Val(tITDesde.Text)

    For y = 1 To H
        For x = 1 To W
            GI = AgregarGrh(iList, False, DGRH)

            If DGRH > 0 Then DGRH = DGRH + 1

            With GrhData(GI)
                .FileNum = Val(tITbmp.Text)
                .sX = Val(tITX.Text) * 32 * W + (x - 1) * 32
                .sY = Val(tITY.Text) * 32 * H + (y - 1) * 32
                .NumFrames = 1
                ReDim .Frames(1 To 1)
                .Frames(1) = GI
                .pixelWidth = 32
                .pixelHeight = 32
                .TileWidth = 32
                .TileHeight = 32

            End With
        
            If x = 1 And y = 1 Then
                Indice = GI

            End If

        Next x
    Next y

    MaxSup = MaxSup + 1
    ReDim Preserve SupData(MaxSup) As SupData

    With SupData(MaxSup)
        .Grh = Indice
        .Block = False
        .Capa = 0
        .Width = W
        .Height = H
        .name = tITNombre.Text

    End With

End Sub

Public Function AgregarGrh(ByRef EnLista As Integer, _
                           Optional Animacion As Boolean = False, _
                           Optional DGRH As Integer = 0) As Integer

'    Dim Index As Integer
'
'    Dim i     As Integer
'
'    If DGRH = 0 Then
'        Index = UBound(GrhData()) + 1
'    Else
'        Index = DGRH
'
'    End If
'
'    'Make sure he is not overwritting anything
'    If Index <= UBound(GrhData()) Then
'        If GrhData(Index).NumFrames > 0 Then
'
'            'If MsgBox("The chosen index is currently in use. Do you want to overwrite it?", vbOKCancel) = vbCancel Then
'            '    Exit Function
'            'End If
'        End If
'
'    Else
'        'Resize array
'        ReDim Preserve GrhData(1 To Index) As GrhData
'
'    End If
'
'    If GrhData(Index).NumFrames = 0 Then
'
'        'Search where to place the grh....
'        For i = 0 To frmMain.LynxGrh.ItemCount - 1
'
'            If Val(frmMain.LynxGrh.List(i)) > Index Then
'                Exit For
'
'            End If
'
'        Next i
'
'        'Add it!
'        Call frmMain.LynxGrh.AddItem(Index & IIf(Animacion, " (ANIMACIÓN)", ""), i)
'    Else
'
'        'Search for the grh index within the grhList
'        For i = 0 To frmMain.LynxGrh.ListCount - 1
'
'            If Val(frmMain.LynxGrh.List(i)) = Index Then
'                Exit For
'
'            End If
'
'        Next i
'
'    End If
'
'    AgregarGrh = Index

End Function

Private Sub LvBSelector_Click(Index As Integer)
    Select Case Index
    
        Case 0 'Cuerpos
        
            If frmCuerpos.Visible Then
                frmCuerpos.Visible = False
            Else
                frmCuerpos.Show , frmMain
            End If
            
        Case 1 'Cabezas
        
            If frmCabezas.Visible Then
                frmCabezas.Visible = False
            Else
                frmCabezas.Show , frmMain
            End If
        
        Case 2 'Cascos
        
            If frmCascos.Visible Then
                frmCascos.Visible = False
            Else
                frmCascos.Show , frmMain
            End If
        
        Case 3 'Armas
        
            If frmArmas.Visible Then
                frmArmas.Visible = False
            Else
                frmArmas.Show , frmMain
            End If
            
        Case 4 'Escudos
        
            If frmEscudos.Visible Then
                frmEscudos.Visible = False
            Else
                frmEscudos.Show , frmMain
            End If
        
        Case 5 'FX's
            If frmFxs.Visible Then
                frmFxs.Visible = False
            Else
                frmFxs.Show , frmMain
            End If
        
        Case 6 'Ataques
            If frmAtaques.Visible Then
                frmAtaques.Visible = False
            Else
                frmAtaques.Show , frmMain
            End If
            
        Case 7 'Particulas
            If frmParticleEditor.Visible Then
                frmParticleEditor.Visible = False
            Else
                frmParticleEditor.Show , frmMain
            End If
    
    End Select
    
End Sub

Private Sub LynxGrh_Click()
    Dim nGrh As Long

    nGrh = LynxGrh.CellText(, 0)

    DoEvents
    Call InitGrh(CurrentGrh, nGrh)

    'Mostramos el grh info
    txtFrame(0).Text = GrhData(nGrh).NumFrames
    txtFrame(1).Text = GrhData(nGrh).FileNum
    txtFrame(2).Text = GrhData(nGrh).sX
    txtFrame(3).Text = GrhData(nGrh).sY
    txtFrame(4).Text = GrhData(nGrh).pixelHeight
    txtFrame(5).Text = GrhData(nGrh).pixelWidth
    
    isList = True
    Particle_Group_Remove_All
    
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

Private Sub mnuIndexMemory_Click()
    If IndexarfromMemory Then
        GRHt.Text = "Graficos.ind compilado..."
    Else
        MsgBox "Error al compilar Graficos.ind..."

    End If
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

'    On Error Resume Next
'
'    Dim i       As Long
'    Dim j       As Long
'    Dim Archivo As String
'
'    Archivo = InputBox("Ingrese el numero de GRH:")
'
'    If IsNumeric(Archivo) = False Then Exit Sub
'    If LenB(Archivo) > 0 And (Archivo < grhCount) And (Archivo > 0) Then
'
'        For i = 1 To grhCount
'
'            If GrhData(i).NumFrames >= 1 And i = Archivo Then
'                DoEvents
'
'                For j = 0 To 39000
'
'                    If ReadField(1, Listado.List(j), Asc(" ")) = Archivo Then
'                        MsgBox "GRH encontrado."
'                        Listado.ListIndex = j
'                        Exit Sub
'
'                    End If
'
'                Next
'
'            End If
'
'        Next
'        MsgBox "No se encontro el GRH."
'    Else
'        MsgBox "Nombre de GRH invalido."
'
'    End If

End Sub

Private Sub mnuBuscarGrhconPNG_Click()

'    On Error Resume Next
'
'    Dim i       As Long
'
'    Dim j       As Long
'
'    Dim Archivo As String
'
'    BuscarPNG = 0
'    mnuIrASBMP.Enabled = False
'    Archivo = InputBox("Ingrese el numero de grafico:")
'
'    If IsNumeric(Archivo) = False Then Exit Sub
'    If LenB(Archivo) > 0 And (Archivo > 0) Then
'
'        For i = 1 To grhCount
'
'            If GrhData(i).FileNum = Archivo Then
'
'                For j = 0 To Listado.ListCount - 1
'
'                    If ReadField(1, Listado.List(j), Asc(" ")) = i Then
'                        BuscarPNG = Archivo
'                        mnuIrASBMP.Enabled = True
'                        Listado.ListIndex = j
'                        Exit Sub
'
'                    End If
'
'                Next j
'
'            End If
'
'        Next i
'
'        MsgBox "No se encontro el PNG."
'    Else
'        MsgBox "Nombre de PNG invalido."
'
'    End If

End Sub

Private Sub mnuIrABMP_Click()

'    On Error Resume Next
'
'    Dim i       As Long
'    Dim j       As Long
'    Dim Archivo As String
'
'    BuscarPNG = 0
'    mnuIrASBMP.Enabled = False
'    Archivo = InputBox("Ingrese el numero de PNG:")
'
'    If IsNumeric(Archivo) = False Then Exit Sub
'    If LenB(Archivo) > 0 And (Archivo > 0) Then
'
'        For i = 1 To grhCount
'
'            If GrhData(i).FileNum = Archivo Then
'
'                For j = 0 To Listado.ListCount - 1
'
'                    If ReadField(1, Listado.List(j), Asc(" ")) = i Then
'                        BuscarPNG = Archivo
'                        mnuIrASBMP.Enabled = True
'                        Listado.ListIndex = j
'                        Exit Sub
'
'                    End If
'
'                Next
'
'            End If
'
'        Next
'        MsgBox "No se encontro el PNG."
'    Else
'        MsgBox "Nombre de PNG invalido."
'
'    End If

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

Private Sub mnuExportarColores_Click()
    Call ExportarColores
End Sub

Private Sub mnuExportarTodo_Click()
    Call ExportarGraficos
    Call ExportarFxs
    Call ExportarCuerpos
    Call ExportarCascos
    Call ExportarCabezas
    Call ExportarParticulas
    Call ExportarColores
    'Call ExportarGUI
End Sub

Private Sub mnuGenerarMinimapa_Click()
    frmMinimapa.Show , frmMain
End Sub

Private Sub mnuIndexarGraficos_Click()
    DoEvents

    If IndexarGraficos Then
        GRHt.Text = "Graficos.ind compilado..."
    Else
        MsgBox "Error al compilar Graficos.ind..."

    End If

End Sub

Private Sub mnuIndexarCabezas_Click()
    If IndexarCabezas Then
        GRHt.Text = "Head.ind compilado..."
    Else
        MsgBox "Error al compilar Head.ind..."

    End If
End Sub

Private Sub mnuIndexarPersonajes_Click()
    DoEvents
    
    If IndexarCuerpos Then
        GRHt.Text = "Personajes.ind compilado..."
    Else
        MsgBox "Error al compilar Personajes.ind..."
    End If
    
End Sub

Private Sub mnuIndexarCascos_Click()
    DoEvents
    
    If IndexarCascos Then
        GRHt.Text = "Helmet.ind compilado..."
    Else
        MsgBox "Error al compilar Helmet.ind..."
    End If
End Sub

Private Sub mnuIndexarArmas_Click()
    DoEvents
    
    If IndexarArmas Then
        GRHt.Text = "Armas.ind compilado..."
    Else
        MsgBox "Error al compilar Armas.ind..."
    End If
End Sub

Private Sub mnuIndexarEscudos_Click()
    DoEvents
    
    If IndexarEscudos Then
        GRHt.Text = "Escudos.ind compilado..."
    Else
        MsgBox "Error al compilar Escudos.ind..."
    End If
End Sub

Private Sub mnuIndexarFXs_Click()
    DoEvents
    
    If IndexarFXs Then
        GRHt.Text = "FXs.ind compilado..."
    Else
        MsgBox "Error al compilar FXs.ind..."
    End If
End Sub

Private Sub mnuIndexarParticulas_Click()
    DoEvents
    
    If IndexarParticulas Then
        GRHt.Text = "Particulas.ind compilado..."
    Else
        MsgBox "Error al compilar Particulas.ind..."
    End If
End Sub

Private Sub mnuIndexarColores_Click()

    If IndexarColores Then
        GRHt.Text = "Colores.ind compilado..."
    Else
        MsgBox "Error al compilar Colores.ind..."
    End If
End Sub

Private Sub mnuIndexarGUI_Click()
    If IndexarGUI Then
        GRHt.Text = "GUI.ind compilado..."
    Else
        MsgBox "Error al compilar GUI.ind..."
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
    Call mnuIndexarColores_Click
    Call mnuIndexarGUI_Click
    
End Sub
