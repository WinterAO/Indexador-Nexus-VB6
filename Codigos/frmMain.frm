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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   471
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   774
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
