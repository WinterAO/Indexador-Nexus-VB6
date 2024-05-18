VERSION 5.00
Begin VB.Form frmMinimapa 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Minimap Color Finder"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4110
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
   Icon            =   "frmMinimapa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   274
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3270
      Left            =   120
      ScaleHeight     =   3270
      ScaleWidth      =   3900
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1110
      Width           =   3900
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2670
      ScaleHeight     =   480
      ScaleWidth      =   585
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4470
      Width           =   585
   End
   Begin Indexador_Nexus.lvButtons_H cmdComenzar 
      Height          =   555
      Left            =   1860
      TabIndex        =   4
      Top             =   180
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   979
      Caption         =   "Comenzar"
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
   Begin Indexador_Nexus.lvButtons_H cmdCerrar 
      Height          =   585
      Left            =   90
      TabIndex        =   5
      Top             =   180
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   1032
      Caption         =   "Cerrrar"
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
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   150
      Left            =   120
      Top             =   1230
      Width           =   15
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Graficos.ind cargado!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   3870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Promedio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   870
      TabIndex        =   2
      Top             =   4590
      Width           =   1665
   End
End
Attribute VB_Name = "frmMinimapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdComenzar_Click()

    On Error GoTo cmdComenzar_Click_Err
    
    Dim i As Long
    
    cmdComenzar.Enabled = False
    cmdCerrar.Enabled = False
    
    If FileExist(DirIndex & "\minimap.bin", vbNormal) Then Kill DirIndex & "\minimap.bin"
    
    Shape1.Width = 0
    
    Open DirIndex & "\minimap.bin" For Binary Access Write As #1
    
        Seek #1, 1
        
        For i = 1 To grhCount
            If GrhData(i).active = True Then
                Picture1.Cls
                Picture2.Cls
                lblstatus.Caption = "Cargando grafico " & i & "/" & grhCount
                Shape1.Width = ((i / 100) / (grhCount / 100)) * 189
                Put #1, , Grh_GetColor(i)
            End If
            DoEvents
        Next i
        
    Close #1
    
    Kill App.Path & "\temp\*.*"
    
    lblstatus = "Finalizado!"
    
    MsgBox "Finalizado!"
    
    cmdComenzar.Enabled = True
    cmdCerrar.Enabled = True
    
    Exit Sub

cmdComenzar_Click_Err:

    Call RegistrarError(Err.Number, Err.Description, "frmMinimapa.cmdComenzar_Click", Erl)
    Resume Next
    
End Sub

