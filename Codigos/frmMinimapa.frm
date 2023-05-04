VERSION 5.00
Begin VB.Form frmMinimapa 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Minimap Color Finder"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3090
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   206
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   525
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
      Width           =   2895
   End
   Begin VB.CommandButton cmdComenzar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Comenzar!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   120
      ScaleHeight     =   2880
      ScaleWidth      =   2880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2880
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
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   585
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4410
      Width           =   585
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
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Graficos.ind cargado!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1170
      Width           =   2880
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
      Left            =   270
      TabIndex        =   4
      Top             =   4500
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
    Dim i As Long
    
    cmdComenzar.Enabled = False
    cmdCerrar.Enabled = False
    
    If FileExist(DirIndex & "\Minimap.bin", vbNormal) Then Kill DirIndex & "\minimap.bin"
    
    Shape1.Width = 0
    
    Open DirIndex & "\MiniMap.bin" For Binary Access Write As #1
    
        Seek #1, 1
        
        For i = 1 To grhCount
            If GrhData(i).Active = True Then
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
End Sub

