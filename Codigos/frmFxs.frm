VERSION 5.00
Begin VB.Form frmFxs 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fx's"
   ClientHeight    =   3195
   ClientLeft      =   26970
   ClientTop       =   8565
   ClientWidth     =   5820
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
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3900
      TabIndex        =   7
      Top             =   2130
      Width           =   1125
   End
   Begin VB.TextBox txtSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3900
      TabIndex        =   5
      Top             =   2460
      Width           =   1125
   End
   Begin VB.ListBox listAnimacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00535353&
      ForeColor       =   &H0000FF00&
      Height          =   1590
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   2700
   End
   Begin VB.ListBox ListaFxs 
      Appearance      =   0  'Flat
      BackColor       =   &H00535353&
      ForeColor       =   &H0000FF00&
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   2700
   End
   Begin Indexador_Nexus.lvButtons_H LvBNuevo 
      Height          =   405
      Left            =   1230
      TabIndex        =   1
      Top             =   2580
      Width           =   1545
      _ExtentX        =   2725
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
   Begin Indexador_Nexus.lvButtons_H LvBBorrar 
      Height          =   405
      Left            =   150
      TabIndex        =   2
      Top             =   2580
      Width           =   1485
      _ExtentX        =   2619
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
   Begin Indexador_Nexus.lvButtons_H LvBGuardar 
      Height          =   315
      Left            =   3690
      TabIndex        =   3
      Top             =   2820
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "Guardar"
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
   Begin Indexador_Nexus.lvButtons_H LvBAnadir 
      Height          =   315
      Left            =   4380
      TabIndex        =   9
      Top             =   1770
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Caption         =   "Añadir"
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
      cBack           =   8454016
   End
   Begin Indexador_Nexus.lvButtons_H LvBDelete 
      Height          =   315
      Left            =   3090
      TabIndex        =   10
      Top             =   1770
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Caption         =   "Eliminar"
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
      cBack           =   8421631
   End
   Begin VB.Label lblFrame 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frame:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3120
      TabIndex        =   8
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label lblVelocidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Velocidad:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Top             =   2490
      Width           =   795
   End
End
Attribute VB_Name = "frmFxs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nGrh As Long

Private Sub ListaFxs_Click()
'**********************************
'Autor: Lorwik
'Fecha: ??
'**********************************

    Dim i As Byte

    nGrh = FxData(ListaFxs.Text).Animacion
    
    For i = 1 To GrhData(nGrh).NumFrames
        listAnimacion.AddItem GrhData(nGrh).Frames(i)
    Next i
    
    txtSpeed.Text = GrhData(nGrh).speed

    DoEvents
    Call InitGrh(CurrentGrh, nGrh)
    
    Renderizando = eRender.eFXs
    Particle_Group_Remove_All
End Sub

Private Sub LvBBorrar_Click()
'**********************************
'Autor: Lorwik
'Fecha: ??
'**********************************

    If MsgBox("¿Seguro que quieres borrar el Fx seleccionado?" & vbCrLf & "Este cambio no tiene vuelta atrás.", vbOKCancel) = vbOK Then

        FxData(ListaFxs.Text).Animacion = 0
        FxData(ListaFxs.Text).OffsetX = 0
        FxData(ListaFxs.Text).OffsetY = 0
        
        If Val(ListaFxs.Text) >= NumFxs Then
            NumFxs = NumFxs - 1
            ReDim Preserve FxData(0 To NumFxs) As tIndiceFx
            ListaFxs.RemoveItem Val(ListaFxs.Text) - 1

        End If
        
    End If

End Sub

Private Sub LvBGuardar_Click()
'**********************************
'Autor: Lorwik
'Fecha: 24/11/2023
'**********************************

    Dim i As Byte
    
    With GrhData(nGrh)
    
        .NumFrames = listAnimacion.ListCount
    
        ReDim .Frames(1 To .NumFrames)
    
        For i = 1 To .NumFrames
            .Frames(i) = listAnimacion.List(i)
        Next i

        .speed = Val(txtSpeed.Text)

    End With
    
End Sub
