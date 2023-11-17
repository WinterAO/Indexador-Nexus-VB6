VERSION 5.00
Begin VB.Form frmCascos 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cascos"
   ClientHeight    =   3015
   ClientLeft      =   6930
   ClientTop       =   11505
   ClientWidth     =   5310
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
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   354
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNGrafico 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3780
      TabIndex        =   3
      Top             =   240
      Width           =   1125
   End
   Begin VB.TextBox txtX 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3780
      TabIndex        =   2
      Top             =   570
      Width           =   1125
   End
   Begin VB.TextBox txtY 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3780
      TabIndex        =   1
      Top             =   900
      Width           =   1125
   End
   Begin VB.ListBox ListaCascos 
      Appearance      =   0  'Flat
      BackColor       =   &H00535353&
      ForeColor       =   &H0000FF00&
      Height          =   2370
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2700
   End
   Begin Indexador_Nexus.lvButtons_H LvBGuardar 
      Height          =   405
      Left            =   3630
      TabIndex        =   4
      Top             =   1530
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
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
   Begin Indexador_Nexus.lvButtons_H LvBNuevo 
      Height          =   405
      Left            =   1170
      TabIndex        =   8
      Top             =   2550
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
      Left            =   90
      TabIndex        =   9
      Top             =   2550
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
   Begin VB.Label lblNGrafico 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Grafico:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2910
      TabIndex        =   7
      Top             =   270
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3540
      TabIndex        =   6
      Top             =   600
      Width           =   150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3540
      TabIndex        =   5
      Top             =   930
      Width           =   150
   End
End
Attribute VB_Name = "frmCascos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListaCascos_Click()

    txtNgrafico.Text = Cascos(ListaCascos.Text).Texture
    txtX.Text = Cascos(ListaCascos.Text).startX
    txtY.Text = Cascos(ListaCascos.Text).startY
    
    isList = False
    Particle_Group_Remove_All
End Sub

Private Sub LvBBorrar_Click()

    If MsgBox("¿Seguro que quieres borrar el casco seleccionada?" & vbCrLf & "Este cambio no tiene vuelta atrás.", vbOKCancel) = vbOK Then

        Cascos(ListaCascos.Text).Texture = 0
        Cascos(ListaCascos.Text).startX = 0
        Cascos(ListaCascos.Text).startY = 0
        Cascos(ListaCascos.Text).Std = 0
        
        If Val(ListaCascos.Text) >= NumCascos Then
            NumCascos = NumCascos - 1
            ReDim Preserve Cascos(0 To NumCascos) As tHead
            ListaCascos.RemoveItem Val(ListaCascos.Text) - 1
        End If
        
    End If
   
End Sub

Private Sub LvBNuevo_Click()
    
    NumCascos = NumCascos + 1

    'Resize array
    ReDim Preserve Cascos(0 To NumCascos) As tHead
            
    Cascos(NumCascos).Std = 0
    Cascos(NumCascos).Texture = 0
    Cascos(NumCascos).startX = 0
    Cascos(NumCascos).startY = 0
            
    ListaCascos.AddItem NumCascos
    
End Sub

