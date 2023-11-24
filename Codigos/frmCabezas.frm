VERSION 5.00
Begin VB.Form frmCabezas 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cabezas"
   ClientHeight    =   3015
   ClientLeft      =   6930
   ClientTop       =   8145
   ClientWidth     =   5295
   ClipControls    =   0   'False
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
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtY 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Top             =   900
      Width           =   1125
   End
   Begin VB.TextBox txtX 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   570
      Width           =   1125
   End
   Begin VB.TextBox txtNGrafico 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   1125
   End
   Begin VB.ListBox ListaHead 
      Appearance      =   0  'Flat
      BackColor       =   &H00535353&
      ForeColor       =   &H0000FF00&
      Height          =   2370
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   2700
   End
   Begin Indexador_Nexus.lvButtons_H LvBGuardar 
      Height          =   405
      Left            =   3690
      TabIndex        =   7
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
      Top             =   2520
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
      Top             =   2520
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3600
      TabIndex        =   6
      Top             =   930
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3600
      TabIndex        =   4
      Top             =   600
      Width           =   150
   End
   Begin VB.Label lblNGrafico 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Grafico:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2970
      TabIndex        =   2
      Top             =   270
      Width           =   795
   End
End
Attribute VB_Name = "frmCabezas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListaHead_Click()
'**********************************
'Autor: Lorwik
'Fecha: ??
'**********************************

    txtNgrafico.Text = heads(ListaHead.Text).Texture
    txtX.Text = heads(ListaHead.Text).startX
    txtY.Text = heads(ListaHead.Text).startY
    
    isList = False
    Particle_Group_Remove_All
End Sub

Private Sub LvBBorrar_Click()
'**********************************
'Autor: Lorwik
'Fecha: ??
'**********************************

    If MsgBox("¿Seguro que quieres borrar la cabeza seleccionada?" & vbCrLf & "Este cambio no tiene vuelta atrás.", vbOKCancel) = vbOK Then

        heads(ListaHead.Text).Texture = 0
        heads(ListaHead.Text).startX = 0
        heads(ListaHead.Text).startY = 0
        heads(ListaHead.Text).Std = 0
        
        If Val(ListaHead.Text) >= NumHeads Then
            NumHeads = NumHeads - 1
            ReDim Preserve heads(0 To NumHeads) As tHead
            ListaHead.RemoveItem Val(ListaHead.Text) - 1
        End If
        
    End If
   
End Sub

Private Sub LvBGuardar_Click()
'**********************************
'Autor: Lorwik
'Fecha: 24/11/2023
'**********************************

    With heads(Val(ListaHead.Text))
        .Std = 1
        .Texture = Val(txtNgrafico.Text)
        .startX = Val(txtX.Text)
        .startY = Val(txtY.Text)
    
    End With
    
    Call AddtoRichTextBox(frmMain.RichConsola, "La cabeza " & Val(ListaHead.Text) & " se ha guardado.", 0, 255, 0)
    
End Sub

Private Sub LvBNuevo_Click()
'**********************************
'Autor: Lorwik
'Fecha: ??
'**********************************

    NumHeads = NumHeads + 1

    'Resize array
    ReDim Preserve heads(0 To NumHeads) As tHead
            
    heads(NumHeads).Std = 0
    heads(NumHeads).Texture = 0
    heads(NumHeads).startX = 0
    heads(NumHeads).startY = 0
            
    ListaHead.AddItem NumHeads
    
End Sub
