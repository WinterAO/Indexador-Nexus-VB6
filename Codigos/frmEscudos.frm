VERSION 5.00
Begin VB.Form frmEscudos 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Escudos"
   ClientHeight    =   2835
   ClientLeft      =   26970
   ClientTop       =   4800
   ClientWidth     =   5280
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
   ScaleHeight     =   189
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtOeste 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3990
      TabIndex        =   6
      Top             =   1320
      Width           =   1125
   End
   Begin VB.TextBox txtEste 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3990
      TabIndex        =   5
      Top             =   930
      Width           =   1125
   End
   Begin VB.TextBox txtSur 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3990
      TabIndex        =   4
      Top             =   540
      Width           =   1125
   End
   Begin VB.TextBox txtNorte 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3990
      TabIndex        =   3
      Top             =   150
      Width           =   1125
   End
   Begin VB.ListBox ListaEscudos 
      Appearance      =   0  'Flat
      BackColor       =   &H00535353&
      ForeColor       =   &H0000FF00&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   2700
   End
   Begin Indexador_Nexus.lvButtons_H LvBNuevo 
      Height          =   405
      Left            =   1260
      TabIndex        =   1
      Top             =   2340
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
      Top             =   2340
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
      Height          =   405
      Left            =   3360
      TabIndex        =   7
      Top             =   1800
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
   Begin VB.Label lblOeste 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oeste:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3120
      TabIndex        =   11
      Top             =   1350
      Width           =   795
   End
   Begin VB.Label lblEste 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Este:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3120
      TabIndex        =   10
      Top             =   960
      Width           =   795
   End
   Begin VB.Label lblSur 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sur:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3120
      TabIndex        =   9
      Top             =   570
      Width           =   795
   End
   Begin VB.Label lblNorte 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Norte:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3120
      TabIndex        =   8
      Top             =   180
      Width           =   795
   End
End
Attribute VB_Name = "frmEscudos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListaEscudos_Click()
    Dim nGrh As Long

    nGrh = ShieldAnimData(ListaEscudos.Text).ShieldWalk(3).GrhIndex
    
    DoEvents
    Call InitGrh(CurrentGrh, nGrh)
    
    isList = False
    Particle_Group_Remove_All
End Sub

Private Sub LvBBorrar_Click()
'**********************************
'Autor: Lorwik
'Fecha: 24/11/2023
'**********************************

    Dim i As Byte
    
    If MsgBox("¿Seguro que quieres borrar el escudo seleccionado?" & vbCrLf & "Este cambio no tiene vuelta atrás.", vbOKCancel) = vbOK Then

        For i = 1 To 4
            ShieldAnimData(ListaEscudos.Text).ShieldWalk(i).GrhIndex = 0
            ShieldAnimData(ListaEscudos.Text).ShieldWalk(i).Started = 0
        Next i
        
        If Val(ListaEscudos.Text) >= NumEscudosAnims Then
            NumEscudosAnims = NumEscudosAnims - 1
            ReDim Preserve ShieldAnimData(0 To NumEscudosAnims) As ShieldAnimData
            ListaEscudos.RemoveItem Val(ListaEscudos.Text) - 1
        End If
        
    End If
End Sub

Private Sub LvBGuardar_Click()
'**********************************
'Autor: Lorwik
'Fecha: 24/11/2023
'Descripción: Guardado del arma que se esta editando
'**********************************

    With ShieldAnimData(Val(ListaEscudos.Text))
    
        .ShieldWalk(1).GrhIndex = Val(txtNorte.Text)
        .ShieldWalk(2).GrhIndex = Val(txtSur.Text)
        .ShieldWalk(3).GrhIndex = Val(txtOeste.Text)
        .ShieldWalk(4).GrhIndex = Val(txtEste.Text)
    
    End With
    
    Call AddtoRichTextBox(frmMain.RichConsola, "El escudo " & Val(ListaEscudos.Text) & " se ha guardado.", 0, 255, 0)
End Sub

Private Sub LvBNuevo_Click()
'**********************************
'Autor: Lorwik
'Fecha: 24/11/2023
'**********************************

    Dim i As Byte
    
    NumEscudosAnims = NumEscudosAnims + 1

    'Resize array
    ReDim Preserve ShieldAnimData(0 To NumEscudosAnims) As ShieldAnimData
            
    For i = 1 To 4
        ShieldAnimData(NumEscudosAnims).ShieldWalk(i).GrhIndex = 0
    Next i
            
    ListaEscudos.AddItem NumEscudosAnims
    
End Sub
