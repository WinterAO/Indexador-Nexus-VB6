VERSION 5.00
Begin VB.Form frmCuerpos 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuerpos"
   ClientHeight    =   3045
   ClientLeft      =   6930
   ClientTop       =   14445
   ClientWidth     =   5370
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
   ScaleHeight     =   203
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   358
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtY 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3930
      TabIndex        =   14
      Top             =   2130
      Width           =   1125
   End
   Begin VB.TextBox txtX 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3930
      TabIndex        =   12
      Top             =   1740
      Width           =   1125
   End
   Begin VB.TextBox txtOeste 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3930
      TabIndex        =   9
      Top             =   1350
      Width           =   1125
   End
   Begin VB.TextBox txtEste 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3930
      TabIndex        =   7
      Top             =   960
      Width           =   1125
   End
   Begin VB.TextBox txtSur 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3930
      TabIndex        =   5
      Top             =   570
      Width           =   1125
   End
   Begin VB.TextBox txtNorte 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3930
      TabIndex        =   3
      Top             =   180
      Width           =   1125
   End
   Begin VB.ListBox ListaCuerpos 
      Appearance      =   0  'Flat
      BackColor       =   &H00535353&
      ForeColor       =   &H0000FF00&
      Height          =   2370
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   2700
   End
   Begin Indexador_Nexus.lvButtons_H LvBNuevo 
      Height          =   405
      Left            =   1200
      TabIndex        =   1
      Top             =   2580
      Width           =   1545
      _extentx        =   2725
      _extenty        =   714
      caption         =   "Nuevo"
      capalign        =   2
      backstyle       =   2
      shape           =   1
      font            =   "frmCuerpos.frx":0000
      cbhover         =   -2147483633
      lockhover       =   1
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   8454016
   End
   Begin Indexador_Nexus.lvButtons_H LvBBorrar 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   2580
      Width           =   1485
      _extentx        =   2619
      _extenty        =   714
      caption         =   "Borrar"
      capalign        =   2
      backstyle       =   2
      shape           =   2
      font            =   "frmCuerpos.frx":0028
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   8421631
   End
   Begin Indexador_Nexus.lvButtons_H LvBGuardar 
      Height          =   405
      Left            =   3420
      TabIndex        =   11
      Top             =   2550
      Width           =   1335
      _extentx        =   2355
      _extenty        =   714
      caption         =   "Guardar"
      capalign        =   2
      backstyle       =   2
      font            =   "frmCuerpos.frx":0050
      cgradient       =   0
      mode            =   0
      value           =   0   'False
      cback           =   -2147483633
   End
   Begin VB.Label lblOffsetY 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Offset Y:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3060
      TabIndex        =   15
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label lblOffsetX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Offset X:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3060
      TabIndex        =   13
      Top             =   1770
      Width           =   795
   End
   Begin VB.Label lblOeste 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oeste:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3060
      TabIndex        =   10
      Top             =   1380
      Width           =   795
   End
   Begin VB.Label lblEste 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Este:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3060
      TabIndex        =   8
      Top             =   990
      Width           =   795
   End
   Begin VB.Label lblSur 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sur:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3060
      TabIndex        =   6
      Top             =   600
      Width           =   795
   End
   Begin VB.Label lblNorte 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Norte:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3060
      TabIndex        =   4
      Top             =   210
      Width           =   795
   End
End
Attribute VB_Name = "frmCuerpos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListaCuerpos_Click()
    Dim nGrh As Long

    nGrh = BodyData(ListaCuerpos.Text).Walk(3).GrhIndex
    
    DoEvents
    Call InitGrh(CurrentGrh, nGrh)
    
    txtNorte.Text = BodyData(ListaCuerpos.Text).Walk(0).GrhIndex
    txtEste.Text = BodyData(ListaCuerpos.Text).Walk(2).GrhIndex
    txtSur.Text = BodyData(ListaCuerpos.Text).Walk(1).GrhIndex
    txtOeste.Text = BodyData(ListaCuerpos.Text).Walk(3).GrhIndex
    
    isList = False
    Particle_Group_Remove_All
End Sub

Private Sub LvBBorrar_Click()
    Dim i As Byte
    
    If MsgBox("¿Seguro que quieres borrar el cuerpo seleccionado?" & vbCrLf & "Este cambio no tiene vuelta atrás.", vbOKCancel) = vbOK Then

        For i = 1 To 4
            BodyData(ListaCuerpos.Text).Walk(i).GrhIndex = 0
            BodyData(ListaCuerpos.Text).Walk(i).Started = 0
        Next i
        
        BodyData(ListaCuerpos.Text).HeadOffset.x = 0
        BodyData(ListaCuerpos.Text).HeadOffset.y = 0
        
        If Val(ListaCuerpos.Text) >= NumCuerpos Then
            NumCuerpos = NumCuerpos - 1
            ReDim Preserve BodyData(0 To NumCuerpos) As BodyData
            ListaCuerpos.RemoveItem Val(ListaCuerpos.Text) - 1
        End If
        
    End If
End Sub


Private Sub LvBNuevo_Click()
    Dim i As Byte
    
    NumCuerpos = NumCuerpos + 1

    'Resize array
    ReDim Preserve BodyData(0 To NumCuerpos) As BodyData
            
    For i = 1 To 4
        BodyData(NumCuerpos).Walk(i).GrhIndex = 0
    Next i
    
    BodyData(NumCuerpos).HeadOffset.x = 0
    BodyData(NumCuerpos).HeadOffset.y = 0
            
    ListaCuerpos.AddItem NumCuerpos
    
End Sub

