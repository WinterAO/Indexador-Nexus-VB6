VERSION 5.00
Begin VB.Form frmCuerpos 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuerpos"
   ClientHeight    =   3015
   ClientLeft      =   6930
   ClientTop       =   14445
   ClientWidth     =   4560
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
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox ListaCuerpos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   2760
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   2700
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
    
    isList = False
    Particle_Group_Remove_All
End Sub
