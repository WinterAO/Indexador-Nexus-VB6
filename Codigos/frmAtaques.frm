VERSION 5.00
Begin VB.Form frmAtaques 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ataques"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
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
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ListaAtaques 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   2760
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   2700
   End
End
Attribute VB_Name = "frmAtaques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListaAtaques_Click()
    Renderizando = eRender.eAtaques
    Particle_Group_Remove_All
End Sub
