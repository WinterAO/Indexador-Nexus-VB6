VERSION 5.00
Begin VB.Form frmCodigo 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Codigo"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8130
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
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   -30
      TabIndex        =   1
      Top             =   5370
      Width           =   8175
   End
   Begin VB.TextBox Codigo 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   5370
      Left            =   -30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
    Unload Me
End Sub
