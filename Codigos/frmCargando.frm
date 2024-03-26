VERSION 5.00
Begin VB.Form frmCargando 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Iniciando Indexador Nexus"
   ClientHeight    =   4125
   ClientLeft      =   15960
   ClientTop       =   11145
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   275
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picLogo 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3780
      Left            =   0
      ScaleHeight     =   3780
      ScaleWidth      =   7740
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7740
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   512
      Y1              =   274
      Y2              =   274
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   512
      Y1              =   253
      Y2              =   252
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   450
      TabIndex        =   1
      Top             =   3840
      Width           =   6795
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    picLogo.Picture = LoadPicture(App.Path & "\init\NexusEditor.jpg")
End Sub
