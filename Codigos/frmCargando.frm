VERSION 5.00
Begin VB.Form frmCargando 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Iniciando Indexador Nexus"
   ClientHeight    =   555
   ClientLeft      =   15960
   ClientTop       =   11145
   ClientWidth     =   7260
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
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   486
      Y1              =   36
      Y2              =   36
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   484
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
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
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   6795
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

