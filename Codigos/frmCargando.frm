VERSION 5.00
Begin VB.Form frmCargando 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando..."
      ForeColor       =   &H8000000B&
      Height          =   195
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

