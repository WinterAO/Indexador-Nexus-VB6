VERSION 5.00
Begin VB.Form frmObj 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor de objetos"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12300
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
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSeCae 
      Caption         =   "Se cae"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   1620
      Width           =   3225
   End
   Begin VB.CheckBox chkAgarrable 
      Caption         =   "Agarrable"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   1350
      Width           =   3225
   End
   Begin VB.ComboBox cTipo 
      BackColor       =   &H80000012&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   900
      Width           =   2355
   End
   Begin VB.TextBox txtNorte 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   2355
   End
   Begin VB.TextBox txtSur 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3960
      TabIndex        =   2
      Top             =   510
      Width           =   2355
   End
   Begin VB.TextBox txtY 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   2070
      Width           =   2355
   End
   Begin VB.ListBox Lista 
      Appearance      =   0  'Flat
      BackColor       =   &H00535353&
      ForeColor       =   &H0000FF00&
      Height          =   8610
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2700
   End
   Begin Indexador_Nexus.lvButtons_H LvBGuardar 
      Height          =   405
      Left            =   3630
      TabIndex        =   4
      Top             =   2580
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
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3090
      TabIndex        =   8
      Top             =   150
      Width           =   615
   End
   Begin VB.Label lbGrhIndex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GrhIndex:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3090
      TabIndex        =   7
      Top             =   540
      Width           =   735
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3090
      TabIndex        =   6
      Top             =   930
      Width           =   360
   End
   Begin VB.Label lblValor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3090
      TabIndex        =   5
      Top             =   2100
      Width           =   420
   End
End
Attribute VB_Name = "frmObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

