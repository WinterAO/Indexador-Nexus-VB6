VERSION 5.00
Begin VB.Form frmAdaptador 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Adaptador"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8040
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Indexador_Nexus.lvButtons_H cmdAdaptar 
      Height          =   315
      Left            =   1890
      TabIndex        =   10
      Top             =   3090
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   556
      Caption         =   "Adaptar"
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
   Begin VB.TextBox TxtNumAnimaciones 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   6180
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtGrafico 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   3870
      TabIndex        =   7
      Top             =   2610
      Width           =   1095
   End
   Begin VB.TextBox txtPos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1710
      TabIndex        =   5
      Top             =   2610
      Width           =   1185
   End
   Begin VB.TextBox txtAdaptado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3540
      Width           =   7815
   End
   Begin VB.TextBox txtOriginal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   330
      Width           =   7815
   End
   Begin VB.Label lblAnimaciones 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frames:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   5400
      TabIndex        =   9
      Top             =   2700
      Width           =   675
   End
   Begin VB.Label lblGrafico 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grafico:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3150
      TabIndex        =   6
      Top             =   2670
      Width           =   645
   End
   Begin VB.Label lblPrimeraPosición 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Primera posición:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   2670
      Width           =   1455
   End
   Begin VB.Label lblAdaptado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adaptado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   3240
      Width           =   825
   End
   Begin VB.Label lblGrhOriginal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grh Original"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   990
   End
End
Attribute VB_Name = "frmAdaptador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdaptar_Click()
'    Dim Lineas As String, i As Long, j As Long
'    Dim resultado As String
'    Dim Contador As Long
'    Dim lineaGrh As String
'    Dim Fr As Integer
'    Dim tmp As String
'
'    'Debe ingresar almenos un numero
'    If txtPos.Text = "" Then
'        MsgBox ("El valor de la posición es nulo.")
'        Exit Sub
'    End If
'
'    'Debe ingresar almenos un numero
'    If txtGrafico.Text = "" Then
'        MsgBox ("El valor del grafico es nulo.")
'        Exit Sub
'    End If
'
'
'    Contador = txtPos.Text
'
'    'Separamos todas las lineas
'    Lineas = Split(txtOriginal.Text, vbCrLf)
'
'    For i = LBound(Lineas) To UBound(Lineas)
'
'        lineaGrh = ReadField(2, Lineas(i), 61) 'Linea apartir del GrhNº
'        Fr = ReadField(1, lineaGrh, 45) 'Numero del frame
'
'        If Fr = 1 Then
'            'Recortamos y reemplazamos
'            resultado = resultado & "Grh" & Contador & "=" & Fr & "-" & txtGrafico.Text & "-" & ReadField(3, lineaGrh, 45) _
'            & "-" & ReadField(4, lineaGrh, 45) & "-" & ReadField(5, lineaGrh, 45) & "-" & ReadField(6, lineaGrh, 45) & vbCrLf
'
'        Else '¿Es una animacion?
'
'            If TxtNumAnimaciones.Text = "" Then
'                MsgBox ("Hay animaciones y no se especifico el numero de estas.")
'                Exit Sub
'            End If
'
'            tmp = "Grh" & Contador & "=" & Fr
'
'            For j = LBound(Lineas) To UBound(Lineas) - TxtNumAnimaciones.Text
'                tmp = tmp + "-" & j
'            Next j
'
'            resultado = resultado + tmp & "-" & ReadField(Fr + 2, lineaGrh, 45)
'
'        End If
'
'        'Aumentamos en 1 el contador
'        Contador = Contador + 1
'    Next i
'
'    txtAdaptado.Text = resultado
End Sub

