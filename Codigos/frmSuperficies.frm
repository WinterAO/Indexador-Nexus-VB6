VERSION 5.00
Begin VB.Form frmSuperficies 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Indexar Superficies"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   471
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   466
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Indexador_Nexus.lvButtons_H LvBAvanzado 
      Height          =   375
      Left            =   4230
      TabIndex        =   35
      Top             =   6570
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      Caption         =   "Avanzado"
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
   Begin VB.OptionButton OptX 
      BackColor       =   &H00404040&
      Caption         =   "64 x 384"
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
      Index           =   7
      Left            =   5640
      TabIndex        =   32
      Top             =   6150
      Width           =   1185
   End
   Begin VB.OptionButton OptX 
      BackColor       =   &H00404040&
      Caption         =   "288 x 96"
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
      Index           =   6
      Left            =   4080
      TabIndex        =   31
      Top             =   6150
      Width           =   1185
   End
   Begin VB.OptionButton OptX 
      BackColor       =   &H00404040&
      Caption         =   "256 x 64"
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
      Index           =   5
      Left            =   5520
      TabIndex        =   30
      Top             =   4590
      Width           =   1185
   End
   Begin VB.OptionButton OptX 
      BackColor       =   &H00404040&
      Caption         =   "64 x 64"
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
      Index           =   4
      Left            =   4050
      TabIndex        =   29
      Top             =   4620
      Width           =   1185
   End
   Begin VB.OptionButton OptX 
      BackColor       =   &H00404040&
      Caption         =   "96 x 96"
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
      Index           =   3
      Left            =   5580
      TabIndex        =   28
      Top             =   3210
      Width           =   1185
   End
   Begin VB.OptionButton OptX 
      BackColor       =   &H00404040&
      Caption         =   "384 x 128"
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
      Index           =   2
      Left            =   3990
      TabIndex        =   27
      Top             =   3180
      Width           =   1185
   End
   Begin VB.OptionButton OptX 
      BackColor       =   &H00404040&
      Caption         =   "128 x 128"
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
      Index           =   1
      Left            =   5520
      TabIndex        =   26
      Top             =   1560
      Width           =   1185
   End
   Begin VB.OptionButton OptX 
      BackColor       =   &H00404040&
      Caption         =   "512 x 128"
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
      Index           =   0
      Left            =   3960
      TabIndex        =   25
      Top             =   1560
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.PictureBox Muestra 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Index           =   7
      Left            =   6000
      Picture         =   "frmSuperficies.frx":0000
      ScaleHeight     =   1200
      ScaleWidth      =   300
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4920
      Width           =   300
   End
   Begin VB.PictureBox Muestra 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   6
      Left            =   3840
      Picture         =   "frmSuperficies.frx":04CE
      ScaleHeight     =   900
      ScaleWidth      =   1650
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5130
      Width           =   1650
   End
   Begin VB.PictureBox Muestra 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   5
      Left            =   5400
      Picture         =   "frmSuperficies.frx":0E5F
      ScaleHeight     =   750
      ScaleWidth      =   1380
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1380
   End
   Begin VB.PictureBox Muestra 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   4
      Left            =   3990
      Picture         =   "frmSuperficies.frx":1B05
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3600
      Width           =   960
   End
   Begin VB.PictureBox Muestra 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Index           =   3
      Left            =   5430
      Picture         =   "frmSuperficies.frx":271B
      ScaleHeight     =   1200
      ScaleWidth      =   1380
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1380
   End
   Begin VB.PictureBox Muestra 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Index           =   2
      Left            =   3840
      Picture         =   "frmSuperficies.frx":376D
      ScaleHeight     =   1200
      ScaleWidth      =   1380
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1380
   End
   Begin VB.PictureBox Muestra 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Index           =   1
      Left            =   5400
      Picture         =   "frmSuperficies.frx":471B
      ScaleHeight     =   1200
      ScaleWidth      =   1380
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   270
      Width           =   1380
   End
   Begin VB.PictureBox Muestra 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Index           =   0
      Left            =   3840
      Picture         =   "frmSuperficies.frx":5ADA
      ScaleHeight     =   1200
      ScaleWidth      =   1380
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   270
      Width           =   1380
   End
   Begin VB.TextBox txtReferencia 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   3
      Left            =   210
      TabIndex        =   12
      Top             =   4620
      Width           =   3375
   End
   Begin VB.TextBox txtReferencia 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   2
      Left            =   210
      TabIndex        =   10
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox txtReferencia 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   1
      Left            =   210
      TabIndex        =   8
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox txtReferencia 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Index           =   0
      Left            =   210
      TabIndex        =   6
      Top             =   2700
      Width           =   3375
   End
   Begin Indexador_Nexus.lvButtons_H LvBIndexar 
      Height          =   735
      Left            =   180
      TabIndex        =   4
      Top             =   1020
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   1296
      Caption         =   "Indexar"
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
   Begin VB.TextBox txtNgrafico 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2010
      TabIndex        =   3
      Top             =   570
      Width           =   1335
   End
   Begin Indexador_Nexus.lvButtons_H LvBOcultarMenu 
      Height          =   375
      Left            =   9390
      TabIndex        =   36
      Top             =   6570
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
      Caption         =   "Ocultar Menu Avanzado"
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
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   1
      X1              =   472
      X2              =   472
      Y1              =   24
      Y2              =   444
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   0
      X1              =   468
      X2              =   468
      Y1              =   24
      Y2              =   444
   End
   Begin VB.Label lblLineas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Líneas: 0"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   960
      TabIndex        =   34
      Top             =   5970
      Width           =   645
   End
   Begin VB.Label lblHazClick 
      BackStyle       =   0  'Transparent
      Caption         =   "Haz click sobre ""Número de referencias"" o ""Ultimo GRH"" para actualizar sus valores."
      ForeColor       =   &H0000FF00&
      Height          =   570
      Left            =   240
      TabIndex        =   33
      Top             =   6480
      Width           =   3315
   End
   Begin VB.Label lblAlto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alto: 0"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   1320
      TabIndex        =   16
      Top             =   5550
      Width           =   975
   End
   Begin VB.Label lblAncho 
      BackStyle       =   0  'Transparent
      Caption         =   "Ancho: 0"
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   270
      TabIndex        =   15
      Top             =   5550
      Width           =   975
   End
   Begin VB.Label lblNumeroDe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número de referencias: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   240
      TabIndex        =   14
      Top             =   5250
      Width           =   2085
   End
   Begin VB.Label lblNombreDe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la cuarta referencia:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   4350
      Width           =   2595
   End
   Begin VB.Label lblNombreDe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la tercera referencia:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   3690
      Width           =   2670
   End
   Begin VB.Label lblNombreDe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la segunda referencia:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   3090
      Width           =   2775
   End
   Begin VB.Label lblNombreDe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de la primera referencia:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   2430
      Width           =   2685
   End
   Begin VB.Label lblReferenciasDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referencias del World Editor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   180
      TabIndex        =   5
      Top             =   1950
      Width           =   3000
   End
   Begin VB.Label lblNºGrafico 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Grafico"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2010
      TabIndex        =   2
      Top             =   180
      Width           =   1110
   End
   Begin VB.Label lblGrhs 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   300
      TabIndex        =   1
      Top             =   600
      Width           =   1170
   End
   Begin VB.Label lblultimoGRH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Último GRH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   1260
   End
End
Attribute VB_Name = "frmSuperficies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nSupAncho As Byte
Private nSupAlto  As Byte
Private nLineas   As Integer
Dim nSup As Byte

Private Sub Form_Load()
    '*************************************************
    'Autor: Lorwik
    'Fecha: 15/11/2023
    '*************************************************
    
    lblGrhs.Caption = grhCount
    Call OptX_Click(0)
End Sub

Private Sub LvBAvanzado_Click()
    Me.Width = 13620
End Sub

Private Sub LvBIndexar_Click()
    '*************************************************
    'Autor: Lorwik
    'Fecha: 15/11/2023
    'Descripcion: Indexado de superficies
    '*************************************************
    
    Dim i          As Long
    
    Dim tX         As Byte

    Dim tY         As Byte

    Dim XX         As Integer

    Dim YY         As Integer
    
    Dim Ancho      As Integer
    
    Dim alto       As Integer
    
    Dim sinCalculo As Boolean
    
    Dim Count      As Integer
    
    Dim refCount   As Byte
    
    Dim resultado  As String
    
    If Val(txtNgrafico.Text) < 1 Then
        MsgBox "Numero de grafico invalido."
        Exit Sub

    End If

    For i = 0 To nSup - 1

        If txtReferencia(i).Visible Then
            If txtReferencia(i).Text = vbNullString Then
                MsgBox "Introduce el nombre de las referencias."
                Exit Sub

            End If

        End If

    Next i
    
    Call AddtoRichTextBox(frmMain.RichConsola, "Indexando superficies...", 0, 255, 0)
    
    If OptX(6).value Then '280 x 96
    
        Ancho = 280
        alto = 96
        resultado = "Grh" & grhCount + 1 & "=" & txtNgrafico.Text & "-0-0-" & Ancho & "-" & alto & vbCrLf
        sinCalculo = True
        
    ElseIf OptX(7).value Then '64 x 384
        Ancho = 64
        alto = 384
        resultado = "Grh" & grhCount + 1 & "=" & txtNgrafico.Text & "-0-0-" & Ancho & "-" & alto & vbCrLf
        sinCalculo = True
        
    End If
    
    'Insertamos el resultado
    ReDim Preserve GrhData(1 To grhCount + nLineas) As GrhData
    
    If Not sinCalculo Then

        'Calculo de los Grh
        For tX = 1 To nSupAncho
            For tY = 1 To nSupAlto
            
                Count = Count + 1
            
                grhCount = grhCount + 1
                resultado = resultado & "Grh" & grhCount & "=" & txtNgrafico.Text & "-" & XX & "-" & YY & "-32-32" & vbCrLf
        
                With GrhData(grhCount)
                        
                    .FileNum = Val(txtNgrafico.Text)
                    .NumFrames = 1
                    .pixelWidth = Ancho
                    .pixelHeight = alto
                    .sX = 32
                    .sY = 32
                    .active = True
                    
                End With
                
                If nSupAlto <> 0 And nSupAncho <> 0 Then
                    If Count Mod (nSupAncho * nSupAlto) = 0 Then
        
                        MaxSup = MaxSup + 1
                    
                        ReDim Preserve SupData(MaxSup) As SupData
                
                        With SupData(MaxSup)
                    
                            .name = txtReferencia(refCount).Text
                            .Grh = grhCount
                            .Height = nSupAlto
                            .Width = nSupAncho
                            .Capa = 1
                            .Block = False
                    
                        End With
                    
                        refCount = refCount + 1
                
                    End If
                    
                Else
                
                    MaxSup = MaxSup + 1
                    
                    ReDim Preserve SupData(MaxSup) As SupData
                
                    With SupData(MaxSup)
                    
                        .name = txtReferencia(0).Text
                        .Grh = grhCount
                        .Height = nSupAlto
                        .Width = nSupAncho
                        .Capa = 1
                        .Block = False
                    
                    End With

                End If
                
                XX = XX + 32
            Next tY
    
            XX = 0
            YY = YY + 32
        Next tX

    Else
    
        grhCount = grhCount + 1
    
        With GrhData(grhCount)
                        
            .FileNum = Val(txtNgrafico.Text)
            .NumFrames = 1
            .pixelWidth = 0
            .pixelHeight = 0
            .sX = Ancho
            .sY = alto
            .active = True
            
        End With
        
        MaxSup = MaxSup + 1
        
        ReDim Preserve SupData(MaxSup) As SupData
        
        With SupData(MaxSup)
        
            .name = txtReferencia(0).Text
            .Grh = grhCount + nLineas
            .Height = 0
            .Width = 0
            .Capa = 1
            .Block = False
        
        End With

    End If
    
    Call recargarLynxGrh
    
    'Resultado
    Call AddtoRichTextBox(frmMain.RichConsola, "Se añadieron las siguientes " & nLineas & " Grh:", 0, 255, 0)
    Call AddtoRichTextBox(frmMain.RichConsola, resultado, 0, 255, 0)
    
End Sub

Private Sub LvBOcultarMenu_Click()
    Me.Width = 7080
End Sub

Private Sub Muestra_Click(Index As Integer)
    '*************************************************
    'Autor: Lorwik
    'Fecha: 15/11/2023
    '*************************************************
    OptX(Index).value = True
End Sub

Private Sub OptX_Click(Index As Integer)
    '*************************************************
    'Autor: Lorwik
    'Fecha: 15/11/2023
    'Descripcion: Selecciona un tipo de superficie
    '*************************************************
    
    Dim i As Byte
    
    Select Case Index
    
        Case 0
            nSupAncho = 4
            nSupAlto = 4

            For i = 1 To 3
                txtReferencia(i).Visible = True
                lblNombreDe(i).Visible = True
            Next i
            
            nLineas = 64
            nSup = 4
            
        Case 1
            nSupAncho = 4
            nSupAlto = 4

            For i = 1 To 3
                txtReferencia(i).Visible = False
                lblNombreDe(i).Visible = False
            Next i
            
            nLineas = 16
            nSup = 1
            
        Case 2
            nSupAncho = 4
            nSupAlto = 4
            
            txtReferencia(3).Visible = False
            lblNombreDe(3).Visible = False

            For i = 1 To 2
                txtReferencia(i).Visible = True
                lblNombreDe(i).Visible = True
            Next i
            
            nLineas = 47
            nSup = 2
            
        Case 3
            nSupAncho = 3
            nSupAlto = 3

            For i = 1 To 3
                txtReferencia(i).Visible = False
                lblNombreDe(i).Visible = False
            Next i
            
            nLineas = 9
            nSup = 1
            
        Case 4
            nSupAncho = 2
            nSupAlto = 2

            For i = 1 To 3
                txtReferencia(i).Visible = False
                lblNombreDe(i).Visible = False
            Next i
            
            nLineas = 4
            nSup = 1
            
        Case 5
            nSupAncho = 4
            nSupAlto = 2

            For i = 1 To 3
                txtReferencia(i).Visible = True
                lblNombreDe(i).Visible = True
            Next i
            
            nLineas = 16
            nSup = 4
            
        Case 6
            nSupAncho = 1
            nSupAlto = 1
            
            For i = 1 To 3
                txtReferencia(i).Visible = False
                lblNombreDe(i).Visible = False
            Next i
            
            nLineas = 1
            nSup = 1
            
        Case 7
            nSupAncho = 1
            nSupAlto = 1
            
            For i = 1 To 3
                txtReferencia(i).Visible = False
                lblNombreDe(i).Visible = False
            Next i
            
            nLineas = 1
            nSup = 1
    
    End Select
    
    lblAncho.Caption = "Ancho: " & nSupAncho
    lblAlto.Caption = "Alto: " & nSupAlto
    lblLineas.Caption = nLineas

End Sub
