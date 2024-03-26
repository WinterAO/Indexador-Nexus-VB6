VERSION 5.00
Begin VB.Form frmIndices 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor de Indices"
   ClientHeight    =   3525
   ClientLeft      =   15495
   ClientTop       =   10530
   ClientWidth     =   7410
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
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   235
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   494
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkAutoColocar 
      BackColor       =   &H00404040&
      Caption         =   "Auto colocar bloqueo"
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4770
      TabIndex        =   16
      Top             =   2340
      Width           =   2475
   End
   Begin VB.ComboBox cFiltro 
      BackColor       =   &H80000012&
      ForeColor       =   &H80000014&
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   2610
      Width           =   3765
   End
   Begin Indexador_Nexus.lvButtons_H LvBGuardar 
      Height          =   405
      Left            =   4890
      TabIndex        =   13
      Top             =   2730
      Width           =   2115
      _ExtentX        =   3731
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
   Begin VB.TextBox txtCapa 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4755
      TabIndex        =   10
      Top             =   1920
      Width           =   2475
   End
   Begin VB.TextBox txtAlto 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4755
      TabIndex        =   8
      Top             =   1530
      Width           =   2475
   End
   Begin VB.TextBox txtAncho 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4770
      TabIndex        =   6
      Top             =   1140
      Width           =   2475
   End
   Begin VB.TextBox txtGrh 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4770
      TabIndex        =   4
      Top             =   750
      Width           =   2475
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4770
      TabIndex        =   2
      Top             =   360
      Width           =   2475
   End
   Begin Indexador_Nexus.LynxGrid LynxIndices 
      Height          =   2445
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4313
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   5460819
      BackColorBkg    =   5460819
      BackColorEdit   =   14737632
      BackColorSel    =   12937777
      ForeColor       =   12632256
      ForeColorSel    =   8438015
      BackColorEvenRows=   3158064
      CustomColorFrom =   4210752
      CustomColorTo   =   8421504
      GridColor       =   14737632
      FocusRectColor  =   9895934
      GridLines       =   2
      ThemeColor      =   5
      ScrollBars      =   1
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      HotHeaderTracking=   0   'False
   End
   Begin Indexador_Nexus.lvButtons_H LvBNuevo 
      Height          =   405
      Left            =   1770
      TabIndex        =   11
      Top             =   3030
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   714
      Caption         =   "Nuevo"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   -2147483633
      LockHover       =   1
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   8454016
   End
   Begin Indexador_Nexus.lvButtons_H LvBBorrar 
      Height          =   405
      Left            =   150
      TabIndex        =   12
      Top             =   3030
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   714
      Caption         =   "Borrar"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
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
      cBack           =   8421631
   End
   Begin VB.Label lblBlock 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Block:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4290
      TabIndex        =   15
      Top             =   2340
      Width           =   615
   End
   Begin VB.Label lblCapa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Capa:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4260
      TabIndex        =   9
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblAlto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alto:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4350
      TabIndex        =   7
      Top             =   1530
      Width           =   615
   End
   Begin VB.Label lblAncho 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ancho:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4200
      TabIndex        =   5
      Top             =   1140
      Width           =   615
   End
   Begin VB.Label lblGrh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grh:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4380
      TabIndex        =   3
      Top             =   780
      Width           =   315
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4110
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmIndices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nSup As Long

Private Sub cFiltro_KeyPress(KeyAscii As Integer)
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2021
    '*************************************************

    If KeyAscii = 13 Then
        Call Filtrar
    End If
End Sub

Private Sub Filtrar()
    '*************************************************
    'Author: Lorwik
    'Last modified: 27/04/2021
    '*************************************************

    Dim vDatos As String

    Dim i      As Integer

    Dim j      As Integer

    Dim K      As Long
    
    If cFiltro.ListCount > 5 Then cFiltro.RemoveItem 0
    
    cFiltro.AddItem cFiltro.Text
    
    With LynxIndices
    
        Call AddtoRichTextBox(frmMain.RichConsola, "Buscando indice...", 0, 162, 232)
    
        .Clear
        .Redraw = False
        .Visible = False
    
        For i = 0 To MaxSup
            vDatos = SupData(i).name
        
            For j = 1 To Len(vDatos)
        
                If UCase$(mid$(vDatos & str(i), j, Len(cFiltro.Text))) = UCase$(cFiltro.Text) Or LenB(cFiltro.Text) = 0 Then
                    .AddItem i
                    K = .Rows - 1
                    .CellText(K, 1) = SupData(i).Grh
                    .CellText(K, 2) = vDatos
                    Exit For

                End If
            
            Next
        
        Next i
    
        .Visible = True
        .Redraw = True
        .ColForceFit
    
    End With
    
    DoEvents

End Sub

Private Sub LvBBorrar_Click()

    Dim K      As Long

    If nSup = 0 Then
        MsgBox "No has seleccionado ningun indice"
        Exit Sub
    End If
    
    If MsgBox("¿Seguro que quieres borrar el Indice: " & SupData(nSup).name & "?" & vbCrLf & "Este cambio no tiene vuelta atrás.", vbOKCancel) = vbOK Then
        'Reset it
        With SupData(nSup)
            .Block = False
            .Capa = 0
            .Grh = 0
            .Height = 0
            .Width = 0
            .name = vbNullString
        End With
        
        With LynxIndices
            K = .Row
            .CellText(K, 1) = 0
            .CellText(K, 2) = ""
            .Redraw = True
            .ColForceFit
        End With
        
    End If
End Sub

Private Sub LvBGuardar_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 13/11/2023
    '*************************************************
    
    On Error GoTo ErrorHandler
    
    Dim K As Long
    
    SupData(nSup).name = txtNombre.Text
    SupData(nSup).Grh = Val(txtGrh.Text)
    SupData(nSup).Height = Val(txtAncho.Text)
    SupData(nSup).Width = Val(txtAlto.Text)
    SupData(nSup).Capa = Val(txtCapa.Text)
    SupData(nSup).Block = IIf(chkAutoColocar.value, "1", "0")
        
    Call WriteVar(DirIndices & "Indices.ini", "INIT", "Referencias", MaxSup)
    Call WriteVar(DirIndices & "Indices.ini", "REFERENCIA" & nSup, "Nombre", SupData(nSup).name)
    Call WriteVar(DirIndices & "Indices.ini", "REFERENCIA" & nSup, "GrhIndice", SupData(nSup).Grh)
    Call WriteVar(DirIndices & "Indices.ini", "REFERENCIA" & nSup, "Alto", SupData(nSup).Height)
    Call WriteVar(DirIndices & "Indices.ini", "REFERENCIA" & nSup, "Ancho", SupData(nSup).Width)
    Call WriteVar(DirIndices & "Indices.ini", "REFERENCIA" & nSup, "Capa", SupData(nSup).Capa)
    Call WriteVar(DirIndices & "Indices.ini", "REFERENCIA" & nSup, "Block", SupData(nSup).Block)
        
    With LynxIndices
        K = .Row
        .CellText(K, 1) = SupData(nSup).Grh
        .CellText(K, 2) = SupData(nSup).name
        .Redraw = True
        .ColForceFit

    End With
        
    Call AddtoRichTextBox(frmMain.RichConsola, "Indice " & nSup & " - " & SupData(nSup).name, 0, 255, 0)
    
    Exit Sub
    
ErrorHandler:
    Call AddtoRichTextBox(frmMain.RichConsola, "Error al guardar el indice.", 255, 0, 0)

End Sub

Private Sub LvBNuevo_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 26/03/2024
    '*************************************************
    
    On Error GoTo ErrorHandler
    
    Dim K As Long
    
    MaxSup = MaxSup + 1
    ReDim Preserve SupData(MaxSup) As SupData
    
    'Asignamos valores por defecto
    With SupData(MaxSup)
        .name = "Nueva Superficie"
        .Grh = 1
        .Height = 0
        .Width = 0
        .Capa = 1
        .Block = 0
    End With
        
    With LynxIndices
        .AddItem
        K = .Rows - 1
        .CellText(K, 0) = MaxSup
        .CellText(K, 1) = SupData(MaxSup).Grh
        .CellText(K, 2) = SupData(MaxSup).name
        .Redraw = True
        .ColForceFit

    End With
    
    Exit Sub
    
ErrorHandler:
    Call AddtoRichTextBox(frmMain.RichConsola, "Error al crear un nuevo indice.", 255, 0, 0)
    
End Sub

Private Sub LynxIndices_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 13/11/2023
    '*************************************************
    
    nSup = LynxIndices.CellText(, 0)
    DoEvents
    
    txtNombre.Text = SupData(nSup).name
    txtGrh.Text = SupData(nSup).Grh
    txtAncho.Text = SupData(nSup).Height
    txtAlto.Text = SupData(nSup).Width
    txtCapa.Text = SupData(nSup).Capa
    chkAutoColocar.value = IIf(SupData(nSup).Block, True, False)
    
    Call InitGrh(CurrentGrh, SupData(nSup).Grh)
    isList = False
    
End Sub
