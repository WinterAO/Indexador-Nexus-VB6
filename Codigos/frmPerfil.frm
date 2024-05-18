VERSION 5.00
Begin VB.Form frmPerfil 
   BackColor       =   &H00424242&
   BorderStyle     =   0  'None
   Caption         =   "Selección de Perfil"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
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
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   StartUpPosition =   1  'CenterOwner
   Begin Indexador_Nexus.lvButtons_H LvBBoton 
      Height          =   495
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   2850
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   873
      Caption         =   "&Salir"
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
      cBack           =   8421631
   End
   Begin VB.Frame FraPerfil 
      BackColor       =   &H00535353&
      Caption         =   "Perfil"
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   5115
      Begin Indexador_Nexus.lvButtons_H LvBNuevo 
         Height          =   375
         Index           =   2
         Left            =   3240
         TabIndex        =   7
         Top             =   270
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         Caption         =   "Nuevo"
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
      Begin VB.ComboBox cmbPerfil 
         Height          =   315
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   2895
      End
   End
   Begin VB.Frame FraConfiguraciónDe 
      BackColor       =   &H00535353&
      Caption         =   "Configuración de video"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5115
      Begin VB.CheckBox chkvSync 
         BackColor       =   &H00535353&
         Caption         =   "Activar sincronización vertical"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   930
         Width           =   3015
      End
      Begin VB.ComboBox cmbProcesado 
         Height          =   315
         ItemData        =   "frmPerfil.frx":0000
         Left            =   1920
         List            =   "frmPerfil.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   400
         Width           =   2895
      End
      Begin VB.Label lblModoDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modo de procesado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1470
      End
   End
   Begin Indexador_Nexus.lvButtons_H LvBBoton 
      Height          =   495
      Index           =   1
      Left            =   2940
      TabIndex        =   9
      Top             =   2850
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   873
      Caption         =   "&Continuar"
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
      cBack           =   8454016
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selección de perfil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Perfiles() As String
Private nPerfiles As Byte

Private Sub cmbPerfil_Click()
    Dim tag As String
    tag = cmbPerfil.Text
    If LenB(tag) > 0 And FileExist(profileFile(tag), vbArchive) Then
        Dim v As Byte
        v = CByte(Val(GetVar(profileFile(tag), "CONFIGURACION", "MeMode")))

        v = CByte(Val(GetVar(profileFile(tag), "VIDEO", "VertexProcessingOverride")))
        cmbProcesado.ListIndex = v
    
        v = CByte(Val(GetVar(profileFile(tag), "VIDEO", "LimitarFPS")))
    
        If (v = 1) Then
            chkvSync.value = Checked
            
        Else
            chkvSync.value = Unchecked
            
        End If
           
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Dim i As Byte
    Dim lastProfile As Byte
        
    nPerfiles = Val(GetVar(profilesFile, "INIT", "profiles"))
    lastProfile = Val(GetVar(profilesFile, "INIT", "lastProfile"))
    
    cmbPerfil.Clear
    
    For i = 1 To nPerfiles
        Dim tag As String
        tag = GetVar(profilesFile, "PROFILE" & i, "name")
        If LenB(tag) > 0 And FileExist(profileFile(tag), vbArchive) Then
            cmbPerfil.AddItem (tag)
        End If
    Next i
    
    ReDim Perfiles(1 To nPerfiles) As String
    
    If lastProfile <= nPerfiles And lastProfile > 0 Then
        cmbPerfil.ListIndex = lastProfile - 1
    End If
    
    cmbProcesado.ListIndex = ClientSetup.OverrideVertexProcess
    
    If ClientSetup.LimiteFPS Then
        chkvSync.value = Checked
        
    Else
        chkvSync.value = Unchecked
        
    End If
End Sub

Private Sub chkvSync_Click()
    If chkvSync.value = Checked Then
        ClientSetup.LimiteFPS = True
        
    Else
        ClientSetup.LimiteFPS = False
        
    End If
End Sub

Private Sub LvBBoton_Click(Index As Integer)

    Select Case Index
    
        Case 0 'Salir
            Call SimpleRegistrarError("Seleccion de modo cancelador, saliendo del Indexador Nexus.")
            End
        
        Case 1
        
            If cmbPerfil.ListIndex < 0 Then
                MsgBox "¡No se ha seleccionado ningún perfil!", vbCritical
                Exit Sub
            End If
        
            ProfileTag = cmbPerfil.List(cmbPerfil.ListIndex)
        
            ModoElegido = True
            
            ClientSetup.OverrideVertexProcess = cmbProcesado.ListIndex
            
            Call WriteVar(profileFile(ProfileTag), "VIDEO", "VertexProcessingOverride", CByte(ClientSetup.OverrideVertexProcess))
            Call WriteVar(profileFile(ProfileTag), "VIDEO", "LimitarFPS", IIf(ClientSetup.LimiteFPS, "1", "0"))
            ' Guarda el índice del perfil seleccionado en "lastProfile"
            Call WriteVar(profilesFile, "INIT", "lastProfile", cmbPerfil.ListIndex + 1)
            
            Unload Me
    End Select
End Sub

Private Sub LvBNuevo_Click(Index As Integer)
    ProfileTag = InputBox("Introduce el nombre para el perfil.")
    
    If ProfileTag = vbNullString Then Exit Sub
    
    Call WriteVar(profilesFile, "INIT", "profiles", nPerfiles + 1)
    Call WriteVar(profilesFile, "PROFILE" & (nPerfiles + 1), "name", ProfileTag)
    
    Call WriteVar(profileFile(ProfileTag), "RUTAS", "DirClient", "")
    Call WriteVar(profileFile(ProfileTag), "RUTAS", "DirExport", "")
    Call WriteVar(profileFile(ProfileTag), "RUTAS", "DirIndex", "")
    Call WriteVar(profileFile(ProfileTag), "RUTAS", "DirIndices", "")

    ReDim Perfiles(1 To nPerfiles) As String
    
    cmbPerfil.AddItem (ProfileTag)
End Sub
