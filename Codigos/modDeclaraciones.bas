Attribute VB_Name = "modDeclaraciones"
Option Explicit

Public isList As Boolean

Public ParticleIndex As Integer

Public DataChanged As Boolean

Public ReferenciaPJ As Boolean

Public Ocupado As Boolean

Public Play    As Boolean

'Control
Public prgRun As Boolean 'When true the program ends

'Caminata fluida
Public Movement_Speed        As Single

'Direcciones
Public Enum E_Heading

    nada = 0
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum

Public Type tColor

    R As Byte
    G As Byte
    B As Byte

End Type

'Colores
Public Const MAXCOLORES As Byte = 56
Public ColoresPJ(0 To MAXCOLORES) As tColor

Public Const SW_NORMAL = 1

Public Declare Function ShellExecute _
               Lib "shell32.dll" _
               Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring _
               Lib "kernel32" _
               Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                   ByVal lpKeyname As Any, _
                                                   ByVal lpString As String, _
                                                   ByVal lpFileName As String) As Long

Public Declare Function getprivateprofilestring _
               Lib "kernel32" _
               Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                 ByVal lpKeyname As Any, _
                                                 ByVal lpdefault As String, _
                                                 ByVal lpreturnedstring As String, _
                                                 ByVal nSize As Long, _
                                                 ByVal lpFileName As String) As Long
                                                 
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
