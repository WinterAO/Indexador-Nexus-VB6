Attribute VB_Name = "modDeclaraciones"
Option Explicit

Public Ocupado As Boolean

Public Play    As Boolean

'Direcciones
Public Enum E_Heading

    nada = 0
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum

Public Type tColor

    R As Long
    G As Long
    B As Long

End Type

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
