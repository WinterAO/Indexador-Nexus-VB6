Attribute VB_Name = "ModParticulas"
Option Explicit

'--> Current StreamFile <--
Public CurStreamFile As String

Public Sub NuevaParticula()

    Dim Nombre As String
    Dim NewStreamNumber As Integer
    Dim grhlist(0) As Long
    Dim LoopC As Long
    
    'Get name for new stream
    Nombre = InputBox("Por favor inserte un nombre a la particula", "New Stream")
    
    If Nombre = "" Then Exit Sub
    
    'Set new stream #
    NewStreamNumber = frmParticleEditor.ListParticulas.ListCount + 1
    
    'Add stream to combo box
    frmParticleEditor.ListParticulas.AddItem NewStreamNumber & " - " & Nombre
    
    'Add 1 to TotalStreams
    TotalStreams = TotalStreams + 1
    
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'Add stream data to StreamData array
    StreamData(NewStreamNumber).name = frmParticleEditor.name
    StreamData(NewStreamNumber).NumOfParticles = 126
    StreamData(NewStreamNumber).x1 = 0
    StreamData(NewStreamNumber).y1 = 0
    StreamData(NewStreamNumber).x2 = 0
    StreamData(NewStreamNumber).y2 = 0
    StreamData(NewStreamNumber).angle = 0
    StreamData(NewStreamNumber).vecx1 = -20
    StreamData(NewStreamNumber).vecx2 = 20
    StreamData(NewStreamNumber).vecy1 = -20
    StreamData(NewStreamNumber).vecy2 = 20
    StreamData(NewStreamNumber).life1 = 10
    StreamData(NewStreamNumber).life2 = 50
    StreamData(NewStreamNumber).friction = 8
    StreamData(NewStreamNumber).spin_speedL = 0.1
    StreamData(NewStreamNumber).spin_speedH = 0.1
    StreamData(NewStreamNumber).grav_strength = 2
    StreamData(NewStreamNumber).bounce_strength = -5
    StreamData(NewStreamNumber).alphaBlend = 1
    StreamData(NewStreamNumber).gravity = 0
    
    
    'Select the new stream type in the combo box
    frmParticleEditor.ListParticulas.ListIndex = NewStreamNumber - 1

End Sub

Public Sub GuardarParticulas()

    Dim LoopC As Long
    Dim StreamFile As String
    Dim Bypass As Boolean
    Dim RetVal
    CurStreamFile = DirIndex & "Particulas.dat"
    
    If FileExist(CurStreamFile, vbNormal) = True Then
        RetVal = MsgBox("¡El archivo " & CurStreamFile & " ya existe!" & vbCrLf & "¿Deseas sobreescribirlo?", vbYesNoCancel Or vbQuestion)
        If RetVal = vbNo Then
            Bypass = False
        ElseIf RetVal = vbCancel Then
            Exit Sub
        ElseIf RetVal = vbYes Then
            StreamFile = CurStreamFile
            Bypass = True
        End If
    End If
    
    If Bypass = False Then
    
        StreamFile = CurStreamFile
        
        If FileExist(StreamFile, vbNormal) = True Then
            RetVal = MsgBox("¡El archivo " & StreamFile & " ya existe!" & vbCrLf & "¿Desea sobreescribirlo?", vbYesNo Or vbQuestion)
            If RetVal = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    Dim GrhListing As String
    Dim i As Long
    
    'Check for existing data file and kill it
    If FileExist(StreamFile, vbNormal) Then Kill StreamFile
    
    'Write particle data to particle.ini
    WriteVar StreamFile, "INIT", "Total", Val(TotalStreams)
    
    For LoopC = 1 To TotalStreams
        WriteVar StreamFile, Val(LoopC), "Name", StreamData(LoopC).name
        WriteVar StreamFile, Val(LoopC), "NumOfParticles", Val(StreamData(LoopC).NumOfParticles)
        WriteVar StreamFile, Val(LoopC), "X1", Val(StreamData(LoopC).x1)
        WriteVar StreamFile, Val(LoopC), "Y1", Val(StreamData(LoopC).y1)
        WriteVar StreamFile, Val(LoopC), "X2", Val(StreamData(LoopC).x2)
        WriteVar StreamFile, Val(LoopC), "Y2", Val(StreamData(LoopC).y2)
        WriteVar StreamFile, Val(LoopC), "Angle", Val(StreamData(LoopC).angle)
        WriteVar StreamFile, Val(LoopC), "VecX1", Val(StreamData(LoopC).vecx1)
        WriteVar StreamFile, Val(LoopC), "VecX2", Val(StreamData(LoopC).vecx2)
        WriteVar StreamFile, Val(LoopC), "VecY1", Val(StreamData(LoopC).vecy1)
        WriteVar StreamFile, Val(LoopC), "VecY2", Val(StreamData(LoopC).vecy2)
        WriteVar StreamFile, Val(LoopC), "Life1", Val(StreamData(LoopC).life1)
        WriteVar StreamFile, Val(LoopC), "Life2", Val(StreamData(LoopC).life2)
        WriteVar StreamFile, Val(LoopC), "Friction", Val(StreamData(LoopC).friction)
        WriteVar StreamFile, Val(LoopC), "Spin", Val(StreamData(LoopC).spin)
        WriteVar StreamFile, Val(LoopC), "Spin_SpeedL", Val(StreamData(LoopC).spin_speedL)
        WriteVar StreamFile, Val(LoopC), "Spin_SpeedH", Val(StreamData(LoopC).spin_speedH)
        WriteVar StreamFile, Val(LoopC), "Grav_Strength", Val(StreamData(LoopC).grav_strength)
        WriteVar StreamFile, Val(LoopC), "Bounce_Strength", Val(StreamData(LoopC).bounce_strength)
        
        WriteVar StreamFile, Val(LoopC), "AlphaBlend", Val(StreamData(LoopC).alphaBlend)
        WriteVar StreamFile, Val(LoopC), "Gravity", Val(StreamData(LoopC).gravity)
        
        WriteVar StreamFile, Val(LoopC), "XMove", Val(StreamData(LoopC).XMove)
        WriteVar StreamFile, Val(LoopC), "YMove", Val(StreamData(LoopC).YMove)
        WriteVar StreamFile, Val(LoopC), "move_x1", Val(StreamData(LoopC).move_x1)
        WriteVar StreamFile, Val(LoopC), "move_x2", Val(StreamData(LoopC).move_x2)
        WriteVar StreamFile, Val(LoopC), "move_y1", Val(StreamData(LoopC).move_y1)
        WriteVar StreamFile, Val(LoopC), "move_y2", Val(StreamData(LoopC).move_y2)
        WriteVar StreamFile, Val(LoopC), "life_counter", Val(StreamData(LoopC).life_counter)
        WriteVar StreamFile, Val(LoopC), "Speed", str(StreamData(LoopC).speed)
        
        WriteVar StreamFile, Val(LoopC), "NumGrhs", Val(StreamData(LoopC).NumGrhs)
        
        GrhListing = vbNullString
        For i = 1 To StreamData(LoopC).NumGrhs
            GrhListing = GrhListing & StreamData(LoopC).grh_list(i) & ","
        Next i
        
        WriteVar StreamFile, Val(LoopC), "Grh_List", GrhListing
        
        WriteVar StreamFile, Val(LoopC), "ColorSet1", StreamData(LoopC).colortint(0).R & "," & StreamData(LoopC).colortint(0).G & "," & StreamData(LoopC).colortint(0).B
        WriteVar StreamFile, Val(LoopC), "ColorSet2", StreamData(LoopC).colortint(1).R & "," & StreamData(LoopC).colortint(1).G & "," & StreamData(LoopC).colortint(1).B
        WriteVar StreamFile, Val(LoopC), "ColorSet3", StreamData(LoopC).colortint(2).R & "," & StreamData(LoopC).colortint(2).G & "," & StreamData(LoopC).colortint(2).B
        WriteVar StreamFile, Val(LoopC), "ColorSet4", StreamData(LoopC).colortint(3).R & "," & StreamData(LoopC).colortint(3).G & "," & StreamData(LoopC).colortint(3).B
        
    Next LoopC
    
    'Report the results
    If TotalStreams > 1 Then
        MsgBox TotalStreams & " Particulas guardadas en: " & vbCrLf & StreamFile, vbInformation
    Else
        MsgBox TotalStreams & " Particulas guardadas en: " & vbCrLf & StreamFile, vbInformation
    End If
    
    'Set DataChanged variable to false
    DataChanged = False
    CurStreamFile = StreamFile
    
End Sub

Sub CargarParticulasLista()

    On Error GoTo CargarParticulasLista_Err

    Dim LoopC    As Long

    Dim DataTemp As Boolean

    DataTemp = DataChanged
    
    With StreamData(frmParticleEditor.ListParticulas.ListIndex + 1)
    
        'Set the values
        frmParticleEditor.txtPCount.Text = .NumOfParticles
        frmParticleEditor.txtX1.Text = .x1
        frmParticleEditor.txtY1.Text = .y1
        frmParticleEditor.txtX2.Text = .x2
        frmParticleEditor.txtY2.Text = .y2
        frmParticleEditor.txtAngle.Text = .angle
        frmParticleEditor.vecx1.Text = .vecx1
        frmParticleEditor.vecx2.Text = .vecx2
        frmParticleEditor.vecy1.Text = .vecy1
        frmParticleEditor.vecy2.Text = .vecy2
        frmParticleEditor.life1.Text = .life1
        frmParticleEditor.life2.Text = .life2
        frmParticleEditor.fric.Text = .friction
        frmParticleEditor.chkSpin.value = .spin
        frmParticleEditor.spin_speedL.Text = .spin_speedL
        frmParticleEditor.spin_speedH.Text = .spin_speedH
        frmParticleEditor.txtGravStrength.Text = .grav_strength
        frmParticleEditor.txtBounceStrength.Text = .bounce_strength
        frmParticleEditor.chkAlphaBlend.value = .alphaBlend
        frmParticleEditor.chkGravity.value = .gravity
        frmParticleEditor.chkXMove.value = .XMove
        frmParticleEditor.chkYMove.value = .YMove
        frmParticleEditor.move_x1.Text = .move_x1
        frmParticleEditor.move_x2.Text = .move_x2
        frmParticleEditor.move_y1.Text = .move_y1
        frmParticleEditor.move_y2.Text = .move_y2
        
        If .life_counter = -1 Then
            frmParticleEditor.life.Enabled = False
            frmParticleEditor.chkNeverDies.value = vbChecked
        Else
            frmParticleEditor.life.Enabled = True
            frmParticleEditor.life.Text = .life_counter
            frmParticleEditor.chkNeverDies.value = vbUnchecked

        End If
        
        frmParticleEditor.speed.Text = .speed
        
        frmParticleEditor.lstSelGrhs.Clear
        
        For LoopC = 1 To .NumGrhs
            frmParticleEditor.lstSelGrhs.AddItem .grh_list(LoopC)
        Next LoopC
    
    End With
    
    DataChanged = DataTemp
    
    Call Particle_Group_Remove_All
    
    ParticleIndex = frmParticleEditor.ListParticulas.ListIndex + 1
    
    General_Particle_Create ParticleIndex, 50, 50
    
CargarParticulasLista_Err:
    Call LogError(Err.Number, Err.Description, "modParticulas.CargarParticulasLista", Erl)
    Resume Next

End Sub

