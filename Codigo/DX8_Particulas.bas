Attribute VB_Name = "DX8_Particulas"
Option Explicit


'*********************************
'Particulas
'*********************************
Private base_tile_size As Integer

Private Type Particle

    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    X1 As Long
    X2 As Long
    Y1 As Long
    Y2 As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    Radio As Integer
    rgb_list(0 To 3) As Long

End Type

Private Type Stream

    name As String
    NumOfParticles As Long
    NumGrhs As Long
    ID As Long
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
   
    Speed As Single
    life_counter As Long

End Type

Private Type particle_group

    Active As Boolean
    ID As Long
    map_x As Long
    map_y As Long
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    live As Long
    liv1 As Integer
    liveend As Long
    
    X1 As Long
    X2 As Long
    Y1 As Long
    Y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    rgb_list(0 To 3) As Long
    
    Speed As Single
    life_counter As Long
    
    Radio As Integer

End Type

'Particle system
 
'Dim StreamData() As particle_group
Dim TotalStreams             As Long

Dim particle_group_count     As Long

Dim particle_group_last      As Long

Public particle_group_list() As particle_group

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''[PARTICULAS]''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Particle_Group_Create(ByVal map_x As Integer, _
                                      ByVal map_y As Integer, _
                                      ByRef grh_index_list() As Long, _
                                      ByRef rgb_list() As Long, _
                                      Optional ByVal particle_count As Long = 20, _
                                      Optional ByVal stream_type As Long = 1, _
                                      Optional ByVal alpha_blend As Boolean, _
                                      Optional ByVal alive_counter As Long = -1, _
                                      Optional ByVal frame_speed As Single = 0.5, _
                                      Optional ByVal ID As Long, _
                                      Optional ByVal X1 As Integer, _
                                      Optional ByVal Y1 As Integer, _
                                      Optional ByVal angle As Integer, _
                                      Optional ByVal vecx1 As Integer, _
                                      Optional ByVal vecx2 As Integer, _
                                      Optional ByVal vecy1 As Integer, _
                                      Optional ByVal vecy2 As Integer, _
                                      Optional ByVal life1 As Integer, _
                                      Optional ByVal life2 As Integer, _
                                      Optional ByVal fric As Integer, _
                                      Optional ByVal spin_speedL As Single, _
                                      Optional ByVal gravity As Boolean, _
                                      Optional grav_strength As Long, _
                                      Optional bounce_strength As Long, _
                                      Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional ByVal Radio As Integer) As Long
                                        
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 12/15/2002
    'Returns the particle_group_index if successful, else 0
    '**************************************************************
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, Radio
        Else
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, Radio

        End If

    End If

End Function

Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True

    End If

End Function
 
Public Function Particle_Group_Remove_All() As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Ultima modificacion 08/05/2020 por ReyarB
    '
    '*****************************************************************
    Dim index As Long
    
    For index = 1 To particle_group_last

        'Make sure it's a legal index
        If Particle_Group_Check(index) Then
            Particle_Group_Destroy index

        End If

    Next index
    
    Particle_Group_Remove_All = True

End Function
 
Public Function Particle_Group_Find(ByVal ID As Long) As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    'Find the index related to the handle
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    loopc = 1

    Do Until particle_group_list(loopc).ID = ID

        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function

        End If

        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
    Exit Function
ErrorHandler:
    Particle_Group_Find = 0

End Function
 
Public Function Particle_Get_Type(ByVal particle_group_index As Long) As Byte

    On Error GoTo ErrorHandler:

    Particle_Get_Type = particle_group_list(particle_group_index).stream_type
    Exit Function
ErrorHandler:
    Particle_Get_Type = 0

End Function

Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    On Error Resume Next

    Dim temp As particle_group

    Dim i    As Integer
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    
    particle_group_list(particle_group_index) = temp
    
    'Update array size
    If particle_group_index = particle_group_last Then

        Do Until particle_group_list(particle_group_last).Active
            particle_group_last = particle_group_last - 1

            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub

            End If

        Loop
        Debug.Print particle_group_last & "," & UBound(particle_group_list)
        ReDim Preserve particle_group_list(1 To particle_group_last) As particle_group

    End If

    particle_group_count = particle_group_count - 1

End Sub
 
Private Sub Particle_Group_Make(ByVal particle_group_index As Long, _
                                ByVal map_x As Integer, _
                                ByVal map_y As Integer, _
                                ByVal particle_count As Long, _
                                ByVal stream_type As Long, _
                                ByRef grh_index_list() As Long, _
                                ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, _
                                Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, _
                                Optional ByVal ID As Long, _
                                Optional ByVal X1 As Integer, _
                                Optional ByVal Y1 As Integer, _
                                Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, _
                                Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, _
                                Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, _
                                Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, _
                                Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, _
                                Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional Radio As Integer)
                               
    '*****************************************************************
    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Last Modify Date: 5/15/2003
    'Makes a new particle effect
    'Modified by Juan Mart�n Sotuyo Dodero
    '*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)

    End If

    particle_group_count = particle_group_count + 1
   
    'Make active
    particle_group_list(particle_group_index).Active = True
   
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y

    End If
   
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    particle_group_list(particle_group_index).Radio = Radio
   
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False

    End If
   
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
   
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).X1 = X1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).X2 = X2
    particle_group_list(particle_group_index).Y2 = Y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    'Color > el R y el B esta intercambiados.
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(3)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(1)
   
    'handle
    particle_group_list(particle_group_index).ID = ID
   
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    'plot particle group on map
    If (map_x <> -1) And (map_y <> -1) Then
        MapData(map_x, map_y).particle_group_index = particle_group_index

    End If
   
End Sub

Public Function Particle_Type_Get(ByVal particle_Index As Long) As Long

    '*****************************************************************
    'Author: Juan Mart�n Sotuyo Dodero (juansotuyo@hotmail.com)
    'Last Modify Date: 8/27/2003
    'Returns the stream type of a particle stream
    '*****************************************************************
    If Particle_Group_Check(particle_Index) Then
        Particle_Type_Get = particle_group_list(particle_Index).stream_type
    Else
        Particle_Type_Get = 0

    End If

End Function

Public Sub Particle_Group_Render(ByVal particle_group_index As Long, _
                                 ByVal screen_x As Long, _
                                 ByVal screen_y As Long)

    '*****************************************************************
    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Modified by: Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 5/15/2003
    'Renders a particle stream at a paticular screen point
    '*****************************************************************
    If particle_group_index = 0 Then Exit Sub
    
    Dim loopc            As Long

    Dim temp_rgb(0 To 3) As Long

    Dim no_move          As Boolean
    
    'Set colors
    temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
    temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
    temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
    temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    
    If particle_group_list(particle_group_index).alive_counter Then
    
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame

        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True

        End If
    
        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
                
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
               screen_x, screen_y, _
               particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
               temp_rgb(), _
               particle_group_list(particle_group_index).alpha_blend, no_move, _
               particle_group_list(particle_group_index).X1, particle_group_list(particle_group_index).Y1, particle_group_list(particle_group_index).angle, _
               particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
               particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
               particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
               particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
               particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
               particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).X2, _
               particle_group_list(particle_group_index).Y2, particle_group_list(particle_group_index).XMove, _
               particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
               particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
               particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
               particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).Radio, _
               particle_group_list(particle_group_index).particle_count, loopc
        Next loopc
        
        If no_move = False Then

            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1

            End If

        End If
    
    Else
        'If it's dead destroy it
        Particle_Group_Destroy particle_group_index

    End If

End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, _
                            ByVal screen_x As Long, _
                            ByVal screen_y As Long, _
                            ByVal grh_index As Long, _
                            ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, _
                            Optional ByVal no_move As Boolean, _
                            Optional ByVal X1 As Integer, _
                            Optional ByVal Y1 As Integer, _
                            Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, _
                            Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, _
                            Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, _
                            Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, _
                            Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, _
                            Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, _
                            Optional ByVal X2 As Integer, _
                            Optional ByVal Y2 As Integer, _
                            Optional ByVal XMove As Boolean, _
                            Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional ByVal Radio As Integer, Optional ByVal count As Integer, Optional ByVal index As Integer)

    '**************************************************************
    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Modified by: Juan Mart�n Sotuyo Dodero
    'Ultima modificacion 08/05/2020 por ReyarB
    '**************************************************************
    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            'Start new particle
            InitGrh temp_particle.Grh, grh_index, alpha_blend

            If Radio = 0 Then
                temp_particle.X = RandomNumber(X1, X2)
                temp_particle.Y = RandomNumber(Y1, Y2)
            Else
                temp_particle.X = (RandomNumber(X1, X2) + Radio) + Radio * Cos(PI * 2 * index / count)
                temp_particle.Y = (RandomNumber(Y1, Y2) + Radio) + Radio * Sin(PI * 2 * index / count)

            End If

            temp_particle.X = RandomNumber(X1, X2) - (base_tile_size \ 2)
            temp_particle.Y = RandomNumber(Y1, Y2) - (base_tile_size \ 2)
            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.friction = fric
        Else

            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength

                If temp_particle.Y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength

                End If

            End If

            'Do rotation
            If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0

            End If
            
            If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)

        End If
        
        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
        temp_particle.alive_counter = temp_particle.alive_counter - 1

    End If
    
    'Draw it
    If temp_particle.Grh.GrhIndex Then
        Draw_Grh temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, 1, 1, rgb_list(), alpha_blend, , , temp_particle.Grh.angle

    End If
    
    If temp_particle.Grh.GrhIndex Then
        Draw_Grh temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, 1, 1, rgb_list(), alpha_blend, , , temp_particle.Grh.angle

    End If

End Sub

Private Function Particle_Group_Next_Open() As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    If particle_group_last = 0 Then
        Particle_Group_Next_Open = 1
        Exit Function

    End If
    
    loopc = 1

    Do Until particle_group_list(loopc).Active = False

        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function

        End If

        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc

    Exit Function

ErrorHandler:

End Function
 
Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).Active Then
            Particle_Group_Check = True

        End If

    End If

End Function

Public Function Map_Particle_Group_Get(ByVal map_x As Long, ByVal map_y As Long) As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 2/20/2003
    'Checks to see if a tile position has a particle_group_index and return it
    '*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Particle_Group_Get = map_current.map_grid(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0

    End If

End Function

Private Function Char_Check(ByVal char_index As Integer) As Boolean

    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Mart�n Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (CharList(char_index).Heading > 0)

    End If
    
End Function

Public Function Map_In_Bounds(ByVal map_x As Long, ByVal map_y As Long) As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Checks to see if a tile position is in the maps bounds
    '*****************************************************************
    If map_x < map_current.map_x_min Or map_x > map_current.map_x_max Or map_y < map_current.map_y_min Or map_y > map_current.map_y_max Then
        Map_In_Bounds = False
        Exit Function

    End If
   
    Map_In_Bounds = True

End Function


