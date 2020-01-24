Attribute VB_Name = "ClsDX8Engine"
'******************************************************************************************************************************************
'Lorwik> Este modulo suele ser siempre un modulo clase, pero como hay ciertas variables que si o si deben de ir aqui _
 y que otros sistemas que se encuentran en otros modulos necesita consultarlos y dichas variables no se pueden poner como publicas _
 en un modulo de clase, opté por convertir todo el modulo de clase en un modulo normal.
'******************************************************************************************************************************************

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

Private Type decoration

    Grh As Grh
    Render_On_Top As Boolean
    subtile_pos As Byte

End Type

Private Type Map_Tile

    Grh(1 To 3) As Grh
    decoration(1 To 5) As decoration
    decoration_count As Byte
    blocked As Boolean
    particle_group_index As Long
    char_index As Long
    light_base_value(0 To 3) As Long
    light_value(0 To 3) As Long
   
    exit_index As Long
    npc_index As Long
    item_index As Long
   
    Trigger As Byte

End Type

Private Type Map

    map_grid() As Map_Tile
    map_x_max As Long
    map_x_min As Long
    map_y_max As Long
    map_y_min As Long
    map_description As String
    'Added by Juan Martín Sotuyo Dodero
    base_light_color As Long

End Type

Dim map_current                As Map

'*********************************************

Private HalfWindowTileWidth    As Integer

Private HalfWindowTileHeight   As Integer

Private TileBufferPixelOffsetX As Integer

Private TileBufferPixelOffsetY As Integer

Private TileBufferSize         As Byte

Public engineBaseSpeed         As Single

Private ScrollPixelsPerFrameX  As Byte

Private ScrollPixelsPerFrameY  As Byte

Public base_light              As Long

Private lFrameTimer            As Long

Public FPS                     As Long

Public FramesPerSecCounter     As Long

Private timerElapsedTime       As Single

Private timerTicksPerFrame     As Double

Public movSpeed                As Single

Private lFrameLimiter          As Long

Private Declare Function QueryPerformanceFrequency _
                Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter _
                Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private Function GetElapsedTime() As Single

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim start_time    As Currency

    Static end_time   As Currency

    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq

    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)

End Function

Public Function Engine_Init() As Boolean
    '*****************************************************
    'Inicia el motor grafico
    '*****************************************************
    On Local Error GoTo ErrorHandler

    Dim aD3dai             As D3DADAPTER_IDENTIFIER8
    Dim DispMode           As D3DDISPLAYMODE
    Dim DispModeBK         As D3DDISPLAYMODE
    Dim D3DWindow          As D3DPRESENT_PARAMETERS
    Dim ColorKeyVal        As Long
    Dim EleccionProcessing As Long
    
    Set SurfaceDB = New clsTexManager
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    Call D3D.GetAdapterIdentifier(D3DADAPTER_DEFAULT, 0, aD3dai)
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = 3200
        .BackBufferHeight = 3200
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.Renderer.hwnd

    End With
    
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.Renderer.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
                                                            
    D3DDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    With D3DDevice
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
        .SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_TFACTOR

    End With
                                                            
    '************************************************************************************************************************************
    
    HalfWindowTileHeight = (frmMain.Renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.Renderer.ScaleWidth / 32) \ 2
    TileBufferSize = 11
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    TilePixelWidth = 32 'Tamaño de tile
    TilePixelHeight = 32 'Tamaño de tile
    engineBaseSpeed = 0.019  'Velocidad a la que va a correr el engine (modifica la velocidad de caminata)
    
    '***********************************
    'Tamaño del mapa
    '***********************************
    MinXBorder = XMinMapSize + (ClienteWidth \ 2)
    MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
    MinYBorder = YMinMapSize + (ClienteHeight \ 2)
    MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)
    '***********************************
    
    ScrollPixelsPerFrameX = 8
    ScrollPixelsPerFrameY = 8
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    UserPos.X = 50
    UserPos.Y = 50
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)
    
    movSpeed = 1
    
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function

    End If
    
    If Err Then
        MsgBox "No se puede iniciar DirectD3D. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function

    End If
    
    If D3DDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectDevice. Por favor asegurese de tener la ultima version correctamente instalada."
        Exit Function

    End If
    
    'Display form handle, View window offset from 0,0 of display form, Tile Size, Display size in tiles, Screen buffer
    
    'Cargamos el indice de graficos.
    If FileExist(InitPath & "Graficos.ini", vbNormal) Then
        Call LoadGrhIni
    Else
        Call LoadGrhData
    End If
    
    'Cargamos Minimap.dat
    Call LoadMiniMap
    
    'Cargamos listado de particulas.
    Call CargarParticulas
    
    DoEvents
    
    frmCargando.P1.Visible = True
    frmCargando.L(0).Visible = True
    
    frmCargando.X.Caption = "Cargando Cuerpos..."
    Call modIndices.CargarIndicesDeCuerpos
    DoEvents
    frmCargando.P2.Visible = True
    frmCargando.L(1).Visible = True
    
    frmCargando.X.Caption = "Cargando Cabezas..."
    modIndices.CargarIndicesDeCabezas
    DoEvents
    frmCargando.P3.Visible = True
    frmCargando.L(2).Visible = True
    
    frmCargando.X.Caption = "Cargando NPC's..."
    modIndices.CargarIndicesNPC
    DoEvents
    frmCargando.P4.Visible = True
    frmCargando.L(3).Visible = True
    
    frmCargando.X.Caption = "Cargando Objetos..."
    modIndices.CargarIndicesOBJ
    DoEvents
    frmCargando.P5.Visible = True
    frmCargando.L(4).Visible = True
    
    frmCargando.X.Caption = "Cargando Triggers..."
    modIndices.CargarIndicesTriggers
    DoEvents
    frmCargando.P6.Visible = True
    frmCargando.L(5).Visible = True
    DoEvents
    Texto.Engine_Init_FontSettings
    Texto.Engine_Init_FontTextures

    Engine_Init = True
    
    Exit Function

ErrorHandler:
    Debug.Print "Error Number Returned: " & Err.Number
    MsgBox "Error in engine initialization: " & Err.Number & ": " & Err.Description & " Dispositivo: " & Trim$(StrConv(aD3dai.Description, vbUnicode)), vbCritical, "Direct3D Initialization"
    Engine_Init = False

End Function

Public Sub Start()

    Dim i As Byte

    base_light = ARGB(255, 255, 255, 255)

    For i = 0 To 3
        LightIluminado(i) = RGB(255, 255, 255)
    Next i

    DoEvents
    
    Do While prgRun

        'Call FlushBuffer
        If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then
            CheckKeys
            Render
            Engine_ActFPS
            
        Else
            Sleep 10&

        End If

        DoEvents
    Loop
    
    Engine_Deinit
    End

End Sub

Public Sub Engine_Deinit()
    '************************************
    'Termina con el motor grafico
    '************************************
    Erase MapData
    Erase CharList
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set dX = Nothing
    End

End Sub

Private Function CreateTLVertex(X As Single, _
                                Y As Single, _
                                Z As Single, _
                                rhw As Single, _
                                color As Long, _
                                Specular As Long, _
                                tu As Single, _
                                tv As Single) As TLVERTEX
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    CreateTLVertex.X = X
    CreateTLVertex.Y = Y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv

End Function

Public Sub Engine_ActFPS()

    If GetTickCount - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount

    End If

End Sub

Function InMapBounds(X As Integer, Y As Integer) As Boolean
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************

    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        InMapBounds = False
        Exit Function

    End If

    InMapBounds = True

End Function

Public Sub Draw_GrhIndex(ByVal grh_index As Integer, _
                         ByVal X As Integer, _
                         ByVal Y As Integer, _
                         ByRef Light() As Long)

    '********************************************
    'Dibuja desde un GrhIndex que le indiquemos
    '********************************************
    If grh_index <= 0 Then Exit Sub
    
    Device_Box_Textured_Render grh_index, X, Y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, Light, GrhData(grh_index).sX, GrhData(grh_index).sY

End Sub

Private Sub Draw_Grh(ByRef Grh As Grh, _
                     ByVal X As Integer, _
                     ByVal Y As Integer, _
                     ByVal Center As Byte, _
                     ByVal Animate As Byte, _
                     ByRef Light() As Long, _
                     Optional ByVal Alpha As Boolean, _
                     Optional ByVal map_x As Integer = 1, _
                     Optional ByVal map_y As Integer = 1, _
                     Optional ByVal angle As Single, _
                     Optional Transp As Byte = 255)

    '**********************************************************
    'Dibuja desde un Grh, pero primero hay que iniciar el Grh
    '**********************************************************
    On Error Resume Next

    Dim CurrentGrhIndex As Long
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
    If Grh.GrhIndex > UBound(GrhData) Or GrhData(Grh.GrhIndex).NumFrames = 0 And GrhData(Grh.GrhIndex).FileNum = 0 Then
        'Call InitGrh(Grh, 20299)
        Call AddtoRichTextBox(frmMain.StatTxt, "Error en Grh. Posicion: X:" & map_x & " Y:" & map_y, 255, 0, 0)
    End If
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * (movSpeed / 1.5)

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> -1 Then
                    
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If

                End If

            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
    
        'Center Grh over X,Y pos
        If Center Then
        
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * (32 \ 2)) + 32 \ 2
            End If
    
            If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * 32) + 32
            End If
    
        End If

        Call Device_Box_Textured_Render(CurrentGrhIndex, X, Y, .pixelWidth, .pixelHeight, Light, .sX, .sY, Alpha, angle, Transp)

    
    End With
    
End Sub

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, _
                                ByRef dest As RECT, _
                                ByRef src As RECT, _
                                ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, _
                                Optional ByRef Textures_Height As Long, _
                                Optional ByVal angle As Single)

    '**************************************************************
    'Crea el plano donde se desarrolla el juego y todo se dibuja.
    'Si jugamos con la configuracion de este sub, podremos provocar
    'ciertos efecto. (ejem: montañas, reflejos y sombras)
    '
    ' * v1      * v3
    ' |\        |
    ' |  \      |
    ' |    \    |
    ' |      \  |
    ' |        \|
    ' * v0      * v2
    '**************************************************************
    Dim x_center    As Single

    Dim y_center    As Single

    Dim radius      As Single

    Dim x_Cor       As Single

    Dim y_Cor       As Single

    Dim left_point  As Single

    Dim right_point As Single

    Dim temp        As Single
   
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.Bottom - dest.Top) / 2
       
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.Bottom - y_center) ^ 2)
       
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = 3.1459 - right_point

    End If
   
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius

    End If
   
    '0 - Bottom left vertex
    If Textures_Width Or Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width + 0.001, (src.Bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)

    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius

    End If
   
    '1 - Top left vertex
    If Textures_Width Or Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width + 0.001, src.Top / Textures_Height + 0.001)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)

    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius

    End If
   
    '2 - Bottom right vertex
    If Textures_Width Or Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)

    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius

    End If
   
    '3 - Top right vertex
    If Textures_Width Or Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height + 0.001)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)

    End If
 
End Sub

Private Function Geometry_Create_TLVertex(ByVal X As Single, _
                                          ByVal Y As Single, _
                                          ByVal Z As Single, _
                                          ByVal rhw As Single, _
                                          ByVal color As Long, _
                                          ByVal Specular As Long, _
                                          tu As Single, _
                                          ByVal tv As Single) As TLVERTEX
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv

End Function

Public Sub Device_Box_Textured_Render(ByVal GrhIndex As Long, _
                                      ByVal dest_x As Integer, _
                                      ByVal dest_y As Integer, _
                                      ByVal src_width As Integer, _
                                      ByVal src_height As Integer, _
                                      ByRef rgb_list() As Long, _
                                      ByVal src_x As Integer, _
                                      ByVal src_y As Integer, _
                                      Optional ByVal alpha_blend As Boolean, _
                                      Optional ByVal angle As Single, _
                                      Optional alphabyte As Byte = 255)

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 2/12/2004
    'Just copies the Textures
    '**************************************************************
    Static src_rect            As RECT

    Static dest_rect           As RECT

    Static temp_verts(3)       As TLVERTEX

    Static d3dTextures         As D3D8Textures

    Static light_value(0 To 3) As Long
    
    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    'Lorwik> Esto de las luces hay que mirarlo, asi no me convence.
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    If (light_value(0) = 0) Then light_value(0) = base_light
    If (light_value(1) = 0) Then light_value(1) = base_light
    If (light_value(2) = 0) Then light_value(2) = base_light
    If (light_value(3) = 0) Then light_value(3) = base_light
        
    'Set up the source rectangle
    With src_rect
        .Bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y

    End With
                
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y

    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
    
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
        'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    End If
    
    D3DDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(alphabyte, 0, 0, 0)
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    End If

End Sub

Public Sub Render()

    Static re As RECT

    re.Left = 0
    re.Top = 0
    re.Bottom = frmMain.Renderer.ScaleHeight
    re.Right = frmMain.Renderer.ScaleWidth

    Engine_ActFPS
    
    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    
    ShowNextFrame
    
    Texto.Engine_Text_Draw 890, 5, "FPS: " & FPS, vbWhite
    Texto.Engine_Text_Draw 5, 5, POSX, vbWhite
    
    D3DDevice.EndScene
    D3DDevice.Present re, ByVal 0, 0, ByVal 0
    
    lFrameLimiter = GetTickCount
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

End Sub

Sub ShowNextFrame()

    Static OffsetCounterX As Single

    Static OffsetCounterY As Single

    '****** Move screen Left and Right if needed ******
    If AddtoUserPos.X <> 0 Then
        OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame

        If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.X) Then
            OffsetCounterX = 0
            AddtoUserPos.X = 0

        End If

    End If
            
    '****** Move screen Up and Down if needed ******
    If AddtoUserPos.Y <> 0 Then
        OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame

        If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.Y) Then
            OffsetCounterY = 0
            AddtoUserPos.Y = 0

        End If

    End If
        
    Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)

End Sub

Sub RenderScreen(ByVal tilex As Integer, _
                 ByVal tiley As Integer, _
                 ByVal PixelOffsetX As Byte, _
                 ByVal PixelOffsetY As Byte)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    '**************************************************************
   
    Dim Y                As Integer     'Keeps track of where on map we are

    Dim X                As Integer     'Keeps track of where on map we are

    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen

    Dim MinY             As Integer  'Start Y pos on current map

    Dim MaxY             As Integer  'End Y pos on current map

    Dim MinX             As Integer  'Start X pos on current map

    Dim MaxX             As Integer  'End X pos on current map

    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen

    Dim minXOffset       As Integer

    Dim minYOffset       As Integer

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs

    Dim CurrentGrhIndex  As Integer

    Dim offx             As Integer

    Dim offy             As Integer

    Dim Grh              As Grh
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    MinY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    MinX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
        
    'Make sure mins and maxs are allways in map bounds
    If MinY < XMinMapSize Then
        minYOffset = YMinMapSize - MinY
        MinY = YMinMapSize

    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If MinX < XMinMapSize Then
        minXOffset = XMinMapSize - MinX
        MinX = XMinMapSize

    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    GenerarVista
    
    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX

            'Layer 1 **********************************
            If InMapBounds(X, Y) Then
                If MapData(X, Y).Graphic(1).GrhIndex And VerCapa1 Then
                    Call Draw_Grh(MapData(X, Y).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, MapData(X, Y).light_value, X, Y)

                End If
                
                If MapData(X, Y).Graphic(2).GrhIndex <> 0 And VerCapa2 Then
                    Call Draw_Grh(MapData(X, Y).Graphic(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, 1, MapData(X, Y).light_value, , X, Y)

                End If

            End If

            '******************************************
            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y

    ScreenY = minYOffset - TileBufferSize

    For Y = MinY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For X = MinX To MaxX

            If InMapBounds(X, Y) Then
                PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
                PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY

                With MapData(X, Y)
                    '******************************************
    
                    'Object Layer **********************************
                    If .ObjGrh.GrhIndex <> 0 And VerObjetos Then
                        Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value, , X, Y)

                    End If
                    
                    'Char layer ************************************
                    If .CharIndex <> 0 And VerNpcs Then
                        Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, X, Y)

                    End If

                    '*************************************************
       
                    'Layer 3 *****************************************
                    If .Graphic(3).GrhIndex <> 0 And VerCapa3 Then
                        Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value, , X, Y)

                    End If

                    '************************************************
    
                End With

            End If

            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y

    ScreenY = minYOffset - 5
    
    'Particulas ************************************************
    ScreenY = minYOffset - TileBufferSize

    For Y = MinY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For X = MinX To MaxX

            'Particulas**************************************
            If MapData(X, Y).particle_group_index Then Particle_Group_Render MapData(X, Y).particle_group_index, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY
            '************************************************
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y

    '***********************************************************

    'Draw blocked tiles and grid
    ScreenY = minYOffset - TileBufferSize

    For Y = MinY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For X = MinX To MaxX
    
            'Layer 4 **********************************
            If MapData(X, Y).Graphic(4).GrhIndex And VerCapa4 And Not bTecho Then
                Call Draw_Grh(MapData(X, Y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, MapData(X, Y).light_value, , X, Y)

            End If

            '**********************************
            If MapData(X, Y).TileExit.Map <> 0 And VerTranslados Then
                Grh.GrhIndex = 3
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call Draw_Grh(Grh, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, LightIluminado(), True, X, Y)

            End If
        
            'Show blocked tiles
            If VerBlockeados And MapData(X, Y).blocked = 1 Then
                Grh.GrhIndex = 4
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call Draw_Grh(Grh, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, LightIluminado(), , X, Y)

            End If
            
            If VerGrilla Then
                Call Draw_Grh(Grh, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, LightIluminado(), , X, Y)

            End If
            
            If VerTriggers Then '4978
                If MapData(X, Y).Trigger > 0 Then Texto.Engine_Text_Draw ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, MapData(X, Y).Trigger, vbWhite, , , 3

            End If
            
            If Seleccionando Then
                If X >= SeleccionIX And Y >= SeleccionIY Then
                    If X <= SeleccionFX And Y <= SeleccionFY Then
                        Grh.GrhIndex = 2
                        Grh.FrameCounter = 1
                        Grh.Started = 0
                        Call Draw_Grh(Grh, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, MapData(X, Y).light_value, , X, Y)

                    End If

                End If

            End If
        
            ScreenX = ScreenX + 1
        Next X

        ScreenY = ScreenY + 1
    Next Y
 
End Sub

Private Sub CharRender(ByVal CharIndex As Long, _
                       ByVal PixelOffsetX As Integer, _
                       ByVal PixelOffsetY As Integer, _
                       ByVal X As Byte, _
                       ByVal Y As Byte)

    '*******************************************************
    'Esto forma parte del RenderScreen.
    'Dibuja todo aquello que tenga cuerpo (por asi decirlo)
    'Bichos y PJ
    '*******************************************************
    Dim moved As Boolean

    Dim Pos   As Integer

    Dim line  As String
    
    With CharList(CharIndex)

        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffset.X = .MoveOffset.X + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffset.X >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffset.X <= 0) Then
                    .MoveOffset.X = 0
                    .scrollDirectionX = 0

                End If

            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffset.Y = .MoveOffset.Y + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffset.Y >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffset.Y <= 0) Then
                    .MoveOffset.Y = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If
        
        'If done moving stop animation
        If Not moved Then
            If .Heading = 0 Then Exit Sub
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1

            .Moving = False

        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffset.X
        PixelOffsetY = PixelOffsetY + .MoveOffset.Y
           
        movSpeed = 1.3

        'Dibujamos el cuerpo
        If .Body.Walk(.Heading).GrhIndex Then Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, LightIluminado(), , X, Y)

        'Dibujamos la Cabeza
        If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, LightIluminado(), , X, Y)

    End With

End Sub

Public Sub MapCapture(ByRef Format As Boolean)

    '*************************************************
    'Author: Torres Patricio(Pato)
    'Last modified:12/03/11
    '*************************************************
    Dim D3DWindow        As D3DPRESENT_PARAMETERS

    Dim Y                As Long     'Keeps track of where on map we are

    Dim X                As Long     'Keeps track of where on map we are

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs
          
    Dim Grh              As Grh      'Temp Grh for show tile and blocked

    Static re            As RECT

    re.Left = 0
    re.Top = 0
    re.Bottom = 3200
    re.Right = 3200

    frmRender.pgbProgress.value = 0
    frmRender.pgbProgress.max = 50000
            
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
    D3DDevice.BeginScene
    'Draw floor layer

    For Y = 1 To 100
        For X = 1 To 100
            
            'Layer 1 **********************************

            If MapData(X, Y).Graphic(1).GrhIndex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(1), (X - 1) * 32 + TilePixelWidth, (Y - 1) * 32 + TilePixelHeight, 0, 1, MapData(X, Y).light_value())

            End If

            '******************************************
            
            frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1

        Next X
    Next Y
        
    'Draw floor layer 2

    For Y = 1 To 100
        For X = 1 To 100
            
            'Layer 2 **********************************

            If (MapData(X, Y).Graphic(2).GrhIndex <> 0) And VerCapa2 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), (X - 1) * 32 + TilePixelWidth, (Y - 1) * 32 + TilePixelHeight, 1, 1, MapData(X, Y).light_value())

            End If

            '******************************************
            
            frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1

        Next X
    Next Y
    
    'Draw Transparent Layers

    For Y = 1 To 100
        For X = 1 To 100
                    
            PixelOffsetXTemp = (X - 1) * 32 + TilePixelWidth
            PixelOffsetYTemp = (Y - 1) * 32 + TilePixelHeight
            
            With MapData(X, Y)
                'Object Layer **********************************

                If (.ObjGrh.GrhIndex <> 0) And VerObjetos Then
                    Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value())

                End If

                '***********************************************
                
                'Layer 3 *****************************************

                If (.Graphic(3).GrhIndex <> 0) And VerCapa3 Then
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value())

                End If

                '************************************************
                
                frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1

            End With

        Next X
    Next Y
        
    'Draw layer 4

    For Y = 1 To 100
        For X = 1 To 100

            With MapData(X, Y)
                'Layer 4 **********************************

                If (.Graphic(4).GrhIndex <> 0) And VerCapa4 Then
                    'Draw
                    Call Draw_Grh(.Graphic(4), (X - 1) * 32 + TilePixelWidth, (Y - 1) * 32 + TilePixelHeight, 1, 1, MapData(X, Y).light_value())

                End If

                '**********************************
                
                frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1

            End With

        Next X
    Next Y
    
    'Draw trans, bloqs, triggers and select tiles

    For Y = 1 To 100
        For X = 1 To 100

            With MapData(X, Y)
                PixelOffsetXTemp = (X - 1) * 32 + TilePixelWidth
                PixelOffsetYTemp = (Y - 1) * 32 + TilePixelHeight
                
                '**********************************
                Grh.FrameCounter = 1
                Grh.Started = 0

                If (.TileExit.Map <> 0) And VerTranslados Then
                    Grh.GrhIndex = 3
                    
                    Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 0, MapData(X, Y).light_value())

                End If
                
                'Show blocked tiles

                If (.blocked = 1) And VerBlockeados Then
                    Grh.GrhIndex = 4
                    
                    Call Draw_Grh(Grh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 0, MapData(X, Y).light_value())

                End If

                '******************************************
                
                frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1

            End With

        Next X
    Next Y
          
    D3DDevice.EndScene
    D3DDevice.Present re, ByVal 0, frmRender.picMap.hwnd, ByVal 0
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''Guardo la imagen'''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call frmRender.Capturar_Imagen(frmRender.picMap, frmRender.picMap)
    SavePicture frmRender.picMap, App.Path & "\Renderizados\" & NumMap_Save & ".bmp"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
End Sub

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
    'Last Modify Date: 1/04/2003
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
    'Modified by Juan Martín Sotuyo Dodero
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
    'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
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
    'Modified by: Juan Martín Sotuyo Dodero
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
    'Modified by: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/15/2003
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
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''/[PARTICULAS]''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
