Attribute VB_Name = "ClsDX8Engine"
'******************************************************************************************************************************************
'Lorwik> Este modulo suele ser siempre un modulo clase, pero como hay ciertas variables que si o si deben de ir aqui _
 y que otros sistemas que se encuentran en otros modulos necesita consultarlos y dichas variables no se pueden poner como publicas _
 en un modulo de clase, opté por convertir todo el modulo de clase en un modulo normal.
'******************************************************************************************************************************************

Option Explicit

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

Public map_current                As Map

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
Public timerElapsedTime       As Single
Public timerTicksPerFrame     As Double
Public movSpeed                As Single
Public lFrameLimiter          As Long

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

'Public Sub RenderToPicture()
'    Dim Y As Integer
'    Dim X As Integer
'
'    Dim destRect As RECT
'
'    destRect.Bottom = 3200 '100 * Radio
'    destRect.Right = 3200 '100 * Radio
'    destRect.Left = 0
'    destRect.Top = 0
'
'    D3DDevice.BeginScene
'    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
'
'    'Capa 1 y 2
'    For Y = 1 To 100
'        For X = 1 To 100
'            If MapData(X, Y).Graphic(1).grhindex > 0 Then Draw_GrhIndexMiniMap GrhData(MapData(X, Y).Graphic(1).grhindex).Frames(1), (X * Radio) - Radio, (Y * Radio) - Radio, (GrhData(GrhData(MapData(X, Y).Graphic(1).grhindex).Frames(1)).PixelWidth / 32) * Radio, (GrhData(GrhData(MapData(X, Y).Graphic(1).grhindex).Frames(1)).PixelHeight / 32) * Radio, Radio
'            If MapData(X, Y).Graphic(2).grhindex > 0 Then Draw_GrhIndexMiniMap GrhData(MapData(X, Y).Graphic(2).grhindex).Frames(1), (X * Radio) - Radio, (Y * Radio) - Radio, (GrhData(GrhData(MapData(X, Y).Graphic(2).grhindex).Frames(1)).PixelWidth / 32) * Radio, (GrhData(GrhData(MapData(X, Y).Graphic(2).grhindex).Frames(1)).PixelHeight / 32) * Radio, Radio
'        Next X
'    Next Y
'
'    'Capa 3
'    For Y = 1 To 100
'        For X = 1 To 100
'            If MapData(X, Y).Graphic(3).grhindex > 0 Then Draw_GrhIndexMiniMap GrhData(MapData(X, Y).Graphic(3).grhindex).Frames(1), (X * Radio) - (Radio / 2), (Y * Radio) + Radio, (GrhData(GrhData(MapData(X, Y).Graphic(3).grhindex).Frames(1)).PixelWidth / 32) * Radio, (GrhData(GrhData(MapData(X, Y).Graphic(3).grhindex).Frames(1)).PixelHeight / 32) * Radio, Radio, 1
'        Next X
'    Next Y
'
'    'capa 4 (si no quieres techos elimina estos for completos hasta el next)
'    For Y = 1 To 100
'        For X = 1 To 100
'            If MapData(X, Y).Graphic(4).grhindex > 0 Then Draw_GrhIndexMiniMap GrhData(MapData(X, Y).Graphic(4).grhindex).Frames(1), (X * Radio) - (Radio / 2), (Y * Radio) + Radio, (GrhData(GrhData(MapData(X, Y).Graphic(4).grhindex).Frames(1)).PixelWidth / 32) * Radio, (GrhData(GrhData(MapData(X, Y).Graphic(4).grhindex).Frames(1)).PixelHeight / 32) * Radio, Radio, 1
'        Next X
'    Next Y
'
'    D3DDevice.EndScene
'    D3DDevice.Present destRect, ByVal 0, frmRenderer.Picture1.hwnd, ByVal 0
'End Sub

Public Sub Draw_GrhIndexMiniMap(ByVal grh_index As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal PixelWidth As Long, ByVal PixelHeight As Long, ByVal Radio As Integer, Optional ByVal Center As Byte = 0)
If grh_index <= 0 Then Exit Sub
Dim rgb_list(3) As Long

rgb_list(0) = D3DColorXRGB(255, 255, 255)
rgb_list(1) = D3DColorXRGB(255, 255, 255)
rgb_list(2) = D3DColorXRGB(255, 255, 255)
rgb_list(3) = D3DColorXRGB(255, 255, 255)

If Center Then
        If GrhData(grh_index).TileWidth <> 1 Then
            X = X - Int(PixelWidth / 2)
        End If
        If GrhData(grh_index).TileHeight <> 1 Then
            Y = Y - Int(PixelHeight + Radio)
        End If
End If

Device_Box_Textured_Render grh_index, _
X, Y, _
GrhData(grh_index).PixelWidth, GrhData(grh_index).PixelHeight, _
rgb_list, _
GrhData(grh_index).sX, GrhData(grh_index).sY, PixelWidth, PixelHeight

End Sub
Public Sub Draw_GrhIndex(ByVal grh_index As Integer, _
                         ByVal X As Integer, _
                         ByVal Y As Integer, _
                         ByRef Light() As Long)

    '********************************************
    'Dibuja desde un GrhIndex que le indiquemos
    '********************************************
    If grh_index <= 0 Then Exit Sub
    
    Device_Box_Textured_Render grh_index, X, Y, GrhData(grh_index).PixelWidth, GrhData(grh_index).PixelHeight, Light, GrhData(grh_index).sX, GrhData(grh_index).sY

End Sub

Public Sub Draw_Grh(ByRef Grh As Grh, _
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

        Call Device_Box_Textured_Render(CurrentGrhIndex, X, Y, .PixelWidth, .PixelHeight, Light, .sX, .sY, Alpha, angle, Transp)

    
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

    Static RE As RECT

    RE.Left = 0
    RE.Top = 0
    RE.Bottom = frmMain.Renderer.ScaleHeight
    RE.Right = frmMain.Renderer.ScaleWidth

    Engine_ActFPS
    
    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    
    ShowNextFrame
    
    Texto.Engine_Text_Draw 890, 5, "FPS: " & FPS, vbWhite
    Texto.Engine_Text_Draw 5, 5, POSX, vbWhite
    
    D3DDevice.EndScene
    D3DDevice.Present RE, ByVal 0, 0, ByVal 0
    
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

Public Sub MapCapture(ByRef Format As Boolean, ByVal ToWorldMapJPG As Boolean)

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
    Static RE            As RECT
    
    With RE
        .Left = 0
        .Top = 0
        .Bottom = 3200
        .Right = 3200
    End With
    
    'With frmRender.pgbProgress
    '    .Value = 0
    '    .Max = 50000
    'End With
         
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0)
    Call D3DDevice.BeginScene
    
    If ToWorldMapJPG = False Then
    
    'Draw floor layer
    For Y = 1 To 100
        For X = 1 To 100
            
            'Layer 1 **********************************
            If MapData(X, Y).Graphic(1).GrhIndex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(1), (X - 1) * 32 + TilePixelWidth - 35, (Y - 1) * 32 + TilePixelHeight - 35, 0, 1, MapData(X, Y).light_value())

            End If

            '******************************************
            
            'frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1

        Next X
    Next Y
        
    'Draw floor layer 2
    For Y = 1 To 100
        For X = 1 To 100
            
            'Layer 2 **********************************
            If (MapData(X, Y).Graphic(2).GrhIndex <> 0) And VerCapa2 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), (X - 1) * 32 + TilePixelWidth - 35, (Y - 1) * 32 + TilePixelHeight - 35, 1, 1, MapData(X, Y).light_value())

            End If

            '******************************************
            
            'frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1

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
                    Call Draw_Grh(.ObjGrh, PixelOffsetXTemp - 35, PixelOffsetYTemp - 35, 1, 1, MapData(X, Y).light_value())

                End If

                '***********************************************
                
                'Layer 3 *****************************************
                If (.Graphic(3).GrhIndex <> 0) And VerCapa3 Then
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp - 35, PixelOffsetYTemp - 35, 1, 1, MapData(X, Y).light_value())

                End If

                '************************************************
                
                'frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1

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
                    Call Draw_Grh(.Graphic(4), (X - 1) * 32 + TilePixelWidth - 35, (Y - 1) * 32 + TilePixelHeight - 35, 1, 1, MapData(X, Y).light_value())
                        
                End If

                '**********************************
                
                'frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1

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
                    
                    Call Draw_Grh(Grh, PixelOffsetXTemp - 35, PixelOffsetYTemp - 35, 1, 0, MapData(X, Y).light_value())

                End If
                
                'Show blocked tiles
                If (.blocked = 1) And VerBlockeados Then
                    Grh.GrhIndex = 4
                    Call Draw_Grh(Grh, PixelOffsetXTemp - 35, PixelOffsetYTemp - 35, 1, 0, MapData(X, Y).light_value())

                End If

                '******************************************
                
                'frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1

            End With
    
        Next X
    Next Y
    
    ToWorldMap2 = False
    ElseIf ToWorldMapJPG = True Then
    
    For Y = 10 To 91
        For X = 10 To 91
            
            'Layer 1 **********************************
            If MapData(X, Y).Graphic(1).GrhIndex <> 0 Then
                Call Draw_Grh(MapData(X, Y).Graphic(1), (X - 1) * 32 + TilePixelWidth - 35, (Y - 1) * 32 + TilePixelHeight - 35, 0, 1, MapData(X, Y).light_value())

            End If

            '******************************************
            
            'frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1

        Next X
    Next Y
        
    'Draw floor layer 2
    For Y = 10 To 91
        For X = 10 To 91
            
            'Layer 2 **********************************
            If (MapData(X, Y).Graphic(2).GrhIndex <> 0) And VerCapa2 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), (X - 1) * 32 + TilePixelWidth - 35, (Y - 1) * 32 + TilePixelHeight - 35, 1, 1, MapData(X, Y).light_value())

            End If

            '******************************************
            
            'frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1

        Next X
    Next Y
    
    'Draw Transparent Layers
    For Y = 10 To 91
        For X = 10 To 91
                    
            PixelOffsetXTemp = (X - 1) * 32 + TilePixelWidth
            PixelOffsetYTemp = (Y - 1) * 32 + TilePixelHeight
            
            With MapData(X, Y)
                
                'Object Layer **********************************
                If (.ObjGrh.GrhIndex <> 0) And VerObjetos Then
                    Call Draw_Grh(.ObjGrh, PixelOffsetXTemp - 35, PixelOffsetYTemp - 35, 1, 1, MapData(X, Y).light_value())

                End If

                '***********************************************
                
                'Layer 3 *****************************************
                If (.Graphic(3).GrhIndex <> 0) And VerCapa3 Then
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp - 35, PixelOffsetYTemp - 35, 1, 1, MapData(X, Y).light_value())

                End If

                '************************************************
                
                'frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1

            End With

        Next X
    Next Y
        
    'Draw layer 4
    For Y = 10 To 91
        For X = 10 To 91

            With MapData(X, Y)
                
                'Layer 4 **********************************
                If (.Graphic(4).GrhIndex <> 0) And VerCapa4 Then
                        
                    'Draw
                    Call Draw_Grh(.Graphic(4), (X - 1) * 32 + TilePixelWidth - 35, (Y - 1) * 32 + TilePixelHeight - 35, 1, 1, MapData(X, Y).light_value())
                        
                End If

                '**********************************
                
                'frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1

            End With

        Next X
    Next Y
    
    'Draw trans, bloqs, triggers and select tiles
    For Y = 10 To 91
        For X = 10 To 91

            With MapData(X, Y)
                PixelOffsetXTemp = (X - 1) * 32 + TilePixelWidth
                PixelOffsetYTemp = (Y - 1) * 32 + TilePixelHeight
                
                '**********************************
                Grh.FrameCounter = 1
                Grh.Started = 0

                If (.TileExit.Map <> 0) And VerTranslados Then
                    Grh.GrhIndex = 3
                    
                    Call Draw_Grh(Grh, PixelOffsetXTemp - 35, PixelOffsetYTemp - 35, 1, 0, MapData(X, Y).light_value())

                End If
                
                'Show blocked tiles
                If (.blocked = 1) And VerBlockeados Then
                    Grh.GrhIndex = 4
                    Call Draw_Grh(Grh, PixelOffsetXTemp - 35, PixelOffsetYTemp - 35, 1, 0, MapData(X, Y).light_value())

                End If

                '******************************************
                
                'frmRender.pgbProgress.Value = frmRender.pgbProgress.Value + 1

            End With
    
        Next X
    Next Y
    
    ToWorldMap2 = True
    End If 'If que cierra el ToWorldMapJPG
    
    Call D3DDevice.EndScene
    Call D3DDevice.Present(RE, ByVal 0, frmRenderer.Picture1.hwnd, ByVal 0)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''Guardo la imagen''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Call frmRender.Capturar_Imagen(frmRenderer.Picture1, frmRenderer.Picture1)
    
    'Si no existe la carpeta de MiniMapas, la hacemos.
    If Not FileExist(DirMinimapas, vbDirectory) Then
        Call MkDir(DirMinimapas)
    End If
    
    Call SavePicture(frmRenderer.Picture1, DirMinimapas & NumMap_Save & ".bmp")
    'Call SavePicture(frmRenderer.Picture1.Image, App.Path & NumMap_Save & ".bmp")
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
End Sub
