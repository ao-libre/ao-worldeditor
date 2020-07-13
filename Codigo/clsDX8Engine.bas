Attribute VB_Name = "DX8_Engine"
'******************************************************************************************************************************************
'Lorwik> Este modulo suele ser siempre un modulo clase, pero como hay ciertas variables que si o si deben de ir aqui _
 y que otros sistemas que se encuentran en otros modulos necesita consultarlos y dichas variables no se pueden poner como publicas _
 en un modulo de clase, opt� por convertir todo el modulo de clase en un modulo normal.
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
Private timerElapsedTime       As Single
Public timerTicksPerFrame      As Double
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
        Call QueryPerformanceFrequency(timer_freq)
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)

End Function

Public Function Engine_Init_D3DDevice(RENDER_MODE As CONST_D3DCREATEFLAGS) As Boolean
    
    On Error GoTo DeviceError
    
    Dim DispMode           As D3DDISPLAYMODE
    Dim D3DWindow          As D3DPRESENT_PARAMETERS

    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8

    Call D3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = 3200 * 3
        .BackBufferHeight = 3200 * 3
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.Renderer.hWnd

    End With
    
    If Not D3DDevice Is Nothing Then
        Set D3DDevice = Nothing
    End If
    
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DWindow.hDeviceWindow, RENDER_MODE, D3DWindow)
    
    Engine_Init_D3DDevice = True
    
    Exit Function
    
DeviceError:
    
    Set D3DDevice = Nothing
    
    Engine_Init_D3DDevice = False

End Function

Public Function Engine_Init() As Boolean
    '*****************************************************
    'Inicia el motor grafico
    '*****************************************************
    On Local Error GoTo ErrorHandler

    ' Tratamos de inicializar el DirectX Device con la mejor configuracion posible.
    If Not Engine_Init_D3DDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
        If Not Engine_Init_D3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not Engine_Init_D3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                
                Call MsgBox("No se pudo inicializar el motor grafico." & vbNewLine & "Compruebe que las librerias esten registradas correctamente.")
                
                End
                
            End If
        End If
    End If
    
    ' Con esto obtenemos el nombre de la placa de video detectada.
    Dim aD3dai As D3DADAPTER_IDENTIFIER8
    Call D3D.GetAdapterIdentifier(D3DADAPTER_DEFAULT, 0, aD3dai)
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    With D3DDevice
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
        .SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_TFACTOR

    End With
                                             
    Set SurfaceDB = New clsTexManager
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)
                                             
    '************************************************************************************************************************************
    
    HalfWindowTileHeight = (frmMain.Renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.Renderer.ScaleWidth / 32) \ 2
    TileBufferSize = 11
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    TilePixelWidth = 32 'Tama�o de tile
    TilePixelHeight = 32 'Tama�o de tile
    engineBaseSpeed = 0.019  'Velocidad a la que va a correr el engine (modifica la velocidad de caminata)
    
    '***********************************
    'Tama�o del mapa
    'Ultima modificacion 08/05/2020 por ReyarB
    '***********************************
    MinXBorder = XMinMapSize + (ClienteWidth \ 2)
    MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
    MinYBorder = YMinMapSize + (ClienteHeight \ 2)
    MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)
    '***********************************
    
    ScrollPixelsPerFrameX = 8
    ScrollPixelsPerFrameY = 8
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapData2(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    UserPos.X = 50
    UserPos.Y = 50

    movSpeed = 1

    
    'Display form handle, View window offset from 0,0 of display form, Tile Size, Display size in tiles, Screen buffer
    Call LoadGrhData
    Call LoadMiniMap
    Call CargarParticulas
    DoEvents
    frmCargando.P1.Visible = True
    frmCargando.L(0).Visible = True
    frmCargando.X.Caption = "Cargando Cuerpos..."
    modIndices.CargarIndicesDeCuerpos
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

    MsgBox "Error: " & err.Number & _
           "Descripcion: " & err.Description & vbNewLine & _
           "Dispositivo: " & Trim$(StrConv(aD3dai.Description, vbUnicode)), vbCritical, "Direct3D Initialization"
    
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
                                z As Single, _
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
    CreateTLVertex.z = z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv

End Function

'*************************************************
'Ultima modificacion 08/05/2020 por ReyarB
'*************************************************
Public Sub Engine_ActFPS()

    If GetTickCount - lFrameTimer > 300 Then ' antes 1000
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
    
    If Grh.GrhIndex > GrhCount Or GrhData(Grh.GrhIndex).NumFrames = 0 And GrhData(Grh.GrhIndex).FileNum = 0 Then
        Call InitGrh(Grh, 23829) 'llamita= 32179) ' error=23829
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

    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2

        End If

        If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32

        End If

    End If

    Device_Box_Textured_Render CurrentGrhIndex, X, Y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, Light, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, Alpha, angle, Transp
    'exits:

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
    'ciertos efecto. (ejem: monta�as, reflejos y sombras)
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
                                          ByVal z As Single, _
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
    Geometry_Create_TLVertex.z = z
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
    'Author: Juan Mart�n Sotuyo Dodero
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
    
    Texto.Engine_Text_Draw 5, 5, PosX, vbWhite
    Texto.Engine_Text_Draw 5, 20, "FPS: " & FPS, vbWhite
    
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
    'Last modified by: Juan Mart�n Sotuyo Dodero (Maraxus)
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
    '******************************************************
    'Ultima modificacion 08/05/2020 por ReyarB
    '******************************************************
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
    
    'If we can, we render around the view area to make it smoother ReyarB ver error
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = screenminY - 1

        '        ScreenY = 1
    End If

    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = screenminX - 1

        '        ScreenX = 1
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
                Grh.GrhIndex = 2 ' Ver resultados Ultima modificacion 08/05/2020 por ReyarB
                Grh.FrameCounter = 1
                Grh.Started = 0
                Call Draw_Grh(Grh, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, LightIluminado(), , X, Y)

            End If
                      
            If VerTriggers Then '4978
                If MapData(X, Y).Trigger > 0 Then Texto.Engine_Text_Draw ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, MapData(X, Y).Trigger, vbWhite, , , 3

            End If
            
            If Seleccionando Then
                If X >= SeleccionIX And Y >= SeleccionIY Then
                    If X <= SeleccionFX And Y <= SeleccionFY Then
                        Grh.GrhIndex = 23828
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

Public Sub MapCapture(ByRef Format As Integer)

    '*************************************************
    'Author: Torres Patricio(Pato)
    'Ultima modificacion 08/05/2020 por ReyarB
    '*************************************************
    Dim D3DWindow        As D3DPRESENT_PARAMETERS

    Dim Y                As Byte     'Keeps track of where on map we are

    Dim X                As Byte     'Keeps track of where on map we are
          
    Dim RX               As Long

    Dim RY               As Long
          
    Dim PosX             As Integer

    Dim PosY             As Integer

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs
          
    Dim Grh              As Grh      'Temp Grh for show tile and blocked

    BY = 0
    BX = 0

    Static re As RECT
        
    Select Case Format
        
        Case 0 ' configuracion original Lorwik
            
            re.Left = 0
            re.Top = 0
            re.Bottom = XMaxMapSize * 32
            re.Right = YMaxMapSize * 32
            frmRender.picMap.Height = 200
            frmRender.picMap.Width = 200
            frmRender.Height = 4000
            frmRender.Width = 12000
            BY = 1
            BX = 1
                
        Case 1 'Modificado por ReyarB para generar minimapas
            re.Left = 0
            re.Top = 0
            re.Bottom = 100 * 32
            re.Right = 100 * 32
            frmRender.picMap.Height = 100
            frmRender.picMap.Width = 100
            frmRender.Height = 3000
            frmRender.Width = 12000
            BY = 1
            BX = 1
                
        Case 2 'Modificado por ReyarB para generar minimapas
            re.Left = 0
            re.Top = 0
            re.Bottom = 80 * 32
            re.Right = 74 * 32
            frmRender.picMap.Height = 100
            frmRender.picMap.Width = 100
            frmRender.Height = 3000
            frmRender.Width = 12000
            BY = 10
            BX = 13

    End Select

    frmRender.pgbProgress.value = 0
    frmRender.pgbProgress.Max = 500000
            
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
    D3DDevice.BeginScene
    'Draw floor layer

    For Y = BY To YMaxMapSize
        For X = BX To XMaxMapSize
                              
            RY = Y - BY
            RX = X - BX
                              
            PosX = (RX - 1) * 32 + TilePixelWidth
            PosY = (RY - 1) * 32 + TilePixelHeight
                                          
            'Layer 1 **********************************
                              
            If (MapData(X, Y).Graphic(1).GrhIndex <> 0) And VerCapa1 Then
                Call Draw_Grh(MapData(X, Y).Graphic(1), PosX, PosY, 1, 1, MapData(X, Y).light_value())

            End If

            '******************************************
            
            'frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1
            
            'Layer 2 **********************************

            If (MapData(X, Y).Graphic(2).GrhIndex <> 0) And VerCapa2 Then
                Call Draw_Grh(MapData(X, Y).Graphic(2), PosX, PosY, 1, 1, MapData(X, Y).light_value())

            End If

            '******************************************
            
            'frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1

            PixelOffsetXTemp = (X - 1) * 32 + TilePixelWidth
            PixelOffsetYTemp = (Y - 1) * 32 + TilePixelHeight
            
            With MapData(X, Y)
                'Object Layer **********************************

                If (.ObjGrh.GrhIndex <> 0) And VerObjetos Then
                    Call Draw_Grh(.ObjGrh, (RX - 1) * 32 + TilePixelWidth, (RY - 1) * 32 + TilePixelHeight, 1, 1, MapData(X, Y).light_value())

                End If

                '***********************************************
                Select Case Format

                    Case 0
                        'Layer 3 *****************************************

                        If (.Graphic(3).GrhIndex <> 0) And VerCapa3 Then
                            Call Draw_Grh(.Graphic(3), (RX - 1) * 32 + TilePixelWidth, (RY - 1) * 32 + TilePixelHeight, 1, 1, MapData(X, Y).light_value())

                        End If

                        '************************************************
                        'Ultima modificacion 08/05/2020 por ReyarB
                        '************************************************
                
                        'frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1
                    Case 1
 
                End Select
                
            End With
        
            Select Case Format

                Case 0

                    With MapData(X, Y)

                        'Layer 4 **********************************
                        If (.Graphic(4).GrhIndex <> 0) And VerCapa4 Then
                            'Draw
                            Call Draw_Grh(.Graphic(4), (RX - 1) * 32 + TilePixelWidth, (RY - 1) * 32 + TilePixelHeight, 1, 1, MapData(X, Y).light_value())

                        End If
        
                        '**********************************
                        'frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1
                    End With

                Case 1

                    With MapData(X, Y)
                        'PixelOffsetXTemp = (X - 1) * 32 + TilePixelWidth
                        'PixelOffsetYTemp = (Y - 1) * 32 + TilePixelHeight
                        '**********************************
                        Grh.FrameCounter = 1
                        Grh.Started = 0

                        If (.TileExit.Map <> 0) And VerTranslados Then
                            Grh.GrhIndex = 3
                            Call Draw_Grh(Grh, (RX - 1) * 32 + TilePixelWidth, (RY - 1) * 32 + TilePixelHeight, 1, 0, MapData(X, Y).light_value())

                        End If
            
                        'Show blocked tiles
                        If (.blocked = 1) And VerBlockeados Then
                            Grh.GrhIndex = 4
                            Call Draw_Grh(Grh, (RX - 1) * 32 + TilePixelWidth, (RY - 1) * 32 + TilePixelHeight, 1, 0, MapData(X, Y).light_value())

                        End If
        
                        '******************************************
                        ' frmRender.pgbProgress.value = frmRender.pgbProgress.value + 1
                    End With

            End Select
     
            D3DDevice.Present re, ByVal 0, frmRender.picMap.hWnd, ByVal 0
        Next X
    Next Y
          
    D3DDevice.EndScene
    D3DDevice.Present re, ByVal 0, frmRender.picMap.hWnd, ByVal 0
    D3DDevice.Present re, ByVal 0, frmMain.Minimap.hWnd, ByVal 0
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''Guardo la imagen'''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call frmRender.Capturar_Imagen(frmRender.picMap, frmRender.picMap)
        
    Select Case Format

        Case 0
            SavePicture frmRender.picMap, App.Path & "\Renderizados\" & NumMap_Save & ".bmp"

        Case 1
            SavePicture frmRender.picMap, App.Path & "\Recursos\Graficos\MiniMapa\" & NumMap_Save & ".bmp"

        Case 2
            SavePicture frmRender.picMap, App.Path & "\Renderizados\Minimapa\" & NumMap_Save & ".bmp"

    End Select
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
End Sub

'***********************************************
'Autor: Lorwik
'Fecha: 11/05/2020
'Descripcion: Ajusta los controles cuando se redimensiona la ventana
'***********************************************
Public Function SetHalfWindowTileHeight(ByVal Height As Integer)
    HalfWindowTileHeight = (Height / 32) \ 2

End Function

Public Function SetHalfWindowTileWidth(ByVal Width As Integer)
    HalfWindowTileWidth = (Width / 32) \ 2

End Function

