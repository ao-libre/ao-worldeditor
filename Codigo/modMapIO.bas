Attribute VB_Name = "modMapIO"
Option Explicit

'***************************
'Map format .CSM
'***************************
Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueados
    X As Integer
    y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    y As Integer
    Trigger As Integer
End Type

Private Type tDatosLuces
    X As Integer
    y As Integer
    light_value(3) As Long
    base_light(0 To 3) As Boolean 'Indica si el tile tiene luz propia.
End Type

Private Type tDatosParticulas
    X As Integer
    y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    y As Integer
    NPCIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    y As Integer
    objindex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDat
    map_name As String
    battle_mode As Boolean
    backup_mode As Boolean
    restrict_mode As String
    music_number As String
    zone As String
    terrain As String
    ambient As String
    lvlMinimo As String
    SePuedeDomar As Boolean
    ResuSinEfecto As Boolean
    MagiaSinEfecto As Boolean
    InviSinEfecto As Boolean
    NoEncriptarMP As Boolean
    version As Long
End Type

Public MapSize As tMapSize
Private MapDat As tMapDat

Private MapTitulo As String     ' GS > Almacena el titulo del mapa para el .dat

Public MapaCargado_Integer  As Boolean

''
' Obtener el tamaño de un archivo
'
' @param FileName Especifica el path del archivo
' @return   Nos devuelve el tamaño

Public Function FileSize(ByVal FileName As String) As Long
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

    On Error GoTo FalloFile
    Dim nFileNum As Integer
    Dim lFileSize As Long
    
    nFileNum = FreeFile
    Open FileName For Input As nFileNum
    lFileSize = LOF(nFileNum)
    Close nFileNum
    FileSize = lFileSize
    
    Exit Function
FalloFile:
    FileSize = -1
End Function

''
' Nos dice si existe el archivo/directorio
'
' @param file Especifica el path
' @param FileType Especifica el tipo de archivo/directorio
' @return   Nos devuelve verdadero o falso

Public Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************
If LenB(Dir(File, FileType)) = 0 Then
    FileExist = False
Else
    FileExist = True
End If

End Function

''
' Guarda el Mapa
'
' @param Path Especifica el path del mapa

Public Sub GuardarMapa(Optional Path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************

Call frmOptimizar.Optimizar

Call Resolucion
Call modEdicion.Bloquear_Bordes(1)

frmMain.Dialog.CancelError = True
On Error GoTo errhandler

If LenB(Path) = 0 Then
    frmMain.ObtenerNombreArchivo True
    Path = frmMain.Dialog.FileName
    If LenB(Path) = 0 Then Exit Sub
End If

If frmMain.Dialog.FilterIndex = 2 Then
    Call Save_CSM(Path)
ElseIf frmMain.Dialog.FilterIndex = 1 Then
    Call MapaV2_Guardar(Path)
        
    
End If

errhandler:
End Sub

''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional Path As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If MapInfo.Changed = 1 Then
    If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
        GuardarMapa Path
    End If
End If
End Sub


''
' Limpia todo el mapa a uno nuevo
'

Public Sub NuevoMapa()
'*************************************************
'Author: ^[GS]^
'Last modified: 21/05/06
'*************************************************

    On Error Resume Next
    
    Dim loopc As Integer
    Dim y As Integer
    Dim X As Integer
    
    bAutoGuardarMapaCount = 0
    
    'frmMain.mnuUtirialNuevoFormato.Checked = True
    frmMain.mnuReAbrirMapa.Enabled = False
    frmMain.TimAutoGuardarMapa.Enabled = False
    frmMain.lblMapVersion.Caption = 0
    frmMain.lblMapAmbient.Caption = 0
    
    MapaCargado = False
    
    For loopc = 0 To frmMain.MapPest.count - 1
        frmMain.MapPest(loopc).Enabled = False
    Next
    
    frmMain.MousePointer = 11
    
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            ' Capa 1
            MapData(X, y).Graphic(1).GrhIndex = 1
            
            ' Bloqueos
            MapData(X, y).blocked = 0
    
            ' Capas 2, 3 y 4
            MapData(X, y).Graphic(2).GrhIndex = 0
            MapData(X, y).Graphic(3).GrhIndex = 0
            MapData(X, y).Graphic(4).GrhIndex = 0
    
            ' NPCs
            If MapData(X, y).NPCIndex > 0 Then
                EraseChar MapData(X, y).CharIndex
                MapData(X, y).NPCIndex = 0
            End If
    
            ' OBJs
            MapData(X, y).OBJInfo.objindex = 0
            MapData(X, y).OBJInfo.Amount = 0
            MapData(X, y).ObjGrh.GrhIndex = 0
    
            ' Translados
            MapData(X, y).TileExit.Map = 0
            MapData(X, y).TileExit.X = 0
            MapData(X, y).TileExit.y = 0
            
            ' Triggers
            MapData(X, y).Trigger = 0
            
            MapData(X, y).particle_Index = 0
            Particle_Group_Remove MapData(X, y).particle_group_index
            
            InitGrh MapData(X, y).Graphic(1), 1
        Next X
    Next y
    
    MapInfo.MapVersion = 0
    MapInfo.Name = "Nuevo Mapa"
    MapInfo.Music = 0
    MapInfo.PK = True
    MapInfo.MagiaSinEfecto = 0
    MapInfo.InviSinEfecto = 0
    MapInfo.ResuSinEfecto = 0
    MapInfo.Terreno = "BOSQUE"
    MapInfo.Zona = "CAMPO"
    MapInfo.Restringir = "No"
    MapInfo.NoEncriptarMP = 0
    
    Call MapInfo_Actualizar
    
    Call ActualizaMinimap ' Radar
    
    'Set changed flag
    MapInfo.Changed = 0
    frmMain.MousePointer = 0
    
    ' Vacio deshacer
    modEdicion.Deshacer_Clear
    
    MapaCargado = True
    EngineRun = True
    
    frmMain.SetFocus

End Sub

''
' Guardar Mapa con el formato V2
'
' @param SaveAs Especifica donde guardar el mapa

Public Sub MapaV2_Guardar(ByVal SaveAs As String, Optional ByVal Preguntar As Boolean = True)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo ErrorSave

    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim loopc       As Long
    Dim TempInt     As Integer
    Dim y           As Long
    Dim X           As Long
    Dim ByFlags     As Byte

    If FileExist(SaveAs, vbNormal) = True Then
    
    If frmConvert.Check2.value = 0 Then
            If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            Else
                Call Kill(SaveAs)
            End If
        
        Else
            Call Kill(SaveAs)
            
        End If
        
    End If

    frmMain.MousePointer = 11

    ' y borramos el .inf tambien
    If FileExist(Left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Call Kill(Left$(SaveAs, Len(SaveAs) - 4) & ".inf")
    End If

    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1

    SaveAs = Left$(SaveAs, Len(SaveAs) - 4)
    SaveAs = SaveAs & ".inf"

    'Open .inf file
    FreeFileInf = FreeFile
    Open SaveAs For Binary As FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    
    ' Version del Mapa
    If frmMain.lblMapVersion.Caption < 32767 Then
        frmMain.lblMapVersion.Caption = frmMain.lblMapVersion + 1
        frmMapInfo.txtMapVersion = frmMain.lblMapVersion.Caption
    End If

    Put FreeFileMap, , CInt(frmMain.lblMapVersion.Caption)
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            With MapData(X, y)
            
                ByFlags = 0
                
                If .blocked = 1 Then ByFlags = ByFlags Or 1
                
                If .Graphic(2).GrhIndex Then ByFlags = ByFlags Or 2
                If .Graphic(3).GrhIndex Then ByFlags = ByFlags Or 4
                If .Graphic(4).GrhIndex Then ByFlags = ByFlags Or 8

                If .Trigger Then ByFlags = ByFlags Or 16
                
                If .particle_group_index Then ByFlags = ByFlags Or 32
                    
'                If MapData(X, Y).light_index Then
'                    Put FreeFileMap, , Lights(MapData(X, Y).light_index).Range
'                    R = Lights(MapData(X, Y).light_index).RGBCOLOR.R
'                    G = Lights(MapData(X, Y).light_index).RGBCOLOR.G
'                    B = Lights(MapData(X, Y).light_index).RGBCOLOR.B
'                    Put FreeFileMap, , R
'                    Put FreeFileMap, , G
'                    Put FreeFileMap, , B
'                End If
                    
                Put FreeFileMap, , ByFlags
                    
                If MapaCargado_Integer Then
                    Put FreeFileMap, , .Graphic(1).GrhIndexIntg
                Else
                    Put FreeFileMap, , .Graphic(1).GrhIndex
                End If
                
                For loopc = 2 To 4
                    
                    If MapaCargado_Integer Then
                        If .Graphic(loopc).GrhIndex Then Put FreeFileMap, , .Graphic(loopc).GrhIndexIntg
                    Else
                        If .Graphic(loopc).GrhIndex Then Put FreeFileMap, , .Graphic(loopc).GrhIndex
                    End If

                Next loopc
                    
                If .Trigger Then Put FreeFileMap, , .Trigger
                
                If .particle_group_index Then
                    Put FreeFileMap, , .particle_Index
                End If
                
                'Escribimos el archivo ".INF"
                ByFlags = 0
                    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                If .NPCIndex Then ByFlags = ByFlags Or 2
                
                If .OBJInfo.objindex Then ByFlags = ByFlags Or 4
                    
                Put FreeFileInf, , ByFlags
                    
                If .TileExit.Map Then
                    Put FreeFileInf, , .TileExit.Map
                    Put FreeFileInf, , .TileExit.X
                    Put FreeFileInf, , .TileExit.y
                End If
                    
                If .NPCIndex Then
                    Put FreeFileInf, , CInt(.NPCIndex)
                End If
                    
                If .OBJInfo.objindex Then
                    Put FreeFileInf, , .OBJInfo.objindex
                    Put FreeFileInf, , .OBJInfo.Amount
                End If
            
            End With
            
            
        Next X
    Next y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close FreeFileInf

    Call Pestañas(SaveAs, ".map")

    'write .dat file
    SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
    Call MapInfo_Guardar(SaveAs)

    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0

    Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & err.Number & " - " & err.Description

End Sub

''
' Abrir Mapa con el formato V2
'
' @param Map Especifica el Path del mapa

Public Sub MapaV2_Cargar(ByVal Map As String, Optional ByVal EsInteger As Boolean = False)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error Resume Next

    Dim loopc       As Integer
    Dim TempInt     As Integer
    Dim Body        As Integer
    Dim Head        As Integer
    Dim Heading     As Byte
    Dim y           As Integer
    Dim X           As Integer
    Dim i           As Byte
    Dim ByFlags     As Byte
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long

    DoEvents
    
    'Change mouse icon
    frmMain.MousePointer = 11
       
    'Con esto, le digo al WE que estamos usando mapas de tipo integer,
    'lo uso mas que nada para que no crashee cargar los mapas siguientes en las Pestañas.
    MapaCargado_Integer = EsInteger
    
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    Map = Left$(Map, Len(Map) - 4)
    Map = Map & ".inf"
    
    FreeFileInf = FreeFile
    Open Map For Binary As FreeFileInf
    Seek FreeFileInf, 1
    
    'Cabecera map
    Get FreeFileMap, , MapInfo.MapVersion
    Get FreeFileMap, , MiCabecera
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    Get FreeFileMap, , TempInt
    
    'Cabecera inf
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt
    Get FreeFileInf, , TempInt

    'Vamos a limpiar las luces y particulas del mapa anterior
    'Engine.Particle_Group_Remove_All
    
    'Load arrays Ver ReyarB
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            With MapData(X, y)
            
                For i = 0 To 3
                    If .light_value(i) = True Then
                        .light_value(i) = False
                    End If
                Next i
            
                Get FreeFileMap, , ByFlags
                .blocked = (ByFlags And 1)
            
                'Layer 1
                If EsInteger Then
                    Get FreeFileMap, , .Graphic(1).GrhIndexIntg
                    Call InitGrh(.Graphic(1), .Graphic(1).GrhIndexIntg)
                Else
                    Get FreeFileMap, , .Graphic(1).GrhIndex
                    Call InitGrh(.Graphic(1), .Graphic(1).GrhIndex)
                End If
            
                'Layer 2 used?
                If ByFlags And 2 Then
                    
                    If EsInteger Then
                        Get FreeFileMap, , .Graphic(2).GrhIndexIntg
                        Call InitGrh(.Graphic(2), .Graphic(2).GrhIndexIntg)
                    Else
                        Get FreeFileMap, , .Graphic(2).GrhIndex
                        Call InitGrh(.Graphic(2), .Graphic(2).GrhIndex)
                    End If
 
                Else
                
                    .Graphic(2).GrhIndex = 0
                    
                End If
                
                'Layer 3 used?
                If ByFlags And 4 Then
                    
                    If EsInteger Then
                        Get FreeFileMap, , .Graphic(3).GrhIndexIntg
                        Call InitGrh(.Graphic(3), .Graphic(3).GrhIndexIntg)
                    Else
                        Get FreeFileMap, , .Graphic(3).GrhIndex
                        Call InitGrh(.Graphic(3), .Graphic(3).GrhIndex)
                    End If

                Else
                
                    .Graphic(3).GrhIndex = 0
                    
                End If
                
                'Layer 4 used?
                If ByFlags And 8 Then
                    
                    If EsInteger Then
                        Get FreeFileMap, , .Graphic(4).GrhIndexIntg
                        Call InitGrh(.Graphic(3), .Graphic(3).GrhIndexIntg)
                    Else
                        Get FreeFileMap, , .Graphic(4).GrhIndex
                        Call InitGrh(.Graphic(4), .Graphic(4).GrhIndex)
                    End If

                Else
                    
                    .Graphic(4).GrhIndex = 0

                End If
             
                'Trigger used?
                If ByFlags And 16 Then
                    Get FreeFileMap, , .Trigger
                Else
                    .Trigger = 0
                End If
                
                'Particles used?
                If ByFlags And 32 Then
                   Get FreeFileMap, , TempInt
                    MapData(X, y).particle_group_index = General_Particle_Create(TempInt, X, y, -1)
                    MapData(X, y).particle_Index = TempInt
                End If
                
            
                'Cargamos el archivo ".INF"
                Get FreeFileInf, , ByFlags
            
                If ByFlags And 1 Then
                    
                    With .TileExit
                    
                        Get FreeFileInf, , .Map
                        Get FreeFileInf, , .X
                        Get FreeFileInf, , .y
                    
                    End With
                    

                End If
    
                If ByFlags And 2 Then
                
                    'Get and make NPC
                    Get FreeFileInf, , .NPCIndex
    
                    If .NPCIndex < 0 Then
                        .NPCIndex = 0
                    Else
                        Body = NpcData(.NPCIndex).Body
                        Head = NpcData(.NPCIndex).Head
                        Heading = NpcData(.NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, X, y)
                    End If

                End If
    
                If ByFlags And 4 Then
                    
                    'Get and make Object
                    Get FreeFileInf, , .OBJInfo.objindex
                    Get FreeFileInf, , .OBJInfo.Amount

                    If .OBJInfo.objindex > 0 Then
                        Call InitGrh(.ObjGrh, ObjData(.OBJInfo.objindex).GrhIndex)
                    End If

                End If
            
            End With
    
        Next X
    Next y
    
    'Close files
    Close FreeFileMap
    Close FreeFileInf
    
    Call Pestañas(Map, ".map")
    
    Map = Left$(Map, Len(Map) - 4) & ".dat"
    
    Call MapInfo_Cargar(Map)
    
    With frmMain
    
        frmMain.lblMapVersion.Caption = MapInfo.MapVersion
    
        ' Avisamos que estamos trabajando con un mapa de tipo integer.
        If EsInteger Then
            .Caption = App.Title & " - Mapa Integer"
        Else
            .Caption = App.Title & " - Mapa Long"
        End If
        
        'Set changed flag
        MapInfo.Changed = 0
        
        ' Vacia el Deshacer
        Call modEdicion.Deshacer_Clear
        
        'Change mouse icon
        .MousePointer = 0
        
        CurMap = ReturnNumberFromString(Map)
        
        Call DibujarMiniMapa
' *****************************************************************************
' Renderizado mapa 200x200 ****************************************************
' *****************************************************************************
        
        If frmMain.chkRenderizarAl.value = 1 Then
            frmRender.Show
            MapCapture (0)
            frmRender.Hide
        End If
 ' *****************************************************************************
' Renderizado mapa 200x200 ****************************************************
' *****************************************************************************
    End With
    
    MapaCargado = True

End Sub

' *****************************************************************************
' MAPINFO *********************************************************************
' *****************************************************************************

''
' Guardar Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Guardar(ByVal Archivo As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

    If LenB(MapTitulo) = 0 Then
        MapTitulo = NameMap_Save
    End If

    Call WriteVar(Archivo, MapTitulo, "Name", MapInfo.Name)
    Call WriteVar(Archivo, MapTitulo, "MusicNumMp3", MapInfo.Music)
    Call WriteVar(Archivo, MapTitulo, "MagiaSinefecto", Val(MapInfo.MagiaSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "InviSinEfecto", Val(MapInfo.InviSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "ResuSinEfecto", Val(MapInfo.ResuSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "NoEncriptarMP", Val(MapInfo.NoEncriptarMP))

    Call WriteVar(Archivo, MapTitulo, "Terreno", MapInfo.Terreno)
    Call WriteVar(Archivo, MapTitulo, "Zona", MapInfo.Zona)
    Call WriteVar(Archivo, MapTitulo, "Restringir", MapInfo.Restringir)
    Call WriteVar(Archivo, MapTitulo, "BackUp", Str(MapInfo.BackUp))

    If MapInfo.PK Then
        Call WriteVar(Archivo, MapTitulo, "Pk", "0")
    Else
        Call WriteVar(Archivo, MapTitulo, "Pk", "1")
    End If
End Sub

''
' Abrir Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Cargar(ByVal Archivo As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 02/06/06
'*************************************************

On Error Resume Next
    Dim Leer As New clsIniReader
    Dim loopc As Integer
    Dim Path As String
    MapTitulo = Empty
    
    If FileExist(Archivo, vbNormal) = False Then
        Call AddtoRichTextBox(frmMain.StatTxt, "Error en el Mapa " & Archivo & ", no se ha encontrado al archivo .dat", 255, 0, 0)
     Exit Sub
    End If
    
    Leer.Initialize Archivo

    For loopc = Len(Archivo) To 1 Step -1
        If mid(Archivo, loopc, 1) = "\" Then
            Path = Left(Archivo, loopc)
            Exit For
        End If
    Next
    Archivo = Right(Archivo, Len(Archivo) - (Len(Path)))
    MapTitulo = UCase(Left(Archivo, Len(Archivo) - 4))

    MapInfo.Name = Leer.GetValue(MapTitulo, "Name")
    MapInfo.Music = Leer.GetValue(MapTitulo, "MusicNumMp3")
    MapInfo.MagiaSinEfecto = Val(Leer.GetValue(MapTitulo, "MagiaSinEfecto"))
    MapInfo.InviSinEfecto = Val(Leer.GetValue(MapTitulo, "InviSinEfecto"))
    MapInfo.ResuSinEfecto = Val(Leer.GetValue(MapTitulo, "ResuSinEfecto"))
    MapInfo.NoEncriptarMP = Val(Leer.GetValue(MapTitulo, "NoEncriptarMP"))
    
    If Val(Leer.GetValue(MapTitulo, "Pk")) = 0 Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False
    End If
    
    MapInfo.Terreno = Leer.GetValue(MapTitulo, "Terreno")
    MapInfo.Zona = Leer.GetValue(MapTitulo, "Zona")
    MapInfo.Restringir = Leer.GetValue(MapTitulo, "Restringir")
    MapInfo.BackUp = Val(Leer.GetValue(MapTitulo, "BACKUP"))
    
    Call MapInfo_Actualizar
    
End Sub

''
' Actualiza el formulario de MapInfo
'

Public Sub MapInfo_Actualizar()
'*************************************************
'Author: ^[GS]^
'Last modified: 02/06/06
'*************************************************

On Error Resume Next
    ' Mostrar en Formularios
    frmMapInfo.txtMapNombre.Text = MapInfo.Name
    frmMapInfo.txtMapMusica.Text = MapInfo.Music
    frmMapInfo.txtMapTerreno.Text = MapInfo.Terreno
    frmMapInfo.txtMapZona.Text = MapInfo.Zona
    frmMapInfo.txtMapRestringir.Text = MapInfo.Restringir
'    frmMapInfo.chkMapBackup.value = MapInfo.BackUp
    frmMapInfo.chkMapMagiaSinEfecto.value = MapInfo.MagiaSinEfecto
'    frmMapInfo.chkMapInviSinEfecto.value = MapInfo.InviSinEfecto
'    frmMapInfo.ChkMapNpc.value = MapInfo.SePuedeDomar
    frmMapInfo.chkMapResuSinEfecto.value = MapInfo.ResuSinEfecto
    frmMapInfo.chkMapNoEncriptarMP.value = MapInfo.NoEncriptarMP
    frmMapInfo.chkMapPK.value = IIf(MapInfo.PK = True, 1, 0)
    frmMapInfo.txtMapVersion = MapInfo.MapVersion
    frmMain.lblMapNombre = MapInfo.Name
    frmMain.lblMapMusica = MapInfo.Music
    frmMain.lblMapAmbient = MapInfo.ambient
    frmMapInfo.TxtlvlMinimo = MapInfo.lvlMinimo
    frmMapInfo.TxtAmbient = MapInfo.ambient

End Sub

''
' Calcula la orden de Pestañas
'
' @param Map Especifica path del mapa

Public Sub Pestañas(ByVal Map As String, ByVal formato As String)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
On Error Resume Next
Dim loopc As Integer

For loopc = Len(Map) To 1 Step -1
    If mid(Map, loopc, 1) = "\" Then
        PATH_Save = Left(Map, loopc)
        Exit For
    End If
Next
Map = Right(Map, Len(Map) - (Len(PATH_Save)))
For loopc = Len(Left(Map, Len(Map) - 4)) To 1 Step -1
    If IsNumeric(mid(Left(Map, Len(Map) - 4), loopc, 1)) = False Then
        NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - loopc)
        NameMap_Save = Left(Map, loopc)
        Exit For
    End If
Next
For loopc = (NumMap_Save - 7) To (NumMap_Save + 10)
        If FileExist(PATH_Save & NameMap_Save & loopc & formato, vbArchive) = True Then
            frmMain.MapPest(loopc - NumMap_Save + 7).Visible = True
            frmMain.MapPest(loopc - NumMap_Save + 7).Enabled = True
            frmMain.MapPest(loopc - NumMap_Save + 7).Caption = NameMap_Save & loopc
        Else
            frmMain.MapPest(loopc - NumMap_Save + 7).Visible = False
        End If
Next
End Sub

Sub Cargar_CSM(ByVal Map As String)

'Particle_Group_Remove_All
'Light_Remove_All

On Error GoTo ErrorHandler

Dim fh As Integer
Dim File As Integer
Dim MH As tMapHeader
Dim Blqs() As tDatosBloqueados
Dim L1() As Long
Dim L2() As tDatosGrh
Dim L3() As tDatosGrh
Dim L4() As tDatosGrh
Dim Triggers() As tDatosTrigger
Dim Luces() As tDatosLuces
Dim Particulas() As tDatosParticulas
Dim Objetos() As tDatosObjs
Dim NPCs() As tDatosNPC
Dim TEs() As tDatosTE

Dim i As Long
Dim J As Long
DoEvents
    
'Change mouse icon
frmMain.MousePointer = 11
    
fh = FreeFile
Open Map For Binary Access Read As fh
    Get #fh, , MH
    Get #fh, , MapSize
    '¿Queremos cargar un mapa de IAO 1.4?
    Get #fh, , MapDat
    
    With MapSize
        If Not .XMax = XMaxMapSize Or Not .YMax = YMaxMapSize Then
            ReDim MapData(.XMin To .XMax, .YMin To .YMax)
        End If
        ReDim L1(.XMin To .XMax, .YMin To .YMax)
    End With
    
    Get #fh, , L1
    
    With MH
        If .NumeroBloqueados > 0 Then
            ReDim Blqs(1 To .NumeroBloqueados)
            Get #fh, , Blqs
            For i = 1 To .NumeroBloqueados
                MapData(Blqs(i).X, Blqs(i).y).blocked = 1
            Next i
        End If
        
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2
            For i = 1 To .NumeroLayers(2)
                InitGrh MapData(L2(i).X, L2(i).y).Graphic(2), L2(i).GrhIndex
            Next i
        End If
        
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3
            For i = 1 To .NumeroLayers(3)
                InitGrh MapData(L3(i).X, L3(i).y).Graphic(3), L3(i).GrhIndex
            Next i
        End If
        
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4
            For i = 1 To .NumeroLayers(4)
                InitGrh MapData(L4(i).X, L4(i).y).Graphic(4), L4(i).GrhIndex
            Next i
        End If
        
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers
            For i = 1 To .NumeroTriggers
                MapData(Triggers(i).X, Triggers(i).y).Trigger = Triggers(i).Trigger
            Next i
        End If
        
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas
            For i = 1 To .NumeroParticulas
                MapData(Particulas(i).X, Particulas(i).y).particle_group_index = General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).y)
            Next i
        End If
            
'        If .NumeroLuces > 0 Then
'            ReDim Luces(1 To .NumeroLuces)
'            Dim p As Byte
'            Get #fh, , Luces
'            For i = 1 To .NumeroLuces
'                For p = 0 To 3
'                    MapData(Luces(i).X, Luces(i).Y).base_light(p) = Luces(i).base_light(p)
'                    If MapData(Luces(i).X, Luces(i).Y).base_light(p) Then _
'                        MapData(Luces(i).X, Luces(i).Y).light_value(p) = Luces(i).light_value(p)
'                Next p
'            Next i
'        End If
            
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos
            For i = 1 To .NumeroOBJs
                MapData(Objetos(i).X, Objetos(i).y).OBJInfo.objindex = Objetos(i).objindex
                MapData(Objetos(i).X, Objetos(i).y).OBJInfo.Amount = Objetos(i).ObjAmmount
                If MapData(Objetos(i).X, Objetos(i).y).OBJInfo.objindex > NumOBJs Then
                    InitGrh MapData(Objetos(i).X, Objetos(i).y).ObjGrh, 23829
                Else
                    InitGrh MapData(Objetos(i).X, Objetos(i).y).ObjGrh, ObjData(MapData(Objetos(i).X, Objetos(i).y).OBJInfo.objindex).GrhIndex
                End If
            Next i
        End If
            
        If .NumeroNPCs > 0 Then
            ReDim NPCs(1 To .NumeroNPCs)
            Get #fh, , NPCs
            For i = 1 To .NumeroNPCs
                If NPCs(i).NPCIndex > 0 Then
                    MapData(NPCs(i).X, NPCs(i).y).NPCIndex = NPCs(i).NPCIndex
                    Call MakeChar(NextOpenChar(), NpcData(NPCs(i).NPCIndex).Body, NpcData(NPCs(i).NPCIndex).Head, NpcData(NPCs(i).NPCIndex).Heading, NPCs(i).X, NPCs(i).y)
                End If
            Next i
        End If

        If .NumeroTE > 0 Then
            ReDim TEs(1 To .NumeroTE)
            Get #fh, , TEs
            For i = 1 To .NumeroTE
                MapData(TEs(i).X, TEs(i).y).TileExit.Map = TEs(i).DestM
                MapData(TEs(i).X, TEs(i).y).TileExit.X = TEs(i).DestX
                MapData(TEs(i).X, TEs(i).y).TileExit.y = TEs(i).DestY
            Next i
        End If
        
    End With

Close fh


For J = MapSize.YMin To MapSize.YMax
    For i = MapSize.XMin To MapSize.XMax
        If L1(i, J) > 0 Then
            InitGrh MapData(i, J).Graphic(1), L1(i, J)
        End If
    Next i
Next J



'*******************************
'Render lights
'Light_Render_All
'*******************************

    
    
    'MapInfo_Cargar Map
    frmMain.lblMapVersion.Caption = MapInfo.MapVersion
    
    Call Pestañas(Map, ".csm")
    
    ' Vacia el Deshacer
    modEdicion.Deshacer_Clear
    
    'Change mouse icon
    frmMain.MousePointer = 0
    
    Call CSMInfoCargar
    
    'Set changed flag
    MapInfo.Changed = 0

    MapaCargado = True
    
    Call DibujarMiniMapa ' Radar
    
    If frmMain.chkRenderizarAl.value = 1 Then
        frmRender.Show
        MapCapture (1)
        frmRender.Hide
   End If
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa " & Map & " cargado...", 43, 0, 255)
ErrorHandler:
    If fh <> 0 Then Close fh
    Call AddtoRichTextBox(frmMain.StatTxt, "Error en el Mapa " & Map & ", se ha generado un informe de errores en: " & App.Path & "\Logs.txt", 255, 0, 0)
    File = FreeFile
    Open App.Path & "\Logs.txt" For Output As #File
        Print #File, err.Description
    Close #File
End Sub


Public Function Save_CSM(ByVal MapRoute As String) As Boolean

On Error GoTo ErrorHandler

Dim fh As Integer
Dim MH As tMapHeader
Dim Blqs() As tDatosBloqueados
Dim L1() As Long
Dim L2() As tDatosGrh
Dim L3() As tDatosGrh
Dim L4() As tDatosGrh
Dim Triggers() As tDatosTrigger
Dim Luces() As tDatosLuces
Dim Particulas() As tDatosParticulas
Dim Objetos() As tDatosObjs
Dim NPCs() As tDatosNPC
Dim TEs() As tDatosTE

Dim i As Integer
Dim J As Integer

If FileExist(MapRoute, vbNormal) = True Then
    If MsgBox("¿Desea sobrescribir " & MapRoute & "?", vbCritical + vbYesNo) = vbNo Then
        Exit Function
    Else
        Kill MapRoute
    End If
End If

frmMain.MousePointer = 11
MapSize.XMax = XMaxMapSize
MapSize.YMax = YMaxMapSize
ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax)

For J = MapSize.YMin To MapSize.YMax
    For i = MapSize.XMin To MapSize.XMax
        With MapData(i, J)
            If .blocked Then
                MH.NumeroBloqueados = MH.NumeroBloqueados + 1
                ReDim Preserve Blqs(1 To MH.NumeroBloqueados)
                Blqs(MH.NumeroBloqueados).X = i
                Blqs(MH.NumeroBloqueados).y = J
            End If
            
            L1(i, J) = .Graphic(1).GrhIndex
            
            If .Graphic(2).GrhIndex > 0 Then
                MH.NumeroLayers(2) = MH.NumeroLayers(2) + 1
                ReDim Preserve L2(1 To MH.NumeroLayers(2))
                L2(MH.NumeroLayers(2)).X = i
                L2(MH.NumeroLayers(2)).y = J
                L2(MH.NumeroLayers(2)).GrhIndex = .Graphic(2).GrhIndex
            End If
            
            If .Graphic(3).GrhIndex > 0 Then
                MH.NumeroLayers(3) = MH.NumeroLayers(3) + 1
                ReDim Preserve L3(1 To MH.NumeroLayers(3))
                L3(MH.NumeroLayers(3)).X = i
                L3(MH.NumeroLayers(3)).y = J
                L3(MH.NumeroLayers(3)).GrhIndex = .Graphic(3).GrhIndex
            End If
            
            If .Graphic(4).GrhIndex > 0 Then
                MH.NumeroLayers(4) = MH.NumeroLayers(4) + 1
                ReDim Preserve L4(1 To MH.NumeroLayers(4))
                L4(MH.NumeroLayers(4)).X = i
                L4(MH.NumeroLayers(4)).y = J
                L4(MH.NumeroLayers(4)).GrhIndex = .Graphic(4).GrhIndex
            End If
            
            If .Trigger > 0 Then
                MH.NumeroTriggers = MH.NumeroTriggers + 1
                ReDim Preserve Triggers(1 To MH.NumeroTriggers)
                Triggers(MH.NumeroTriggers).X = i
                Triggers(MH.NumeroTriggers).y = J
                Triggers(MH.NumeroTriggers).Trigger = .Trigger
            End If
            
            If .particle_group_index > 0 Then
                MH.NumeroParticulas = MH.NumeroParticulas + 1
                ReDim Preserve Particulas(1 To MH.NumeroParticulas)
                Particulas(MH.NumeroParticulas).X = i
                Particulas(MH.NumeroParticulas).y = J
                Particulas(MH.NumeroParticulas).Particula = CLng(particle_group_list(.particle_group_index).stream_type)
            End If
           
'            If .base_light(0) Or .base_light(1) _
'                    Or .base_light(2) Or .base_light(3) Then
'                MH.NumeroLuces = MH.NumeroLuces + 1
'                ReDim Preserve Luces(1 To MH.NumeroLuces)
'                Dim p As Byte
'                Luces(MH.NumeroLuces).X = i
'                Luces(MH.NumeroLuces).Y = J
'                For p = 0 To 3
'                    Luces(MH.NumeroLuces).base_light(p) = .base_light(p)
'                    If .base_light(p) Then _
'                        Luces(MH.NumeroLuces).light_value(p) = .light_value(p)
'                Next p
'            End If
            
            If .OBJInfo.objindex > 0 Then
                MH.NumeroOBJs = MH.NumeroOBJs + 1
                ReDim Preserve Objetos(1 To MH.NumeroOBJs)
                Objetos(MH.NumeroOBJs).objindex = .OBJInfo.objindex
                Objetos(MH.NumeroOBJs).ObjAmmount = .OBJInfo.Amount
                Objetos(MH.NumeroOBJs).X = i
                Objetos(MH.NumeroOBJs).y = J
            End If
            
            If .NPCIndex > 0 Then
                MH.NumeroNPCs = MH.NumeroNPCs + 1
                ReDim Preserve NPCs(1 To MH.NumeroNPCs)
                NPCs(MH.NumeroNPCs).NPCIndex = .NPCIndex
                NPCs(MH.NumeroNPCs).X = i
                NPCs(MH.NumeroNPCs).y = J
            End If
            
            If .TileExit.Map > 0 Then
                MH.NumeroTE = MH.NumeroTE + 1
                ReDim Preserve TEs(1 To MH.NumeroTE)
                TEs(MH.NumeroTE).DestM = .TileExit.Map
                TEs(MH.NumeroTE).DestX = .TileExit.X
                TEs(MH.NumeroTE).DestY = .TileExit.y
                TEs(MH.NumeroTE).X = i
                TEs(MH.NumeroTE).y = J
            End If
        End With
    Next i
Next J

Call CSMInfoSave
          
fh = FreeFile
Open MapRoute For Binary As fh
    
    Put #fh, , MH
    Put #fh, , MapSize
    Put #fh, , MapDat
    Put #fh, , L1

    With MH
        If .NumeroBloqueados > 0 Then _
            Put #fh, , Blqs
        If .NumeroLayers(2) > 0 Then _
            Put #fh, , L2
        If .NumeroLayers(3) > 0 Then _
            Put #fh, , L3
        If .NumeroLayers(4) > 0 Then _
            Put #fh, , L4
        If .NumeroTriggers > 0 Then _
            Put #fh, , Triggers
        If .NumeroParticulas > 0 Then _
            Put #fh, , Particulas
        If .NumeroLuces > 0 Then _
            Put #fh, , Luces
        If .NumeroOBJs > 0 Then _
            Put #fh, , Objetos
        If .NumeroNPCs > 0 Then _
            Put #fh, , NPCs
        If .NumeroTE > 0 Then _
            Put #fh, , TEs
    End With

Close fh

Call Pestañas(MapRoute, ".csm")

'Change mouse icon
frmMain.MousePointer = 0
MapInfo.Changed = 0

Save_CSM = True

 Call AddtoRichTextBox(frmMain.StatTxt, "Mapa " & MapRoute & " guardado...", 43, 0, 255)
Exit Function

ErrorHandler:
    If fh <> 0 Then Close fh

End Function


Public Sub CSMInfoSave()
    MapDat.map_name = MapInfo.Name
    MapDat.music_number = MapInfo.Music
    MapDat.MagiaSinEfecto = MapInfo.MagiaSinEfecto
    MapDat.InviSinEfecto = MapInfo.InviSinEfecto
    MapDat.ResuSinEfecto = MapInfo.ResuSinEfecto
    MapDat.NoEncriptarMP = MapInfo.NoEncriptarMP
    MapDat.SePuedeDomar = MapInfo.SePuedeDomar
    MapDat.lvlMinimo = MapInfo.lvlMinimo
    MapDat.version = MapInfo.MapVersion
    
    If MapInfo.PK = True Then
        MapDat.battle_mode = True
    Else
        MapDat.battle_mode = False
    End If
    
    MapDat.ambient = MapInfo.ambient
    
    MapDat.terrain = MapInfo.Terreno
    MapDat.zone = MapInfo.Zona
    MapDat.restrict_mode = MapInfo.Restringir
    MapDat.backup_mode = MapInfo.BackUp
End Sub

Public Sub CSMInfoCargar()
    MapInfo.Name = MapDat.map_name
    MapInfo.Music = MapDat.music_number
    MapInfo.MagiaSinEfecto = MapDat.MagiaSinEfecto
    MapInfo.InviSinEfecto = MapDat.InviSinEfecto
    MapInfo.ResuSinEfecto = MapDat.ResuSinEfecto
    MapInfo.NoEncriptarMP = MapDat.NoEncriptarMP
    MapInfo.SePuedeDomar = MapDat.SePuedeDomar
    MapInfo.lvlMinimo = MapDat.lvlMinimo
    MapInfo.MapVersion = MapDat.version
    
    If MapDat.battle_mode = True Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False
    End If
    
    MapInfo.ambient = MapDat.ambient
    
    MapInfo.Terreno = MapDat.terrain
    MapInfo.Zona = MapDat.zone
    MapInfo.Restringir = MapDat.restrict_mode
    MapInfo.BackUp = MapDat.backup_mode
    
    Call MapInfo_Actualizar
End Sub
Public Sub Resolucion()

    If frmMain.Option2.value = True Then
        ClienteHeight = 13
        ClienteWidth = 17
        Else
        ClienteHeight = 19
        ClienteWidth = 24
    End If
    
    'MinXBorder = XMinMapSize + (Round(700 / 32) \ 2) '700 = Width render cliente
    'MaxXBorder = XMaxMapSize - (Round(700 / 32) \ 2)
    'MinYBorder = YMinMapSize + (Round(524 / 32) \ 2) '524 = Heigth render cliente
    'MaxYBorder = YMaxMapSize - (Round(524 / 32) \ 2)
   
    MinXBorder = XMinMapSize + (ClienteWidth \ 2)
    MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
    MinYBorder = YMinMapSize + (ClienteHeight \ 2)
    MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)

End Sub

''
' Guardar Mapa con el formato V5
'
' @param SaveAs Especifica donde guardar el mapa

Public Sub MapaV5_Guardar(ByVal SaveAs As String, Optional ByVal Preguntar As Boolean)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo ErrorSave

    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim loopc       As Long
    Dim TempInt     As Integer
    Dim y           As Long
    Dim X           As Long
    Dim ByFlags     As Byte
    
    'If frmConvert.Check2.value = 0 Then Preguntar = False
    
    

    If FileExist(SaveAs, vbNormal) = True Then
        
        If frmConvert.Check2.value = 0 Then
            If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
                Exit Sub
            Else
                Call Kill(SaveAs)
            End If
        
        Else
            Call Kill(SaveAs)
            
        End If
        
    End If

    frmMain.MousePointer = 11

    ' y borramos el .inf tambien
    If FileExist(Left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Call Kill(Left$(SaveAs, Len(SaveAs) - 4) & ".inf")
    End If

    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1

    SaveAs = Left$(SaveAs, Len(SaveAs) - 4)
    SaveAs = SaveAs & ".inf"

    'Open .inf file
    FreeFileInf = FreeFile
    Open SaveAs For Binary As FreeFileInf
    Seek FreeFileInf, 1

    'map Header
    
    ' Version del Mapa
    If frmMain.lblMapVersion.Caption < 32767 Then
        frmMain.lblMapVersion.Caption = frmMain.lblMapVersion + 1
        frmMapInfo.txtMapVersion = frmMain.lblMapVersion.Caption
    End If

    Put FreeFileMap, , CInt(frmMain.lblMapVersion.Caption)
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    Put FreeFileInf, , TempInt
    
    'Write .map file
    For y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            With MapData(X, y)
            
                ByFlags = 0
                
                If .blocked = 1 Then ByFlags = ByFlags Or 1
                
                If .Graphic(2).GrhIndex Then ByFlags = ByFlags Or 2
                If .Graphic(3).GrhIndex Then ByFlags = ByFlags Or 4
                If .Graphic(4).GrhIndex Then ByFlags = ByFlags Or 8

                If .Trigger Then ByFlags = ByFlags Or 16
                                   
                Put FreeFileMap, , ByFlags
                    
                If MapaCargado_Integer Then
                    Put FreeFileMap, , .Graphic(1).GrhIndexIntg
                Else
                    Put FreeFileMap, , .Graphic(1).GrhIndex
                End If
                
                For loopc = 2 To 4
                    
                    If MapaCargado_Integer Then
                        If .Graphic(loopc).GrhIndex Then Put FreeFileMap, , .Graphic(loopc).GrhIndexIntg
                    Else
                        If .Graphic(loopc).GrhIndex Then Put FreeFileMap, , .Graphic(loopc).GrhIndex
                    End If

                Next loopc
                    
                If .Trigger Then Put FreeFileMap, , .Trigger
                
                'Escribimos el archivo ".INF"
                ByFlags = 0
                    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                If .NPCIndex Then ByFlags = ByFlags Or 2
                
                If .OBJInfo.objindex Then ByFlags = ByFlags Or 4
                    
                Put FreeFileInf, , ByFlags
                    
                If .TileExit.Map Then
                    Put FreeFileInf, , .TileExit.Map
                    Put FreeFileInf, , .TileExit.X
                    Put FreeFileInf, , .TileExit.y
                End If
                    
                If .NPCIndex Then
                    Put FreeFileInf, , CInt(.NPCIndex)
                End If
                    
                If .OBJInfo.objindex Then
                    Put FreeFileInf, , .OBJInfo.objindex
                    Put FreeFileInf, , .OBJInfo.Amount
                End If
            
            End With
            
            
        Next X
    Next y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close FreeFileInf

    Call Pestañas(SaveAs, ".map")

    'write .dat file
    SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
    Call MapInfo_Guardar(SaveAs)

    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0

    Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV5, nro. " & err.Number & " - " & err.Description

End Sub


