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
    Y As Integer

End Type

Private Type tDatosGrh

    X As Integer
    Y As Integer
    GrhIndex As Long

End Type

Private Type tDatosTrigger

    X As Integer
    Y As Integer
    Trigger As Integer

End Type

Private Type tDatosLuces

    X As Integer
    Y As Integer
    light_value(3) As Long
    base_light(0 To 3) As Boolean 'Indica si el tile tiene luz propia.

End Type

Private Type tDatosParticulas

    X As Integer
    Y As Integer
    Particula As Long

End Type

Private Type tDatosNPC

    X As Integer
    Y As Integer
    NPCIndex As Integer

End Type

Private Type tDatosObjs

    X As Integer
    Y As Integer
    objindex As Integer
    ObjAmmount As Integer

End Type

Private Type tDatosTE

    X As Integer
    Y As Integer
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

Public MapSize    As tMapSize

Private MapDat    As tMapDat

Private MapTitulo As String     ' GS > Almacena el titulo del mapa para el .dat

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

    Dim nFileNum  As Integer

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

Public Function FileExist(ByVal File As String, _
                          ByVal FileType As VbFileAttribute) As Boolean

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

    frmMain.Dialog.CancelError = True

    On Error GoTo ErrHandler

    If LenB(Path) = 0 Then
        
        Call frmMain.ObtenerNombreArchivo(True)
        
        Path = frmMain.Dialog.FileName

        If LenB(Path) = 0 Then Exit Sub

    End If
    
    Select Case frmMain.Dialog.FilterIndex
    
        Case 1
            Call MapaV2_Guardar(Path)
        Case 2
            Call Save_CSM(Path)
            
    End Select

ErrHandler:

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

    Dim LoopC As Integer

    Dim Y     As Integer

    Dim X     As Integer

    bAutoGuardarMapaCount = 0

    'frmMain.mnuUtirialNuevoFormato.Checked = True
    frmMain.mnuReAbrirMapa.Enabled = False
    frmMain.TimAutoGuardarMapa.Enabled = False
    frmMain.lblMapVersion.Caption = 0
    frmMain.lblMapAmbient.Caption = 0

    MapaCargado = False

    For LoopC = 0 To frmMain.MapPest.count - 1
        frmMain.MapPest(LoopC).Enabled = False
    Next

    frmMain.MousePointer = 11

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            ' Capa 1
            MapData(X, Y).Graphic(1).GrhIndex = 1
        
            ' Bloqueos
            MapData(X, Y).blocked = 0

            ' Capas 2, 3 y 4
            MapData(X, Y).Graphic(2).GrhIndex = 0
            MapData(X, Y).Graphic(3).GrhIndex = 0
            MapData(X, Y).Graphic(4).GrhIndex = 0

            ' NPCs
            If MapData(X, Y).NPCIndex > 0 Then
                EraseChar MapData(X, Y).CharIndex
                MapData(X, Y).NPCIndex = 0

            End If

            ' OBJs
            MapData(X, Y).OBJInfo.objindex = 0
            MapData(X, Y).OBJInfo.Amount = 0
            MapData(X, Y).ObjGrh.GrhIndex = 0

            ' Translados
            MapData(X, Y).TileExit.Map = 0
            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0
        
            ' Triggers
            MapData(X, Y).Trigger = 0
        
            MapData(X, Y).particle_Index = 0
        
            InitGrh MapData(X, Y).Graphic(1), 1
        Next X
    Next Y

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

Public Sub MapaV2_Guardar(ByVal SaveAs As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo ErrorSave

    Dim FreeFileMap As Long

    Dim FreeFileInf As Long

    Dim LoopC       As Long

    Dim TempInt     As Integer

    Dim Y           As Long

    Dim X           As Long

    Dim ByFlags     As Byte

    If FileExist(SaveAs, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill SaveAs

        End If

    End If

    frmMain.MousePointer = 11

    ' y borramos el .inf tambien
    If FileExist(Left(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Kill Left(SaveAs, Len(SaveAs) - 4) & ".inf"

    End If

    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1

    SaveAs = Left(SaveAs, Len(SaveAs) - 4)
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
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            ByFlags = 0
                
            If MapData(X, Y).blocked = 1 Then ByFlags = ByFlags Or 1
            If MapData(X, Y).Graphic(2).GrhIndex Then ByFlags = ByFlags Or 2
            If MapData(X, Y).Graphic(3).GrhIndex Then ByFlags = ByFlags Or 4
            If MapData(X, Y).Graphic(4).GrhIndex Then ByFlags = ByFlags Or 8
            If MapData(X, Y).Trigger Then ByFlags = ByFlags Or 16
                
            Put FreeFileMap, , ByFlags
                
            Put FreeFileMap, , MapData(X, Y).Graphic(1).GrhIndex
                
            For LoopC = 2 To 4

                If MapData(X, Y).Graphic(LoopC).GrhIndex Then Put FreeFileMap, , MapData(X, Y).Graphic(LoopC).GrhIndex
            Next LoopC
                
            If MapData(X, Y).Trigger Then Put FreeFileMap, , MapData(X, Y).Trigger
                
            '.inf file
                
            ByFlags = 0
                
            If MapData(X, Y).TileExit.Map Then ByFlags = ByFlags Or 1
            If MapData(X, Y).NPCIndex Then ByFlags = ByFlags Or 2
            If MapData(X, Y).OBJInfo.objindex Then ByFlags = ByFlags Or 4
                
            Put FreeFileInf, , ByFlags
                
            If MapData(X, Y).TileExit.Map Then
                Put FreeFileInf, , MapData(X, Y).TileExit.Map
                Put FreeFileInf, , MapData(X, Y).TileExit.X
                Put FreeFileInf, , MapData(X, Y).TileExit.Y

            End If
                
            If MapData(X, Y).NPCIndex Then
                
                Put FreeFileInf, , CInt(MapData(X, Y).NPCIndex)

            End If
                
            If MapData(X, Y).OBJInfo.objindex Then
                Put FreeFileInf, , MapData(X, Y).OBJInfo.objindex
                Put FreeFileInf, , MapData(X, Y).OBJInfo.Amount

            End If
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close FreeFileInf

    Call Pestañas(SaveAs, ".map")

    'write .dat file
    SaveAs = Left$(SaveAs, Len(SaveAs) - 4) & ".dat"
    MapInfo_Guardar SaveAs

    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0

    Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & Err.Number & " - " & Err.Description

End Sub

''
' Abrir Mapa con el formato V2
'
' @param Map Especifica el Path del mapa

Public Sub MapaV2_Cargar(ByVal Map As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error Resume Next

    Dim LoopC       As Integer
    Dim TempInt     As Integer
    Dim Body        As Integer
    Dim Head        As Integer
    Dim Heading     As Byte
    Dim Y           As Integer
    Dim X           As Integer
    Dim ByFlags     As Byte
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long

    DoEvents
    
    'Change mouse icon
    frmMain.MousePointer = 11
       
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    Map = Left(Map, Len(Map) - 4)
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
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            Get FreeFileMap, , ByFlags
            
            MapData(X, Y).blocked = (ByFlags And 1)
            
            Get FreeFileMap, , MapData(X, Y).Graphic(1).GrhIndex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get FreeFileMap, , MapData(X, Y).Graphic(2).GrhIndex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            Else
                MapData(X, Y).Graphic(2).GrhIndex = 0

            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get FreeFileMap, , MapData(X, Y).Graphic(3).GrhIndex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            Else
                MapData(X, Y).Graphic(3).GrhIndex = 0

            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get FreeFileMap, , MapData(X, Y).Graphic(4).GrhIndex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
            Else
                MapData(X, Y).Graphic(4).GrhIndex = 0

            End If
             
            'Trigger used?
            If ByFlags And 16 Then
                Get FreeFileMap, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0

            End If
            
            '.inf file
            Get FreeFileInf, , ByFlags
            
            If ByFlags And 1 Then
                Get FreeFileInf, , MapData(X, Y).TileExit.Map
                Get FreeFileInf, , MapData(X, Y).TileExit.X
                Get FreeFileInf, , MapData(X, Y).TileExit.Y

            End If
    
            If ByFlags And 2 Then
                'Get and make NPC
                Get FreeFileInf, , MapData(X, Y).NPCIndex
    
                If MapData(X, Y).NPCIndex < 0 Then
                    MapData(X, Y).NPCIndex = 0
                Else
                    Body = NpcData(MapData(X, Y).NPCIndex).Body
                    Head = NpcData(MapData(X, Y).NPCIndex).Head
                    Heading = NpcData(MapData(X, Y).NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)

                End If

            End If
    
            If ByFlags And 4 Then
                'Get and make Object
                Get FreeFileInf, , MapData(X, Y).OBJInfo.objindex
                Get FreeFileInf, , MapData(X, Y).OBJInfo.Amount

                If MapData(X, Y).OBJInfo.objindex > 0 Then
                    InitGrh MapData(X, Y).ObjGrh, ObjData(MapData(X, Y).OBJInfo.objindex).GrhIndex

                End If

            End If
    
        Next X
    Next Y
    
    'Close files
    Close FreeFileMap
    Close FreeFileInf
    
    Call Pestañas(Map, ".map")
    
    Call ActualizaMinimap ' Radar
    
    Map = Left$(Map, Len(Map) - 4) & ".dat"
    
    MapInfo_Cargar Map
    frmMain.lblMapVersion.Caption = MapInfo.MapVersion
    
    'Set changed flag
    MapInfo.Changed = 0
    
    ' Vacia el Deshacer
    modEdicion.Deshacer_Clear
    
    'Change mouse icon
    frmMain.MousePointer = 0
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
    Call WriteVar(Archivo, MapTitulo, "MusicNum", MapInfo.Music)
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

    Dim Leer  As New clsIniManager
    Dim LoopC As Integer
    Dim Path  As String

    MapTitulo = Empty
    
    If FileExist(Archivo, vbNormal) = False Then
        Call AddtoRichTextBox(frmMain.StatTxt, "Error en el Mapa " & Archivo & ", no se ha encontrado al archivo .dat", 255, 0, 0)
        Exit Sub

    End If
    
    Call Leer.Initialize(Archivo)

    For LoopC = Len(Archivo) To 1 Step -1

        If mid(Archivo, LoopC, 1) = "\" Then
            Path = Left(Archivo, LoopC)
            Exit For

        End If

    Next
    Archivo = Right(Archivo, Len(Archivo) - (Len(Path)))
    MapTitulo = UCase(Left(Archivo, Len(Archivo) - 4))

    MapInfo.Name = Leer.GetValue(MapTitulo, "Name")
    MapInfo.Music = Leer.GetValue(MapTitulo, "MusicNum")
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
    frmMapInfo.chkMapMagiaSinEfecto.Value = MapInfo.MagiaSinEfecto
    '    frmMapInfo.chkMapInviSinEfecto.value = MapInfo.InviSinEfecto
    '    frmMapInfo.ChkMapNpc.value = MapInfo.SePuedeDomar
    frmMapInfo.chkMapResuSinEfecto.Value = MapInfo.ResuSinEfecto
    frmMapInfo.chkMapNoEncriptarMP.Value = MapInfo.NoEncriptarMP
    frmMapInfo.chkMapPK.Value = IIf(MapInfo.PK = True, 1, 0)
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

    Dim LoopC As Integer

    For LoopC = Len(Map) To 1 Step -1

        If mid(Map, LoopC, 1) = "\" Then
            PATH_Save = Left(Map, LoopC)
            Exit For

        End If

    Next
    Map = Right(Map, Len(Map) - (Len(PATH_Save)))

    For LoopC = Len(Left(Map, Len(Map) - 4)) To 1 Step -1

        If IsNumeric(mid(Left(Map, Len(Map) - 4), LoopC, 1)) = False Then
            NumMap_Save = Right(Left(Map, Len(Map) - 4), Len(Left(Map, Len(Map) - 4)) - LoopC)
            NameMap_Save = Left(Map, LoopC)
            Exit For

        End If

    Next

    For LoopC = (NumMap_Save - 7) To (NumMap_Save + 10)

        If FileExist(PATH_Save & NameMap_Save & LoopC & formato, vbArchive) = True Then
            frmMain.MapPest(LoopC - NumMap_Save + 7).Visible = True
            frmMain.MapPest(LoopC - NumMap_Save + 7).Enabled = True
            frmMain.MapPest(LoopC - NumMap_Save + 7).Caption = NameMap_Save & LoopC
        Else
            frmMain.MapPest(LoopC - NumMap_Save + 7).Visible = False

        End If

    Next

End Sub

Sub Cargar_CSM(ByVal Map As String)

    'Particle_Group_Remove_All
    'Light_Remove_All

    On Error GoTo ErrorHandler

    Dim fh           As Integer

    Dim File         As Integer

    Dim MH           As tMapHeader

    Dim Blqs()       As tDatosBloqueados

    Dim L1()         As Long

    Dim L2()         As tDatosGrh

    Dim L3()         As tDatosGrh

    Dim L4()         As tDatosGrh

    Dim Triggers()   As tDatosTrigger

    Dim Luces()      As tDatosLuces

    Dim Particulas() As tDatosParticulas

    Dim Objetos()    As tDatosObjs

    Dim NPCs()       As tDatosNPC

    Dim TEs()        As tDatosTE

    Dim i            As Long

    Dim j            As Long

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
        ReDim MapData(.XMin To .XMax, .YMin To .YMax)
        ReDim L1(.XMin To .XMax, .YMin To .YMax)

    End With
    
    Get #fh, , L1
    
    With MH

        If .NumeroBloqueados > 0 Then
            ReDim Blqs(1 To .NumeroBloqueados)
            Get #fh, , Blqs

            For i = 1 To .NumeroBloqueados
                MapData(Blqs(i).X, Blqs(i).Y).blocked = 1
            Next i

        End If
        
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2

            For i = 1 To .NumeroLayers(2)
                InitGrh MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).GrhIndex
            Next i

        End If
        
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3

            For i = 1 To .NumeroLayers(3)
                InitGrh MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).GrhIndex
            Next i

        End If
        
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4

            For i = 1 To .NumeroLayers(4)
                InitGrh MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).GrhIndex
            Next i

        End If
        
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers

            For i = 1 To .NumeroTriggers
                MapData(Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
            Next i

        End If
        
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas

            For i = 1 To .NumeroParticulas
                MapData(Particulas(i).X, Particulas(i).Y).particle_group_index = General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).Y)
            Next i

        End If
            
        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)

            Dim p As Byte

            Get #fh, , Luces
            'For i = 1 To .NumeroLuces
            'For p = 0 To 3
            'MapData(Luces(i).X, Luces(i).y).base_light(p) = Luces(i).base_light(p)
            'If MapData(Luces(i).X, Luces(i).y).base_light(p) Then _
             MapData(Luces(i).X, Luces(i).y).light_value(p) = Luces(i).light_value(p)

            'Next p
            'Next i
        End If
            
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos

            For i = 1 To .NumeroOBJs
                MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.objindex = Objetos(i).objindex
                MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.Amount = Objetos(i).ObjAmmount

                If MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.objindex > NumOBJs Then
                    InitGrh MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, 20299
                Else
                    InitGrh MapData(Objetos(i).X, Objetos(i).Y).ObjGrh, ObjData(MapData(Objetos(i).X, Objetos(i).Y).OBJInfo.objindex).GrhIndex

                End If

            Next i

        End If
            
        If .NumeroNPCs > 0 Then
            ReDim NPCs(1 To .NumeroNPCs)
            Get #fh, , NPCs

            For i = 1 To .NumeroNPCs

                If NPCs(i).NPCIndex > 0 Then
                    MapData(NPCs(i).X, NPCs(i).Y).NPCIndex = NPCs(i).NPCIndex
                    Call MakeChar(NextOpenChar(), NpcData(NPCs(i).NPCIndex).Body, NpcData(NPCs(i).NPCIndex).Head, NpcData(NPCs(i).NPCIndex).Heading, NPCs(i).X, NPCs(i).Y)

                End If

            Next i

        End If

        If .NumeroTE > 0 Then
            ReDim TEs(1 To .NumeroTE)
            Get #fh, , TEs

            For i = 1 To .NumeroTE
                MapData(TEs(i).X, TEs(i).Y).TileExit.Map = TEs(i).DestM
                MapData(TEs(i).X, TEs(i).Y).TileExit.X = TEs(i).DestX
                MapData(TEs(i).X, TEs(i).Y).TileExit.Y = TEs(i).DestY
            Next i

        End If
        
    End With

    Close fh

    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax

            If L1(i, j) > 0 Then
                InitGrh MapData(i, j).Graphic(1), L1(i, j)

            End If

        Next i
    Next j

    '*******************************
    'Render lights
    'Light_Render_All
    '*******************************

    Call DibujarMiniMapa ' Radar
    
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
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa " & Map & " cargado...", 0, 255, 0)
ErrorHandler:

    If fh <> 0 Then Close fh
    Call AddtoRichTextBox(frmMain.StatTxt, "Error en el Mapa " & Map & ", se ha generado un informe de errores en: " & App.Path & "\Logs.txt", 255, 0, 0)
    File = FreeFile
    Open App.Path & "\Logs.txt" For Output As #File
    Print #File, Err.Description
    Close #File

End Sub

Public Function Save_CSM(ByVal MapRoute As String) As Boolean

    On Error GoTo ErrorHandler

    Dim fh           As Integer

    Dim MH           As tMapHeader

    Dim Blqs()       As tDatosBloqueados

    Dim L1()         As Long

    Dim L2()         As tDatosGrh

    Dim L3()         As tDatosGrh

    Dim L4()         As tDatosGrh

    Dim Triggers()   As tDatosTrigger

    Dim Luces()      As tDatosLuces

    Dim Particulas() As tDatosParticulas

    Dim Objetos()    As tDatosObjs

    Dim NPCs()       As tDatosNPC

    Dim TEs()        As tDatosTE

    Dim i            As Integer

    Dim j            As Integer

    If FileExist(MapRoute, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & MapRoute & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Function
        Else
            Kill MapRoute

        End If

    End If

    frmMain.MousePointer = 11

    ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax)

    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax

            With MapData(i, j)

                If .blocked Then
                    MH.NumeroBloqueados = MH.NumeroBloqueados + 1
                    ReDim Preserve Blqs(1 To MH.NumeroBloqueados)
                    Blqs(MH.NumeroBloqueados).X = i
                    Blqs(MH.NumeroBloqueados).Y = j

                End If
            
                L1(i, j) = .Graphic(1).GrhIndex
            
                If .Graphic(2).GrhIndex > 0 Then
                    MH.NumeroLayers(2) = MH.NumeroLayers(2) + 1
                    ReDim Preserve L2(1 To MH.NumeroLayers(2))
                    L2(MH.NumeroLayers(2)).X = i
                    L2(MH.NumeroLayers(2)).Y = j
                    L2(MH.NumeroLayers(2)).GrhIndex = .Graphic(2).GrhIndex

                End If
            
                If .Graphic(3).GrhIndex > 0 Then
                    MH.NumeroLayers(3) = MH.NumeroLayers(3) + 1
                    ReDim Preserve L3(1 To MH.NumeroLayers(3))
                    L3(MH.NumeroLayers(3)).X = i
                    L3(MH.NumeroLayers(3)).Y = j
                    L3(MH.NumeroLayers(3)).GrhIndex = .Graphic(3).GrhIndex

                End If
            
                If .Graphic(4).GrhIndex > 0 Then
                    MH.NumeroLayers(4) = MH.NumeroLayers(4) + 1
                    ReDim Preserve L4(1 To MH.NumeroLayers(4))
                    L4(MH.NumeroLayers(4)).X = i
                    L4(MH.NumeroLayers(4)).Y = j
                    L4(MH.NumeroLayers(4)).GrhIndex = .Graphic(4).GrhIndex

                End If
            
                If .Trigger > 0 Then
                    MH.NumeroTriggers = MH.NumeroTriggers + 1
                    ReDim Preserve Triggers(1 To MH.NumeroTriggers)
                    Triggers(MH.NumeroTriggers).X = i
                    Triggers(MH.NumeroTriggers).Y = j
                    Triggers(MH.NumeroTriggers).Trigger = .Trigger

                End If
            
                If .particle_group_index > 0 Then
                    MH.NumeroParticulas = MH.NumeroParticulas + 1
                    ReDim Preserve Particulas(1 To MH.NumeroParticulas)
                    Particulas(MH.NumeroParticulas).X = i
                    Particulas(MH.NumeroParticulas).Y = j
                    Particulas(MH.NumeroParticulas).Particula = CLng(particle_group_list(.particle_group_index).stream_type)

                End If
           
                'If .base_light(0) Or .base_light(1) _
                '        Or .base_light(2) Or .base_light(3) Then
                '    MH.NumeroLuces = MH.NumeroLuces + 1
                '    ReDim Preserve Luces(1 To MH.NumeroLuces)
                '    Dim p As Byte
                '    Luces(MH.NumeroLuces).X = i
                '    Luces(MH.NumeroLuces).y = j
                '    For p = 0 To 3
                '        Luces(MH.NumeroLuces).base_light(p) = .base_light(p)
                '        If .base_light(p) Then _
                '            Luces(MH.NumeroLuces).light_value(p) = .light_value(p)
                '    Next p
                'End If
            
                If .OBJInfo.objindex > 0 Then
                    MH.NumeroOBJs = MH.NumeroOBJs + 1
                    ReDim Preserve Objetos(1 To MH.NumeroOBJs)
                    Objetos(MH.NumeroOBJs).objindex = .OBJInfo.objindex
                    Objetos(MH.NumeroOBJs).ObjAmmount = .OBJInfo.Amount
                    Objetos(MH.NumeroOBJs).X = i
                    Objetos(MH.NumeroOBJs).Y = j

                End If
            
                If .NPCIndex > 0 Then
                    MH.NumeroNPCs = MH.NumeroNPCs + 1
                    ReDim Preserve NPCs(1 To MH.NumeroNPCs)
                    NPCs(MH.NumeroNPCs).NPCIndex = .NPCIndex
                    NPCs(MH.NumeroNPCs).X = i
                    NPCs(MH.NumeroNPCs).Y = j

                End If
            
                If .TileExit.Map > 0 Then
                    MH.NumeroTE = MH.NumeroTE + 1
                    ReDim Preserve TEs(1 To MH.NumeroTE)
                    TEs(MH.NumeroTE).DestM = .TileExit.Map
                    TEs(MH.NumeroTE).DestX = .TileExit.X
                    TEs(MH.NumeroTE).DestY = .TileExit.Y
                    TEs(MH.NumeroTE).X = i
                    TEs(MH.NumeroTE).Y = j

                End If

            End With

        Next i
    Next j

    Call CSMInfoSave
          
    fh = FreeFile
    Open MapRoute For Binary As fh
    
    Put #fh, , MH
    Put #fh, , MapSize
    Put #fh, , MapDat
    Put #fh, , L1

    With MH

        If .NumeroBloqueados > 0 Then Put #fh, , Blqs

        If .NumeroLayers(2) > 0 Then Put #fh, , L2

        If .NumeroLayers(3) > 0 Then Put #fh, , L3

        If .NumeroLayers(4) > 0 Then Put #fh, , L4

        If .NumeroTriggers > 0 Then Put #fh, , Triggers

        If .NumeroParticulas > 0 Then Put #fh, , Particulas

        If .NumeroLuces > 0 Then Put #fh, , Luces

        If .NumeroOBJs > 0 Then Put #fh, , Objetos

        If .NumeroNPCs > 0 Then Put #fh, , NPCs

        If .NumeroTE > 0 Then Put #fh, , TEs

    End With

    Close fh

    Call Pestañas(MapRoute, ".csm")

    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0

    Save_CSM = True

    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa " & MapRoute & " guardado...", 0, 255, 0)
    Exit Function

ErrorHandler:

    If fh <> 0 Then Close fh

End Function

Public Sub MapaInteger_Cargar(ByVal Map As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error Resume Next

    Dim LoopC       As Integer

    Dim TempInt     As Integer

    Dim Body        As Integer

    Dim Head        As Integer

    Dim Heading     As Byte

    Dim Y           As Integer

    Dim X           As Integer

    Dim ByFlags     As Byte

    Dim FreeFileMap As Long

    Dim FreeFileInf As Long

    DoEvents
    
    'Change mouse icon
    frmMain.MousePointer = 11
       
    'Open files
    FreeFileMap = FreeFile
    Open Map For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    Map = Left(Map, Len(Map) - 4)
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
            
    Particle_Group_Remove_All
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
    
            Dim i As Byte

            For i = 0 To 3

                If MapData(X, Y).light_value(i) = True Then
                    MapData(X, Y).light_value(i) = False

                End If

            Next i
    
            Get FreeFileMap, , ByFlags
            
            MapData(X, Y).blocked = (ByFlags And 1)
            
            Get FreeFileMap, , MapData(X, Y).Graphic(1).GrhIndexIntg
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndexIntg
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get FreeFileMap, , MapData(X, Y).Graphic(2).GrhIndexIntg
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndexIntg
            Else
                MapData(X, Y).Graphic(2).GrhIndexIntg = 0

            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get FreeFileMap, , MapData(X, Y).Graphic(3).GrhIndexIntg
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndexIntg
            Else
                MapData(X, Y).Graphic(3).GrhIndexIntg = 0

            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get FreeFileMap, , MapData(X, Y).Graphic(4).GrhIndexIntg
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndexIntg
            Else
                MapData(X, Y).Graphic(4).GrhIndexIntg = 0

            End If
             
            'Trigger used?
            If ByFlags And 16 Then
                Get FreeFileMap, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0

            End If
            
            If ByFlags And 32 Then
                Get FreeFileMap, , TempInt
                MapData(X, Y).particle_group_index = General_Particle_Create(TempInt, X, Y, -1)
                MapData(X, Y).particle_Index = TempInt

            End If
            
            If ByFlags And 64 Then

                '  Get FreeFileMap, , MapData(X, y).base_light(0)
                '  Get FreeFileMap, , MapData(X, y).base_light(1)
                '  Get FreeFileMap, , MapData(X, y).base_light(2)
                '  Get FreeFileMap, , MapData(X, y).base_light(3)
                '
                '  If MapData(X, y).base_light(0) Then _
                '      Get FreeFileMap, , MapData(X, y).light_value(0)
                '
                '  If MapData(X, y).base_light(1) Then _
                '      Get FreeFileMap, , MapData(X, y).light_value(1)
                '
                '  If MapData(X, y).base_light(2) Then _
                '      Get FreeFileMap, , MapData(X, y).light_value(2)
                '
                '  If MapData(X, y).base_light(3) Then _
                '      Get FreeFileMap, , MapData(X, y).light_value(3)
            End If
            
            '.inf file
            Get FreeFileInf, , ByFlags
            
            If ByFlags And 1 Then
                Get FreeFileInf, , MapData(X, Y).TileExit.Map
                Get FreeFileInf, , MapData(X, Y).TileExit.X
                Get FreeFileInf, , MapData(X, Y).TileExit.Y

            End If
    
            If ByFlags And 2 Then
                'Get and make NPC
                Get FreeFileInf, , MapData(X, Y).NPCIndex
    
                If MapData(X, Y).NPCIndex < 0 Then
                    MapData(X, Y).NPCIndex = 0
                Else
                    Body = NpcData(MapData(X, Y).NPCIndex).Body
                    Head = NpcData(MapData(X, Y).NPCIndex).Head
                    Heading = NpcData(MapData(X, Y).NPCIndex).Heading
                    Call MakeChar(NextOpenChar(), Body, Head, Heading, X, Y)

                End If

            End If
    
            If ByFlags And 4 Then
                'Get and make Object
                Get FreeFileInf, , MapData(X, Y).OBJInfo.objindex
                Get FreeFileInf, , MapData(X, Y).OBJInfo.Amount

                If MapData(X, Y).OBJInfo.objindex > 0 Then
                    InitGrh MapData(X, Y).ObjGrh, ObjData(MapData(X, Y).OBJInfo.objindex).GrhIndex

                End If

            End If
    
        Next X
    Next Y
    
    'Close files
    Close FreeFileMap
    Close FreeFileInf
    
    Call Pestañas(Map, ".map")

    Call DibujarMiniMapa ' Radar
    
    Map = Left$(Map, Len(Map) - 4) & ".dat"
    
    MapInfo_Cargar Map
    frmMain.lblMapVersion.Caption = MapInfo.MapVersion
    
    'Set changed flag
    MapInfo.Changed = 0
    
    ' Vacia el Deshacer
    modEdicion.Deshacer_Clear
    
    'Change mouse icon
    frmMain.MousePointer = 0
    
    MapaCargado = True
    
End Sub

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
