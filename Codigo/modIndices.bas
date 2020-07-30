Attribute VB_Name = "modIndices"
Option Explicit

Public GrhCount As Long

''
' Carga los indices de Graficos
'
Public Function LoadGrhData() As Boolean

    On Error GoTo ErrorHandler

    Dim Grh         As Long
    Dim Frame       As Long
    Dim handle      As Integer
    Dim fileVersion As Long
    Dim File        As String

    'Open files
    handle = FreeFile()
    Open DirIndex & "Graficos.ind" For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , GrhCount
    
    'Resize arrays
    ReDim GrhData(1 To GrhCount) As GrhData
    
    While Not EOF(handle)

        Get handle, , Grh

        If Grh <> 0 Then

            With GrhData(Grh)
            
                'Get number of frames
                Get handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                .Active = True
                
                ReDim .Frames(1 To .NumFrames)
                
                If .NumFrames > 1 Then

                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                    
                        Get handle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > GrhCount Then GoTo ErrorHandler

                    Next Frame
                    
                    Get handle, , .Speed
                    If .Speed <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                Else
                    'Read in normal GRH data
                    Get handle, , .FileNum
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , GrhData(Grh).sX
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .sY
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                End If
            End With
        End If
    Wend
    
    Close handle

    LoadGrhData = True
    
    Exit Function
 
ErrorHandler:
    LoadGrhData = False
    Debug.Print "Error en LoadGrhData... Grh: " & Grh

End Function


''
' Carga los indices de Superficie
'

Public Sub CargarIndicesSuperficie()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(IniPath & "Datos\indices.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Datos\indices.ini'", vbCritical
        End
    End If
    Dim Leer As New clsIniReader
    Dim i As Integer
    Leer.Initialize IniPath & "Datos\indices.ini"
    MaxSup = Leer.GetValue("INIT", "Referencias")
    ReDim SupData(MaxSup) As SupData
    frmMain.lListado(0).Clear
    For i = 0 To MaxSup
        SupData(i).name = Leer.GetValue("REFERENCIA" & i, "Nombre")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
        frmMain.lListado(0).AddItem SupData(i).name & " - #" & i
    Next
    
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de Datos\indices.ini" & vbCrLf & "Err: " & err.Number & " - " & err.Description, vbCritical + vbOKOnly
End Sub

Public Sub CargarBloqueables()
    If FileExist(IniPath & "Datos\Bloqueables.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Datos\Bloqueables.ini'", vbCritical
        End
    End If
    
    Dim Leer As New clsIniReader
    Dim i As Integer
    
    Leer.Initialize IniPath & "Datos\Bloqueables.ini"
    MaxBloqueables = Leer.GetValue("INIT", "Num")
    
    ReDim Bloqueables(MaxBloqueables) As Long
    
    
    For i = 1 To MaxBloqueables
        Bloqueables(i) = Leer.GetValue("BLOQUEABLES", i)
    Next i
End Sub

''
' Carga los indices de Objetos
'

Public Sub CargarIndicesOBJ()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(DirDats & "\OBJ.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'OBJ.dat' en " & DirDats, vbCritical
        End
    End If
    Dim Obj As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(DirDats & "\OBJ.dat")
    frmMain.lListado(3).Clear
    NumOBJs = Val(Leer.GetValue("INIT", "NumOBJs"))
    ReDim ObjData(1 To NumOBJs) As ObjData
    For Obj = 1 To NumOBJs
        frmCargando.X.Caption = "Cargando Datos de Objetos..." & Obj & "/" & NumOBJs
        DoEvents
        ObjData(Obj).name = Leer.GetValue("OBJ" & Obj, "Name")
        ObjData(Obj).GrhIndex = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
        ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
        ObjData(Obj).WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
        ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        ObjData(Obj).GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))
        frmMain.lListado(3).AddItem ObjData(Obj).name & " - #" & Obj
    Next Obj
    Exit Sub
Fallo:
MsgBox "Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & DirDats & vbCrLf & "Err: " & err.Number & " - " & err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Triggers
'

Public Sub CargarIndicesTriggers()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(App.Path & "\Datos\Triggers.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Triggers.ini' en \Datos\Triggers.ini", vbCritical
        End
    End If
    Dim NumT As Integer
    Dim T As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(App.Path & "\Datos\Triggers.ini")
    frmMain.lListado(4).Clear
    NumT = Val(Leer.GetValue("INIT", "NumTriggers"))
    For T = 1 To NumT
         frmMain.lListado(4).AddItem Leer.GetValue("Trig" & T, "Name") & " - #" & (T - 1)
    Next T

Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el Trigger " & T & " de Triggers.ini en " & App.Path & "\Datos\Triggers.ini" & vbCrLf & "Err: " & err.Number & " - " & err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Cuerpos
'

Public Sub CargarIndicesDeCuerpos()

    Dim n            As Integer
    Dim i            As Long
    Dim J            As Byte
    Dim File         As String
    Dim NumCuerpos   As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    n = FreeFile()
    Open DirIndex & "Personajes.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As tBodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #n, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then

            For J = 1 To 4
                InitGrh BodyData(i).Walk(J), MisCuerpos(i).Body(J), 0
            Next J

            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY

        End If

    Next i
    
    Close #n

End Sub

''
' Carga los indices de Cabezas
'

Public Sub CargarIndicesDeCabezas()

    Dim n            As Integer
    Dim i            As Long
    Dim J            As Byte
    Dim Miscabezas() As tIndiceCabeza
    Dim File         As String
    
    'Cabezas
    n = FreeFile()
    Open DirIndex & "Cabezas.ind" For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As tHeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #n, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then

            For J = 1 To 4
                Call InitGrh(HeadData(i).Head(J), Miscabezas(i).Head(J), 0)
            Next J

        End If

    Next i
    
    Close #n

End Sub

''
' Carga los indices de NPCs
'

Public Sub CargarIndicesNPC()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
On Error Resume Next
'On Error GoTo Fallo
    If FileExist(DirDats & "\NPCs.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & DirDats, vbCritical
        End
    End If
    'If FileExist(DirDats & "\NPCs-HOSTILES.dat", vbArchive) = False Then
    '    MsgBox "Falta el archivo 'NPCs-HOSTILES.dat' en " & DirDats, vbCritical
    '    End
    'End If
    Dim Trabajando As String
    Dim NPC As Integer
    Dim Leer As New clsIniReader
    frmMain.lListado(1).Clear
    frmMain.lListado(2).Clear
    Call Leer.Initialize(DirDats & "\NPCs.dat")
    NumNPCs = Val(Leer.GetValue("INIT", "NumNPCs"))
    'Call Leer.Initialize(DirDats & "\NPCs-HOSTILES.dat")
    'NumNPCsHOST = Val(Leer.GetValue("INIT", "NumNPCs"))
    ReDim NpcData(1000) As NpcData
    Trabajando = "Dats\NPCs.dat"
    'Call Leer.Initialize(DirDats & "\NPCs.dat")
    'MsgBox "  "
    For NPC = 1 To NumNPCs
        NpcData(NPC).name = Leer.GetValue("NPC" & NPC, "Name")
        
        NpcData(NPC).Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
        NpcData(NPC).Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
        NpcData(NPC).Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))
        If LenB(NpcData(NPC).name) <> 0 Then
            frmMain.lListado(1).AddItem NpcData(NPC).name & " - #" & NPC
        End If
    Next
    'MsgBox "  "
    'Trabajando = "Dats\NPCs-HOSTILES.dat"
    'Call Leer.Initialize(DirDats & "\NPCs-HOSTILES.dat")
    'For NPC = 1 To NumNPCsHOST
    '    NpcData(NPC + 499).name = Leer.GetValue("NPC" & (NPC + 499), "Name")
    '    NpcData(NPC + 499).Body = Val(Leer.GetValue("NPC" & (NPC + 499), "Body"))
    '    NpcData(NPC + 499).Head = Val(Leer.GetValue("NPC" & (NPC + 499), "Head"))
    '    NpcData(NPC + 499).Heading = Val(Leer.GetValue("NPC" & (NPC + 499), "Heading"))
    '    If LenB(NpcData(NPC + 499).name) <> 0 Then frmMain.lListado(2).AddItem NpcData(NPC + 499).name & " - #" & (NPC + 499)
    'Next NPC
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & " de " & Trabajando & " en " & DirDats & vbCrLf & "Err: " & err.Number & " - " & err.Description, vbCritical + vbOKOnly

End Sub

Public Sub LoadMiniMap()

    On Error GoTo ErrorHandler
    
    If Not FileExist(DirIndex & "minimap.dat", vbNormal) Then Exit Sub
    
    Dim File   As String
    Dim count  As Long
    Dim handle As Integer
        
    'Open files
    handle = FreeFile()
    
    Open DirIndex & "minimap.dat" For Binary As handle
    Seek handle, 1

    For count = 1 To GrhCount

        If GrhData(count).Active Then
            Get handle, , GrhData(count).MiniMap_color
        End If

    Next count

    Close handle
    
ErrorHandler:
    Debug.Print "Error en LoadMiniMap."

End Sub

Public Sub CargarParticulas()
    '*************************************
    'Coded by OneZero (onezero_ss@hotmail.com)
    'Last Modified: 6/4/03
    'Loads the Particles.ini file to the ComboBox
    'Edited by Juan Martín Sotuyo Dodero to add speed and life
    '*************************************
    
    On Error GoTo ErrorHandler
    
    Dim loopc      As Long
    Dim i          As Long
    Dim GrhListing As String
    Dim TempSet    As String
    Dim ColorSet   As Long
    Dim StreamFile As String
    Dim Leer       As New clsIniReader

    Call Leer.Initialize(App.Path & "\Recursos\Init\Particulas.ini")
    
    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    For loopc = 1 To TotalStreams

        With StreamData(loopc)
        
            .name = Leer.GetValue(Val(loopc), "Name")
        
            Call frmMain.lstParticle.AddItem(loopc & "-" & .name)
        
            .NumOfParticles = Leer.GetValue(Val(loopc), "NumOfParticles")
            .X1 = Leer.GetValue(Val(loopc), "X1")
            .Y1 = Leer.GetValue(Val(loopc), "Y1")
            .X2 = Leer.GetValue(Val(loopc), "X2")
            .Y2 = Leer.GetValue(Val(loopc), "Y2")
            .angle = Leer.GetValue(Val(loopc), "Angle")
            .vecx1 = Leer.GetValue(Val(loopc), "VecX1")
            .vecx2 = Leer.GetValue(Val(loopc), "VecX2")
            .vecy1 = Leer.GetValue(Val(loopc), "VecY1")
            .vecy2 = Leer.GetValue(Val(loopc), "VecY2")
            .life1 = Leer.GetValue(Val(loopc), "Life1")
            .life2 = Leer.GetValue(Val(loopc), "Life2")
            .friction = Leer.GetValue(Val(loopc), "Friction")
            .spin = Leer.GetValue(Val(loopc), "Spin")
            .spin_speedL = Leer.GetValue(Val(loopc), "Spin_SpeedL")
            .spin_speedH = Leer.GetValue(Val(loopc), "Spin_SpeedH")
            .AlphaBlend = Leer.GetValue(Val(loopc), "AlphaBlend")
            .gravity = Leer.GetValue(Val(loopc), "Gravity")
            .grav_strength = Leer.GetValue(Val(loopc), "Grav_Strength")
            .bounce_strength = Leer.GetValue(Val(loopc), "Bounce_Strength")
            .XMove = Leer.GetValue(Val(loopc), "XMove")
            .YMove = Leer.GetValue(Val(loopc), "YMove")
            .move_x1 = Leer.GetValue(Val(loopc), "move_x1")
            .move_x2 = Leer.GetValue(Val(loopc), "move_x2")
            .move_y1 = Leer.GetValue(Val(loopc), "move_y1")
            .move_y2 = Leer.GetValue(Val(loopc), "move_y2")
            .Radio = Val(Leer.GetValue(Val(loopc), "Radio"))
            .life_counter = Leer.GetValue(Val(loopc), "life_counter")
            .Speed = Val(Leer.GetValue(Val(loopc), "Speed"))
            .NumGrhs = Leer.GetValue(Val(loopc), "NumGrhs")
           
            ReDim .grh_list(1 To .NumGrhs)
            GrhListing = Leer.GetValue(Val(loopc), "Grh_List")
           
            For i = 1 To .NumGrhs
                .grh_list(i) = ReadField(Str(i), GrhListing, 44)
            Next i

            .grh_list(i - 1) = .grh_list(i - 1)

            For ColorSet = 1 To 4
                TempSet = Leer.GetValue(Val(loopc), "ColorSet" & ColorSet)
                .colortint(ColorSet - 1).R = ReadField(1, TempSet, 44)
                .colortint(ColorSet - 1).G = ReadField(2, TempSet, 44)
                .colortint(ColorSet - 1).B = ReadField(3, TempSet, 44)
            Next ColorSet
        
        End With
        
    Next loopc

    Set Leer = Nothing
    
    Exit Sub
    
ErrorHandler:

    If err.Number <> 0 Then
        
        Select Case err.Number
        
            Case 9
                Call MsgBox("Se han encontrado valores invalidos en el Particulas.ini - Index: " & loopc)
                Exit Sub
                
            Case 53
                Call MsgBox("No se encuentra el archivo Particulas.ini en Recursos\Init", vbApplicationModal)
                Exit Sub
                
        End Select

    End If
    
End Sub

