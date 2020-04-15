Attribute VB_Name = "modGeneral"
Option Explicit

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_DISPLAYFREQUENCY = &H400000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Public InitPath As String

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

''
' Realiza acciones de desplasamiento segun las teclas que hallamos precionado
'

Public Sub CheckKeys()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************

    If HotKeysAllow = False Then Exit Sub
    '[Loopzer]
    'If GetKeyState(vbKeyControl) < 0 Then
    '    If Seleccionando Then
    '        If GetKeyState(vbKeyC) < 0 Then CopiarSeleccion
    '        If GetKeyState(vbKeyX) < 0 Then CortarSeleccion
    '        If GetKeyState(vbKeyB) < 0 Then BlockearSeleccion
    '        If GetKeyState(vbKeyD) < 0 Then AccionSeleccion
    ''    Else
    '        If GetKeyState(vbKeyS) < 0 Then DePegar ' GS
    '        If GetKeyState(vbKeyV) < 0 Then PegarSeleccion
    '    End If
    'End If
    '[/Loopzer]
    
    If GetKeyState(vbKeyUp) < 0 Then
        
        If UserPos.Y < 1 Then Exit Sub ' 10
        
        If LegalPos(UserPos.X, UserPos.Y - 1) And WalkMode = True Then
            
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            
            UserPos.Y = UserPos.Y - 1
            
            Call MoveCharbyPos(UserCharIndex, UserPos.X, UserPos.Y)
            
            dLastWalk = GetTickCount
        
        ElseIf WalkMode = False Then
            
            UserPos.Y = UserPos.Y - 1
        
        End If
        
        bRefreshRadar = True ' Radar
        
        Call ActualizaMinimap

        frmMain.SetFocus
        
        Exit Sub
    End If

    If GetKeyState(vbKeyRight) < 0 Then
        
        If UserPos.X > 100 Then Exit Sub ' 89
        
        If LegalPos(UserPos.X + 1, UserPos.Y) And WalkMode = True Then
            
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            
            UserPos.X = UserPos.X + 1
            
            Call MoveCharbyPos(UserCharIndex, UserPos.X, UserPos.Y)
            
            dLastWalk = GetTickCount
        
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X + 1
        
        End If
        
        bRefreshRadar = True ' Radar
        
        Call ActualizaMinimap
        
        frmMain.SetFocus
        
        Exit Sub
        
    End If

    If GetKeyState(vbKeyDown) < 0 Then
        
        If UserPos.Y > 100 Then Exit Sub ' 92
        
        If LegalPos(UserPos.X, UserPos.Y + 1) And WalkMode = True Then
            
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            
            UserPos.Y = UserPos.Y + 1
            
            Call MoveCharbyPos(UserCharIndex, UserPos.X, UserPos.Y)
            
            dLastWalk = GetTickCount
            
        ElseIf WalkMode = False Then
            UserPos.Y = UserPos.Y + 1
        
        End If
        
        bRefreshRadar = True ' Radar
        
        Call ActualizaMinimap

        frmMain.SetFocus
        
        Exit Sub
        
    End If

    If GetKeyState(vbKeyLeft) < 0 Then
        
        If UserPos.X < 1 Then Exit Sub ' 12
        
        If LegalPos(UserPos.X - 1, UserPos.Y) And WalkMode = True Then
            
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            
            UserPos.X = UserPos.X - 1
            
            Call MoveCharbyPos(UserCharIndex, UserPos.X, UserPos.Y)
            
            dLastWalk = GetTickCount
        
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X - 1
        End If
        
        bRefreshRadar = True ' Radar
        
        Call ActualizaMinimap
 
        frmMain.SetFocus
        
        Exit Sub
        
    End If
    
End Sub

Public Function ReadField(Pos As Integer, Text As String, SepASCII As Integer) As String

    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    Dim i         As Integer

    Dim LastPos   As Integer

    Dim CurChar   As String * 1

    Dim FieldNum  As Integer

    Dim Seperator As String

    Seperator = Chr(SepASCII)
    LastPos = 0
    FieldNum = 0

    For i = 1 To Len(Text)
        CurChar = mid$(Text, i, 1)

        If CurChar = Seperator Then
            FieldNum = FieldNum + 1

            If FieldNum = Pos Then
                ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function

            End If

            LastPos = i

        End If

    Next i

    FieldNum = FieldNum + 1

    If FieldNum = Pos Then
        ReadField = mid$(Text, LastPos + 1)

    End If

End Function

''
' Completa y corrije un path
'
' @param Path Especifica el path con el que se trabajara
' @return   Nos devuelve el path completado

Private Function autoCompletaPath(ByVal Path As String) As String
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    Path = Replace(Path, "/", "\")

    If Left$(Path, 1) = "\" Then
        ' agrego app.path & path
        Path = App.Path & Path

    End If

    If Right$(Path, 1) <> "\" Then
        ' me aseguro que el final sea con "\"
        Path = Path & "\"

    End If

    autoCompletaPath = Path

End Function

''
' Carga la configuracion del WorldEditor de WorldEditor.ini
'

Private Sub CargarMapIni()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    On Error GoTo Fallo

    Dim tStr As String

    Dim Leer As New clsIniManager

    If FileExist(App.Path & "\WorldEditor.ini", vbArchive) = False Then
        frmMain.mnuGuardarUltimaConfig.Checked = True
        MaxGrhs = 32000
        UserPos.X = 50
        UserPos.Y = 50
        PantallaX = 19
        PantallaY = 22
        MsgBox "Falta el archivo 'WorldEditor.ini' de configuración.", vbInformation
        Exit Sub

    End If
    
    Call Leer.Initialize(App.Path & "\WorldEditor.ini")
    
    ' Obj de Translado
    Cfg_TrOBJ = Val(Leer.GetValue("CONFIGURACION", "ObjTranslado"))
    frmMain.mnuAutoCapturarTranslados.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarTrans"))
    frmMain.mnuAutoCapturarSuperficie.Checked = Val(Leer.GetValue("CONFIGURACION", "AutoCapturarSup"))
    frmMain.mnuUtilizarDeshacer.Checked = Val(Leer.GetValue("CONFIGURACION", "UtilizarDeshacer"))
    
    ' Guardar Ultima Configuracion
    frmMain.mnuGuardarUltimaConfig.Checked = Val(Leer.GetValue("CONFIGURACION", "GuardarConfig"))

    ' Index
    MaxGrhs = Val(Leer.GetValue("INDEX", "MaxGrhs"))

    If MaxGrhs < 1 Then MaxGrhs = 32000
    
    tStr = Leer.GetValue("MOSTRAR", "LastPos") ' x-y
    UserPos.X = Val(ReadField(1, tStr, Asc("-")))
    UserPos.Y = Val(ReadField(2, tStr, Asc("-")))

    If UserPos.X < XMinMapSize Or UserPos.X > XMaxMapSize Then
        UserPos.X = 50
    End If

    If UserPos.Y < YMinMapSize Or UserPos.Y > YMaxMapSize Then
        UserPos.Y = 50
    End If
    
    ' Menu Mostrar
    frmMain.mnuVerAutomatico.Checked = Val(Leer.GetValue("MOSTRAR", "ControlAutomatico"))
    frmMain.mnuVerCapa2.Checked = Val(Leer.GetValue("MOSTRAR", "Capa2"))
    frmMain.mnuVerCapa3.Checked = Val(Leer.GetValue("MOSTRAR", "Capa3"))
    frmMain.mnuVerCapa4.Checked = Val(Leer.GetValue("MOSTRAR", "Capa4"))
    frmMain.mnuVerTranslados.Checked = Val(Leer.GetValue("MOSTRAR", "Translados"))
    frmMain.mnuVerObjetos.Checked = Val(Leer.GetValue("MOSTRAR", "Objetos"))
    frmMain.mnuVerNPCs.Checked = Val(Leer.GetValue("MOSTRAR", "NPCs"))
    frmMain.mnuVerTriggers.Checked = Val(Leer.GetValue("MOSTRAR", "Triggers"))
    frmMain.mnuVerGrilla.Checked = Val(Leer.GetValue("MOSTRAR", "Grilla")) ' Grilla
    VerGrilla = frmMain.mnuVerGrilla.Checked
    frmMain.mnuVerBloqueos.Checked = Val(Leer.GetValue("MOSTRAR", "Bloqueos"))
    frmMain.cVerTriggers.Value = frmMain.mnuVerTriggers.Checked
    frmMain.cVerBloqueos.Value = frmMain.mnuVerBloqueos.Checked
    
    ' Tamaño de visualizacion
    PantallaX = Val(Leer.GetValue("MOSTRAR", "PantallaX"))
    PantallaY = Val(Leer.GetValue("MOSTRAR", "PantallaY"))

    If PantallaX > 23 Or PantallaX <= 2 Then PantallaX = 23
    If PantallaY > 32 Or PantallaY <= 2 Then PantallaY = 32
    
    ' [GS] 02/10/06
    ' Tamaño de visualizacion en el cliente
    ClienteHeight = Val(Leer.GetValue("RENDER", "ClienteHeight"))
    ClienteWidth = Val(Leer.GetValue("RENDER", "ClienteWidth"))

    If frmMain.Option2.Value = True Then
        ClienteHeight = 13
        ClienteWidth = 17
        Else
        ClienteHeight = 19
        ClienteWidth = 21
    End If
    
    If ClienteHeight <= 0 Then ClienteHeight = 13
    If ClienteWidth <= 0 Then ClienteWidth = 17
    
    Exit Sub
Fallo:
    MsgBox "ERROR " & Err.Number & " en WorldEditor.ini" & vbCrLf & Err.Description, vbCritical

    Resume Next

End Sub
    
Public Sub Main()

    '*************************************************
    'Author: Unkwown
    'Last modified: 25/11/08 - GS
    '*************************************************
    On Error Resume Next

    If App.PrevInstance = True Then End

    Dim OffsetCounterX As Integer
    Dim OffsetCounterY As Integer
    Dim Chkflag        As Integer

    Call CargarMapIni
    
    With frmCargando

        If FileExist(IniPath & "WorldEditor.jpg", vbArchive) Then
            .Picture1.Picture = LoadPicture(IniPath & "WorldEditor.jpg")
        End If
    
        .verX = "v" & App.Major & "." & App.Minor & "." & App.Revision
        
        Call .Show
        Call .SetFocus
        DoEvents
                
        .X.Caption = "Inicializando constantes..."
            IniPath = App.Path & "\"
            InitPath = App.Path & "\Recursos\Init\"
            DirDats = App.Path & "\Dats\"
            DirMinimapas = App.Path & "\Recursos\Graficos\MiniMapa\"
            DirAudio = App.Path & "\Recursos\Audio\"
        DoEvents
        
        .X.Caption = "Cargando Indice de Superficies..."
            Call modIndices.CargarIndicesSuperficie
        DoEvents

        .X.Caption = "Indexando Cargado de Imagenes..."
        DoEvents

        If Not Engine_Init Then ' 30/05/2006
            Call MsgBox("¡No se ha logrado iniciar el engine gráfico! Reinstale los últimos controladores de DirectX y actualize sus controladores de video.", vbCritical, "Saliendo")
            End
        End If

        .X.Caption = "Iniciando motor de audio"
        DoEvents
        
        Call Audio.Initialize(dX, frmMain.hwnd, "", DirAudio & "MIDI\", DirAudio & "MP3\")
        
    
        With MapSize
            .XMax = XMaxMapSize
            .XMin = XMinMapSize
            .YMax = YMaxMapSize
            .YMin = YMinMapSize
            ReDim MapData(.XMin To .XMax, .YMin To .YMax)
        End With

        Call .SetFocus

        .X.Caption = "Iniciando Ventana de Edición..."
        DoEvents
        
        Call .Hide
    
    End With
    
    'If frmRender.Visible Then RenderToPicture
    frmMain.Show
    DoEvents
    
    Call modMapIO.NuevoMapa
    
    prgRun = True
    Call Start
    
    Call AddtoRichTextBox(frmMain.StatTxt, "World Editor By Lorwik iniciado...", 0, 255, 0)

End Sub

Public Function GetVar(File As String, Main As String, Var As String) As String

    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    Dim L        As Integer
    Dim Char     As String
    Dim sSpaces  As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found

    szReturn = vbNullString
    sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
    Call GetPrivateProfileString(Main, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

Public Sub WriteVar(File As String, Main As String, Var As String, Value As String)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    
    Call writeprivateprofilestring(Main, Var, Value, File)

End Sub

Public Sub ToggleWalkMode()

    '*************************************************
    'Author: Unkwown
    'Last modified: 28/05/06 - GS
    '*************************************************
    On Error GoTo fin:

    If WalkMode = False Then
        WalkMode = True
    Else
        frmMain.mnuModoCaminata.Checked = False
        WalkMode = False

    End If

    If WalkMode = False Then
        'Erase character
        Call EraseChar(UserCharIndex)
        MapData(UserPos.X, UserPos.Y).CharIndex = 0
    Else

        'MakeCharacter
        If LegalPos(UserPos.X, UserPos.Y) Then
            Call MakeChar(NextOpenChar(), 1, 1, SOUTH, UserPos.X, UserPos.Y)
            UserCharIndex = MapData(UserPos.X, UserPos.Y).CharIndex
            frmMain.mnuModoCaminata.Checked = True
        Else
            MsgBox "ERROR: Ubicacion ilegal."
            WalkMode = False

        End If

    End If

fin:

End Sub

Public Sub FixCoasts(ByVal GrhIndex As Long, ByVal X As Integer, ByVal Y As Integer)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************

    If GrhIndex = 7284 Or GrhIndex = 7290 Or GrhIndex = 7291 Or GrhIndex = 7297 Or GrhIndex = 7300 Or GrhIndex = 7301 Or GrhIndex = 7302 Or GrhIndex = 7303 Or GrhIndex = 7304 Or GrhIndex = 7306 Or GrhIndex = 7308 Or GrhIndex = 7310 Or GrhIndex = 7311 Or GrhIndex = 7313 Or GrhIndex = 7314 Or GrhIndex = 7315 Or GrhIndex = 7316 Or GrhIndex = 7317 Or GrhIndex = 7319 Or GrhIndex = 7321 Or GrhIndex = 7325 Or GrhIndex = 7326 Or GrhIndex = 7327 Or GrhIndex = 7328 Or GrhIndex = 7332 Or GrhIndex = 7338 Or GrhIndex = 7339 Or GrhIndex = 7345 Or GrhIndex = 7348 Or GrhIndex = 7349 Or GrhIndex = 7350 Or GrhIndex = 7351 Or GrhIndex = 7352 Or GrhIndex = 7349 Or GrhIndex = 7350 Or GrhIndex = 7351 Or GrhIndex = 7354 Or GrhIndex = 7357 Or GrhIndex = 7358 Or GrhIndex = 7360 Or GrhIndex = 7362 Or GrhIndex = 7363 Or GrhIndex = 7365 Or GrhIndex = 7366 Or GrhIndex = 7367 Or GrhIndex = 7368 Or GrhIndex = 7369 Or GrhIndex = 7371 Or GrhIndex = 7373 Or GrhIndex = 7375 Or GrhIndex = 7376 Then MapData(X, Y).Graphic(2).GrhIndex = 0

End Sub

Public Function RandomNumber(ByVal LowerBound As Variant, _
                             ByVal UpperBound As Variant) As Single
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    Call Randomize(Timer)
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound

End Function

''
' Actualiza el Caption del menu principal
'
' @param Trabajando Indica el path del mapa con el que se esta trabajando
' @param Editado Indica si el mapa esta editado

Public Sub CaptionWorldEditor(ByVal Trabajando As String, ByVal Editado As Boolean)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If Trabajando = vbNullString Then
        Trabajando = "Nuevo Mapa"

    End If

    frmMain.Caption = "WorldEditor v" & App.Major & "." & App.Minor & " Build " & App.Revision & " - [" & Trabajando & "]"

    If Editado = True Then
        frmMain.Caption = frmMain.Caption & " (modificado)"

    End If

End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                     ByVal Text As String, _
                     Optional ByVal red As Integer = -1, _
                     Optional ByVal green As Integer, _
                     Optional ByVal blue As Integer, _
                     Optional ByVal bold As Boolean = False, _
                     Optional ByVal italic As Boolean = False, _
                     Optional ByVal bCrLf As Boolean = True)

    '******************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D
    'apperance!
    'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
    'Juan Martín Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
    '******************************************r
    With RichTextBox

        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF

        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        
        RichTextBox.Refresh

    End With

End Sub

Public Function ColorToDX8(ByVal long_color As Long) As Long

    ' DX8 engine
    Dim temp_color As String
    Dim red        As Integer, blue As Integer, green As Integer
    
    temp_color = Hex$(long_color)

    If Len(temp_color) < 6 Then
        'Give is 6 digits for easy RGB conversion.
        temp_color = String$(6 - Len(temp_color), "0") + temp_color
    End If
    
    red = CLng("&H" + mid$(temp_color, 1, 2))
    green = CLng("&H" + mid$(temp_color, 3, 2))
    blue = CLng("&H" + mid$(temp_color, 5, 2))
    
    ColorToDX8 = D3DColorXRGB(red, green, blue)

End Function

'Solo se usa para el minimapa
Public Function ReturnNumberFromString(ByVal sString As String) As String
   
   Dim i As Integer
   
   For i = 1 To LenB(sString)
   
       If mid$(sString, i, 1) Like "[0-9]" Then
           ReturnNumberFromString = ReturnNumberFromString + mid$(sString, i, 1)
       End If
       
   Next i
   
End Function

