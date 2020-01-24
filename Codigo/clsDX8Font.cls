VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX8Font"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private DirectD3D As D3DX8
 
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (destination As Any, _
                                       source As Any, _
                                       ByVal length As Long)
 
'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
 
    Width As Long
    Height As Long
 
End Type
 
Private Type POINTAPI
 
    X As Long
    Y As Long
 
End Type
 
Private Type CharVA
 
    Vertex(0 To 3) As TLVERTEX
 
End Type
 
Private Type VFH
 
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
 
End Type
 
Private Type CustomFont
 
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
 
End Type
 
'public Const Font_Default_TextureNum As Long = -1   'The texture number used to represent this font - only used for AlternateRendering - keep negative to prevent interfering with game textures
Private cfonts(1 To 3) As CustomFont  ' _Default2 As CustomFont
 
Sub Engine_Text_Draw(ByVal Left As Long, _
                     ByVal Top As Long, _
                     ByVal Text As String, _
                     ByVal color As Long, _
                     Optional ByVal Alpha As Byte = 255, _
                     Optional ByVal Center As Boolean = False, _
                     Optional ByVal Font As Byte = 1)
   
    Engine_Render_Text cfonts(Font), Text, Left, Top, color, Center, Alpha, Font
 
End Sub

Public Sub Text_Render_Special(ByVal intX As Integer, _
                               ByVal intY As Integer, _
                               ByRef strText As String, _
                               ByVal lngColor As Long, _
                               Optional bolCentred As Boolean = False)  ' GSZAO
    '*****************************************************************
    'Text_Render_Special by ^[GS]^
    '*****************************************************************
    
    If LenB(strText) <> 0 Then
        lngColor = ColorToDX8(lngColor)
        Call Engine_Render_Text(cfonts(1), strText, intX, intY, lngColor, bolCentred)

    End If
    
End Sub
 
Private Sub Engine_Render_Text(ByRef UseFont As CustomFont, _
                               ByVal Text As String, _
                               ByVal X As Long, _
                               ByVal Y As Long, _
                               ByVal color As Long, _
                               Optional ByVal Center As Boolean = False, _
                               Optional ByVal Alpha As Byte = 255, _
                               Optional ByVal Font As Byte = 1)
 
    Dim TempVA(0 To 3) As TLVERTEX

    Dim tempstr()      As String

    Dim count          As Integer

    Dim ascii()        As Byte

    Dim i              As Long

    Dim j              As Long

    Dim TempColor      As Long

    Dim ResetColor     As Byte

    Dim YOffset        As Single
 
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
   
    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
 
    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
   
    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)
   
    'Set the temp color (or else the first character has no color)
    TempColor = color
 
    'Set the texture
    D3DDevice.SetTexture 0, UseFont.Texture
   
    If Center Then
        X = X - Engine_GetTextWidth(cfonts(Font), Text) * 0.5

    End If
   
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)
 
        If Len(tempstr(i)) > 0 Then
            YOffset = i * UseFont.CharHeight
            count = 0
       
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
       
            'Loop through the characters
            For j = 1 To Len(tempstr(i))
   
                'Copy from the cached vertex array to the temp vertex array
                CopyMemory TempVA(0), UseFont.HeaderInfo.CharVA(ascii(j - 1)).Vertex(0), 32 * 4
 
                'Set up the verticies
                TempVA(0).X = X + count
                TempVA(0).Y = Y + YOffset
                       
                TempVA(1).X = TempVA(1).X + X + count
                TempVA(1).Y = TempVA(0).Y
 
                TempVA(2).X = TempVA(0).X
                TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
 
                TempVA(3).X = TempVA(1).X
                TempVA(3).Y = TempVA(2).Y
                       
                'Set the colors
                TempVA(0).color = TempColor
                TempVA(1).color = TempColor
                TempVA(2).color = TempColor
                TempVA(3).color = TempColor
                       
                'Draw the verticies
                D3DDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(Alpha, 0, 0, 0)
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0))
                                
                'Shift over the the position to render the next character
                count = count + UseFont.HeaderInfo.CharWidth(ascii(j - 1))
 
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = color

                End If
               
            Next j
           
        End If
 
    Next i
   
End Sub
 
Private Function Engine_GetTextWidth(ByRef UseFont As CustomFont, _
                                     ByVal Text As String) As Integer
 
    Dim i As Long
 
    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
   
    'Loop through the text
    For i = 1 To Len(Text)
       
        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
       
    Next i
 
End Function

Sub Engine_Init_FontSettings()

    Dim FileNum  As Byte

    Dim LoopChar As Long

    Dim Row      As Single

    Dim u        As Single

    Dim v        As Single

    Dim i        As Byte

    Dim file     As String

    '*** Default font ***
        
    For i = 1 To UBound(cfonts)
        
        'Load the header information
        FileNum = FreeFile
        Open App.Path & "\Recursos\Fonts\Font" & i & ".dat" For Binary As #FileNum
        Get #FileNum, , cfonts(i).HeaderInfo
        Close #FileNum

        'Calculate some common values
        cfonts(i).CharHeight = cfonts(i).HeaderInfo.CellHeight - 4
        cfonts(i).RowPitch = cfonts(i).HeaderInfo.BitmapWidth \ cfonts(i).HeaderInfo.CellWidth
        cfonts(i).ColFactor = cfonts(i).HeaderInfo.CellWidth / cfonts(i).HeaderInfo.BitmapWidth
        cfonts(i).RowFactor = cfonts(i).HeaderInfo.CellHeight / cfonts(i).HeaderInfo.BitmapHeight
          
        'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
        For LoopChar = 0 To 255
              
            'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
            Row = (LoopChar - cfonts(i).HeaderInfo.BaseCharOffset) \ cfonts(i).RowPitch
            u = ((LoopChar - cfonts(i).HeaderInfo.BaseCharOffset) - (Row * cfonts(i).RowPitch)) * cfonts(i).ColFactor
            v = Row * cfonts(i).RowFactor
        
            'Set the verticies
            With cfonts(i).HeaderInfo.CharVA(LoopChar)
                .Vertex(0).color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
                .Vertex(0).rhw = 1
                .Vertex(0).tu = u
                .Vertex(0).tv = v
                .Vertex(0).X = 0
                .Vertex(0).Y = 0
                .Vertex(0).Z = 0
                  
                .Vertex(1).color = D3DColorARGB(255, 0, 0, 0)
                .Vertex(1).rhw = 1
                .Vertex(1).tu = u + cfonts(i).ColFactor
                .Vertex(1).tv = v
                .Vertex(1).X = cfonts(i).HeaderInfo.CellWidth
                .Vertex(1).Y = 0
                .Vertex(1).Z = 0
                  
                .Vertex(2).color = D3DColorARGB(255, 0, 0, 0)
                .Vertex(2).rhw = 1
                .Vertex(2).tu = u
                .Vertex(2).tv = v + cfonts(i).RowFactor
                .Vertex(2).X = 0
                .Vertex(2).Y = cfonts(i).HeaderInfo.CellHeight
                .Vertex(2).Z = 0
                  
                .Vertex(3).color = D3DColorARGB(255, 0, 0, 0)
                .Vertex(3).rhw = 1
                .Vertex(3).tu = u + cfonts(i).ColFactor
                .Vertex(3).tv = v + cfonts(i).RowFactor
                .Vertex(3).X = cfonts(i).HeaderInfo.CellWidth
                .Vertex(3).Y = cfonts(i).HeaderInfo.CellHeight
                .Vertex(3).Z = 0
                
            End With
            
        Next LoopChar
        
    Next i
 
End Sub

Sub Engine_Init_FontTextures()

    '*****************************************************************
    'Init the custom font textures
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures
    '*****************************************************************
    On Error GoTo eDebug:
    
    Dim i       As Long

    Dim TexInfo As D3DXIMAGE_INFO_A

    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '*** Default font ***
    For i = 1 To UBound(cfonts)
        
        'Set the texture
        Set cfonts(i).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Recursos\Fonts\font" & i & ".bmp", D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)

        'Store the size of the texture
        cfonts(i).TextureSize.X = TexInfo.Width
        cfonts(i).TextureSize.Y = TexInfo.Height
    Next
    
    Exit Sub

eDebug:

    If Err.Number = "-2005529767" Then
        MsgBox "Error en la textura de fuente utilizada " & App.Path & "\Recursos\Fonts\Font.png", vbCritical
        End

    End If
    
    End

End Sub
