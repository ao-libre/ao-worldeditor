VERSION 5.00
Begin VB.Form frmAmpliacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ampliacion de mapa"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Pegar mapa 100x100"
      Height          =   360
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   2535
   End
   Begin WorldEditor.lvButtons_H LvBTraslados 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   3120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "Traslados"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H LvBCopiarMapa 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Caption         =   "Copiar mapa 100x100"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H LvBBorrarBloqueos 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Caption         =   "Borrar bloqueos inherentes al mapa de 100x100"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Text            =   "0"
      Top             =   2640
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pegar"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "frmAmpliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Borrado As Boolean
    Borrado = False
    Call PegarMapa(0 + Val(Text10.Text), 0 + (Val(Text9.Text)))
End Sub

Private Sub LvBTraslados_Click()
Dim X As Long, Y As Long

    For X = 1 To XMaxMapSize
        For Y = 1 To 8
        
           If MapData(X, Y).TileExit.Map > 0 Then
              MapData(X, 9).TileExit = MapData(X, Y).TileExit
              MapData(X, Y).TileExit.Map = 0
              MapData(X, Y).TileExit.X = 0
              MapData(X, Y).TileExit.Y = 0
           End If
        
        Next Y
    Next X
    
    For X = 1 To XMaxMapSize
        For Y = 93 To YMaxMapSize
        
           If MapData(X, Y).TileExit.Map > 0 Then
              MapData(X, 92).TileExit = MapData(X, Y).TileExit
              MapData(X, Y).TileExit.Map = 0
              MapData(X, Y).TileExit.X = 0
              MapData(X, Y).TileExit.Y = 0
           End If
        
        Next Y
    Next X
    
    For X = 1 To 11
        For Y = 1 To YMaxMapSize
        
           If MapData(X, Y).TileExit.Map > 0 Then
              MapData(12, Y).TileExit = MapData(X, Y).TileExit
              MapData(X, Y).TileExit.Map = 0
              MapData(X, Y).TileExit.X = 0
              MapData(X, Y).TileExit.Y = 0
           End If
        
        Next Y
    Next X
    
    For X = 90 To XMaxMapSize
        For Y = 1 To YMaxMapSize
        
           If MapData(X, Y).TileExit.Map > 0 Then
              MapData(89, Y).TileExit = MapData(X, Y).TileExit
              MapData(X, Y).TileExit.Map = 0
              MapData(X, Y).TileExit.X = 0
              MapData(X, Y).TileExit.Y = 0
           End If
        
        Next Y
    Next X

    'modMapIO.GuardarMapa Dialog.FileName

End Sub
Private Sub LvBCopiarMapa_Click()
Dim X As Integer
Dim Y As Integer

    For X = 1 To XMaxMapSize
        For Y = 1 To YMaxMapSize
            With MapData2(X, Y)
                .Graphic(1) = MapData(X, Y).Graphic(1)
                .Graphic(2) = MapData(X, Y).Graphic(2)
                .Graphic(3) = MapData(X, Y).Graphic(3)
                .Graphic(4) = MapData(X, Y).Graphic(4)
                .blocked = MapData(X, Y).blocked
                .NPCIndex = MapData(X, Y).NPCIndex
                .Trigger = MapData(X, Y).Trigger
                .ObjGrh = MapData(X, Y).ObjGrh
                .OBJInfo = MapData(X, Y).OBJInfo
            End With
        Next
    Next
End Sub

Private Sub LvBBorrarBloqueos_Click()
Dim X As Integer
Dim Y As Integer
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
        
        If MapData(X, Y).Graphic(2).GrhIndex > 0 Or _
           MapData(X, Y).Graphic(3).GrhIndex > 0 Or _
           MapData(X, Y).Graphic(4).GrhIndex > 0 Or _
           MapData(X, Y).OBJInfo.objindex > 0 Then GoTo Jump
        
        If X >= 10 And Y >= 93 And Y <= 108 Then MapData(X, Y).blocked = 0
        If X >= 92 And X <= 108 And Y >= 8 Then MapData(X, Y).blocked = 0
        If X >= 192 And X <= 208 And Y >= 8 Then MapData(X, Y).blocked = 0
        If X >= 9 And X <= 91 And Y >= 182 And Y <= 193 Then MapData(X, Y).blocked = 0
        If X >= 109 And X <= 191 And Y >= 188 And Y <= 193 Then MapData(X, Y).blocked = 0
        If X >= 209 And X <= 274 And Y >= 195 And Y <= 206 Then MapData(X, Y).blocked = 0
        If X >= 109 And X <= 192 Then MapData(X, Y).blocked = 0
        If Y >= 182 And Y <= 188 Then MapData(X, Y).blocked = 0
        
Jump:
        Next
    Next
End Sub

Private Sub Label2_Click(index As Integer)
    Dim Borrado As Boolean
    Borrado = False
    If Not Borrado Then
        Select Case index
            Case 0
                Call PegarMapa(0 + Val(Text10.Text), 0 + (Val(Text9.Text)))
            Case 1
                Call PegarMapa(100 + Val(Text10.Text), 0 + Val(Text9.Text))
            Case 2
                Call PegarMapa(200 + Val(Text10.Text), 0 + Val(Text9.Text))
            Case 3
                Call PegarMapa(0 + Val(Text10.Text), 100 + Val(Text9.Text))
            Case 4
                Call PegarMapa(100 + Val(Text10.Text), 100 + Val(Text9.Text))
            Case 5
                Call PegarMapa(200 + Val(Text10.Text), 100 + Val(Text9.Text))
            Case 6
             Call PegarMapa(0 + Val(Text10.Text), 200 + Val(Text9.Text))
            Case 7
                Call PegarMapa(100 + Val(Text10.Text), 200 + Val(Text9.Text))
            Case 8
                Call PegarMapa(200 + Val(Text10.Text), 200 + Val(Text9.Text))
        End Select
    Else
        Select Case index
        
            Case 0
                Call BorrarMapa(0, 0)
            Case 1
                Call BorrarMapa(100, 0)
            Case 2
                Call BorrarMapa(200, 0)
            Case 3
                Call BorrarMapa(0, 100)
            Case 4
                Call BorrarMapa(100, 100)
            Case 5
                Call BorrarMapa(200, 100)
            Case 6
                Call BorrarMapa(0, 200)
            Case 7
                Call BorrarMapa(100, 200)
            Case 8
                Call BorrarMapa(200, 200)
        End Select
    End If

End Sub

Private Sub PegarMapa(ByVal mX As Integer, ByVal mY As Integer)
On Error GoTo err
Dim OffsetX As Integer
Dim OffsetY As Integer
Dim X As Integer, Y As Integer


    OffsetX = X + mX
    OffsetY = Y + mY

    For X = 1 To XMaxMapSize
        For Y = 1 To YMaxMapSize
        
            If OffsetX + X > 0 And OffsetX + X < 201 Then
              If OffsetY + Y > 0 And OffsetY + Y < 201 Then
              
                With MapData(X + OffsetX, Y + OffsetY)
    
                    .Graphic(1) = MapData2(X, Y).Graphic(1)
                    .Graphic(2) = MapData2(X, Y).Graphic(2)
                    .Graphic(3) = MapData2(X, Y).Graphic(3)
                    .Graphic(4) = MapData2(X, Y).Graphic(4)
                    .blocked = MapData2(X, Y).blocked
                    .NPCIndex = MapData2(X, Y).NPCIndex
                    .Trigger = MapData2(X, Y).Trigger
                    .ObjGrh = MapData2(X, Y).ObjGrh
                    .OBJInfo = MapData2(X, Y).OBJInfo
                End With
              End If
            End If
          
        Next
    Next
err:
Debug.Print err.Description
Debug.Print "error en pegarmapa"
End Sub

Private Sub BorrarMapa(ByVal mX As Integer, ByVal mY As Integer)
Dim GrhNull As Grh
Dim ObjectNull As Obj
Dim X As Integer, Y As Integer

For X = 1 To XMaxMapSize
    For Y = 1 To YMaxMapSize
        With MapData(X + mX, Y + mY)
            .Graphic(1) = GrhNull
            .Graphic(2) = GrhNull
            .Graphic(3) = GrhNull
            .Graphic(4) = GrhNull
            .blocked = 0
            .NPCIndex = 0
            .Trigger = 0
            .ObjGrh = GrhNull
            .OBJInfo = ObjectNull
        End With
    Next
Next
End Sub

Private Sub Picture1_Click()

End Sub
