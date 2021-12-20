VERSION 5.00
Begin VB.Form frmInformes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   Icon            =   "frmInformes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInfo 
      Height          =   3855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin WorldEditor.lvButtons_H cmdObjetos 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
      _extentx        =   3625
      _extenty        =   873
      caption         =   "&Objetos"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmInformes.frx":1E74E
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdCerrar 
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   4920
      Width           =   2175
      _extentx        =   3836
      _extenty        =   873
      caption         =   "&Cerrar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmInformes.frx":1E77A
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdTranslados 
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   4200
      Width           =   2175
      _extentx        =   3836
      _extenty        =   873
      caption         =   "&Translados"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmInformes.frx":1E7A6
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdNPCs 
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
      _extentx        =   3413
      _extenty        =   873
      caption         =   "&NPCs/Hostiles"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmInformes.frx":1E7D2
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6480
      Y1              =   4815
      Y2              =   4815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6480
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6480
      Y1              =   4070
      Y2              =   4070
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6480
      Y1              =   4080
      Y2              =   4080
   End
End
Attribute VB_Name = "frmInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************
Option Explicit

Private Sub cmdCerrar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Unload Me

End Sub

''
'   Genera el informe de Objetos
'

Private Sub ActalizarObjetos()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    Dim Y As Integer

    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    txtInfo.Text = "Informe de Objetos (X,Y)"

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, Y).OBJInfo.objindex > 0 Then
                txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " tiene " & MapData(X, Y).OBJInfo.Amount & " del Objeto " & MapData(X, Y).OBJInfo.objindex & " - " & ObjData(MapData(X, Y).OBJInfo.objindex).name

            End If

        Next X
    Next Y

End Sub

''
'   Genera el informe de NPCs
'

Private Sub ActalizarNPCs()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    Dim Y As Integer

    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    txtInfo.Text = "Informe de NPCs/Hostiles (X,Y)"

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, Y).NPCIndex > 0 Then
                If MapData(X, Y).NPCIndex >= 500 Then
                    txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " tiene " & NpcData(MapData(X, Y).NPCIndex).name & " (Hostil)"
                Else
                    txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " tiene " & NpcData(MapData(X, Y).NPCIndex).name

                End If

            End If

        Next X
    Next Y

End Sub

''
'   Genera el informe de Translados
'

Private Sub ActalizarTranslados()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    Dim Y As Integer

    Dim X As Integer

    If Not MapaCargado Then
        Exit Sub

    End If

    txtInfo.Text = "Informe de Translados (X,Y)"

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If MapData(X, Y).TileExit.Map > 0 Then
                txtInfo.Text = txtInfo.Text & vbCrLf & X & "," & Y & " nos traslada a la posición " & MapData(X, Y).TileExit.X & "," & MapData(X, Y).TileExit.Y & " del Mapa " & MapData(X, Y).TileExit.Map

                If ((X < 20 And MapData(X, Y).TileExit.X < 20) Or (X > 80 And MapData(X, Y).TileExit.X > 80)) And (X <> MapData(X, Y).TileExit.X) Then
                    txtInfo.Text = txtInfo.Text & " (X sospechoso)"

                End If

                If ((Y < 20 And MapData(X, Y).TileExit.Y < 20) Or (Y > 80 And MapData(X, Y).TileExit.Y > 80)) And (Y <> MapData(X, Y).TileExit.Y) Then
                    txtInfo.Text = txtInfo.Text & " (Y sospechoso)"

                End If

            End If

        Next X
    Next Y

End Sub

Private Sub cmdNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call ActalizarNPCs

End Sub

Private Sub cmdObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call ActalizarObjetos

End Sub

Private Sub cmdTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call ActalizarTranslados

End Sub
