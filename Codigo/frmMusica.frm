VERSION 5.00
Begin VB.Form frmMusica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Musica"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmMusica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin WorldEditor.lvButtons_H cmdCerrar 
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
      _extentx        =   4260
      _extenty        =   873
      caption         =   "&Cerrar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMusica.frx":1E74E
      mode            =   0
      value           =   0
      cback           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdAplicarYCerrar 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   2415
      _extentx        =   4260
      _extenty        =   873
      caption         =   "&Aplicar y Cerrar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMusica.frx":1E77A
      mode            =   0
      value           =   0
      enabled         =   0
      cback           =   12648447
   End
   Begin WorldEditor.lvButtons_H cmdDetener 
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      _extentx        =   2143
      _extenty        =   873
      caption         =   "&Detener"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMusica.frx":1E7A6
      mode            =   0
      value           =   0
      enabled         =   0
      cback           =   12632319
   End
   Begin WorldEditor.lvButtons_H cmdEscuchar 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _extentx        =   2143
      _extenty        =   873
      caption         =   "&Escuchar"
      capalign        =   2
      backstyle       =   2
      cgradient       =   0
      font            =   "frmMusica.frx":1E7D2
      mode            =   0
      value           =   0
      enabled         =   0
      cback           =   12648384
   End
   Begin VB.FileListBox fleMusicas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   120
      Pattern         =   "*.mid"
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmMusica"
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

Public Modo As MusicTypes
Private Actual As String
Private Play As Boolean

''
' Aplica la Musica seleccionada y oculta la ventana
'

Private Sub cmdAplicarYCerrar_Click()

    If Actual <> vbNullString Then
        If Modo = midi Then
            MapInfo.midi = Left$(Actual, Len(Actual) - 4)
            frmMapInfo.txtMapMusica.Text = MapInfo.midi
            frmMain.lblMapMusica = MapInfo.midi
        Else
            MapInfo.mp3 = Left$(Actual, Len(Actual) - 4)
            frmMapInfo.txtMapMP3.Text = MapInfo.mp3
        End If

        Actual = vbNullString

    End If

    Unload Me

End Sub

''
' Oculta la ventana
'

Private Sub cmdCerrar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Unload Me

End Sub

''
' Detiene la Musica que se encuentra Reproduciendo
'

Private Sub cmdDetener_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    If Modo = midi Then
        Call Audio.StopMidi
    Else
        Call Audio.StopMp3
    End If

    cmdEscuchar.Enabled = True
    cmdDetener.Enabled = False
    Play = False

End Sub

''
' Inicia la reproduccion de la Musica Seleccionada
'

Private Sub cmdEscuchar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If Modo = midi Then
        Call Audio.PlayMIDI(fleMusicas.List(fleMusicas.ListIndex))
    Else
        Call Audio.PlayMp3(fleMusicas.List(fleMusicas.ListIndex))
    End If

    cmdDetener.Enabled = True
    cmdEscuchar.Enabled = False
    Play = True

End Sub

''
' Selecciona una nueva Musica del listado
'

Private Sub fleMusicas_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Actual = fleMusicas.List(fleMusicas.ListIndex)
    cmdAplicarYCerrar.Enabled = True
    If Play = False Then cmdEscuchar.Enabled = True
End Sub

Private Sub Form_Load()
    If Modo = midi Then
        fleMusicas.Path = DirAudio & "MIDI"
        fleMusicas.Pattern = "*.mid"
    Else
        fleMusicas.Path = DirAudio & "MP3"
        fleMusicas.Pattern = "*.mp3"
    End If
    
    Actual = vbNullString
    Play = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modo = midi Then
        Call Audio.StopMidi
    Else
        Call Audio.StopMp3
    End If
End Sub
