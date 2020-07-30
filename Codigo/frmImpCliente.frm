VERSION 5.00
Begin VB.Form frmImpCliente 
   Caption         =   "Importar archivos nesesarios al Editor"
   ClientHeight    =   6390
   ClientLeft      =   10725
   ClientTop       =   6030
   ClientWidth     =   8160
   Icon            =   "frmImpCliente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8160
   Begin VB.CommandButton Command2 
      Caption         =   "Importar del Server"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Importar del Cliente"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   3720
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.DirListBox Dir2 
      Appearance      =   0  'Flat
      Height          =   2790
      Left            =   4320
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "SERVIDOR"
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "CLIENTE"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label LblCliente 
      Caption         =   "Seleccionar la carpeta Cliente y Servidor luego Importar"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "Barra de estado:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   7815
   End
   Begin VB.Menu mnuMover 
      Caption         =   "Importar del Cliente"
   End
   Begin VB.Menu mnuServer 
      Caption         =   "Importar del Server"
   End
End
Attribute VB_Name = "frmImpCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call verClienteyServer
Call mnuMover_Click
End Sub

Private Sub Command2_Click()
Call verClienteyServer
Call mnuServer_Click
End Sub

Private Sub Form_Load()
On Error Resume Next
Move (Screen.Width - Width) \ 29, (Screen.Height - Height) \ 29 'Centra el formulario completamente
End Sub



Private Function xfilecopy(origen$, destino$, Archivo$, informa As Label)
' Copia varios archivos de una carpeta a otra
' Origen$= directorio de origen , terminado en "\"
' Destino$= directorio de destino , terminado en "\"
' archivo$= especificacion de archivos a copiar, con simb. comodin
' informa= un label en el que se muestra el progreso
'
' result = xfilecopy("c:\pat\", "h:\doc\", "*.exe", label1)
' copia todos los archivos exe de c:\pat en h:\doc
' muestra lo que esta haciendo en label1

Dim n, result, cuenta, pcent
' cuenta los archivos a copiar
cuenta = 0
n = Dir$(origen$ & Archivo$)
While (n <> "")
 cuenta = cuenta + 1
 n = Dir$
Wend

' Copia
result = 0
n = Dir$(origen$ & Archivo$)
On Error GoTo malxfilecopy
While (n <> "") And (result > -1)
 pcent = (result + 1) & "/" & cuenta & " "
 pcent = pcent & Format$(100 * result / cuenta, "#0.0") & "%"
 informa.Caption = pcent & " Copiando " & origen$ & n & " a " & destino$
 DoEvents

 FileCopy origen$ & n, destino$ & n
 result = result + 1
 n = Dir$
continuaxfilecopy:
Wend
informa.Caption = ""
xfilecopy = result
Exit Function

malxfilecopy:
 result = -1
 Resume continuaxfilecopy
End Function


Private Sub mnuMover_Click()
Dim Ruta, Ruta1, X, z
Dim wgraficos, wegraficos
Dim wminimapa, weminimapa
If MsgBox("Desea copiar los archivos del directorio:" + Chr$(10) + Dir1.Path + Chr$(10) + "A:" + Chr$(10) + App.Path, 4 + 64 + 256, "Copiar archivos a otro directorio") = 6 Then
On Error Resume Next
If Right(Dir1.Path, 1) = "\" Then
  Ruta = Dir1.Path & ""
 Else
  Ruta = Dir1.Path & "\"
End If

'**************Rutas origen******************
Y = Ruta & "INIT\"
wgraficos = Ruta & "Graficos\"
wminimapa = Ruta & "Graficos\MiniMapa\"
wMapa = Ruta & "Mapas\Alkon\"
'*************Rutas destinos******************
z = App.Path & "\INIT\"
wegraficos = App.Path & "\Recursos\graficos\"
weminimapa = App.Path & "\Recursos\MiniMapa\"
weMapa = App.Path & "\Conversor\Mapas Long\"

'*************copiado*************************
result = xfilecopy("" & Y & "", "" & z & "", "Graficos.ind", Label1)
result = xfilecopy("" & Y & "", "" & z & "", "Cabezas.ind", Label1)
result = xfilecopy("" & Y & "", "" & z & "", "Cuerpos.ind", Label1)
result = xfilecopy("" & Y & "", "" & z & "", "Particulas.ini", Label1)
result = xfilecopy("" & Y & "", "" & z & "", "Personajes.ind", Label1)
result = xfilecopy("" & wgraficos & "", "" & wegraficos & "", "*.png", Label1)
result = xfilecopy("" & wminimapa & "", "" & weminimapa & "", "*.bmp", Label1)
result = xfilecopy("" & wMapa & "", "" & weMapa & "", "*.map", Label1)

If err Then MsgBox "No existe el directorio de fuente ni del directorio destino", 16, "¡No copie nada!"


End If
End Sub

Private Sub mnuServer_Click()
Dim Ruta, Ruta1, X, z
Dim wgraficos, wegraficos
Dim wminimapa, weminimapa
Dim wMapa, weMapa
If MsgBox("Desea copiar los archivos del directorio:" + Chr$(10) + Dir1.Path + Chr$(10) + "A:" + Chr$(10) + App.Path, 4 + 64 + 256, "Copiar archivos a otro directorio") = 6 Then
On Error Resume Next

If Right(Dir2.Path, 1) = "\" Then
  Ruta1 = Dir2.Path & ""
 Else
  Ruta1 = Dir2.Path & "\"
End If
'**************Rutas origen******************
Y = Ruta & "Dat\"
wMapa = Ruta1 & "Mundos\Alkon\"

'*************Rutas destinos******************
z = App.Path & "\Recursos\Dat\"
weMapa = App.Path & "\Conversor\Mapas Long\"

'*************copiado*************************
result = xfilecopy("" & Y & "", "" & z & "", "NPCs.dat", Label1)
result = xfilecopy("" & Y & "", "" & z & "", "obj.dat", Label1)
result = xfilecopy("" & wMapa & "", "" & weMapa & "", "*.inf", Label1)
result = xfilecopy("" & wMapa & "", "" & weMapa & "", "*.dat", Label1)

If err Then MsgBox "No existe el directorio de fuente ni del directorio destino", 16, "¡No copie nada!"


End If
End Sub


Sub verClienteyServer()


    If FileExist(IniPath & "Conversor\", vbDirectory) = False Then
        MsgBox "Falta La Carpeta 'Conversor\' Se creara y se guardara los mapas para el Cliente.", vbCritical
        MkDir (IniPath & "Conversor\")
    End If
    If FileExist(IniPath & "Conversor\Mapas Server\", vbDirectory) = False Then
        MsgBox "Falta La Carpeta 'Conversor\Mapas Server\' Se creara y se guardara los mapas para el Server sin Particulas.", vbCritical
        MkDir (IniPath & "Conversor\Mapas Server\")
    End If
    If FileExist(IniPath & "Conversor\Mapas Cliente\", vbDirectory) = False Then
        MsgBox "Falta La Carpeta 'Conversor\Mapas Cliente\' Se creara y se guardara los mapas para el Cliente.", vbCritical
        MkDir (IniPath & "Conversor\Mapas Cliente\")
    End If

    If FileExist(IniPath & "Recursos\", vbDirectory) = False Then
        MsgBox "Falta La Carpeta 'Recursos\Dat\' Se crera y se guardara los mapas para el Server sin Particulas.", vbCritical
        MkDir (IniPath & "Recursos\")
    End If
    
    If FileExist(IniPath & "Recursos\AUDIO\", vbDirectory) = False Then
        MsgBox "Falta La Carpeta 'Recursos\AUDIO\' Se crera y se guardara los mapas para el Server sin Particulas.", vbCritical
        MkDir (IniPath & "Recursos\AUDIO\")
    End If
    
    If FileExist(IniPath & "Recursos\cursores\", vbDirectory) = False Then
        MsgBox "Falta La Carpeta 'Recursos\cursores\' Se crera y se guardara los mapas para el Server sin Particulas.", vbCritical
        MkDir (IniPath & "Recursos\cursores\")
    End If
    
    If FileExist(IniPath & "Recursos\Dat\", vbDirectory) = False Then
        MsgBox "Falta La Carpeta 'Recursos\Dat\' Se crera y se guardara los mapas para el Server sin Particulas.", vbCritical
        MkDir (IniPath & "Recursos\Dat\")
    End If

    If FileExist(IniPath & "Recursos\Graficos\", vbDirectory) = False Then
        MsgBox "Falta La Carpeta 'Recursos\Graficos\' Se crera y se guardara los mapas para el Server sin Particulas.", vbCritical
        MkDir (IniPath & "Recursos\Graficos\")
    End If

    If FileExist(IniPath & "Recursos\INIT\", vbDirectory) = False Then
        MsgBox "Falta La Carpeta 'Recursos\INIT\' Se crera y se guardara los mapas para el Server sin Particulas.", vbCritical
        MkDir (IniPath & "Recursos\INIT\")
    End If

    If FileExist(IniPath & "Recursos\MiniMapa\", vbDirectory) = False Then
        MsgBox "Falta La Carpeta 'Recursos\MiniMapa\' Se crera y se guardara los mapas para el Server sin Particulas.", vbCritical
        MkDir (IniPath & "Recursos\MiniMapa\")
    End If

    If FileExist(IniPath & "Renderizados\", vbDirectory) = False Then
        MsgBox "Falta La Carpeta 'Renderizados\' Se crera y se guardara los Mimi-mapas para WE.", vbCritical
        MkDir (IniPath & "Renderizados\")
    End If
    
    If FileExist(IniPath & "Renderizados\Minimapa\", vbDirectory) = False Then
        MsgBox "Falta La Carpeta 'Renderizados\Minimapa\' Se crera y se guardara los Mimi-mapas para WE.", vbCritical
        MkDir (IniPath & "Renderizados\Minimapa\")
    End If

End Sub

