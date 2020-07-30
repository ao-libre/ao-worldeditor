VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRender 
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   7995
   ClientTop       =   3000
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton RenderizarMinimapM 
      Caption         =   "Renderizar Para Mundo WE"
      Height          =   735
      Left            =   7680
      TabIndex        =   16
      Top             =   2040
      Width           =   3855
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   120
      ScaleHeight     =   10.041
      ScaleMode       =   0  'User
      ScaleWidth      =   10.041
      TabIndex        =   1
      Top             =   840
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6960
      TabIndex        =   15
      Text            =   "1"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Automatizar proceso"
      Height          =   195
      Left            =   4920
      TabIndex        =   14
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9960
      TabIndex        =   6
      Text            =   "2"
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton RenderizarMinimap 
      Caption         =   "Renderizar Minimap para Cliente"
      Height          =   615
      Left            =   9000
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   615
      Left            =   10920
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Renderizar WE"
      Height          =   615
      Left            =   7440
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar pgbProgress 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Numero del mapa:"
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   ".map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Info 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label Label6 
      Caption         =   "Hasta:"
      Height          =   255
      Left            =   8640
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Mapa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   ".map"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10680
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblmapa 
      Height          =   255
      Left            =   7920
      TabIndex        =   3
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    '*************************************************************
     ' Capturar la imagen de controles
       
     '  1 - Colocar un picturebox llamado picture1, un Command1 y un Command2 _
        2 - Agragar algunos controles _
        3 - Indicar en la Sub " Capturar_Imagen " .. el control a capturar
    '*************************************************************
      
      
    ' Declaraciones del Api
      
    '*************************************************************
    ' Función BitBlt para copiar la imagen del control en un picturebox
    Private Automatico As Boolean
    Private Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long
      
    ' Recupera la imagen del área del control
    Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Sub Check1_Click()
    If Check1.value = False Then
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Text2.Visible = False
        Automatico = False
    Else
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Text2.Visible = True
        Automatico = True
    End If
End Sub


    '*************************************************************
      
    ' Sub que copia la imagen del control en un picturebox
    '*************************************************************
    Public Sub Capturar_Imagen(Control As Control, destino As Object)
          
        Dim hdc As Long
        Dim Escala_Anterior As Integer
        Dim ancho As Long
        Dim alto As Long
          
        ' Para que se mantenga la imagen por si se repinta la ventana
        destino.AutoRedraw = True
          
        On Error Resume Next
        ' Si da error es por que el control está dentro de un Frame _
          ya que  los Frame no tiene  dicha propiedad
        Escala_Anterior = Control.Container.ScaleMode
          
        If err.Number = 438 Then
           ' Si el control está en un Frame, convierte la escala
           ancho = ScaleX(Control.Width, vbTwips, vbPixels)
           alto = ScaleY(Control.Height, vbTwips, vbPixels)
        Else
           ' Si no cambia la escala del  contenedor a pixeles
           Control.Container.ScaleMode = vbPixels
           ancho = Control.Width
           alto = Control.Height
        End If
          
        ' limpia el error
        On Error GoTo 0
        ' Captura el área de pantalla correspondiente al control
        hdc = GetWindowDC(Control.hwnd)
        ' Copia esa área al picturebox
        BitBlt destino.hdc, 0, 0, ancho, alto, hdc, 0, 0, vbSrcCopy
        'BitBlt destino.hdc, 10, 10, ancho, alto, hdc, 10, 10, vbSrcCopy
        ' Convierte la imagen anterior en un Mapa de bits
        destino.Picture = destino.Image
        ' Borra la imagen ya que ahora usa el Picture
        destino.Cls
          
        On Error Resume Next
        If err.Number = 0 Then
           ' Si el control no está en un  Frame, restaura la escala del contenedor
           Control.Container.ScaleMode = Escala_Anterior
        End If
          
    End Sub

Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub cmdAceptar_Click()
    frmRender.Height = 4000
    Dim i As Integer

    If Automatico = False Then
        i = Text1.Text
            If FileExist(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
                Call MapCapture(0)
                Info.Caption = "Minimapas renderizados correctamente!"
                End If
    Else
        For i = Text1.Text To Text2.Text
            If FileExist(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
                Call MapCapture(0)
                Info.Caption = "Mapa" & i & " renderizado correctamente!"
            End If
        Next i
    End If

End Sub

'******************************************************
'Ultima modificacion 08/05/2020 por ReyarB
'******************************************************
Private Sub RenderizarMinimap_Click()
frmRender.Height = 3000

Dim i As Integer

    If Automatico = False Then
        i = Text1.Text
            If FileExist(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
                Call MapCapture(1)
                Info.Caption = "Minimapas renderizados correctamente!"
            End If
    Else
        For i = Text1.Text To Text2.Text
            If FileExist(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
                Call MapCapture(1)
                Info.Caption = "Mapa" & i & " renderizado correctamente!"
            End If
        Next i
    End If
End Sub


Private Sub RenderizarMinimapM_Click()
'******************************************************
'Ultima modificacion 08/05/2020 por ReyarB
'******************************************************
frmRender.Height = 3000

Dim i As Integer

    If Automatico = False Then
        i = Text1.Text
            If FileExist(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
                Call MapCapture(2)
                Info.Caption = "Minimapas renderizados correctamente!"
            End If
    Else
        For i = Text1.Text To Text2.Text
            If FileExist(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
                Call MapCapture(2)
                Info.Caption = "Mapa" & i & " renderizado correctamente!"
            End If
        Next i
    End If
End Sub
