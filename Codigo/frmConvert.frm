VERSION 5.00
Begin VB.Form frmConvert 
   Caption         =   "Conversor Integer > Long / Long > CSM"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Separar Automáticamente mapas para Cliente y para el Server sin partículas"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   6615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Sobrescribir automáticamente"
      Height          =   195
      Left            =   4080
      TabIndex        =   14
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Convertir Long > CSM"
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5160
      TabIndex        =   10
      Text            =   "2"
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Automatizar proceso"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convertir Integer > Long"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Text            =   "1"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Por ahora las particulas van solo en el mapa cliente por lo que los mapas para el server no deben llevarlas. "
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   3360
      Width           =   6255
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
      Left            =   5880
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
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
      Left            =   4440
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Hasta:"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   6375
   End
   Begin VB.Label Label5 
      Caption         =   $"frmConvert.frx":0000
      Height          =   855
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label4 
      Caption         =   "Instrucciones:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   975
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
      Left            =   2880
      TabIndex        =   3
      Top             =   2040
      Width           =   855
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
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Numero del mapa:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Automatico As Boolean

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

'************************************************
'Ultima modificacion 08/05/2020 por ReyarB
'************************************************
Private Sub Command1_Click()
Dim i As Integer
    If Automatico = False Then
        i = Text1.Text
        Call modMapIO.NuevoMapa
        Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Integer\Mapa" & Text1.Text & ".map", True)
        Call MapaV2_Guardar(App.Path & "\Conversor\Mapas Long\Mapa" & Text1.Text & ".map")
        Info.Caption = "Conversion realizada correctamente!"
    Else
        For i = Text1.Text To Text2.Text
            If FileExist(App.Path & "\Conversor\Mapas Integer\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Integer\Mapa" & i & ".map", True)
                Call MapaV2_Guardar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
                Info.Caption = "Mapa" & i & " convertido correctamente!"
            End If
        Next i
    End If
End Sub

Private Sub Command2_Click()
Dim i As Integer
    If Automatico = False Then
        Call modMapIO.NuevoMapa
        Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & Text1.Text & ".map")
        Call Save_CSM(App.Path & "\Conversor\Mapas CSM\Mapa" & Text1.Text & ".csm")
        
        Info.Caption = "Conversion realizada correctamente!"
    Else
        For i = Text1.Text To Text2.Text
            
            If FileExist(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
                Call Save_CSM(App.Path & "\Conversor\Mapas CSM\Mapa" & i & ".csm")
            
                Info.Caption = "Mapa" & i & " convertido correctamente!"
            End If
        Next i
    End If
End Sub

Private Sub Command3_Click()


Dim i As Integer
    If Automatico = False Then
        Call modMapIO.NuevoMapa
        Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & Text1.Text & ".map")
        Call MapaV5_Guardar(App.Path & "\Conversor\Mapas Server\Mapa" & Text1.Text & ".map")
        Call MapaV2_Guardar(App.Path & "\Conversor\Mapas Cliente\Mapa" & Text1.Text & ".map")
        Info.Caption = "Conversion realizada correctamente!"
    Else
        For i = Text1.Text To Text2.Text
            
            If FileExist(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map", vbNormal) = True Then
                Call modMapIO.NuevoMapa
                Call MapaV2_Cargar(App.Path & "\Conversor\Mapas Long\Mapa" & i & ".map")
                Call MapaV5_Guardar(App.Path & "\Conversor\Mapas Server\Mapa" & i & ".map")
                Call MapaV2_Guardar(App.Path & "\Conversor\Mapas Cliente\Mapa" & i & ".map")
                Info.Caption = "Mapa" & i & " convertido correctamente!"
            End If
        Next i
    End If
End Sub


