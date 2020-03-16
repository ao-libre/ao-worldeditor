VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WorldEditor"
   ClientHeight    =   13020
   ClientLeft      =   390
   ClientTop       =   840
   ClientWidth     =   19095
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   868
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1273
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Magic 
      Caption         =   "Magic Button"
      Height          =   375
      Left            =   17760
      TabIndex        =   142
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   17640
      Picture         =   "frmMain.frx":628A
      ScaleHeight     =   660
      ScaleWidth      =   1455
      TabIndex        =   123
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox Minimap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   120
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   97
      Top             =   120
      Width           =   1500
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   1245
         Left            =   120
         Top             =   120
         Width           =   1245
      End
      Begin VB.Shape UserArea 
         BorderColor     =   &H80000004&
         Height          =   225
         Left            =   600
         Top             =   720
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1290
      Left            =   1680
      TabIndex        =   89
      Top             =   30
      Width           =   3225
      Begin WorldEditor.lvButtons_H cmdInformacionDelMapa 
         Height          =   375
         Left            =   100
         TabIndex        =   90
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Caption         =   "&Información del Mapa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Label lblMapAmbient 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2520
         TabIndex        =   141
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblAmbient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ambient:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   270
         Left            =   1680
         TabIndex        =   140
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label lblFNombreMapa 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre del Mapa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   270
         Left            =   105
         TabIndex        =   96
         Top             =   60
         Width           =   3015
      End
      Begin VB.Label lblFVersion 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Versión:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   285
         Left            =   105
         TabIndex        =   95
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lblFMusica 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Musica:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   270
         Left            =   105
         TabIndex        =   94
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label lblMapNombre 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo Mapa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1440
         TabIndex        =   93
         Top             =   90
         Width           =   900
      End
      Begin VB.Label lblMapMusica 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1080
         TabIndex        =   92
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblMapVersion 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1440
         TabIndex        =   91
         Top             =   1010
         Width           =   105
      End
   End
   Begin VB.PictureBox PreviewGrh 
      BackColor       =   &H00000000&
      FillColor       =   &H00C0C0C0&
      Height          =   4500
      Left            =   120
      ScaleHeight     =   4440
      ScaleWidth      =   4425
      TabIndex        =   87
      Top             =   7200
      Visible         =   0   'False
      Width           =   4485
   End
   Begin VB.PictureBox Renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   10215
      Left            =   4680
      ScaleHeight     =   679
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   951
      TabIndex        =   86
      Top             =   1440
      Width           =   14295
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   6
      Left            =   11760
      TabIndex        =   37
      Top             =   30
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1826
      Caption         =   "Tri&gger's (F12)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":9C38
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   5
      Left            =   10320
      TabIndex        =   36
      Top             =   30
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Objetos (F11)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":A1FE
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   3
      Left            =   8955
      TabIndex        =   35
      Top             =   30
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1826
      Caption         =   "&NPC's (F8)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":A6FF
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   2
      Left            =   7440
      TabIndex        =   34
      Top             =   30
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Bloqueos (F7)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":AAB3
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   1
      Left            =   5925
      TabIndex        =   33
      Top             =   30
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Translados (F6)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":AE34
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   0
      Left            =   5160
      TabIndex        =   32
      Top             =   30
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1826
      Caption         =   "&Superficie (F5)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":E494
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdQuitarFunciones 
      Height          =   435
      Left            =   1800
      TabIndex        =   31
      ToolTipText     =   "Quitar Todas las Funciones Activadas"
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   767
      Caption         =   "&Quitar Funciones (F4)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      cBack           =   12632319
   End
   Begin VB.Timer TimAutoGuardarMapa 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3960
      Top             =   1920
   End
   Begin VB.PictureBox pPaneles 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5385
      Left            =   120
      Picture         =   "frmMain.frx":119DA
      ScaleHeight     =   5355
      ScaleWidth      =   4425
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
      Begin VB.Frame cLuces 
         BackColor       =   &H00000000&
         Caption         =   "Luces"
         ForeColor       =   &H00FFFFFF&
         Height          =   4155
         Left            =   120
         TabIndex        =   124
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Frame Frame3 
            BackColor       =   &H00000000&
            Caption         =   "Luz Base"
            ForeColor       =   &H00FFFFFF&
            Height          =   1335
            Left            =   120
            TabIndex        =   132
            Top             =   2760
            Width           =   3855
            Begin WorldEditor.lvButtons_H lvButtons_H1 
               Height          =   360
               Left            =   360
               TabIndex        =   133
               Top             =   360
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   635
               Caption         =   "Mañana"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cGradient       =   0
               Mode            =   1
               Value           =   0   'False
               cBack           =   8438015
            End
            Begin WorldEditor.lvButtons_H lvButtons_H2 
               Height          =   360
               Left            =   2040
               TabIndex        =   134
               Top             =   360
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   635
               Caption         =   "Dia"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cGradient       =   0
               Mode            =   1
               Value           =   0   'False
               cBack           =   16777088
            End
            Begin WorldEditor.lvButtons_H lvButtons_H3 
               Height          =   360
               Left            =   360
               TabIndex        =   135
               Top             =   840
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   635
               Caption         =   "Tarde"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cGradient       =   0
               Mode            =   1
               Value           =   0   'False
               cBack           =   8421504
            End
            Begin WorldEditor.lvButtons_H lvButtons_H4 
               Height          =   360
               Left            =   2040
               TabIndex        =   136
               Top             =   840
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   635
               Caption         =   "Noche"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cGradient       =   0
               Mode            =   1
               Value           =   0   'False
               cBack           =   4210752
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00000000&
            Caption         =   "Rango"
            ForeColor       =   &H00FFFFFF&
            Height          =   660
            Left            =   1320
            TabIndex        =   129
            Top             =   1080
            Width           =   1380
            Begin VB.TextBox cRango 
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   315
               Left            =   105
               TabIndex        =   130
               Text            =   "1"
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "(1 al 50)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   131
               Top             =   270
               Width           =   615
            End
         End
         Begin VB.Frame RGBCOLOR 
            BackColor       =   &H00000000&
            Caption         =   "RGB"
            ForeColor       =   &H00FFFFFF&
            Height          =   690
            Left            =   1200
            TabIndex        =   125
            Top             =   360
            Width           =   1680
            Begin VB.TextBox R 
               BackColor       =   &H80000012&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   315
               Left            =   105
               TabIndex        =   128
               Text            =   "1"
               Top             =   270
               Width           =   450
            End
            Begin VB.TextBox B 
               BackColor       =   &H80000012&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   315
               Left            =   1095
               TabIndex        =   127
               Text            =   "1"
               Top             =   270
               Width           =   450
            End
            Begin VB.TextBox G 
               BackColor       =   &H80000012&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   315
               Left            =   600
               TabIndex        =   126
               Text            =   "1"
               Top             =   270
               Width           =   450
            End
         End
         Begin WorldEditor.lvButtons_H cInsertarLuz 
            Height          =   360
            Left            =   2160
            TabIndex        =   137
            Top             =   1800
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   635
            Caption         =   "Insertar Luz"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H cQuitarLuz 
            Height          =   360
            Left            =   360
            TabIndex        =   138
            Top             =   1800
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   635
            Caption         =   "Quitar Luz"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nota: Para quitar una luz ya guardada, insertar una luz encima y despues quitar."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   360
            TabIndex        =   139
            Top             =   2280
            Width           =   3615
         End
      End
      Begin VB.TextBox Life 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   118
         Text            =   "-1"
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ListBox lstParticle 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2205
         Left            =   120
         TabIndex        =   117
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Frame CopyBorder 
         BackColor       =   &H00000000&
         Caption         =   "Copiar bordes"
         ForeColor       =   &H00FFFFFF&
         Height          =   3975
         Left            =   120
         TabIndex        =   99
         Top             =   240
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox TXTArriba 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1680
            TabIndex        =   103
            Text            =   "80"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox TxTAbajo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1680
            TabIndex        =   102
            Text            =   "80"
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox TxTDerecha 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3240
            TabIndex        =   101
            Text            =   "78"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox TxTIzquierda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   360
            TabIndex        =   100
            Text            =   "78"
            Top             =   1440
            Width           =   615
         End
         Begin WorldEditor.lvButtons_H COPIAR_GRH 
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   104
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Caption         =   "Pegar borde abajo"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
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
         Begin WorldEditor.lvButtons_H COPIAR_GRH 
            Height          =   375
            Index           =   1
            Left            =   2160
            TabIndex        =   105
            Top             =   1800
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Pegar borde derecha"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
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
         Begin WorldEditor.lvButtons_H COPIAR_GRH 
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   106
            Top             =   1800
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            Caption         =   "Pegar borde izquierda"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
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
         Begin WorldEditor.lvButtons_H COPIAR_GRH 
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   107
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Caption         =   "Pegar borde arriba"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
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
         Begin WorldEditor.lvButtons_H MemoriaAuxiliar 
            Height          =   495
            Left            =   360
            TabIndex        =   108
            Top             =   360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   873
            Caption         =   "Copiar bordes del mapa en memoria auxiliar"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   16744576
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Estos valores se miden en Tiles y JAMAS podran superar los 100 tiles o estar por debajo de 1 tiles."
            ForeColor       =   &H00FFFFFF&
            Height          =   555
            Left            =   480
            TabIndex        =   110
            Top             =   3360
            Width           =   3810
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "¡ATENCION!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   109
            Top             =   3120
            Width           =   1065
         End
      End
      Begin VB.TextBox tTY 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   76
         Text            =   "1"
         Top             =   960
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTX 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   75
         Text            =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTMapa 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   74
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   2900
      End
      Begin WorldEditor.lvButtons_H cInsertarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   77
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Insertar Translado"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTransOBJ 
         Height          =   375
         Left            =   240
         TabIndex        =   78
         Top             =   1680
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Colocar automaticamente &Objeto"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionManual 
         Height          =   375
         Left            =   240
         TabIndex        =   79
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Union con Mapa Adyacente (manual)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionAuto 
         Height          =   375
         Left            =   240
         TabIndex        =   80
         Top             =   2520
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Union con Mapas &Adyacentes (auto)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   81
         Top             =   3000
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Quitar Translados"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCapas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         ItemData        =   "frmMain.frx":5F074
         Left            =   1080
         List            =   "frmMain.frx":5F084
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Left            =   2880
         TabIndex        =   63
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         Left            =   600
         TabIndex        =   62
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   0
         ItemData        =   "frmMain.frx":5F094
         Left            =   120
         List            =   "frmMain.frx":5F096
         Sorted          =   -1  'True
         TabIndex        =   61
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin WorldEditor.lvButtons_H cQuitarEnTodasLasCapas 
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Quitar en &Capas 2 y 3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarEnEstaCapa 
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar en esta Capa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cSeleccionarSuperficie 
         Height          =   735
         Left            =   2400
         TabIndex        =   66
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Superficie"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":5F098
         Left            =   3360
         List            =   "frmMain.frx":5F09A
         TabIndex        =   57
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":5F09C
         Left            =   840
         List            =   "frmMain.frx":5F09E
         TabIndex        =   0
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   3
         ItemData        =   "frmMain.frx":5F0A0
         Left            =   120
         List            =   "frmMain.frx":5F0A2
         TabIndex        =   56
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   3
         Left            =   600
         TabIndex        =   55
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":5F0A4
         Left            =   840
         List            =   "frmMain.frx":5F0A6
         TabIndex        =   48
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":5F0A8
         Left            =   3360
         List            =   "frmMain.frx":5F0AA
         TabIndex        =   47
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         Left            =   600
         TabIndex        =   46
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   1
         ItemData        =   "frmMain.frx":5F0AC
         Left            =   120
         List            =   "frmMain.frx":5F0AE
         TabIndex        =   45
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   3210
         Index           =   4
         ItemData        =   "frmMain.frx":5F0B0
         Left            =   120
         List            =   "frmMain.frx":5F0B2
         TabIndex        =   44
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.PictureBox Picture5 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   3
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture6 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   4
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture7 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   5
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture8 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   6
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture9 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   7
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture11 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   40
         Top             =   0
         Width           =   0
      End
      Begin WorldEditor.lvButtons_H cQuitarTrigger 
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerTriggers 
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Mostrar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTrigger 
         Height          =   735
         Left            =   2400
         TabIndex        =   43
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Trigger"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   0
         Left            =   2400
         TabIndex        =   51
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerBloqueos 
         Height          =   495
         Left            =   120
         TabIndex        =   52
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         Caption         =   "&Mostrar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Insertar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   54
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Quitar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   58
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar OBJ's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   59
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar OBJ's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   2
         Left            =   2400
         TabIndex        =   60
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar Objetos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   735
         Index           =   1
         Left            =   2400
         TabIndex        =   73
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   72
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   71
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":5F0B4
         Left            =   840
         List            =   "frmMain.frx":5F0B6
         TabIndex        =   67
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         Left            =   600
         TabIndex        =   68
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   2
         ItemData        =   "frmMain.frx":5F0B8
         Left            =   120
         List            =   "frmMain.frx":5F0BA
         TabIndex        =   69
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":5F0BC
         Left            =   3360
         List            =   "frmMain.frx":5F0BE
         TabIndex        =   70
         Text            =   "500"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin WorldEditor.lvButtons_H cmdAdd 
         Height          =   375
         Left            =   1320
         TabIndex        =   120
         Top             =   3720
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&Agregar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cmdDel 
         Height          =   375
         Left            =   1320
         TabIndex        =   121
         Top             =   4080
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&Quitar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "LiveCounter:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   119
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lYver 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Y vertical:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   84
         Top             =   1005
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lXhor 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "X horizontal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   83
         Top             =   645
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lMapN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Mapa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   82
         Top             =   285
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lbCapas 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Capa Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   3195
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lbGrh 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Sup Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   2040
         TabIndex        =   17
         Top             =   3195
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   2160
         TabIndex        =   16
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de OBJ:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   2160
         TabIndex        =   13
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   2160
         TabIndex        =   9
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2565
      Top             =   1905
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   675
      Index           =   4
      Left            =   9840
      TabIndex        =   85
      Top             =   240
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1191
      Caption         =   "none"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":5F0C0
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin RichTextLib.RichTextBox StatTxt 
      Height          =   1155
      Left            =   120
      TabIndex        =   88
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   11760
      Width           =   18795
      _ExtentX        =   33152
      _ExtentY        =   2037
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":5F474
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   7
      Left            =   13080
      TabIndex        =   98
      Top             =   30
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1826
      Caption         =   "&Copiar Bordes"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":5F4F1
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   8
      Left            =   14445
      TabIndex        =   116
      Top             =   30
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1826
      Caption         =   "&Particulas"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":5FB32
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   9
      Left            =   15810
      TabIndex        =   122
      Top             =   30
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1826
      Caption         =   "Luces "
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":601B4
      ImgSize         =   24
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   18240
      TabIndex        =   115
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   17460
      TabIndex        =   114
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   16680
      TabIndex        =   113
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   15120
      TabIndex        =   112
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   15900
      TabIndex        =   111
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Line Separacion1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   329
      X2              =   329
      Y1              =   8
      Y2              =   88
   End
   Begin VB.Line Separacion2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   337
      X2              =   337
      Y1              =   8
      Y2              =   88
   End
   Begin VB.Line Separacion2 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   336
      X2              =   336
      Y1              =   8
      Y2              =   88
   End
   Begin VB.Line Separacion1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   328
      X2              =   328
      Y1              =   8
      Y2              =   88
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   14340
      TabIndex        =   39
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   13575
      TabIndex        =   38
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5925
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   6690
      TabIndex        =   29
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   7455
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   8220
      TabIndex        =   27
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   8985
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   9750
      TabIndex        =   25
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   10515
      TabIndex        =   24
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   11280
      TabIndex        =   23
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   12045
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   12810
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNuevoMapa 
         Caption         =   "&Nuevo Mapa"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAbrirMapaNew 
         Caption         =   "&Abrir Mapa"
      End
      Begin VB.Menu mnuArchivoLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReAbrirMapa 
         Caption         =   "&Re-Abrir Mapa"
      End
      Begin VB.Menu mnuArchivoLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarMapa 
         Caption         =   "&Guardar Mapa"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGuardarMapaComo 
         Caption         =   "Guardar Mapa &como..."
      End
      Begin VB.Menu mnuArchivoLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "&Conversor"
      End
      Begin VB.Menu mnuRenderMapa 
         Caption         =   "Renderizar Mapa"
      End
      Begin VB.Menu mnuArchivoLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
      Begin VB.Menu mnuArchivoLine6 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuComo 
         Caption         =   "¿ Como seleccionar ? ---- Mantener SHIFT y arrastrar el cursor."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCortar 
         Caption         =   "C&ortar Selección"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "&Copiar Selección"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "&Pegar Selección"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuBloquearS 
         Caption         =   "&Bloquear Selección"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuRealizarOperacion 
         Caption         =   "&Realizar Operación en Selección"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDeshacerPegado 
         Caption         =   "Deshacer P&egado de Selección"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLineEdicion0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeshacer 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuUtilizarDeshacer 
         Caption         =   "&Utilizar Deshacer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuInfoMap 
         Caption         =   "&Información del Mapa"
      End
      Begin VB.Menu mnuLineEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertar 
         Caption         =   "&Insertar"
         Begin VB.Menu mnuInsertarTransladosAdyasentes 
            Caption         =   "&Translados a Mapas Adyasentes"
         End
         Begin VB.Menu mnuInsertarSuperficieAlAzar 
            Caption         =   "Superficie al &Azar"
         End
         Begin VB.Menu mnuInsertarSuperficieEnBordes 
            Caption         =   "Superficie en los &Bordes del Mapa"
         End
         Begin VB.Menu mnuInsertarSuperficieEnTodo 
            Caption         =   "Superficie en Todo el Mapa"
         End
         Begin VB.Menu mnuBloquearBordes 
            Caption         =   "Bloqueo en &Bordes del Mapa"
         End
         Begin VB.Menu mnuBloquearMapa 
            Caption         =   "Bloqueo en &Todo el Mapa"
         End
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "&Quitar"
         Begin VB.Menu mnuQuitarTranslados 
            Caption         =   "Todos los &Translados"
         End
         Begin VB.Menu mnuQuitarBloqueos 
            Caption         =   "Todos los &Bloqueos"
         End
         Begin VB.Menu mnuQuitarNPCs 
            Caption         =   "Todos los &NPC's"
         End
         Begin VB.Menu mnuQuitarNPCsHostiles 
            Caption         =   "Todos los NPC's &Hostiles"
         End
         Begin VB.Menu mnuQuitarObjetos 
            Caption         =   "Todos los &Objetos"
         End
         Begin VB.Menu mnuQuitarTriggers 
            Caption         =   "Todos los Tri&gger's"
         End
         Begin VB.Menu mnuQuitarSuperficieBordes 
            Caption         =   "Superficie de los B&ordes"
         End
         Begin VB.Menu mnuQuitarSuperficieDeCapa 
            Caption         =   "Superficie de la &Capa Seleccionada"
         End
         Begin VB.Menu mnuLineEdicion2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuitarTODO 
            Caption         =   "TODO"
         End
      End
      Begin VB.Menu mnuLineEdicion3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunciones 
         Caption         =   "&Funciones"
         Begin VB.Menu mnuQuitarFunciones 
            Caption         =   "&Quitar Funciones"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuAutoQuitarFunciones 
            Caption         =   "Auto-&Quitar Funciones"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuConfigAvanzada 
         Caption         =   "Configuracion A&vanzada de Superficie"
      End
      Begin VB.Menu mnuLineEdicion4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCompletarSuperficies 
         Caption         =   "Auto-Completar &Superficies"
      End
      Begin VB.Menu mnuAutoCapturarSuperficie 
         Caption         =   "Auto-C&apturar información de la Superficie"
      End
      Begin VB.Menu mnuAutoCapturarTranslados 
         Caption         =   "Auto-&Capturar información de los Translados"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoGuardarMapas 
         Caption         =   "Configuración de Auto-&Guardar Mapas"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuCapas 
         Caption         =   "...&Capas"
         Begin VB.Menu mnuVerCapa1 
            Caption         =   "Capa &1 (Piso)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa2 
            Caption         =   "Capa &2 (costas, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa3 
            Caption         =   "Capa &3 (arboles, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa4 
            Caption         =   "Capa &4 (techos, etc)"
         End
      End
      Begin VB.Menu mnuVerTranslados 
         Caption         =   "...&Translados"
      End
      Begin VB.Menu mnuVerBloqueos 
         Caption         =   "...&Bloqueos"
      End
      Begin VB.Menu mnuVerNPCs 
         Caption         =   "...&NPC's"
      End
      Begin VB.Menu mnuVerObjetos 
         Caption         =   "...&Objetos"
      End
      Begin VB.Menu mnuVerTriggers 
         Caption         =   "...Tri&gger's"
      End
      Begin VB.Menu mnuVerGrilla 
         Caption         =   "...Gri&lla"
      End
      Begin VB.Menu mnuLinMostrar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerAutomatico 
         Caption         =   "Control &Automaticamente"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPaneles 
      Caption         =   "&Paneles"
      Begin VB.Menu mnuSuperficie 
         Caption         =   "&Superficie"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTranslados 
         Caption         =   "&Translados"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBloquear 
         Caption         =   "&Bloquear"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuNPCs 
         Caption         =   "&NPC's"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuNPCsHostiles 
         Caption         =   "NPC's &Hostiles"
         Shortcut        =   {F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuObjetos 
         Caption         =   "&Objetos"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuTriggers 
         Caption         =   "Tri&gger's"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPanelesLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQSuperficie 
         Caption         =   "Ocultar Superficie"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuQTranslados 
         Caption         =   "Ocultar Translados"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuQBloquear 
         Caption         =   "Ocultar Bloquear"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuQNPCs 
         Caption         =   "Ocultar NPC's"
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuQNPCsHostiles 
         Caption         =   "Ocultar NPC's Hostiles"
         Shortcut        =   +{F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQObjetos 
         Caption         =   "Ocultar Objetos"
         Shortcut        =   +{F11}
      End
      Begin VB.Menu mnuQTriggers 
         Caption         =   "Ocultar Trigger's"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu mnuFuncionesLine1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuInformes 
         Caption         =   "&Informes"
      End
      Begin VB.Menu mnuActualizarIndices 
         Caption         =   "&Actualizar Indices de..."
         Begin VB.Menu mnuActualizarSuperficies 
            Caption         =   "&Superficies"
         End
         Begin VB.Menu mnuActualizarNPCs 
            Caption         =   "&NPC's"
         End
         Begin VB.Menu mnuActualizarObjs 
            Caption         =   "&Objetos"
         End
         Begin VB.Menu mnuActualizarTriggers 
            Caption         =   "&Trigger's"
         End
         Begin VB.Menu mnuActualizarCabezas 
            Caption         =   "C&abezas"
         End
         Begin VB.Menu mnuActualizarCuerpos 
            Caption         =   "C&uerpos"
         End
         Begin VB.Menu mnuActualizarGraficos 
            Caption         =   "&Graficos"
         End
      End
      Begin VB.Menu mnuModoCaminata 
         Caption         =   "Modalidad &Caminata"
      End
      Begin VB.Menu mnuGRHaBMP 
         Caption         =   "&GRH => BMP"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptimizar 
         Caption         =   "Optimi&zar Mapa"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarUltimaConfig 
         Caption         =   "&Guardar Ultima Configuración"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuManual 
         Caption         =   "&Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLineAyuda1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu mnuObjSc 
      Caption         =   "mnuObjSc"
      Visible         =   0   'False
      Begin VB.Menu mnuConfigObjTrans 
         Caption         =   "&Utilizar como Objeto de Translados"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Public MouseX As Integer

Public MouseY As Integer

Private Sub PonerAlAzar(ByVal n As Integer, T As Byte)

    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06 by GS
    '*************************************************
    Dim objindex As Long
    Dim NPCIndex As Long
    Dim X, Y, i
    Dim Head    As Integer
    Dim Body    As Integer
    Dim Heading As Byte
    Dim Leer    As New clsIniManager

    i = n

    Call modEdicion.Deshacer_Add("Aplicar " & IIf(T = 0, "Objetos", "NPCs") & " al Azar") ' Hago deshacer

    Do While i > 0
        X = CInt(RandomNumber(XMinMapSize, XMaxMapSize - 1))
        Y = CInt(RandomNumber(YMinMapSize, YMaxMapSize - 1))
    
        Select Case T

            Case 0

                If MapData(X, Y).OBJInfo.objindex = 0 Then
                    i = i - 1

                    If cInsertarBloqueo.Value = True Then
                        MapData(X, Y).blocked = 1
                    Else
                        MapData(X, Y).blocked = 0

                    End If

                    If cNumFunc(2).Text > 0 Then
                        objindex = cNumFunc(2).Text
                        InitGrh MapData(X, Y).ObjGrh, ObjData(objindex).GrhIndex
                        MapData(X, Y).OBJInfo.objindex = objindex
                        MapData(X, Y).OBJInfo.Amount = Val(cCantFunc(2).Text)

                        Select Case ObjData(objindex).ObjType ' GS

                            Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                MapData(X, Y).Graphic(3) = MapData(X, Y).ObjGrh

                        End Select

                    End If

                End If

            Case 1

                If MapData(X, Y).blocked = 0 Then
                    i = i - 1

                    If cNumFunc(T - 1).Text > 0 Then
                        NPCIndex = cNumFunc(T - 1).Text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                        MapData(X, Y).NPCIndex = NPCIndex

                    End If

                End If

            Case 2

                If MapData(X, Y).blocked = 0 Then
                    i = i - 1

                    If cNumFunc(T - 1).Text >= 0 Then
                        NPCIndex = cNumFunc(T - 1).Text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(Y))
                        MapData(X, Y).NPCIndex = NPCIndex

                    End If

                End If

        End Select

        DoEvents
    Loop

End Sub

Private Sub cAgregarFuncalAzar_Click(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    If IsNumeric(cCantFunc(index).Text) = False Or cCantFunc(index).Text > 200 Then
        MsgBox "El Valor de Cantidad introducido no es soportado!" & vbCrLf & "El valor maximo es 200.", vbCritical
        Exit Sub

    End If

    cAgregarFuncalAzar(index).Enabled = False
    Call PonerAlAzar(CInt(cCantFunc(index).Text), 1 + (IIf(index = 2, -1, index)))
    cAgregarFuncalAzar(index).Enabled = True

End Sub

Private Sub cCantFunc_Change(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    If Val(cCantFunc(index)) < 1 Then
        cCantFunc(index).Text = 1

    End If

    If Val(cCantFunc(index)) > 10000 Then
        cCantFunc(index).Text = 10000

    End If

End Sub

Private Sub cCapas_Change()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 31/05/06
    '*************************************************
    If Val(cCapas.Text) < 1 Then
        cCapas.Text = 1

    End If

    If Val(cCapas.Text) > 4 Then
        cCapas.Text = 4

    End If

    cCapas.Tag = vbNullString

End Sub

Private Sub cCapas_KeyPress(KeyAscii As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0

End Sub

Private Sub cFiltro_GotFocus(index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    HotKeysAllow = False

End Sub

Private Sub cFiltro_KeyPress(index As Integer, KeyAscii As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If KeyAscii = 13 Then
        Call Filtrar(index)

    End If

End Sub

Private Sub cFiltro_LostFocus(index As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    HotKeysAllow = True

End Sub

Private Sub cGrh_KeyPress(KeyAscii As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo Fallo

    If KeyAscii = 13 Then
        Call fPreviewGrh(cGrh.Text)

        If frmMain.PreviewGrh.Visible = True Then
            Call modPaneles.VistaPreviaDeSup

        End If

        If frmMain.cGrh.ListCount > 5 Then
            frmMain.cGrh.RemoveItem 0

        End If

        frmMain.cGrh.AddItem frmMain.cGrh.Text

    End If

    Exit Sub
Fallo:
    cGrh.Text = 1

End Sub

Private Sub cInsertarFunc_Click(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cInsertarFunc(index).Value = True Then
        cQuitarFunc(index).Enabled = False
        cAgregarFuncalAzar(index).Enabled = False

        If index <> 2 Then cCantFunc(index).Enabled = False
        Call modPaneles.EstSelectPanel((index) + 3, True)
    Else
        cQuitarFunc(index).Enabled = True
        cAgregarFuncalAzar(index).Enabled = True

        If index <> 2 Then cCantFunc(index).Enabled = True
        Call modPaneles.EstSelectPanel((index) + 3, False)

    End If

End Sub

Private Sub cInsertarLuz_Click()

    If cInsertarLuz.Value Then
        cQuitarLuz.Enabled = False
    Else
        cQuitarLuz.Enabled = True

    End If

End Sub

Private Sub cInsertarTrans_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    If cInsertarTrans.Value = True Then
        cQuitarTrans.Enabled = False
        Call modPaneles.EstSelectPanel(1, True)
    Else
        cQuitarTrans.Enabled = True
        Call modPaneles.EstSelectPanel(1, False)

    End If

End Sub

Private Sub cInsertarTrigger_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cInsertarTrigger.Value = True Then
        cQuitarTrigger.Enabled = False
        Call modPaneles.EstSelectPanel(6, True)
    Else
        cQuitarTrigger.Enabled = True
        Call modPaneles.EstSelectPanel(6, False)

    End If

End Sub

Private Sub cmdAdd_Click()

    If cmdAdd.Value = True Then
        cmdDel.Enabled = False
        Call modPaneles.EstSelectPanel(8, True)
    Else
        cmdDel.Enabled = True
        Call modPaneles.EstSelectPanel(8, False)

    End If

End Sub

Private Sub cmdDel_Click()

    If cmdDel.Value = True Then
        lstParticle.Enabled = False
        cmdAdd.Enabled = False
        Call modPaneles.EstSelectPanel(8, True)
    Else
        lstParticle.Enabled = True
        cmdAdd.Enabled = True
        Call modPaneles.EstSelectPanel(8, False)

    End If

End Sub

Private Sub cmdInformacionDelMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmMapInfo.Show
    frmMapInfo.Visible = True

End Sub

Private Sub cmdQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call mnuQuitarFunciones_Click

End Sub

Private Sub COPIAR_GRH_Click(index As Integer)

    On Error Resume Next

    Dim Y As Integer
    Dim X As Integer
 
    Select Case index

        Case 0 'Arriba

            For Y = 1 To 10
                For X = 1 To 100
                    MapData(X, Y).Graphic(1) = MapData_Adyacente(X, TXTArriba + Y).Graphic(1)
                    MapData(X, Y).Graphic(2) = MapData_Adyacente(X, TXTArriba + Y).Graphic(2)
                    MapData(X, Y).Graphic(3) = MapData_Adyacente(X, TXTArriba + Y).Graphic(3)
                    MapData(X, Y).Graphic(4) = MapData_Adyacente(X, TXTArriba + Y).Graphic(4)
                Next
            Next
            MapInfo.Changed = 1
               
        Case 1 'Derecha

            For Y = 1 To 100
                For X = 90 To 100
                    MapData(X, Y).Graphic(1) = MapData_Adyacente(X - TxTDerecha, Y).Graphic(1)
                    MapData(X, Y).Graphic(2) = MapData_Adyacente(X - TxTDerecha, Y).Graphic(2)
                    MapData(X, Y).Graphic(3) = MapData_Adyacente(X - TxTDerecha, Y).Graphic(3)
                    MapData(X, Y).Graphic(4) = MapData_Adyacente(X - TxTDerecha, Y).Graphic(4)
                Next
            Next
            MapInfo.Changed = 1
               
        Case 2 'Izquierda

            For Y = 1 To 100
                For X = 1 To 11
                    MapData(X, Y).Graphic(1) = MapData_Adyacente(X + TxTIzquierda, Y).Graphic(1)
                    MapData(X, Y).Graphic(2) = MapData_Adyacente(X + TxTIzquierda, Y).Graphic(2)
                    MapData(X, Y).Graphic(3) = MapData_Adyacente(X + TxTIzquierda, Y).Graphic(3)
                    MapData(X, Y).Graphic(4) = MapData_Adyacente(X + TxTIzquierda, Y).Graphic(4)
                Next
            Next
            MapInfo.Changed = 1
               
        Case 3 'Abajo

            For Y = 90 To 100
                For X = 1 To 100
                    MapData(X, Y).Graphic(1) = MapData_Adyacente(X, Y - TxTAbajo).Graphic(1)
                    MapData(X, Y).Graphic(2) = MapData_Adyacente(X, Y - TxTAbajo).Graphic(2)
                    MapData(X, Y).Graphic(3) = MapData_Adyacente(X, Y - TxTAbajo).Graphic(3)
                    MapData(X, Y).Graphic(4) = MapData_Adyacente(X, Y - TxTAbajo).Graphic(4)
                Next
            Next
            MapInfo.Changed = 1

    End Select

End Sub

Private Sub cQuitarLuz_Click()

    '*************************************************
    'Author: Lorwik
    '*************************************************
    If cQuitarLuz.Value Then
        cInsertarLuz.Enabled = False
    Else
        cInsertarLuz.Enabled = True

    End If

End Sub

Private Sub cUnionManual_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    cInsertarTrans.Value = (cUnionManual.Value = True)
    Call cInsertarTrans_Click

End Sub

Private Sub cverBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuVerBloqueos.Checked = cVerBloqueos.Value

End Sub

Private Sub cverTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuVerTriggers.Checked = cVerTriggers.Value

End Sub

Private Sub cNumFunc_KeyPress(index As Integer, KeyAscii As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    If KeyAscii = 13 Then

        Dim Cont As String

        Cont = frmMain.cNumFunc(index).Text
        Call cNumFunc_LostFocus(index)

        If Cont <> frmMain.cNumFunc(index).Text Then Exit Sub
        If frmMain.cNumFunc(index).ListCount > 5 Then
            frmMain.cNumFunc(index).RemoveItem 0

        End If

        frmMain.cNumFunc(index).AddItem frmMain.cNumFunc(index).Text
        Exit Sub
    ElseIf KeyAscii = 8 Then
    
    ElseIf IsNumeric(Chr(KeyAscii)) = False Then
        KeyAscii = 0
        Exit Sub

    End If

End Sub

Private Sub cNumFunc_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    If cNumFunc(index).Text = vbNullString Then
        frmMain.cNumFunc(index).Text = IIf(index = 1, 500, 1)

    End If

End Sub

Private Sub cNumFunc_LostFocus(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    If index = 0 Then
        If frmMain.cNumFunc(index).Text > 499 Or frmMain.cNumFunc(index).Text < 1 Then
            frmMain.cNumFunc(index).Text = 1

        End If

    ElseIf index = 1 Then

        If frmMain.cNumFunc(index).Text < 500 Or frmMain.cNumFunc(index).Text > 32000 Then
            frmMain.cNumFunc(index).Text = 500

        End If

    ElseIf index = 2 Then

        If frmMain.cNumFunc(index).Text < 1 Or frmMain.cNumFunc(index).Text > 32000 Then
            frmMain.cNumFunc(index).Text = 1

        End If

    End If

End Sub

Private Sub cInsertarBloqueo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    cInsertarBloqueo.Tag = vbNullString

    If cInsertarBloqueo.Value = True Then
        cQuitarBloqueo.Enabled = False
        Call modPaneles.EstSelectPanel(2, True)
    Else
        cQuitarBloqueo.Enabled = True
        Call modPaneles.EstSelectPanel(2, False)

    End If

End Sub

Private Sub cQuitarBloqueo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    cInsertarBloqueo.Tag = vbNullString

    If cQuitarBloqueo.Value = True Then
        cInsertarBloqueo.Enabled = False
        Call modPaneles.EstSelectPanel(2, True)
    Else
        cInsertarBloqueo.Enabled = True
        Call modPaneles.EstSelectPanel(2, False)

    End If

End Sub

Private Sub cQuitarEnEstaCapa_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarEnEstaCapa.Value = True Then
        lListado(0).Enabled = False
        cFiltro(0).Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnTodasLasCapas.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        lListado(0).Enabled = True
        cFiltro(0).Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnTodasLasCapas.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

End Sub

Private Sub cQuitarEnTodasLasCapas_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarEnTodasLasCapas.Value = True Then
        cCapas.Enabled = False
        lListado(0).Enabled = False
        cFiltro(0).Enabled = False
        cGrh.Enabled = False
        cSeleccionarSuperficie.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        cCapas.Enabled = True
        lListado(0).Enabled = True
        cFiltro(0).Enabled = True
        cGrh.Enabled = True
        cSeleccionarSuperficie.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

End Sub

Private Sub cQuitarFunc_Click(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarFunc(index).Value = True Then
        cInsertarFunc(index).Enabled = False
        cAgregarFuncalAzar(index).Enabled = False
        cCantFunc(index).Enabled = False
        cNumFunc(index).Enabled = False
        cFiltro((index) + 1).Enabled = False
        lListado((index) + 1).Enabled = False
        Call modPaneles.EstSelectPanel((index) + 3, True)
    Else
        cInsertarFunc(index).Enabled = True
        cAgregarFuncalAzar(index).Enabled = True
        cCantFunc(index).Enabled = True
        cNumFunc(index).Enabled = True
        cFiltro((index) + 1).Enabled = True
        lListado((index) + 1).Enabled = True
        Call modPaneles.EstSelectPanel((index) + 3, False)

    End If

End Sub

Private Sub cQuitarTrans_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarTrans.Value = True Then
        cInsertarTransOBJ.Enabled = False
        cInsertarTrans.Enabled = False
        cUnionManual.Enabled = False
        cUnionAuto.Enabled = False
        tTMapa.Enabled = False
        tTX.Enabled = False
        tTY.Enabled = False
        mnuInsertarTransladosAdyasentes.Enabled = False
        Call modPaneles.EstSelectPanel(1, True)
    Else
        tTMapa.Enabled = True
        tTX.Enabled = True
        tTY.Enabled = True
        cUnionAuto.Enabled = True
        cUnionManual.Enabled = True
        cInsertarTrans.Enabled = True
        cInsertarTransOBJ.Enabled = True
        mnuInsertarTransladosAdyasentes.Enabled = True
        Call modPaneles.EstSelectPanel(1, False)

    End If

End Sub

Private Sub cQuitarTrigger_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cQuitarTrigger.Value = True Then
        lListado(4).Enabled = False
        cInsertarTrigger.Enabled = False
        Call modPaneles.EstSelectPanel(6, True)
    Else
        lListado(4).Enabled = True
        cInsertarTrigger.Enabled = True
        Call modPaneles.EstSelectPanel(6, False)

    End If

End Sub

Private Sub cSeleccionarSuperficie_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If cSeleccionarSuperficie.Value = True Then
        cQuitarEnTodasLasCapas.Enabled = False
        cQuitarEnEstaCapa.Enabled = False
        Call modPaneles.EstSelectPanel(0, True)
    Else
        cQuitarEnTodasLasCapas.Enabled = True
        cQuitarEnEstaCapa.Enabled = True
        Call modPaneles.EstSelectPanel(0, False)

    End If

End Sub

Private Sub cUnionAuto_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmUnionAdyacente.Show

End Sub

Private Sub Form_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Me.SetFocus

End Sub

Private Sub Form_DblClick()
    'MsgBox "Sos 1 pelotudo."
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    Dim tX As Integer

    Dim tY As Integer

    If Not MapaCargado Then Exit Sub

    If SobreX > 0 And SobreY > 0 Then
        DobleClick Val(SobreX), Val(SobreY)

    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    ' HotKeys
    If HotKeysAllow = False Then Exit Sub

    Select Case UCase(Chr(KeyAscii))

        Case "S" ' Activa/Desactiva Insertar Superficie
            cSeleccionarSuperficie.Value = (cSeleccionarSuperficie.Value = False)
            Call cSeleccionarSuperficie_Click

        Case "T" ' Activa/Desactiva Insertar Translados
            cInsertarTrans.Value = (cInsertarTrans.Value = False)
            Call cInsertarTrans_Click

        Case "B" ' Activa/Desactiva Insertar Bloqueos
            cInsertarBloqueo.Value = (cInsertarBloqueo.Value = False)
            Call cInsertarBloqueo_Click

        Case "N" ' Activa/Desactiva Insertar NPCs
            cInsertarFunc(0).Value = (cInsertarFunc(0).Value = False)
            Call cInsertarFunc_Click(0)

            ' Case "H" ' Activa/Desactiva Insertar NPCs Hostiles
            '     cInsertarFunc(1).value = (cInsertarFunc(1).value = False)
            '     Call cInsertarFunc_Click(1)
        Case "O" ' Activa/Desactiva Insertar Objetos
            cInsertarFunc(2).Value = (cInsertarFunc(2).Value = False)
            Call cInsertarFunc_Click(2)

        Case "G" ' Activa/Desactiva Insertar Triggers
            cInsertarTrigger.Value = (cInsertarTrigger.Value = False)
            Call cInsertarTrigger_Click

        Case "Q" ' Quitar Funciones
            Call mnuQuitarFunciones_Click

    End Select

End Sub

Private Sub Form_Load()
    frmMain.Dialog.FilterIndex = 1
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'If Seleccionando Then CopiarSeleccion
End Sub

Private Sub lListado_Click(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    On Error Resume Next

    If HotKeysAllow = False Then
        lListado(index).Tag = lListado(index).ListIndex

        Select Case index

            Case 0
                cGrh.Text = DameGrhIndex(ReadField(2, lListado(index).Text, Asc("#")))

                If SupData(ReadField(2, lListado(index).Text, Asc("#"))).Capa <> 0 Then
                    If LenB(ReadField(2, lListado(index).Text, Asc("#"))) = 0 Then cCapas.Tag = cCapas.Text
                    cCapas.Text = SupData(ReadField(2, lListado(index).Text, Asc("#"))).Capa
                Else

                    If LenB(cCapas.Tag) <> 0 Then
                        cCapas.Text = cCapas.Tag
                        cCapas.Tag = vbNullString

                    End If

                End If

                If SupData(ReadField(2, lListado(index).Text, Asc("#"))).Block = True Then
                    If LenB(cInsertarBloqueo.Tag) = 0 Then cInsertarBloqueo.Tag = IIf(cInsertarBloqueo.Value = True, 1, 0)
                    cInsertarBloqueo.Value = True
                    Call cInsertarBloqueo_Click
                Else

                    If LenB(cInsertarBloqueo.Tag) <> 0 Then
                        cInsertarBloqueo.Value = IIf(Val(cInsertarBloqueo.Tag) = 1, True, False)
                        cInsertarBloqueo.Tag = vbNullString
                        Call cInsertarBloqueo_Click

                    End If

                End If

                Call fPreviewGrh(cGrh.Text)
                Call modPaneles.VistaPreviaDeSup

            Case 1
                cNumFunc(0).Text = ReadField(2, lListado(index).Text, Asc("#"))

            Case 2
                cNumFunc(1).Text = ReadField(2, lListado(index).Text, Asc("#"))

            Case 3
                cNumFunc(2).Text = ReadField(2, lListado(index).Text, Asc("#"))

        End Select

    Else
        lListado(index).ListIndex = lListado(index).Tag

    End If

End Sub

Private Sub lListado_MouseDown(index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    If index = 3 And Button = 2 Then
        If lListado(3).ListIndex > -1 Then Me.PopupMenu mnuObjSc

    End If

End Sub

Private Sub lListado_MouseMove(index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    On Error Resume Next

    HotKeysAllow = False

End Sub

Private Sub lvButtons_H1_Click()
    base_light = ARGB(230, 200, 200, 255)

End Sub

Private Sub lvButtons_H2_Click()
    base_light = ARGB(255, 255, 255, 255)

End Sub

Private Sub lvButtons_H3_Click()
    base_light = ARGB(200, 200, 200, 255)

End Sub

Private Sub lvButtons_H4_Click()
    base_light = ARGB(165, 165, 165, 255)

End Sub

' Lee los traslados del mapa y retorna los mapas adyacentes o cero si no tiene en esa direccion
Private Sub LeerAdyacentes(ByRef Norte As Integer, ByRef Sur As Integer, ByRef Este As Integer, ByRef Oeste As Integer)
    Dim X As Integer
    Dim Y As Integer

    ' Norte
    Y = MinYBorder
    For X = (MinXBorder + 1) To (MaxXBorder - 1)
        If MapData(X, Y).TileExit.Map > 0 Then
            Norte = MapData(X, Y).TileExit.Map
            Exit For
        End If
    Next

    ' Este
    X = MaxXBorder
    For Y = (MinYBorder + 1) To (MaxYBorder - 1)
        If MapData(X, Y).TileExit.Map > 0 Then
            Este = MapData(X, Y).TileExit.Map
            Exit For
        End If
    Next

    ' Sur
    Y = MaxYBorder
    For X = (MinXBorder + 1) To (MaxXBorder - 1)
        If MapData(X, Y).TileExit.Map > 0 Then
            Sur = MapData(X, Y).TileExit.Map
            Exit For
        End If
    Next

    ' Oeste
    X = MinXBorder
    For Y = (MinYBorder + 1) To (MaxYBorder - 1)
        If MapData(X, Y).TileExit.Map > 0 Then
            Oeste = MapData(X, Y).TileExit.Map
            Exit For
        End If
    Next
End Sub

Private Sub Magic_Click()
    Dim Path As String
    Path = InputBox("Ingrese el path absoluto a la carpeta de mapas", "Que fiaca hacer un formulario de abrir jaja")
    
    Dim Files() As String, File As String
    
    File = Dir$(Path & "\*.MAP")
    
    Dim Iterator As Integer
    
    Do While File <> ""
        ReDim Preserve Files(Iterator) As String
        Files(Iterator) = File
        Iterator = Iterator + 1
        File = Dir
    Loop
    
    Dim Norte As Integer, Sur As Integer, Este As Integer, Oeste As Integer

    For Iterator = 0 To UBound(Files)
        File = Path & "\" & Files(Iterator)
    
        Call modMapIO.NuevoMapa
        Call modMapIO.MapaV2_Cargar(File)
    
        Norte = 0
        Sur = 0
        Este = 0
        Oeste = 0
        
        Call LeerAdyacentes(Norte, Sur, Este, Oeste)
        
        Call LimpiarTraslados(Norte, Sur, Este, Oeste)
        
        Call AplicarTraslados(Norte, Sur, Este, Oeste)
        
        Call BloquearBordes
        
        Call modMapIO.MapaV2_Guardar(File, False)
    
    Next
        
    Call modMapIO.NuevoMapa
    Call modMapIO.MapaV2_Cargar(Dir$(Path & "\*.MAP"))
    
End Sub

Private Sub LimpiarTraslados(ByVal Norte As Integer, ByVal Sur As Integer, ByVal Este As Integer, ByVal Oeste As Integer)
    Dim Y As Integer
    Dim X As Integer

    ' Norte
    If Norte > 0 Then
        Y = MinYBorder

        For X = (MinXBorder + 1) To (MaxXBorder - 1)

            MapData(X, Y).TileExit.Map = 0

            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0

        Next

    End If

    ' Este
    If Este > 0 Then
        X = MaxXBorder

        For Y = (MinYBorder + 1) To (MaxYBorder - 1)

            MapData(X, Y).TileExit.Map = 0

            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0

        Next

    End If

    ' Sur
    If Sur > 0 Then
        Y = MaxYBorder

        For X = (MinXBorder + 1) To (MaxXBorder - 1)

            MapData(X, Y).TileExit.Map = 0

            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0

        Next

    End If

    ' Oeste
    If Oeste > 0 Then
        X = MinXBorder

        For Y = (MinYBorder + 1) To (MaxYBorder - 1)

            MapData(X, Y).TileExit.Map = 0

            MapData(X, Y).TileExit.X = 0
            MapData(X, Y).TileExit.Y = 0

        Next

    End If

End Sub

Private Sub AplicarTraslados(ByVal Norte As Integer, ByVal Sur As Integer, ByVal Este As Integer, ByVal Oeste As Integer)
    Dim Y As Integer
    Dim X As Integer

    ' Norte
    If Norte > 0 Then
        Y = NewMinYBorder

        For X = (NewMinXBorder + 1) To (NewMaxXBorder - 1)

            If MapData(X, Y).blocked = 0 Then
                MapData(X, Y).TileExit.Map = Norte
                MapData(X, Y).TileExit.X = X
                MapData(X, Y).TileExit.Y = NewMaxYBorder - 1
            End If

        Next
        
    Else
        Y = NewMinYBorder
        ' Si no tiene traslado para este lado, bloqueamos las posiciones
        For X = (NewMinXBorder + 1) To (NewMaxXBorder - 1)
            MapData(X, Y).blocked = 1
        Next
    End If

    ' Este
    If Este > 0 Then
        X = NewMaxXBorder

        For Y = (NewMinYBorder + 1) To (NewMaxYBorder - 1)

            If MapData(X, Y).blocked = 0 Then
                MapData(X, Y).TileExit.Map = Este
                MapData(X, Y).TileExit.X = NewMinXBorder + 1
                MapData(X, Y).TileExit.Y = Y
            End If

        Next
    
    Else
        ' Si no tiene traslado para este lado, bloqueamos las posiciones
        X = NewMaxXBorder
        For Y = (NewMinYBorder + 1) To (NewMaxYBorder - 1)
            MapData(X, Y).blocked = 1
        Next
    End If

    ' Sur
    If Sur > 0 Then
        Y = NewMaxYBorder

        For X = (NewMinXBorder + 1) To (NewMaxXBorder - 1)

            If MapData(X, Y).blocked = 0 Then
                MapData(X, Y).TileExit.Map = Sur
                MapData(X, Y).TileExit.X = X
                MapData(X, Y).TileExit.Y = NewMinYBorder + 1
            End If

        Next
        
    Else
        ' Si no tiene traslado para este lado, bloqueamos las posiciones
        Y = NewMaxYBorder
        For X = (NewMinXBorder + 1) To (NewMaxXBorder - 1)
            MapData(X, Y).blocked = 1
        Next

    End If

    ' Oeste
    If Oeste > 0 Then
        X = NewMinXBorder

        For Y = (NewMinYBorder + 1) To (NewMaxYBorder - 1)

            If MapData(X, Y).blocked = 0 Then
                MapData(X, Y).TileExit.Map = Oeste
                MapData(X, Y).TileExit.X = NewMaxXBorder - 1
                MapData(X, Y).TileExit.Y = Y
            End If

        Next
        
    Else
        ' Si no tiene traslado para este lado, bloqueamos las posiciones
        X = NewMinXBorder
        For Y = (NewMinYBorder + 1) To (NewMaxYBorder - 1)
            MapData(X, Y).blocked = 1
        Next

    End If

End Sub

Private Sub BloquearBordes()
    Dim Y As Integer
    Dim X As Integer

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            If X < NewMinXBorder Or X > NewMaxXBorder Or Y < NewMinYBorder Or Y > NewMaxYBorder Then
                MapData(X, Y).blocked = 1
            End If

        Next X
    Next Y
    
    ' Bloqueo las 4 esquinitas que queda feo sino :v
    MapData(NewMinXBorder, NewMinYBorder).blocked = 1
    MapData(NewMaxXBorder, NewMinYBorder).blocked = 1
    MapData(NewMinXBorder, NewMaxYBorder).blocked = 1
    MapData(NewMaxXBorder, NewMaxYBorder).blocked = 1
End Sub

Private Sub MapPest_Click(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim formato As String

    
    Select Case frmMain.Dialog.FilterIndex
    
        Case 0
            formato = ".csm"
            
        Case 1
            formato = ".map"
            
    End Select
    
    If (index + NumMap_Save - 4) <> NumMap_Save Then
        Dialog.CancelError = True

        On Error GoTo ErrHandler

        Dialog.FileName = PATH_Save & NameMap_Save & (index + NumMap_Save - 7) & formato

        If MapInfo.Changed = 1 Then
            
            If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
                Call modMapIO.GuardarMapa(Dialog.FileName)
            End If

        End If

        Call modMapIO.NuevoMapa
        
        DoEvents
        
        Select Case frmMain.Dialog.FilterIndex
        
            Case 0
                Call modMapIO.Cargar_CSM(Dialog.FileName)
                
            Case 1
                Call modMapIO.MapaV2_Cargar(Dialog.FileName)
            
        End Select
        
        EngineRun = True
        
    End If
    
        Exit Sub
    
ErrHandler:
        Call MsgBox(Err.Description)

End Sub

Private Sub MemoriaAuxiliar_Click()

    On Error GoTo Error
 
    MapData_Adyacente = MapData
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa copiado a la memoria", 0, 255, 0)
     
    Exit Sub
    
Error:
    Call AddtoRichTextBox(frmMain.StatTxt, "Error guardando mapa", 255, 0, 0)

End Sub

Private Sub minimap_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

    If X < 11 Then X = 11
    If X > 89 Then X = 89
    If Y < 10 Then Y = 10
    If Y > 92 Then Y = 92
    
    UserPos.X = X
    UserPos.Y = Y
    
    Call ActualizaMinimap

End Sub

Private Sub mnuAbrirMapaNew_Click()
    '*************************************************
    'Author: Lorwik
    'Last modified: 04/11/2015
    '*************************************************
    Dialog.CancelError = True

    On Error GoTo ErrHandler

    Call DeseaGuardarMapa(Dialog.FileName)

    Call ObtenerNombreArchivo(False)

    If Len(Dialog.FileName) < 3 Then Exit Sub

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode

    End If
    
    Call modMapIO.NuevoMapa

    Select Case frmMain.Dialog.FilterIndex
    
        Case 1
            Call modMapIO.MapaV2_Cargar(Dialog.FileName)
            
        Case 2
            Call modMapIO.Cargar_CSM(Dialog.FileName)
            
    End Select
    
    DoEvents
    
    mnuReAbrirMapa.Enabled = True
    
    EngineRun = True
    
    Exit Sub
ErrHandler:

End Sub

Private Sub mnuActualizarCabezas_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    Call modIndices.CargarIndicesDeCabezas

End Sub

Private Sub mnuActualizarCuerpos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    Call modIndices.CargarIndicesDeCuerpos

End Sub

Private Sub mnuActualizarGraficos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    Call modIndices.LoadGrhData

End Sub

Private Sub mnuActualizarSuperficies_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modIndices.CargarIndicesSuperficie

End Sub

Private Sub mnuacercade_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmAbout.Show

End Sub

Private Sub mnuActualizarNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modIndices.CargarIndicesNPC

End Sub

Private Sub mnuActualizarObjs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modIndices.CargarIndicesOBJ

End Sub

Private Sub mnuActualizarTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modIndices.CargarIndicesTriggers

End Sub

Private Sub mnuAutoCapturarTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    mnuAutoCapturarTranslados.Checked = (mnuAutoCapturarTranslados.Checked = False)

End Sub

Private Sub mnuAutoCapturarSuperficie_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    mnuAutoCapturarSuperficie.Checked = (mnuAutoCapturarSuperficie.Checked = False)

End Sub

Private Sub mnuAutoCompletarSuperficies_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuAutoCompletarSuperficies.Checked = (mnuAutoCompletarSuperficies.Checked = False)

End Sub

Private Sub mnuAutoGuardarMapas_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmAutoGuardarMapa.Show

End Sub

Private Sub mnuAutoQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuAutoQuitarFunciones.Checked = (mnuAutoQuitarFunciones.Checked = False)

End Sub

Private Sub mnuBloquear_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 6

        If i <> 2 Then
            frmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next

    modPaneles.VerFuncion 2, True

End Sub

Private Sub mnuBloquearBordes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Bloquear_Bordes

End Sub

Private Sub mnuBloquearMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Bloqueo_Todo(1)

End Sub

Private Sub mnuBloquearS_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call modEdicion.Deshacer_Add("Bloquear Selección")
    Call BlockearSeleccion

End Sub

Private Sub mnuConfigAvanzada_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmConfigSup.Show

End Sub

Private Sub mnuConfigObjTrans_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 29/05/06
    '*************************************************
    Cfg_TrOBJ = cNumFunc(2).Text

End Sub

Private Sub mnuConvert_Click()
    frmConvert.Show

End Sub

Private Sub mnuCopiar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call CopiarSeleccion

End Sub

Private Sub mnuCortar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call modEdicion.Deshacer_Add("Cortar Selección")
    Call CortarSeleccion

End Sub

Private Sub mnuDeshacer_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 15/10/06
    '*************************************************
    Call modEdicion.Deshacer_Recover

End Sub

Private Sub mnuDeshacerPegado_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call modEdicion.Deshacer_Add("Deshacer Pegado de Selección")
    Call DePegar

End Sub

Private Sub mnuGRHaBMP_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call frmGRHaBMP.Show

End Sub

Private Sub mnuGuardarMapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modMapIO.GuardarMapa(Dialog.FileName)

End Sub

Private Sub mnuGuardarMapaComo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modMapIO.GuardarMapa

End Sub

Private Sub mnuGuardarUltimaConfig_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 23/05/06
    '*************************************************
    mnuGuardarUltimaConfig.Checked = (mnuGuardarUltimaConfig.Checked = False)

End Sub

Private Sub mnuInfoMap_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmMapInfo.Show
    frmMapInfo.Visible = True

End Sub

Private Sub mnuInformes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call frmInformes.Show

End Sub

Private Sub mnuInsertarSuperficieAlAzar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Superficie_Azar

End Sub

Private Sub mnuInsertarSuperficieEnBordes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Superficie_Bordes

End Sub

Private Sub mnuInsertarSuperficieEnTodo_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Superficie_Todo

End Sub

Private Sub mnuInsertarTransladosAdyasentes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    frmUnionAdyacente.Show

End Sub

Private Sub mnuManual_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    If LenB(Dir(App.Path & "\manual\index.html", vbArchive)) <> 0 Then
        Call Shell("explorer " & App.Path & "\manual\index.html")
        DoEvents

    End If

End Sub

Private Sub mnuModoCaminata_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    Call ToggleWalkMode

End Sub

Private Sub mnuNPCs_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 6

        If i <> 3 Then
            frmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    
    Call modPaneles.VerFuncion(3, True)

End Sub

'Private Sub mnuNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Dim i As Byte
'For i = 0 To 6
'    If i <> 4 Then
'        frmMain.SelectPanel(i).value = False
'        Call VerFuncion(i, False)
'    End If
'Next
'modPaneles.VerFuncion 4, True
'End Sub

Private Sub mnuNuevoMapa_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    Dim loopC As Integer

    Call DeseaGuardarMapa(Dialog.FileName)

    For loopC = 0 To frmMain.MapPest.count
        frmMain.MapPest(loopC).Visible = False
    Next

    frmMain.Dialog.FileName = Empty

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode

    End If

    Call modMapIO.NuevoMapa

    Call cmdInformacionDelMapa_Click

End Sub

Private Sub mnuObjetos_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 6

        If i <> 5 Then
            frmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    
    Call modPaneles.VerFuncion(5, True)

End Sub

Private Sub mnuOptimizar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/09/06
    '*************************************************
    frmOptimizar.Show

End Sub

Private Sub mnuPegar_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call modEdicion.Deshacer_Add("Pegar Selección")
    Call PegarSeleccion

End Sub

Private Sub mnuQBloquear_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modPaneles.VerFuncion(2, False)

End Sub

Private Sub mnuQNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modPaneles.VerFuncion(3, False)

End Sub

'Private Sub mnuQNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'modPaneles.VerFuncion 4, False
'End Sub

Private Sub mnuQObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modPaneles.VerFuncion(5, False)

End Sub

Private Sub mnuQSuperficie_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modPaneles.VerFuncion(0, False)

End Sub

Private Sub mnuQTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modPaneles.VerFuncion(1, False)

End Sub

Private Sub mnuQTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modPaneles.VerFuncion(6, False)

End Sub

Private Sub mnuQuitarBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Bloqueo_Todo(0)

End Sub

Private Sub mnuQuitarFunciones_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    
    ' Superficies
    cSeleccionarSuperficie.Value = False
    Call cSeleccionarSuperficie_Click
    cQuitarEnEstaCapa.Value = False
    Call cQuitarEnEstaCapa_Click
    cQuitarEnTodasLasCapas.Value = False
    Call cQuitarEnTodasLasCapas_Click
    
    ' Translados
    cQuitarTrans.Value = False
    Call cQuitarTrans_Click
    cInsertarTrans.Value = False
    Call cInsertarTrans_Click
    
    ' Bloqueos
    cQuitarBloqueo.Value = False
    Call cQuitarBloqueo_Click
    cInsertarBloqueo.Value = False
    Call cInsertarBloqueo_Click
    
    ' Otras funciones
    cInsertarFunc(0).Value = False
    Call cInsertarFunc_Click(0)
    cInsertarFunc(1).Value = False
    Call cInsertarFunc_Click(1)
    cInsertarFunc(2).Value = False
    Call cInsertarFunc_Click(2)
    cQuitarFunc(0).Value = False
    Call cQuitarFunc_Click(0)
    cQuitarFunc(1).Value = False
    Call cQuitarFunc_Click(1)
    cQuitarFunc(2).Value = False
    Call cQuitarFunc_Click(2)
    
    ' Triggers
    cInsertarTrigger.Value = False
    Call cInsertarTrigger_Click
    cQuitarTrigger.Value = False
    Call cQuitarTrigger_Click

End Sub

Private Sub mnuQuitarNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Quitar_NPCs(False)

End Sub

Private Sub mnuQuitarNPCsHostiles_Click()
    Call modEdicion.Quitar_NPCs(True)

End Sub

'Private Sub mnuQuitarNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Call modEdicion.Quitar_NPCs(True)
'End Sub

Private Sub mnuQuitarObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Quitar_Objetos

End Sub

Private Sub mnuQuitarSuperficieBordes_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Quitar_Bordes

End Sub

Private Sub mnuQuitarSuperficieDeCapa_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Quitar_Capa(cCapas.Text)

End Sub

Private Sub mnuQuitarTODO_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Borrar_Mapa

End Sub

Private Sub mnuQuitarTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    Call modEdicion.Quitar_Translados

End Sub

Private Sub mnuQuitarTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Call modEdicion.Quitar_Triggers

End Sub

Private Sub mnuReAbrirMapa_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    On Error GoTo ErrHandler

    If FileExist(Dialog.FileName, vbArchive) = False Then Exit Sub
    
    If MapInfo.Changed = 1 Then
        
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            Call modMapIO.GuardarMapa(Dialog.FileName)
        End If

    End If

    Call modMapIO.NuevoMapa

    Select Case frmMain.Dialog.FilterIndex
    
        Case 0
            Call modMapIO.Cargar_CSM(Dialog.FileName)
            
        Case 1
            Call modMapIO.MapaV2_Cargar(Dialog.FileName)
            
    End Select
    
    DoEvents
    
    mnuReAbrirMapa.Enabled = True
    
    EngineRun = True
    
    Exit Sub
    
ErrHandler:

End Sub

Private Sub mnuRealizarOperacion_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************
    Call modEdicion.Deshacer_Add("Realizar Operación en Selección")
    Call AccionSeleccion

End Sub

Private Sub mnuRenderMapa_Click()
    Radio = Val(InputBox("Escriba la escala de 1 a 5 en la que generemos su mapa", "la escala se multiplica x 32"))

    If Radio = 0 Then Radio = 1
    If Radio >= 5 Then Radio = 5

    frmRender.picMap.Width = (Radio * 100)
    frmRender.picMap.Height = (Radio * 100)

    frmRender.Show
End Sub

Private Sub mnuSalir_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Unload Me

End Sub

Private Sub mnuSuperficie_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 6

        If i <> 0 Then
            frmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    
    Call modPaneles.VerFuncion(0, True)

End Sub

Private Sub mnuTranslados_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 6

        If i <> 1 Then
            frmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    
    Call modPaneles.VerFuncion(1, True)

End Sub

Private Sub mnuTriggers_Click()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 6

        If i <> 6 Then
            frmMain.SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next
    
    Call modPaneles.VerFuncion(6, True)

End Sub

Private Sub mnuUtilizarDeshacer_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 16/10/06
    '*************************************************
    mnuUtilizarDeshacer.Checked = (mnuUtilizarDeshacer.Checked = False)

End Sub

Private Sub mnuVerAutomatico_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuVerAutomatico.Checked = (mnuVerAutomatico.Checked = False)

End Sub

Private Sub mnuVerBloqueos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    cVerBloqueos.Value = (cVerBloqueos.Value = False)
    mnuVerBloqueos.Checked = cVerBloqueos.Value

End Sub

Private Sub mnuVerCapa1_Click()
    mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)

End Sub

Private Sub mnuVerCapa2_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)

End Sub

Private Sub mnuVerCapa3_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)

End Sub

Private Sub mnuVerCapa4_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)

End Sub

Private Sub mnuVerGrilla_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    VerGrilla = (VerGrilla = False)
    mnuVerGrilla.Checked = VerGrilla

End Sub

Private Sub mnuVerNPCs_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    mnuVerNPCs.Checked = (mnuVerNPCs.Checked = False)

End Sub

Private Sub mnuVerObjetos_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)

End Sub

Private Sub mnuVerTranslados_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 26/05/06
    '*************************************************
    mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)

End Sub

Private Sub mnuVerTriggers_Click()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    cVerTriggers.Value = (cVerTriggers.Value = False)
    mnuVerTriggers.Checked = cVerTriggers.Value

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06 - GS
    'Last modified: 20/11/07 - Loopzer
    '*************************************************

    Dim tX As Integer

    Dim tY As Integer

    If Not MapaCargado Then Exit Sub

    ConvertCPtoTP X, Y, tX, tY

    'If Shift = 1 And Button = 2 Then PegarSeleccion tX, tY: Exit Sub
    If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionIX = tX '+ UserPos.X
        SeleccionIY = tY '+ UserPos.Y
    Else
        Call ClickEdit(Button, tX, tY)
    End If

End Sub

Private Sub Renderer_DblClick()

'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    Dim tX As Integer

    Dim tY As Integer

    If Not MapaCargado Then Exit Sub

    If SobreX > 0 And SobreY > 0 Then
        DobleClick Val(SobreX), Val(SobreY)
    End If
    
End Sub

Private Sub Renderer_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    Call Form_MouseMove(Button, Shift, X, Y)
    MouseX = X
    MouseY = Y

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06 - GS
    '*************************************************

    Dim tX As Integer

    Dim tY As Integer

    'Make sure map is loaded
    If Not MapaCargado Then Exit Sub
    HotKeysAllow = True

    ConvertCPtoTP X, Y, tX, tY

    POSX = "X: " & tX & " - Y: " & tY

    If Shift = 1 And Button = 1 Then
        Seleccionando = True
        SeleccionFX = tX '+ TileX
        SeleccionFY = tY '+ TileY
    Else
        Call ClickEdit(Button, tX, tY)
    End If

End Sub

Private Sub Renderer_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    Call Form_MouseDown(Button, Shift, X, Y)
    Call DibujarMiniMapa

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    
    Dim IniManager As clsIniManager
    Set IniManager = New clsIniManager
    
    ' Guardar configuración
    Call IniManager.ChangeValue("CONFIGURACION", "GuardarConfig", IIf(frmMain.mnuGuardarUltimaConfig.Checked = True, "1", "0"))

    If frmMain.mnuGuardarUltimaConfig.Checked = True Then
    
        Call IniManager.ChangeValue("PATH", "UltimoMapa", Dialog.FileName)
        
        Call IniManager.ChangeValue("MOSTRAR", "ControlAutomatico", IIf(frmMain.mnuVerAutomatico.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("MOSTRAR", "Capa2", IIf(frmMain.mnuVerCapa2.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("MOSTRAR", "Capa3", IIf(frmMain.mnuVerCapa3.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("MOSTRAR", "Capa4", IIf(frmMain.mnuVerCapa4.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("MOSTRAR", "Translados", IIf(frmMain.mnuVerTranslados.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("MOSTRAR", "Objetos", IIf(frmMain.mnuVerObjetos.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("MOSTRAR", "NPCs", IIf(frmMain.mnuVerNPCs.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("MOSTRAR", "Triggers", IIf(frmMain.mnuVerTriggers.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("MOSTRAR", "Grilla", IIf(frmMain.mnuVerGrilla.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("MOSTRAR", "Bloqueos", IIf(frmMain.mnuVerBloqueos.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("MOSTRAR", "LastPos", UserPos.X & "-" & UserPos.Y)
        
        Call IniManager.ChangeValue("CONFIGURACION", "UtilizarDeshacer", IIf(frmMain.mnuUtilizarDeshacer.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("CONFIGURACION", "AutoCapturarTrans", IIf(frmMain.mnuAutoCapturarTranslados.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("CONFIGURACION", "AutoCapturarSup", IIf(frmMain.mnuAutoCapturarSuperficie.Checked = True, "1", "0"))
        Call IniManager.ChangeValue("CONFIGURACION", "ObjTranslado", Val(Cfg_TrOBJ))

    End If
    
    Call IniManager.DumpFile(IniPath & "WorldEditor.ini")
    
    'Allow MainLoop to close program
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If

End Sub

Private Sub SelectPanel_Click(index As Integer)

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Byte

    For i = 0 To 9

        If i <> index Then
            SelectPanel(i).Value = False
            Call VerFuncion(i, False)

        End If

    Next

    If mnuAutoQuitarFunciones.Checked = True Then Call mnuQuitarFunciones_Click
    Call VerFuncion(index, SelectPanel(index).Value)

End Sub

Private Sub TimAutoGuardarMapa_Timer()

    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If mnuAutoGuardarMapas.Checked = True Then
        bAutoGuardarMapaCount = bAutoGuardarMapaCount + 1

        If bAutoGuardarMapaCount >= bAutoGuardarMapa Then
            If MapInfo.Changed = 1 Then ' Solo guardo si el mapa esta modificado
                modMapIO.GuardarMapa Dialog.FileName

            End If

            bAutoGuardarMapaCount = 0

        End If

    End If

End Sub

Public Sub ObtenerNombreArchivo(ByVal Guardar As Boolean)

    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    On Error Resume Next

    With Dialog
        .Filter = "Mapas de Argentum Online (*.map)|*.map|Mapas de IAO Clon (*.csm)|*.csm"

        If Guardar Then
            .DialogTitle = "Guardar"
            .DefaultExt = ".txt"
            .FileName = vbNullString
            .flags = cdlOFNPathMustExist
            .ShowSave
        Else
            .DialogTitle = "Cargar"
            .FileName = vbNullString
            .flags = cdlOFNFileMustExist
            .ShowOpen

        End If

    End With

End Sub
