VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WorldEditor Versión 2.0.56  ReyarB"
   ClientHeight    =   11940
   ClientLeft      =   3270
   ClientTop       =   690
   ClientWidth     =   21915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   796
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1461
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WorldEditor.lvButtons_H SelectPanelextra 
      Height          =   1275
      Index           =   2
      Left            =   24840
      TabIndex        =   210
      Top             =   0
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   2249
      Caption         =   "Abrir Datos en Exel"
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
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanelextra 
      Height          =   1275
      Index           =   1
      Left            =   23160
      TabIndex        =   209
      Top             =   0
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   2249
      Caption         =   "&Translados Mapa Adtacentes"
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
      Image           =   "frmMain.frx":7F6A
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin RichTextLib.RichTextBox StatTxt 
      Height          =   960
      Left            =   4680
      TabIndex        =   88
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   10920
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   1693
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmMain.frx":B5CA
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
   Begin VB.PictureBox PreviewNPCs 
      BackColor       =   &H00000000&
      FillColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   120
      ScaleHeight     =   2160
      ScaleWidth      =   4425
      TabIndex        =   205
      Top             =   9600
      Visible         =   0   'False
      Width           =   4485
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   207
         Top             =   555
         Width           =   525
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Largo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   206
         Top             =   900
         Width           =   480
      End
   End
   Begin VB.PictureBox PreviewObj 
      BackColor       =   &H00000000&
      FillColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   120
      ScaleHeight     =   2160
      ScaleWidth      =   4425
      TabIndex        =   202
      Top             =   9600
      Visible         =   0   'False
      Width           =   4485
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Largo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   203
         Top             =   900
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   204
         Top             =   555
         Width           =   525
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   20520
      Picture         =   "frmMain.frx":B647
      ScaleHeight     =   495
      ScaleWidth      =   1335
      TabIndex        =   201
      Top             =   10800
      Width           =   1335
   End
   Begin VB.PictureBox pPaneles 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   120
      Picture         =   "frmMain.frx":1013D
      ScaleHeight     =   415
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   295
      TabIndex        =   2
      Top             =   3360
      Width           =   4455
      Begin VB.ListBox lstParticle 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2205
         Left            =   120
         TabIndex        =   115
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
         Height          =   2160
         Index           =   0
         ItemData        =   "frmMain.frx":6A109
         Left            =   120
         List            =   "frmMain.frx":6A10B
         Sorted          =   -1  'True
         TabIndex        =   61
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Frame MapasFrame 
         BackColor       =   &H80000012&
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   120
         TabIndex        =   176
         Top             =   5520
         Width           =   4215
         Begin VB.OptionButton Option2 
            BackColor       =   &H00000000&
            Caption         =   "Mapas para 680"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   2400
            TabIndex        =   178
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00000000&
            Caption         =   "Mapas para 1024"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   240
            TabIndex        =   177
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.TextBox Life 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   116
         Text            =   "-1"
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
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
         ItemData        =   "frmMain.frx":6A10D
         Left            =   1080
         List            =   "frmMain.frx":6A11D
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   2640
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
         Left            =   2760
         TabIndex        =   63
         Text            =   "1"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
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
         Top             =   2280
         Visible         =   0   'False
         Width           =   3615
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
         ItemData        =   "frmMain.frx":6A12D
         Left            =   3360
         List            =   "frmMain.frx":6A12F
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
      Begin VB.PictureBox Picture5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   3
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   4
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   5
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   6
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   7
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         ItemData        =   "frmMain.frx":6A131
         Left            =   840
         List            =   "frmMain.frx":6A133
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
         ItemData        =   "frmMain.frx":6A135
         Left            =   3360
         List            =   "frmMain.frx":6A137
         TabIndex        =   70
         Text            =   "500"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin WorldEditor.lvButtons_H cmdAdd 
         Height          =   375
         Left            =   1320
         TabIndex        =   118
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
         TabIndex        =   119
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
         Height          =   2370
         Index           =   3
         ItemData        =   "frmMain.frx":6A139
         Left            =   120
         List            =   "frmMain.frx":6A13B
         TabIndex        =   56
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
         Height          =   2370
         Index           =   1
         ItemData        =   "frmMain.frx":6A13D
         Left            =   120
         List            =   "frmMain.frx":6A13F
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
         Height          =   2370
         Index           =   4
         ItemData        =   "frmMain.frx":6A141
         Left            =   120
         List            =   "frmMain.frx":6A143
         TabIndex        =   44
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
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
         ItemData        =   "frmMain.frx":6A145
         Left            =   840
         List            =   "frmMain.frx":6A147
         TabIndex        =   0
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
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
         ItemData        =   "frmMain.frx":6A149
         Left            =   840
         List            =   "frmMain.frx":6A14B
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
         Index           =   2
         ItemData        =   "frmMain.frx":6A14D
         Left            =   3360
         List            =   "frmMain.frx":6A14F
         TabIndex        =   57
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin WorldEditor.lvButtons_H cQuitarEnTodasLasCapas 
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   3480
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
      Begin WorldEditor.lvButtons_H cSeleccionarSuperficie 
         Height          =   855
         Left            =   2400
         TabIndex        =   66
         Top             =   3000
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1508
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
      Begin WorldEditor.lvButtons_H cQuitarEnEstaCapa 
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   3000
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
         Height          =   2370
         Index           =   2
         ItemData        =   "frmMain.frx":6A151
         Left            =   120
         List            =   "frmMain.frx":6A153
         TabIndex        =   69
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
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
      Begin VB.Frame CopyBorder 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   5415
         Left            =   120
         TabIndex        =   99
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdIrAlMapa 
            Caption         =   "Ir al Mapa"
            Height          =   375
            Left            =   2520
            TabIndex        =   199
            Top             =   4920
            Width           =   1335
         End
         Begin VB.TextBox TxtMapa 
            Height          =   285
            Left            =   1440
            TabIndex        =   197
            Text            =   "1"
            Top             =   4920
            Width           =   735
         End
         Begin WorldEditor.lvButtons_H COPIAR_GRH 
            Height          =   1455
            Index           =   2
            Left            =   3000
            TabIndex        =   106
            Top             =   960
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   2566
            Caption         =   "Pegar mapa Este"
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
            Index           =   8
            Left            =   2520
            TabIndex        =   192
            Top             =   3600
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            Caption         =   "3"
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
            Index           =   9
            Left            =   2520
            TabIndex        =   193
            Top             =   3960
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            Caption         =   "6"
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
            Index           =   10
            Left            =   1320
            TabIndex        =   194
            Top             =   4320
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            Caption         =   "7"
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
            Index           =   11
            Left            =   1920
            TabIndex        =   195
            Top             =   4320
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            Caption         =   "8"
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
            Index           =   12
            Left            =   2520
            TabIndex        =   196
            Top             =   4320
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            Caption         =   "9"
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
            Index           =   5
            Left            =   1920
            TabIndex        =   189
            Top             =   3600
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            Caption         =   "2"
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
            Index           =   6
            Left            =   1320
            TabIndex        =   190
            Top             =   3960
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            Caption         =   "3"
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
            Index           =   7
            Left            =   1920
            TabIndex        =   191
            Top             =   3960
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            Caption         =   "4"
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
            Index           =   4
            Left            =   1320
            TabIndex        =   188
            Top             =   3600
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            Caption         =   "1"
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
            Height          =   1455
            Index           =   1
            Left            =   120
            TabIndex        =   105
            Top             =   960
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   2566
            Caption         =   "Pegar mapa Oeste"
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
         Begin VB.OptionButton OpcExit 
            BackColor       =   &H80000007&
            Caption         =   "Espejar hasta el exit "
            ForeColor       =   &H8000000B&
            Height          =   375
            Left            =   1200
            TabIndex        =   175
            Top             =   1320
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton OpcBorde 
            BackColor       =   &H80000012&
            Caption         =   "Espejar con el borde"
            ForeColor       =   &H8000000B&
            Height          =   375
            Left            =   1200
            TabIndex        =   174
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox TXTArriba 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1680
            TabIndex        =   103
            Text            =   "180"
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox TxTAbajo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1680
            TabIndex        =   102
            Text            =   "180"
            Top             =   2760
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox TxTDerecha 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3240
            TabIndex        =   101
            Text            =   "180"
            Top             =   1440
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox TxTIzquierda 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   360
            TabIndex        =   100
            Text            =   "180"
            Top             =   1440
            Visible         =   0   'False
            Width           =   615
         End
         Begin WorldEditor.lvButtons_H COPIAR_GRH 
            Height          =   735
            Index           =   3
            Left            =   120
            TabIndex        =   104
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   1296
            Caption         =   "Pegar en mapa Norte"
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
            Height          =   735
            Index           =   0
            Left            =   120
            TabIndex        =   107
            Top             =   2400
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   1296
            Caption         =   "Pegar mapa Sur"
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
            Height          =   2895
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   5106
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
         Begin VB.Label lblIrMapa 
            BackColor       =   &H80000007&
            Caption         =   "Ir al Mapa"
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   360
            TabIndex        =   198
            Top             =   4920
            Width           =   855
         End
         Begin VB.Label lvlMapaCompleto 
            BackStyle       =   0  'Transparent
            Caption         =   "Pegar mapa completo en Zona Nº"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   555
            Left            =   240
            TabIndex        =   110
            Top             =   3240
            Width           =   1770
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
            Index           =   0
            Left            =   2400
            TabIndex        =   109
            Top             =   3240
            Visible         =   0   'False
            Width           =   1065
         End
      End
      Begin VB.Frame cLuces 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Luces"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   4155
         Left            =   120
         TabIndex        =   122
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Frame Frame4 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   2520
            TabIndex        =   181
            Top             =   240
            Width           =   1455
            Begin WorldEditor.lvButtons_H lvButtons_H5 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   182
               Top             =   360
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               CapAlign        =   2
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
               cBack           =   255
            End
            Begin WorldEditor.lvButtons_H lvButtons_H5 
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   183
               Top             =   360
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               CapAlign        =   2
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
               cBack           =   65535
            End
            Begin WorldEditor.lvButtons_H lvButtons_H5 
               Height          =   255
               Index           =   2
               Left            =   720
               TabIndex        =   184
               Top             =   1080
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               CapAlign        =   2
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
               cBack           =   12632256
            End
            Begin WorldEditor.lvButtons_H lvButtons_H5 
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   185
               Top             =   1080
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               CapAlign        =   2
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
               cBack           =   16711935
            End
            Begin WorldEditor.lvButtons_H lvButtons_H5 
               Height          =   255
               Index           =   4
               Left            =   720
               TabIndex        =   186
               Top             =   720
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               CapAlign        =   2
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
               cBack           =   16777215
            End
            Begin WorldEditor.lvButtons_H lvButtons_H5 
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   187
               Top             =   720
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               CapAlign        =   2
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
               cBack           =   16776960
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00000000&
            Caption         =   "Rango"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   660
            Left            =   600
            TabIndex        =   127
            Top             =   1080
            Width           =   1380
            Begin VB.TextBox cRango 
               Appearance      =   0  'Flat
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
               TabIndex        =   128
               Text            =   "5"
               Top             =   240
               Width           =   555
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "(1 al 50)"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   720
               TabIndex        =   129
               Top             =   270
               Width           =   615
            End
         End
         Begin VB.Frame RGBCOLOR 
            BackColor       =   &H00000000&
            Caption         =   "RGB"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   690
            Left            =   600
            TabIndex        =   123
            Top             =   360
            Width           =   1680
            Begin VB.TextBox R 
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
               Height          =   315
               Left            =   105
               TabIndex        =   126
               Text            =   "200"
               Top             =   270
               Width           =   450
            End
            Begin VB.TextBox B 
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
               Height          =   315
               Left            =   1095
               TabIndex        =   125
               Text            =   "14"
               Top             =   270
               Width           =   450
            End
            Begin VB.TextBox G 
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
               Height          =   315
               Left            =   600
               TabIndex        =   124
               Text            =   "235"
               Top             =   270
               Width           =   450
            End
         End
         Begin WorldEditor.lvButtons_H cInsertarLuz 
            Height          =   360
            Left            =   2160
            TabIndex        =   135
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
            TabIndex        =   136
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
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Luz Base"
            ForeColor       =   &H00FFFFFF&
            Height          =   1335
            Left            =   120
            TabIndex        =   130
            Top             =   2760
            Width           =   3855
            Begin WorldEditor.lvButtons_H lvButtons_H1 
               Height          =   360
               Left            =   360
               TabIndex        =   131
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
               TabIndex        =   132
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
               TabIndex        =   133
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
               TabIndex        =   134
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
            TabIndex        =   137
            Top             =   2280
            Width           =   3615
         End
      End
      Begin VB.Frame FraRellenar 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Rellenar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   120
         TabIndex        =   150
         Top             =   3960
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox DY2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3480
            TabIndex        =   155
            Text            =   "5"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox DY1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   154
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox DX2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   153
            Text            =   "5"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox DX1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   600
            TabIndex        =   152
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin WorldEditor.lvButtons_H LvBAreas 
            Height          =   375
            Index           =   2
            Left            =   2280
            TabIndex        =   151
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Pintar Area"
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
         Begin WorldEditor.lvButtons_H LvBAreas 
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   165
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Quitar Bloqueos"
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
         Begin WorldEditor.lvButtons_H LvBAreas 
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   166
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Insertar Bloqueos"
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
         Begin WorldEditor.lvButtons_H LvBAreas 
            Height          =   375
            Index           =   3
            Left            =   2280
            TabIndex        =   167
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Quitar Area"
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
         Begin VB.Label lblX2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X1:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   159
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblX2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X2:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   1275
            TabIndex        =   158
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblY1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y1:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2160
            TabIndex        =   157
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblY2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y2:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   3195
            TabIndex        =   156
            Top             =   240
            Width           =   255
         End
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
         Top             =   2700
         Visible         =   0   'False
         Width           =   930
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
         TabIndex        =   117
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
         Top             =   2325
         Visible         =   0   'False
         Width           =   450
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
         Top             =   2700
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
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   19560
      Picture         =   "frmMain.frx":6A155
      ScaleHeight     =   900
      ScaleWidth      =   855
      TabIndex        =   173
      TabStop         =   0   'False
      Top             =   10920
      Width           =   855
   End
   Begin VB.Frame FraFormatoDel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Formato del Mapa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   17160
      TabIndex        =   170
      Top             =   10920
      Width           =   2295
      Begin VB.OptionButton OptX 
         BackColor       =   &H00FFFFFF&
         Caption         =   "300 x 300 (Maximo)"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   200
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton OptX 
         BackColor       =   &H00FFFFFF&
         Caption         =   "200 x 200 (agrandado)"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   172
         Top             =   480
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton OptX 
         BackColor       =   &H00FFFFFF&
         Caption         =   "100 x 100 (clasico)"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   171
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.PictureBox MinimapCapture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   22320
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   169
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   3023
   End
   Begin VB.CheckBox chkOptMinimap 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bloq"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   164
      Top             =   3075
      Width           =   735
   End
   Begin WorldEditor.lvButtons_H LvBVerMapa 
      Height          =   615
      Left            =   20520
      TabIndex        =   163
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      Caption         =   "Ver Mapa"
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
   Begin VB.CheckBox chkOptMinimap 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Capa 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   162
      Top             =   3075
      Width           =   975
   End
   Begin VB.CheckBox chkOptMinimap 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Obj"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   161
      Top             =   3075
      Width           =   975
   End
   Begin VB.CheckBox chkOptMinimap 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NPC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   160
      Top             =   3075
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.Frame FraOpciones 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3240
      TabIndex        =   140
      Top             =   0
      Width           =   1335
      Begin VB.CommandButton cmdDM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   600
         Picture         =   "frmMain.frx":6CBC9
         Style           =   1  'Graphical
         TabIndex        =   215
         Top             =   240
         Width           =   240
      End
      Begin VB.CommandButton cmdDM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   360
         Picture         =   "frmMain.frx":6CEBB
         Style           =   1  'Graphical
         TabIndex        =   214
         Top             =   480
         Width           =   240
      End
      Begin VB.CommandButton cmdDM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "frmMain.frx":6D1AB
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   480
         Width           =   240
      End
      Begin VB.CommandButton cmdDM 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   600
         Picture         =   "frmMain.frx":6D49A
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   720
         Width           =   240
      End
      Begin VB.CommandButton cmdDM 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   480
         Width           =   240
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   142
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "frmMain.frx":6D781
         cBack           =   -2147483633
      End
      Begin VB.CheckBox chkRenderizarAl 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Renderizar al cargar"
         Height          =   315
         Left            =   120
         TabIndex        =   141
         Top             =   2640
         Width           =   1095
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   143
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "frmMain.frx":6E3D3
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   144
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "frmMain.frx":6F025
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   145
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "frmMain.frx":6FC77
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   146
         Top             =   1800
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "1"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   147
         Top             =   1800
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "2"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   148
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
      Begin WorldEditor.lvButtons_H LvBOpcion 
         Height          =   375
         Index           =   7
         Left            =   720
         TabIndex        =   149
         Top             =   2160
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "4"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   20400
      Picture         =   "frmMain.frx":708C9
      ScaleHeight     =   660
      ScaleWidth      =   1455
      TabIndex        =   121
      Top             =   11160
      Width           =   1455
   End
   Begin VB.PictureBox Minimap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   300
      TabIndex        =   97
      Top             =   30
      Width           =   3023
      Begin VB.Shape UserArea 
         BorderColor     =   &H80000004&
         Height          =   297
         Left            =   1320
         Top             =   1440
         Width           =   400
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   2775
         Left            =   120
         Top             =   119
         Width           =   2745
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   17640
      TabIndex        =   89
      Top             =   0
      Width           =   2865
      Begin WorldEditor.lvButtons_H cmdInformacionDelMapa 
         Height          =   375
         Left            =   105
         TabIndex        =   90
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
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
         TabIndex        =   139
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
         TabIndex        =   138
         Top             =   315
         Width           =   1095
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
         Height          =   390
         Left            =   120
         TabIndex        =   96
         Top             =   0
         Width           =   2655
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
         Height          =   405
         Left            =   120
         TabIndex        =   95
         Top             =   960
         Width           =   2655
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   120
      ScaleHeight     =   2160
      ScaleWidth      =   4425
      TabIndex        =   87
      Top             =   9600
      Visible         =   0   'False
      Width           =   4485
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Largo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   180
         Top             =   900
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   179
         Top             =   555
         Width           =   525
      End
   End
   Begin VB.PictureBox Renderer 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9360
      Left            =   4680
      ScaleHeight     =   629.873
      ScaleMode       =   0  'User
      ScaleWidth      =   1143
      TabIndex        =   86
      Top             =   1440
      Width           =   17175
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   6
      Left            =   11760
      TabIndex        =   37
      Top             =   0
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
      Image           =   "frmMain.frx":74277
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   5
      Left            =   10320
      TabIndex        =   36
      Top             =   0
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
      Image           =   "frmMain.frx":7483D
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   3
      Left            =   8955
      TabIndex        =   35
      Top             =   0
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
      Image           =   "frmMain.frx":74D3E
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   2
      Left            =   7440
      TabIndex        =   34
      Top             =   0
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
      Image           =   "frmMain.frx":750F2
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   1
      Left            =   5880
      TabIndex        =   33
      Top             =   0
      Width           =   2627
      _ExtentX        =   4630
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
      Image           =   "frmMain.frx":78EE1
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   0
      Left            =   5191
      TabIndex        =   32
      Top             =   0
      Width           =   1753
      _ExtentX        =   3096
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
      Image           =   "frmMain.frx":7C541
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H cmdQuitarFunciones 
      Height          =   555
      Left            =   20520
      TabIndex        =   31
      ToolTipText     =   "Quitar Todas las Funciones Activadas"
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   979
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
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2565
      Top             =   3345
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
      Image           =   "frmMain.frx":7FA87
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   7
      Left            =   13080
      TabIndex        =   98
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1826
      Caption         =   "&Copiar Mapa"
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
      Image           =   "frmMain.frx":7FE3B
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   8
      Left            =   14445
      TabIndex        =   114
      Top             =   0
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
      Image           =   "frmMain.frx":8047C
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   9
      Left            =   15810
      TabIndex        =   120
      Top             =   0
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
      Image           =   "frmMain.frx":80AFE
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H LvBGuardarMinimapa 
      Height          =   375
      Index           =   8
      Left            =   3360
      TabIndex        =   168
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "G/MiniM"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
   Begin WorldEditor.lvButtons_H SelectPanelextra 
      Height          =   1275
      Index           =   0
      Left            =   22440
      TabIndex        =   208
      Top             =   0
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   2249
      Caption         =   "&Insertar Bloqueos en Bordes Mapa"
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
      Image           =   "frmMain.frx":80FA0
      ImgSize         =   24
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
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   320
      X2              =   320
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
      Begin VB.Menu mnuAbrirMapaInt 
         Caption         =   "&Abrir Mapa (Int)"
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
         Caption         =   "&Conversor y Utilidades"
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
      Begin VB.Menu cmdampliacion 
         Caption         =   "Ampliar Mapa"
      End
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
         Caption         =   "&Realizar Operación en Seleccón"
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
         Begin VB.Menu mnuCostas 
            Caption         =   "Costas Automaticas (BETA)"
         End
         Begin VB.Menu mnuBloquearBordes 
            Caption         =   "Bloqueo en &Bordes del Mapa"
         End
         Begin VB.Menu mnuInsertarTransladosAdyasentes 
            Caption         =   "&Translados a Mapas Adyasentes"
         End
         Begin VB.Menu mnuLinea11 
            Caption         =   "-"
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
         Begin VB.Menu mnuBloquearMapa 
            Caption         =   "Bloqueo en &Todo el Mapa"
         End
      End
      Begin VB.Menu mnuquitar 
         Caption         =   "&Quitar"
         Begin VB.Menu mnuTrasladosMap 
            Caption         =   "Traslados legales"
         End
         Begin VB.Menu mnuQuitarBloqueosBorde 
            Caption         =   "Bloqueos de los borde"
         End
         Begin VB.Menu mnuQuitarSuperficieBordes 
            Caption         =   "Superficie de los B&ordes"
         End
         Begin VB.Menu mnuQuitarSuperficieDeCapa 
            Caption         =   "Superficie de la &Capa Seleccionada"
         End
         Begin VB.Menu mnuLine10 
            Caption         =   "-"
         End
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
         Begin VB.Menu mnuquitararboles 
            Caption         =   "Todos los Arboles"
         End
         Begin VB.Menu mnuQuitarTriggers 
            Caption         =   "Todos los Tri&gger's"
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
      Begin VB.Menu mnulmpcostas 
         Caption         =   "Limpiar costas"
      End
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
         Caption         =   "&Guardar Ultima Configuracón"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuIniciarWE 
         Caption         =   "Carga Inicial"
      End
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
Dim X, y, i
Dim Head As Integer
Dim Body As Integer
Dim Heading As Byte
Dim Leer As New clsIniReader
i = n

modEdicion.Deshacer_Add "Aplicar " & IIf(T = 0, "Objetos", "NPCs") & " al Azar" ' Hago deshacer

Do While i > 0
    X = CInt(RandomNumber(XMinMapSize, XMaxMapSize - 1))
    y = CInt(RandomNumber(YMinMapSize, YMaxMapSize - 1))
    
    Select Case T
        Case 0
            If MapData(X, y).OBJInfo.objindex = 0 Then
                  i = i - 1
                  If cInsertarBloqueo.value = True Then
                    MapData(X, y).blocked = 1
                  Else
                    MapData(X, y).blocked = 0
                  End If
                  If cNumFunc(2).Text > 0 Then
                      objindex = cNumFunc(2).Text
                      InitGrh MapData(X, y).ObjGrh, ObjData(objindex).GrhIndex
                      MapData(X, y).OBJInfo.objindex = objindex
                      MapData(X, y).OBJInfo.Amount = Val(cCantFunc(2).Text)
                      Select Case ObjData(objindex).objtype ' GS
                            Case 4, 8, 10, 22 ' Arboles, Carteles, Foros, Yacimientos
                                MapData(X, y).Graphic(3) = MapData(X, y).ObjGrh
                      End Select
                  End If
            End If
        Case 1
           If MapData(X, y).blocked = 0 Then
                  i = i - 1
                  If cNumFunc(T - 1).Text > 0 Then
                        NPCIndex = cNumFunc(T - 1).Text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(y))
                        MapData(X, y).NPCIndex = NPCIndex
                  End If
            End If
        Case 2
           If MapData(X, y).blocked = 0 Then
                  i = i - 1
                  If cNumFunc(T - 1).Text >= 0 Then
                        NPCIndex = cNumFunc(T - 1).Text
                        Body = NpcData(NPCIndex).Body
                        Head = NpcData(NPCIndex).Head
                        Heading = NpcData(NPCIndex).Heading
                        Call MakeChar(NextOpenChar(), Body, Head, Heading, CInt(X), CInt(y))
                        MapData(X, y).NPCIndex = NPCIndex
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
If IsNumeric(cCantFunc(index).Text) = False Or cCantFunc(index).Text > 200 Then ' ver ReyarB
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


Private Sub chkOptMinimap_Click(index As Integer)
    Call DibujarMiniMapa
End Sub

Private Sub cInsertarFunc_Click(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cInsertarFunc(index).value = True Then
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
    If cInsertarLuz.value Then
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
If cInsertarTrans.value = True Then
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
If cInsertarTrigger.value = True Then
    cQuitarTrigger.Enabled = False
    Call modPaneles.EstSelectPanel(6, True)
Else
    cQuitarTrigger.Enabled = True
    Call modPaneles.EstSelectPanel(6, False)
End If
End Sub



Private Sub cmdDM_Click(index As Integer)
frmConfigSup.DespMosaic.value = vbChecked
Select Case index
    Case 0 'A

frmConfigSup.DMLargo.Text = Val(frmConfigSup.DMLargo.Text) + 1
    Case 1 '<
    frmConfigSup.DMAncho.Text = Val(frmConfigSup.DMAncho.Text) + 1
    Case 2 '>
    frmConfigSup.DMAncho.Text = Val(frmConfigSup.DMAncho.Text) - 1
    Case 3 'V
    frmConfigSup.DMLargo.Text = Val(frmConfigSup.DMLargo.Text) - 1
    Case 4 '0
frmConfigSup.DMAncho.Text = 0
frmConfigSup.DMLargo.Text = 0
End Select
End Sub

Private Sub cmdIrAlMapa_Click()
NumMap_Save = 7
Call MapPest_Click(TxtMapa)
End Sub

Private Sub cQuitarLuz_Click()
'*************************************************
'Author: Lorwik
'*************************************************
    If cQuitarLuz.value Then
        cInsertarLuz.Enabled = False
    Else
        cInsertarLuz.Enabled = True
    End If
End Sub
Private Sub cmdAdd_Click()

If cmdAdd.value = True Then
    lstParticle.Enabled = True
    cmdDel.Enabled = False
    Call modPaneles.EstSelectPanel(8, True)
Else
    lstParticle.Enabled = False
    cmdDel.Enabled = True
    Call modPaneles.EstSelectPanel(8, False)
End If
End Sub

Private Sub cmdampliacion_Click()
    frmAmpliacion.Show
End Sub


Private Sub cmdDel_Click()

If cmdDel.value = True Then
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
'*******************************************************
'Ultima modificacion 08/05/2020 por ReyarB
'*******************************************************
Private Sub COPIAR_GRH_Click(index As Integer)

    Dim y As Integer
    Dim X As Integer
    Dim An As Integer
    Dim Bn As Integer
    Dim Cn As Integer
    Dim Dn As Integer
    Dim ExtraBorde As Integer
    Dim Norte As Integer, Sur As Integer, Este As Integer, Oeste As Integer
    
    frmMain.MemoriaAuxiliar.Visible = True
    frmMain.COPIAR_GRH(0).Visible = False
    frmMain.COPIAR_GRH(1).Visible = False
    frmMain.COPIAR_GRH(2).Visible = False
    frmMain.COPIAR_GRH(3).Visible = False
    frmMain.COPIAR_GRH(4).Visible = False
    frmMain.COPIAR_GRH(5).Visible = False
    frmMain.COPIAR_GRH(6).Visible = False
    frmMain.COPIAR_GRH(7).Visible = False
    frmMain.COPIAR_GRH(8).Visible = False
    frmMain.COPIAR_GRH(9).Visible = False
    frmMain.COPIAR_GRH(10).Visible = False
    frmMain.COPIAR_GRH(11).Visible = False
    frmMain.COPIAR_GRH(12).Visible = False
    frmMain.lvlMapaCompleto.Visible = False
    frmMain.OpcBorde.Visible = False
    frmMain.OpcExit.Visible = False
    
    If frmMain.Option2.value = True Then
        If frmMain.OpcBorde.value = True Then
            Select Case index
                Case 0
                    ExtraBorde = 6
                Case 1
                    ExtraBorde = 8
                Case 2
                    ExtraBorde = 8
                Case 3
                    ExtraBorde = 6
            End Select
            Else
            ExtraBorde = 0
        End If
    Else
        If frmMain.OpcBorde.value = True Then
            
            Select Case index
                Case 0
                    ExtraBorde = 9
                Case 1
                    ExtraBorde = 13
                Case 2
                    ExtraBorde = 13
                Case 3
                    ExtraBorde = 9
            End Select
            Else
            ExtraBorde = 0
        End If
    
    End If
Call Resolucion

    If frmMain.Option2.value = True Then
            An = 6 + ExtraBorde
            Bn = 92 + ExtraBorde
            Cn = 6 + ExtraBorde
            Dn = 92 + ExtraBorde
            TXTArriba = 87
            TxTDerecha = 78
            TxTIzquierda = 78
            TxTAbajo = 87
        Else
            TXTArriba = 80
            TxTDerecha = 74
            TxTIzquierda = 74
            TxTAbajo = 80
            An = 11 + ExtraBorde 'arriba
            Bn = 88 - ExtraBorde 'derecha
            Cn = 13 + ExtraBorde 'izq
            Dn = 90 - ExtraBorde 'abajo
    End If
    
    On Error Resume Next
 
    Select Case index

        Case 0 'Sur
        
        Norte = 0
        Sur = 0
        Este = 0
        Oeste = 0
                    
        Call LeerAdyacentes(Norte, Sur, Este, Oeste)
        
        If Sur = 0 Then
        Call MsgBox("No hay traslados al mapa Sur, no se puede pegar el borde, compruebe si es correcto el Formato del mapa.")
            Exit Sub
        End If
        
        Call MapEspejo(Sur)
        
        
            For y = 1 To An  ' borrado
                For X = 1 To XMaxMapSize
                
                    'Quitar NPCs
                    If MapData(X, y).NPCIndex > 0 Then
                        EraseChar MapData(X, y).CharIndex
                        MapData(X, y).NPCIndex = 0
                    End If
    
                    ' Quitar Objetos
                    MapData(X, y).OBJInfo.objindex = 0
                    MapData(X, y).OBJInfo.Amount = 0
                    MapData(X, y).ObjGrh.GrhIndex = 0
    
                    ' Quitar Triggers
                    MapData(X, y).Trigger = 0
              
                    ' Quitar Graficos
                    MapData(X, y).Graphic(1).GrhIndex = 0
                    MapData(X, y).Graphic(2).GrhIndex = 0
                    MapData(X, y).Graphic(3).GrhIndex = 0
                    MapData(X, y).OBJInfo.objindex = 0

                Next
            Next

            For y = 1 To An
                For X = 1 To XMaxMapSize
                    MapData(X, y).Graphic(1) = MapData_Adyacente(X, TXTArriba + y).Graphic(1)
                    MapData(X, y).Graphic(2) = MapData_Adyacente(X, TXTArriba + y).Graphic(2)
                    MapData(X, y).Graphic(3) = MapData_Adyacente(X, TXTArriba + y).Graphic(3)
                    MapData(X, y).Graphic(4) = MapData_Adyacente(X, TXTArriba + y).Graphic(4)
                    MapData(X, y).Trigger = MapData_Adyacente(X, TXTArriba + y).Trigger
                    MapData(X, y).ObjGrh = MapData_Adyacente(X, TXTArriba + y).ObjGrh
                    MapData(X, y).OBJInfo = MapData_Adyacente(X, TXTArriba + y).OBJInfo

                Next
            Next
            MapInfo.Changed = 1
            UserPos.y = 12

        Case 1 'Oeste
        
        Norte = 0
        Sur = 0
        Este = 0
        Oeste = 0
                    
        Call LeerAdyacentes(Norte, Sur, Este, Oeste)
        
        If Oeste = 0 Then
        Call MsgBox("No hay traslados al mapa Oeste, no se puede pegar el borde, compruebe si es correcto el Formato del mapa.")
            Exit Sub
        End If
        
        Call MapEspejo(Oeste)
        
            For y = 1 To YMaxMapSize
                For X = Bn To XMaxMapSize
                    'Quitar NPCs
                    If MapData(X, y).NPCIndex > 0 Then
                        EraseChar MapData(X, y).CharIndex
                        MapData(X, y).NPCIndex = 0
                    End If
    
                    ' Quitar Objetos
                    MapData(X, y).OBJInfo.objindex = 0
                    MapData(X, y).OBJInfo.Amount = 0
                    MapData(X, y).ObjGrh.GrhIndex = 0
    
                    ' Quitar Triggers
                    MapData(X, y).Trigger = 0
              
                    ' Quitar Graficos
                    MapData(X, y).Graphic(1).GrhIndex = 0
                    MapData(X, y).Graphic(2).GrhIndex = 0
                    MapData(X, y).Graphic(3).GrhIndex = 0
                    MapData(X, y).OBJInfo.objindex = 0
                Next
            Next

            For y = 1 To YMaxMapSize
                For X = Bn To XMaxMapSize
                    MapData(X, y).Graphic(1) = MapData_Adyacente(X - TxTDerecha, y).Graphic(1)
                    MapData(X, y).Graphic(2) = MapData_Adyacente(X - TxTDerecha, y).Graphic(2)
                    MapData(X, y).Graphic(3) = MapData_Adyacente(X - TxTDerecha, y).Graphic(3)
                    MapData(X, y).Graphic(4) = MapData_Adyacente(X - TxTDerecha, y).Graphic(4)
                    MapData(X, y).ObjGrh = MapData_Adyacente(X - TxTDerecha, y).ObjGrh
                    MapData(X, y).OBJInfo = MapData_Adyacente(X - TxTDerecha, y).OBJInfo

                Next
            Next
            MapInfo.Changed = 1
            UserPos.X = 85
                        
        Case 2 'Este
        
        Norte = 0
        Sur = 0
        Este = 0
        Oeste = 0
                    
        Call LeerAdyacentes(Norte, Sur, Este, Oeste)
        
        If Este = 0 Then
        Call MsgBox("No hay traslados al mapa Este, no se puede pegar el borde, compruebe si es correcto el Formato del mapa.")
            Exit Sub
        End If
        
        Call MapEspejo(Este)

            For y = 1 To YMaxMapSize
                For X = 1 To Cn
                    'Quitar NPCs
                    If MapData(X, y).NPCIndex > 0 Then
                        EraseChar MapData(X, y).CharIndex
                        MapData(X, y).NPCIndex = 0
                    End If
    
                    ' Quitar Objetos
                    MapData(X, y).OBJInfo.objindex = 0
                    MapData(X, y).OBJInfo.Amount = 0
                    MapData(X, y).ObjGrh.GrhIndex = 0
    
                    ' Quitar Triggers
                    MapData(X, y).Trigger = 0
              
                    ' Quitar Graficos
                    MapData(X, y).Graphic(1).GrhIndex = 0
                    MapData(X, y).Graphic(2).GrhIndex = 0
                    MapData(X, y).Graphic(3).GrhIndex = 0
                    MapData(X, y).OBJInfo.objindex = 0
                Next
            Next

            For y = 1 To YMaxMapSize
                For X = 1 To Cn
                    MapData(X, y).Graphic(1) = MapData_Adyacente(X + TxTIzquierda, y).Graphic(1)
                    MapData(X, y).Graphic(2) = MapData_Adyacente(X + TxTIzquierda, y).Graphic(2)
                    MapData(X, y).Graphic(3) = MapData_Adyacente(X + TxTIzquierda, y).Graphic(3)
                    MapData(X, y).Graphic(4) = MapData_Adyacente(X + TxTIzquierda, y).Graphic(4)
                    MapData(X, y).ObjGrh = MapData_Adyacente(X + TxTIzquierda, y).ObjGrh
                    MapData(X, y).OBJInfo = MapData_Adyacente(X + TxTIzquierda, y).OBJInfo

                Next
            Next
            MapInfo.Changed = 1
            UserPos.X = 21
                        
        Case 3 'Norte
        
        Norte = 0
        Sur = 0
        Este = 0
        Oeste = 0
                    
        Call LeerAdyacentes(Norte, Sur, Este, Oeste)
        
        If Norte = 0 Then
        Call MsgBox("No hay traslados al mapa Norte, no se puede pegar el borde, compruebe si es correcto el Formato del mapa.")
            Exit Sub
        End If
        
        Call MapEspejo(Norte)

            For y = Dn To YMaxMapSize
                For X = 1 To XMaxMapSize
                    'Quitar NPCs
                    If MapData(X, y).NPCIndex > 0 Then
                        EraseChar MapData(X, y).CharIndex
                        MapData(X, y).NPCIndex = 0
                    End If
    
                    ' Quitar Objetos
                    MapData(X, y).OBJInfo.objindex = 0
                    MapData(X, y).OBJInfo.Amount = 0
                    MapData(X, y).ObjGrh.GrhIndex = 0
    
                    ' Quitar Triggers
                    MapData(X, y).Trigger = 0
              
                    ' Quitar Graficos
                    MapData(X, y).Graphic(1).GrhIndex = 0
                    MapData(X, y).Graphic(2).GrhIndex = 0
                    MapData(X, y).Graphic(3).GrhIndex = 0
                    MapData(X, y).OBJInfo.objindex = 0
                Next
            Next
            For y = Dn To YMaxMapSize
                For X = 1 To XMaxMapSize
                    MapData(X, y).Graphic(1) = MapData_Adyacente(X, y - TxTAbajo).Graphic(1)
                    MapData(X, y).Graphic(2) = MapData_Adyacente(X, y - TxTAbajo).Graphic(2)
                    MapData(X, y).Graphic(3) = MapData_Adyacente(X, y - TxTAbajo).Graphic(3)
                    MapData(X, y).Graphic(4) = MapData_Adyacente(X, y - TxTAbajo).Graphic(4)
                    MapData(X, y).ObjGrh = MapData_Adyacente(X, y - TxTAbajo).ObjGrh
                    MapData(X, y).OBJInfo = MapData_Adyacente(X, y - TxTAbajo).OBJInfo

                Next
            Next
            MapInfo.Changed = 1
            UserPos.y = 88
                       
        Case 4 'Mapa entero en posicion 1
        
            Call BorrarMapa(0, 0)
            Call PegarMapa(0, 0)
            MapInfo.Changed = 1
            
        Case 5 'Mapa entero en posicion 2
        
            Call BorrarMapa(100, 0)
            Call PegarMapa(100, 0)
            MapInfo.Changed = 1
            
        Case 8 'Mapa entero en posicion 3
        
            Call BorrarMapa(200, 0)
            Call PegarMapa(200, 0)
            MapInfo.Changed = 1
            
        Case 6 'Mapa entero en posicion 4
        
            Call BorrarMapa(0, 100)
            Call PegarMapa(0, 100)
            MapInfo.Changed = 1
            
        Case 7 'Mapa entero en posicion 5
        
            Call BorrarMapa(100, 100)
            Call PegarMapa(100, 100)
            MapInfo.Changed = 1
            
        Case 9 'Mapa entero en posicion 6
        
            Call BorrarMapa(200, 100)
            Call PegarMapa(200, 100)
            MapInfo.Changed = 1
        MapInfo.Changed = 1
        
        Case 10 'Mapa entero en posicion 7
        
            Call BorrarMapa(0, 200)
            Call PegarMapa(0, 200)
            MapInfo.Changed = 1
            
        Case 11 'Mapa entero en posicion 8
        
            Call BorrarMapa(100, 200)
            Call PegarMapa(100, 200)
            MapInfo.Changed = 1
            
        Case 12 'Mapa entero en posicion 9
        
            Call BorrarMapa(200, 200)
            Call PegarMapa(200, 200)
            MapInfo.Changed = 1
        MapInfo.Changed = 1

            
    End Select
    
    Call modEdicion.Bloquear_Bordes(1)

End Sub


Private Sub cUnionManual_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cInsertarTrans.value = (cUnionManual.value = True)
Call cInsertarTrans_Click
End Sub

Private Sub cverBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerBloqueos.Checked = cVerBloqueos.value
End Sub

Private Sub cverTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerTriggers.Checked = cVerTriggers.value
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
If cInsertarBloqueo.value = True Then
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
If cQuitarBloqueo.value = True Then
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
If cQuitarEnEstaCapa.value = True Then
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
If cQuitarEnTodasLasCapas.value = True Then
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
If cQuitarFunc(index).value = True Then
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
If cQuitarTrans.value = True Then
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
If cQuitarTrigger.value = True Then
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
If cSeleccionarSuperficie.value = True Then
    cQuitarEnTodasLasCapas.Enabled = False
    cQuitarEnEstaCapa.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
    frmConfigSup.Visible = True
    
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

Private Sub Form_Load()
    frmMain.Dialog.FilterIndex = 1
End Sub

Private Sub LvBGuardarMinimapa_Click(index As Integer)
    Shape1.Visible = False
    UserArea.Visible = False
    DoEvents
    Call frmRender.Capturar_Imagen(frmMain.Minimap, frmMain.MinimapCapture)
    SavePicture frmMain.MinimapCapture, App.Path & "\Recursos\Graficos\MiniMapa\" & NumMap_Save & ".bmp"
    Shape1.Visible = True
    UserArea.Visible = True
    Call DibujarMiniMapa
End Sub

Private Sub LvBOpcion_Click(index As Integer)
    Select Case index
        Case 0
            cVerBloqueos.value = (cVerBloqueos.value = False)
            mnuVerBloqueos.Checked = cVerBloqueos.value
        Case 1
            mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)
        Case 2
            mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)
        Case 3
            cVerTriggers.value = (cVerTriggers.value = False)
            mnuVerTriggers.Checked = cVerTriggers.value
        Case 4
            mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)
        Case 5
            mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)
        Case 6
            mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)
        Case 7
            mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)
    End Select
End Sub

Private Sub LvBAreas_Click(index As Integer)
    If IsNumeric(DX1.Text) = False Or _
       IsNumeric(DX2.Text) = False Or _
       IsNumeric(DY1.Text) = False Or _
       IsNumeric(DY2.Text) = False Then
    
        Call MsgBox("Debes introducir valores númericos. Estos pueden tener un mÃ¯Â¿Â½nimo de 1 y un mÃ¯Â¿Â½ximo de " & (YMinMapSize + XMinMapSize) / 2 & ".")
    
       Exit Sub
    End If
    
    Select Case index
        Case 0
            Call Bloqueos_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
        Case 1
            Call Bloqueos_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
        Case 2
            Call Superficie_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, True)
        Case 3
            Call Superficie_Area(DX1.Text, DX2.Text, DY1.Text, DY2.Text, False)
    End Select
End Sub
'*****************************************************
'Ultima modificacion 08/05/2020 por ReyarB
'*****************************************************
Private Sub lvButtons_H5_Click(index As Integer)

    Select Case index
    
        Case 0
            R = 255
            G = 0
            B = 0
        Case 1
            R = 255
            G = 255
            B = 0
        Case 2
            R = 192
            G = 192
            B = 192
        Case 3
            R = 255
            G = 0
            B = 255
        Case 4
            R = 255
            G = 255
            B = 255
        Case 5
            R = 127
            G = 255
            B = 255

    
    End Select

End Sub



Private Sub LvBVerMapa_Click()
  If frmMapa.Visible Then
    frmMapa.Hide
  Else
    frmMapa.Show
  End If
End Sub



Private Sub mnuAbrirMapaInt_Click()
'*************************************************
'Author: Lorwik
'Last modified: 25/04/2020
'*************************************************
Dialog.CancelError = True
On Error GoTo errhandler

DeseaGuardarMapa Dialog.FileName

ObtenerNombreArchivo False

If Len(Dialog.FileName) < 3 Then Exit Sub

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode
    End If
    
    Call modMapIO.NuevoMapa
    
    'Tambien podrá elegir CSM, pero no habra diferencia
    If frmMain.Dialog.FilterIndex = 2 Then
        modMapIO.Cargar_CSM Dialog.FileName
    ElseIf frmMain.Dialog.FilterIndex = 1 Then
        modMapIO.MapaV2_Cargar Dialog.FileName, True
    End If
    
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
    
Exit Sub
errhandler:
End Sub

Private Sub mnuCostas_Click()
    Call PutCoast
End Sub

Private Sub mnuImpServer_Click()
Call frmImpCliente.Show
End Sub

Private Sub mnuIniciarWE_Click()
 frmImpCliente.Show
End Sub

Private Sub mnulmpcostas_Click()
    Call AddtoRichTextBox(frmMain.StatTxt, "Limpiando costas", 43, 0, 255)
    Call LimpiarCostas
End Sub

Private Sub mnuMasterObjetos_Click()
'    Dim Verifica_archivo
'    Dim cadena1 As String
'    Verifica_archivo = Dir(DirDats & "\obj.dat")
'    If Verifica_archivo <> "" Then
'        Dim Nuevo As New FileSystemObject, Nuevo1
'        Set Nuevo1 = Nuevo.GetFile(DirDats & "\obj.dat")
'        Nuevo1.Copy (DirDats & "\objetos.csv")
'    End If
'
'        Call Reemplazar_Texto(DirDats & "\objetos.csv", "=", ",")
'        Call ObjetosExel
'        Call CrearMasterDatos
End Sub

Private Sub mnuquitararboles_Click()
    Call modEdicion.Quitar_Arboles
End Sub

Private Sub mnuQuitarBloqueosBorde_Click()
 Call Resolucion
 Call Bloquear_Bordes(0)
End Sub



Private Sub mnuTrasladosMap_Click()
Call Quitar_TrasladosMap
End Sub

Private Sub Option1_Click()
Call Resolucion
End Sub

Private Sub Option2_Click()
Call Resolucion
End Sub

Private Sub OptX_Click(index As Integer)
'*************************************************
'Author: Lorwik
'Last modified: 25/04/2020
'*************************************************
'Nota: Hay que cambiar muchas cosas, el engine cuando inicia hace calculos con el tamaño de los mapas
'ademas hay mas funciones que manejan estos datos, no basta con cambiar el XMax & YMax.
   Call Resolucion
      Select Case index
    
        Case 0
            XMaxMapSize = 100
            YMaxMapSize = 100
            frmMain.Minimap.ScaleHeight = 100
            frmMain.Minimap.ScaleWidth = 100
            frmMain.UserArea.Height = 32
            frmMain.UserArea.Width = 40
            frmMain.Shape1.Height = 83
            frmMain.Shape1.Width = 76
            frmMain.Shape1.Left = 12
            frmMain.Shape1.Top = 9
                If frmMain.CopyBorder.Visible = True Then
                frmMain.COPIAR_GRH(4).Visible = False
                frmMain.COPIAR_GRH(5).Visible = False
                frmMain.COPIAR_GRH(6).Visible = False
                frmMain.COPIAR_GRH(7).Visible = False
                frmMain.COPIAR_GRH(8).Visible = False
                frmMain.COPIAR_GRH(9).Visible = False
                frmMain.COPIAR_GRH(10).Visible = False
                frmMain.COPIAR_GRH(11).Visible = False
                frmMain.COPIAR_GRH(12).Visible = False
                frmMain.lvlMapaCompleto.Visible = False
            
            End If
        
        Case 1
            XMaxMapSize = 200
            YMaxMapSize = 200
            frmMain.Minimap.ScaleHeight = 200
            frmMain.Minimap.ScaleWidth = 200
            frmMain.UserArea.Height = 30
            frmMain.UserArea.Width = 40
            frmMain.Shape1.Height = 183
            frmMain.Shape1.Width = 176
            frmMain.Shape1.Left = 12
            frmMain.Shape1.Top = 9
            If frmMain.CopyBorder.Visible = True Then
                frmMain.COPIAR_GRH(4).Visible = True
                frmMain.COPIAR_GRH(5).Visible = True
                frmMain.COPIAR_GRH(6).Visible = True
                frmMain.COPIAR_GRH(7).Visible = True
            End If

        Case 2
            XMaxMapSize = 300
            YMaxMapSize = 300
            frmMain.Minimap.ScaleHeight = 300
            frmMain.Minimap.ScaleWidth = 300
            frmMain.UserArea.Height = 30
            frmMain.UserArea.Width = 40
            frmMain.Shape1.Height = 283
            frmMain.Shape1.Width = 276
            frmMain.Shape1.Left = 12
            frmMain.Shape1.Top = 9
        If frmMain.CopyBorder.Visible = True Then
            frmMain.COPIAR_GRH(4).Visible = True
            frmMain.COPIAR_GRH(5).Visible = True
            frmMain.COPIAR_GRH(6).Visible = True
            frmMain.COPIAR_GRH(7).Visible = True
            frmMain.COPIAR_GRH(8).Visible = True
            frmMain.COPIAR_GRH(9).Visible = True
            frmMain.COPIAR_GRH(10).Visible = True
            frmMain.COPIAR_GRH(11).Visible = True
            frmMain.COPIAR_GRH(12).Visible = True
            frmMain.lvlMapaCompleto.Visible = True
            
        End If
    
       End Select
End Sub



Private Sub PreviewGrh_Click()
frmConfigSup.Visible = True
End Sub

Private Sub renderer_DblClick()
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
        cSeleccionarSuperficie.value = (cSeleccionarSuperficie.value = False)
        Call cSeleccionarSuperficie_Click
    Case "T" ' Activa/Desactiva Insertar Translados
        cInsertarTrans.value = (cInsertarTrans.value = False)
        Call cInsertarTrans_Click
    Case "B" ' Activa/Desactiva Insertar Bloqueos
        cInsertarBloqueo.value = (cInsertarBloqueo.value = False)
        Call cInsertarBloqueo_Click
    Case "N" ' Activa/Desactiva Insertar NPCs
        cInsertarFunc(0).value = (cInsertarFunc(0).value = False)
        Call cInsertarFunc_Click(0)
   ' Case "H" ' Activa/Desactiva Insertar NPCs Hostiles
   '     cInsertarFunc(1).value = (cInsertarFunc(1).value = False)
   '     Call cInsertarFunc_Click(1)
    Case "O" ' Activa/Desactiva Insertar Objetos
        cInsertarFunc(2).value = (cInsertarFunc(2).value = False)
        Call cInsertarFunc_Click(2)
    Case "G" ' Activa/Desactiva Insertar Triggers
        cInsertarTrigger.value = (cInsertarTrigger.value = False)
        Call cInsertarTrigger_Click
    Case "Q" ' Quitar Funciones
        Call mnuQuitarFunciones_Click
End Select
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
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
                If LenB(cInsertarBloqueo.Tag) = 0 Then cInsertarBloqueo.Tag = IIf(cInsertarBloqueo.value = True, 1, 0)
                cInsertarBloqueo.value = True
                Call cInsertarBloqueo_Click
            Else
                If LenB(cInsertarBloqueo.Tag) <> 0 Then
                    cInsertarBloqueo.value = IIf(Val(cInsertarBloqueo.Tag) = 1, True, False)
                    cInsertarBloqueo.Tag = vbNullString
                    Call cInsertarBloqueo_Click
                End If
            End If
            Call fPreviewGrh(cGrh.Text)
            Call modPaneles.VistaPreviaDeSup
        Case 1
            cNumFunc(0).Text = DameNPCsIndex(ReadField(2, lListado(index).Text, Asc("#")))
            Call fPreviewNPCs(cNumFunc(0).Text)
            Call modPaneles.VistaPreviaDeNPCs
            cNumFunc(0).Text = ReadField(2, lListado(index).Text, Asc("#"))
        Case 2
            cNumFunc(1).Text = ReadField(2, lListado(index).Text, Asc("#"))
        Case 3
            cNumFunc(2).Text = DameOBJIndex(ReadField(2, lListado(index).Text, Asc("#")))
            Call fPreviewObj(cNumFunc(2).Text)
            Call modPaneles.VistaPreviaDeObj
            cNumFunc(2).Text = ReadField(2, lListado(index).Text, Asc("#"))
    End Select
Else
    lListado(index).ListIndex = lListado(index).Tag
End If

End Sub

Private Sub lListado_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
If index = 3 And Button = 2 Then
    If lListado(3).ListIndex > -1 Then Me.PopupMenu mnuObjSc
End If
End Sub

Private Sub lListado_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
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

Private Sub MapPest_Click(index As Integer)


    '*************************************************
    'Author: ^[GS]^
    'Ultima modificacion 08/05/2020 por ReyarB
    '*************************************************
    Dim formato As String

    Select Case frmMain.Dialog.FilterIndex
    
        Case 2
            formato = ".csm"
            
        Case 1
            formato = ".map"
            
    End Select
    
    
        If MapInfo.Changed = 1 Then
            
            If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
                Call modMapIO.GuardarMapa(Dialog.FileName)
            End If

        End If
        
       
    If (index + NumMap_Save - 4) <> NumMap_Save Then
        Dialog.CancelError = True

        On Error GoTo errhandler

        Dialog.FileName = PATH_Save & NameMap_Save & (index + NumMap_Save - 7) & formato

        If FileSize(Dialog.FileName) > 300000 Then
            'MsgBox "File Size =" & FileSize(Dialog.FileName), vbInformation
            XMaxMapSize = 300
            YMaxMapSize = 300
            OptX(2).value = True
        ElseIf FileSize(Dialog.FileName) > 200000 And FileSize(Dialog.FileName) < 300000 Then
                'MsgBox "File Size =" & FileSize(Dialog.FileName), vbInformation
                XMaxMapSize = 200
                YMaxMapSize = 200
                OptX(1).value = True
        Else
            'MsgBox "File Size =" & FileSize(Dialog.FileName), vbInformation
            XMaxMapSize = 100
            YMaxMapSize = 100
            OptX(0).value = True
        End If
        
        Call modMapIO.NuevoMapa
        
        
        DoEvents
        Select Case frmMain.Dialog.FilterIndex
        
            Case 2
                Call modMapIO.Cargar_CSM(Dialog.FileName)
                
            Case 1
                Call modMapIO.MapaV2_Cargar(Dialog.FileName, MapaCargado_Integer)
            
        End Select
        
        EngineRun = True
        
    End If
Call ActualizaMinimap
        Exit Sub
    
errhandler:
        Call MsgBox(err.Description)

End Sub
'******************************************
'Ultima modificacion 08/05/2020 por ReyarB
'*******************************************
Private Sub MemoriaAuxiliar_Click()
On Error GoTo error
 
    MapData_Adyacente = MapData
    
    frmMain.MemoriaAuxiliar.Visible = False
    frmMain.COPIAR_GRH(0).Visible = True
    frmMain.COPIAR_GRH(1).Visible = True
    frmMain.COPIAR_GRH(2).Visible = True
    frmMain.COPIAR_GRH(3).Visible = True
    If XMaxMapSize = 200 Then
        frmMain.COPIAR_GRH(4).Visible = True
        frmMain.COPIAR_GRH(5).Visible = True
        frmMain.COPIAR_GRH(6).Visible = True
        frmMain.COPIAR_GRH(7).Visible = True
    ElseIf XMaxMapSize = 300 Then
        frmMain.COPIAR_GRH(4).Visible = False
        frmMain.COPIAR_GRH(5).Visible = False
        frmMain.COPIAR_GRH(6).Visible = False
        frmMain.COPIAR_GRH(7).Visible = False
        frmMain.COPIAR_GRH(8).Visible = False
        frmMain.COPIAR_GRH(9).Visible = False
        frmMain.COPIAR_GRH(10).Visible = False
        frmMain.COPIAR_GRH(11).Visible = False
        frmMain.COPIAR_GRH(12).Visible = False
        frmMain.lvlMapaCompleto.Visible = False
    End If

    frmMain.OpcBorde.Visible = True
    frmMain.OpcExit.Visible = True
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Mapa copiado a la memoria", 43, 0, 255)
     
    Exit Sub
error:
    Call AddtoRichTextBox(frmMain.StatTxt, "Error guardando mapa", 255, 0, 0)
End Sub
'*************************************************
'Author: ^[GS]^
'Modificado el 10/05/2020 por ReyarB
'*************************************************
Private Sub minimap_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If X < MinXBorder Then X = MinXBorder '11
If X > MaxXBorder Then X = MaxXBorder '89
If y < MinYBorder Then y = MinYBorder '10
If y > MaxYBorder Then y = MaxYBorder '92
    
    UserPos.X = X
    UserPos.y = y
    
    Call ActualizaMinimap
End Sub

Private Sub minimap_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'*************************************************
'Author: ^[GS]^
'Modificado el 10/05/2020 por ReyarB
'*************************************************
MiRadarX = X
MiRadarY = y
End Sub


Private Sub mnuAbrirMapaNew_Click()
'*************************************************
'Author: Lorwik
'Ultima modificacion 08/05/2020 por ReyarB
'*************************************************
Dialog.CancelError = True
On Error GoTo errhandler

DeseaGuardarMapa Dialog.FileName

ObtenerNombreArchivo False

If Len(Dialog.FileName) < 3 Then Exit Sub

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode
    End If
    
        If FileSize(Dialog.FileName) > 300000 Then
            'MsgBox "File Size =" & FileSize(Dialog.FileName), vbInformation
            XMaxMapSize = 300
            YMaxMapSize = 300
            OptX(2).value = True
        ElseIf FileSize(Dialog.FileName) > 200000 And FileSize(Dialog.FileName) < 300000 Then
                'MsgBox "File Size =" & FileSize(Dialog.FileName), vbInformation
                XMaxMapSize = 200
                YMaxMapSize = 200
                OptX(1).value = True
        Else
            'MsgBox "File Size =" & FileSize(Dialog.FileName), vbInformation
            XMaxMapSize = 100
            YMaxMapSize = 100
            OptX(0).value = True
        End If
           
    
    Call modMapIO.NuevoMapa
    If frmMain.Dialog.FilterIndex = 2 Then
    
        modMapIO.Cargar_CSM Dialog.FileName
        
    ElseIf frmMain.Dialog.FilterIndex = 1 Then
    
        modMapIO.MapaV2_Cargar Dialog.FileName
        
    End If
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True

Exit Sub
errhandler:
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
bAutoCompletarSuperficies = mnuAutoCompletarSuperficies.Checked
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
For i = 0 To 9
    If i <> 2 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next

modPaneles.VerFuncion 2, True
End Sub

Private Sub mnuBloquearBordes_Click()
'*************************************************
'Author: ^[GS]^
'Ultima modificacion 08/05/2020 por ReyarB
'*************************************************
Call Resolucion
Call modEdicion.Bloquear_Bordes(1)
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
frmGRHaBMP.Show
End Sub

Private Sub mnuGuardarMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modMapIO.GuardarMapa Dialog.FileName
End Sub

Private Sub mnuGuardarMapaComo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modMapIO.GuardarMapa
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
frmInformes.Show
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
Call Resolucion
frmUnionAdyacente.Show
End Sub

Private Sub mnuManual_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
If LenB(Dir(App.Path & "\Manual\index.htm", vbArchive)) <> 0 Then
    Call Shell("explorer " & App.Path & "\Manual\index.htm")
    DoEvents
End If
End Sub

Private Sub mnuModoCaminata_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
ToggleWalkMode
End Sub

Private Sub mnuNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 9
    If i <> 3 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 3, True
End Sub



'Private Sub mnuNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Dim i As Byte
'For i = 0 To 9
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
Dim loopc As Integer

DeseaGuardarMapa Dialog.FileName

For loopc = 0 To frmMain.MapPest.count
    frmMain.MapPest(loopc).Visible = False
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
For i = 0 To 9
    If i <> 5 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 5, True
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
modPaneles.VerFuncion 2, False
End Sub

Private Sub mnuQNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 3, False
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
modPaneles.VerFuncion 5, False
End Sub

Private Sub mnuQSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 0, False
End Sub

Private Sub mnuQTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 1, False
End Sub

Private Sub mnuQTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 6, False
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
cSeleccionarSuperficie.value = False
Call cSeleccionarSuperficie_Click
cQuitarEnEstaCapa.value = False
Call cQuitarEnEstaCapa_Click
cQuitarEnTodasLasCapas.value = False
Call cQuitarEnTodasLasCapas_Click
' Translados
cQuitarTrans.value = False
Call cQuitarTrans_Click
cInsertarTrans.value = False
Call cInsertarTrans_Click
' Bloqueos
cQuitarBloqueo.value = False
Call cQuitarBloqueo_Click
cInsertarBloqueo.value = False
Call cInsertarBloqueo_Click
' Otras funciones
cInsertarFunc(0).value = False
Call cInsertarFunc_Click(0)
cInsertarFunc(1).value = False
Call cInsertarFunc_Click(1)
cInsertarFunc(2).value = False
Call cInsertarFunc_Click(2)
cQuitarFunc(0).value = False
Call cQuitarFunc_Click(0)
cQuitarFunc(1).value = False
Call cQuitarFunc_Click(1)
cQuitarFunc(2).value = False
Call cQuitarFunc_Click(2)
' Triggers
cInsertarTrigger.value = False
Call cInsertarTrigger_Click
cQuitarTrigger.value = False
Call cQuitarTrigger_Click
'Luces
cInsertarLuz.value = False
Call cInsertarLuz_Click
cQuitarLuz.value = False
Call cQuitarLuz_Click
'Particulas
cmdAdd.value = False
Call cmdAdd_Click
cmdDel.value = False
Call cmdDel_Click
'cQuitarParticula.value = False
'Call cQuitarParticula_Click
'cInsertarParticula.value = False
'Call cInsertarParticula_Click

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
'Ultima modificacion 08/05/2020 por ReyarB
'*************************************************
Call Resolucion
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
On Error GoTo errhandler
    If FileExist(Dialog.FileName, vbArchive) = False Then Exit Sub
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa Dialog.FileName
        End If
    End If
    Call modMapIO.NuevoMapa
    
    If frmMain.Dialog.FilterIndex = 1 Then
        modMapIO.MapaV2_Cargar Dialog.FileName
    ElseIf frmMain.Dialog.FilterIndex = 2 Then
        modMapIO.Cargar_CSM Dialog.FileName
    End If
    
    DoEvents
    mnuReAbrirMapa.Enabled = True
    EngineRun = True
Exit Sub
errhandler:
End Sub

Private Sub mnuRealizarOperacion_Click()
'*************************************************
'Author: ^[GS]^
'Ultima modificacion 08/05/2020 por ReyarB
'*************************************************
Call modEdicion.Deshacer_Add("Realizar Operación en Selección")
Call AccionSeleccion
End Sub

Private Sub mnuRenderMapa_Click()
    Call frmRender.Show(vbModal)
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
For i = 0 To 9
    If i <> 0 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 0, True
End Sub

Private Sub mnuTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 9
    If i <> 1 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 1, True
End Sub

Private Sub mnuTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 9
    If i <> 6 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 6, True
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
cVerBloqueos.value = (cVerBloqueos.value = False)
mnuVerBloqueos.Checked = cVerBloqueos.value

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
cVerTriggers.value = (cVerTriggers.value = False)
mnuVerTriggers.Checked = cVerTriggers.value
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 - GS
'Last modified: 20/11/07 - Loopzer
'*************************************************

Dim tX As Integer
Dim tY As Integer

If Not MapaCargado Then Exit Sub

ConvertCPtoTP X, y, tX, tY

'If Shift = 1 And Button = 2 Then PegarSeleccion tX, tY: Exit Sub
If Shift = 1 And Button = 1 Then
    Seleccionando = True
    SeleccionIX = tX '+ UserPos.X
    SeleccionIY = tY '+ UserPos.Y
    DX1.Text = tX
    DY1.Text = tY
Else
    ClickEdit Button, tX, tY
End If

End Sub

Private Sub Renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Call Form_MouseMove(Button, Shift, X, y)
    MouseX = X
    MouseY = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 - GS
'*************************************************

Dim tX As Integer
Dim tY As Integer

'Make sure map is loaded
If Not MapaCargado Then Exit Sub
HotKeysAllow = True

ConvertCPtoTP X, y, tX, tY

PosX = "X: " & tX & " - Y: " & tY

 If Shift = 1 And Button = 1 Then
    Seleccionando = True
    SeleccionFX = tX '+ TileX
    SeleccionFY = tY '+ TileY
    DX2.Text = tX
    DY2.Text = tY
Else
    ClickEdit Button, tX, tY
End If
End Sub

Private Sub Renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Call Form_MouseDown(Button, Shift, X, y)
    Call DibujarMiniMapa
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************

' Guardar configuración
WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "GuardarConfig", IIf(frmMain.mnuGuardarUltimaConfig.Checked = True, "1", "0")
If frmMain.mnuGuardarUltimaConfig.Checked = True Then
    WriteVar IniPath & "WorldEditor.ini", "PATH", "UltimoMapa", Dialog.FileName
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "ControlAutomatico", IIf(frmMain.mnuVerAutomatico.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Capa2", IIf(frmMain.mnuVerCapa2.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Capa3", IIf(frmMain.mnuVerCapa3.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Capa4", IIf(frmMain.mnuVerCapa4.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Translados", IIf(frmMain.mnuVerTranslados.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Objetos", IIf(frmMain.mnuVerObjetos.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "NPCs", IIf(frmMain.mnuVerNPCs.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Triggers", IIf(frmMain.mnuVerTriggers.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Grilla", IIf(frmMain.mnuVerGrilla.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "Bloqueos", IIf(frmMain.mnuVerBloqueos.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "MOSTRAR", "LastPos", UserPos.X & "-" & UserPos.y
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "UtilizarDeshacer", IIf(frmMain.mnuUtilizarDeshacer.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarTrans", IIf(frmMain.mnuAutoCapturarTranslados.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarSup", IIf(frmMain.mnuAutoCapturarSuperficie.Checked = True, "1", "0")
    WriteVar IniPath & "WorldEditor.ini", "CONFIGURACION", "ObjTranslado", Val(Cfg_TrOBJ)
End If

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
        SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
If mnuAutoQuitarFunciones.Checked = True Then Call mnuQuitarFunciones_Click
Call VerFuncion(index, SelectPanel(index).value)
End Sub



Private Sub SelectPanelExtra_Click(index As Integer)
Select Case index
    Case 0
        Call Resolucion
        Call modEdicion.Bloquear_Bordes(1)
    Case 1
        Call Resolucion
        frmUnionAdyacente.Show
    Case 2
    

   End Select
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
    .Filter = "Mapas formato (*.map)|*.map|Mapas formato (*.csm)|*.csm"
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

' Lee los traslados del mapa y retorna los mapas adyacentes o cero si no tiene en esa direccion
Private Sub LeerAdyacentes(ByRef Norte As Integer, ByRef Sur As Integer, ByRef Este As Integer, ByRef Oeste As Integer)
    Dim X As Integer
    Dim y As Integer

    ' Norte
    y = MinYBorder
    For X = (MinXBorder + 1) To (MaxXBorder - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Norte = MapData(X, y).TileExit.Map
            Exit For
        End If
    Next

    ' Este
    X = MaxXBorder
    For y = (MinYBorder + 1) To (MaxYBorder - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Este = MapData(X, y).TileExit.Map
            Exit For
        End If
    Next

    ' Sur
    y = MaxYBorder
    For X = (MinXBorder + 1) To (MaxXBorder - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Sur = MapData(X, y).TileExit.Map
            Exit For
        End If
    Next

    ' Oeste
    X = MinXBorder
    For y = (MinYBorder + 1) To (MaxYBorder - 1)
        If MapData(X, y).TileExit.Map > 0 Then
            Oeste = MapData(X, y).TileExit.Map
            Exit For
        End If
    Next
End Sub

Private Sub MapEspejo(index As Integer)

    '*************************************************
    'Author: ^[ReyarB]^
    'Last modified: 20/04/2020
    '*************************************************
    Dim formato As String
    
    Select Case frmMain.Dialog.FilterIndex
        Case 2
            formato = ".csm"
        Case 1
            formato = ".map"
    End Select
    
        If MapInfo.Changed = 1 Then
            If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
                Call modMapIO.GuardarMapa(Dialog.FileName)
            End If
        End If

        Dialog.FileName = PATH_Save & NameMap_Save & (index) & formato
        Call modMapIO.NuevoMapa
        
        DoEvents
        Select Case frmMain.Dialog.FilterIndex
            Case 2
                Call modMapIO.Cargar_CSM(Dialog.FileName)
                
            Case 1
                Call modMapIO.MapaV2_Cargar(Dialog.FileName, MapaCargado_Integer)
            
        End Select
        
        EngineRun = True


    Exit Sub
    
errhandler:
        Call MsgBox(err.Description)
End Sub
'**********************************************************************
'**********************************************************************
Private Sub PegarMapa(ByVal mX As Integer, ByVal mY As Integer)
On Error GoTo err
Dim OffsetX As Integer
Dim OffsetY As Integer
Dim X As Integer, y As Integer


    OffsetX = X + mX
    OffsetY = y + mY

    For X = 1 To 100
        For y = 1 To 100
        
            If OffsetX + X > 0 And OffsetX + X < 301 Then
              If OffsetY + y > 0 And OffsetY + y < 301 Then
              
                With MapData(X + OffsetX, y + OffsetY)
    
                    .Graphic(1) = MapData_Adyacente(X, y).Graphic(1)
                    .Graphic(2) = MapData_Adyacente(X, y).Graphic(2)
                    .Graphic(3) = MapData_Adyacente(X, y).Graphic(3)
                    .Graphic(4) = MapData_Adyacente(X, y).Graphic(4)
                    .blocked = MapData_Adyacente(X, y).blocked
                    .NPCIndex = MapData_Adyacente(X, y).NPCIndex
                    .Trigger = MapData_Adyacente(X, y).Trigger
                    .ObjGrh = MapData_Adyacente(X, y).ObjGrh
                    .OBJInfo = MapData_Adyacente(X, y).OBJInfo
                End With
              End If
            End If
          
        Next
    Next
    Call BorrarBloqueos
err:
Debug.Print err.Description
Debug.Print "error en pegarmapa"
End Sub

Private Sub BorrarMapa(ByVal mX As Integer, ByVal mY As Integer)
Dim GrhNull As Grh
Dim ObjectNull As Obj
Dim X As Integer, y As Integer

            For y = 1 + mY To 100 + mY ' borrado
                For X = 1 + mX To 100 + mX

                    'Quitar NPCs
                    If MapData(X, y).NPCIndex > 0 Then
                        EraseChar MapData(X, y).CharIndex
                        MapData(X, y).NPCIndex = 0
                    End If
                    ' Quitar Objetos
                    MapData(X, y).OBJInfo.objindex = 0
                    MapData(X, y).OBJInfo.Amount = 0
                    MapData(X, y).ObjGrh.GrhIndex = 0
                    ' Quitar Triggers
                    MapData(X, y).Trigger = 0
                    ' Quitar Bloqueos
                    MapData(X, y).blocked = 0
                    ' Quitar Graficos
                    MapData(X, y).Graphic(1).GrhIndex = 0
                    MapData(X, y).Graphic(2).GrhIndex = 0
                    MapData(X, y).Graphic(3).GrhIndex = 0
                    MapData(X, y).Graphic(4).GrhIndex = 0
                    
                Next
            Next

End Sub

Private Sub BorrarBloqueos()
Dim X As Integer
Dim y As Integer
    For X = XMinMapSize To XMaxMapSize
        For y = YMinMapSize To YMaxMapSize
        
        If MapData(X, y).Graphic(2).GrhIndex > 0 Or _
           MapData(X, y).Graphic(3).GrhIndex > 0 Or _
           MapData(X, y).Graphic(4).GrhIndex > 0 Or _
           MapData(X, y).OBJInfo.objindex > 0 Then GoTo Jump
        
        If X >= 13 And y >= 92 And y <= 109 Then MapData(X, y).blocked = 0
        If X >= 89 And X <= 112 And y >= 10 Then MapData(X, y).blocked = 0
'        If X >= 192 And X <= 211 And y >= 10 Then MapData(X, y).blocked = 0
'        If X >= 13 And X <= 92 And y >= 192 And y <= 193 Then MapData(X, y).blocked = 0

'        If X >= 109 And X <= 191 And y >= 188 And y <= 193 Then MapData(X, y).blocked = 0
'        If X >= 209 And X <= 274 And y >= 195 And y <= 206 Then MapData(X, y).blocked = 0
'        If X >= 109 And X <= 192 Then MapData(X, y).blocked = 0
'        If y >= 182 And y <= 188 Then MapData(X, y).blocked = 0
        
Jump:
        Next
    Next
End Sub
Private Sub Form_Resize()
'***********************************************
'Autor: Lorwik
'Fecha: 11/05/2020
'Descripcion: Ajusta los controles cuando se redimensiona la ventana
'***********************************************
    Me.Renderer.Height = Me.Height / 19 'Alto
    Me.Renderer.Width = Me.Width / 18 'Ancho

    StatTxt.Top = Me.Renderer.Height + 100
    FraFormatoDel.Top = Me.Renderer.Height + 100
    Picture2.Top = Me.Renderer.Height + 100
    Picture1.Top = Me.Renderer.Height + 115
    Picture3.Top = Me.Renderer.Height + 100
    'agregamos accesos directos
     
     
    'Modificamos los parametros del engine
    SetHalfWindowTileHeight (frmMain.Renderer.ScaleHeight)
    SetHalfWindowTileWidth (frmMain.Renderer.ScaleWidth)
End Sub


Private Sub TxtMapa_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    NumMap_Save = 7
    Call MapPest_Click(TxtMapa)
End If

End Sub

