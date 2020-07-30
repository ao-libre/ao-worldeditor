VERSION 5.00
Begin VB.Form frmConfigSup 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Configuración Avanzada de Superficie "
   ClientHeight    =   1395
   ClientLeft      =   10845
   ClientTop       =   12255
   ClientWidth     =   3915
   Icon            =   "frmConfigSup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
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
      Height          =   480
      Index           =   3
      Left            =   2640
      Picture         =   "frmConfigSup.frx":628A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   360
      Width           =   480
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
      Height          =   480
      Index           =   2
      Left            =   2640
      Picture         =   "frmConfigSup.frx":9F79
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   480
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
      Height          =   480
      Index           =   1
      Left            =   2160
      Picture         =   "frmConfigSup.frx":DD01
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   480
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
      Height          =   480
      Index           =   0
      Left            =   3120
      Picture         =   "frmConfigSup.frx":11A1C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   480
   End
   Begin VB.TextBox DMLargo 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   1560
      TabIndex        =   7
      Text            =   "0"
      Top             =   840
      Width           =   420
   End
   Begin VB.TextBox DMAncho 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   1560
      TabIndex        =   6
      Text            =   "0"
      Top             =   480
      Width           =   420
   End
   Begin VB.CheckBox DespMosaic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      Caption         =   "Desplazamiento de Mosaico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      UseMaskColor    =   -1  'True
      Value           =   1  'Checked
      Width           =   2880
   End
   Begin VB.TextBox mAncho 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   960
      TabIndex        =   1
      Text            =   "4"
      Top             =   480
      Width           =   345
   End
   Begin VB.TextBox mLargo 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   960
      TabIndex        =   0
      Text            =   "4"
      Top             =   840
      Width           =   345
   End
   Begin VB.CheckBox MOSAICO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      Caption         =   "Mosaico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   1335
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   525
   End
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   480
   End
End
Attribute VB_Name = "frmConfigSup"
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

Private Sub cmdDM_Click(index As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

On Error Resume Next
Select Case index
        Case 0
            DMAncho.Text = Str(Val(DMAncho.Text) + 1)
        Case 1
            DMAncho.Text = Str(Val(DMAncho.Text) - 1)
        Case 2
            DMLargo.Text = Str(Val(DMLargo.Text) - 1)
        Case 3
            DMLargo.Text = Str(Val(DMLargo.Text) + 1)
End Select
End Sub

Private Sub Form_Deactivate()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

If UnloadMode <> 0 Then
'    Cancel = True
    Me.Hide
End If
End Sub

Private Sub DespMosaic_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
If LenB(DMAncho.Text) = 0 Then DMAncho.Text = "0"
If LenB(DMLargo.Text) = 0 Then DMLargo.Text = "0"
End Sub


Private Sub mAncho_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
' Impedir que se ingrese un valor no numerico
If KeyAscii <> 8 And IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub

Private Sub mLargo_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
' Impedir que se ingrese un valor no numerico
If KeyAscii <> 8 And IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub

Private Sub cmdAceptar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Me.Hide
End Sub

Private Sub MOSAICO_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
If LenB(mAncho.Text) = 0 Then mAncho.Text = "0"
If LenB(mLargo.Text) = 0 Then mLargo.Text = "0"
End Sub
