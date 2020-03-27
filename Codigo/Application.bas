Attribute VB_Name = "Application"
Option Explicit
'**********************************************************************************************************************************************************
'En este modulo vamos a meter TODAS las declaraciones de API y funciones que tengan que ver excluxivamente con la interaccion entre Windows y nuestra app.
'**********************************************************************************************************************************************************

'***************************************
'Para obetener memoria libre en la RAM
'***************************************
Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS

    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long

End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

'***********************************************************
' Obtener informacion del adaptador de video y resolucion.
'***********************************************************

Public Type typDevMODE

    dmDeviceName       As String * 32
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * 32
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long

End Type
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

'*********************************************************************
'Funciones que manejan la memoria
'*********************************************************************

Public Function General_Bytes_To_Megabytes(Bytes As Double) As Double

    Dim dblAns As Double
        dblAns = (Bytes / 1024) / 1024
        
    General_Bytes_To_Megabytes = Format(dblAns, "###,###,##0.00")

End Function

Public Function General_GetFreeRam() As Double

    Call GlobalMemoryStatus(pUdtMemStatus)
    
    'Return Value in Megabytes
    Dim dblAns As Double
        dblAns = pUdtMemStatus.dwAvailPhys
    
    General_GetFreeRam = General_Bytes_To_Megabytes(dblAns)

End Function

Public Function General_GetFreeRam_Bytes() As Long
    
    Call GlobalMemoryStatus(pUdtMemStatus)
    
    General_GetFreeRam_Bytes = pUdtMemStatus.dwAvailPhys

End Function

