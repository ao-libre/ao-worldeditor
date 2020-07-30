Attribute VB_Name = "modcopiarObj"
Dim ruta As String

Sub MostrarRuta()

    MsgBox ThisWorkbook.Path

End Sub

Sub CrearDirectorio()

ruta = App.Path
carpeta = ruta & "\Planillas"
X = Dir(carpeta, vbDirectory)
    If X = "" Then
        MsgBox ("La carpeta " & carpeta & " no existe")
        MkDir ruta & "\Planillas"
    Else
        'MsgBox ("La carpeta " & carpeta & " existe")
    End If
End Sub
Sub CrearHojasObjyNPCs()

 Call MacroFileExist
 Call MacroFileExist2
 
 
End Sub

Sub AbrirHojasNPCs()

ruta = App.Path

    Workbooks.OpenText FileName:=ruta & "\Planillas\NPCs.csv", _
          DataType:=xlDelimited, Semicolon:=True
          
    Windows("objetosDat.xlsm").Activate
        For Each she In Worksheets
            If she.Name = "NPCs" Then
            'she.Delete
            Windows("NPCs.csv").Close savechanges:=False
            Exit Sub
            End If
        Next
        
        ActiveWindow.WindowState = xlNormal
        Windows("NPCs.csv").Activate
        Sheets("NPCs").Select
        Sheets("NPCs").Copy Before:=Workbooks("objetosDat.xlsm").Sheets(1)
        Windows("NPCs.csv").Close savechanges:=False
          
End Sub
Sub AbrirHojasObjetos()

ruta = App.Path

    Workbooks.OpenText FileName:=ruta & "\Planillas\Objetos.csv", _
          DataType:=xlDelimited, Semicolon:=True
    
    Windows("objetosDat.xlsm").Activate
        For Each she In Worksheets
            If she.Name = "Objetos" Then
            'she.Delete
            Windows("Objetos.csv").Close savechanges:=False
            Exit Sub
            End If
        Next
        ActiveWindow.WindowState = xlNormal
        Windows("Objetos.csv").Activate
        Sheets("Objetos").Select
        Sheets("Objetos").Copy Before:=Workbooks("objetosDat.xlsm").Sheets(1)
        Windows("Objetos.csv").Close savechanges:=False

          
End Sub

Sub ExtraerDatosOtrosLibros()
    Dim LibroDatos As Workbook
    
    Set LibroDatos = Workbook.Open(ruta & "\Planilla\Objetos.csv")
    LibroDatos.Sheets(1).Range("A6:D3000").Copy
    LibroDatos.Close savechange:=False
    Avtivesheet.Paste
    Range("A6").Select
    
End Sub
Function FileExists(FPath As String) As Boolean

Dim FName As String
FName = Dir(FPath)
If FName <> "" Then FileExists = True _
Else: FileExists = False
End Function

Sub MacroFileExist()
ruta = App.Path

If FileExists(ruta & "\Planillas\Objetos.csv") = True Then
    'MsgBox "Objetos.csv ya existe."
Else
    MsgBox "Objetos.csv no exite y sera creada."
        FileCopy ruta & "\obj.dat", ruta & "\Planillas\Objetos.csv"
        Call Reemplazar_Texto(ruta & "\Planillas" & "\objetos.csv", "=", ",")
        Windows("Objetos.csv").Close savechanges:=False
End If
Call AbrirHojasObjetos
End Sub

Sub MacroFileExist2()
ruta = App.Path

If FileExists(ruta & "\Planillas\NPCs.csv") = True Then
    'MsgBox "NPCs.csv ya existe."
Else
    MsgBox "NPCs.csv no existe y sera creado."
        FileCopy ruta & "\NPCs.dat", ruta & "\Planillas\NPCs.csv"
        Call Reemplazar_Texto(ruta & "\Planillas" & "\NPCs.csv", "=", ",")
        Call Reemplazar_Texto(ruta & "\Planillas" & "\NPCs.csv", "-", ",")
        Call Reemplazar_Texto(ruta & "\Planillas" & "\NPCs.csv", "'", ",")
        Windows("NPCs.csv").Close savechanges:=False
End If
Call AbrirHojasNPCs
End Sub
' Verifica si Existe Libro y Crea Libro

Sub CreaLibroExcel()
Dim she As Worksheet
Dim f, B, ruta, path1 As String
Dim verexi As Object
Application.StatusBar = "Creando libro y finalizando…"
f = "objetosDat.xlsm"
ruta = ActiveWorkbook.Path
path1 = ruta & "\" & f
B = "Objetos"
Set verexi = CreateObject("Scripting.FileSystemObject")
If verexi.FileExists(path1) Then
   Workbooks.Open FileName:=path1, UpdateLinks:=0
   ActiveWorkbook.Sheets.Add Before:=Worksheets(1)
   For Each she In Worksheets
   If she.Name = B Then she.Delete
   Next
   ActiveSheet.Name = B
Else
Workbooks.Add
ActiveWorkbook.SaveAs FileName:=path1
ActiveWorkbook.Sheets.Add Before:=Worksheets(1)
ActiveSheet.Name = B
End If
Application.StatusBar = Clear
End Sub

' compruebo que NPCsIndex este Abierta
Sub AbrirHojasNPCsIndex()

ruta = App.Path
          
    Windows("objetosDat.xlsm").Activate
        For Each she In Worksheets
            If she.Name = "NPCsIndex" Then
                'she.Delete
                Exit Sub
            End If
        Next
              
        MsgBox ("NO existe")
        Call modNPCs.CrearMasterNPCs
End Sub
' compruebo que NPCsIndex este Abierta
Sub AbrirHojasObjIndex()

ruta = App.Path
          
    Windows("objetosDat.xlsm").Activate
        For Each she In Worksheets
            If she.Name = "ObjIndex" Then
                'she.Delete
                Exit Sub
            End If
        Next
              
        MsgBox ("NO existe")
        Call UserForm1.CrearMasterDatos
End Sub


'**************************************************************
'* Leemos los datos de objtos.dat en un exel y los pasamos a una hoja master
'* creadopor ReyarB
'* ultima modificacion 30/05/2020
'***************************************************************

Sub CrearMasterDatos()

    Call modcopiarObj.CrearDirectorio
    Call modcopiarObj.CrearHojasObjyNPCs

    Call CrearHoja("ObjIndex")
    
    Range("A5:BE5").HorizontalAlignment = xlCenter
    Range("A5:BE5").VerticalAlignment = xlCenter
    Range("A5:BE5").RowHeight = 25
    Range("A5:BE5").Borders.Weight = XlBorderWeight.xlThick
    Range("A2").ColumnWidth = 20
    Range("B2").ColumnWidth = 40
    Range("C8:D2").ColumnWidth = 8
    Range("E2").ColumnWidth = 10
            
    palabraBusqueda = "*" & "OBJ" & "*"
    ultimaFila = Sheets("Objetos").Range("B" & Rows.count).End(xlUp).Row
    
    If ultimaFila < 6 Then
        Exit Sub
    End If
    
    For Cont = 6 To ultimaFila
        For Y = 1 To 58
            
            If Sheets("Objetos").Cells(Cont, 1) = "" Then Exit For
            
            If Sheets("Objetos").Cells(Cont, 1) Like palabraBusqueda Then Obj = Sheets("Objetos").Cells(Cont, 1)
            If Sheets("Objetos").Cells(Cont, 1) = "Name" Then Name = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "GrhIndex" Then GrhIndex = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "ObjType" Then ObjType = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "NumRopaje" Then NumRopaje = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Agarrable" Then Agarrable = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Caos" Then Caos = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "MaxDef" Then MaxDef = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "MinDef" Then MinDef = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Valor" Then Valor = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Crucial" Then Crucial = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "CP1" Then CP1 = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "CP2" Then CP2 = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "CP3" Then CP3 = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "CP4" Then CP4 = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "CP5" Then CP5 = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "CP6" Then CP6 = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "CP7" Then CP7 = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "CP8" Then CP8 = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "CP9" Then CP9 = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "CP10" Then CP10 = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "CP11" Then CP11 = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "LingH" Then LingH = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "LingP" Then LingP = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "LingO" Then LingO = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "SkHerreria" Then SkHerreria = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "IndexAbierta" Then IndexAbierta = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "IndexCerrada" Then IndexCerrada = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "IndexCerradaLlave" Then IndexCerradaLlave = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Llave" Then Llave = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "CantItems" Then CantItems = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "NroItems" Then NROITEMS = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Texto" Then Texto = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "VGrande" Then VGrande = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Anim" Then Anim = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "RazaEnanaAnim" Then RazaEnanaAnim = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "StaffPower" Then StaffPower = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "StaffDamageBonus" Then StaffDamageBonus = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Proyectil" Then Proyectil = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Municiones" Then Municiones = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Apunala" Then Apunala = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Envenena" Then Envenena = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Acuchilla" Then Acuchilla = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Madera" Then Madera = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "MaderaElfica" Then MaderaElfica = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "SkCarpinteria" Then SkCarpinteria = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Log" Then Log = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "NoRobable" Then NoRobable = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "Upgrade" Then Upgrade = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "MinHit" Then MinHIT = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "MaxHit" Then MaxHIT = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "HechizoIndex" Then HechizoIndex = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "TipoPocion" Then TipoPocion = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "MaxModificador" Then MaxModificador = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "MinModificador" Then MinModificador = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "DuracionEfecto" Then DuracionEfecto = Sheets("Objetos").Cells(Cont, 2)
            If Sheets("Objetos").Cells(Cont, 1) = "NoSeCae" Then NoSeCae = Sheets("Objetos").Cells(Cont, 2)
            

             Cont = Cont + 1
            
         Next Y

        ultimaFilaAuxiliar = Sheets("ObjIndex").Range("A" & Rows.count).End(xlUp).Row
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 1) = Obj
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 2) = Name
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 3) = GrhIndex
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 4) = ObjType
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 5) = NumRopaje
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 6) = Agarrable
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 7) = Caos
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 8) = MaxDef
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 9) = MinDef
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 10) = Valor
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 11) = Crucial
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 12) = CP1
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 13) = CP2
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 14) = CP3
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 15) = CP4
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 16) = CP5
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 17) = CP6
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 18) = CP7
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 19) = CP8
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 20) = CP9
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 21) = CP10
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 22) = CP11
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 23) = LingH
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 24) = LingP
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 25) = LingO
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 26) = SkHerreria
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 27) = IndexAbierta
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 28) = IndexCerrada
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 29) = IndexCerradaLlave
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 30) = Llave
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 31) = CantItems
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 32) = NROITEMS
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 33) = Texto
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 34) = VGrande
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 35) = Anim
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 36) = RazaEnanaAnim
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 37) = StaffPower
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 38) = StaffDamageBonus
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 39) = Proyectil
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 40) = Municiones
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 41) = Apunala
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 42) = Envenena
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 43) = Acuchilla
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 44) = Madera
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 45) = MaderaElfica
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 46) = SkCarpinteria
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 47) = Log
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 48) = NoRobable
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 49) = Upgrade
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 50) = MinHIT
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 51) = MaxHIT
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 52) = HechizoIndex
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 53) = TipoPocion
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 54) = MaxModificador
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 55) = MinModificador
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 56) = DuracionEfecto
        Sheets("ObjIndex").Cells(ultimaFilaAuxiliar + 1, 57) = NoSeCae

        
        
        
        Obj = ""
        Name = ""
        GrhIndex = ""
        ObjType = ""
        NumRopaje = ""
        Agarrable = ""
        Caos = ""
        MaxDef = ""
        MinDef = ""
        Valor = ""
        Crucial = ""
        CP1 = ""
        CP2 = ""
        CP3 = ""
        CP4 = ""
        CP5 = ""
        CP6 = ""
        CP7 = ""
        CP8 = ""
        CP9 = ""
        CP10 = ""
        CP11 = ""
        LingH = ""
        LingP = ""
        LingO = ""
        SkHerreria = ""
        IndexAbierta = ""
        IndexCerrada = ""
        IndexCerradaLlave = ""
        Llave = ""
        CantItems = ""
        NROITEMS = ""
        Texto = ""
        VGrande = ""
        Anim = ""
        RazaEnanaAnim = ""
        StaffPower = ""
        StaffDamageBonus = ""
        Proyectil = ""
        Municiones = ""
        Apunala = ""
        Envenena = ""
        Acuchilla = ""
        Madera = ""
        MaderaElfica = ""
        SkCarpinteria = ""
        Log = ""
        NoRobable = ""
        Upgrade = ""
        MinHIT = ""
        MaxHIT = ""
        HechizoIndex = ""
        TipoPocion = ""
        MaxModificador = ""
        MinModificador = ""
        DuracionEfecto = ""
        NoSeCae = ""
        
    Next Cont
    
        ultimaFilaAuxiliar = Sheets("ObjIndex").Range("A" & Rows.count).End(xlUp).Row
        With Sheets("ObjIndex").Range("A6:BE" & ultimaFilaAuxiliar).Font
        .Name = "arial"
        .Size = 9
        .italic = True
        
        MsgBox "Proceso Terminado. Master de ObjIndex", vbInformation, "Resultado"
    End With

End Sub


Sub Reemplazar_Texto(ByVal El_Archivo As String, _
                            ByVal La_cadena As String, _
                            ByVal Nueva_Cadena As String)
  
On Error GoTo errSub
Dim f As Integer
Dim Contenido As String
  
      
    f = FreeFile
      
    'Abre el archivo para leer los datos
    Open El_Archivo For Input As f
      
    'carga el contenido del archivo en la variable
    Contenido = Input$(LOF(f), #f)
      
    'Cierra el archivo
    Close #f
      
    ' Ejecuta la función Replace, pasandole los datos
    Contenido = Replace(Contenido, La_cadena, Nueva_Cadena)
  
      
    f = FreeFile
    'Abre un nuevo archivo
    Open El_Archivo For Output As f
    'Graba los nuevos datos
    Print #f, Contenido
      
    'cierra el archivo
    Close #f
      
    'MsgBox " Archivo modificado ", vbInformation
Exit Sub
  
'Error
  
errSub:
MsgBox err.Description, vbCritical
Close
End Sub
