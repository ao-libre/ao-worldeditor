Attribute VB_Name = "modCostas"
Option Explicit

Public Sub PutCoast()
 
    Dim aux As Integer
    Dim X As Integer
    Dim Y As Integer
    On Error Resume Next
     
    Static antAux(4) As Byte 'Para el movimiento del grh _
                        (como las costas de pasto estan compuestas por 2 grhs c/u en forma de : _
                        1 0 _
                        1 0 _
                        donde 1 es la costa y el 0 es el grh en negro (que no nos interesa)
     
    'If Not HayAgua(x, y) Then Exit Sub 'Se supone que si ponemos costas, debe haber agua en la posicion de búsqueda.
     
    Call AddtoRichTextBox(frmMain.StatTxt, "Colocando costas", 0, 255, 0)
     
    'Buscamos si hay alrededor, si es asi, salimos.
    For Y = YMinMapSize To YMaxMapSize
     
        If Y > YMinMapSize And Y < YMaxMapSize Then
        
        For X = XMinMapSize To XMaxMapSize
            If X > XMinMapSize And X < XMaxMapSize Then
            
            
            If HayAgua2(X - 1, Y) And HayAgua2(X - 1, Y - 1) And HayAgua2(X, Y - 1) And HayAgua2(X + 1, Y - 1) _
                And HayAgua2(X + 1, Y) And HayAgua2(X + 1, Y + 1) And HayAgua2(X, Y + 1) And HayAgua2(X - 1, Y + 1) Then
                    Debug.Print "No se pone costa :D"
                    'Exit Sub
            End If
            
        
            'Buscamos a la izquierda
            
            If Not HayAgua2(X - 1, Y) And MapData(X - 1, Y).Graphic(2).GrhIndex = 0 Then
                If HayAgua2(X, Y) Then
                    If MapData(X + 1, Y).Graphic(2).GrhIndex > 0 Then
                        MapData(X + 1, Y).Graphic(2).GrhIndex = 0
                    End If
                    
                        aux = IIf(antAux(0) > 0, 7309, 7307)
                        antAux(0) = Not antAux(0)
                        InitGrh MapData(X, Y).Graphic(2), aux
                End If
                'Exit Sub
            End If
            
            'Buscamos a la derecha
            If Not HayAgua2(X + 1, Y) And MapData(X + 1, Y).Graphic(2).GrhIndex = 0 Then
                If HayAgua2(X, Y) Then
                    If MapData(X - 1, Y).Graphic(2).GrhIndex > 0 Then
                        MapData(X - 1, Y).Graphic(2).GrhIndex = 0
                    End If
                        
                            aux = IIf(antAux(1) > 0, 7322, 7320)
                            antAux(1) = Not antAux(1)
                            InitGrh MapData(X, Y).Graphic(2), aux
                End If
                'Exit Sub
            End If
     
            'Buscamos a la Arriba
            If Not HayAgua2(X, Y - 1) And MapData(X, Y - 1).Graphic(2).GrhIndex = 0 Then
                If HayAgua2(X, Y) Then
                    If MapData(X, Y + 1).Graphic(2).GrhIndex > 0 Then
                        MapData(X, Y + 1).Graphic(2).GrhIndex = 0
                    End If
                        
                            aux = IIf(antAux(2) > 0, 7324, 7323)
                            antAux(2) = Not antAux(2)
                            InitGrh MapData(X, Y).Graphic(2), aux
                End If
                'Exit Sub
            End If
     
            'Buscamos a la Abajo
            If Not HayAgua2(X, Y + 1) And MapData(X, Y + 1).Graphic(2).GrhIndex = 0 Then
                If HayAgua2(X, Y) Then
                    If MapData(X, Y - 1).Graphic(2).GrhIndex > 0 Then
                        MapData(X, Y - 1).Graphic(2).GrhIndex = 0
                    End If
                        
                            aux = IIf(antAux(3) > 0, 7330, 7329)
                            antAux(3) = Not antAux(3)
                            InitGrh MapData(X, Y).Graphic(2), aux
                End If
                'Exit Sub
            End If

            'Buscamos los casos especificos (Corners)******************************
            'Arriba Izquierda
            
            If Not HayAgua2(X - 1, Y) And Not HayAgua2(X, Y - 1) And HayAgua2(X, Y) And HayAgua2(X + 1, Y) And HayAgua2(X, Y + 1) Then
                    MapData(X, Y + 1).Graphic(2).GrhIndex = 7289
                    MapData(X, Y).Graphic(2).GrhIndex = 7287
                    MapData(X + 1, Y).Graphic(2).GrhIndex = 7288
            End If
            
            'Arriba Derecha
            If Not HayAgua2(X + 1, Y) And Not HayAgua2(X, Y - 1) And HayAgua2(X, Y) And HayAgua2(X - 1, Y) And HayAgua2(X, Y + 1) Then
                    MapData(X, Y + 1).Graphic(2).GrhIndex = 7298
                    MapData(X, Y).Graphic(2).GrhIndex = 7296
                    MapData(X - 1, Y).Graphic(2).GrhIndex = 7295
            End If
     
            'Abajo Izquierda
            If Not HayAgua2(X - 1, Y) And Not HayAgua2(X, Y + 1) And HayAgua2(X, Y) And HayAgua2(X + 1, Y) And HayAgua2(X, Y - 1) Then
                    MapData(X, Y - 1).Graphic(2).GrhIndex = 7283
                    MapData(X, Y).Graphic(2).GrhIndex = 7285
                    MapData(X + 1, Y).Graphic(2).GrhIndex = 7286
            End If
     
            'Abajo Derecha
            If Not HayAgua2(X + 1, Y) And Not HayAgua2(X, Y + 1) And HayAgua2(X, Y) And HayAgua2(X - 1, Y) And HayAgua2(X, Y - 1) Then
                    MapData(X, Y - 1).Graphic(2).GrhIndex = 7292
                    MapData(X, Y).Graphic(2).GrhIndex = 7294
                    MapData(X - 1, Y).Graphic(2).GrhIndex = 7293
            End If
            
            End If
        Next

        End If
        
    Next
     
    For Y = 2 To YMaxMapSize
        For X = 2 To XMaxMapSize
            '**Corners***
            
            'Arriba Izquierda
            If CostaDerecha(X, Y + 1) And CostaAbajo(X + 1, Y) And HayAgua2(X, Y) Then MapData(X, Y).Graphic(2).GrhIndex = 7318
     
            'Arriba Derecha
            If CostaIzquierda(X, Y + 1) And CostaAbajo(X - 1, Y) And HayAgua2(X, Y) Then MapData(X, Y).Graphic(2).GrhIndex = 7305
     
            'Abajo Izquierda
            If CostaDerecha(X, Y - 1) And CostaArriba(X + 1, Y) Then MapData(X, Y).Graphic(2).GrhIndex = 7312
     
            'Abajo Derecha
            If CostaIzquierda(X, Y - 1) And CostaArriba(X - 1, Y) Then MapData(X, Y).Graphic(2).GrhIndex = 7299
        Next
    Next
    Call AddtoRichTextBox(frmMain.StatTxt, "Limpiando costas", 0, 255, 0)
    Call LimpiarCostas
    
    Call AddtoRichTextBox(frmMain.StatTxt, "Costas colocadas", 0, 255, 0)
End Sub
 
Private Function CostaIzquierda(ByVal X As Integer, ByVal Y As Integer) As Boolean
    CostaIzquierda = ((MapData(X, Y).Graphic(2).GrhIndex = 7307) Or (MapData(X, Y).Graphic(2).GrhIndex = 7309))
End Function
 
Private Function CostaDerecha(ByVal X As Integer, ByVal Y As Integer) As Boolean
    CostaDerecha = ((MapData(X, Y).Graphic(2).GrhIndex = 7322) Or (MapData(X, Y).Graphic(2).GrhIndex = 7320))
End Function
 
Private Function CostaArriba(ByVal X As Integer, ByVal Y As Integer) As Boolean
    CostaArriba = ((MapData(X, Y).Graphic(2).GrhIndex = 7324) Or (MapData(X, Y).Graphic(2).GrhIndex = 7323))
End Function
 
Private Function CostaAbajo(ByVal X As Integer, ByVal Y As Integer) As Boolean
    CostaAbajo = ((MapData(X, Y).Graphic(2).GrhIndex = 7330) Or (MapData(X, Y).Graphic(2).GrhIndex = 7329))
End Function
 
Function HayAgua2(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua2 = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562))
End Function

Public Sub LimpiarCostas()
Dim Y As Integer
Dim X As Integer

    'Buscamos si hay alrededor, si es asi, salimos.
    For Y = YMinMapSize To YMaxMapSize
        If Y > YMinMapSize And Y < YMaxMapSize Then
        
            For X = XMinMapSize To XMaxMapSize
                If X > XMinMapSize And X < XMaxMapSize Then
                    MapData(X, Y).Graphic(2).GrhIndex = 7284
                    MapData(X, Y).Graphic(2).GrhIndex = 7290
                    MapData(X, Y).Graphic(2).GrhIndex = 7291
                    MapData(X, Y).Graphic(2).GrhIndex = 7297
                    MapData(X, Y).Graphic(2).GrhIndex = 7300
                    MapData(X, Y).Graphic(2).GrhIndex = 7301
                    MapData(X, Y).Graphic(2).GrhIndex = 7302
                    MapData(X, Y).Graphic(2).GrhIndex = 7303
                    MapData(X, Y).Graphic(2).GrhIndex = 7304
                    MapData(X, Y).Graphic(2).GrhIndex = 7306
                    MapData(X, Y).Graphic(2).GrhIndex = 7308
                    MapData(X, Y).Graphic(2).GrhIndex = 7310
                    MapData(X, Y).Graphic(2).GrhIndex = 7311
                    MapData(X, Y).Graphic(2).GrhIndex = 7313
                    MapData(X, Y).Graphic(2).GrhIndex = 7314
                    
                    MapData(X, Y).Graphic(2).GrhIndex = 7316
                    MapData(X, Y).Graphic(2).GrhIndex = 7315
                    MapData(X, Y).Graphic(2).GrhIndex = 7317
                    MapData(X, Y).Graphic(2).GrhIndex = 7319
                    MapData(X, Y).Graphic(2).GrhIndex = 7321
                    MapData(X, Y).Graphic(2).GrhIndex = 7325
                    MapData(X, Y).Graphic(2).GrhIndex = 7326
                    MapData(X, Y).Graphic(2).GrhIndex = 7327
                    MapData(X, Y).Graphic(2).GrhIndex = 7328
                End If
            Next X
        End If
    Next Y
            
End Sub
