Sub fondo()

' Vamos a cambiar el color de fondo de una celda

Range("A5").Interior.Color = 210

' También podemos asignar colores genéricos

Range("B7").Interior.Color = vbBlue


' Se puede utilizar la escala RGB

Range("B8").Interior.Color = RGB(150, 200, 220)

End Sub
