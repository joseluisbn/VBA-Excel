Sub borrar()

' Seleccionamos el rango

Range("A5").Select

' Borramos el contenido seleccionado con Selection

Selection.ClearContents

' Se puede hacer en una sola línea, ya que el objeto Range también tiene el método ClearContents

Range("A5").ClearContents
    
' Para borrar contenido y formato
    
 Range("A5").Clear

End Sub
