Sub borrar()

' Seleccionamos el rango

Range("A5").Select

' Borramos el contenido seleccionado

Selection.ClearContents

' Se puede hacer en una sola línea, ya que el objeto Range también tiene el método ClearContents

Range("A5").ClearContents

End Sub
