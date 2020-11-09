Sub holamundo()

' Esto es un comentario. Los comentarios son omitidos durante la ejecución.

' Vamos a introducir el valor "Hola, mundo" en la celda A1 siguiendo la estructura jerárquica de objetos en Excel.

Application.Workbooks("Libro1").Worksheets("Hoja1").Range("A1").Value = "Hola, mundo"

' Ahora vamos a introducir un valor en el libro activo.
' Si estamos trabajando en un libro activo, podemos omitir toda la estructura jerárquica.
' Para el objeto Range podemos omitir "value", ya que es su propiedad por defecto.

Range("A5") = "Hola, mundo"

End Sub
