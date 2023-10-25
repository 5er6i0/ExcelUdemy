Sub PintarArcoiris()
' Pintará el color de fondo de las celdas que van desde A1 hasta A7
' con los siguientes colores respectivamente:
' Rojo (Red), Rosa (Magenta), Amarillo (Yellow),
' Verde (Green), Cian (Cyan), Azul (blue), y Negro (Black)
    ' Seleccionar la hoja
    ActiveSheet.Select
        
    ' Pintar las celdas
    Range("A1").Interior.Color = vbRed
    Range("A2").Interior.Color = vbMagenta
    Range("A3").Interior.Color = vbYellow
    Range("A4").Interior.Color = vbGreen
    Range("A5").Interior.Color = vbCyan
    Range("A6").Interior.Color = vbBlue
    Range("A7").Interior.Color = vbBlack

End Sub

Sub EscribirArcoiris()
' Escribirá en cada celda pintada,
' el nombre del color correspondiente (ayuda: usa la propiedad Value)
    
    ' Seleccionar la hoja
    ActiveSheet.Select
    
    ' Insetar texto
    Range("A1").Value = "Rojo"
    Range("A2").Value = "Magenta"
    Range("A3").Value = "Amarillo"
    Range("A4").Value = "Verde"
    Range("A5").Value = "Cyan"
    Range("A6").Value = "Blue"
    Range("A7").Value = "Negro"

End Sub

Sub ColorearFuenteArcoiris()
' Asignará a todas las celdas escritas,
' el color de fuente blanco (White)
    
    ' Seleccionamos la hoja
    ActiveSheet.Select
    
    ' Asignamos el color de fuente
    Range("A1:A7").Font.Color = vbWhite
    
End Sub

'*************************
'**Respuestas del curso**
'*************************

Sub PintarArcoiris()

Range("A1").Select
Range("A1").Interior.Color = vbRed
ActiveCell.Offset(1, 0).Interior.Color = vbMagenta
ActiveCell.Offset(2, 0).Interior.Color = vbYellow
ActiveCell.Offset(3, 0).Interior.Color = vbGreen
ActiveCell.Offset(4, 0).Interior.Color = vbCyan
ActiveCell.Offset(5, 0).Interior.Color = vbBlue
ActiveCell.Offset(6, 0).Interior.Color = vbBlack

End Sub

Sub EscribirArcoiris()

Range("A1").Select
Range("A1").Value = "Red"
ActiveCell.Offset(1, 0).Value = "Magenta"
ActiveCell.Offset(2, 0).Value = "Yellow"
ActiveCell.Offset(3, 0).Value = "Green"
ActiveCell.Offset(4, 0).Value = "Cyan"
ActiveCell.Offset(5, 0).Value = "Blue"
ActiveCell.Offset(6, 0).Value = "Black"

End Sub

Sub colorearfuentearcoiris()

Range("A1").Select
Range("A1").Font.Color = vbWhite
ActiveCell.Offset(1, 0).Font.Color = vbWhite
ActiveCell.Offset(2, 0).Font.Color = vbWhite
ActiveCell.Offset(3, 0).Font.Color = vbWhite
ActiveCell.Offset(4, 0).Font.Color = vbWhite
ActiveCell.Offset(5, 0).Font.Color = vbWhite
ActiveCell.Offset(6, 0).Font.Color = vbWhite

End Sub
