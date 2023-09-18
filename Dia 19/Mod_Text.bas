Private Sub Texto()




'Cambiar Mayusc/Minusc
'Convertir lo que cargue el usuario a may�scula
Range("A1").Value = UCase(Range("A1").Value)
'Convertir lo que cargue el usuario a min�scula
Range("A1").Value = LCase(Range("A1").Value)
'Convertir las iniciales de lo que cargue el usuario a may�scula
Range("A1").Value = Application.WorksheetFunction.Proper(Range("A1").Value)

End Sub
