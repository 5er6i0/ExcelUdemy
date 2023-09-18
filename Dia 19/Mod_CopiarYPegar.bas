Private Sub ProbandoEsto()
'Pega porciones de codigo aqui para probarlos



End Sub

Private Sub CopiandoYPegando()

'************ COPIAR Y PEGAR SIMPLE ************

'Copiar y pegar simple en la misma hoja

'Rango origen	Rango Destino
Range("A1").Copy Range("B1")
'Rango origen	Rango Destino
Range("A1:A3").Copy Range("B1:B3")
'M�todo alternativo
Range("A1:A3").Copy Range("B1")


'Copiar y pegar simple en otra hoja (del mismo libro)

'Hoja origen 	rango origen           hoja destino	rango destino
Worksheets("Hoja1").Range("A1").Copy Worksheets("Hoja2").Range("A1")


'Copiar y pegar simple en otra hoja (de otro libro)

'Especificar el nombre del libro, de la hoja, y el rango, tanto del origen como del destino
Workbooks("Libro1.xlsx").Worksheets("Hoja1").Range("A1").Copy _
        Workbooks("Mi-Libro-de-Macros.xlsm").Worksheets("Hoja2").Range("B1")
        
'Como antes pero usando la propiedad "Current Region" (region actual)
Workbooks("Libro1.xlsx").Worksheets("Hoja1").Range("A1").CurrentRegion.Copy _
        Workbooks("Mi-Libro-de-Macros.xlsm").Worksheets("Hoja2").Range("B1")





'************ COPIAR Y PEGAR ESPECIAL ************

'NOTE: copiar y pegar individual NO produce los guiones que giran alrededor de la seleccion, 
'Pero copiar y pegar lineas separada SI los produce. Eliminalos con esto:
Application.CutCopyMode = False


'Copiar y pegar el FORMATO con PasteSpecial (Pegado especial)
'Origen
Range("A1").Copy
'Destino
Range("B1").PasteSpecial xlPasteFormats
'Eliminar guiones giratorios
Application.CutCopyMode = False


'Copiar y pegar el VALOR con PasteSpecial (Pegado especial)
'Origen
Range("A1").Copy
'Destino
Range("B1").PasteSpecial xlPasteValues
'Eliminar guiones giratorios
Application.CutCopyMode = False



End Sub
