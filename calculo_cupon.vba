'Valuacion de Cupon

Function Cupon(Cupon, Valor_Nominal, interes, periodo)
Cupon = (Cupon * Valor_Nominal) / ((1 + interes) ^ periodo)
End Function

'Calculadora del cupon, se ponen los inputs
'Range sirve para poner el valor de la celda

Sub CalcularRentabilidadIntroduccion()
Range("B53").Select
ActiveCell.Value = InputBox("Proceda a poner el cupon semestral")

Range("B54").Select
ActiveCell.Value = InputBox(" Por favor digite la cantidad nominal")

'Para calcular el cupon
Range("B55") = Range("B53") * Range("B54")

'Para calcular el rango donde empieza
Range("A59") = Range("B53") * -1
Range("A60") = Range("B53")
Range("A61") = Range("B53")
Range("A62") = Range("B53")
Range("A63") = Range("B53")
Range("A64") = Range("B53")
Range("A65") = Range("B53")
Range("A66") = Range("B53")
Range("A67") = Range("B53")
Range("A68") = Range("B53")
Range("A69") = Range("B53")
Range("A70") = Range("B53") + Range("B54")

'Para calcular el valor del Bono
'Range("E57").Value = Application.WorksheetFunction.Sum(Range(A59:A70))
Set myRange = Woorksheet("Introduccion").Range("C21:C31")
result = Application.WorksheetFunction.IRR(myRange)
Range("E58") = result
'Range("E58")="#0.0000"

Range("I59") = (1 + Range("F24")) ^ 2 - 1
End Sub
