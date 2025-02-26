' MERCADO DE CAPITALES

' MODULO DE FORMULAS...

' FORMULA PARA EL VFAV...

Function unam_vfav(ahorro, tasa, años, capitalizaciones)
unam_vfav = ahorro * (((1 + tasa / capitalizaciones) ^ (capitalizaciones * años) - 1) / (tasa / capitalizaciones))
End Function

' FORMULA PARA EL VFAV CRECIENTE EN 2 PERIODOS...

Function unam_vfavc(pago1, pago2, n1, n2, i, m)
n = m * años

anu = (((1 + (i / m)) ^ (m * n1)) - 1) / (i / m)
anualidad1 = pago1 * anu
Anualidad_cap = anualidad1 * (1 + i / m) ^ (n2)
anualidad2 = pago2 * anu
unam_vfavc = Anualidad_cap + anualidad2


'Interes compuesto formula
Function unam_int_comp(capital, tasa, capitalizacion, periodo)
interes_compuesto = capital * ((1 + (tasa / capitalizacion)) ^ (capitalizacion * periodo))

'Tasa de interes formula
Function rate(capital, monto, capitalizacion, periodo)
rate = ((monto / capital) ^ (1 / (capitalizacion * periodo)) - 1) * m

'funcion de las tasas
Function unam_tasa(capital, monto, n, m)

unam_tasa = ((monto / capital) ^ (1 / (m * n)) - 1) * m

End Function


End Function

Sub anualidad1()

End Sub
