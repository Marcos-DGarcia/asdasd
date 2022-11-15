Sub valoriza_Caba()
Dim valorCaba
Dim cont
Dim ultimaFila
Dim kilos
Dim valorminimo
Dim valor100
Dim valor1000
Dim valor4000
Dim valor8000
Dim valor14000
Dim tarifaCaba
dim localidad
dim provincia

'CABA
valor100 = 5.02
valor1000 = 4.98
valor4000 = 3.23
valor8000 = 2.64
valor14000 = 2.11
valorminimo = 1021.38
'GRAN BUENOS AIRES
ultimaFila = Hoja27.Range("A" & Rows.Count).End(xlUp).Row
For cont = 9 To ultimaFila
    kilos = Hoja27.Cells(cont, 7)
    localidad = Hoja27.Cells(cont, 6)
    provincia = Hoja27.Cells(cont, 8)
    select case provincia
    case "BUENOS AIRES"
        select case localidad
        Case "HURLINGHAM" Or "LA PLATA" OR "VILLA MARTELLI"
            If kilos >= 4000 Then
                tarifaInterior = kilos * valorGBA02
            End If
            If kilos >= 204 And kilos <= 4000 Then
                tarifaInterior = kilos * valorGBA01
            End If
            If kilos < 204 Then
                tarifaInterior = valorGbaMinimo
            end if
        Case "CAÑUELAS"
            If kilos >= 14000 Then
                tarifaInterior = kilos * valorCanuelas0
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorCanuelas04
             End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorCanuelas03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorCanuelas02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorCanuelas01
            End If
            If kilos < 204 Then
                tarifaInterior = valorCanuelasMinimo
            End If
                    'TARIFAS DE AMBA 5.09-5.07
            case else
            If kilos >= 14000 Then
                tarifaInterior = kilos * valorCanuelas0
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorCanuelas04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorCanuelas03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorCanuelas02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorCanuelas01
            End If
            If kilos < 204 Then
                tarifaInterior = valorCanuelasMinimo
        end select
        ' 5.02-4.98 ETC VALORES DE EXPRESOS
    case else
        If kilos >= 14000 Then
            tarifaCaba = kilos * valor14000
        End If
        If kilos < 14000 And kilos >= 8000 Then
            tarifaCaba = kilos * valor8000
        End If
        If kilos < 8000 And kilos >= 4000 Then
            tarifaCaba = kilos * valor4000
        End If
        If kilos < 4000 And kilos >= 1000 Then
            tarifaCaba = kilos * valor1000
        End If
        If kilos < 1000 And kilos >= 214 Then
            tarifaCaba = kilos * valor100
        End If
        If kilos < 214 Then
            tarifaCaba = valorminimo
        End If
        Hoja27.Cells(cont, 8) = tarifaCaba    
    end select
Next cont
End Sub
