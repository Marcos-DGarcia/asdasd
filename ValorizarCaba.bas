Sub valoriza_Caba()
Dim valorCaba
Dim cont
Dim ultimaFila
Dim kilos
Dim tarifaCaba
dim localidad
dim provincia

Dim valorminimo
Dim valor1
Dim valor2
Dim valor3
Dim valor4
Dim valor5

Dim valorGBA01
Dim valorGBA02
Dim valorGBA03
Dim valorGBA04
Dim valorGBA05
Dim valorGbaMinimo

Dim valorCanuelas01
Dim valorCanuelas02
Dim valorCanuelas03
Dim valorCanuelas04
Dim valorCanuelas05
Dim valorCanuelasMinimo

'CABA
valor1 = 5.02
valor2 = 4.98
valor3 = 3.23
valor4 = 2.64
valor5 = 2.11
valorminimo = 1021.38
'GRAN BUENOS AIRES
valorGBA01 = 5.09
valorGBA02 = 5.07
valorGBA03 = 3.29
valorGBA04 = 2.81
valorGBA05 = 2.04
valorGbaMinimo = 1036.48
'CAÑUELAS
valorCanuelas01 = 9.91
valorCanuelas02 = 9.87
valorCanuelas03 = 6.4
valorCanuelas04 = 5.46
valorCanuelas05 = 3.98
valorCanuelasMinimo = 2017.74

ultimaFila = Hoja27.Range("A" & Rows.Count).End(xlUp).Row
For cont = 9 To ultimaFila
    kilos = Hoja27.Cells(cont, 8)
    localidad = Hoja27.Cells(cont, 7)
    Ucase(localidad)
    provincia = Hoja27.Cells(cont, 6)
    Ucase(provincia)
    select case provincia
    case "BUENOS AIRES"
        select case localidad
        'if
             Case "HURLINGHAM" Or "LA PLATA" OR "VILLA MARTELLI"
            If kilos >= 4000 Then
                tarifaInterior = kilos * valorGBA02
            End If
            If kilos >= 204 And kilos <= 4000 Then
                tarifaInterior = kilos * valorGBA01
            End If
            If kilos < 204 Then
                tarifaInterior = valorGbaMinimo
            End if
        'if
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
        'else if         'TARIFAS DE AMBA 5.09-5.07
            case else
            If kilos >= 14000 Then
                tarifaInterior = kilos * valorGBA05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorGBA04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorGBA03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorGBA02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorGBA01
            End If
            If kilos < 204 Then
                tarifaInterior = valorGbaMinimo
            end if
        end select
        ' 5.02-4.98 ETC VALORES DE EXPRESOS
    case else
        If kilos >= 14000 Then
            tarifaCaba = kilos * valor5
        End If
        If kilos < 14000 And kilos >= 8000 Then
            tarifaCaba = kilos * valor4
        End If
        If kilos < 8000 And kilos >= 4000 Then
            tarifaCaba = kilos * valor3
        End If
        If kilos < 4000 And kilos >= 1000 Then
            tarifaCaba = kilos * valor2
        End If
        If kilos < 1000 And kilos >= 214 Then
            tarifaCaba = kilos * valor1
        End If
        If kilos < 214 Then
            tarifaCaba = valorminimo
        End If
        Hoja27.Cells(cont, 9) = tarifaCaba    
    end select
Next cont
End Sub
