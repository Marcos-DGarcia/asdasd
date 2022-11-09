Sub valorizar_int()
    Dim localidad
    Dim provincia
    '
    Dim valorGBA01
    Dim valorGBA02
    Dim valorGBA03
    Dim valorGBA04
    Dim valorGBA05
    Dim valorGbaMinimo
    '
    Dim valorCanuelas01
    Dim valorCanuelas02
    Dim valorCanuelas03
    Dim valorCanuelas04
    Dim valorCanuelas05
    Dim valorCanuelasMinimo
    '
    Dim valor9deJulio01
    Dim valor9deJulio02
    Dim valor9deJulio03
    Dim valor9deJulio04
    Dim valor9deJulio05
    Dim valor9deJulioMinimo
    '
    Dim valorCACHARI01
    Dim valorCACHARI02
    Dim valorCACHARI03
    Dim valorCACHARI04
    Dim valorCACHARI05
    Dim valorCACHARIMinimo
    '
    dim valorCarhue01 
    dim valorCarhue02 
    dim valorCarhue03 
    dim valorCarhue04 
    dim valorCarhue05 
    dim valorCarhueMinimo 
    '
    dim valorBrandsen01
    dim valorBrandsen02
    dim valorBrandsen03
    dim valorBrandsen04
    dim valorBrandsen05
    dim valorBrandsenMinimo
    '
    dim valorJunin01
    dim valorJunin02
    dim valorJunin03
    dim valorJunin04
    dim valorJunin05
    dim valorJuninMinimo
    '
    dim valorAzul01
    dim valorAzul02
    dim valorAzul03
    dim valorAzul04
    dim valorAzul05
    dim valorAzulMinimo
    '
    dim valorBahia01
    dim valorBahia02  
    dim valorBahia03 
    dim valorBahia04  
    dim valorBahia05 
    dim valorBahiaMinimo 
    '
    dim valorBalcarse01
    dim valorBalcarse02  
    dim valorBalcarse03 
    dim valorBalcarse04  
    dim valorBalcarse05 
    dim valorBalcarseMinimo
    '
    dim valorMardelPlata01
    dim valorMardelPlata02  
    dim valorMardelPlata03 
    dim valorMardelPlata04  
    dim valorMardelPlata05 
    dim valorMardelPlataMinimo
    '
    dim valorSanNicolas01
    dim valorSanNicolas02  
    dim valorSanNicolas03 
    dim valorSanNicolas04  
    dim valorSanNicolas05 
    dim valorSanNicolasMinimo  
    '
    dim valorTandil01
    dim valorTandil02  
    dim valorTandil03 
    dim valorTandil04  
    dim valorTandil05 
    dim valorTandilMinimo  
    '
    Dim cont
    Dim ultimaFila
    Dim kilos
    Dim tarifaInterior

    'BUENOS AIRES
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
    '9 DE JULIO, CARMEN DE ARECO, CHIVILICOY
    valor9deJulio01 = 21.93
    valor9deJulio02 = 21.83
    valor9deJulio03 = 21.79
    valor9deJulio04 = 21.97
    valor9deJulio05 = 16.68
    valor9deJulioMinimo = 2334.13
    'CACHARI, ARRECIFES, GENERAL VILLEGAS, LA VIOLETA, PERGAMINO, TRENQUE LAUQUEN,
    valorCACHARI01 = 25.52
    valorCACHARI02 = 25.4
    valorCACHARI03 = 25.35
    valorCACHARI04 = 25.56
    valorCACHARI05 = 18.7
    valorCACHARIMinimo = 2715.39
    'CARHUE,CORONEL PRINGLES, CORONEL SUEREZ, PEDRO LURO, TRES ARROYOS,
    valorCarhue01 = 25.48
    valorCarhue02 = 25.37
    valorCarhue03 = 25.32
    valorCarhue04 = 25.53
    valorCarhue05 = 17.62
    valorCarhueMinimo = 2711.97
    'BRANDSEN, GENERAL BELGRANO
    valorBrandsen01 = 18.96
    valorBrandsen02 = 18.87
    valorBrandsen03 = 17.67
    valorBrandsen04 = 16.51
    valorBrandsen05 = 13.11
    valorBrandsenMinimo = 2017.74
    'JUNIN, FERRE
    valorJunin01 = 16.60
    valorJunin02 = 16.53
    valorJunin03 = 16.50
    valorJunin04 = 16.64
    valorJunin05 = 11.48
    valorJuninMinimo = 1767.03
    'AZUL
    valorAzul01 = 20.51
    valorAzul02 = 20.41
    valorAzul03 = 19.11
    valorAzul04 = 17.85
    valorAzul05 = 14.18
    valorAzulMinimo = 2182.15
    'BAHIA BLANCA
    valorBahia01 = 24.50
    valorBahia02 = 24.39
    valorBahia03 = 24.35
    valorBahia04 = 24.55
    valorBahia05 = 16.94
    valorBahiaMinimo = 2607.55
    'BALCARSE
    valorBalcarse01 = 17.92
    valorBalcarse02 = 17.84
    valorBalcarse03 = 16.70
    valorBalcarse04 = 15.60
    valorBalcarse05 = 12.39
    valorBalcarseMinimo = 1906.89
    'MAR DEL PLATA
    valorMardelPlata01 = 16.96 
    valorMardelPlata02  = 16.88
    valorMardelPlata03 = 15.80
    valorMardelPlata04  = 14.76
    valorMardelPlata05 = 11.73
    valorMardelPlataMinimo= 1804.63
    'SAN NICOLAS
    valorSanNicolas01 = 15.64
    valorSanNicolas02 = 15.57
    valorSanNicolas03 = 14.58
    valorSanNicolas04 = 13.62
    valorSanNicolas05 = 10.82
    valorSanNicolasMinimo = 1664.54
    'TANDIL
    valorTandil01 = 19.10
    valorTandil02 = 19.02
    valorTandil03 = 17.80
    valorTandil04 = 16.63
    valorTandil05 = 13.21
    valorTandilMinimo = 2032.88
    
    ultimaFila = Hoja27.Range("A" & Rows.Count).End(xlUp).Row
    For cont = 9 To ultimaFila
        kilos = Hoja27.Cells(cont, 8)
        provincia = Hoja27.Cells(cont, 7)
        UCase (provincia)
        localidad = Hoja27.Cells(cont, 6)
        UCase (localidad)
        Select Case provincia
         Case "BUENOS AIRES"
            Select Case localidad
             Case "HURLINGHAM" Or "LA PLATA"
                If kilos >= 4000 Then
                    tarifaInterior = kilos * valorGBA02
                End If
                If kilos >= 204 And kilos <= 4000 Then
                    tarifaInterior = kilos * valorGBA01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorGbaMinimo
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
             Case "9 DE JULIO" Or "CARMEN DE ARECO" Or "CHIVILICOY"
                    If kilos >= 14000 Then
                        tarifaInterior = kilos * valor9deJulio05
                    End If
                    If kilos < 14000 And kilos >= 8000 Then
                        tarifaInterior = kilos * valor9deJulio04
                    End If
                    If kilos < 8000 And kilos >= 4000 Then
                        tarifaInterior = kilos * valor9deJulio03
                    End If
                    If kilos < 4000 And kilos >= 1000 Then
                        tarifaInterior = kilos * valor9deJulio02
                    End If
                    If kilos < 1000 And kilos >= 214 Then
                        tarifaInterior = kilos * valor9deJulio01
                    End If
                    If kilos < 204 Then
                        tarifaInterior = valor9deJulioMinimo
                    End If
             Case "CACHARI" Or "ARRECIFES" Or "GENERAL VILLEGAS" Or "LA VIOLETA" Or "PERGAMINO" Or "TRENQUE LAUQUEN"
                    If kilos >= 14000 Then
                        tarifaInterior = kilos * valorCACHARI05
                    End If
                    If kilos < 14000 And kilos >= 8000 Then
                        tarifaInterior = kilos * valorCACHARI04
                    End If
                    If kilos < 8000 And kilos >= 4000 Then
                        tarifaInterior = kilos * valorCACHARI03
                    End If
                    If kilos < 4000 And kilos >= 1000 Then
                        tarifaInterior = kilos * valorCACHARI02
                    End If
                    If kilos < 1000 And kilos >= 214 Then
                        tarifaInterior = kilos * valorCACHARI01
                    End If
                    If kilos < 204 Then
                        tarifaInterior = valorCACHARIMinimo
                    End If
             Case "CARHUE" OR "CORONEL PRINGLES" OR "CORONEL SUAREZ" OR "PEDRO LURO" OR "TRES AROYOS"
                    If kilos >= 14000 Then
                       tarifaInterior = kilos * valorCarhue05
                    End If
                    If kilos < 14000 And kilos >= 8000 Then
                        tarifaInterior = kilos * valorCarhue04
                    End If
                    If kilos < 8000 And kilos >= 4000 Then
                        tarifaInterior = kilos * valorCarhue03
                    End If
                    If kilos < 4000 And kilos >= 1000 Then
                        tarifaInterior = kilos * valorCarhue02
                    End If
                    If kilos < 1000 And kilos >= 214 Then
                        tarifaInterior = kilos * valorCarhue01
                    End If
                    If kilos < 204 Then
                        tarifaInterior = valorCarhueMinimo
                    End If
             case "BRANDSEN" OR "GENERAL BELGRANO"
                    If kilos >= 14000 Then
                       tarifaInterior = kilos * valorBrandsen05
                    End If
                    If kilos < 14000 And kilos >= 8000 Then
                        tarifaInterior = kilos * valorBrandsen04
                    End If
                    If kilos < 8000 And kilos >= 4000 Then
                        tarifaInterior = kilos * valorBrandsen03
                    End If
                    If kilos < 4000 And kilos >= 1000 Then
                        tarifaInterior = kilos * valorBrandsen02
                    End If
                    If kilos < 1000 And kilos >= 214 Then
                        tarifaInterior = kilos * valorBrandsen01
                    End If
                    If kilos < 204 Then
                        tarifaInterior = valorBrandsenMinimo
                    End If
             case "JUNIN" OR "FERRE" OR "CHACHABUCO" OR "ROJAS"
                    If kilos >= 14000 Then
                       tarifaInterior = kilos * valorJunin05
                    End If
                    If kilos < 14000 And kilos >= 8000 Then
                        tarifaInterior = kilos * valorJunin04
                    End If
                    If kilos < 8000 And kilos >= 4000 Then
                        tarifaInterior = kilos * valorJunin03
                    End If
                    If kilos < 4000 And kilos >= 1000 Then
                        tarifaInterior = kilos * valorJunin02
                    End If
                    If kilos < 1000 And kilos >= 214 Then
                        tarifaInterior = kilos * valorJunin01
                    End If
                    If kilos < 204 Then
                        tarifaInterior = valorJuninMinimo
                    End If
             case "AZUL"
                    If kilos >= 14000 Then
                       tarifaInterior = kilos * valorAzul05
                    End If
                    If kilos < 14000 And kilos >= 8000 Then
                        tarifaInterior = kilos * valorAzul04
                    End If
                    If kilos < 8000 And kilos >= 4000 Then
                        tarifaInterior = kilos * valorAzul03
                    End If
                    If kilos < 4000 And kilos >= 1000 Then
                        tarifaInterior = kilos * valorAzul02
                    End If
                    If kilos < 1000 And kilos >= 214 Then
                        tarifaInterior = kilos * valorAzul01
                    End If
                    If kilos < 204 Then
                        tarifaInterior = valorAzul01
                    End If
             case "BAHIA BLANCA"
                    If kilos >= 14000 Then
                       tarifaInterior = kilos * valorBahia05
                    End If
                    If kilos < 14000 And kilos >= 8000 Then
                        tarifaInterior = kilos * valorBahia04
                    End If
                    If kilos < 8000 And kilos >= 4000 Then
                        tarifaInterior = kilos * valorBahia03
                    End If
                    If kilos < 4000 And kilos >= 1000 Then
                        tarifaInterior = kilos * valorBahia02
                    End If
                    If kilos < 1000 And kilos >= 214 Then
                        tarifaInterior = kilos * valorBahia01
                    End If
                    If kilos < 204 Then
                        tarifaInterior = valorBahiaMinimo
                    End If
             case "BALCARSE"
                    If kilos >= 14000 Then
                       tarifaInterior = kilos * valorBalcarse05
                    End If
                    If kilos < 14000 And kilos >= 8000 Then
                        tarifaInterior = kilos * valorBalcarse04
                    End If
                    If kilos < 8000 And kilos >= 4000 Then
                        tarifaInterior = kilos * valorBalcarse03
                    End If
                    If kilos < 4000 And kilos >= 1000 Then
                        tarifaInterior = kilos * valorBalcarse02
                    End If
                    If kilos < 1000 And kilos >= 214 Then
                        tarifaInterior = kilos * valorBalcarse01
                    End If
                    If kilos < 204 Then
                        tarifaInterior = valorBalcarseMinimo
                    End If
             case "MAR DEL PLATA"
                    If kilos >= 14000 Then
                       tarifaInterior = kilos * valorMardelPlata05
                    End If
                    If kilos < 14000 And kilos >= 8000 Then
                        tarifaInterior = kilos * valorMardelPlata04
                    End If
                    If kilos < 8000 And kilos >= 4000 Then
                        tarifaInterior = kilos * valorMardelPlata03
                    End If
                    If kilos < 4000 And kilos >= 1000 Then
                        tarifaInterior = kilos * valorMardelPlata02
                    End If
                    If kilos < 1000 And kilos >= 214 Then
                        tarifaInterior = kilos * valorMardelPlata01
                    End If
                    If kilos < 204 Then
                        tarifaInterior = valorMardelPlataMinimo
                    End If
             case "SAN NICOLAS"
                    If kilos >= 14000 Then
                       tarifaInterior = kilos * valorSanNicolas05
                    End If
                    If kilos < 14000 And kilos >= 8000 Then
                        tarifaInterior = kilos * valorSanNicolas04
                    End If
                    If kilos < 8000 And kilos >= 4000 Then
                        tarifaInterior = kilos * valorSanNicolas03
                    End If
                    If kilos < 4000 And kilos >= 1000 Then
                        tarifaInterior = kilos * valorSanNicolas02
                    End If
                    If kilos < 1000 And kilos >= 214 Then
                        tarifaInterior = kilos * valorSanNicolas01
                    End If
                    If kilos < 204 Then
                        tarifaInterior = valorSanNicolasMinimo
                    End If
             case "TANDIL"
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
                    End If
             Case else
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
                




                End Select
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
            Next cont


End Sub
