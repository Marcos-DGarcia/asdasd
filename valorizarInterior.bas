Sub valorizar_int()
    Dim localidad
    Dim provincia
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
    dim valorBalnearia01
    dim valorBalnearia02  
    dim valorBalnearia03 
    dim valorBalnearia04  
    dim valorBalnearia05 
    dim valorBalneariaMinimo  

    dim valorJesusMaria01
    dim valorJesusMaria02  
    dim valorJesusMaria03 
    dim valorJesusMaria04  
    dim valorJesusMaria05 
    dim valorJesusMariaMinimo

    dim valorCanals01
    dim valorCanals02  
    dim valorCanals03 
    dim valorCanals04  
    dim valorCanals05 
    dim valorCanalsMinimo  

    dim valorCatamarca01
    dim valorCatamarca02  
    dim valorCatamarca03 
    dim valorCatamarca04  
    dim valorCatamarca05 
    dim valorCatamarcaMinimo  

    dim valorCorrientes01
    dim valorCorrientes02  
    dim valorCorrientes03 
    dim valorCorrientes04  
    dim valorCorrientes05 
    dim valorCorrientesMinimo
    
    dim valorBovril01
    dim valorBovril02
    dim valorBovril03
    dim valorBovril04
    dim valorBovril05
    dim valorBovrilMinimo

    dim valorColon01
    dim valorColon02  
    dim valorColon03 
    dim valorColon04  
    dim valorColon05 
    dim valorColonMinimo

    dim valorGralRamirez01
    dim valorGralRamirez02  
    dim valorGralRamirez03 
    dim valorGralRamirez04  
    dim valorGralRamirez05 
    dim valorGralRamirezMinimo

    dim valorParana01
    dim valorParana02  
    dim valorParana03 
    dim valorParana04  
    dim valorParana05 
    dim valorParanaMinimo

    dim valorJujuy01
    dim valorJujuy02  
    dim valorJujuy03 
    dim valorJujuy04  
    dim valorJujuy05 
    dim valorJujuyMinimo

    dim valorMaipu01
    dim valorMaipu02  
    dim valorMaipu03 
    dim valorMaipu04  
    dim valorMaipu05 
    dim valorMaipuMinimo

    dim valorRawson01
    dim valorRawson02  
    dim valorRawson03 
    dim valorRawson04  
    dim valorRawson05 
    dim valorRawsonMinimo

    dim valorSanRafael01
    dim valorSanRafael02  
    dim valorSanRafael03 
    dim valorSanRafael04  
    dim valorSanRafael05 
    dim valorSanRafaelMinimo

    dim valorLaRioja01
    dim valorLaRioja02  
    dim valorLaRioja03 
    dim valorLaRioja04  
    dim valorLaRioja05 
    dim valorLaRiojaMinimo

    dim valorSalta01
    dim valorSalta02  
    dim valorSalta03 
    dim valorSalta04  
    dim valorSalta05 
    dim valorSaltaMinimo

    dim valorPocito01
    dim valorPocito02  
    dim valorPocito03 
    dim valorPocito04  
    dim valorPocito05 
    dim valorPocitoMinimo

    dim valorVillaMer01
    dim valorVillaMer02  
    dim valorVillaMer03 
    dim valorVillaMer04  
    dim valorVillaMer05 
    dim valorVillaMerMinimo

    dim valorArequito01
    dim valorArequito02  
    dim valorArequito03 
    dim valorArequito04  
    dim valorArequito05 
    dim valorArequitoMinimo

    Dim cont
    Dim ultimaFila
    Dim kilos
    Dim tarifaInterior
 '' BUENOS AIRES 
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
 '' CORDOBA
    ' BALNEARIA, GENERAL DEHEZA, LAS VARILLAS, MARCOS JUAREZ, VILLA MARIA, RIO TERCERO, SAN FRANCISCO
    valorBalnearia01 = 20.84
    valorBalnearia02 = 20.74
    valorBalnearia03 = 19.42
    valorBalnearia04 = 18.14
    valorBalnearia05 = 15.96
    valorBalneariaMinimo = 2217.43
    'CERRO DE LAS ROSAS, JESUS MARIA, ONCATIVO,CORDOBA, SAN MARTIN, 20 DE JUNIO, SINSACATE, VILLA CABRERA
    valorJesusMaria01 = 18.51
    valorJesusMaria02 = 18.42
    valorJesusMaria03 = 17.25
    valorJesusMaria04 = 16.11
    valorJesusMaria05 = 14.18
    valorJesusMariaMinimo = 1969.39
    'CANALS, LABOULAYE, MONTE MAIZ, HUINCA RENANCO, RIO CUARTO, VILLA HUIDOBRO
    valorCanals01 = 21.02
    valorCanals02 = 20.92
    valorCanals03 = 19.59
    valorCanals04 = 18.30
    valorCanals05 = 16.10
    valorCanalsMinimo = 2236.96
 '' CATAMARCA
    valorCatamarca01 = 26.57
    valorCatamarca02 = 26.45
    valorCatamarca03 = 24.76
    valorCatamarca04 = 23.13
    valorCatamarca05 = 24.03
    valorCatamarcaMinimo = 2827.26
 '' CORRIENTES
    valorCorrientes01 = 20.29
    valorCorrientes02 = 20.19
    valorCorrientes03 = 20.16
    valorCorrientes04 = 20.32
    valorCorrientes05 = 18.35
    valorCorrientesMinimo = 2158.90
 '' ENTRE RIOS 
    'COLON, CONCEPCION DEL URUGUAY, SAN SALVADOR, CHAJARI, COLONIA ADELA, CONCORDIA, GUALEGUAYCHU, VILLAGUAY
    valorColon01 = 19.49
    valorColon02 = 19.40
    valorColon03 = 18.17
    valorColon04 = 16.97
    valorColon05 = 11.41
    valorColonMinimo = 2074.36
    'BOVRIL
    valorBovril01 = 21.28
    valorBovril02 = 21.18
    valorBovril03 = 19.83
    valorBovril04 = 18.52
    valorBovril05 = 12.45
    valorBovrilMinimo = 2264.38
    'Gral Ramirez
    valorGralRamirez01 = 17.27
    valorGralRamirez02 = 17.19
    valorGralRamirez03 = 16.09
    valorGralRamirez04 = 15.03
    valorGralRamirez05 = 10.10
    valorGralRamirezMinimo = 1837.76
    'PARANA
    valorParana01 = 15.64
    valorParana02 = 15.57
    valorParana03 = 14.58
    valorParana04 = 13.62
    valorParana05 = 9.15
    valorParanaMinimo = 1664.54
 '' JUJUY
    valorJujuy01 = 22.44
    valorJujuy02 = 26.48
    valorJujuy03 = 24.79
    valorJujuy04 = 23.15
    valorJujuy05 = 24.05
    valorJujuyMinimo = 2830.38
 '' MENDOZA
    'LUJAN DE CUYO,MAIPU,MENDOZA,SAN MARTIN, GODOY CRUZ,GUAYMALLEN,CHACRAS DE CORIA,VILLA NUEVA
    valorMaipu01 = 20.43
    valorMaipu02 = 20.34
    valorMaipu03 = 19.04
    valorMaipu04 = 17.79
    valorMaipu05 = 15.65
    valorMaipuMinimo = 2174.61
    'RAWSON
    valorRawson01 = 20.66
    valorRawson02 = 20.57
    valorRawson03 = 19.26
    valorRawson04 = 17.99
    valorRawson05 = 15.83
    valorRawsonMinimo = 2198.76
    'SAN RAFAEL
    valorSanRafael01 = 21.02
    valorSanRafael02 = 20.92
    valorSanRafael03 = 19.59
    valorSanRafael04 = 18.30
    valorSanRafael05 = 16.10
    valorSanRafaelMinimo = 2236.43
 '' LA RIOJA
    valorLaRioja01 = 30.84
    valorLaRioja02 = 30.69
    valorLaRioja03 = 28.74
    valorLaRioja04 = 26.84
    valorLaRioja05 = 21.32
    valorLaRiojaMinimo = 3281.42
 '' SALTA
    valorSalta01 = 24.16 
    valorSalta02 = 24.05
    valorSalta03 = 22.52
    valorSalta04 = 21.03
    valorSalta05 = 21.85
    valorSaltaMinimo = 2572.16
 '' SAN JUAN 
    'POCITO,VILLA ABERASTAIN,RAWSON,SAN JUAN
    valorPocito01 = 20.66
    valorPocito02 = 20.57
    valorPocito03 = 19.26
    valorPocito04 = 17.99
    valorPocito05 = 15.83
    valorPocitoMinimo = 2198.76
    'SAN LUIS,VILLA MERCEDES
    valorVillaMer01 = 22.41
    valorVillaMer02 = 22.31
    valorVillaMer03 = 20.89
    valorVillaMer04 = 19.51
    valorVillaMer05 = 17.17
    valorVillaMerMinimo = 2385.11
 '' SANTA FE 
    'AREQUITO,BOMBAL,CAÑADA DE GOMEZ,RECREO,LA GUARDA, LAS PAREJAS,CASILDA,FIRMAT,RAFAELA,SAN LORENZO,SANTA FE, SANTA TERESA, SANTO TOME,TOTORAS,VENADO TUERTO,VILLA CAÑAS
    valorArequito01 = 15.64
    valorArequito02 = 15.57
    valorArequito03 = 14.58
    valorArequito04 = 13.62
    valorArequito05 = 9.15
    valorArequitoMinimo = 1664.54 
    'CARLOS PELLEGRINI,VILLA GOBERNADOR GALVEZ


    'RECONQUISTA

    'ROSARIO
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
                   tarifaInterior = kilos * valorTandil05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorTandil04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorTandil03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorTandil02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorTandil01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorTandilMinimo
                End If
                ' VALOR MAS CARO PARA LOCACALIDADES QUE NO ESTEN EN LISTA DE PRECIOS
             CASE ELSE 
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
            END SELECT
        Case "CORDOBA"
            Select case localidad
             case "BALNEARIA" or "GENERAL DEHEZA" or "LAS VARILLAS" or "MARCOS JUAREZ" or "VILLA MARIA" or "RIO TERCERO" or"SAN FRANCISCO"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorBalnearia05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorBalnearia04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorBalnearia03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorBalnearia02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorBalnearia01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorBalneariaMinimo
                End If
             case "CERRO DE LAS ROSAS" "JESUS MARIA" "ONCATIVO" "CORDOBA" "BARRIO SAN MARTIN" "BARRIO 20 DE JUNIO" "SINSACATE" "VILLA CABRERA"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorJesusMaria05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorJesusMaria04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorJesusMaria03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorJesusMaria02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorJesusMaria01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorJesusMariaMinimo
                End If
             case "CANALS" "LABOULAYE" "MONTE MAIZ" "HUINCA RENANCO" "RIO CUARTO" "VILLA HUIDOBRO"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorCanals05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorCanals04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorCanals03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorCanals02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorCanals01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorCanalsMinimo
                End If 
                ' OTRAS LOCALIDADES QUE NO ESTAN EN LA LISTA, VAN CON VALOR MAS CARO   
             case else 
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorCanals05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorCanals04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorCanals03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorCanals02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorCanals01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorCanalsMinimo
                End If 
            END SELECT
        Case "CATAMARCA"
            If kilos >= 14000 Then
               tarifaInterior = kilos * valorCatamarca05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorCatamarca04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorCatamarca03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorCatamarca02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorCatamarca01
            End If
            If kilos < 204 Then
                tarifaInterior = valorCatamarcaMinimo
            End If
        case "CORRIENTES"
            If kilos >= 14000 Then
               tarifaInterior = kilos * valorCorrientes05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorCorrientes04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorCorrientes03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorCorrientes02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorCorrientes01
            End If
            If kilos < 204 Then
                tarifaInterior = valorCorrientesMinimo
            End If
        case "ENTRE RIOS"
            select case localidad
             case "COLON" or "CONCEPCION DEL URUGUAY" or "SAN SALVADOR" or "CHAJARI" or "COLONIA ADELA" or "CONCORDIA" or "GUALEGUAYCHU" or "VILLAGUAY"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorColon05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorColon04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorColon03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorColon02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorColon01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorColonMinimo
                End If
             case "BOVRIL" 
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorBovril05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorBovril04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorBovril03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorBovril02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorBovril01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorBovrilMinimo
                End If 
             case "GRAL. RAMIREZ"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorGralRamirez05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorGralRamirez04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorGralRamirez03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorGralRamirez02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorGralRamirez01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorGralRamirezMinimo
                End If
             case "PARANA"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorParana05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorParana04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorParana03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorParana02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorParana01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorParanaMinimo
                End If           
                'VALORES QUE NO ESTAN EN LA LISTA DE PRECIOS 
             case else
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorBovril05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorBovril04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorBovril03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorBovril02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorBovril01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorBovrilMinimo
                End If 
            END SELECT
        case "JUJUY"
           If kilos >= 14000 Then
               tarifaInterior = kilos * valorJujuy05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorJujuy04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorJujuy03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorJujuy02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorJujuy01
            End If
            If kilos < 204 Then
                tarifaInterior = valorJujuyMinimo
            End If
        case "MENDOZA"
            select case localidad
             case "LUJAN DE CUYO" OR "MAIPU" OR "MENDOZA" OR "SAN MARTIN" OR "GODOY CRUZ" OR "GUAYMALLEN" OR "CHACRAS DE CORIA" OR "VILLA NUEVA"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorMaipu05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorMaipu04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorMaipu03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorMaipu02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorMaipu01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorMaipuMinimo
                End If   
             case "RAWSON"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorRawson05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorRawson04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorRawson03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorRawson02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorRawson01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorRawsonMinimo
                End If
             case "SAN RAFAEL"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorSanRafael05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorSanRafael04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorSanRafael03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorSanRafael02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorSanRafael01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorRawsonMinimo
                End If
                'valores para localidades que no estan en la lista de precios
             Case else
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorSanRafael05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorSanRafael04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorSanRafael03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorSanRafael02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorSanRafael01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorRawsonMinimo
                End If             
        case "LA RIOJA"
           If kilos >= 14000 Then
               tarifaInterior = kilos * valorJujuy05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorJujuy04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorJujuy03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorJujuy02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorJujuy01
            End If
            If kilos < 204 Then
                tarifaInterior = valorJujuyMinimo
            End If

        END SELECT



        
    Next cont



End Sub
