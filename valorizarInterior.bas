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
    
    dim valorCarlosPellegrini01
    dim valorCarlosPellegrini02  
    dim valorCarlosPellegrini03 
    dim valorCarlosPellegrini04  
    dim valorCarlosPellegrini05 
    dim valorCarlosPellegriniMinimo

    dim valorReconquista01
    dim valorReconquista02  
    dim valorReconquista03 
    dim valorReconquista04  
    dim valorReconquista05 
    dim valorReconquistaMinimo

    dim valorRosario01
    dim valorRosario02  
    dim valorRosario03 
    dim valorRosario04  
    dim valorRosario05 
    dim valorRosarioMinimo

    dim valorSantiago01
    dim valorSantiago02  
    dim valorSantiago03 
    dim valorSantiago04  
    dim valorSantiago05 
    dim valorSantiagoMinimo

    dim valorTucuman01
    dim valorTucuman02  
    dim valorTucuman03 
    dim valorTucuman04  
    dim valorTucuman05 
    dim valorTucumanMinimo

    dim valorCharata01
    dim valorCharata02  
    dim valorCharata03 
    dim valorCharata04  
    dim valorCharata05 
    dim valorCharataMinimo

    dim valorResistencia01
    dim valorResistencia02  
    dim valorResistencia03 
    dim valorResistencia04  
    dim valorResistencia05 
    dim valorResistenciaMinimo

    dim valorComodoro01
    dim valorComodoro02  
    dim valorComodoro03 
    dim valorComodoro04  
    dim valorComodoro05 
    dim valorComodoroMinimo

    dim valorTrelew01
    dim valorTrelew02  
    dim valorTrelew03 
    dim valorTrelew04  
    dim valorTrelew05 
    dim valorTrelewMinimo

    dim valorEsquel01
    dim valorEsquel02  
    dim valorEsquel03 
    dim valorEsquel04  
    dim valorEsquel05 
    dim valorEsquelMinimo

    dim valorFormoza01
    dim valorFormoza02  
    dim valorFormoza03 
    dim valorFormoza04  
    dim valorFormoza05 
    dim valorFormozaMinimo

    dim valorMisiones01
    dim valorMisiones02  
    dim valorMisiones03 
    dim valorMisiones04  
    dim valorMisiones05 
    dim valorMisionesMinimo

    dim valorCentenario01
    dim valorCentenario02  
    dim valorCentenario03 
    dim valorCentenario04  
    dim valorCentenario05 
    dim valorCentenarioMinimo

    dim valorZapala01
    dim valorZapala02  
    dim valorZapala03 
    dim valorZapala04  
    dim valorZapala05 
    dim valorZapalaMinimo

    dim valorLaPampa01
    dim valorLaPampa02  
    dim valorLaPampa03 
    dim valorLaPampa04  
    dim valorLaPampa05 
    dim valorLaPampaMinimo


    dim valorGralRoca01
    dim valorGralRoca02  
    dim valorGralRoca03 
    dim valorGralRoca04  
    dim valorGralRoca05 
    dim valorGralRocaMinimo

    dim valorBariloche01
    dim valorBariloche02  
    dim valorBariloche03 
    dim valorBariloche04  
    dim valorBariloche05 
    dim valorBarilocheMinimo

    dim valorCipolleti01
    dim valorCipolleti02  
    dim valorCipolleti03 
    dim valorCipolleti04  
    dim valorCipolleti05 
    dim valorCipolletiMinimo

    dim valorViedma01
    dim valorViedma02  
    dim valorViedma03 
    dim valorViedma04  
    dim valorViedma05 
    dim valorViedmaMinimo
    
    dim valorSantaCruz01
    dim valorSantaCruz02  
    dim valorSantaCruz03 
    dim valorSantaCruz04  
    dim valorSantaCruz05 
    dim valorSantaCruzMinimo

    dim valorTierraDelF01
    dim valorTierraDelF02  
    dim valorTierraDelF03 
    dim valorTierraDelF04  
    dim valorTierraDelF05 
    dim valorTierraDelFMinimo

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
    valorCarlosPellegrini01 = 17.27
    valorCarlosPellegrini02 = 17.19
    valorCarlosPellegrini03 = 16.09
    valorCarlosPellegrini04 = 15.03
    valorCarlosPellegrini05 = 10.10
    valorCarlosPellegriniMinimo = 1837.76
    'RECONQUISTA
    valorReconquista01 = 21.28
    valorReconquista02 = 21.18
    valorReconquista03 = 19.83
    valorReconquista04 = 18.52
    valorReconquista05 = 12.45
    valorReconquistaMinimo = 2264.38
    'ROSARIO
    valorRosario01 = 15.05
    valorRosario02 = 14.98
    valorRosario03 = 14.03
    valorRosario04 = 13.10
    valorRosario05 = 7.21
    valorRosarioMinimo = 1601.63
 '' SANTIAGO DEL ESTERO
    valorSantiago01 = 25.90
    valorSantiago02 = 25.78
    valorSantiago03 = 24.14
    valorSantiago04 = 22.55
    valorSantiago05 = 19.84
    valorSantiagoMinimo = 2755.96
 '' TUCUMAN
    valorTucuman01 = 22.02
    valorTucuman02 = 21.92
    valorTucuman03 = 20.52
    valorTucuman04 = 19.17
    valorTucuman05 = 19.91
    valorTucumanMinimo = 2343.20
 '' CHACO
    'CHARATA,VILLA ANGELA
    valorCharata01 = 23.90
    valorCharata02 = 23.79
    valorCharata03 = 23.75
    valorCharata04 = 23.94
    valorCharata05 = 21.61
    valorCharataMinimo = 2543.35

    'RESISTENCIA    
    valorResistencia01 = 20.29
    valorResistencia02 = 20.19
    valorResistencia03 = 20.16
    valorResistencia04 = 20.32
    valorResistencia05 = 18.35
    valorResistenciaMinimo = 2158.90
 '' CHUBUT
    'COMODORO RIVADAVIA
    valorComodoro01 = 34.96
    valorComodoro02 = 34.80
    valorComodoro03 = 34.74
    valorComodoro04 = 35.02
    valorComodoro05 = 33.47
    valorComodoroMinimo = 3720.25
    'TRELEW
    valorTrelew01 = 30.29
    valorTrelew02 = 30.15
    valorTrelew03 = 30.10
    valorTrelew04 = 30.35
    valorTrelew05 = 29.01
    valorTrelewMinimo = 3223.65
    'ESQUEL
    valorEsquel01 = 37.11
    valorEsquel02 = 36.94
    valorEsquel03 = 36.87
    valorEsquel04 = 37.18
    valorEsquel05 = 35.53
    valorEsquelMinimo = 3949.14
 '' FORMOSA
    valorFormoza01 = 23.90
    valorFormoza02 = 23.79
    valorFormoza03 = 23.75
    valorFormoza04 = 23.94
    valorFormoza05 = 21.61
    valorFormozaMinimo = 2543.35
 '' MISIONES
    valorMisiones01 = 21.50
    valorMisiones02 = 21.40
    valorMisiones03 = 21.36
    valorMisiones04 = 21.54
    valorMisiones05 = 19.44
    valorMisionesMinimo = 2287.54
 '' NEUQUEN
    '' CENTENARIO,NEUQUEN
    valorCentenario01 = 26.62
    valorCentenario02 = 26.53
    valorCentenario03 = 26.48
    valorCentenario04 = 26.70
    valorCentenario05 = 25.51
    valorCentenarioMinimo = 2835.69
    ''ZAPALA
    valorZapala01 = 31.06
    valorZapala02 = 30.92
    valorZapala03 = 30.87
    valorZapala04 = 31.12
    valorZapala05 = 29.74
    valorZapalaMinimo = 3305.72
 ''LA PAMPA
    valorLaPampa01 = 27.58
    valorLaPampa02 = 27.45
    valorLaPampa03 = 27.40
    valorLaPampa04 = 27.63
    valorLaPampa05 = 16.14
    valorLaPampaMinimo = 2934.63
 '' RIO NEGRO
    'GRAL ROCA
    valorGralRoca01 = 28.48
    valorGralRoca02 = 28.34
    valorGralRoca03 = 28.29
    valorGralRoca04 = 28.52
    valorGralRoca05 = 27.26
    valorGralRocaMinimo = 3029.83
    'BARILOCHE
    valorBariloche01 = 35.51
    valorBariloche02 = 35.35
    valorBariloche03 = 35.29
    valorBariloche04 = 35.58
    valorBariloche05 = 34.00
    valorBarilocheMinimo = 3779.08
    'CIPOLLETTI
    valorCipolleti01 = 28.09
    valorCipolleti02 = 27.96
    valorCipolleti03 = 27.91
    valorCipolleti04 = 28.14
    valorCipolleti05 = 26.89
    valorCipolletiMinimo = 2988.96
    ''VIEDMA
    valorViedma01 = 26.89
    valorViedma02 = 26.77
    valorViedma03 = 26.72
    valorViedma04 = 26.94
    valorViedma05 = 25.75
    valorViedmaMinimo = 2861.93
 '' SANTA CRUZ
    valorSantaCruz01 = 44.18
    valorSantaCruz02 = 43.98
    valorSantaCruz03 = 43.90
    valorSantaCruz04 = 44.26
    valorSantaCruz05 = 42.30
    valorSantaCruzMinimo = 7051.79
 '' TIERRA DEL FUEGO
    valorTierraDelF01 = 48.42
    valorTierraDelF02 = 48.20
    valorTierraDelF03 = 48.11
    valorTierraDelF04 = 45.51
    valorTierraDelF05 = 44.51
    valorTierraDelFMinimo = 7729.24
 ''
    ultimaFila = Hoja27.Range("A" & Rows.Count).End(xlUp).Row
    For cont = 9 To ultimaFila
        kilos = Hoja27.Cells(cont, 8)
        provincia = Hoja27.Cells(cont, 7)
        UCase (provincia)
        localidad = Hoja27.Cells(cont, 6)
        UCase (localidad)
        Select Case provincia
        case "BUENOS AIRES"
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
        case "CORDOBA"
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
        case "CATAMARCA"
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
            END SELECT         
        case "LA RIOJA"
           If kilos >= 14000 Then
               tarifaInterior = kilos * valorLaRioja05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorLaRioja04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorLaRioja03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorLaRioja02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorLaRioja01
            End If
            If kilos < 204 Then
                tarifaInterior = valorLaRiojaMinimo
            End If
        case "SALTA"
           If kilos >= 14000 Then
               tarifaInterior = kilos * valorSalta05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorSalta04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorSalta03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorSalta02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorSalta01
            End If
            If kilos < 204 Then
                tarifaInterior = valorSaltaMinimo
            End If
        case "SAN JUAN"
            select case  localidad
            case "POCITO" or "VILLA ABERASTAIN" or "RAWSON" or "SAN JUAN"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorPocito05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorPocito04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorPocito03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorPocito02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorPocito01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorPocitoMinimo
                End If                   
            case "SAN LUIS" or "VILLA MERCEDES"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorVillaMer05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorVillaMer04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorVillaMer03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorVillaMer02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorVillaMer01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorVillaMerMinimo
                End If        
                'localidades que no estan en la lista                       
            case else 
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorVillaMer05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorVillaMer04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorVillaMer03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorVillaMer02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorVillaMer01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorVillaMerMinimo
                End If 
            END SELECT
        case "SANTA FE"
            select case localidad
             case "AREQUITO" OR "BOMBAL" OR "CAÑADA DE GOMEZ" OR "RECREO" OR "LA GUARDA" OR "LAS PAREJAS" OR "CASILDA"OR "FIRMAT" OR "RAFAELA" OR "SAN LORENZO" OR "SANTA FE" OR "SANTA TERESA" OR  "SANTO TOME" OR "TOTORAS" OR "VENADO TUERTO" OR "VILLA CAÑAS"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorArequito05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorArequito04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorArequito03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorArequito02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorArequito01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorArequitoMinimo
                End If
             case "CARLOS PELLEGRINI" or "VILLA GOBERNADOR GALVEZ"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorCarlosPellegrini05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorCarlosPellegrini04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorCarlosPellegrini03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorCarlosPellegrini02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorCarlosPellegrini01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorCarlosPellegriniMinimo
                End If
             case "RECONQUISTA"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorReconquista05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorReconquista04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorReconquista03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorReconquista02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorReconquista01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorReconquistaMinimo
                End If
             case "ROSARIO"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorRosario05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorRosario04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorRosario03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorRosario02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorRosario01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorRosarioMinimo
                End If
             case else
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorReconquista05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorReconquista04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorReconquista03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorReconquista02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorReconquista01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorReconquistaMinimo
                End If             
            END SELECT
        case "SANTIAGO DEL ESTERO"
           If kilos >= 14000 Then
               tarifaInterior = kilos * valorSantiago05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorSantiago04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorSantiago03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorSantiago02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorSantiago01
            End If
            If kilos < 204 Then
                tarifaInterior = valorSantiagoMinimo
            End If
        case "TUCUMAN"
           If kilos >= 14000 Then
               tarifaInterior = kilos * valorTucuman05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorTucuman04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorTucuman03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorTucuman02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorTucuman01
            End If
            If kilos < 204 Then
                tarifaInterior = valorTucumanMinimo
            End If   
        case "CHACO"
            select case localidad
             case "CHARATA" OR "VILLA ANGELA"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorCharata05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorCharata04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorCharata03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorCharata02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorCharata01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorCharataMinimo
                End If
             case "RESISTENCIA"    
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorResistencia05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorResistencia04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorResistencia03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorResistencia02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorResistencia01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorResistenciaMinimo
                End If
             case else
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorCharata05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorCharata04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorCharata03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorCharata02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorCharata01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorCharataMinimo
                End If
             END SELECT
        case "CHUBUT"
            select case localidad
             case "COMODORO RIVADAVIA"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorComodoro05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorComodoro04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorComodoro03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorComodoro02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorComodoro01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorComodoroMinimo
                End If
             case "TRELEW"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorTrelew05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorTrelew04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorTrelew03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorTrelew02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorTrelew01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorTrelewMinimo
                End If                             
             case "ESQUEL"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorEsquel05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorEsquel04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorEsquel03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorEsquel02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorEsquel01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorEsquelMinimo
                End If
             case else 
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorEsquel05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorEsquel04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorEsquel03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorEsquel02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorEsquel01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorEsquelMinimo
                End If
            END SELECT
        case "FORMOSA"
           If kilos >= 14000 Then
               tarifaInterior = kilos * valorFormoza05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorFormoza04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorFormoza03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorFormoza02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorFormoza01
            End If
            If kilos < 204 Then
                tarifaInterior = valorFormozaMinimo
            End If 
        case "MISIONES"       
           If kilos >= 14000 Then
               tarifaInterior = kilos * valorMisiones05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorMisiones04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorMisiones03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorMisiones02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorMisiones01
            End If
            If kilos < 204 Then
                tarifaInterior = valorMisionesMinimo
            End If
        case "NEUQUEN"
            select case localidad
             case " CENTENARIO" or "NEUQUEN"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorCentenario05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorCentenario04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorCentenario03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorCentenario02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorCentenario01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorCentenarioMinimo
                End If
             case "ZAPALA"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorZapala05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorZapala04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorZapala03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorZapala02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorZapala01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorZapalaMinimo
                End If
             case else 
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorZapala05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorZapala04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorZapala03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorZapala02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorZapala01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorZapalaMinimo
                End If
            END SELECT
        case "LA PAMPA"
           If kilos >= 14000 Then
               tarifaInterior = kilos * valorLaPampa05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorLaPampa04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorLaPampa03
            End If
            If kilos < 4000 And kilos >= 1000 Then  
                tarifaInterior = kilos * valorLaPampa02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorLaPampa01
            End If
            If kilos < 204 Then
                tarifaInterior = valorLaPampaMinimo
            End If 
        case "RIO NEGRO"
            select case localidad
             case "GRAL ROCA"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorGralRoca05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorGralRoca04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorGralRoca03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorGralRoca02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorGralRoca01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorGralRocaMinimo
                End If
             case "BARILOCHE"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorBariloche05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorBariloche04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorBariloche03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorBariloche02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorBariloche01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorBarilocheMinimo
                End If
             case "CIPOLLETI"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorCipolleti05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorCipolleti04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorCipolleti03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorCipolleti04
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorCipolleti05
                End If
                If kilos < 204 Then
                    tarifaInterior = valorCipolletiMinimo
                End If
             case "VIEDMA"
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorViedma05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorViedma04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorViedma03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorViedma02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorViedma01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorViedmaMinimo
                End If
             case else
                If kilos >= 14000 Then
                   tarifaInterior = kilos * valorBariloche05
                End If
                If kilos < 14000 And kilos >= 8000 Then
                    tarifaInterior = kilos * valorBariloche04
                End If
                If kilos < 8000 And kilos >= 4000 Then
                    tarifaInterior = kilos * valorBariloche03
                End If
                If kilos < 4000 And kilos >= 1000 Then
                    tarifaInterior = kilos * valorBariloche02
                End If
                If kilos < 1000 And kilos >= 214 Then
                    tarifaInterior = kilos * valorBariloche01
                End If
                If kilos < 204 Then
                    tarifaInterior = valorBarilocheMinimo
                End If
            END SELECT
        case "SANTA CRUZ"
           If kilos >= 14000 Then
               tarifaInterior = kilos * valorSantaCruz05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorSantaCruz04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorSantaCruz03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorSantaCruz02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorSantaCruz01
            End If
            If kilos < 204 Then
                tarifaInterior = valorSantaCruzMinimo
            End If 
        case "TIERRA DEL FUEGO"
           If kilos >= 14000 Then
               tarifaInterior = kilos * valorTierraDelF05
            End If
            If kilos < 14000 And kilos >= 8000 Then
                tarifaInterior = kilos * valorTierraDelF04
            End If
            If kilos < 8000 And kilos >= 4000 Then
                tarifaInterior = kilos * valorTierraDelF03
            End If
            If kilos < 4000 And kilos >= 1000 Then
                tarifaInterior = kilos * valorTierraDelF02
            End If
            If kilos < 1000 And kilos >= 214 Then
                tarifaInterior = kilos * valorTierraDelF01
            End If
            If kilos < 204 Then
                tarifaInterior = valorTierraDelFMinimo
            End If                     
        END SELECT



        
    Next cont



End Sub
