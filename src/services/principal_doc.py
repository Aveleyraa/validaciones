# -*- coding: utf-8 -*-
"""
Created on Wed Nov  3 11:07:44 2021

@author: AARON.RAMIREZ
"""

from services.catalogos import validar_catalogo
from utility.common_utils import CommonUtils, x_valida,final_tabla,frame


def procesarcoor(hopan, hoja, seccion):  # funcion principal!!
    """
    hopan: dataframe de pandas
    hoja: hoja de excel de openpyxl
    
    
    """
    

    pro = CommonUtils.preguntas(hopan)
    t = CommonUtils.espacio(hopan, pro)
    con = 0
    for i in t:

        pregunta = pro[con]  # aqui va fila inicio de la pregunta para iterar
        nter = hopan.iat[pregunta, 0]
        frame.agregar_valor_acolumna(nter,'pregunta')
        frame.agregar_valor_acolumna(seccion,'seccion')
        r = CommonUtils.analizarcor(con, hopan, pro)
        totales = CommonUtils.paratotales(con, hopan, pro)
        part_tab = r[7]
        part = r[
            7
        ]  # se hace copia porque la variable se modifica y a las condicionales de abajo ya no llega bien
        print("este es r de pregunta ", nter, ":", r)
        letras_validadas = []
        # alfa = clasificador(r,totales)
        # print('el alfa: ', alfa)
        if "sabe" in r[8] or "sabe1" in r[8] or "noaplica" in r[8] or "NA" in r[8]:
            diccio = r[8]
            tuplas = diccio["tuplas"]
            ad = CommonUtils.masdeunatablauni(tuplas)
            for i in ad:
                if nter != '8.1.-':
                    freal = pregunta + i + 1
                    CommonUtils.validarcondicional(diccio, part, freal, i, hoja)
                else:  # porque pusieron un encabezado todo combinado y había que sumar un uno para ajustar, ya que prolema de merge cell
                    freal = pregunta + i + 1 + 1
                    for tuplas in part:
                        npart = [[(tupla[0]-1,tupla[1]) for tupla in tuplas]]
                    #este ajuste es exclusivo para pregunta 8.1 de modulo 1
                    CommonUtils.validarcondicional(diccio, npart, freal, i, hoja)
                # funcion para validar
        if 'temporales' in r[8]:
            # diccionario = r[8]['temporales']
            CommonUtils.validar_temporal(pregunta, hoja, r[8])
        

        if r[0] == 1 and r[1] == 0:  # si es tabla general
            a = CommonUtils.getcor(con, hopan, pro)
            inicios = a["inicio"]
            tupla = a["coord"]
            tupla.sort()
            print("inicios para tabla normal: ", a)
            c = 0
            for i in inicios[:1]:
                freal = pregunta + i  # fila real
                fin = r[6]
                ntuplas = []
                for tupl in tupla:
                    tfil = tupl[0]
                    if tfil == i:
                        ntuplas.append(tupl)
                tuplas = CommonUtils.subtuplas(totales, ntuplas, part_tab,r[8])
                # print("salida subtuplas: ", tuplas, len(tuplas), len(ntuplas))
                if fin == []:  # para ajustar perguntas de tabla que no tiene autosuma
                    fin = [0 - pregunta for i in range(0, len(inicios))]
                if len(tuplas) == len(ntuplas):
                    ldeletras = CommonUtils.letraslide(c)
                    if fin == []:
                        hay = CommonUtils.variables(
                            i, tuplas, freal, 0, hoja, part_tab, ldeletras, "si"
                        )
                    else:
                        hay = CommonUtils.variables(
                            i,
                            tuplas,
                            freal,
                            fin[c] + pregunta,
                            hoja,
                            part_tab,
                            ldeletras,
                            "si",
                        )
                    hay = [hay[0]]
                if c == len(fin):
                    break
                if len(tuplas) != len(ntuplas):
                    part_tab = 0
                    c1 = 0
                    extraer_letras = []
                    for lista in tuplas:
                        print("entro aqui 1")
                        ldeletras = CommonUtils.letraslide(c1)

                        # if lista ==tuplas[-1] and len(totales['subtotal']) > 0 and lista[-1] != 'a':
                        #     print('entro aqui 2')
                        #     hay = validarTS(i,lista,freal,fin[c]+pregunta,hoja,ldeletras)

                        if lista[-1] == "a":
                            lista.remove("a")
                            print("entro aqui 3")
                            hay = CommonUtils.validarTS(
                                i, lista, freal, fin[c] + pregunta, hoja, ldeletras
                            )
                        else:
                            print("entro aqui 4")
                            hay = CommonUtils.variables(
                                i,
                                lista,
                                freal,
                                fin[c] + pregunta,
                                hoja,
                                part_tab,
                                ldeletras,
                                "no",
                            )
                            hay = hay[0]
                        letras_validadas.append(hay)
                        extraer_letras.append(ldeletras[:6])
                        c1 += 1

                if fin[c] + pregunta == 0:
                    try:
                        CommonUtils.menerrorsub(
                            freal + 2,
                            a["final"][0] + 2 + pregunta,
                            hoja,
                            extraer_letras,
                            r[9] + pregunta,
                        )
                    except:
                        CommonUtils.menerrorsub(
                            freal + 2,
                            4 + 2 + pregunta,
                            hoja,
                            extraer_letras,
                            r[9] + pregunta,
                        )#correccion especial para pregunta 1.54 de CNIJF 2022
                if fin[c] + pregunta != 0:
                    CommonUtils.menerrorsub(
                        freal + 2,
                        fin[c] + pregunta + 2,
                        hoja,
                        extraer_letras,
                        r[9] + pregunta,
                    )
                c += 1
        if r[1] != 0 and r[0] == 1:  # tablas de filas unicas
            ad = CommonUtils.masdeunatablauni(r[1])
            for i in ad:

                freal = (
                    pregunta + i + 1
                )  # mas uno porque es la fila del titulo de columnas
                r[1].sort()
                tuplas = CommonUtils.subtuplas(totales, r[1], part_tab,r[8])
                ntuplas = []
                for tupl in r[1]:
                    tfil = tupl[0]
                    if tfil == i:
                        ntuplas.append(tupl)
                if len(tuplas) == len(r[1]):
                    ldeletras = CommonUtils.letraslide(0)
                    hay = CommonUtils.variables(
                        i, ntuplas, freal, 0, hoja, part_tab, ldeletras, "si"
                    )
                    hay = hay[0]
                else:
                    extraer_letras = []
                    part_tab = 0
                    c = 0
                    for lista in tuplas:
                        ldeletras = CommonUtils.letraslide(c)

                        # if lista ==tuplas[-1] and len(totales['subtotal']) > 0:
                        #     hay = validarTS(i,lista,freal,0,hoja,ldeletras)

                        if lista[-1] == "a":
                            lista.remove("a")
                            hay = CommonUtils.validarTS(
                                i, lista, freal, 0, hoja, ldeletras
                            )

                        else:
                            hay = CommonUtils.variables(
                                i, lista, freal, 0, hoja, part_tab, ldeletras, "no"
                            )
                            hay = hay[0]
                        letras_validadas.append(hay)
                        extraer_letras.append(ldeletras[:6])
                        c += 1
                CommonUtils.menerrorsub(
                    freal + 2, freal + 2, hoja, extraer_letras, r[9] + pregunta
                )

        if r[0] == 2:
            p = r[4]
            u = p["desagregados"]
            ad = CommonUtils.masdeunatablauni1(p["tuplas"])
            delito = CommonUtils.buscarpalabra(con, "Tipo de delito", hopan, pro)
            k = 0

            for i in ad:
                tp = p["tuplas"]
                tp.sort()
                tuplas = CommonUtils.subtuplas(totales, tp, part_tab,r[8])
                codelitos = u[k]
                de = delito[k]
                freal = i + 1 + pregunta
                fin = r[6]
                if len(tuplas) == len(tp):
                    ldeletras = CommonUtils.letraslide(0)
                    hay = CommonUtils.vardelitos(
                        i + 1,
                        p["tuplas"],
                        freal,
                        de,
                        codelitos,
                        pregunta + 2,
                        fin[k] + pregunta,
                        hoja,
                        part_tab,
                        ldeletras,
                        "si",
                    )
                    letras_validadas.append(hay)
                else:
                    extraer_letras = []
                    part_tab = 0
                    c = 0
                    for lista in tuplas:
                        ldeletras = CommonUtils.letraslide(c)

                        # if lista ==tuplas[-1] and len(totales['subtotal']) > 0:
                        #     hay = validarTS(i,lista,freal,0,hoja,ldeletras)

                        if k == len(fin):
                            break

                        if lista[-1] == "a":
                            lista.remove("a")
                            hay = CommonUtils.validarTSD(
                                i,
                                lista,
                                freal,
                                fin[k] + pregunta,
                                hoja,
                                ldeletras,
                                pregunta + 2,
                                codelitos,
                            )
                            letras_validadas.append(hay)
                        else:
                            hay = CommonUtils.vardelitos(
                                i + 1,
                                lista,
                                freal,
                                de,
                                codelitos,
                                pregunta + 2,
                                fin[k] + pregunta + 1,
                                hoja,
                                part_tab,
                                ldeletras,
                                "no",
                            )
                            letras_validadas.append(hay)
                        extraer_letras.append(ldeletras[:6])
                        c += 1

                CommonUtils.menerrorsub(
                    freal + 2,
                    fin[k] + pregunta + 2,
                    hoja,
                    extraer_letras,
                    r[9] + pregunta,
                )
                k += 1
            na_desagregados = x_valida.juntar()
            CommonUtils.gris_desagregados(
                na_desagregados, letras_validadas, hoja
                )
        if r[2] == 1 and r[0] == 1 and r[1]==0:  # tablas de unica columna de respuesta
            ad = CommonUtils.masdeunatablauni(r[3])
            c = 0
            ad.sort()
            ldeletras = CommonUtils.letraslide(0)
            for i in ad:
                freal = pregunta + i + 1
                fin = r[6]
                CommonUtils.variablesl1(
                    i, r[3], freal, fin[c] + pregunta, hoja, ldeletras
                )
                c += 1

            if len(r[6]) > len(ad):
                print("Tabla con columnas de respuesta unica no detectada")

        if r[0] == 0:  # Las que no son tablas
            tuplas = r[5]  # esto debe ser lista de tuplas
            corde = CommonUtils.clasi(tuplas)
            freal = pregunta + 2
            CommonUtils.variablesnotab(corde, freal, hoja, r[9])


        if r[0] == 3:
            diccio = r[8]
            tuplas = diccio["tuplas"]
            if r[3]:
                diccio['tuplas'] = r[3]
            ad = CommonUtils.masdeunatablauni(tuplas)
            fin = r[6]

            if fin == []:
                fin = [0 for i in ad]

            c = 0
            for i in ad:
                freal = pregunta + i + 1
                if c == len(fin):
                    break
                CommonUtils.soloautosuma(diccio, part_tab, freal, i, hoja, fin[c])
                c += 1
        if "catalogos" in r[8]:
            di = r[8]
            validar_catalogo(pregunta, di, hoja)

        if totales != "No":
            fn = final_tabla.juntar()
            if not fn:
                fn =[0]
            CommonUtils.poner_gris(letras_validadas,hoja,fn[0])

        print("Todo bien con pregunta ", nter)
        con += 1
        frame.conjuntar_db()
        frame.ajustar_db()
    

    return


def p_especificas(hopan,hoja,preguntas_validar):         
    """
    hopan: dataframe de pandas
    hoja: hoja de excel de openpyxl
    preguntas: list. Lista con preguntas exactamente como están en la 
    columna A del excel
    
    
    
    """

    pro = CommonUtils.preguntas(hopan)
    t = CommonUtils.espacio(hopan, pro)
    con = 0
    for i in t:

        pregunta = pro[con]  # aqui va fila inicio de la pregunta para iterar
        nter = hopan.iat[pregunta, 0]
        if nter in preguntas_validar:
            r = CommonUtils.analizarcor(con, hopan, pro)
            totales = CommonUtils.paratotales(con, hopan, pro)
            part_tab = r[7]
            part = r[
                7
            ]  # se hace copia porque la variable se modifica y a las condicionales de abajo ya no llega bien
            print("este es r de pregunta ", nter, ":", r)
            letras_validadas = []
            # alfa = clasificador(r,totales)
            # print('el alfa: ', alfa)
            if "sabe" in r[8] or "sabe1" in r[8] or "noaplica" in r[8] or "NA" in r[8]:
                diccio = r[8]
                tuplas = diccio["tuplas"]
                ad = CommonUtils.masdeunatablauni(tuplas)
                for i in ad:
                    if nter != '8.1.-':
                        freal = pregunta + i + 1
                        CommonUtils.validarcondicional(diccio, part, freal, i, hoja)
                    else:  # porque pusieron un encabezado todo combinado y había que sumar un uno para ajustar, ya que prolema de merge cell
                        freal = pregunta + i + 1 + 1
                        for tuplas in part:
                            npart = [[(tupla[0]-1,tupla[1]) for tupla in tuplas]]
                        #este ajuste es exclusivo para pregunta 8.1 de modulo 1
                        CommonUtils.validarcondicional(diccio, npart, freal, i, hoja)
                    # funcion para validar
            if 'temporales' in r[8]:
                # diccionario = r[8]['temporales']
                CommonUtils.validar_temporal(pregunta, hoja, r[8])
            
    
            if r[0] == 1 and r[1] == 0:  # si es tabla general
                a = CommonUtils.getcor(con, hopan, pro)
                inicios = a["inicio"]
                tupla = a["coord"]
                tupla.sort()
                print("inicios para tabla normal: ", a)
                c = 0
                for i in inicios[:1]:
                    freal = pregunta + i  # fila real
                    fin = r[6]
                    ntuplas = []
                    for tupl in tupla:
                        tfil = tupl[0]
                        if tfil == i:
                            ntuplas.append(tupl)
                    tuplas = CommonUtils.subtuplas(totales, ntuplas, part_tab,r[8])
                    # print("salida subtuplas: ", tuplas, len(tuplas), len(ntuplas))
                    if fin == []:  # para ajustar perguntas de tabla que no tiene autosuma
                        fin = [0 - pregunta for i in range(0, len(inicios))]
                    if len(tuplas) == len(ntuplas):
                        ldeletras = CommonUtils.letraslide(c)
                        if fin == []:
                            hay = CommonUtils.variables(
                                i, tuplas, freal, 0, hoja, part_tab, ldeletras, "si"
                            )
                        else:
                            hay = CommonUtils.variables(
                                i,
                                tuplas,
                                freal,
                                fin[c] + pregunta,
                                hoja,
                                part_tab,
                                ldeletras,
                                "si",
                            )
                        hay = [hay[0]]
                    if c == len(fin):
                        break
                    if len(tuplas) != len(ntuplas):
                        part_tab = 0
                        c1 = 0
                        extraer_letras = []
                        for lista in tuplas:
                            print("entro aqui 1")
                            ldeletras = CommonUtils.letraslide(c1)
    
                            # if lista ==tuplas[-1] and len(totales['subtotal']) > 0 and lista[-1] != 'a':
                            #     print('entro aqui 2')
                            #     hay = validarTS(i,lista,freal,fin[c]+pregunta,hoja,ldeletras)
    
                            if lista[-1] == "a":
                                lista.remove("a")
                                print("entro aqui 3")
                                hay = CommonUtils.validarTS(
                                    i, lista, freal, fin[c] + pregunta, hoja, ldeletras
                                )
                            else:
                                print("entro aqui 4")
                                hay = CommonUtils.variables(
                                    i,
                                    lista,
                                    freal,
                                    fin[c] + pregunta,
                                    hoja,
                                    part_tab,
                                    ldeletras,
                                    "no",
                                )
                                hay = hay[0]
                            letras_validadas.append(hay)
                            extraer_letras.append(ldeletras[:6])
                            c1 += 1
    
                    if fin[c] + pregunta == 0:
                        CommonUtils.menerrorsub(
                            freal + 2,
                            a["final"][0] + 2 + pregunta,
                            hoja,
                            extraer_letras,
                            r[9] + pregunta,
                        )
                    if fin[c] + pregunta != 0:
                        CommonUtils.menerrorsub(
                            freal + 2,
                            fin[c] + pregunta + 2,
                            hoja,
                            extraer_letras,
                            r[9] + pregunta,
                        )
                    c += 1
            if r[1] != 0 and r[0] == 1:  # tablas de filas unicas
                ad = CommonUtils.masdeunatablauni(r[1])
                for i in ad:
    
                    freal = (
                        pregunta + i + 1
                    )  # mas uno porque es la fila del titulo de columnas
                    r[1].sort()
                    tuplas = CommonUtils.subtuplas(totales, r[1], part_tab,r[8])
                    ntuplas = []
                    for tupl in r[1]:
                        tfil = tupl[0]
                        if tfil == i:
                            ntuplas.append(tupl)
                    if len(tuplas) == len(r[1]):
                        ldeletras = CommonUtils.letraslide(0)
                        hay = CommonUtils.variables(
                            i, ntuplas, freal, 0, hoja, part_tab, ldeletras, "si"
                        )
                        hay = hay[0]
                    else:
                        extraer_letras = []
                        part_tab = 0
                        c = 0
                        for lista in tuplas:
                            ldeletras = CommonUtils.letraslide(c)
    
                            # if lista ==tuplas[-1] and len(totales['subtotal']) > 0:
                            #     hay = validarTS(i,lista,freal,0,hoja,ldeletras)
    
                            if lista[-1] == "a":
                                lista.remove("a")
                                hay = CommonUtils.validarTS(
                                    i, lista, freal, 0, hoja, ldeletras
                                )
    
                            else:
                                hay = CommonUtils.variables(
                                    i, lista, freal, 0, hoja, part_tab, ldeletras, "no"
                                )
                                hay = hay[0]
                            letras_validadas.append(hay)
                            extraer_letras.append(ldeletras[:6])
                            c += 1
                    CommonUtils.menerrorsub(
                        freal + 2, freal + 2, hoja, extraer_letras, r[9] + pregunta
                    )
    
            if r[0] == 2:
                p = r[4]
                u = p["desagregados"]
                ad = CommonUtils.masdeunatablauni1(p["tuplas"])
                delito = CommonUtils.buscarpalabra(con, "Tipo de delito", hopan, pro)
                k = 0
    
                for i in ad:
                    tp = p["tuplas"]
                    tp.sort()
                    tuplas = CommonUtils.subtuplas(totales, tp, part_tab,r[8])
                    codelitos = u[k]
                    de = delito[k]
                    freal = i + 1 + pregunta
                    fin = r[6]
                    if len(tuplas) == len(tp):
                        ldeletras = CommonUtils.letraslide(0)
                        hay = CommonUtils.vardelitos(
                            i + 1,
                            p["tuplas"],
                            freal,
                            de,
                            codelitos,
                            pregunta + 2,
                            fin[k] + pregunta,
                            hoja,
                            part_tab,
                            ldeletras,
                            "si",
                        )
                        letras_validadas.append(hay)
                    else:
                        extraer_letras = []
                        part_tab = 0
                        c = 0
                        for lista in tuplas:
                            ldeletras = CommonUtils.letraslide(c)
    
                            # if lista ==tuplas[-1] and len(totales['subtotal']) > 0:
                            #     hay = validarTS(i,lista,freal,0,hoja,ldeletras)
    
                            if k == len(fin):
                                break
    
                            if lista[-1] == "a":
                                lista.remove("a")
                                hay = CommonUtils.validarTSD(
                                    i,
                                    lista,
                                    freal,
                                    fin[k] + pregunta,
                                    hoja,
                                    ldeletras,
                                    pregunta + 2,
                                    codelitos,
                                )
                                letras_validadas.append(hay)
                            else:
                                hay = CommonUtils.vardelitos(
                                    i + 1,
                                    lista,
                                    freal,
                                    de,
                                    codelitos,
                                    pregunta + 2,
                                    fin[k] + pregunta + 1,
                                    hoja,
                                    part_tab,
                                    ldeletras,
                                    "no",
                                )
                                letras_validadas.append(hay)
                            extraer_letras.append(ldeletras[:6])
                            c += 1
    
                    CommonUtils.menerrorsub(
                        freal + 2,
                        fin[k] + pregunta + 2,
                        hoja,
                        extraer_letras,
                        r[9] + pregunta,
                    )
                    k += 1
                na_desagregados = x_valida.juntar()
                CommonUtils.gris_desagregados(
                    na_desagregados, letras_validadas, hoja
                    )
            if r[2] == 1 and r[0] == 1 and r[1] ==0:  # tablas de unica columna de respuesta
                ad = CommonUtils.masdeunatablauni(r[3])
                c = 0
                ad.sort()
                ldeletras = CommonUtils.letraslide(0)
                for i in ad:
                    freal = pregunta + i + 1
                    fin = r[6]
                    CommonUtils.variablesl1(
                        i, r[3], freal, fin[c] + pregunta, hoja, ldeletras
                    )
                    c += 1
    
                if len(r[6]) > len(ad):
                    print("Tabla con columnas de respuesta unica no detectada")
    
            if r[0] == 0:  # Las que no son tablas
                tuplas = r[5]  # esto debe ser lista de tuplas
                corde = CommonUtils.clasi(tuplas)
                freal = pregunta + 2
                CommonUtils.variablesnotab(corde, freal, hoja, r[9])
    
    
            if r[0] == 3:
                diccio = r[8]
                tuplas = diccio["tuplas"]
                ad = CommonUtils.masdeunatablauni(tuplas)
                fin = r[6]
    
                if fin == []:
                    fin = [0 for i in ad]
    
                c = 0
                for i in ad:
                    freal = pregunta + i + 1
                    if c == len(fin):
                        break
                    CommonUtils.soloautosuma(diccio, part_tab, freal, i, hoja, fin[c])
                    c += 1
            if "catalogos" in r[8]:
                di = r[8]
                validar_catalogo(pregunta, di, hoja)
    
            if totales != "No":
                fn = final_tabla.juntar()
                if not fn:
                    fn =[0]
                CommonUtils.poner_gris(letras_validadas,hoja,fn[0])
    
            print("Todo bien con pregunta ", nter)
            con += 1
        else:
            con+=1
            pass

    return

