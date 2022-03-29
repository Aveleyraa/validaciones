import string

class CommonUtils_encontrar:
    @staticmethod
    def genp():
        """
        regresa res que es una lista con letras mayusculas y minusculas 
        para buscar coincidencias en el documento excel (las marcas)
        ejemplo ["A1","b1"...]
        """
        abc1 = list(string.ascii_uppercase)
        abc = list(string.ascii_lowercase)
        nn = ['Ñ','ñ'] 
        l = abc1 + abc + nn
        res = []
        for i in l:
            for d in range(1,10):
                res.append(i+str(d))

        return res

    
    @staticmethod
    def filtro(cadena):
        """


        Parameters
        ----------
        cadena : str 

        Returns
        -------
        str
            Regresa una string dependiendo el signo que encuentre en la cadena
            que se pasa como argumento de entrada. En caso de que no detecte
            ninguno, regresa ref, que ace alución a las marcas que son referentes
            y no necesitan ser comparados

        """
        a = '>'
        b = '<'
        c = '='
        if a in cadena:
            return 'mayor'

        if b in cadena:
            return 'menor'

        if c in cadena:
            return 'igual'

        else:
            return 'ref'

    @staticmethod
    def imagen(sec,datos):
        """


        Parameters
        ----------
        sec : str
            Es el nombre de la hoja de excel que se va a leer.
        datos : dataframe pandas
            es la hoja de excel a leer, expresada en un dataframe de pandas.

        Returns
        -------
        entot : dic
            regresa un diccionario que contiene filas, columnas y sección
            donde se encontraron las letras de variable "posibles". además 
            contiene dentro de sí, otros diccionarios referentes alas letras
            que encontró, determianndo si son referentes o comparadores,
            la id a la que hacen referencia los comparadores, así como su
            operacion de acuerdo al signo detectado dentro de la string.

        """
        posibles = CommonUtils_encontrar.genp() + ['W'] #variable con la lista para buscar letras con numeros en el documento
        mdf = datos
        reg = 0
        cont = 0
        entot = {'fila':[],'columna':[],
                'sec':[]}
        asig = {}
        asig1 = {}
        asid = {}
        for re in posibles:
            asig[re] = []
            asid[re+'id'] = []
            asig1[re+'Op'] = []
        for lista in mdf:
            vo = 0
            for col in datos[lista]:

                for l in posibles:

                    try:
                        t = col.split('/')
                        for i in t:

                            if l == i:
                                asig[l] = [reg,*asig[l]]
                                asid[l+'id'] = [i,*asid[l+'id']]
                                asig1[l+'Op'] = [CommonUtils_encontrar.filtro(i),*asig1[l+'Op']]
                            if i.startswith(l) and i!=l:
                                asig[l].append(reg)
                                asid[l+'id'].append(i)
                                asig1[l+'Op'].append(CommonUtils_encontrar.filtro(i))
                            if l in i:
                                entot['sec'].append(sec)
                                entot['fila'].append(vo)
                                entot['columna'].append(cont)
                                reg += 1

                    except:
                        pass
                vo+=1
            cont+=1
        borrar = [val for val in asig if not asig[val]]
        for i in borrar:
            del asig[i],asig1[i+'Op'],asid[i+'id']

        entot['asig'] = asig
        entot['asig1'] = asig1
        entot['asid'] = asid

        return entot

    @staticmethod
    def cordenadas(sec,datos):
        """


        Parameters
        ----------
        sec : str
            nombre de la hoja de excel a leer.
        datos : dataframe de pandas
            la hoja de excel a leer expresada en dataframe de pandas.

        Returns
        -------
        d_salida : dic
            Regresa diccionario con las coordenadas expresadas a manera de 
            documento en excel, ejemplo "A25" para referir a una celda que
            está en columna A, fila 25. También tiene la sección donde está
            sacando esas coordenadas y la id de las macas que encontró. 
            La id en este caso es la marca como se ha puesto en el documento,
            por ejemplo "a1.1>"

        """
        abc1 = list(string.ascii_uppercase)+['AA','AB','AC','AD','AE']
        d = CommonUtils_encontrar.imagen(sec,datos)
        d_salida = {}
        sal = []
        c = 0
        for n in d['fila']:
            try:
                cor = abc1[d['columna'][c]]+str(n+2)
                sal.append(cor)
            except:
                sal.append('Fuera de margen')
            c +=1
        for ele in d['asig']:
            d_salida[ele] = []
            d_salida[ele+'sec'] = []
            # d_salida[ele+'id'] = []
            for v in d['asig'][ele]:
                d_salida[ele].append(sal[v]) 
                d_salida[ele+'sec'].append(d['sec'][v]) 
                # d_salida[ele+'id'].append(d['asid'][v])
        d_salida.update(d['asig1'])
        d_salida.update(d['asid'])

        return d_salida

    @staticmethod
    def nframe(di):
        "esta funcion ordena el diccionario que da de salida la funcion coordenadas"
        b = {'seccion':[],'coordenada':[],
            'comparacion':[],'operacion':[],'ID':[]}
        for k in di:
            if 'sec' not in k and 'Op' not in k and 'id' not in k: #se itera solo para el elemento en el diccionario que corresponde a la letra buscada en la marca, no a su id, ni secion ni operacion
                c = 0
                for val in di[k]:

                    b['coordenada'].append(val)
                    b['comparacion'].append(k)
                    b['operacion'].append(di[k+'Op'][c])
                    b['seccion'].append(di[k+'sec'][c])
                    b['ID'].append(di[k+'id'][c])
                    c += 1

        return b

    @staticmethod
    def determinar(cadena):
        """


        Parameters
        ----------
        cadena : str
            operacion que se va a realizar para comparar.

        Returns
        -------
        str
            elementos para realizar formula que se escribirá en excel. 

        """
        if cadena == 'menor':
            return '<='
        if cadena == 'mayor':
            return '>='
        if cadena == 'igual':
            return '='
        else:
            return 'posible mala referencia'

    @staticmethod
    def clasif(cadena):
        """


        Parameters
        ----------
        cadena : str
            operacion que se va a realizar para comparar.

        Returns
        -------
        str
        regresarlo en palabras es importante para el proceso de revision.

        """
        if '<' in cadena:
            return 'menor'
        if '>' in cadena:
            return 'mayor'
        if '=' in cadena:
            return 'igual'
        else:
            return 'ref'

    @staticmethod
    def formulaS(cadena,seccion):
        """


        Parameters
        ----------
        cadena : str
            cadena con más de una coordenada.
        seccion : str
            nombre de la hoja de excel.

        Returns
        -------
        formula : str
            Dado que hay más de una coordenada, se tiene que hacer un proceso
            adicional para generar una formula que nos dé el valor concreto
            de la suma de los valores que se detecten en las coordenadas
            que están siendo pasadas como argumento de esta funcion. Regresa
            por lo tanto una formula que hace eso.

        """
        if seccion != '':
            ad = seccion+'!'
        else:
            ad = seccion
        c = cadena.split(',')
        r = f'COUNTIF({ad}{c[0]}:{c[-1]},"NS")'
        ca = f'{ad}{c[0]}:{c[-1]}'
        bl = f'ISBLANK({ad}{c[0]})'
        o = f'COUNTIF({ad}{c[0]}:{c[-1]},"NA")'
        for co in c[1:]: 
            bl += f',ISBLANK({ad}{co})' #Esto porque es blanco solo funciona con una celda
        #metodo para coordenadas si no fueran continuos
        # r = f'COUNTIF({ad}{c[0]},"NS")'
        # ca = f'{ad}{c[0]}'
        # bl = f'ISBLANK({ad}{c[0]})'
        # o = f'COUNTIF({ad}{c[0]},"NA")'
        # for co in c[1:]: 
        #     r += f'+COUNTIF({ad}{co},"NS")'
        #     ca += f',{ad}{co}'
        #     bl += f',ISBLANK({ad}{co})'
        #     o += f'+COUNTIF({ad}{co},"NA")'

        formula = f'IF(AND(SUM({ca})=0,{r}>0),"NS",IF(AND(SUM({ca})=0,{o}>0),"NA",IF(AND({bl}),"",SUM({ca}))))'
        return formula

    @staticmethod
    def getnum(cad):
        "regresa el numero de fila de una coordenada de excel"
        numero = ''
        for caracter in cad:
            if caracter.isnumeric():
                numero += caracter
        numero = int(numero)
        return numero

    @staticmethod
    def sumco(co,num):
        """


        Parameters
        ----------
        co : str
            coordenada de excel ejemplo A25.
        num : int
            numero que se va a sumar a la fila de la coordenada.

        Returns
        -------
        cor : str
            nueva coordenda con el num sumado a la fila de la coordenada de 
            entrada.

        """
        letra = ''
        fila = CommonUtils_encontrar.getnum(co)
        for caracter in co:
            if caracter.isalpha():
                letra += caracter
        cor = f'{letra}{fila+num}'
        return cor

    @staticmethod
    def columnas(unicos,base,secc):
        """


        Parameters
        ----------
        unicos : list
            Lista de valores unicos en donde se detectó existencia 
            del caracter ":"
        base : dic
            Diccionario donde están conenidos todos los elementos registrados
            de una hoja de excel.
        secc : str
            nombre de la hoja de excel.

        Returns
        -------
        base : dic
            Regresa el diccionario de entrada pero modificado ya que agrega 
            las celdas contenidas en las columnas marcadas. Las marcas con ":"
            representan una columna a comparar con otra, donde en realidad cada
            fila de esa columna tiene que ser comparada con la fila de otra.
            Por esa razón se generan las referencias necesarias a cada fila dentro
            de las columnas que fueron marcadas. Además, se hace el borrado del
            caracter ":" para no generar errores en los procesos siguentes
            de creación de formulas.

        """
        for columna in unicos:

            seccion = secc
            coord = []
            ide = ''.join(l for l in columna if l != ':')
            op = CommonUtils_encontrar.clasif(columna)
            c = 0
            indices = []
            for ele in base['ID']:

                if ele == columna:
                    coord.append(base['coordenada'][c])
                    indices.append(c)
                c += 1
            cor = coord[0]
            a1 = CommonUtils_encontrar.getnum(coord[0])
            a2 = CommonUtils_encontrar.getnum(coord[1])
            resta = a2-a1
            ide1 = ide
            if '.' in ide:
                iterar = ide.split('.')
                n = iterar[0]
                ide1 = n

            d = {'seccion':seccion,'coordenada':cor,
                'comparacion':ide1,'operacion':op,'ID':ide}
            integrar = [d]
            for i in range(1,resta+1):
                e = {'seccion':seccion,'coordenada':cor,
                    'comparacion':ide,'operacion':op,'ID':ide}
                e['coordenada'] = CommonUtils_encontrar.sumco(coord[0],i)
                e['ID'] = ide+str(i)
                if '.' in e['comparacion']:
                    e['ID'] = ide
                    iterar = e['comparacion'].split('.')
                    n = iterar[0]
                    e['comparacion'] = n + str(i)
                else:
                    e['comparacion'] = ide+str(i)
                integrar.append(e)
            for ind in reversed(indices):
                for lla in base:
                    base[lla].pop(ind)
            for fila in reversed(integrar):

                for ke in fila:
                    base[ke].insert(indices[0],fila[ke])
        return base

    @staticmethod
    def p_rel(guia, excel, pagina):
        """


        Parameters
        ----------
        guia : dataframe de pandas, con los elementos tal y 
        como los genera proceso llamado "encontrar.py"

        excel : la página del archivo excel a validar con openpyxl.

        pagina : nombre de la página del excel a validar.

        Returns
        -------
        None.

        """
        orden = CommonUtils_encontrar.categorizar(guia,pagina)
    
        for w in orden:
            CommonUtils_encontrar.escribir(w,excel)
    
        return

    @staticmethod
    def escribir(lista,excel):
        "escribe la formula contenida en la lista en el excel de openpyxl"
        fila = lista[0]
        columna = 32 #porque columna AF en excel es 32
        for formula in lista[1:]:
            excel.cell(row=fila,column=columna,value=formula)
            columna += 1
        return

    @staticmethod
    def categorizar(datos,pag):
        """


        Parameters
        ----------
        datos : dataframe de pandas.
        pag : str
            nombre de la hoja de excel en la que se va a trabajar.

        Returns
        -------
        rl : list
            Lista de listas por cada "W" encontrada con las coordenadas 
            ordenadas y la fila en donde se debe escribir.

        """
        rl = []
        mdf = datos.loc[datos['seccion']==pag]
        w = mdf.loc[mdf['ID']=='W']
        rangoW = []
        for val in w['coordenada']:
            fila = ''
            for caracter in val:
                if caracter.isnumeric():
                    fila += caracter
            rangoW.append(int(fila))
        rangoW.sort()
        mdf1 = mdf.loc[mdf['ID']!='W']
        mdf1 = mdf1.loc[mdf1['operacion']!='ref']
        numeros = []
        for val in mdf1['coordenada']:
            if ',' in val: #para omitir sumas
                pass
            else:
                fila = ''
                for caracter in val:
                    if caracter.isnumeric():
                        fila += caracter
                numeros.append(int(fila))

        ldel = []
        for rango in rangoW:
            ap = []
            for numero in numeros: #conseguir numeros para los rangos
                if numero > rango:
                    ap.append(numero)
            ap.sort()
            ldel.append(ap)
        for lista in ldel: #depurar numeros de los rangos
            c = 0
            for val in lista[1:]:
                resta = val - lista[c]
                if resta > 7:
                    lista.remove(val)

        mdf1 = mdf1.reset_index(drop=True)

        medidor = 0
        for lista in ldel:

            sul = []
            sul.append(rangoW[medidor])
            comparar = list(set(lista))
            fila = 0
            for val in mdf1['coordenada']:

                for co in comparar:
                    if val.endswith(str(co)):
                        sul.append(mdf1['formulas'][fila])
                fila += 1
            rl.append(sul)
            medidor += 1

        return rl