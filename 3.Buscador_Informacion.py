import pandas as pd
import numpy as np
from pathlib import Path
import re
import zipfile 
from pathlib import Path
import os
class buscador_informacion:
    def __init__(self, ruta_excel_principal, ruta_raiz, ruta_ministerio, ruta_destino_carpeta_contributivo, ruta_destino_carpeta_subsidiado, ruta_guardar_excel):
        self.ruta_excel_principal = ruta_excel_principal
        self.ruta_raiz = ruta_raiz
        self.ruta_ministerio = ruta_ministerio
        self.ruta_destino_carpeta_contributivo = ruta_destino_carpeta_contributivo
        self.ruta_destino_carpeta_subsidiado = ruta_destino_carpeta_subsidiado
        self.ruta_guardar_excel = ruta_guardar_excel
        # --- ATRIBUTOS PARA MEMORIA (CACHE) ---
        self._df_principal_bruto = None
        self._df_carpetas_escaneadas = None
        self._df_disco_escaneado = None
        self._contributivo_limpio = None
        self._subsidiado_limpio = None
    def _leer_excel_principal(self):
        """Carga el Excel una sola vez en memoria"""
        if self._df_principal_bruto is None:
            self._df_principal_bruto = pd.read_excel(self.ruta_excel_principal)
        return self._df_principal_bruto   
    def contributivo_sistema(self):
        if self._contributivo_limpio is None:
            excel_principal = self._leer_excel_principal().copy()
            opciones = ['CONTRIBUTIVO_1', 'CONTRIBUTIVO_2']
            contributivo_sistema = excel_principal[excel_principal['Contrato'].isin(opciones)].copy()
            #Ordenar valores de acuerdo al orden de la factura
            contributivo_sistema.sort_values(by='NF', inplace=True, ascending=True)
            # Cambiar 'NF' por '# Factura'
            contributivo_sistema.rename(columns={'NF': '# Factura'}, inplace=True)
            columnas_a_eliminar = ['CodigoEmpresa','EPS','Column_1','Column_2','Column_3','Column_4','Column_5','Column_6','Column_7','Contrato']
            self._contributivo_limpio = contributivo_sistema.drop(columnas_a_eliminar, axis=1)
        return self._contributivo_limpio
    def subsidiado_sistema(self):
        if self._subsidiado_limpio is None:
            excel_principal = self._leer_excel_principal().copy()
            # Filtrar por Subsidiado
            opcion = ['SUBSIDIADO_1','SUBSIDIADO_2']
            subsidiado_sistema = excel_principal[excel_principal['Contrato'].isin(opcion)].copy()
            # Ordenar valores de acuerdo al orden de la factura
            subsidiado_sistema.sort_values(by='NF', inplace=True, ascending=True)
            #  Cambiar 'NF' por '# Factura'
            subsidiado_sistema.rename(columns={'NF': '# Factura'}, inplace=True)
            columnas_a_eliminar = ['CodigoEmpresa','EPS','Column_1','Column_2','Column_3','Column_4','Column_5','Column_6','Column_7','Contrato']
            self._subsidiado_limpio = subsidiado_sistema.drop(columnas_a_eliminar,axis=1)
        return self._subsidiado_limpio
    def carpeta_x_facturador(self):
        # Si ya se escaneó el disco, devolver el resultado guardado
        if self._df_carpetas_escaneadas is not None:
            return self._df_carpetas_escaneadas
        ruta_raiz = Path(self.ruta_raiz)
        print("🔍 Escaneando carpetas en el disco... por favor espera.")
        carpetas_encontradas = []

        # 2. Bucle de escaneo
        for carpeta in ruta_raiz.rglob("*"): 
            if carpeta.is_dir():
                # Reemplazar # por la cantidad de digitos que tenga la factura, por ejemplo si es de 7 digitos luego #=7
                match = re.search(r'[a-zA-Z_](\d{#})', carpeta.name)
            
                if match:
                    partes = list(carpeta.parts)
                
                    # El facturador está en la quinta posición (índice 4 si contamos desde 0, 
                    # o índice 5 si la ruta empieza con \\ que cuenta como uno o dos elementos)
                    # Basado en tu conteo de "quintas de izquierda a derecha":
                    try:
                        nombre_facturador = partes[3] 
                    except IndexError:
                        nombre_facturador = "Nivel no disponible"

                    carpetas_encontradas.append({
                        '# Factura': int(match.group(1)),
                        'Facturador': nombre_facturador, 
                        'Ubicacion_Real': str(carpeta),
                        'Niveles': partes 
                    })

        # 3. Creación y expansión del DataFrame
        if carpetas_encontradas:
            df_disco = pd.DataFrame(carpetas_encontradas)
    
            # Expansión de niveles
            df_niveles = pd.DataFrame(df_disco['Niveles'].tolist())
            df_niveles.columns = [f'Nivel_{i}' for i in range(df_niveles.shape[1])]
    
            # Unión final
            df_final = pd.concat([df_disco, df_niveles], axis=1).drop(columns=['Niveles'])
            print('Proceso finalizado con éxito.')
        else:
            print('No se encontraron coincidencias.')
            df_final = pd.DataFrame()
        columnas_a_eliminar=['Nivel_0','Nivel_1','Nivel_2','Nivel_3']
        self._df_carpetas_escaneadas = df_final.drop(columns=columnas_a_eliminar, errors='ignore')
        
        print('✅ Escaneo de disco finalizado.')
        return self._df_carpetas_escaneadas        
    def archivos_x_facturador(self):
        # Si ya se escaneó el disco, devolver el resultado guardado
        if self._df_disco_escaneado is not None:
            return self._df_disco_escaneado
        ruta_raiz = Path(self.ruta_raiz)
        print("🔍 Escaneando carpetas en el disco... por favor espera.")
        resultados = []

        # 3. Búsqueda dirigida
        # Solo recorremos las carpetas una vez, pero buscamos coincidencias activamente
        for archivo in ruta_raiz.rglob("*pdf"):
        
            # Extraemos el número del archivo actual
            # Reemplazar # por la cantidad de digitos que tenga la factura, por ejemplo si es de 7 digitos luego #=7
            match = re.search(r'[a-zA-Z](\d{#})', archivo.name)
        
            # Verificamos si este número está en tu lista de interés

            if match:
        
                partes = list(archivo.parts)
                try:
                    nombre_facturador = partes[3] 
                except IndexError:
                    nombre_facturador = "Nivel no disponible"
                resultados.append({
                    'Carpeta_Padre': archivo.parent.name,
                    'Nombre_Archivo': archivo.name,
                    '# Factura': int(match.group(1)),
                    'Facturador': nombre_facturador,
                    'Ubicacion_Real_(pdf)': str(archivo),
                    'Niveles': partes
                })
        # 3. Creación y expansión del DataFrame
        if resultados:
            df_localizados = pd.DataFrame(resultados)

            # 3. Crear un resumen de conteo por carpeta
            df_conteo = df_localizados.groupby('Carpeta_Padre').size().reset_index(name='Total_Archivos')
            # Expansión de niveles
            df_niveles = pd.DataFrame(df_localizados['Niveles'].tolist())
            df_niveles.columns = [f'Nivel_{i}' for i in range(df_niveles.shape[1])]
    
            # Unión final
            df_final = pd.concat([df_localizados, df_niveles], axis=1).drop(columns=['Niveles'])
            # Unimos el conteo al dataframe final para que veas en cada fila cuántos hay en su carpeta
            df_final = df_localizados.merge(df_conteo, on='Carpeta_Padre')
            print('Proceso finalizado con éxito.')
        else:
            print('No se encontraron coincidencias.')
            df_localizados = pd.DataFrame()
        columnas_a_eliminar=['Nivel_0','Nivel_1','Nivel_2','Nivel_3']
        self._df_disco_escaneado = df_final.drop(columns=columnas_a_eliminar, errors='ignore')
        return self._df_disco_escaneado
    def facturas_contributivo(self, nombre_contributivo=None):
        # 1. Llamamos a los métodos previos para obtener los datos frescos
        # Nota: Asegúrate de que carpeta_x_facturador() devuelva el df_final limpio
        df_carpetas_sistema = self.carpeta_x_facturador()
        df_carpetas_excel = self.contributivo_sistema()
        
        # 2. Verificamos que no estén vacíos para evitar errores en el merge
        if df_carpetas_sistema.empty or df_carpetas_excel.empty:
            print("⚠️ Uno de los conjuntos de datos está vacío. No se puede realizar el cruce.")
            return pd.DataFrame()

        # 3. Realizamos la intersección (Inner Join)
        df_facturas_contributivo = pd.merge(
            df_carpetas_sistema, 
            df_carpetas_excel, 
            on='# Factura', 
            how='inner'
        )
        
        print(f"✅ Cruce exitoso. Se encontraron {len(df_facturas_contributivo)} facturas coincidentes.")
        df_facturas_contributivo = df_facturas_contributivo.rename(columns={
            'Nivel_4': '1',
            'Nivel_5': '2',
            'Nivel_6': '3',
            'Nivel_7': '4',
            'Nivel_8': '5',
            'Fecha':'DD-MM-AAAA'
            })
        df_facturas_contributivo['DD-MM-AAAA'] = pd.to_datetime(df_facturas_contributivo['DD-MM-AAAA']).dt.strftime('%d-%m-%Y')

        # 2. Ordenar de menor a mayor (ascending=True es el valor por defecto)
        df_facturas_contributivo = df_facturas_contributivo.sort_values(by='DD-MM-AAAA', ascending=True)

        # 3. (Opcional) Resetear el índice para que los números de la izquierda vayan de 0 en adelante
        df_facturas_contributivo = df_facturas_contributivo.reset_index(drop=True)
        if nombre_contributivo:
            ruta_final = Path(self.ruta_guardar_excel) / nombre_contributivo
            df_facturas_contributivo.to_excel(ruta_final, index=False)
        return df_facturas_contributivo
    def facturas_subsidiado(self, nombre_subsidiado=None):
        # 1. Llamamos a los métodos previos para obtener los datos frescos
        # Nota: Asegúrate de que carpeta_x_facturador() devuelva el df_final limpio
        df_carpetas_sistema = self.carpeta_x_facturador()
        df_carpetas_excel = self.subsidiado_sistema()
        
        # 2. Verificamos que no estén vacíos para evitar errores en el merge
        if df_carpetas_sistema.empty or df_carpetas_excel.empty:
            print("⚠️ Uno de los conjuntos de datos está vacío. No se puede realizar el cruce.")
            return pd.DataFrame()

        # 3. Realizamos la intersección (Inner Join)
        df_facturas_subsidiado = pd.merge(
            df_carpetas_sistema, 
            df_carpetas_excel, 
            on='# Factura', 
            how='inner'
        )
        
        print(f"✅ Cruce exitoso. Se encontraron {len(df_facturas_subsidiado)} facturas coincidentes.")
        df_facturas_subsidiado = df_facturas_subsidiado.rename(columns={
            'Nivel_4': '1',
            'Nivel_5': '2',
            'Nivel_6': '3',
            'Nivel_7': '4',
            'Nivel_8': '5',
            'Fecha':'DD-MM-AAAA'
            })
        df_facturas_subsidiado['DD-MM-AAAA'] = pd.to_datetime(df_facturas_subsidiado['DD-MM-AAAA']).dt.strftime('%d-%m-%Y')

        # 2. Ordenar de menor a mayor (ascending=True es el valor por defecto)
        df_facturas_subsidiado = df_facturas_subsidiado.sort_values(by='DD-MM-AAAA', ascending=True)

        # 3. (Opcional) Resetear el índice para que los números de la izquierda vayan de 0 en adelante
        df_facturas_subsidiado = df_facturas_subsidiado.reset_index(drop=True)
        if nombre_subsidiado:
            ruta_final = Path(self.ruta_guardar_excel) / nombre_subsidiado
            df_facturas_subsidiado.to_excel(ruta_final, index=False)
        return df_facturas_subsidiado
    def soportes_contributivo(self, nombre_contributivo=None):
        # 1. Llamamos a los métodos previos para obtener los datos frescos
        # Nota: Asegúrate de que carpeta_x_facturador() devuelva el df_final limpio
        df_archivo_sistema = self.archivos_x_facturador()
        df_carpetas_excel = self.contributivo_sistema()
        
        # 2. Verificamos que no estén vacíos para evitar errores en el merge
        if df_archivo_sistema.empty or df_carpetas_excel.empty:
            print("⚠️ Uno de los conjuntos de datos está vacío. No se puede realizar el cruce.")
            return pd.DataFrame()

        # 3. Realizamos la intersección (Inner Join)
        df_soportes_contributivo = pd.merge(
            df_archivo_sistema, 
            df_carpetas_excel, 
            on='# Factura', 
            how='inner'
        )
        
        print(f"✅ Cruce exitoso. Se encontraron {len(df_soportes_contributivo)} facturas coincidentes.")
        df_soportes_contributivo = df_soportes_contributivo.rename(columns={
            'Nivel_4': '1',
            'Nivel_5': '2',
            'Nivel_6': '3',
            'Nivel_7': '4',
            'Nivel_8': '5',
            'Fecha':'DD-MM-AAAA'
            })
        df_soportes_contributivo['DD-MM-AAAA'] = pd.to_datetime(df_soportes_contributivo['DD-MM-AAAA']).dt.strftime('%d-%m-%Y')

        # 2. Ordenar de menor a mayor (ascending=True es el valor por defecto)
        df_soportes_contributivo = df_soportes_contributivo.sort_values(by='DD-MM-AAAA', ascending=True)

        # 3. (Opcional) Resetear el índice para que los números de la izquierda vayan de 0 en adelante
        df_soportes_contributivo = df_soportes_contributivo.reset_index(drop=True)
        if nombre_contributivo:
            ruta_final = Path(self.ruta_guardar_excel) / nombre_contributivo
            df_soportes_contributivo.to_excel(ruta_final, index=False)
        return df_soportes_contributivo
    def soportes_subsidiado(self, nombre_subsidiado=None):
        # 1. Llamamos a los métodos previos para obtener los datos frescos
        # Nota: Asegúrate de que carpeta_x_facturador() devuelva el df_final limpio
        df_archivos_sistema = self.archivos_x_facturador()
        df_carpetas_excel = self.subsidiado_sistema()
        
        # 2. Verificamos que no estén vacíos para evitar errores en el merge
        if df_archivos_sistema.empty or df_carpetas_excel.empty:
            print("⚠️ Uno de los conjuntos de datos está vacío. No se puede realizar el cruce.")
            return pd.DataFrame()

        # 3. Realizamos la intersección (Inner Join)
        df_soportes_subsidiado = pd.merge(
            df_archivos_sistema, 
            df_carpetas_excel, 
            on='# Factura', 
            how='inner'
        )
        
        print(f"✅ Cruce exitoso. Se encontraron {len(df_soportes_subsidiado)} facturas coincidentes.")
        df_soportes_subsidiado = df_soportes_subsidiado.rename(columns={
            'Nivel_4': '1',
            'Nivel_5': '2',
            'Nivel_6': '3',
            'Nivel_7': '4',
            'Nivel_8': '5',
            'Fecha':'DD-MM-AAAA'
            })
        df_soportes_subsidiado['DD-MM-AAAA'] = pd.to_datetime(df_soportes_subsidiado['DD-MM-AAAA']).dt.strftime('%d-%m-%Y')

        # 2. Ordenar de menor a mayor (ascending=True es el valor por defecto)
        df_soportes_subsidiado = df_soportes_subsidiado.sort_values(by='DD-MM-AAAA', ascending=True)

        # 3. (Opcional) Resetear el índice para que los números de la izquierda vayan de 0 en adelante
        df_soportes_subsidiado = df_soportes_subsidiado.reset_index(drop=True)
        if nombre_subsidiado:
            ruta_final = Path(self.ruta_guardar_excel) / nombre_subsidiado
            df_soportes_subsidiado.to_excel(ruta_final, index=False)
        return df_soportes_subsidiado        
    def lista_facturadores_contributivo(self):
        lista_facturadores_contributivo = self.facturas_contributivo()
        return lista_facturadores_contributivo[['Facturador', 'CUsuario']]
    def lista_facturadores_subsidiado(self):
        lista_facturadores_subsidiado = self.facturas_subsidiado()
        return lista_facturadores_subsidiado[['Facturador','CUsuario']]
    def facturas_contributivo_faltantes(self, nombre_contributivo):
        print(f"Se encontraron {len(self.contributivo_sistema())} en el sistema\n En las carpetas se encontraron {len(self.facturas_contributivo())}\n Por lo tanto hacen falta {len(self.contributivo_sistema())-len(self.facturas_contributivo())} correspondientes a contributivo")
        # Nota: Asegúrate de que carpeta_x_facturador() devuelva el df_final limpio)
        df_contributivo_sistema = self.contributivo_sistema()
        facturas_contributivo = self.facturas_contributivo()
        # Realizamos un merge tipo 'outer' o 'right' para traer lo que está en el sistema
        df_complemento = pd.merge(
        df_contributivo_sistema, 
        facturas_contributivo, 
        on='# Factura', 
        how='left', 
        indicator=True
        )

        # Filtramos solo los registros que están 'left_only' 
        df_complemento = df_complemento[df_complemento['_merge'] == 'left_only'].drop(columns=['_merge'])
        columnas_a_eliminar=['Facturador','1','2','3','4','5','DD-MM-AAAA','CUsuario_y']
        df_complemento = df_complemento.drop(columns=columnas_a_eliminar, errors='ignore')
        df_complemento = df_complemento.rename(columns={'CUsuario_x': 'CUsuario'})
        # Hacemos el "merge" (el cruce)
        # Esto solo mantendrá las filas donde el ID coincida en ambos
        contributivo_faltantes = pd.merge(df_complemento, self.lista_facturadores_contributivo(), on='CUsuario')
        contributivo_faltantes = contributivo_faltantes.drop_duplicates()
        print(f"✅ Se encontraron {len(contributivo_faltantes)} facturas faltantes.")
        if self.ruta_guardar_excel:
            ruta_final = Path(self.ruta_guardar_excel) / nombre_contributivo
            contributivo_faltantes.to_excel(ruta_final, index=False)
        return contributivo_faltantes
    def facturas_subsidiado_faltantes(self, nombre_subsidiado):
        print(f"Se encontraron {len(self.subsidiado_sistema())} en el sistema\n En las carpetas se encontraron {len(self.facturas_subsidiado())}\n Por lo tanto hacen falta {len(self.subsidiado_sistema())-len(self.facturas_subsidiado())} correspondientes a subsidiado")
        # Nota: Asegúrate de que carpeta_x_facturador() devuelva el df_final limpio
        df_subsidiado_sistema = self.subsidiado_sistema()
        facturas_subsidiado = self.facturas_subsidiado()
        # Realizamos un merge tipo 'outer' o 'right' para traer lo que está en el sistema
        df_complemento = pd.merge(
        df_subsidiado_sistema, 
        facturas_subsidiado, 
        on='# Factura', 
        how='left', 
        indicator=True
        )

        # Filtramos solo los registros que están 'right_only' 
        # (Es decir, están en el sistema pero no en df_final_limpio)
        df_complemento = df_complemento[df_complemento['_merge'] == 'left_only'].drop(columns=['_merge'])
        columnas_a_eliminar=['Facturador','1','2','3','4','5','DD-MM-AAAA','CUsuario_y']
        df_complemento = df_complemento.drop(columns=columnas_a_eliminar, errors='ignore')
        df_complemento = df_complemento.rename(columns={'CUsuario_x': 'CUsuario'})
        # Hacemos el "merge" (el cruce)
        # Esto solo mantendrá las filas donde el ID coincida en ambos
        subsidiado_faltantes = pd.merge(df_complemento, self.lista_facturadores_subsidiado(), on='CUsuario')
        subsidiado_faltantes = subsidiado_faltantes.drop_duplicates()
        print(f"✅ Se encontraron {len(subsidiado_faltantes)} facturas faltantes.")
        if self.ruta_guardar_excel:
            ruta_final = Path(self.ruta_guardar_excel) / nombre_subsidiado
            subsidiado_faltantes.to_excel(ruta_final, index=False)
        return subsidiado_faltantes
    def archivos_ministerio(self):
        ruta_ministerio = Path(self.ruta_ministerio)
        # Si ya se escaneó el disco, devolver el resultado guardado
        #if self._df_disco_escaneado is not None:
        #    return self._df_disco_escaneado
        resultados = []
        # Definimos el patrón que queremos eliminar
        # 3. Búsqueda dirigida
        # Solo recorremos las carpetas una vez, pero buscamos coincidencias activamente
        for archivo in ruta_ministerio.rglob("*"):

            #Filtramos solo extensiones de interes
            if archivo.suffix.lower() in [".json",".txt"]:
        
                # Reemplazar # por la cantidad de digitos que tenga la factura, por ejemplo si es de 7 digitos luego #=7
                match = re.search(r'([a-zA-Z])(\d{#})', archivo.name)
        
                # Verificamos si este número está en tu lista de interés

                if match:
                    #numero_factura = archivo.parent.name.replace(patron_a_eliminar, "")
                    resultados.append({
                        '# Factura': int(match.group(2)),
                        'Carpeta_Padre': archivo.parent.name,
                        'Nombre_Archivo': archivo.name,
                        'Ubicacion_Real_(json,txt)': str(archivo)
                    })
        # 3. Creación y expansión del DataFrame
        if resultados:
            df_localizados = pd.DataFrame(resultados)
        else:
            print('No se encontraron coincidencias.')
            df_localizados = pd.DataFrame()

        df_localizados = df_localizados.sort_values(by='Carpeta_Padre')
        return df_localizados
    def facturas_famisanar(self):
        # Esta parte tiene que ir en otra función, para poder realizar una clase mas completa que me permita leer los archivos de mejor forma
        df_disco = self._df_disco_escaneado
        df_disco['Detalle'] = df_disco['Nombre_Archivo'].str.extract(r'(FVS)', expand=False)
        df_facturas = df_disco[df_disco['Detalle'].str.contains('FVS', na=False)]
        return df_facturas
    def detectar_errores_contributivo(self, nombre_archivo):
        facturas_contributivo = self.facturas_contributivo()
        facturas_contributivo_niveles = facturas_contributivo[['1','2','3','4','5']]
        # Tomamos la última columna del DataFrame
        columna_menos_uno = facturas_contributivo_niveles.columns[-1]
        columna_menos_dos = facturas_contributivo_niveles.columns[-2]
        columna_menos_tres = facturas_contributivo_niveles.columns[-3]
        # Definir el patrón de la función
        patron = re.compile(r"(\d+)(\s{1,})(.+)$")
        
        errores_1 = []
        for valor in facturas_contributivo_niveles[columna_menos_uno]:
            match = patron.search(str(valor))
            if match:
                errores_1.append({
                    "# Factura":int(match.group(1)),
                    "comentario":match.group(3),
                })
            else:
                errores_1.append({
                    "# Factura":000000,
                    "comentario":""
                })
        errores_2 = []
        for valor in facturas_contributivo_niveles[columna_menos_dos]:
            match = patron.search(str(valor))
            if match:
                errores_2.append({
                    "# Factura":int(match.group(1)),
                    "comentario":match.group(3),
                })
            else:
                errores_2.append({
                    "# Factura":000000,
                    "comentario":""
                })
        errores_3 = []
        for valor in facturas_contributivo_niveles[columna_menos_tres]:
            match = patron.search(str(valor))
            if match:
                errores_3.append({
                    "# Factura":int(match.group(1)),
                    "comentario":match.group(3),
                })
            else:
                errores_3.append({
                    "# Factura":000000,
                    "comentario":""
                })
        df_carpetas_errores = pd.concat([pd.DataFrame(errores_1),pd.DataFrame(errores_2),pd.DataFrame(errores_3)])
        df_carpetas_errores_contributivo = pd.merge(
           df_carpetas_errores,
           facturas_contributivo,
           on='# Factura',
           how='inner' 
        )
        df_carpetas_errores_contributivo_ = df_carpetas_errores_contributivo[['# Factura','comentario','Facturador','DD-MM-AAAA']]
        if self.ruta_guardar_excel:
            ruta_final = Path(self.ruta_guardar_excel) / nombre_archivo
            df_carpetas_errores_contributivo_.to_excel(ruta_final, index=False)
        return df_carpetas_errores_contributivo_
    def detectar_errores_subsidiado(self, nombre_archivo):
        facturas_subsidiado = self.facturas_subsidiado()
        facturas_subsidiado_niveles = facturas_subsidiado[['1','2','3','4','5']]
        # Tomamos la última columna del dataframe
        columna_menos_uno = facturas_subsidiado_niveles.columns[-1]
        columna_menos_dos = facturas_subsidiado_niveles.columns[-2]
        columna_menos_tres = facturas_subsidiado_niveles.columns[-3]

        # Definir el patrón de la función
        patron = re.compile(r"(\d+)(\s{1,})(.+)$")
        errores_1=[]
        for valor in facturas_subsidiado_niveles[columna_menos_uno]:
            match = patron.search(str(valor))
            if match:
                errores_1.append({
                    "# Factura":int(match.group(1)),
                    "comentario":match.group(3),
                })
        errores_2 = []
        for valor in facturas_subsidiado_niveles[columna_menos_dos]:
            match = patron.search(str(valor))
            if match:
                errores_2.append({
                    "# Factura":int(match.group(1)),
                    "comentario":match.group(3),
                })
        errores_3 = []
        for valor in facturas_subsidiado_niveles[columna_menos_tres]:
            match = patron.search(str(valor))
            if match:
                errores_3.append({
                    "# Factura":int(match.group(1)),
                    "comentario":match.group(3),
                })
        df_carpetas_errores = pd.concat([pd.DataFrame(errores_1),pd.DataFrame(errores_2),pd.DataFrame(errores_3)])
        df_carpetas_errores_subsidiado = pd.merge(
           df_carpetas_errores,
           facturas_subsidiado,
           on='# Factura',
           how='inner' 
        )
        df_carpetas_errores_subsidiado_ = df_carpetas_errores_subsidiado[['# Factura','comentario','Facturador','DD-MM-AAAA']]
        if self.ruta_guardar_excel:
            ruta_final = Path(self.ruta_guardar_excel) / nombre_archivo
            df_carpetas_errores_subsidiado_.to_excel(ruta_final, index=False)
        return df_carpetas_errores_subsidiado_
    # La siguiente lógica se extrae a través del comportamiento de los archivos
    def buscador_carpeta_sin_OTR_contributivo(self, nombre_archivo):
        df_soportes_contributivo = self.soportes_contributivo()
        # Reemplazar # por la cantidad de digitos que tenga la factura, por ejemplo si es de 7 digitos luego #=7
        df_filtrado_contributivo = df_soportes_contributivo[df_soportes_contributivo['Nombre_Archivo'].str.contains(r'AUT_\d{#}')]
        if self.ruta_guardar_excel:
            ruta_final = Path(self.ruta_guardar_excel) / nombre_archivo
            df_filtrado_contributivo.to_excel(ruta_final, index=False)
        return df_filtrado_contributivo
    def buscador_carpeta_sin_OTR_subsidiado(self, nombre_archivo):
        df_soportes_subsidiado = self.soportes_subsidiado()
        # Reemplazar # por la cantidad de digitos que tenga la factura, por ejemplo si es de 7 digitos luego #=7
        df_filtrado_subsidiado = df_soportes_subsidiado[df_soportes_subsidiado['Nombre_Archivo'].str.contains(r'AUT_\d{#}')]
        if self.ruta_guardar_excel:
            ruta_final = Path(self.ruta_guardar_excel) / nombre_archivo
            df_filtrado_subsidiado.to_excel(ruta_final, index=False)
        return df_filtrado_subsidiado

Buscador_Informacion = buscador_informacion(
    ruta_excel_principal=r"C:\Usuario\Numero_Usuario\Nube\Carpeta_donde_esta_guardados_los_Archivos\...\DANIEL\Daniel_Excel\2026\MARZO\GUATAVITA\18_al_25.xlsx", 
    ruta_raiz=r"\\Servidor\Carpeta_del_servidor_donde_Se_guardan_los_Archivos\...\2026\MARZO",
    ruta_ministerio=r"\\Servidor\Carpeta_del_servidor_donde_Se_guardan_los_Archivos\...\MINISTERIO\FMS 2026\MARZO\17-MARZO-2026",
    ruta_guardar_excel=r"C:\Usuario\Numero_Usuario\Nube\Carpeta_donde_esta_guardados_los_Archivos\...\DANIEL\Daniel_Excel\2026\MARZO\GUATAVITA")
Facturas_contributivo = Buscador_Informacion.facturas_contributivo('Eps_inicio_al_final_contributivo.xlsx')
Facturas_subsidiado = Buscador_Informacion.facturas_subsidiado('Eps_inicio_al_final_subsidiado.xlsx')
Facturas_sin_carpeta_contributivo = Buscador_Informacion.facturas_contributivo_faltantes("Facturas_sin_carpeta_contributivo_inicio_al_final.xlsx")
Facturas_sin_carpeta_subsidiado = Buscador_Informacion.facturas_subsidiado_faltantes("Facturas_sin_carpeta_subsidiado_inicio_al_final.xlsx")
Detectar_errores_contributivo = Buscador_Informacion.detectar_errores_contributivo('Errores_carpetas_contributivo_inicio_al_final.xlsx')
Detectar_errores_subsidiado = Buscador_Informacion.detectar_errores_subsidiado('Errores_carpetas_subsidiado_inicio_al_final.xlsx')
Buscador_sin_OTR_contributivo = Buscador_Informacion.buscador_carpeta_sin_OTR_contributivo('Sin_OTR_Contributivo_inicio_al_final.xlsx')
Buscador_sin_OTR_subsidiado = Buscador_Informacion.buscador_carpeta_sin_OTR_subsidiado('Sin_OTR_Subsidiado_inicio_al_final.xlsx')
Archivos_contributivo = Buscador_Informacion.soportes_contributivo('Archivos_contributivo_inicio_al_final.xlsx')
Archivos_subsidiado = Buscador_Informacion.soportes_subsidiado('Archivos_Subsidiado_inicio_al_final.xlsx')