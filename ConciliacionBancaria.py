import pandas as pd
from InformeConciliacionBancaria import InformeConciliacionBancaria


class ConciliacionBancaria:

    def __init__(
            self,
            pathArchivoExtracto,
            pathArchivoContable,
            nombreArchivoGenerado
                 ):
        self.nombreArchivoGenerado = nombreArchivoGenerado
        # Cargar datos desde archivos CSV (o cualquier otro formato compatible con pandas)
        self.datos_bancarios = pd.read_excel(pathArchivoExtracto+'.xlsx')
        self.registros_contables = pd.read_excel(pathArchivoContable+'.xlsx')

        self.banco_entradas = None
        self.banco_salidas = None
        self.contable_entradas = None
        self.contable_salidas = None

        self.banco_entradasRepeditos = None
        self.banco_salidasRepeditos = None
        self.duplicados_entradas = []
        self.duplicados_salidas = []
        self.duplicados_entradas_resumen=None
        self.duplicados_salidas_resumen=None

        self.elementosConciliados = {
            'conciliado_entradas': None,
            'conciliado_salidas': None,
            'duplicado_entradas': None,
            'duplicado_salidas': None,
        }

        self.partidasConciliatorias = {
            'partidaConciliatoriaEntradas_bancarias': None,
            'partidaConciliatoriaSalidas_bancarias': None,
            'partidaConciliatoriaEntradas_contables': None,
            'partidaConciliatoriaSalidas_contables': None
        }

        self.informacionEmpresa = {
            'razonSocial': 'Empresa de Pruebas',
            'nit': '800.006.126',
            'periodo': '2023/12'
        }

    def definir_id(self):
        # Asegúrate de tener la columna 'fecha' como objetos datetime, si no está en este formato
        self.datos_bancarios['fecha'] = pd.to_datetime(self.datos_bancarios['fecha'])
        self.registros_contables['fecha'] = pd.to_datetime(self.registros_contables['fecha'])

        # Extraer el día y el mes
        self.datos_bancarios['dia'] = self.datos_bancarios['fecha'].dt.day
        self.datos_bancarios['mes'] = self.datos_bancarios['fecha'].dt.month

        self.registros_contables['dia'] = self.registros_contables['fecha'].dt.day
        self.registros_contables['mes'] = self.registros_contables['fecha'].dt.month

        # Crear la tercera columna combinando día, mes y número
        self.datos_bancarios.loc[self.datos_bancarios['fecha'].dt.day > 0,'id'] = self.datos_bancarios['dia'].astype(str) + self.datos_bancarios['mes'].astype(str) + self.datos_bancarios['valor'].astype(str)
        self.registros_contables.loc[self.registros_contables['fecha'].dt.day > 0,'id'] = self.registros_contables['dia'].astype(str) + self.registros_contables['mes'].astype(str) + self.registros_contables['valor'].astype(str)


        self.banco_entradas = self.datos_bancarios[self.datos_bancarios['tipo'] == 2]
        self.banco_salidas = self.datos_bancarios[self.datos_bancarios['tipo'] == 1]
        self.contable_entradas = self.registros_contables[self.registros_contables['tipo'] == 2]
        self.contable_salidas = self.registros_contables[self.registros_contables['tipo'] == 1]

    def sin_definir_id(self):
        # Asegúrate de tener la columna 'fecha' como objetos datetime, si no está en este formato
        self.datos_bancarios['fecha'] = pd.to_datetime(self.datos_bancarios['fecha'])
        self.registros_contables['fecha'] = pd.to_datetime(self.registros_contables['fecha'])

        # Extraer el día y el mes
        self.datos_bancarios['dia'] = self.datos_bancarios['fecha'].dt.day
        self.datos_bancarios['mes'] = self.datos_bancarios['fecha'].dt.month

        self.registros_contables['dia'] = self.registros_contables['fecha'].dt.day
        self.registros_contables['mes'] = self.registros_contables['fecha'].dt.month

        # Crear la tercera columna combinando día, mes y número
        self.datos_bancarios.loc[self.datos_bancarios['fecha'].dt.day > 0,'id'] =  self.datos_bancarios['valor'].astype(str)
        self.registros_contables.loc[self.registros_contables['fecha'].dt.day > 0,'id'] = self.registros_contables['valor'].astype(str)


        self.banco_entradas = self.datos_bancarios[self.datos_bancarios['tipo'] == 2]
        self.banco_salidas = self.datos_bancarios[self.datos_bancarios['tipo'] == 1]
        self.contable_entradas = self.registros_contables[self.registros_contables['tipo'] == 2]
        self.contable_salidas = self.registros_contables[self.registros_contables['tipo'] == 1]

    def modular_duplicados(self):
        self.contable_entradas = self.contable_entradas.copy()
        self.banco_entradas = self.banco_entradas.copy()

        # REGISTROS DUPLICADOS --------------------------------------------------------
        #contables - entradas
        duplicados = self.contable_entradas['id'].duplicated(keep=False)
        elementosRepetidosEntradasContable = self.contable_entradas[duplicados]
        #contables - salidas
        duplicados_sal = self.contable_salidas['id'].duplicated(keep=False)
        elementosRepetidosSalidasContable = self.contable_salidas[duplicados_sal]
        #bancos - entradas
        duplicados_banco = self.banco_entradas['id'].duplicated(keep=False)
        elementosRepetidosEntradasBancos = self.banco_entradas[duplicados_banco]
        #bancos - salidas
        duplicados_banco_s = self.banco_salidas['id'].duplicated(keep=False)
        elementosRepetidosSalidasBancos = self.banco_salidas[duplicados_banco_s]

        #REPETIDOS
        #contar las veces que esta repetido cada id   CONTABLE ENTRADAS
        frecuencia_ent = elementosRepetidosEntradasContable['id'].value_counts()
        # Crear una copia del DataFrame antes de realizar la asignación
        elementosRepetidosEntradasContable = elementosRepetidosEntradasContable.copy()
        elementosRepetidosEntradasContable.loc[elementosRepetidosEntradasContable['fecha'].dt.day > 0, "estado"] = elementosRepetidosEntradasContable['id'].map(frecuencia_ent)
        #contar las veces que esta repetido cada id   CONTABLE SALIDAS
        frecuencia_sal = elementosRepetidosSalidasContable['id'].value_counts()
        # Crear una copia del DataFrame antes de realizar la asignación
        elementosRepetidosSalidasContable = elementosRepetidosSalidasContable.copy()
        elementosRepetidosSalidasContable.loc[elementosRepetidosSalidasContable['fecha'].dt.day > 0, "estado"] = elementosRepetidosSalidasContable['id'].map(frecuencia_sal)

        #contar las veces que esta repetido cada id BANCARIO ENTRADAS
        frecuenciaa = elementosRepetidosEntradasBancos['id'].value_counts()
        # Crear una copia del DataFrame antes de realizar la asignación
        elementosRepetidosEntradasBancos = elementosRepetidosEntradasBancos.copy()
        elementosRepetidosEntradasBancos.loc[elementosRepetidosEntradasBancos['fecha'].dt.day > 0, "estado"] = elementosRepetidosEntradasBancos['id'].map(frecuenciaa)
        #print(elementosRepetidosEntradasBancos)
        #contar las veces que esta repetido cada id BANCARIO SALIDAS
        frecuencia_salBan = elementosRepetidosSalidasBancos['id'].value_counts()
        # Crear una copia del DataFrame antes de realizar la asignación
        elementosRepetidosSalidasBancos = elementosRepetidosSalidasBancos.copy()
        elementosRepetidosSalidasBancos.loc[elementosRepetidosSalidasBancos['fecha'].dt.day > 0, "estado"] = elementosRepetidosSalidasBancos['id'].map(frecuencia_salBan)


        # Obtener la unión de los índices de frecuencia_ent y frecuenciaa
        todos_los_indices_entradas = frecuencia_ent.index.union(frecuenciaa.index)
        todos_los_indices_salidas = frecuencia_sal.index.union(frecuencia_salBan.index)

        # Comparar frecuencias entre contable entradas y bancos entradas
        comparacion_frecuencias_entradas = pd.DataFrame({
            'ID': todos_los_indices_entradas,
            'Frecuencia_Contable_Entradas': frecuencia_ent.reindex(todos_los_indices_entradas, fill_value=0).values,
            'Frecuencia_Bancos_Entradas': frecuenciaa.reindex(todos_los_indices_entradas, fill_value=0).values
        })
        #print(comparacion_frecuencias_entradas)
        comparacion_frecuencias_salidas = pd.DataFrame({
            'ID': todos_los_indices_salidas,
            'Frecuencia_Contable_Salidas': frecuencia_sal.reindex(todos_los_indices_salidas, fill_value=0).values,
            'Frecuencia_Bancos_Salidas': frecuencia_salBan.reindex(todos_los_indices_salidas, fill_value=0).values
        })

        # Filtrar los casos donde las frecuencias son iguales
        frecuencias_coincidentes_entradas = comparacion_frecuencias_entradas[
            comparacion_frecuencias_entradas['Frecuencia_Contable_Entradas'] == comparacion_frecuencias_entradas[
                'Frecuencia_Bancos_Entradas']]
        frecuencias_coincidentes_salidas = comparacion_frecuencias_salidas[
            comparacion_frecuencias_salidas['Frecuencia_Contable_Salidas'] == comparacion_frecuencias_salidas[
                'Frecuencia_Bancos_Salidas']]
        #print(comparacion_frecuencias_salidas)
        # Filtrar los casos donde las frecuencias son iguales
        frecuencias_NoCoincidentes_entradas = comparacion_frecuencias_entradas[
            comparacion_frecuencias_entradas['Frecuencia_Contable_Entradas'] != comparacion_frecuencias_entradas[
                'Frecuencia_Bancos_Entradas']]
        frecuencias_NoCoincidentes_salidas = comparacion_frecuencias_salidas[
            comparacion_frecuencias_salidas['Frecuencia_Contable_Salidas'] != comparacion_frecuencias_salidas[
                'Frecuencia_Bancos_Salidas']]
        self.duplicados_entradas_resumen = frecuencias_NoCoincidentes_entradas
        self.duplicados_salidas_resumen = frecuencias_NoCoincidentes_salidas
        # Mostrar los casos donde las frecuencias son iguales
        #print(frecuencias_coincidentes_entradas)
        #print(frecuencias_NoCoincidentes_entradas)
        #print(frecuencias_coincidentes_salidas)
        #print(frecuencias_NoCoincidentes_salidas)

        # Crear DataFrames con las filas que no coinciden entradas
        for idx, row in frecuencias_NoCoincidentes_entradas.iterrows():
            id_actual = row['ID']
            #print(row['Frecuencia_Contable_Entradas'], row['Frecuencia_Bancos_Entradas'])
            # Filtrar las entradas contables por el ID actual
            contable_filtrado = self.contable_entradas[self.contable_entradas['id'] == id_actual]
            # Filtrar las entradas de bancos por el ID actual
            bancos_filtrado = self.banco_entradas[self.banco_entradas['id'] == id_actual]

            #print(len(contable_filtrado))
            #print(len(bancos_filtrado))
            if(len(contable_filtrado) > 0 and len(bancos_filtrado)>0):
                # Eliminar las filas que coinciden en 'id' con entradas contables
                self.contable_entradas = self.contable_entradas[self.contable_entradas['id'] != id_actual]

                # Eliminar las filas que coinciden en 'id' con entradas de bancos
                self.banco_entradas = self.banco_entradas[self.banco_entradas['id'] != id_actual]
                self.duplicados_entradas.append(
                    {id_actual: {'contable': contable_filtrado, 'bancos': bancos_filtrado}})

        # Crear DataFrames con las filas que no coinciden  salidas
        for idx, row in frecuencias_NoCoincidentes_salidas.iterrows():
            id_actual = row['ID']
            #print(row)
            # Filtrar las entradas contables por el ID actual
            contable_filtrado = self.contable_salidas[self.contable_salidas['id'] == id_actual]
            # Filtrar las entradas de bancos por el ID actual
            bancos_filtrado = self.banco_salidas[self.banco_salidas['id'] == id_actual]
            #print(contable_filtrado, bancos_filtrado)
            if (len(contable_filtrado) > 0 and len(bancos_filtrado) > 0):
                # Eliminar las filas que coinciden en 'id' con entradas contables
                self.contable_salidas = self.contable_salidas[self.contable_salidas['id'] != id_actual]

                # Eliminar las filas que coinciden en 'id' con entradas de bancos
                self.banco_salidas = self.banco_salidas[self.banco_salidas['id'] != id_actual]

                self.duplicados_salidas.append(
                    {id_actual: {'contable': contable_filtrado, 'bancos': bancos_filtrado}})

        #print(self.duplicados_entradas)
        #print(self.duplicados_entradas)
        """ 
        if elementosRepetidosEntradasContable.shape[0] != elementosRepetidosEntradasBancos.shape[0] :
            # Realizar una fusión (merge) basada en la columna 'B'
            merged_contable = pd.merge(self.contable_entradas, elementosRepetidosEntradasContable, on='id', how='left', indicator=True).copy()
            merged_bancos = pd.merge(self.banco_entradas, elementosRepetidosEntradasBancos, on='id', how='left', indicator=True).copy()
            # Seleccionar solo las filas de df1 donde no hay coincidencias en 'B' con df2
            self.contable_entradas = merged_contable[merged_contable['_merge'] == 'left_only'].drop(columns=['_merge','fecha_y','descripcion_y','valor_y','tipo_y','dia_y','mes_y','estado']).copy()
            self.banco_entradas = merged_bancos[merged_bancos['_merge'] == 'left_only'].drop(columns=['_merge','fecha_y','descripcion_y','valor_y','tipo_y','dia_y','mes_y','estado']).copy()
            #setear los repetidos
            self.contable_entradasRepeditos = elementosRepetidosEntradasContable
            self.banco_entradasRepeditos = elementosRepetidosEntradasBancos

            # Renombrar las cabeceras de los dataframes
            nuevos_nombres = {'valor_x': 'valor', 'tipo_x': 'tipo', 'descripcion_x': 'descripcion', 'fecha_x': 'fecha',
                              'dia_x': 'dia', 'mes_x': 'mes'}
            self.contable_entradas.rename(columns=nuevos_nombres, inplace=True)
            self.banco_entradas.rename(columns=nuevos_nombres, inplace=True)
            self.contable_entradasRepeditos.rename(columns=nuevos_nombres, inplace=True)
            self.banco_entradasRepeditos.rename(columns=nuevos_nombres, inplace=True)
        if elementosRepetidosSalidasContable.shape[0] != elementosRepetidosSalidasBancos.shape[0] :
            # Realizar una fusión (merge) basada en la columna 'B'
            merged_contable = pd.merge(self.contable_salidas , elementosRepetidosSalidasContable, on='id', how='left', indicator=True).copy()
            merged_bancos = pd.merge(self.banco_salidas, elementosRepetidosSalidasBancos, on='id', how='left', indicator=True).copy()
            # Seleccionar solo las filas de df1 donde no hay coincidencias en 'B' con df2
            self.contable_salidas = merged_contable[merged_contable['_merge'] == 'left_only'].drop(columns=['_merge','fecha_y','descripcion_y','valor_y','tipo_y','dia_y','mes_y','estado']).copy()
            self.banco_salidas = merged_bancos[merged_bancos['_merge'] == 'left_only'].drop(columns=['_merge','fecha_y','descripcion_y','valor_y','tipo_y','dia_y','mes_y','estado']).copy()
            #setear los repetidos
            self.contable_salidasRepeditos = elementosRepetidosSalidasContable
            self.banco_salidasRepeditos = elementosRepetidosSalidasBancos

            # Cambiar los nombres de las columnas 'A', 'B' y 'C'
            nuevos_nombres = {'valor_x': 'valor', 'tipo_x': 'tipo', 'descripcion_x': 'descripcion', 'fecha_x': 'fecha',
                              'dia_x': 'dia', 'mes_x': 'mes'}
            self.contable_salidas.rename(columns=nuevos_nombres, inplace=True)
            self.banco_salidas.rename(columns=nuevos_nombres, inplace=True)
            self.contable_salidasRepeditos.rename(columns=nuevos_nombres, inplace=True)
            self.banco_salidasRepeditos.rename(columns=nuevos_nombres, inplace=True)
        """

    def conciliacion_bancaria(self):
        # Realizar la conciliación
        conciliacion_entradas = pd.merge(self.banco_entradas, self.contable_entradas, on='id', how='outer')
        conciliacion_salidas = pd.merge(self.banco_salidas, self.contable_salidas, on='id', how='outer')
        #print(conciliacion_entradas)
        # Filtrar el DataFrame para incluir solo las filas que hicieron merge correctamente
        conciliacion_entradas_exitoso = conciliacion_entradas.dropna(subset=['tipo_y', 'tipo_x'])
        conciliacion_salidas_exitoso = conciliacion_salidas.dropna(subset=['tipo_y', 'tipo_x'])


        self.elementosConciliados = {
            'conciliado_entradas': conciliacion_entradas_exitoso,
            'conciliado_salidas': conciliacion_salidas_exitoso,
            'duplicado_entradas': self.duplicados_entradas,
            'duplicado_salidas': self.duplicados_salidas,
        }


        """
        # Identificar discrepancias
        partidaConciliatoriaEntradas = conciliacion_entradas[
            conciliacion_entradas['id'] != conciliacion_entradas['id']]
        partidaConciliatoriaSalidas = conciliacion_salidas[
            conciliacion_salidas['id'] != conciliacion_salidas['id']]
            """
        # Identificar discrepanciasss
        partidaConciliatoriaEntradas = conciliacion_entradas.drop(conciliacion_entradas_exitoso.index)
        partidaConciliatoriaSalidas = conciliacion_salidas.drop(conciliacion_salidas_exitoso.index)
        #partidaConciliatoriaEntradas = pd.concat(
        #    [conciliacion_entradas_exitoso, conciliacion_entradas]).drop_duplicates(keep=False)
        #partidaConciliatoriaSalidas = pd.concat([conciliacion_salidas_exitoso, conciliacion_salidas]).drop_duplicates(
        #    keep=False)


        # Filtro de las partidas pendientes bancos
        partidaConciliatoriaEntradas_bancarias = partidaConciliatoriaEntradas.dropna(subset=['valor_x'])
        partidaConciliatoriaSalidas_bancarias = partidaConciliatoriaSalidas.dropna(subset=['valor_x'])

        partidaConciliatoriaEntradas_contables = partidaConciliatoriaEntradas.dropna(subset=['valor_y'])
        partidaConciliatoriaSalidas_contables = partidaConciliatoriaSalidas.dropna(subset=['valor_y'])

        self.partidasConciliatorias = {
            'partidaConciliatoriaEntradas_bancarias': partidaConciliatoriaEntradas_bancarias,
            'partidaConciliatoriaSalidas_bancarias': partidaConciliatoriaSalidas_bancarias,
            'partidaConciliatoriaEntradas_contables': partidaConciliatoriaEntradas_contables,
            'partidaConciliatoriaSalidas_contables': partidaConciliatoriaSalidas_contables
        }

    def generar_conciliacion(self):

        a = InformeConciliacionBancaria(
            nombre_archivo_generado=self.nombreArchivoGenerado,
            montosBancarios_entradas=self.partidasConciliatorias.get('partidaConciliatoriaEntradas_bancarias'),
            montosBancarios_salidas=self.partidasConciliatorias.get('partidaConciliatoriaSalidas_bancarias'),
            montosContables_entradas=self.partidasConciliatorias.get('partidaConciliatoriaEntradas_contables'),
            montosContables_salidas=self.partidasConciliatorias.get('partidaConciliatoriaSalidas_contables'),
            informacionEmpresa=self.informacionEmpresa,
            informacionConciliada=self.elementosConciliados,
            duplicados_entradas_resumen = self.duplicados_entradas_resumen,
            duplicados_salidas_resumen= self.duplicados_salidas_resumen
        )
        #a.cabeceras()
        a.hoja_partidas_conciliatorias()
        a.hoja_elementos_conciliados()
        a.hoja_elementos_duplicados()
        a.guardar()





obj = ConciliacionBancaria(
    pathArchivoExtracto='extracto_bancario',
    pathArchivoContable='registros_contables',
    nombreArchivoGenerado='conciliacionBancaria',
)
obj.definir_id()
obj.modular_duplicados()
obj.conciliacion_bancaria()
obj.generar_conciliacion()


"""

# Crear un escritor de Excel
with pd.ExcelWriter('archivoAuxiliar.xlsx') as writer:
    # Guardar las transacciones conciliadas en una hoja de Excel
    conciliacion_entradas_exitoso.to_excel(writer, sheet_name='concil_ent_exit', index=False)
    conciliacion_salidas_exitoso.to_excel(writer, sheet_name='concil_sal_exit', index=False)
    partidaConciliatoriaEntradas_bancarias.to_excel(writer, sheet_name='partidaConcilEnt_banc', index=False)
    partidaConciliatoriaSalidas_bancarias.to_excel(writer, sheet_name='partidaConcilSal_banc', index=False)
    partidaConciliatoriaEntradas_contables.to_excel(writer, sheet_name='partidaConcilEnt_cont', index=False)
    partidaConciliatoriaSalidas_contables.to_excel(writer, sheet_name='partidaConcilSal_cont', index=False)

"""

