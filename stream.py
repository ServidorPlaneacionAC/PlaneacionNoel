import streamlit as st
import pandas as pd
import io
import xlsxwriter

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Simulador de Planta - Escenarios", layout="wide")

st.title("🏭 Simulador programador de planta Noel")
st.markdown("""
Esta aplicación permite cargar el archivo maestro de planta, simular tres escenarios operativos 
(Paro Programado, Capacidad Full y Demand Driven) y comparar los indicadores financieros.
""")

st.info("**Paso 1:** Descargar la plantilla base y llenar con los datos necesarios según las inidicaciones.")

# ==========================================
# --- 1. GENERACIÓN Y DESCARGA DE PLANTILLA (CON COMENTARIOS) ---
# ==========================================

def generar_plantilla():
    buffer = io.BytesIO()
    
    # 1. Definimos la estructura (Columnas) de cada hoja
    estructuras = {
        'Produccion': ['Material', 'Semana', 'Demanda semanal'],
        'Capacidad': ['Material', 'Unidades por hora', 'Unidades por pallet', 
                      'Inventario inicial', 'Valor inventario inicial', 
                      'Costo variable unitario', 'Inventario promedio'],
        'Disponibilidad': ['Semana', 'Paro programado', 'No paro'],
        'Parametros': ['Parametro', 'Valor']
    }
    
    # 2. Definimos las instrucciones (Comentarios) para cada columna
    instrucciones = {
        'Produccion': {
            'Material': 'Código material (Ej: 1001568).',
            'Semana': 'Formato AñoSemana (Ej: 202545). Solo caracteres numéricos.',
            'Demanda semanal': 'Número de unidades demandadas de la semana'
        },
        'Capacidad': {
            'Material': 'Código único del material (Ej: 1001568). Solo escribir el material una vez. Registro único por matrial',
            'Unidades por hora': 'Capacidad en unidades de la línea a producir en una hora',
            'Unidades por pallet': 'Cantidad de unidades que caben en un pallet.',
            'Inventario inicial': 'Cantidad de inventario actual en UMB',
            'Valor inventario inicial': 'Valor ($) del inventario actual. Las unidades de las monedas deben ser las mismas, solo trabajar con una unidad de valor (COP, USD, etc.)',
            'Costo variable unitario': 'Costo directo de producir una unidad ($).',
            'Inventario promedio': 'Política de inventario en unidades'
        },
        'Disponibilidad': {
            'Semana': 'Formato AñoSemana (Ej: 202545). Solo caracteres numéricos. Registro único por semana. Escribir todas las semanas a proyectar del horizonte',
            'Paro programado': 'Turnos disponibles reales considerando los paros programados (Ej: 15.3).',
            'No paro': 'Turnos disponibles teóricos con capacidad full (Ej: 21).'
        },
        'Parametros': {
            'Parametro': 'Nombres fijos: "Horas por turno", "Costo fijo", "Costo Capital".',
            'Valor': 'Valor numérico correspondiente (Ej: 8, 750000000, 0.0029).'
        }
    }

    # 3. Escribimos el archivo con XlsxWriter
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Formato para el encabezado (Negrita y borde)
        header_fmt = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })

        for nombre_hoja, columnas in estructuras.items():
            # Crear DataFrame vacío solo con encabezados
            df_vacio = pd.DataFrame(columns=columnas)
            df_vacio.to_excel(writer, sheet_name=nombre_hoja, index=False)
            
            worksheet = writer.sheets[nombre_hoja]
            
            # Escribir comentarios en cada columna
            for idx, col_name in enumerate(columnas):
                # Escribimos el encabezado con formato bonito
                worksheet.write(0, idx, col_name, header_fmt)
                
                # Buscamos el comentario correspondiente
                comentario = instrucciones.get(nombre_hoja, {}).get(col_name, "Diligenciar este dato.")
                
                # Insertamos el comentario (Burbuja)
                worksheet.write_comment(0, idx, comentario, {'x_scale': 2, 'y_scale': 1.5})
                
                # Ajustamos ancho de columna un poco
                worksheet.set_column(idx, idx, 20)

    buffer.seek(0)
    return buffer

# Botón de descarga en la interfaz
col_descarga, col_vacia = st.columns([1, 4])
with col_descarga:
    st.download_button(
        label="📥 Descargar Plantilla Excel",
        data=generar_plantilla(),
        file_name="Plantilla_Input_Planta.xlsx",
        mime="application/vnd.ms-excel",
        help="Haz clic para bajar un archivo vacío con las columnas correctas."
    )

st.divider() # Línea visual separadora

# ==========================================
# --- 2. CARGA DE ARCHIVO (Lógica original) ---
# ==========================================

st.info("**Paso 2:** Suba el archivo base diligenciado con su información. *Formato .XLSX*")

uploaded_file = st.file_uploader("📂 Carga tu archivo diligenciado aquí", type=["xlsx"])

if uploaded_file is not None:
    st.success("Archivo cargado correctamente. Procesando datos...")

    # Cache para no recargar el Excel cada vez que tocas un botón
    @st.cache_data
    def load_data(file):
        xls = pd.ExcelFile(file)
        return {
            "prod": pd.read_excel(xls, 'Produccion'),
            "cap": pd.read_excel(xls, 'Capacidad'),
            "disp": pd.read_excel(xls, 'Disponibilidad'),
            "params": pd.read_excel(xls, 'Parametros')
        }

    try:
        data = load_data(uploaded_file)
        df_prod = data["prod"]
        df_cap = data["cap"]
        df_disp_raw = data["disp"]
        df_params = data["params"]

        # Procesamiento de Parámetros
        df_params['Parametro'] = df_params['Parametro'].astype(str).str.strip()
        
        def obtener_param(nombre, default=None):
            fila = df_params[df_params['Parametro'] == nombre]
            if fila.empty: return default
            valor = fila['Valor'].values[0]
            if isinstance(valor, str): valor = valor.replace(',', '.')
            try: return float(valor)
            except: return valor

        horas_por_turno = obtener_param('Horas por turno', default=8)
        costo_fijo_param = obtener_param('Costo fijo', default=742373394)
        tasa_costo_capital = obtener_param('Costo Capital', default=0.0029)

        # --- PRE-PROCESAMIENTO ---
        df_prod['Material'] = df_prod['Material'].astype(str).str.strip()
        df_cap['Material'] = df_cap['Material'].astype(str).str.strip()

        cols_capacidad = ['Material', 'Unidades por hora']
        nuevas_cols = ['Unidades por pallet', 'Inventario inicial', 'Valor inventario inicial', 
                       'Costo variable unitario', 'Inventario promedio'] 

        for col in nuevas_cols:
            if col in df_cap.columns: cols_capacidad.append(col)
            else: df_cap[col] = 0; cols_capacidad.append(col)

        df_base_materiales = df_prod.merge(df_cap[cols_capacidad], on='Material', how='left')
        
        # Constantes
        suma_inv_promedio_objetivo = df_base_materiales.groupby('Material')['Inventario promedio'].first().sum()
        suma_valor_inicial_maestro = df_base_materiales.groupby('Material')['Valor inventario inicial'].first().sum()

        df_base_materiales['Semana'] = df_base_materiales['Semana'].astype(str).str.strip()
        df_base_materiales['Semana_Num'] = pd.to_numeric(df_base_materiales['Semana'], errors='coerce')
        df_disp_raw['Semana'] = df_disp_raw['Semana'].astype(str).str.strip()
        df_base_materiales['Horas_Necesarias'] = df_base_materiales['Demanda semanal'] / df_base_materiales['Unidades por hora']

        # --- FUNCIÓN NÚCLEO ---
        def calcular_escenario(nombre_escenario, df_disp_scenario, modo_demanda=False):
            df_main = df_base_materiales.copy()
            resumen_semanal = df_main.groupby('Semana')['Horas_Necesarias'].sum().reset_index()
            resumen_semanal.rename(columns={'Horas_Necesarias': 'Total_Horas_Req_Semana'}, inplace=True)
            
            if modo_demanda:
                resumen_semanal['Turnos disponibles'] = resumen_semanal['Total_Horas_Req_Semana'] / horas_por_turno
            else:
                resumen_semanal = resumen_semanal.merge(df_disp_scenario[['Semana', 'Turnos disponibles']], on='Semana', how='left')
                resumen_semanal['Turnos disponibles'] = resumen_semanal['Turnos disponibles'].fillna(0)
            
            resumen_semanal['Horas_Capacidad_Total'] = resumen_semanal['Turnos disponibles'] * horas_por_turno
            resumen_semanal['Diferencia_Horas'] = resumen_semanal['Horas_Capacidad_Total'] - resumen_semanal['Total_Horas_Req_Semana']
            
            df_final = df_main.merge(resumen_semanal, on='Semana', how='left')
            
            df_final['Participacion'] = df_final['Horas_Necesarias'] / df_final['Total_Horas_Req_Semana']
            df_final['Participacion'] = df_final['Participacion'].fillna(0)
            df_final['Horas_Ajuste'] = df_final['Participacion'] * df_final['Diferencia_Horas']
            
            df_final['Horas_Finales_Asignadas'] = df_final['Horas_Necesarias'] + df_final['Horas_Ajuste']
            df_final['Horas_Finales_Asignadas'] = df_final['Horas_Finales_Asignadas'].apply(lambda x: max(x, 0))
            
            df_final['Unidades_A_Producir'] = df_final['Horas_Finales_Asignadas'] * df_final['Unidades por hora']
            df_final['Unidades_Ajuste'] = df_final['Horas_Ajuste'] * df_final['Unidades por hora']
            df_final['Costo variable total'] = df_final['Unidades_A_Producir'] * df_final['Costo variable unitario']
            
            df_final['Unidades por pallet'] = df_final['Unidades por pallet'].replace(0, 1)
            df_final['Pallets_Requeridos'] = df_final['Unidades_A_Producir'] / df_final['Unidades por pallet']
            df_final['Pallets_Extra'] = df_final['Unidades_Ajuste'] / df_final['Unidades por pallet']
            
            tabla_costos = df_final.groupby('Semana').agg({'Unidades_A_Producir': 'sum'}).reset_index()
            tabla_costos['Costo fijo'] = costo_fijo_param
            tabla_costos['Costo fijo unitario'] = tabla_costos.apply(lambda x: x['Costo fijo']/x['Unidades_A_Producir'] if x['Unidades_A_Producir']>0 else 0, axis=1)
            
            df_final = df_final.merge(tabla_costos[['Semana', 'Costo fijo unitario']], on='Semana', how='left')
            df_final['Costo unitario total'] = df_final['Costo variable unitario'] + df_final['Costo fijo unitario']
            
            df_final = df_final.sort_values(by=['Material', 'Semana_Num']).reset_index(drop=True)
            
            cols_calc = ['Inv_Inicial_Und', 'Inv_Inicial_Valor', 'Costo_Promedio_Ponderado', 
                         'Inv_Final_Und', 'Inv_Final_Valor', 'Costo_Mercancia_Vendida']
            for c in cols_calc: df_final[c] = 0.0

            prev_material = None
            prev_final_und = 0.0
            prev_final_val = 0.0

            for i in range(len(df_final)):
                row = df_final.iloc[i]
                curr_material = row['Material']
                
                if curr_material != prev_material:
                    qty_ini = row['Inventario inicial']
                    val_ini = row['Valor inventario inicial']
                else:
                    qty_ini = prev_final_und
                    val_ini = prev_final_val
                
                qty_prod = row['Unidades_A_Producir']
                cost_prod = row['Costo unitario total']
                val_prod = qty_prod * cost_prod
                
                total_qty = qty_ini + qty_prod
                total_val = val_ini + val_prod
                
                wac = total_val / total_qty if total_qty > 0 else cost_prod
                
                qty_sold = min(row['Demanda semanal'], total_qty)
                cogs_total = qty_sold * wac
                
                qty_final = total_qty - qty_sold
                val_final = qty_final * wac
                
                df_final.at[i, 'Inv_Inicial_Und'] = qty_ini
                df_final.at[i, 'Inv_Inicial_Valor'] = val_ini
                df_final.at[i, 'Costo_Promedio_Ponderado'] = wac
                df_final.at[i, 'Inv_Final_Und'] = qty_final
                df_final.at[i, 'Inv_Final_Valor'] = val_final
                df_final.at[i, 'Costo_Mercancia_Vendida'] = cogs_total
                
                prev_material = curr_material
                prev_final_und = qty_final
                prev_final_val = val_final

            resumen_agrupado = df_final.groupby('Semana').agg({
                'Unidades_A_Producir': 'sum', 'Pallets_Requeridos': 'sum', 'Pallets_Extra': 'sum',
                'Turnos disponibles': 'first', 'Total_Horas_Req_Semana': 'first',
                'Diferencia_Horas': 'first', 'Semana_Num': 'first',
                'Costo variable total': 'sum', 'Costo_Mercancia_Vendida': 'sum',
                'Inv_Final_Valor': 'sum', 'Inv_Final_Und': 'sum'
            }).reset_index()
            
            resumen_agrupado = resumen_agrupado.sort_values('Semana_Num').reset_index(drop=True)
            
            resumen_agrupado['Costo fijo total($)'] = costo_fijo_param
            resumen_agrupado['Costo fijo unitario ($/Und)'] = resumen_agrupado.apply(lambda x: x['Costo fijo total($)']/x['Unidades_A_Producir'] if x['Unidades_A_Producir']>0 else 0, axis=1)
            resumen_agrupado['Turnos necesarios'] = resumen_agrupado['Total_Horas_Req_Semana'] / horas_por_turno
            resumen_agrupado['Costo total producción ($)'] = resumen_agrupado['Costo fijo total($)'] + resumen_agrupado['Costo variable total']
            resumen_agrupado['Costo Capital ($)'] = (resumen_agrupado['Inv_Final_Valor'] * tasa_costo_capital).clip(lower=0)
            
            resumen_agrupado['Inv_Previo_Calculo'] = resumen_agrupado['Inv_Final_Valor'].shift(1).fillna(suma_valor_inicial_maestro)
            resumen_agrupado['Variación inventario'] = resumen_agrupado['Inv_Final_Valor'] - resumen_agrupado['Inv_Previo_Calculo']
            
            resumen_agrupado['EBITDA (CMV)'] = resumen_agrupado['Costo_Mercancia_Vendida']
            resumen_agrupado['Flujo de caja'] = resumen_agrupado['EBITDA (CMV)'] - resumen_agrupado['Variación inventario']
            
            # Nuevas columnas valorización promedio
            resumen_agrupado['Costo Unitario Promedio Inventario ($/Und)'] = resumen_agrupado.apply(
                lambda x: x['Inv_Final_Valor'] / x['Inv_Final_Und'] if x['Inv_Final_Und'] > 0 else 0, axis=1
            )
            resumen_agrupado['Valor Inventario (Inv. Promedio Input) ($)'] = resumen_agrupado['Costo Unitario Promedio Inventario ($/Und)'] * suma_inv_promedio_objetivo

            df_final['Escenario'] = nombre_escenario
            resumen_agrupado['Escenario'] = nombre_escenario
            
            return df_final, resumen_agrupado

        # --- EJECUCIÓN ---
        with st.spinner('Ejecutando simulaciones...'):
            lista_materiales = []
            lista_semanal = []

            # 1. Paro
            if 'Paro programado' in df_disp_raw.columns:
                disp_s1 = df_disp_raw[['Semana', 'Paro programado']].rename(columns={'Paro programado': 'Turnos disponibles'})
                mat_s1, sem_s1 = calcular_escenario("Paro Programado", disp_s1)
                lista_materiales.append(mat_s1); lista_semanal.append(sem_s1)
            
            # 2. No Paro
            if 'No paro' in df_disp_raw.columns:
                disp_s2 = df_disp_raw[['Semana', 'No paro']].rename(columns={'No paro': 'Turnos disponibles'})
                mat_s2, sem_s2 = calcular_escenario("No Paro (Full)", disp_s2)
                lista_materiales.append(mat_s2); lista_semanal.append(sem_s2)
            
            # 3. Demanda
            mat_s3, sem_s3 = calcular_escenario("Siguiendo Demanda", None, modo_demanda=True)
            lista_materiales.append(mat_s3); lista_semanal.append(sem_s3)

            df_master_materiales = pd.concat(lista_materiales, ignore_index=True)
            df_master_semanal = pd.concat(lista_semanal, ignore_index=True)

        # --- TABLAS FINALES ---
        tabla_comparativa = df_master_semanal.groupby('Escenario')[['Costo_Mercancia_Vendida', 'EBITDA (CMV)', 'Costo Capital ($)']].sum().reset_index()
        tabla_comparativa = tabla_comparativa.rename(columns={
            'Costo_Mercancia_Vendida': 'CMV Total ($)',
            'EBITDA (CMV)': 'EBITDA Total ($)', 
            'Costo Capital ($)': 'Costo Capital Total ($)'
        })

        deltas = []
        for escenario in df_master_semanal['Escenario'].unique():
            df_esc = df_master_semanal[df_master_semanal['Escenario'] == escenario].sort_values('Semana_Num')
            if not df_esc.empty:
                val_est_s1 = df_esc.iloc[0]['Valor Inventario (Inv. Promedio Input) ($)']
                val_est_sf = df_esc.iloc[-1]['Valor Inventario (Inv. Promedio Input) ($)']
                delta = val_est_sf - val_est_s1
            else:
                delta = 0
            deltas.append({'Escenario': escenario, 'Delta Valor Inv': delta})
        
        df_deltas = pd.DataFrame(deltas)
        tabla_comparativa = tabla_comparativa.merge(df_deltas, on='Escenario')

        tabla_comparativa['Impacto Total FCL'] = (
            tabla_comparativa['CMV Total ($)'] + 
            tabla_comparativa['Costo Capital Total ($)'] + 
            tabla_comparativa['Delta Valor Inv']
        )

        # --- VISUALIZACIÓN ---
        st.subheader("📊 Comparativa de Escenarios")
        
        col1, col2, col3 = st.columns(3)
        mejores = tabla_comparativa.sort_values('Impacto Total FCL')
        
        with col1:
            st.metric("Mejor Escenario", mejores.iloc[0]['Escenario'])
        with col2:
            st.metric("Impacto FCL (Menor es mejor)", f"${mejores.iloc[0]['Impacto Total FCL']:,.0f}")
        with col3:
            ahorro = mejores.iloc[-1]['Impacto Total FCL'] - mejores.iloc[0]['Impacto Total FCL']
            st.metric("Diferencia vs Peor Escenario", f"${ahorro:,.0f}", delta_color="normal")

        st.bar_chart(tabla_comparativa, x="Escenario", y="Impacto Total FCL", color="Escenario")
        
        with st.expander("Ver Datos Comparativos"):
        # Definimos qué columnas queremos formatear como dinero
            columnas_dinero = [
                'CMV Total ($)', 
                'EBITDA Total ($)', 
                'Costo Capital Total ($)', 
                'Delta Valor Inv', 
                'Impacto Total FCL'
            ]
            
            # Creamos un diccionario de formato solo para esas columnas
            formato = {col: "${:,.2f}" for col in columnas_dinero}
            
            # Aplicamos el estilo de forma segura
            st.dataframe(tabla_comparativa.style.format(formato))

        with st.expander("Ver Detalle Semanal"):
            st.dataframe(df_master_semanal)

        # --- DESCARGA ---
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_master_materiales.to_excel(writer, sheet_name='Detalle_Materiales', index=False)
            df_master_semanal.to_excel(writer, sheet_name='Resumen_Semanal', index=False)
            tabla_comparativa.to_excel(writer, sheet_name='Comparativa_Escenarios', index=False)
        
        buffer.seek(0)
        st.download_button(
            label="📥 Descargar Reporte Completo (Excel)",
            data=buffer,
            file_name="Reporte_Simulacion.xlsx",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"❌ Error al procesar: {e}")

else:
    st.markdown("<p style='color: red;font-size: 14px;'>👆 Esperando archivo...</h4>", unsafe_allow_html=True)