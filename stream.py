# Import necessary libraries for the application
import streamlit as st
import pandas as pd
import pyomo.environ as pyo
import io
import altair as alt

# Initial Streamlit page configuration (browser title and page width)
st.set_page_config(page_title="Optimización de producción", layout="wide")
st.title("🏭 Simulador programación de plantas")

# Introductory description of the application
st.markdown("""
Esta aplicación simular tres escenarios operativos (Paro Programado, Capacidad Full y Demand Driven) y comparar 
los indicadores financieros.""")

st.info("**Paso 1:** Descargar la plantilla base y llenar con los datos necesarios según las inidicaciones.")

# ==========================================
# --- 1. GENERATION AND DOWNLOAD OF TEMPLATE ---
# ==========================================

def generar_plantilla():
    """
    Creates an in-memory Excel file (buffer) with the base structure, 
    formatted headers, and comments containing instructions for the user.
    """
    buffer = io.BytesIO()
    
    # Definition of expected columns for each Excel tab
    estructuras = {
        'Produccion': ['Material', 'Semana', 'Demanda semanal'],
        'Capacidad': ['Material', 'Unidades por hora', 'Unidades por pallet', 
                      'Inventario inicial', 'Valor inventario inicial', 
                      'Costo variable unitario', 'Inventario promedio'],
        'Disponibilidad': ['Semana', 'Turnos disponibles'],
        'Parametros': ['Parametro', 'Valor']
    }
    
    # Help texts that will appear as comments in the header cells
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
            'Turnos disponibles': 'Turnos disponibles teóricos con capacidad full (Ej: 21).'
        },
        'Parametros': {
            'Parametro': 'Valores fijos y únicos que no dependen de los periodos de producción ni los materiales. (Ej: "Horas por turno", "Costo fijo", "Costo Capital)".',
            'Valor': 'Valor numérico correspondiente (Ej: 8, 750000000, 0.0029).'
        }
    }

    # Creation of the Excel document using xlsxwriter as engine
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Visual style for the header row (bold, text wrap, top vertical align, background color, borders)
        header_fmt = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })

        # Iterate over the structures to create each Excel sheet
        for nombre_hoja, columnas in estructuras.items():
            df_vacio = pd.DataFrame(columns=columnas)
            df_vacio.to_excel(writer, sheet_name=nombre_hoja, index=False)
            
            worksheet = writer.sheets[nombre_hoja]
            
            # Apply header format, column width, and add comments
            for idx, col_name in enumerate(columnas):
                worksheet.write(0, idx, col_name, header_fmt)
                comentario = instrucciones.get(nombre_hoja, {}).get(col_name, "Diligenciar este dato.")
                worksheet.write_comment(0, idx, comentario, {'x_scale': 2, 'y_scale': 1.5})
                worksheet.set_column(idx, idx, 20)

    buffer.seek(0)
    return buffer

# UI: Button to download the empty template
col_descarga, col_vacia = st.columns([1, 4])
with col_descarga:
    st.download_button(
        label="📥 Descargar Plantilla Excel",
        data=generar_plantilla(),
        file_name="Plantilla_Input_Planta.xlsx",
        mime="application/vnd.ms-excel",
        help="Haz clic para bajar un archivo vacío con las columnas correctas."
    )

st.divider()

# ==========================================
# 2. DATA LOADING
# ==========================================

# UI: Step 2 Instructions
st.markdown(
    """
    <div style="background-color: #e0f2fe; padding: 16px; border-radius: 8px; color: #0369a1 ; margin-bottom: 12px">
        <span style="font-size: 16px; font-weight: bold;">Paso 2:</span> 
        <span style="font-size: 16px;">👇 Agregar archivo input en formato .xlsx</span>
    </div>
    """,
    unsafe_allow_html=True
)

# Widget to upload the filled Excel file
uploaded_file = st.file_uploader(
    "Label oculto", 
    type=['xlsx'], 
    label_visibility="collapsed"
)

# Execute the processing flow only if a file has been uploaded
if uploaded_file is not None:
    # Read Excel tabs into Pandas DataFrames
    xls = pd.ExcelFile(uploaded_file)
    df_prod = pd.read_excel(xls, 'Produccion')
    df_disp = pd.read_excel(xls, 'Disponibilidad')
    df_cap  = pd.read_excel(xls, 'Capacidad')
    df_par  = pd.read_excel(xls, 'Parametros')

    # Initial cleaning and strict data type conversion for main keys
    df_prod['Semana'] = df_prod['Semana'].fillna(0).astype(int)
    df_disp['Semana'] = df_disp['Semana'].fillna(0).astype(int)
    df_prod['Material'] = df_prod['Material'].astype(str)
    df_cap['Material']  = df_cap['Material'].astype(str)

    # Force float on numeric columns so format="localized" 
    # applies the thousands separator consistently in the UI.
    df_par['Valor'] = pd.to_numeric(df_par['Valor'], errors='coerce').astype(float)

    for col in ['Unidades por hora', 'Unidades por pallet', 'Inventario inicial',
                'Valor inventario inicial', 'Costo variable unitario', 'Inventario promedio']:
        df_cap[col] = pd.to_numeric(df_cap[col], errors='coerce').astype(float)

    df_prod['Demanda semanal']    = pd.to_numeric(df_prod['Demanda semanal'],    errors='coerce').astype(float)
    df_disp['Turnos disponibles'] = pd.to_numeric(df_disp['Turnos disponibles'], errors='coerce').astype(float)

    st.subheader("📝 Vista previa y edición de datos")
    st.info("Modificar, agregar o eliminar los valores directamente de las tablas si es necesario.")

    # Creation of tabs to view/edit each DataFrame in the application
    tab1, tab2, tab3, tab4 = st.tabs(["Producción", "Capacidad", "Disponibilidad", "Parámetros"])

    # ── Tab 1: Production ──────────────────────────────────────────────────────
    with tab1:
        # data_editor allows modifying data directly; overwrites the original dataframe
        df_prod = st.data_editor(
            df_prod,
            num_rows="dynamic",
            use_container_width=True,
            key="edit_prod",
            column_config={
                "Semana":          st.column_config.NumberColumn(format="%d"),
                "Demanda semanal": st.column_config.NumberColumn(format="localized"),
            }
        )

    # ── Tab 2: Capacity ───────────────────────────────────────────────────────
    with tab2:
        df_cap = st.data_editor(
            df_cap, 
            num_rows="dynamic", 
            use_container_width=True, 
            key="edit_cap",
            column_config={
                "Unidades por hora":        st.column_config.NumberColumn(format="localized"),
                "Unidades por pallet":      st.column_config.NumberColumn(format="localized"),
                "Inventario inicial":       st.column_config.NumberColumn(format="localized"),
                "Valor inventario inicial": st.column_config.NumberColumn(format="localized"),
                "Costo variable unitario":  st.column_config.NumberColumn(format="localized"),
                "Inventario promedio":      st.column_config.NumberColumn(format="localized"),
            }
        )

    # ── Tab 3: Availability ──────────────────────────────────────────────────
    with tab3:
        df_disp = st.data_editor(
            df_disp,
            num_rows="dynamic",
            use_container_width=True,
            key="edit_disp",
            column_config={
                "Semana":             st.column_config.NumberColumn(format="%d"),
                "Turnos disponibles": st.column_config.NumberColumn(format="localized"),
            }
        )

    # ── Tab 4: Parameters ──────────────────────────────────────────────────────
    with tab4:
        df_par = st.data_editor(
            df_par, 
            num_rows="dynamic", 
            use_container_width=True, 
            key="edit_par",
            column_config={
                "Valor": st.column_config.NumberColumn(format="localized"),
            }
        )

    # Re-apply data types in case the user inserted anomalous values in the editor
    df_prod['Semana'] = df_prod['Semana'].fillna(0).astype(int)
    df_disp['Semana'] = df_disp['Semana'].fillna(0).astype(int)
    df_prod['Material'] = df_prod['Material'].astype(str)
    df_cap['Material']  = df_cap['Material'].astype(str)

    # Helper function to extract specific parameters from the parameters table
    def get_param(name, default):
        try:
            val = df_par.loc[df_par['Parametro'] == name, 'Valor'].iloc[0]
            if isinstance(val, str): val = float(val.replace(',', '.'))
            return float(val)
        except: return default

    # Extraction of global parameters
    h = get_param("Horas por turno", 8.0)
    C_fijo = get_param("Costo fijo", 742373394)
    r = get_param("Costo Capital", 0.0029)
    cap_cedi = 5000
    c_pallet  = 15000

    # Construction of lists and dictionaries required by the mathematical model
    M_set = sorted(df_cap['Material'].unique().tolist()) # Set of Materials
    T_set = sorted(df_disp['Semana'].unique().tolist())  # Set of Weeks (Periods)

    UPH    = df_cap.set_index('Material')['Unidades por hora'].to_dict()
    UPP    = df_cap.set_index('Material')['Unidades por pallet'].to_dict()
    CV     = df_cap.set_index('Material')['Costo variable unitario'].to_dict()
    I0     = df_cap.set_index('Material')['Inventario inicial'].to_dict()
    Pol    = df_cap.set_index('Material')['Inventario promedio'].fillna(0).to_dict()
    Val_I0 = df_cap.set_index('Material')['Valor inventario inicial'].to_dict()

    # Creation of the demand dictionary by material and week
    Dem = {(m, t): 0 for m in M_set for t in T_set}
    for index, row in df_prod.iterrows():
        mat, sem, cant = str(row['Material']), int(row['Semana']), row['Demanda semanal']
        if sem in T_set and mat in M_set: Dem[(mat, sem)] = cant

    # Extraction of available base shifts per week
    base_shifts = {}
    for index, row in df_disp.iterrows():
        t = int(row['Semana'])
        if t in T_set:
            base_shifts[t] = int(row['Turnos disponibles'])

    # ==========================================
    # SCHEDULED DOWNTIME CONFIGURATION
    # ==========================================
    st.divider()
    st.subheader("🛑 Configuración de Paros Programados")
    st.write("Digita la cantidad de turnos a parar en las semanas correspondientes. El modelo restará estos turnos de la capacidad total.")

    # Create a base DataFrame for the interactive downtime UI
    df_paros_base = pd.DataFrame({
        "Semana": list(base_shifts.keys()),
        "Turnos disponibles": list(base_shifts.values()),
        "Turnos a parar": [0] * len(base_shifts)
    })

    # Show editor to collect the input of shifts to stop
    df_paros_edit = st.data_editor(
        df_paros_base,
        use_container_width=True,
        hide_index=True,
        disabled=["Semana", "Turnos disponibles"], # Prevents altering base data here
        column_config={
            "Turnos a parar": st.column_config.NumberColumn(
                "Turnos a parar",
                min_value=0, # Validation to prevent negative numbers
                step=1
            )
        }
    )

    errores_paro = False
    paro_shifts = {}

    # Validation and instantiation of the new shift scheme adjusted by downtime
    for index, row in df_paros_edit.iterrows():
        semana = int(row['Semana'])
        disp = int(row['Turnos disponibles'])
        parar = int(row['Turnos a parar'])
        
        # Block execution if more shifts are subtracted than exist
        if parar > disp:
            st.error(f"⚠️ Error en la semana {semana}: El número de turnos a parar ({parar}) no puede ser mayor a los turnos disponibles ({disp}).")
            errores_paro = True
        
        paro_shifts[semana] = disp - parar

    if errores_paro:
        st.stop() # Stop the app if there is a logical problem detected above

    # ==========================================
    # SCENARIO DEFINITION
    # ==========================================
    # Dictionary mapping the 3 scenarios with their particular shift and slack rules
    scenarios = {
        "Demand Driven":   {"shifts": base_shifts, "force_max": False, "fill_cap": False},
        "Paro Programado": {"shifts": paro_shifts, "force_max": True,  "fill_cap": True}, 
        "Full Capacity":   {"shifts": base_shifts, "force_max": True,  "fill_cap": True},
        "Paro Óptimo":     {"shifts": base_shifts, "force_max": False, "fill_cap": True} # <-- NUEVO ESCENARIO
    }

    # Period map to establish inventory balance with respect to the previous week
    sorted_weeks = sorted(T_set)
    prev_week_map = {wk: sorted_weeks[i-1] for i, wk in enumerate(sorted_weeks) if i > 0}

# ==========================================
# 3. OPTIMIZATION MODEL
# ==========================================
    def generate_scenario_report(name, max_shifts, force_max, fill_cap):
        """
        Defines and solves the mathematical model using Pyomo for a given scenario.
        Calculates financial indicators and returns summary and detail tables.
        """
        model = pyo.ConcreteModel(name=name)
        model.M = pyo.Set(initialize=M_set) # Set of Materials
        model.T = pyo.Set(initialize=sorted_weeks, ordered=True) # Set of Time

        # Decision variables
        model.X = pyo.Var(model.M, model.T, domain=pyo.NonNegativeIntegers) # Production
        model.Y = pyo.Var(model.T, domain=pyo.NonNegativeIntegers)          # Shifts to operate
        model.I = pyo.Var(model.M, model.T, domain=pyo.NonNegativeIntegers) # Final inventory
        model.P = pyo.Var(model.M, model.T, domain=pyo.NonNegativeIntegers) # Pallets 
        model.E = pyo.Var(model.T, domain=pyo.NonNegativeIntegers)          # External storage

        # Objective Function: Minimize Variable Production Cost + Inventory Capital Cost + External Warehouse Cost
        def obj_rule(mdl):
            return sum(mdl.X[m, t] * CV[m] for m in mdl.M for t in mdl.T) + \
                sum(r * mdl.I[m, t] * CV[m] for m in mdl.M for t in mdl.T) + \
                sum(c_pallet * mdl.E[t] for t in mdl.T)
        model.Obj = pyo.Objective(rule=obj_rule, sense=pyo.minimize)

        # Constraint: Limit of shifts to operate
        def shift_limit_rule(mdl, t): 
            if force_max: return mdl.Y[t] == max_shifts[t]
            else:         return mdl.Y[t] <= max_shifts[t]
        model.ShiftLimit = pyo.Constraint(model.T, rule=shift_limit_rule)

        # Constraint: Total production in hours cannot exceed the hours of operational shifts
        def capacity_rule(mdl, t):
            req = sum(mdl.X[m, t] / UPH[m] for m in mdl.M)
            return req <= mdl.Y[t] * h
        model.Capacity = pyo.Constraint(model.T, rule=capacity_rule)

        # Constraint: Force idle capacity filling
        def fill_capacity_rule(mdl, t):
            if fill_cap:
                # 1. Calculate the total continuous time (in hours) needed to produce all units
                req = sum(mdl.X[m, t] / UPH[m] for m in mdl.M) 
                
                # 2. Find the time it takes to produce ONE single unit of the slowest material
                max_unit_time = max(1 / UPH[m] for m in mdl.M)
                
                # 3. Force the required production time to match the total available shift hours.
                return req >= (mdl.Y[t] * h) - (max_unit_time + 0.001) 
            else:
                return pyo.Constraint.Skip   
        model.FillCapacity = pyo.Constraint(model.T, rule=fill_capacity_rule)

        # Constraint: Inventory balance equation
        def inv_balance_rule(mdl, m, t):
            prod = mdl.X[m, t]
            if t == sorted_weeks[0]: return mdl.I[m, t] == I0[m] + prod - Dem[(m, t)]
            else:                    return mdl.I[m, t] == mdl.I[m, prev_week_map[t]] + prod - Dem[(m, t)]
        model.InvBalance = pyo.Constraint(model.M, model.T, rule=inv_balance_rule)

        # Constraint: Minimum safety stock policy
        def inv_policy_rule(mdl, m, t): return mdl.I[m, t] >= Pol[m]
        model.InvPolicy = pyo.Constraint(model.M, model.T, rule=inv_policy_rule)

        # Constraint: Round up to calculate total required pallets per material
        def pallet_ceil_rule(mdl, m, t):
            return mdl.P[m, t] >= mdl.I[m, t] / UPP[m]
        model.PalletCeil = pyo.Constraint(model.M, model.T, rule=pallet_ceil_rule)

        # Constraint: Determination of external pallets
        def external_wh_rule(mdl, t):
            total_pallets = sum(mdl.P[m, t] for m in mdl.M)
            return mdl.E[t] >= total_pallets - cap_cedi
        model.ExternalWH = pyo.Constraint(model.T, rule=external_wh_rule)

        # Constraint: To force the strict use of the minimum possible number of shifts
        def strict_shifts_rule(mdl, t):
            if not force_max and not fill_cap: 
                req = sum(mdl.X[m, t] / UPH[m] for m in mdl.M)
                return req >= ((mdl.Y[t] - 1) * h) + 0.001
            else:
                return pyo.Constraint.Skip
        model.StrictShifts = pyo.Constraint(model.T, rule=strict_shifts_rule)

        # Instantiate the solver
        solver = pyo.SolverFactory('appsi_highs') 
        tiempo_limite_segundos = 180
        solver.options['time_limit'] = tiempo_limite_segundos 
        
        results = solver.solve(model, load_solutions=False)

        # Verification of the result's health
        if results.solver.termination_condition != pyo.TerminationCondition.optimal:
            print(f"⚠️ {name} is Infeasible or Time Limit Reached.")
            error_df = pd.DataFrame([{"Error": f"Escenario {name} es inviable o excedió el tiempo límite de {tiempo_limite_segundos}s."}])
            return error_df, error_df, None # <-- Agregamos un tercer valor nulo para los errores
        
        # Load results back into the Pyomo variable objects
        model.solutions.load_from(results)

        valor_funcion_objetivo = pyo.value(model.Obj)

        summary_data = []
        details_data = []
        
        prev_inv_value = sum(Val_I0[m] for m in M_set) 
        prev_inv_units = {m: I0[m] for m in M_set}
        
        prev_inv_unit_cost = {}
        for m in M_set:
            if I0[m] > 0:
                prev_inv_unit_cost[m] = Val_I0[m] / I0[m]
            else:
                prev_inv_unit_cost[m] = CV[m]

        for t in sorted_weeks:
            y_val = int(pyo.value(model.Y[t]))
            disp_shifts = max_shifts[t]
            prod_und_total = sum(pyo.value(model.X[m, t]) for m in M_set)
            
            time_req = sum(pyo.value(model.X[m, t]) / UPH[m] for m in M_set)
            holgura = (y_val * h) - time_req 
            
            pallets_extra_total = int(pyo.value(model.E[t]))
            total_pallets_inv = sum(pyo.value(model.I[m, t]) / UPP[m] for m in M_set)
            inventario_und_total = sum(pyo.value(model.I[m, t]) for m in M_set)
            
            cf_total = C_fijo 
            cf_unitario = cf_total / prod_und_total if prod_und_total > 0 else 0
            cv_total = sum(pyo.value(model.X[m, t]) * CV[m] for m in M_set)
            costo_total_prod = cv_total + cf_total
            costo_bodega_externa = pallets_extra_total * c_pallet
            
            val_inv_semana = 0
            costo_cap_semana = 0
            cmv_semana = 0
            
            for m in M_set:
                demand = Dem[(m, t)]
                uph = UPH[m]
                cv_unit = CV[m]
                upp = UPP[m]
                
                prod_und = pyo.value(model.X[m, t])
                prod_pallets = prod_und / upp if upp > 0 else 0
                
                horas_necesarias_demanda = demand / uph
                horas_totales_mat = prod_und / uph
                peso_tiempo = (horas_totales_mat / time_req) if time_req > 0 else 0
                
                horas_adic = holgura * peso_tiempo
                prod_adic = horas_adic * uph
                pallets_almacenar = prod_adic / upp
                
                inv_final = pyo.value(model.I[m, t])
                inv_anterior = prev_inv_units[m]
                
                costo_var_mat = prod_und * cv_unit
                costo_unitario_total = cv_unit + cf_unitario
                
                if t == sorted_weeks[0]:
                    costo_inv_anterior = Val_I0[m]
                    costo_unitario_inv_ant = prev_inv_unit_cost[m]
                else:
                    costo_unitario_inv_ant = prev_inv_unit_cost[m]
                    costo_inv_anterior = inv_anterior * costo_unitario_inv_ant

                total_units_available = inv_anterior + prod_und
                if total_units_available > 0:
                    current_inv_unit_cost = (costo_inv_anterior + (costo_unitario_total * prod_und)) / total_units_available
                else:
                    current_inv_unit_cost = costo_unitario_inv_ant
                    
                costo_inv_final = inv_final * current_inv_unit_cost
                cmv_mat = demand * current_inv_unit_cost
                
                val_inv_semana   += costo_inv_final
                costo_cap_semana += (r * costo_inv_final)
                cmv_semana       += cmv_mat
                
                details_data.append({
                    "Material": m,
                    "Periodo \"aaaass\")": t,
                    "Demanda (Und)": demand,
                    "Capacidad (Und / hr)": uph,
                    "Horas necesarias (hr)": horas_necesarias_demanda,
                    "Tiempo adicional asignado (hr)": horas_adic,
                    "Producción en tiempo adicional (Und)": prod_adic,
                    "Producción total (Und)": prod_und,
                    "Pallets a almacenar": pallets_almacenar,
                    "Inventario Final (Und)": inv_final,
                    "Inventario mes anterior (Und)": inv_anterior,
                    "costo variable unitario ($/Und)": cv_unit,
                    "costo variable total ($)": costo_var_mat,
                    "costo unitario total ($/Und)": costo_unitario_total,
                    "Costo inventario mes anterior ($)": costo_inv_anterior,
                    "Costo unitario inventario ($/Und)": current_inv_unit_cost,
                    "Costo inventario ($)": costo_inv_final,
                    "Costo Mercancía Vendida ($)": cmv_mat
                })
                
                prev_inv_units[m]     = inv_final
                prev_inv_unit_cost[m] = current_inv_unit_cost

            # Financial calculations for the global summary
            var_inv = val_inv_semana - prev_inv_value
            prev_inv_value = val_inv_semana 
            
            total_pol_und = sum(Pol[m] for m in M_set)
            costo_promedio_global = (val_inv_semana / inventario_und_total) if inventario_und_total > 0 else 0
            valor_politica_inv = total_pol_und * costo_promedio_global

            # Insertion into executive summary table
            summary_data.append({
                "semana": t,
                "Costo fijo total($)": cf_total,
                "Total Producido (Und)": prod_und_total,
                "Costo fijo unitario ($/Und)": cf_unitario,
                "Turnos disponibles": disp_shifts,
                "Turnos necesarios": y_val,
                "Turnos a apagar (Recomendación)": disp_shifts - y_val,
                "Tiempo holgura (hr)": holgura,
                "Pallets a almacenar en tiempo extra": pallets_extra_total,
                "Total pallets almacenados": total_pallets_inv,
                "Costo Capital ($)": val_inv_semana * r,
                "Costo total producción ($)": costo_total_prod,
                "Pallets externos": pallets_extra_total,
                "Costo Bodega Externa ($)": costo_bodega_externa,
                "CMV ($)": cmv_semana,
                "Valor inventario": val_inv_semana,
                "Variación inventario": var_inv,
                "EBITDA (CMV)": cmv_semana,
                "Flujo de caja": cmv_semana - var_inv,
                "Inventario": inventario_und_total,
                "Valor política inventario ($)": valor_politica_inv
            })

        return pd.DataFrame(summary_data), pd.DataFrame(details_data), valor_funcion_objetivo

# ==========================================
# 4. EXECUTION (Execution and Export)
# ==========================================

def dar_formato_excel(writer, df, sheet_name):
    """
    Utility function to inject native cell formatting in Excel 
    to a DataFrame exported using xlsxwriter.
    """
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    workbook = writer.book
    
    formato_miles = workbook.add_format({'num_format': '#,##0'})
    formato_moneda = workbook.add_format({'num_format': '"$"#,##0'})
    formato_decimal = workbook.add_format({'num_format': '#,##0.00'})
    
    for idx, col in enumerate(df.columns):
        ancho_maximo = max(df[col].astype(str).map(len).max(), len(col)) + 2
        
        if "$" in col:
            worksheet.set_column(idx, idx, ancho_maximo, formato_moneda)
        elif "hr" in col.lower() or "costo" in col.lower() and "$" not in col:
            worksheet.set_column(idx, idx, ancho_maximo, formato_decimal)
        elif pd.api.types.is_numeric_dtype(df[col]):
            worksheet.set_column(idx, idx, ancho_maximo, formato_miles)
        else:
            worksheet.set_column(idx, idx, ancho_maximo)

if st.button("Ejecutar Optimización"):
    all_summaries_list = []
    comparison_data = []
    dict_summaries = {} 
    
    output_buffer = io.BytesIO()
    
    with st.status("⏳ Iniciando motor de optimización...", expanded=True) as status:
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            
            for name, scenario_config in scenarios.items():
                st.write(f"⚙️ Optimizando escenario: **{name}**...")
                
                shifts_dict = scenario_config["shifts"]
                force_flag  = scenario_config["force_max"]
                fill_flag   = scenario_config["fill_cap"]
                
                # Recibimos el tercer valor: obj_val
                df_summary, df_detail, obj_val = generate_scenario_report(name, shifts_dict, force_flag, fill_flag)
                
                # Handle infeasible scenarios
                if "Error" in df_summary.columns:
                    df_summary.to_excel(writer, sheet_name=f"Error - {name[:20]}", index=False)
                    st.write(f"❌ **{name}**: Matemáticamente inviable o sin solución.")
                    continue
                
                df_summary.insert(0, "Escenario", name)
                all_summaries_list.append(df_summary)
                dict_summaries[name] = df_summary 
                
                cmv_total           = df_summary["CMV ($)"].sum()
                otros_egresos_total = df_summary["Costo Capital ($)"].sum()
                valor_pol_primero   = df_summary["Valor política inventario ($)"].iloc[0]
                valor_pol_ultimo    = df_summary["Valor política inventario ($)"].iloc[-1]
                delta_valor_inv     = valor_pol_ultimo - valor_pol_primero
                
                impacto_fcl         = cmv_total + otros_egresos_total + delta_valor_inv
                
                comparison_data.append({
                    "Escenario":             name,
                    "Costo Real (Función Obj)": obj_val,  # <-- NUEVA COLUMNA
                    "CMV":                   cmv_total,
                    "Otros egresos":         otros_egresos_total,
                    "Delta valor inventario": delta_valor_inv,
                    "Impacto total FCL":     impacto_fcl
                })

                safe_name = name[:20].replace("/", "-")
                dar_formato_excel(writer, df_summary, f"Sum - {safe_name}")
                dar_formato_excel(writer, df_detail, f"Det - {safe_name}")
                
                st.write(f"✅ **{name}** calculado exitosamente.")

            if all_summaries_list:
                df_consolidado = pd.concat(all_summaries_list, ignore_index=True)
                dar_formato_excel(writer, df_consolidado, "Consolidado_General")
                
                df_comparacion = pd.DataFrame(comparison_data)
                dar_formato_excel(writer, df_comparacion, "Comparacion_Escenarios")
        
        status.update(label="🎉 Optimización completada en todos los escenarios.", state="complete", expanded=False)
        
        st.session_state['opt_ejecutada'] = True
        st.session_state['comparison_data'] = comparison_data
        st.session_state['dict_summaries'] = dict_summaries
        st.session_state['excel_buffer'] = output_buffer.getvalue()

# ==========================================
# 5. UI RESULT VISUALIZATION
# ==========================================
# This section is rendered only if the saved 'opt_ejecutada' flag exists
if st.session_state.get('opt_ejecutada', False):
    st.divider()
    st.subheader("📊 Resultados de la Optimización")
    
    tab_fin, tab_esc = st.tabs(["💰 Reporte Financiero", "📋 Consolidado por Escenario"])
    
    # Tab 1: Comparative Executive Summary
    with tab_fin:
        st.write("#### Comparación de Impacto Financiero")
        formato_comparacion = {
            "Costo Real (Función Obj)": "${:,.0f}", # <-- NUEVA COLUMNA CON FORMATO
            "CMV": "${:,.0f}",
            "Otros egresos": "${:,.0f}",
            "Delta valor inventario": "${:,.0f}",
            "Impacto total FCL": "${:,.0f}"
        }
        
        # Convert comparison_data list to a Pandas DataFrame
        df_comp = pd.DataFrame(st.session_state['comparison_data'])
        
        # Mostrar la tabla
        st.dataframe(
            df_comp.style.format(formato_comparacion), 
            use_container_width=True
        )
        
        # --- NUEVO: Gráfico de barras con Altair para mostrar valores en moneda ---
        st.write("#### Gráfico: Impacto Total FCL por Escenario")
        
        # 1. Definir las barras (Aumentamos el tamaño de la letra en los ejes)
        barras = alt.Chart(df_comp).mark_bar(color='#0369a1').encode(
            x=alt.X('Escenario:N', title='', axis=alt.Axis(labelAngle=0, labelFontSize=16)),
            y=alt.Y('Impacto total FCL:Q', title='Impacto FCL ($)', axis=alt.Axis(format='$,.0f', labelFontSize=14, titleFontSize=16))
        )
        
        # 2. Definir las etiquetas numéricas (Aumentamos fontSize de 14 a 20)
        etiquetas = barras.mark_text(
            align='center',
            baseline='bottom',
            dy=-10, # Separa el texto de la barra ligeramente hacia arriba
            fontSize=20, # <-- Tamaño del número sobre la barra más grande
            fontWeight='bold'
        ).encode(
            text=alt.Text('Impacto total FCL:Q', format='$,.0f')
        )
        
        # 3. Renderizar ambas capas superpuestas
        st.altair_chart(barras + etiquetas, use_container_width=True)
        
    # Tab 2: Detailed visualization per scenario using the dictionary saved in memory
    with tab_esc:
        if st.session_state['dict_summaries']:
            st.write("#### Detalle operativo por escenario")
            
            # Selector that does not reset execution because data comes from session_state
            escenario_seleccionado = st.selectbox("Selecciona el escenario a visualizar:", list(st.session_state['dict_summaries'].keys()))
            
            df_mostrar = st.session_state['dict_summaries'][escenario_seleccionado]
            
            formato_resumen = {
                "Costo fijo total($)": "${:,.0f}",
                "Costo fijo unitario ($/Und)": "${:,.0f}",
                "Costo Capital ($)": "${:,.0f}",
                "Costo total producción ($)": "${:,.0f}",
                "Costo Bodega Externa ($)": "${:,.0f}",
                "CMV ($)": "${:,.0f}",
                "Valor inventario": "${:,.0f}",
                "Variación inventario": "${:,.0f}",
                "EBITDA (CMV)": "${:,.0f}",
                "Flujo de caja": "${:,.0f}",
                "Valor política inventario ($)": "${:,.0f}",
                "Total Producido (Und)": "{:,.0f}",
                "Total pallets almacenados": "{:,.0f}",
                "Inventario": "{:,.0f}"
            }
            
            st.dataframe(df_mostrar.style.format(formato_resumen), use_container_width=True)
        else:
            st.warning("No hay escenarios calculados para mostrar.")
            
    # Global download button that reads from the saved session
    st.download_button(
        label="📥 Descargar Reporte Final (Excel Completo)",
        data=st.session_state['excel_buffer'],
        file_name="Reporte_Final_Escenarios.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
