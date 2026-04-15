# Import necessary libraries for the application
import streamlit as st
import pandas as pd
import pyomo.environ as pyo
import io
import altair as alt

# Initial Streamlit page configuration (browser tab title and wide layout)
st.set_page_config(page_title="Optimización de producción", layout="wide")
st.title("🏭 Simulador programación de plantas")

# Introductory description of the application for the UI
st.markdown("""
Esta aplicación permite simular distintos escenarios operativos (Paro Programado, Capacidad Full y Demand Driven) 
con un motor matemático no lineal y comparar los indicadores financieros.
""")

# ==========================================
# --- 1. GENERATION AND DOWNLOAD OF TEMPLATE ---
# ==========================================

st.info("**Paso 1:** Descargar la plantilla base y llenar con los datos necesarios según las indicaciones.")

def generar_plantilla():
    """
    Creates an in-memory Excel file (buffer) containing the base structure, 
    formatted headers, and cell comments with instructions for the user.
    Note: The 'Parametros' sheet was intentionally removed to use Streamlit UI widgets instead.
    """
    buffer = io.BytesIO()
    
    # Define the expected columns for each Excel sheet required by the optimization model
    estructuras = {
        'Produccion': ['Material', 'Semana', 'Demanda semanal'],
        'Capacidad': ['Material', 'Unidades por hora', 'Unidades por pallet', 
                      'Inventario inicial', 'Valor inventario inicial', 
                      'Costo variable unitario', 'Inventario promedio'],
        'Disponibilidad': ['Semana', 'Turnos disponibles']
    }
    
    # Help texts that will be embedded as comments in the Excel header cells to guide the user
    instrucciones = {
        'Produccion': {
            'Material': 'Código material (Ej: 1001568).',
            'Semana': 'Formato AñoSemana (Ej: 202545). Solo caracteres numéricos.',
            'Demanda semanal': 'Número de unidades demandadas de la semana'
        },
        'Capacidad': {
            'Material': 'Código único del material (Ej: 1001568). Registro único por matrial',
            'Unidades por hora': 'Capacidad en unidades de la línea a producir en una hora',
            'Unidades por pallet': 'Cantidad de unidades que caben en un pallet.',
            'Inventario inicial': 'Cantidad de inventario actual en UMB',
            'Valor inventario inicial': 'Valor ($) del inventario actual.',
            'Costo variable unitario': 'Costo directo de producir una unidad ($).',
            'Inventario promedio': 'Política de inventario en unidades'
        },
        'Disponibilidad': {
            'Semana': 'Formato AñoSemana (Ej: 202545). Solo caracteres numéricos.',
            'Turnos disponibles': 'Turnos disponibles teóricos con capacidad full (Ej: 21).'
        }
    }

    # Create the Excel document using xlsxwriter as the engine for formatting capabilities
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Visual style for the header row: Green background with borders
        header_fmt = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'top',
            'fg_color': '#D7E4BC', 'border': 1
        })

        # Iterate over the structures to create each Excel sheet
        for nombre_hoja, columnas in estructuras.items():
            df_vacio = pd.DataFrame(columns=columnas)
            df_vacio.to_excel(writer, sheet_name=nombre_hoja, index=False)
            
            worksheet = writer.sheets[nombre_hoja]
            
            # Apply header format and dynamic comments to each cell
            for idx, col_name in enumerate(columnas):
                worksheet.write(0, idx, col_name, header_fmt)
                comentario = instrucciones.get(nombre_hoja, {}).get(col_name, "Diligenciar este dato.")
                worksheet.write_comment(0, idx, comentario, {'x_scale': 2, 'y_scale': 1.5})
                worksheet.set_column(idx, idx, 20)

    buffer.seek(0) # Reset buffer pointer to the start for downloading
    return buffer

# UI Layout: Download button for the empty template
col_descarga, col_vacia = st.columns([1, 4])
with col_descarga:
    st.download_button(
        label="📥 Descargar Plantilla Excel",
        data=generar_plantilla(),
        file_name="Plantilla_Input_Planta.xlsx",
        mime="application/vnd.ms-excel"
    )

st.divider()

# ==========================================
# 2. DATA LOADING & WIDGETS
# ==========================================

# UI Layout: Step 2 Instructions (Unified design)
st.info("**Paso 2:** 👇 Agregar archivo input en formato .xlsx")

# Shrink the file uploader using CSS injection for a cleaner UI
st.markdown(
    """
    <style>
        [data-testid="stFileUploadDropzone"] { min-height: 50px !important; padding: 10px !important; }
        [data-testid="stFileUploadDropzone"] svg { width: 30px !important; height: 30px !important; }
        [data-testid="stFileUploadDropzone"] div { font-size: 14px !important; }
        [data-testid="stFileUploader"] { margin-bottom: 0px !important; }
    </style>
    """,
    unsafe_allow_html=True
)

# Place the uploader inside a smaller column layout
col_uploader, col_vacia_2 = st.columns([1, 1])
with col_uploader:
    uploaded_file = st.file_uploader("Label oculto", type=['xlsx'], label_visibility="collapsed")

# Only execute the processing flow if a file has been uploaded
if uploaded_file is not None:
    # Read the different tabs of the Excel file using Pandas
    xls = pd.ExcelFile(uploaded_file)
    df_prod = pd.read_excel(xls, 'Produccion')
    df_disp = pd.read_excel(xls, 'Disponibilidad')
    df_cap  = pd.read_excel(xls, 'Capacidad')

    # Initial data cleaning and strict type conversion to avoid math errors
    df_prod['Semana'] = df_prod['Semana'].fillna(0).astype(int)
    df_disp['Semana'] = df_disp['Semana'].fillna(0).astype(int)
    df_prod['Material'] = df_prod['Material'].astype(str)
    df_cap['Material']  = df_cap['Material'].astype(str)

    # Convert operational columns to numeric, coercing errors to NaN and then to float
    for col in ['Unidades por hora', 'Unidades por pallet', 'Inventario inicial',
                'Valor inventario inicial', 'Costo variable unitario', 'Inventario promedio']:
        df_cap[col] = pd.to_numeric(df_cap[col], errors='coerce').astype(float)

    df_prod['Demanda semanal']    = pd.to_numeric(df_prod['Demanda semanal'], errors='coerce').astype(float)
    df_disp['Turnos disponibles'] = pd.to_numeric(df_disp['Turnos disponibles'], errors='coerce').astype(float)

    st.subheader("📝 Vista previa y edición de datos")
    
    # Create tabs to display and allow editing of uploaded data before optimization
    tab1, tab2, tab3 = st.tabs(["Producción", "Capacidad", "Disponibilidad"])

    with tab1:
        # data_editor allows the user to change demand values directly in the UI
        df_prod = st.data_editor(
            df_prod, num_rows="dynamic", use_container_width=True, key="edit_prod",
            column_config={
                "Semana": st.column_config.NumberColumn(format="%d"),
                "Demanda semanal": st.column_config.NumberColumn(format="localized"),
            }
        )

    with tab2:
        df_cap = st.data_editor(
            df_cap, num_rows="dynamic", use_container_width=True, key="edit_cap",
            column_config={
                "Unidades por hora":        st.column_config.NumberColumn(format="localized"),
                "Unidades por pallet":      st.column_config.NumberColumn(format="localized"),
                "Inventario inicial":       st.column_config.NumberColumn(format="localized"),
                "Valor inventario inicial": st.column_config.NumberColumn(format="localized"),
                "Costo variable unitario":  st.column_config.NumberColumn(format="localized"),
                "Inventario promedio":      st.column_config.NumberColumn(format="localized"),
            }
        )

    with tab3:
        df_disp = st.data_editor(
            df_disp, num_rows="dynamic", use_container_width=True, key="edit_disp",
            column_config={
                "Semana":             st.column_config.NumberColumn(format="%d"),
                "Turnos disponibles": st.column_config.NumberColumn(format="localized"),
            }
        )

    # Re-apply strict data types after user edits to maintain data integrity
    df_prod['Semana'] = df_prod['Semana'].fillna(0).astype(int)
    df_disp['Semana'] = df_disp['Semana'].fillna(0).astype(int)
    df_prod['Material'] = df_prod['Material'].astype(str)
    df_cap['Material']  = df_cap['Material'].astype(str)

    # ==========================================
    # GLOBAL PARAMETERS WIDGETS
    # ==========================================
    st.divider()
    st.subheader("⚙️ Parámetros Globales de la Planta")
    st.write("Estos valores aplican para toda la simulación. Modifícalos si es necesario.")

    # Column layout for global plant settings
    col_p1, col_p2, col_p3 = st.columns(3)
    
    with col_p1:
        h = st.number_input("Horas por turno", value=8.0, step=0.5, format="%.1f")
        cap_cedi = st.number_input("Capacidad interna CEDI (Pallets)", value=5000.0, step=100.0)
        
    with col_p2:
        C_fijo = st.number_input("Costo fijo semanal ($)", value=742373394.0, step=1000000.0, format="%.0f")
        c_pallet = st.number_input("Costo bodega externa ($/Pallet)", value=15000.0, step=1000.0, format="%.0f")
        
    with col_p3:
        r = st.number_input("Costo de Capital (Tasa semanal)", value=0.0029, step=0.0001, format="%.4f")


    # Construct the mathematical Sets required by the Pyomo model (Materials and Time Periods)
    M_set = sorted(df_cap['Material'].unique().tolist()) 
    T_set = sorted(df_disp['Semana'].unique().tolist())  

    # Convert capacity dataframe columns into fast-lookup dictionaries for the model
    UPH    = df_cap.set_index('Material')['Unidades por hora'].to_dict()
    UPP    = df_cap.set_index('Material')['Unidades por pallet'].to_dict()
    CV     = df_cap.set_index('Material')['Costo variable unitario'].to_dict()
    I0     = df_cap.set_index('Material')['Inventario inicial'].to_dict()
    Pol    = df_cap.set_index('Material')['Inventario promedio'].fillna(0).to_dict()
    Val_I0 = df_cap.set_index('Material')['Valor inventario inicial'].to_dict()

    # Create the demand dictionary indexed by (Material, Week)
    Dem = {(m, t): 0 for m in M_set for t in T_set}
    for index, row in df_prod.iterrows():
        mat, sem, cant = str(row['Material']), int(row['Semana']), row['Demanda semanal']
        if sem in T_set and mat in M_set: Dem[(mat, sem)] = cant

    # Extract available base shifts per week
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
    
    # Pre-populate a table where users specify how many shifts to reduce per week
    df_paros_base = pd.DataFrame({
        "Semana": list(base_shifts.keys()),
        "Turnos disponibles": list(base_shifts.values()),
        "Turnos a parar": [0] * len(base_shifts)
    })

    df_paros_edit = st.data_editor(
        df_paros_base, use_container_width=True, hide_index=True,
        disabled=["Semana", "Turnos disponibles"], 
        column_config={
            "Turnos a parar": st.column_config.NumberColumn("Turnos a parar", min_value=0, step=1)
        }
    )

    # Validate that downtime doesn't exceed total availability
    errores_paro = False
    paro_shifts = {}

    for index, row in df_paros_edit.iterrows():
        semana = int(row['Semana'])
        disp = int(row['Turnos disponibles'])
        parar = int(row['Turnos a parar'])
        
        if parar > disp:
            st.error(f"⚠️ Error en la semana {semana}: El número de turnos a parar no puede ser mayor a los disponibles.")
            errores_paro = True
        
        paro_shifts[semana] = disp - parar

    if errores_paro:
        st.stop() # Halt execution if validation fails

    # Define the configurations for each operational scenario
    scenarios = {
        "Demand Driven":   {"shifts": base_shifts, "force_max": False, "fill_cap": False},
        "Paro Programado": {"shifts": paro_shifts, "force_max": True,  "fill_cap": True}, 
        "Full Capacity":   {"shifts": base_shifts, "force_max": True,  "fill_cap": True},
        "Paro Óptimo":     {"shifts": base_shifts, "force_max": False, "fill_cap": True} 
    }

    sorted_weeks = sorted(T_set)
    prev_week_map = {wk: sorted_weeks[i-1] for i, wk in enumerate(sorted_weeks) if i > 0}

# ==========================================
# 3. OPTIMIZATION MODEL (MATHEMATICAL ENGINE)
# ==========================================
def generate_scenario_report(name, max_shifts, force_max, fill_cap, r_val, c_pallet_val, c_fijo_val, h_val, cap_cedi_val):
        """
        Constructs and solves an optimization model.
        Determines optimal production (X), inventory (I), and shifts (Y).
        """
        model = pyo.ConcreteModel(name=name)
        
        model.M = pyo.Set(initialize=M_set) 
        model.T = pyo.Set(initialize=sorted_weeks, ordered=True)

        # Variables
        model.X = pyo.Var(model.M, model.T, domain=pyo.NonNegativeReals, initialize=100)
        
        # CHANGE: Domain changed to NonNegativeIntegers to ensure only whole shifts are used
        model.Y = pyo.Var(model.T, domain=pyo.NonNegativeIntegers) 
        
        model.I = pyo.Var(model.M, model.T, domain=pyo.NonNegativeReals) 
        model.P = pyo.Var(model.M, model.T, domain=pyo.NonNegativeReals) 
        model.E = pyo.Var(model.T, domain=pyo.NonNegativeReals)          

        # --- NON-LINEAR OBJECTIVE FUNCTION ---
        def obj_rule(mdl):
            # 1. Variable Production Cost
            costo_produccion_var = sum(mdl.X[m, t] * CV[m] for m in mdl.M for t in mdl.T)
            
            # 2. Fixed Cost (Constant)
            costo_fijo_horizonte = sum(c_fijo_val for t in mdl.T)
            
            # 3. Non-Linear Inventory Cost using Total Unitary Cost
            costo_inventario_nl = 0
            for t in mdl.T:
                total_prod_t = sum(mdl.X[m, t] for m in mdl.M)
                # Fixed Cost Per Unit = Weekly Fixed Cost / Total Weekly Production
                cf_unitario_t = c_fijo_val / (total_prod_t + 0.0001)
                for m in mdl.M:
                    valor_unitario_total = CV[m] + cf_unitario_t
                    costo_inventario_nl += r_val * valor_unitario_total * mdl.I[m, t]
            
            # 4. External Warehouse Cost
            costo_bodega = sum(c_pallet_val * mdl.E[t] for t in mdl.T)
            
            return costo_produccion_var + costo_fijo_horizonte + costo_inventario_nl + costo_bodega
            
        model.Obj = pyo.Objective(rule=obj_rule, sense=pyo.minimize)

        def shift_limit_rule(mdl, t): 
            # If force_max is True (Paro Programado or Full Capacity)
            if force_max: 
                return mdl.Y[t] == max_shifts[t]  # MUST be equality (==)
            else:
                return mdl.Y[t] <= max_shifts[t]  # Allowed to use less (<=)
        model.ShiftLimit = pyo.Constraint(model.T, rule=shift_limit_rule)
        
        # Constraint: Manufacturing Capacity
        def capacity_rule(mdl, t):
            req = sum(mdl.X[m, t] / UPH[m] for m in mdl.M)
            return req <= mdl.Y[t] * h_val
        model.Capacity = pyo.Constraint(model.T, rule=capacity_rule)

        # Constraint: Fill Capacity
        def fill_capacity_rule(mdl, t):
            if fill_cap:
                req = sum(mdl.X[m, t] / UPH[m] for m in mdl.M) 
                # When using whole shifts, we ensure production is as close to the full shift as possible
                return req >= (mdl.Y[t] * h_val) - (h_val - 0.001) 
            else:
                return pyo.Constraint.Skip   
        model.FillCapacity = pyo.Constraint(model.T, rule=fill_capacity_rule)

        # Constraint: Inventory Balance
        def inv_balance_rule(mdl, m, t):
            prod = mdl.X[m, t]
            if t == sorted_weeks[0]: return mdl.I[m, t] == I0[m] + prod - Dem[(m, t)]
            else:                    return mdl.I[m, t] == mdl.I[m, prev_week_map[t]] + prod - Dem[(m, t)]
        model.InvBalance = pyo.Constraint(model.M, model.T, rule=inv_balance_rule)

        # Constraint: Inventory Policy
        def inv_policy_rule(mdl, m, t): return mdl.I[m, t] >= Pol[m]
        model.InvPolicy = pyo.Constraint(model.M, model.T, rule=inv_policy_rule)

        # Constraint: Pallets
        def pallet_ceil_rule(mdl, m, t):
            return mdl.P[m, t] >= mdl.I[m, t] / UPP[m]
        model.PalletCeil = pyo.Constraint(model.M, model.T, rule=pallet_ceil_rule)

        # Constraint: External warehouse
        def external_wh_rule(mdl, t):
            total_pallets = sum(mdl.P[m, t] for m in mdl.M)
            return mdl.E[t] >= total_pallets - cap_cedi_val
        model.ExternalWH = pyo.Constraint(model.T, rule=external_wh_rule)

        # Solver Configuration
        # NOTE: IPOPT is a continuous solver. If it ignores the Integer domain, 
        # you may need to use 'bonmin' or 'couenne' for true discrete shifts.
        solver = pyo.SolverFactory('ipopt') 
        solver.options['max_cpu_time'] = 180 
        
        results = solver.solve(model, load_solutions=False)

        # Check solver termination status
        is_optimal = results.solver.termination_condition == pyo.TerminationCondition.optimal
        is_timeout = results.solver.termination_condition == pyo.TerminationCondition.maxTimeLimit
        
        # Handle cases where no feasible solution is found
        try:
            model.solutions.load_from(results)
        except:
            error_df = pd.DataFrame([{"Error": f"Escenario {name} es matemáticamente inviable."}])
            return error_df, error_df, None, False, False
        
        # Calculate final results and financial KPIs
        valor_funcion_objetivo = pyo.value(model.Obj)
        summary_data = []
        details_data = []
        
        # Track inventory valuation across periods
        prev_inv_value = sum(Val_I0[m] for m in M_set) 
        prev_inv_units = {m: I0[m] for m in M_set}
        
        prev_inv_unit_cost = {}
        for m in M_set:
            if I0[m] > 0: prev_inv_unit_cost[m] = Val_I0[m] / I0[m]
            else:         prev_inv_unit_cost[m] = CV[m]

        for t in sorted_weeks:
            # IPOPT returns continuous floats; round for practical reporting
            y_val = round(pyo.value(model.Y[t]), 2)
            disp_shifts = max_shifts[t]
            prod_und_total = sum(pyo.value(model.X[m, t]) for m in M_set)
            
            time_req = sum(pyo.value(model.X[m, t]) / UPH[m] for m in M_set)
            holgura = (y_val * h_val) - time_req 
            
            pallets_extra_total = pyo.value(model.E[t])
            total_pallets_inv = sum(pyo.value(model.I[m, t]) / UPP[m] for m in M_set)
            inventario_und_total = sum(pyo.value(model.I[m, t]) for m in M_set)
            
            # Allocation of Fixed Costs per unit produced
            cf_total = c_fijo_val 
            cf_unitario = cf_total / prod_und_total if prod_und_total > 0 else 0
            cv_total = sum(pyo.value(model.X[m, t]) * CV[m] for m in M_set)
            costo_total_prod = cv_total + cf_total
            costo_bodega_externa = pallets_extra_total * c_pallet_val
            
            val_inv_semana = 0
            costo_cap_semana = 0
            cmv_semana = 0
            
            # Individual material detailed calculation
            for m in M_set:
                demand = Dem[(m, t)]
                uph = UPH[m]
                cv_unit = CV[m]
                upp = UPP[m]
                
                prod_und = pyo.value(model.X[m, t])
                horas_necesarias_demanda = demand / uph
                horas_totales_mat = prod_und / uph
                peso_tiempo = (horas_totales_mat / time_req) if time_req > 0 else 0
                
                horas_adic = holgura * peso_tiempo
                prod_adic = horas_adic * uph
                pallets_almacenar = (prod_und / upp) if upp > 0 else 0
                
                inv_final = pyo.value(model.I[m, t])
                inv_anterior = prev_inv_units[m]
                
                costo_var_mat = prod_und * cv_unit
                costo_unitario_total = cv_unit + cf_unitario
                
                # Valuing inventory using logic consistent with the Objective Function
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
                costo_cap_semana += (r_val * costo_inv_final)
                cmv_semana       += cmv_mat
                
                # Build the granular detail list
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

            # Weekly aggregation
            var_inv = val_inv_semana - prev_inv_value
            prev_inv_value = val_inv_semana 
            
            total_pol_und = sum(Pol[m] for m in M_set)
            costo_promedio_global = (val_inv_semana / inventario_und_total) if inventario_und_total > 0 else 0
            valor_politica_inv = total_pol_und * costo_promedio_global

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
                "Costo Capital ($)": val_inv_semana * r_val,
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

        return pd.DataFrame(summary_data), pd.DataFrame(details_data), valor_funcion_objetivo, is_optimal, is_timeout

# ==========================================
# 4. EXECUTION AND REPORT GENERATION
# ==========================================

def dar_formato_excel(writer, df, sheet_name):
    """
    Applies conditional formatting and numeric masks to the output Excel sheets.
    """
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    workbook = writer.book
    
    # Define formatting styles
    formato_miles = workbook.add_format({'num_format': '#,##0'})
    formato_moneda = workbook.add_format({'num_format': '"$"#,##0'})
    formato_decimal = workbook.add_format({'num_format': '#,##0.00'})
    
    # Auto-adjust column width and apply specific masks based on column names
    for idx, col in enumerate(df.columns):
        ancho_maximo = max(df[col].astype(str).map(len).max(), len(col)) + 2
        
        if "$" in col: worksheet.set_column(idx, idx, ancho_maximo, formato_moneda)
        elif "hr" in col.lower() or "costo" in col.lower() and "$" not in col:
            worksheet.set_column(idx, idx, ancho_maximo, formato_decimal)
        elif pd.api.types.is_numeric_dtype(df[col]):
            worksheet.set_column(idx, idx, ancho_maximo, formato_miles)
        else: worksheet.set_column(idx, idx, ancho_maximo)

# Main optimization trigger
if st.button("Ejecutar Optimización"):
    all_summaries_list = []
    comparison_data = []
    dict_summaries = {} 
    
    output_buffer = io.BytesIO()
    
    # Progress status bar for UI
    with st.status("⏳ Iniciando motor de optimización...", expanded=True) as status:
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            
            for name, scenario_config in scenarios.items():
                st.write(f"⚙️ Optimizando escenario: **{name}**...")
                
                shifts_dict = scenario_config["shifts"]
                force_flag  = scenario_config["force_max"]
                fill_flag   = scenario_config["fill_cap"]
                
                # Injects UI parameters into the model engine
                df_summary, df_detail, obj_val, is_optimal, is_timeout = generate_scenario_report(
                    name, shifts_dict, force_flag, fill_flag, r, c_pallet, C_fijo, h, cap_cedi
                )
                
                # Catch mathematical errors
                if "Error" in df_summary.columns:
                    st.error(f"❌ **{name}**: No se pudo calcular. Revisa los datos o aumenta el tiempo.")
                    continue
                
                if is_timeout and not is_optimal:
                    st.warning(f"⚠️ **{name}**: Tiempo límite alcanzado. Solución sub-óptima.")
                
                # Compile results for cross-scenario comparison
                df_summary.insert(0, "Escenario", name)
                all_summaries_list.append(df_summary)
                dict_summaries[name] = df_summary 
                
                # Calculating total impact (Cash Flow / EBITDA)
                cmv_total           = df_summary["CMV ($)"].sum()
                otros_egresos_total = df_summary["Costo Capital ($)"].sum()
                valor_pol_primero   = df_summary["Valor política inventario ($)"].iloc[0]
                valor_pol_ultimo    = df_summary["Valor política inventario ($)"].iloc[-1]
                delta_valor_inv     = valor_pol_ultimo - valor_pol_primero
                impacto_fcl         = cmv_total + otros_egresos_total + delta_valor_inv
                
                comparison_data.append({
                    "Escenario":             name,
                    "Costo Real (Función Obj)": obj_val,
                    "CMV":                   cmv_total,
                    "Otros egresos":         otros_egresos_total,
                    "Delta valor inventario": delta_valor_inv,
                    "Impacto total FCL":      impacto_fcl
                })

                # Export scenario sheets to Excel
                safe_name = name[:20].replace("/", "-")
                dar_formato_excel(writer, df_summary, f"Sum - {safe_name}")
                dar_formato_excel(writer, df_detail, f"Det - {safe_name}")
                
                st.write(f"✅ **{name}** calculado exitosamente.")

            # Create global sheets in Excel
            if all_summaries_list:
                df_consolidado = pd.concat(all_summaries_list, ignore_index=True)
                dar_formato_excel(writer, df_consolidado, "Consolidado_General")
                
                df_comparacion = pd.DataFrame(comparison_data)
                dar_formato_excel(writer, df_comparacion, "Comparacion_Escenarios")
        
        status.update(label="🎉 Optimización completada.", state="complete", expanded=False)
        
        # Persistence of data in session state for UI display
        st.session_state['opt_ejecutada'] = True
        st.session_state['comparison_data'] = comparison_data
        st.session_state['dict_summaries'] = dict_summaries
        st.session_state['excel_buffer'] = output_buffer.getvalue()

# ==========================================
# 5. UI RESULT VISUALIZATION
# ==========================================
if st.session_state.get('opt_ejecutada', False):
    st.divider()
    st.subheader("📊 Resultados de la Optimización")
    
    tab_fin, tab_esc = st.tabs(["💰 Reporte Financiero", "📋 Consolidado por Escenario"])
    
    with tab_fin:
        st.write("#### Comparación de Impacto Financiero")
        
        formato_comparacion = {
            "Costo Real (Función Obj)": "${:,.0f}",
            "CMV": "${:,.0f}", "Otros egresos": "${:,.0f}",
            "Delta valor inventario": "${:,.0f}", "Impacto total FCL": "${:,.0f}"
        }
        
        df_comp = pd.DataFrame(st.session_state['comparison_data'])
        st.dataframe(df_comp.style.format(formato_comparacion), use_container_width=True)
        
        # High-level comparison chart using Altair
        st.write("#### Gráfico: Impacto Total FCL por Escenario")
        
        barras = alt.Chart(df_comp).mark_bar(color='#0369a1').encode(
            x=alt.X('Escenario:N', title='', axis=alt.Axis(labelAngle=0, labelFontSize=16)),
            y=alt.Y('Impacto total FCL:Q', title='Impacto FCL ($)', axis=alt.Axis(format='$,.0f', labelFontSize=14, titleFontSize=16))
        )
        
        etiquetas = barras.mark_text(
            align='center', baseline='bottom', dy=-10, fontSize=20, fontWeight='bold'
        ).encode(text=alt.Text('Impacto total FCL:Q', format='$,.0f'))
        
        st.altair_chart(barras + etiquetas, use_container_width=True)
        
    with tab_esc:
        if st.session_state['dict_summaries']:
            st.write("#### Detalle operativo por escenario")
            # Filter results by selecting specific scenarios
            escenario_seleccionado = st.selectbox("Selecciona el escenario a visualizar:", list(st.session_state['dict_summaries'].keys()))
            df_mostrar = st.session_state['dict_summaries'][escenario_seleccionado]
            
            formato_resumen = {
                "Costo fijo total($)": "${:,.0f}", "Costo fijo unitario ($/Und)": "${:,.0f}",
                "Costo Capital ($)": "${:,.0f}", "Costo total producción ($)": "${:,.0f}",
                "Costo Bodega Externa ($)": "${:,.0f}", "CMV ($)": "${:,.0f}",
                "Valor inventario": "${:,.0f}", "Variación inventario": "${:,.0f}",
                "EBITDA (CMV)": "${:,.0f}", "Flujo de caja": "${:,.0f}",
                "Valor política inventario ($)": "${:,.0f}", "Total Producido (Und)": "{:,.0f}",
                "Total pallets almacenados": "{:,.0f}", "Inventario": "{:,.0f}"
            }
            
            st.dataframe(df_mostrar.style.format(formato_resumen), use_container_width=True)
        else:
            st.warning("No hay escenarios calculados para mostrar.")
            
    # Allow user to download the generated Excel report
    st.download_button(
        label="📥 Descargar Reporte Final (Excel Completo)",
        data=st.session_state['excel_buffer'],
        file_name="Reporte_Final_Escenarios_IPOPT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
