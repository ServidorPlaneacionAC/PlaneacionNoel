# Import necessary libraries for the application
import streamlit as st
import pandas as pd
import pyomo.environ as pyo
import io
import altair as alt
import math
import shutil, os

# Initial Streamlit page configuration (browser tab title and wide layout)
st.set_page_config(page_title="Optimización de producción", layout="wide")
st.title("🏭 Simulador programación de plantas")

# Introductory description of the application for the UI
st.markdown("""
Esta aplicación simula distintos escenarios operativos con variables mixtas 
(MINLP) y un motor comercial para minimizar el Costo Unitario Promedio.
""")

st.info("**Paso 1:** Descargar la plantilla base y llenar con los datos necesarios según las indicaciones.")

# ==========================================
# --- 1. GENERATION AND DOWNLOAD OF TEMPLATE ---
# ==========================================

def generar_plantilla():
    buffer = io.BytesIO()
    
    estructuras = {
        'Produccion': ['Material', 'Semana', 'Demanda semanal'],
        'Capacidad': ['Material', 'Unidades por hora', 'Unidades por pallet', 
                      'Inventario inicial', 'Valor inventario inicial', 
                      'Costo variable unitario', 'Inventario promedio'],
        'Disponibilidad': ['Semana', 'Turnos disponibles'],
    }
    
    instrucciones = {
        'Produccion': {
            'Material': 'Código material (Ej: 1001568).',
            'Semana': 'Formato AñoSemana (Ej: 202545). Solo caracteres numéricos.',
            'Demanda semanal': 'Número de unidades demandadas de la semana'
        },
        'Capacidad': {
            'Material': 'Código único del material (Ej: 1001568).',
            'Unidades por hora': 'Capacidad en unidades de la línea a producir en una hora',
            'Unidades por pallet': 'Cantidad de unidades que caben en un pallet.',
            'Inventario inicial': 'Cantidad de inventario actual en UMB',
            'Valor inventario inicial': 'Valor ($) del inventario actual.',
            'Costo variable unitario': 'Costo directo de producir una unidad ($).',
            'Inventario promedio': 'Política de inventario en unidades'
        },
        'Disponibilidad': {
            'Semana': 'Formato AñoSemana (Ej: 202545).',
            'Turnos disponibles': 'Turnos disponibles teóricos con capacidad full (Ej: 21).'
        },
    }

    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'top',
            'fg_color': '#D7E4BC', 'border': 1
        })

        for nombre_hoja, columnas in estructuras.items():
            df_vacio = pd.DataFrame(columns=columnas)
            df_vacio.to_excel(writer, sheet_name=nombre_hoja, index=False)
            worksheet = writer.sheets[nombre_hoja]
            for idx, col_name in enumerate(columnas):
                worksheet.write(0, idx, col_name, header_fmt)
                comentario = instrucciones.get(nombre_hoja, {}).get(col_name, "Diligenciar dato.")
                worksheet.write_comment(0, idx, comentario, {'x_scale': 2, 'y_scale': 1.5})
                worksheet.set_column(idx, idx, 20)

    buffer.seek(0)
    return buffer

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
# 2. DATA LOADING
# ==========================================

st.markdown(
    """
    <div style="background-color: #e0f2fe; padding: 16px; border-radius: 8px; color: #0369a1 ; margin-bottom: 12px">
        <span style="font-size: 16px; font-weight: bold;">Paso 2:</span> 
        <span style="font-size: 16px;">👇 Agregar archivo input en formato .xlsx</span>
    </div>
    """,
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader("Label oculto", type=['xlsx'], label_visibility="collapsed")

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    df_prod = pd.read_excel(xls, 'Produccion')
    df_disp = pd.read_excel(xls, 'Disponibilidad')
    df_cap  = pd.read_excel(xls, 'Capacidad')
    # df_par removed — parameters are now widgets

    df_prod['Semana'] = df_prod['Semana'].fillna(0).astype(int)
    df_disp['Semana'] = df_disp['Semana'].fillna(0).astype(int)
    df_prod['Material'] = df_prod['Material'].astype(str)
    df_cap['Material']  = df_cap['Material'].astype(str)

    for col in ['Unidades por hora', 'Unidades por pallet', 'Inventario inicial',
                'Valor inventario inicial', 'Costo variable unitario', 'Inventario promedio']:
        df_cap[col] = pd.to_numeric(df_cap[col], errors='coerce').astype(float)

    df_prod['Demanda semanal']    = pd.to_numeric(df_prod['Demanda semanal'], errors='coerce').astype(float)
    df_disp['Turnos disponibles'] = pd.to_numeric(df_disp['Turnos disponibles'], errors='coerce').astype(float)

    st.subheader("📝 Vista previa y edición de datos")
    tab1, tab2, tab3 = st.tabs(["Producción", "Capacidad", "Disponibilidad"])

    with tab1:
        df_prod = st.data_editor(
            df_prod, num_rows="dynamic", use_container_width=True, key="edit_prod",
            column_config={"Semana": st.column_config.NumberColumn(format="%d"),
                           "Demanda semanal": st.column_config.NumberColumn(format="localized")}
        )
    with tab2:
        df_cap = st.data_editor(
            df_cap, num_rows="dynamic", use_container_width=True, key="edit_cap",
            column_config={
                "Unidades por hora": st.column_config.NumberColumn(format="localized"),
                "Unidades por pallet": st.column_config.NumberColumn(format="localized"),
                "Inventario inicial": st.column_config.NumberColumn(format="localized"),
                "Valor inventario inicial": st.column_config.NumberColumn(format="localized"),
                "Costo variable unitario": st.column_config.NumberColumn(format="localized"),
                "Inventario promedio": st.column_config.NumberColumn(format="localized"),
            }
        )
    with tab3:
        df_disp = st.data_editor(
            df_disp, num_rows="dynamic", use_container_width=True, key="edit_disp",
            column_config={"Semana": st.column_config.NumberColumn(format="%d"),
                           "Turnos disponibles": st.column_config.NumberColumn(format="localized")}
        )

    df_prod['Semana'] = df_prod['Semana'].fillna(0).astype(int)
    df_disp['Semana'] = df_disp['Semana'].fillna(0).astype(int)
    df_prod['Material'] = df_prod['Material'].astype(str)
    df_cap['Material']  = df_cap['Material'].astype(str)

    st.sidebar.header("⚙️ Parámetros del modelo")

    h = st.sidebar.number_input(
        "Horas por turno",
        min_value=0.1, max_value=24.0,
        value=8.0, step=0.5,
        help="Duración en horas de cada turno de producción."
    )
    C_fijo = st.sidebar.number_input(
        "Costo fijo ($)",
        min_value=0,
        value=742373394, step=100,
        help="Costo fijo total de la planta por semana."
    )
    r = st.sidebar.number_input(
        "Costo de capital (tasa semanal)",
        min_value=0.0, max_value=1.0,
        value=0.0029, step=0.0001, format="%.4f",
        help="Tasa semanal del costo de capital sobre el inventario."
    )
    cap_cedi = st.sidebar.number_input(
        "Capacidad CEDI (pallets)",
        min_value=0,
        value=5000, step=1,
        help="Capacidad máxima de almacenamiento en CEDI en pallets."
    )
    c_pallet = st.sidebar.number_input(
        "Costo pallet externo ($)",
        min_value=0,
        value=15000, step=100,
        help="Costo por pallet almacenado en bodega externa por semana."
    )

    M_set = sorted(df_cap['Material'].unique().tolist()) 
    T_set = sorted(df_disp['Semana'].unique().tolist())  

    UPH    = df_cap.set_index('Material')['Unidades por hora'].to_dict()
    UPP    = df_cap.set_index('Material')['Unidades por pallet'].to_dict()
    CV     = df_cap.set_index('Material')['Costo variable unitario'].to_dict()
    I0     = df_cap.set_index('Material')['Inventario inicial'].to_dict()
    Pol    = df_cap.set_index('Material')['Inventario promedio'].fillna(0).to_dict()
    Val_I0 = df_cap.set_index('Material')['Valor inventario inicial'].to_dict()

    Dem = {(m, t): 0 for m in M_set for t in T_set}
    for index, row in df_prod.iterrows():
        mat, sem, cant = str(row['Material']), int(row['Semana']), row['Demanda semanal']
        if sem in T_set and mat in M_set: Dem[(mat, sem)] = cant

    base_shifts = {}
    for index, row in df_disp.iterrows():
        t = int(row['Semana'])
        if t in T_set:
            base_shifts[t] = int(row['Turnos disponibles'])

    # ==========================================
    # SCHEDULED DOWNTIME CONFIGURATION (MANUAL)
    # ==========================================
    st.divider()
    st.subheader("🛑 Configuración de Paros Programados")
    
    df_paros_base = pd.DataFrame({
        "Semana": list(base_shifts.keys()),
        "Turnos disponibles": list(base_shifts.values()),
        "Turnos a parar": [0] * len(base_shifts)
    })

    df_paros_edit = st.data_editor(
        df_paros_base, use_container_width=True, hide_index=True,
        disabled=["Semana", "Turnos disponibles"], 
        column_config={"Turnos a parar": st.column_config.NumberColumn("Turnos a parar", min_value=0, step=1)}
    )

    errores_paro = False
    paro_shifts = {}

    for index, row in df_paros_edit.iterrows():
        semana = int(row['Semana'])
        disp = int(row['Turnos disponibles'])
        parar = int(row['Turnos a parar'])
        
        if parar > disp:
            st.error(f"⚠️ Error en la semana {semana}: El número de turnos a parar ({parar}) no puede ser mayor a los turnos disponibles ({disp}).")
            errores_paro = True
        paro_shifts[semana] = disp - parar

    if errores_paro:
        st.stop() 

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
    def generate_scenario_report(name, max_shifts, force_max, fill_cap):
        # ✅ Initialize ALL variables BEFORE any try/except
        is_optimal = False
        is_timeout = False
        valor_funcion_objetivo = 0
        summary_data = []
        details_data = []
    
        model = pyo.ConcreteModel(name=name)
            
        model = pyo.ConcreteModel(name=name)
        model.M = pyo.Set(initialize=M_set) 
        model.T = pyo.Set(initialize=sorted_weeks, ordered=True) 

        # 2. VARIABLES MIXTAS
        # X e I continuos (Reales) para fluidez. Turnos y Estibas (Enteros)
        model.X = pyo.Var(model.M, model.T, domain=pyo.NonNegativeIntegers, bounds=(0, 5000000), initialize=100) 
        model.I = pyo.Var(model.M, model.T, domain=pyo.NonNegativeIntegers, bounds=(0, 5000000), initialize=100) 
        
        model.Y = pyo.Var(model.T, domain=pyo.NonNegativeIntegers, bounds=(0, 100), initialize=10)          
        model.P = pyo.Var(model.M, model.T, domain=pyo.NonNegativeIntegers, bounds=(0, 100000), initialize=0) 
        model.E = pyo.Var(model.T, domain=pyo.NonNegativeIntegers, bounds=(0, 100000), initialize=0)          

        model.CF_unit = pyo.Var(model.T, domain=pyo.NonNegativeReals, bounds=(0, C_fijo), initialize=C_fijo)

        # FUNCIÓN OBJETIVO
        def obj_rule(mdl):
            var_cost = sum((mdl.X[m, t] * CV[m]) + C_fijo for m in mdl.M for t in mdl.T)
            inv_cost = sum(r * mdl.I[m, t] * (CV[m] + mdl.CF_unit[t]) for m in mdl.M for t in mdl.T)
            ext_cost = sum(c_pallet * mdl.E[t] for t in mdl.T)
            return var_cost + inv_cost + ext_cost
        model.Obj = pyo.Objective(rule=obj_rule, sense=pyo.minimize)
        
        # 3. RESTRICCIONES
        def min_prod_rule(mdl, t):
            return sum(mdl.X[m, t] for m in mdl.M) >= 1
        model.MinProd = pyo.Constraint(model.T, rule=min_prod_rule)

        # Restricción Bilineal (Multiplicación entre variables para diluir costo fijo)
        def cf_unit_rule(mdl, t):
            prod_total = sum(mdl.X[m, t] for m in mdl.M)
            return mdl.CF_unit[t] * prod_total >= C_fijo
        model.CF_Unit_Constraint = pyo.Constraint(model.T, rule=cf_unit_rule)

        def shift_limit_rule(mdl, t): 
            if force_max: return mdl.Y[t] == max_shifts[t]
            else:         return mdl.Y[t] <= max_shifts[t]
        model.ShiftLimit = pyo.Constraint(model.T, rule=shift_limit_rule)

        def capacity_rule(mdl, t):
            req = sum(mdl.X[m, t] / UPH[m] for m in mdl.M)
            max_unit_time = max(1 / UPH[m] for m in mdl.M)
            return req <= (mdl.Y[t] * h) + max_unit_time
        model.Capacity = pyo.Constraint(model.T, rule=capacity_rule)

        def fill_capacity_rule(mdl, t):
            if fill_cap:
                req = sum(mdl.X[m, t] / UPH[m] for m in mdl.M) 
                max_unit_time = max(1 / UPH[m] for m in mdl.M)
                return req >= (mdl.Y[t] * h) - (max_unit_time + 0.001) 
            else:
                return pyo.Constraint.Skip   
        model.FillCapacity = pyo.Constraint(model.T, rule=fill_capacity_rule)

        def inv_balance_rule(mdl, m, t):
            prod = mdl.X[m, t]
            if t == sorted_weeks[0]: return mdl.I[m, t] == I0[m] + prod - Dem[(m, t)]
            else:                    return mdl.I[m, t] == mdl.I[m, prev_week_map[t]] + prod - Dem[(m, t)]
        model.InvBalance = pyo.Constraint(model.M, model.T, rule=inv_balance_rule)

        def inv_policy_rule(mdl, m, t): return mdl.I[m, t] >= Pol[m]
        model.InvPolicy = pyo.Constraint(model.M, model.T, rule=inv_policy_rule)

        def pallet_ceil_rule(mdl, m, t):
            return mdl.P[m, t] >= mdl.I[m, t] / UPP[m]
        model.PalletCeil = pyo.Constraint(model.M, model.T, rule=pallet_ceil_rule)

        def external_wh_rule(mdl, t):
            total_pallets = sum(mdl.P[m, t] for m in mdl.M)
            return mdl.E[t] >= total_pallets - cap_cedi
        model.ExternalWH = pyo.Constraint(model.T, rule=external_wh_rule)

        def strict_shifts_rule(mdl, t):
            if not force_max and not fill_cap: 
                req = sum(mdl.X[m, t] / UPH[m] for m in mdl.M)
                return req >= ((mdl.Y[t] - 1) * h) + 0.001
            else:
                return pyo.Constraint.Skip
        model.StrictShifts = pyo.Constraint(model.T, rule=strict_shifts_rule)
        
        # 4. SOLVER SCIP (MINLP Open Source)
        def _find_scip_exe():
            # 1. Check if it's already on PATH
            p = shutil.which('scip')
            if p:
                return p
            # 2. Look inside pyscipopt's installation directory
            try:
                import pyscipopt
                base = os.path.dirname(pyscipopt.__file__)
                candidates = [
                    os.path.join(base, 'scip'),
                    os.path.join(base, '..', 'bin', 'scip'),
                    os.path.join(base, '..', '..', 'bin', 'scip'),
                ]
                for c in candidates:
                    c = os.path.normpath(c)
                    if os.path.isfile(c):
                        return c
            except ImportError:
                pass
            return None
        
        scip_exe = _find_scip_exe()
        solver = pyo.SolverFactory('scip', executable=scip_exe) if scip_exe else pyo.SolverFactory('scip')
        
        try:
            # Ponemos tee=True para ver en el log cómo SCIP ataca el problema
            results = solver.solve(model, load_solutions=False, tee=True)
            
            is_optimal = results.solver.termination_condition == pyo.TerminationCondition.optimal
            is_timeout = results.solver.termination_condition == pyo.TerminationCondition.maxTimeLimit
            
            if is_optimal or is_timeout:
                model.solutions.load_from(results)
                valor_funcion_objetivo = pyo.value(model.Obj)
            else:
                error_msg = f"Escenario inviable. Estado: {results.solver.termination_condition}"
                error_df = pd.DataFrame([{"Error": error_msg}])
                return error_df, error_df, 0, False, False
                
        except Exception as e:
            error_msg = str(e)
            error_df = pd.DataFrame([{"Error": f"Fallo en SCIP: {error_msg}"}])
            return error_df, error_df, 0, False, False
        
        # 5. CONSTRUCCIÓN DEL REPORTE
        prev_inv_value = sum(Val_I0[m] for m in M_set) 
        prev_inv_units = {m: I0[m] for m in M_set}

        prev_inv_unit_cost = {}

        for m in M_set:
            if I0[m] > 0:
                prev_inv_unit_cost[m] = Val_I0[m] / I0[m]
            else:
                prev_inv_unit_cost[m] = CV[m]

        for t in sorted_weeks:

            y_val = int(round(pyo.value(model.Y[t])))
            disp_shifts = max_shifts[t]
            prod_und_total = sum(pyo.value(model.X[m, t]) for m in M_set)
            
            Weekly_time_req = sum(Dem[(m, t)] / UPH[m] for m in M_set)
            shifts_for_demand = math.ceil(Weekly_time_req / h)
            holgura = (shifts_for_demand * h) - Weekly_time_req
            
            pallets_extra_total = int(round(pyo.value(model.E[t])))
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
                peso_tiempo = (horas_necesarias_demanda / Weekly_time_req) if Weekly_time_req > 0 else 0
                
                horas_adic = holgura * peso_tiempo
                prod_adic = horas_adic * uph
                pallets_almacenar = (prod_adic+prod_und) / upp
                pallets_tiempo_extra = (prod_adic) / upp
                
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
                    "Producción total (Und)": prod_und,
                    "Pallets a almacenar": pallets_almacenar,
                    "Tiempo adicional asignado (hr)": horas_adic,
                    "Producción en tiempo adicional (Und)": prod_adic,
                    "Pallets en tiempo extra": pallets_tiempo_extra,
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


        return pd.DataFrame(summary_data), pd.DataFrame(details_data), valor_funcion_objetivo, is_optimal, is_timeout

# ==========================================
# 4. EXECUTION AND REPORT GENERATION
# ==========================================

def dar_formato_excel(writer, df, sheet_name):
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    workbook = writer.book
    
    formato_miles = workbook.add_format({'num_format': '#,##0'})
    formato_moneda = workbook.add_format({'num_format': '"$"#,##0'})
    formato_decimal = workbook.add_format({'num_format': '#,##0.00'})
    
    for idx, col in enumerate(df.columns):
        ancho_maximo = max(df[col].astype(str).map(len).max(), len(col)) + 2
        if "$" in col: worksheet.set_column(idx, idx, ancho_maximo, formato_moneda)
        elif "hr" in col.lower() or "costo" in col.lower() and "$" not in col:
            worksheet.set_column(idx, idx, ancho_maximo, formato_decimal)
        elif pd.api.types.is_numeric_dtype(df[col]):
            worksheet.set_column(idx, idx, ancho_maximo, formato_miles)
        else: worksheet.set_column(idx, idx, ancho_maximo)

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
                
                df_summary, df_detail, obj_val, is_optimal, is_timeout = generate_scenario_report(name, shifts_dict, force_flag, fill_flag)
                
                if "Error" in df_summary.columns:
                    mensaje_error = df_summary["Error"].iloc[0]
                    st.error(f"❌ **{name}**: {mensaje_error}")
                    continue
                
                if is_timeout and not is_optimal:
                    st.warning(f"⚠️ **{name}**: Se alcanzó el tiempo límite de 180s. Solución sub-óptima.")
                
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
                    "Costo Real (Función Obj)": obj_val, 
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
if st.session_state.get('opt_ejecutada', False):
    st.divider()
    st.subheader("📊 Resultados de la Optimización")
    
    tab_fin, tab_esc = st.tabs(["💰 Reporte Financiero", "📋 Consolidado por Escenario"])
    
    with tab_fin:
        st.write("#### Comparación de Impacto Financiero")
        
        formato_comparacion = {
            "Costo Real (Función Obj)": "${:,.0f}",
            "CMV": "${:,.0f}",
            "Otros egresos": "${:,.0f}",
            "Delta valor inventario": "${:,.0f}",
            "Impacto total FCL": "${:,.0f}"
        }
        
        df_comp = pd.DataFrame(st.session_state['comparison_data'])
        st.dataframe(df_comp.style.format(formato_comparacion), use_container_width=True)
        
        st.write("#### Costo Unitario por Semana y Escenario")

        # Build a long-format dataframe from all scenario summaries
        linechart_rows = []
        for escenario, df_esc in st.session_state['dict_summaries'].items():
            if "semana" in df_esc.columns and "Costo fijo unitario ($/Und)" in df_esc.columns:
                for _, row in df_esc.iterrows():
                    linechart_rows.append({
                        "Escenario": escenario,
                        "Semana": str(int(row["semana"])),
                        "Costo Unitario ($/Und)": row["Costo fijo unitario ($/Und)"]
                    })
                    
        # BLINDAJE APLICADO: Obligamos a crear las columnas aunque no haya datos
        df_line = pd.DataFrame(linechart_rows, columns=["Escenario", "Semana", "Costo Unitario ($/Und)"])

        # Si está vacío, mostramos advertencia en vez de explotar
        if df_line.empty:
            st.warning("⚠️ No hay datos válidos para graficar. Revisa si los escenarios fallaron o la licencia WLS no cargó.")
        else:
            # Filter widget
            all_scenarios = df_line["Escenario"].unique().tolist()
            selected_scenarios = st.multiselect(
                "Filtrar escenarios:",
                options=all_scenarios,
                default=all_scenarios
            )

            df_filtered = df_line[df_line["Escenario"].isin(selected_scenarios)]

            if df_filtered.empty:
                st.warning("Selecciona al menos un escenario para visualizar.")
            else:
                line = alt.Chart(df_filtered).mark_line(point=True, strokeWidth=2.5).encode(
                    x=alt.X("Semana:O", title="Semana", axis=alt.Axis(labelAngle=-45, labelFontSize=13)),
                    y=alt.Y("Costo Unitario ($/Und):Q", title="Costo Unitario ($/Und)",
                            axis=alt.Axis(format="$,.0f", labelFontSize=13, titleFontSize=14)),
                    color=alt.Color("Escenario:N", legend=alt.Legend(title="Escenario", labelFontSize=13)),
                    tooltip=[
                        alt.Tooltip("Escenario:N", title="Escenario"),
                        alt.Tooltip("Semana:O", title="Semana"),
                        alt.Tooltip("Costo Unitario ($/Und):Q", title="Costo Unitario", format="$,.0f")
                    ]
                ).properties(height=400).interactive()

                st.altair_chart(line, use_container_width=True)
        
    with tab_esc:
        if st.session_state['dict_summaries']:
            st.write("#### Detalle operativo por escenario")
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
            
    st.download_button(
        label="📥 Descargar Reporte Final (Excel Completo)",
        data=st.session_state['excel_buffer'],
        file_name="Reporte_Final_Escenarios.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
