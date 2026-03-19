import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import calendar
from datetime import datetime
import io
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Gestor de Francos AR", 
    page_icon="🗓️", 
    layout="wide"
)

# --- OCULTAR ELEMENTOS DE INTERFAZ (VERSIÓN AGRESIVA) ---
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            
            /* Oculta TODO el contenedor de botones de la derecha (Fork, GitHub, Deploy) */
            header [data-testid="stHeaderActionElements"], 
            .stDeployButton, 
            .viewerBadge_container__1QS13, 
            [data-testid="manage-app-button"],
            .stAppDeployButton {
                display: none !important;
            }

            /* Esto oculta específicamente el botón de Fork y GitHub que aparecen en repos públicos */
            header a, header button {
                display: none !important;
            }

            /* Pero... necesitamos que la FLECHA del menú lateral sí se vea */
            /* La flecha no está dentro de 'stHeaderActionElements', así que debería sobrevivir */
            
            header {
                background-color: rgba(0,0,0,0) !important;
            }
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# --- LÓGICA: CALCULAR ARRASTRE DESDE MES ANTERIOR ---
def procesar_historial_mes_anterior(df_mes_pasado):
    try:
        columnas_dias = [c for c in df_mes_pasado.columns if str(c).isdigit()]
        columnas_dias.sort(key=int, reverse=True) 
        
        datos_procesados = []
        for _, fila in df_mes_pasado.iterrows():
            conteo = 0
            for dia in columnas_dias:
                if fila[dia] == 'T':
                    conteo += 1
                else:
                    break
            
            datos_procesados.append({
                'Agente': fila['Agente'],
                'Tipo': fila['Tipo'],
                'Dias_Acumulados': conteo
            })
        return pd.DataFrame(datos_procesados)
    except Exception as e:
        st.error(f"Error al procesar el archivo del mes anterior: {e}")
        return None

# --- LÓGICA DEL OPTIMIZADOR ---
def optimizar_francos(df_empleados, mes_num, anio):
    num_dias = calendar.monthrange(anio, mes_num)[1]
    domingos = [d for d in range(1, num_dias + 1) if calendar.weekday(anio, mes_num, d) == 6]
    
    cant_francos_objetivo = len(domingos)
    model = cp_model.CpModel()
    empleados = df_empleados.to_dict('records')
    num_emp = len(empleados)
    
    x = {}
    for i in range(num_emp):
        for d in range(1, num_dias + 1):
            x[i, d] = model.NewBoolVar(f'x_{i}_{d}')

    for i, emp in enumerate(empleados):
        tipo = str(emp['Tipo']).strip().capitalize()
        arrastre = int(emp.get('Dias_Acumulados', 0))
        
        model.Add(sum(x[i, d] for d in range(1, num_dias + 1)) == (num_dias - cant_francos_objetivo))

        if tipo == 'Tercerizado':
            if arrastre >= 7:
                model.Add(x[i, 1] == 0)
            elif arrastre > 0:
                limite = min(8 - arrastre, num_dias + 1)
                if limite > 1:
                    model.Add(sum(x[i, d] for d in range(1, limite)) <= (7 - arrastre))
            
            for d in range(1, num_dias - 6):
                model.Add(sum(x[i, d + j] for j in range(8)) <= 7)

            semanas = [range(1,8), range(8,15), range(15,22), range(22,num_dias + 1)]
            for s in semanas:
                dias_v = [d for d in s if d <= num_dias]
                if len(dias_v) >= 5:
                    model.Add(sum(x[i, d] for d in dias_v) <= len(dias_v) - 1)

        elif tipo == 'Propio':
            if arrastre >= 8:
                model.Add(x[i, 1] == 0)
            elif arrastre > 0:
                limite = min(9 - arrastre, num_dias + 1)
                if limite > 1:
                    model.Add(sum(x[i, d] for d in range(1, limite)) <= (8 - arrastre))
            
            for d in range(1, num_dias - 7):
                model.Add(sum(x[i, d + j] for j in range(9)) <= 8)

    min_presentes = max(1, int(num_emp * 0.20))
    for d in range(1, num_dias + 1):
        model.Add(sum(x[i, d] for i in range(num_emp)) >= min_presentes)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 10.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        res = []
        for i, emp in enumerate(empleados):
            fila = {"Agente": emp['Agente'], "Tipo": emp['Tipo'], "Arrastre Inicial": int(emp.get('Dias_Acumulados', 0))}
            for d in range(1, num_dias + 1):
                fila[f"{d}"] = "T" if solver.Value(x[i, d]) == 1 else "F"
            res.append(fila)
        return pd.DataFrame(res)
    return None

def exportar_excel_formateado(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Planificacion')
        ws = writer.sheets['Planificacion']
        f_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        f_font = Font(color="9C0006", bold=True)
        t_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        t_font = Font(color="006100")

        for row in range(2, ws.max_row + 1):
            for col in range(4, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                if cell.value == "F":
                    cell.fill = f_fill
                    cell.font = f_font
                elif cell.value == "T":
                    cell.fill = t_fill
                    cell.font = t_font
                cell.alignment = Alignment(horizontal='center')
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 20 if col == 1 else 4
    return output.getvalue()

# --- INTERFAZ DE USUARIO ---
st.title("🗓️ Gestor de Francos AR")
st.markdown("---")

with st.sidebar:
    st.header("⚙️ Configuración")
    meses_es = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    
    # Preseleccionar el mes siguiente al actual
    mes_default = (datetime.now().month) % 12
    mes_nombre = st.selectbox("Mes a planificar", meses_es, index=mes_default)
    mes_num = meses_es.index(mes_nombre) + 1
    anio = st.number_input("Año", value=datetime.now().year, step=1)
    
    st.divider()
    st.subheader("📥 Carga de Datos")
    modo_carga = st.radio("Método para calcular arrastre:", 
                          ["Usar planilla mes pasado", "Manual (Subir personal.xlsx)"])
    
    st.divider()
    st.caption("© 2026 Fernando. Todos los derechos reservados.")

# Procesamiento de archivos según el modo seleccionado
df_input = None
if modo_carga == "Usar planilla mes pasado":
    file = st.file_uploader("Sube el Excel generado el mes anterior", type=["xlsx"])
    if file:
        df_raw = pd.read_excel(file)
        df_input = procesar_historial_mes_anterior(df_raw)
        if df_input is not None:
            st.success("✅ Historial procesado. Arrastre calculado automáticamente.")
            with st.expander("Ver datos de arrastre detectados"):
                st.table(df_input)
else:
    file = st.file_uploader("Sube el archivo de personal (debe tener columna 'Dias_Acumulados')", type=["xlsx"])
    if file:
        df_input = pd.read_excel(file)

# Botón de acción principal
if df_input is not None:
    if st.button("🚀 Generar Planificación de Francos"):
        with st.spinner("Calculando la mejor combinación de francos..."):
            res_df = optimizar_francos(df_input, mes_num, anio)
            
            if res_df is not None:
                st.balloons()
                st.subheader(f"📅 Planificación de {mes_nombre} {anio}")
                st.dataframe(res_df, use_container_width=True)
                
                excel_data = exportar_excel_formateado(res_df)
                st.download_button(
                    label="📥 Descargar Excel con Colores", 
                    data=excel_data, 
                    file_name=f"Planificacion_{mes_nombre}_{anio}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("❌ No se encontró una solución viable. Prueba revisando si hay demasiados agentes con arrastre crítico (7 u 8 días) al inicio del mes.")