import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import calendar
from datetime import datetime
import io
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# --- FUNCIÓN NUEVA: CALCULAR ARRASTRE DESDE MES ANTERIOR ---
def procesar_historial_mes_anterior(df_mes_pasado):
    """
    Toma el Excel generado el mes anterior y cuenta cuántos días 
    seguidos trabajó cada agente hasta el último día del mes.
    """
    try:
        # Identificar columnas de días (numéricas) y ordenarlas de mayor a menor
        columnas_dias = [c for c in df_mes_pasado.columns if str(c).isdigit()]
        columnas_dias.sort(key=int, reverse=True) 
        
        datos_procesados = []
        for _, fila in df_mes_pasado.iterrows():
            conteo = 0
            for dia in columnas_dias:
                if fila[dia] == 'T':
                    conteo += 1
                else:
                    break # Encontró un franco (F), deja de contar hacia atrás
            
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
            fila = {"Agente": emp['Agente'], "Tipo": emp['Tipo'], "Arrastre": int(emp.get('Dias_Acumulados', 0))}
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
            ws.column_dimensions[get_column_letter(col)].width = 18 if col == 1 else 4
    return output.getvalue()

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Gestor de Francos Pro", layout="wide")
st.title("🗓️ Asignador Automático de Francos")

with st.sidebar:
    st.header("Configuración")
    meses_es = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    mes_nombre = st.selectbox("Mes a Planificar", meses_es, index=datetime.now().month % 12)
    mes_num = meses_es.index(mes_nombre) + 1
    anio = st.number_input("Año", value=datetime.now().year, step=1)
    
    st.divider()
    modo_carga = st.radio("Método de Arrastre:", ["Cargar planilla mes pasado", "Manual (Subir personal.xlsx)"])

if modo_carga == "Cargar planilla mes pasado":
    file = st.file_uploader("Sube el Excel que descargaste el MES PASADO", type=["xlsx"])
    if file:
        df_raw = pd.read_excel(file)
        df_input = procesar_historial_mes_anterior(df_raw)
        if df_input is not None:
            st.info("✅ Arrastre calculado automáticamente desde el historial.")
            st.dataframe(df_input, height=200)
else:
    file = st.file_uploader("Subir personal.xlsx (debe tener columna 'Dias_Acumulados')", type=["xlsx"])
    if file:
        df_input = pd.read_excel(file)

if file and st.button("🚀 Generar Planificación"):
    res_df = optimizar_francos(df_input, mes_num, anio)
    if res_df is not None:
        st.success(f"¡Planificación de {mes_nombre} generada!")
        st.dataframe(res_df)
        excel_data = exportar_excel_formateado(res_df)
        st.download_button(label="📥 Descargar Excel", data=excel_data, file_name=f"Planificacion_{mes_nombre}_{anio}.xlsx")
    else:
        st.error("No se encontró solución viable.")