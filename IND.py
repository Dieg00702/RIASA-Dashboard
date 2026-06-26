import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# Cargar tus datos reales desde la Tabla8
def cargar_tabla_excel(ruta_archivo, nombre_hoja, nombre_tabla):
    wb = load_workbook(ruta_archivo, read_only=False, data_only=True)
    ws = wb[nombre_hoja]
    if nombre_tabla in ws.tables:
        tabla_obj = ws.tables[nombre_tabla]
        rango_celdas = ws[tabla_obj.ref] 
        filas = [[celda.value for celda in fila] for fila in rango_celdas]
        wb.close()
        return pd.DataFrame(filas[1:], columns=filas[0])
    wb.close()
    return pd.DataFrame()



def mostrar_indicadores():
        df = cargar_tabla_excel("Rep RIASA.xlsm", "Medidores Grafica", "Tabla8")
        if not df.empty:
            # 1. Extracción de las métricas desde tu Excel
            real_pza = pd.to_numeric(df.loc[0, "Real Pza"], errors="coerce") or 0
            plan_pza_min = pd.to_numeric(df.loc[0, "Plan Pza"], errors="coerce") or 0
            plan_mensual_max = pd.to_numeric(df.loc[0, "Plan mensual"], errors="coerce") or 0
            
            # CORRECCIÓN: El porcentaje ahora se calcula sobre el "Plan Pza" (Objetivo Mínimo)
            porcentaje = (real_pza / plan_pza_min * 100) if plan_pza_min > 0 else 0
            
            # La barra de progreso se llena tomando como tope máximo (100%) el Plan Mensual
            ancho_barra = min(max((real_pza / plan_mensual_max * 100), 0), 100) if plan_mensual_max > 0 else 0

            # Formatear números para la interfaz
            real_formateado = f"{real_pza:,.0f}"
            min_formateado = f"{plan_pza_min:,.0f}"
            max_formateado = f"{plan_mensual_max:,.0f}"
            porcentaje_formateado = f"{porcentaje:.1f}%"

            # 2. Mantener la lógica de color dinámico basada en tu nuevo porcentaje
            if porcentaje >= 95:
                color_tema = "#4cd964"       # Verde
                fondo_porcentaje = "rgba(76, 217, 100, 0.15)"
            elif porcentaje >= 80:
                color_tema = "#ffcc00"       # Amarillo/Naranja (Tu 70% entraría en Rojo por estar abajo de 80)
                fondo_porcentaje = "rgba(255, 204, 0, 0.15)"
            else:
                color_tema = "#ff3b30"       # Rojo
                fondo_porcentaje = "rgba(255, 59, 48, 0.15)"
            # ==========================================
            # 3. DISEÑO ESTRUCTURADO HTML / CSS TRASLADADO A STREAMLIT
            # ==========================================
            html_tarjeta_inyectora = f"""
            <div style="
                background-color: #1a1b23; 
                border: 1px solid #2d2f39; 
                border-radius: 12px; 
                padding: 24px 28px; 
                width: 100%;
                max-width: 460px;
                box-shadow: 0 4px 15px rgba(0,0,0,0.25);
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            ">
                <p style="color: #808495; font-size: 26px; font-weight: bold; letter-spacing:.5px; margin: 0; text-transform: uppercase;">INYECTORA</p>
                
            
            <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 12px; margin-bottom: 8px;">
                <span style="color:#ffffff; font-size: 70px; font-weight: 500; margin: 5px 0 2px 0; line-height: 1.1;">{real_formateado}</span>
                    <span style="
                        color: {color_tema}; 
                        background-color: {fondo_porcentaje}; 
                        border: 1px solid {color_tema}50;
                        padding: 6px 14px; 
                        border-radius: 20px; 
                        font-size: 15px; 
                        font-weight: bold;
                    ">{porcentaje_formateado}</span>
                </div>
                
            <p style="color: #636674; font-size: 26 px; margin: 0; line-height: 1.4;">
                    Obj. min: <span style="color: #808495;">{min_formateado}</span> &middot; Obj. max: <span style="color: #808495;">{max_formateado}</span>
                </p>
                
            <div style="background-color: #262730; width: 100%; border-radius: 6px; height: 8px; margin-top: 18px; margin-bottom: 4px;">
                    <div style="background-color: {color_tema}; width: {ancho_barra}%; border-radius: 6px; height: 100%;"></div>
                </div>
            </div>
            """


            
        st.markdown(html_tarjeta_inyectora, unsafe_allow_html=True)
        st.markdown('<div style="margin-top: 40px;"></div>', unsafe_allow_html=True)
        df = cargar_tabla_excel("Rep RIASA.xlsm", "Medidores Grafica", "Tabla12")
        if not df.empty:
            # 1. Extracción de las métricas desde tu Excel
            real_pza = pd.to_numeric(df.loc[0, "Real Pza"], errors="coerce") or 0
            plan_pza_min = pd.to_numeric(df.loc[0, "Plan Pza"], errors="coerce") or 0
            plan_mensual_max = pd.to_numeric(df.loc[0, "Plan mensual"], errors="coerce") or 0
            
            # CORRECCIÓN: El porcentaje ahora se calcula sobre el "Plan Pza" (Objetivo Mínimo)
            porcentaje = (real_pza / plan_pza_min * 100) if plan_pza_min > 0 else 0
            
            # La barra de progreso se llena tomando como tope máximo (100%) el Plan Mensual
            ancho_barra = min(max((real_pza / plan_mensual_max * 100), 0), 100) if plan_mensual_max > 0 else 0

            # Formatear números para la interfaz
            real_formateado = f"{real_pza:,.0f}"
            min_formateado = f"{plan_pza_min:,.0f}"
            max_formateado = f"{plan_mensual_max:,.0f}"
            porcentaje_formateado = f"{porcentaje:.1f}%"

            # 2. Mantener la lógica de color dinámico basada en tu nuevo porcentaje
            if porcentaje >= 95:
                color_tema = "#4cd964"       # Verde
                fondo_porcentaje = "rgba(76, 217, 100, 0.15)"
            elif porcentaje >= 80:
                color_tema = "#ffcc00"       # Amarillo/Naranja (Tu 70% entraría en Rojo por estar abajo de 80)
                fondo_porcentaje = "rgba(255, 204, 0, 0.15)"
            else:
                color_tema = "#ff3b30"       # Rojo
                fondo_porcentaje = "rgba(255, 59, 48, 0.15)"
            # ==========================================
            # 3. DISEÑO ESTRUCTURADO HTML / CSS TRASLADADO A STREAMLIT
            # ==========================================
            html_tarjeta_inyectora = f"""
            <div style="
                background-color: #1a1b23; 
                border: 1px solid #2d2f39; 
                border-radius: 12px; 
                padding: 24px 28px; 
                width: 100%;
                max-width: 460px;
                box-shadow: 0 4px 15px rgba(0,0,0,0.25);
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            ">
                <p style="color: #808495; font-size: 26px; font-weight: bold; letter-spacing:.5px; margin: 0; text-transform: uppercase;">PELETIZADO</p>
                
            
            <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 12px; margin-bottom: 8px;">
                <span style="color:#ffffff; font-size: 70px; font-weight: 500; margin: 5px 0 2px 0; line-height: 1.1;">{real_formateado}</span>
                    <span style="
                        color: {color_tema}; 
                        background-color: {fondo_porcentaje}; 
                        border: 1px solid {color_tema}50;
                        padding: 6px 14px; 
                        border-radius: 20px; 
                        font-size: 15px; 
                        font-weight: bold;
                    ">{porcentaje_formateado}</span>
                </div>
                
            <p style="color: #636674; font-size: 26 px; margin: 0; line-height: 1.4;">
                    Obj. min: <span style="color: #808495;">{min_formateado}</span> &middot; Obj. max: <span style="color: #808495;">{max_formateado}</span>
                </p>
                
            <div style="background-color: #262730; width: 100%; border-radius: 6px; height: 8px; margin-top: 18px; margin-bottom: 4px;">
                    <div style="background-color: {color_tema}; width: {ancho_barra}%; border-radius: 6px; height: 100%;"></div>
                </div>
            </div>
            """
        st.markdown(html_tarjeta_inyectora, unsafe_allow_html=True)
        st.markdown('<div style="margin-top: 40px;"></div>', unsafe_allow_html=True)
        df = cargar_tabla_excel("Rep RIASA.xlsm", "Medidores Grafica", "Tabla9")
        if not df.empty:
            # 1. Extracción de las métricas desde tu Excel
            real_pza = pd.to_numeric(df.loc[0, "Real Pza"], errors="coerce") or 0
            plan_pza_min = pd.to_numeric(df.loc[0, "Plan Pza"], errors="coerce") or 0
            plan_mensual_max = pd.to_numeric(df.loc[0, "Plan mensual"], errors="coerce") or 0
            
            # CORRECCIÓN: El porcentaje ahora se calcula sobre el "Plan Pza" (Objetivo Mínimo)
            porcentaje = (real_pza / plan_pza_min * 100) if plan_pza_min > 0 else 0
            
            # La barra de progreso se llena tomando como tope máximo (100%) el Plan Mensual
            ancho_barra = min(max((real_pza / plan_mensual_max * 100), 0), 100) if plan_mensual_max > 0 else 0

            # Formatear números para la interfaz
            real_formateado = f"{real_pza:,.0f}"
            min_formateado = f"{plan_pza_min:,.0f}"
            max_formateado = f"{plan_mensual_max:,.0f}"
            porcentaje_formateado = f"{porcentaje:.1f}%"

            # 2. Mantener la lógica de color dinámico basada en tu nuevo porcentaje
            if porcentaje >= 95:
                color_tema = "#4cd964"       # Verde
                fondo_porcentaje = "rgba(76, 217, 100, 0.15)"
            elif porcentaje >= 80:
                color_tema = "#ffcc00"       # Amarillo/Naranja (Tu 70% entraría en Rojo por estar abajo de 80)
                fondo_porcentaje = "rgba(255, 204, 0, 0.15)"
            else:
                color_tema = "#ff3b30"       # Rojo
                fondo_porcentaje = "rgba(255, 59, 48, 0.15)"
            # ==========================================
            # 3. DISEÑO ESTRUCTURADO HTML / CSS TRASLADADO A STREAMLIT
            # ==========================================
            html_tarjeta_inyectora = f"""
            <div style="
                background-color: #1a1b23; 
                border: 1px solid #2d2f39; 
                border-radius: 12px; 
                padding: 24px 28px; 
                width: 100%;
                max-width: 460px;
                box-shadow: 0 4px 15px rgba(0,0,0,0.25);
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            ">
                <p style="color: #808495; font-size: 26px; font-weight: bold; letter-spacing:.5px; margin: 0; text-transform: uppercase;">TRITURADO</p>
                
            
            <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 12px; margin-bottom: 8px;">
                <span style="color:#ffffff; font-size: 70px; font-weight: 500; margin: 5px 0 2px 0; line-height: 1.1;">{real_formateado}</span>
                    <span style="
                        color: {color_tema}; 
                        background-color: {fondo_porcentaje}; 
                        border: 1px solid {color_tema}50;
                        padding: 6px 14px; 
                        border-radius: 20px; 
                        font-size: 15px; 
                        font-weight: bold;
                    ">{porcentaje_formateado}</span>
                </div>
                
            <p style="color: #636674; font-size: 26 px; margin: 0; line-height: 1.4;">
                    Obj. min: <span style="color: #808495;">{min_formateado}</span> &middot; Obj. max: <span style="color: #808495;">{max_formateado}</span>
                </p>
                
            <div style="background-color: #262730; width: 100%; border-radius: 6px; height: 8px; margin-top: 18px; margin-bottom: 4px;">
                    <div style="background-color: {color_tema}; width: {ancho_barra}%; border-radius: 6px; height: 100%;"></div>
                </div>
            </div>
            """
        st.markdown(html_tarjeta_inyectora, unsafe_allow_html=True)