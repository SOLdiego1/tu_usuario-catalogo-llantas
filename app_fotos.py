"""
app_fotos.py
Dashboard Inteligente de Inventarios y Precios con FOTOS POR DISEÑO (3 Fotos) o SFR.
Llantas y Neumáticos - Sincronizado con OneDrive
"""

import os
import textwrap
import streamlit as st
import pandas as pd

from data_loader import (
    get_latest_excel_file,
    get_file_sync_metadata,
    parse_excel_inventory,
    format_cop,
    clean_code
)
from search_engine import search_products, find_supplier_info
from photo_engine import PhotoManager

# =========================================================================
# CONFIGURACIÓN GENERAL DE LA PÁGINA
# =========================================================================
st.set_page_config(
    page_title="Inventario y Fotos por Diseño | Llantas & Neumáticos",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded"
)

def render_html(html_str: str):
    """
    Renderiza HTML sin sangrías para garantizar que la especificación CommonMark
    o el analizador Markdown de Streamlit nunca lo interprete erróneamente
    como un bloque de código (<pre><code>).
    """
    cleaned = "\n".join(line.strip() for line in html_str.strip().splitlines())
    st.markdown(cleaned, unsafe_allow_html=True)

# =========================================================================
# ESTILOS CSS - TEMA AZUL COMERCIAL
# =========================================================================
render_html("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    }
    
    .hero-banner {
        background: linear-gradient(135deg, #0A2540 0%, #0045A5 50%, #0070F3 100%);
        border-radius: 16px;
        padding: 22px 28px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(0, 112, 243, 0.25);
    }
    .hero-title {
        font-size: 26px;
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.5px;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .hero-subtitle {
        font-size: 14px;
        color: #BAE6FD;
        margin-top: 5px;
        font-weight: 500;
    }
    
    .sync-badge {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 5px 12px;
        border-radius: 9999px;
        font-size: 12px;
        font-weight: 600;
        margin-top: 10px;
    }
    .sync-badge.success {
        background: rgba(16, 185, 129, 0.2);
        color: #A7F3D0;
        border: 1px solid rgba(16, 185, 129, 0.4);
    }
    .sync-badge.warning {
        background: rgba(245, 158, 11, 0.2);
        color: #FDE68A;
        border: 1px solid rgba(245, 158, 11, 0.4);
    }
    
    /* Caja SFR. */
    .sfr-box {
        background: #FFF1F2;
        border: 2px dashed #F43F5E;
        border-radius: 14px;
        padding: 26px;
        text-align: center;
        margin: 8px 0;
    }
    .sfr-title {
        font-size: 38px;
        font-weight: 900;
        color: #E11D48;
        letter-spacing: 3px;
        line-height: 1;
    }
    .sfr-sub {
        font-size: 14px;
        font-weight: 700;
        color: #9F1239;
        margin-top: 6px;
    }
    
    /* Precios destacados */
    .price-card {
        background: #F8FAFC;
        border-radius: 12px;
        padding: 14px;
        border: 1px solid #E2E8F0;
        text-align: center;
    }
    .price-card.highlight {
        background: linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%);
        border: 2px solid #3B82F6;
    }
    .price-card.offer-active {
        background: linear-gradient(135deg, #FEF3C7 0%, #FDE68A 100%);
        border: 2px solid #F59E0B;
        box-shadow: 0 4px 12px rgba(245, 158, 11, 0.2);
    }
    .price-card-label {
        font-size: 11.5px;
        font-weight: 700;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .price-card-value {
        font-size: 22px;
        font-weight: 800;
        color: #0F172A;
        margin-top: 3px;
    }
    .price-card-value.offer {
        color: #B45309;
    }
    .price-card-sub {
        font-size: 11.5px;
        color: #64748B;
        margin-top: 2px;
    }
    
    /* Pastillas y badges */
    .pill-tab {
        background: #0284C7;
        color: white;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 700;
    }
    .pill-cg1 {
        background: #1E293B;
        color: #38BDF8;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 800;
        font-family: monospace;
    }
    
    /* Sedes */
    .sede-badge {
        display: inline-flex;
        justify-content: space-between;
        align-items: center;
        background: #F1F5F9;
        border: 1px solid #CBD5E1;
        padding: 7px 12px;
        border-radius: 8px;
        margin: 4px;
        min-width: 135px;
    }
    .sede-badge-name {
        font-size: 12.5px;
        font-weight: 600;
        color: #1E293B;
    }
    .sede-badge-qty {
        font-size: 13px;
        font-weight: 800;
        color: #0284C7;
        background: #E0F2FE;
        padding: 2px 7px;
        border-radius: 5px;
    }
</style>
""")

# =========================================================================
# GESTIÓN DE ESTADO Y CARGA
# =========================================================================
if 'search_query' not in st.session_state:
    st.session_state.search_query = ""
if 'selected_product_index' not in st.session_state:
    st.session_state.selected_product_index = 0
if 'custom_excel_path' not in st.session_state:
    st.session_state.custom_excel_path = None

@st.cache_resource
def get_photo_manager():
    return PhotoManager()

@st.cache_data(ttl=3600, show_spinner="Cargando inventario y listas de precios...")
def load_inventory_data(file_path):
    return parse_excel_inventory(file_path)

photo_mgr = get_photo_manager()

default_file = get_latest_excel_file()
active_file = st.session_state.custom_excel_path or default_file

# =========================================================================
# BARRA LATERAL
# =========================================================================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2965/2965300.png", width=48)
    st.title("Inventario & Fotos")
    st.caption("Control de Datos y Galería por Diseño")

    uploaded_file = st.file_uploader(
        "Cargar nuevo archivo Excel:",
        type=["xlsx", "xlsm"],
        help="Sube una nueva versión del inventario para actualizar en vivo."
    )
    if uploaded_file is not None:
        temp_dir = os.path.join(os.path.dirname(__file__), "temp")
        os.makedirs(temp_dir, exist_ok=True)
        temp_path = os.path.join(temp_dir, uploaded_file.name)
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        active_file = temp_path
        st.session_state.custom_excel_path = temp_path
        st.success(f"Archivo cargado: {uploaded_file.name}")

    sync_meta = get_file_sync_metadata(active_file)

    st.markdown("---")
    st.markdown("### 📊 Estado de OneDrive")
    st.markdown(f"**Archivo:** `{sync_meta['file_name']}`")
    st.markdown(f"**Modificado:** {sync_meta['last_modified_str']}")

    if sync_meta['status'] == 'OUTDATED':
        st.warning(sync_meta['status_label'])
    else:
        st.success(sync_meta['status_label'])

    if st.button("🔄 Recargar Datos y Fotos", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    st.markdown("### 📸 Base Fotográfica")
    all_designs = photo_mgr.get_all_designs()
    st.metric("Diseños con Fotos", f"{len(all_designs):,} diseños")
    st.caption("Cada diseño cuenta con exactamente 3 fotos registradas.")

    pdf_report_path = os.path.join(os.path.dirname(__file__), "INFORME_AUDITORIA_INVENTARIO_Y_FOTOS.pdf")
    pdf_guide_path = os.path.join(os.path.dirname(__file__), "GUIA_DESPLIEGUE_NUBE_24_7.pdf")
    if os.path.exists(pdf_report_path) or os.path.exists(pdf_guide_path):
        st.markdown("---")
        st.markdown("### 📄 Documentos y Guías PDF")
        if os.path.exists(pdf_report_path):
            with open(pdf_report_path, "rb") as f_pdf:
                st.download_button(
                    label="📊 Descargar Informe de Auditoría (PDF)",
                    data=f_pdf.read(),
                    file_name="INFORME_AUDITORIA_INVENTARIO_Y_FOTOS.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
        if os.path.exists(pdf_guide_path):
            with open(pdf_guide_path, "rb") as f_guide:
                st.download_button(
                    label="☁️ Descargar Guía Nube 24/7 (PDF)",
                    data=f_guide.read(),
                    file_name="GUIA_DESPLIEGUE_NUBE_24_7.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

    st.markdown("---")
    sheet_filter = st.selectbox(
        "Filtrar por Pestaña:",
        [
            "Todas las pestañas",
            "LLANTAS",
            "NEU ESPEC",
            "NEU COMERC",
            "LP GY AUTO-CAM",
            "LP CONTI CAMION",
            "LP GY CAMION",
            "LP CONTI AUTO-CAM",
            "E-S",
            "OFERTA",
            "OBSOL"
        ]
    )
    only_in_stock = False

# =========================================================================
# CARGA DE PRODUCTOS
# =========================================================================
if not active_file or not os.path.exists(active_file):
    st.error("No se encontró el archivo de inventario. Por favor sube un archivo desde el panel lateral.")
    st.stop()

try:
    all_products, supplier_inventory = load_inventory_data(active_file)
except Exception as e:
    st.error(f"Error al leer el archivo de inventario: {e}")
    st.stop()

# =========================================================================
# BANNER HERO PRINCIPAL
# =========================================================================
badge_class = "success" if sync_meta['status'] == 'UP_TO_DATE' else "warning"

render_html(f"""
<div class="hero-banner">
    <div class="hero-title">
        <span>🚗</span> Buscador de Catálogo y Fotos por Diseño
    </div>
    <div class="hero-subtitle">
        Visualización de 3 fotos oficiales por diseño (o rótulo SFR. cuando no exista foto) y ficha técnica por producto.
    </div>
    <div class="sync-badge {badge_class}">
        <span>OneDrive:</span> {sync_meta['status_label']} | <b>{len(all_products):,}</b> productos | <b>{len(all_designs)}</b> diseños con fotos
    </div>
</div>
""")

# =========================================================================
# PESTAÑAS: 1. BUSCADOR | 2. CATÁLOGO DE DISEÑOS
# =========================================================================
tab_search, tab_gallery = st.tabs([
    "🔍 Buscador de Productos y Fotos por Diseño (3 Fotos / SFR.)",
    f"🖼️ Catálogo y Galería de Fotos por Diseño ({len(all_designs)} Diseños)"
])

# =========================================================================
# TAB 1: BUSCADOR
# =========================================================================
with tab_search:
    col_search, col_btn_find, col_btn_clear, col_btn_sample = st.columns([5, 1.2, 1, 1.3])

    with col_search:
        search_input = st.text_input(
            "Buscar por Código CG1, Referencia o Código de Fabricante:",
            value=st.session_state.search_query,
            placeholder="Ej: 1600, 500-12, 600-12, 7.50-16, 2467, 04121760000...",
            label_visibility="collapsed"
        )

    with col_btn_find:
        btn_search = st.button("🔍 BUSCAR", type="primary", use_container_width=True)

    with col_btn_clear:
        btn_clear = st.button("🧹 LIMPIAR", use_container_width=True)

    with col_btn_sample:
        btn_sample = st.button("🎲 Ejemplo", use_container_width=True)

    if btn_clear:
        st.session_state.search_query = ""
        st.session_state.selected_product_index = 0
        st.rerun()

    if btn_sample:
        import random
        samples = ["600-12", "500-12", "7.50-16", "1100-16", "1600", "2467", "04121760000", "5985"]
        st.session_state.search_query = random.choice(samples)
        st.session_state.selected_product_index = 0
        st.rerun()

    if btn_search and search_input:
        st.session_state.search_query = search_input
        st.session_state.selected_product_index = 0
        st.rerun()

    if search_input != st.session_state.search_query:
        st.session_state.search_query = search_input
        st.session_state.selected_product_index = 0
        st.rerun()

    current_query = st.session_state.search_query.strip()

    if not current_query:
        with st.container(border=True):
            st.markdown("### 💡 Consulta Comercial con Fotos por Diseño")
            st.write(
                "Ingresa el **Código CG1** o la **Referencia comercial** de la llanta. "
                "El sistema consultará precios, sedes con existencias (>0), mercancía en tránsito y "
                "**las 3 fotos correspondientes al diseño** suministradas en la carpeta."
            )
            st.caption(
                "Si el producto tiene fotos asignadas a su diseño, se mostrarán las **3 fotos oficiales** organizadas en columnas. "
                "Si no tiene fotos registradas, se indicará claramente el rótulo **SFR.**"
            )

        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric("Total Productos Registrados", f"{len(all_products):,}")
        with m2:
            st.metric("Diseños con Fotos Registradas", f"{len(all_designs):,}")
        with m3:
            st.metric("Sincronización OneDrive", sync_meta['status_label'])

    else:
        results = search_products(all_products, current_query, selected_sheet=sheet_filter, only_in_stock=only_in_stock)

        if not results:
            st.error(f"❌ No se encontraron productos para: '{current_query}'")
        else:
            num_res = len(results)
            if num_res > 1:
                # Pre-calcular estado de fotos para cada resultado
                for r in results:
                    r['_photo_data'] = photo_mgr.get_photos_for_product(r['product'])
                    r['_has_photo'] = r['_photo_data'] is not None

                total_con_foto = sum(1 for r in results if r['_has_photo'])
                total_sin_foto = num_res - total_con_foto

                c_m1, c_m2 = st.columns([3, 2])
                with c_m1:
                    st.markdown(
                        f"#### 🔎 Se encontraron **{num_res} coincidencias** para `{current_query}`: "
                        f"(<span style='color: #059669; font-weight: 800;'>🟢 {total_con_foto} con foto</span> | "
                        f"<span style='color: #475569;'>⚫ {total_sin_foto} sin foto</span>)",
                        unsafe_allow_html=True
                    )
                with c_m2:
                    if total_con_foto > 0:
                        solo_con_foto = st.checkbox("🟢 Ver únicamente referencias con foto", value=False, key="chk_solo_con_foto")
                    else:
                        solo_con_foto = False

                if solo_con_foto:
                    active_results = [r for r in results if r['_has_photo']]
                else:
                    # Ordenar para que las referencias que tienen foto aparezcan de primeras
                    active_results = sorted(results, key=lambda r: 0 if r['_has_photo'] else 1)

                options = []
                photo_indices = []
                for idx, r in enumerate(active_results):
                    p = r['product']
                    cg1_lbl = f"CG1: {p['cod_cg1']}" if p['cod_cg1'] else f"Mat: {p['fabricante_code']}"
                    diseno_lbl = f" - Diseño: {p['diseno']}" if p['diseno'] and p['diseno'] != 'Sin información' else ""
                    if r['_has_photo']:
                        photo_indices.append(idx)
                        options.append(f"🟢 [CON FOTO] [{p['source_sheet']}] {cg1_lbl} - {p['referencia']} ({p['marca']}){diseno_lbl}")
                    else:
                        options.append(f"⚫ [SFR.] [{p['source_sheet']}] {cg1_lbl} - {p['referencia']} ({p['marca']}){diseno_lbl}")

                # Inyectar CSS dinámico para resaltar las opciones con foto en verde en el menú desplegable
                if photo_indices:
                    selectors = ", ".join([f'div[data-baseweb="popover"] ul[role="listbox"] > li:nth-child({i + 1})' for i in photo_indices])
                    hover_selectors = ", ".join([f'div[data-baseweb="popover"] ul[role="listbox"] > li:nth-child({i + 1}):hover' for i in photo_indices])
                    render_html(f"""
                    <style>
                        {selectors} {{
                            background-color: #ECFDF5 !important;
                            color: #065F46 !important;
                            font-weight: 700 !important;
                            border-left: 5px solid #10B981 !important;
                        }}
                        {hover_selectors} {{
                            background-color: #D1FAE5 !important;
                            color: #047857 !important;
                        }}
                    </style>
                    """)

                sel_idx = st.selectbox(
                    "Selecciona el ítem exacto que deseas consultar (🟢 en verde con foto, ⚫ en negro sin foto):",
                    range(len(active_results)),
                    format_func=lambda i: options[i],
                    index=min(st.session_state.selected_product_index, len(active_results) - 1)
                )
                st.session_state.selected_product_index = sel_idx
                target = active_results[sel_idx]
            else:
                target = results[0]
                target['_photo_data'] = photo_mgr.get_photos_for_product(target['product'])
                target['_has_photo'] = target['_photo_data'] is not None

            prod = target['product']
            match_type = target['match_type']

            photo_data = target.get('_photo_data') or photo_mgr.get_photos_for_product(prod)

            # =================================================================
            # 1. FICHA PRINCIPAL DEL PRODUCTO (USANDO CONTAINER NATIVO)
            # =================================================================
            with st.container(border=True):
                # Fila de badges
                b_col1, b_col2 = st.columns([4, 2])
                with b_col1:
                    render_html(f"""
                    <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
                        <span class="pill-tab">{prod['source_sheet']}</span>
                        <span class="pill-cg1">CG1: {prod['cod_cg1'] or 'No registrado'}</span>
                        {f'<span class="pill-cg1">Fabricante: {prod["fabricante_code"]}</span>' if prod['fabricante_code'] else ''}
                        {'<span style="background: #E0F2FE; color: #0369A1; padding: 4px 10px; border-radius: 6px; font-weight: 800; font-size: 12px;">📸 3 Fotos Disponibles</span>' if photo_data else '<span style="background: #FFE4E6; color: #BE123C; padding: 4px 10px; border-radius: 6px; font-weight: 900; font-size: 12px;">SFR.</span>'}
                    </div>
                    """)
                with b_col2:
                    st.markdown(f"<div style='text-align: right; font-size: 13px; font-weight: 700; color: #0070F3;'>Coincidencia: {match_type}</div>", unsafe_allow_html=True)

                st.markdown(f"<h2 style='margin: 8px 0 2px 0; font-size: 26px; font-weight: 800; color: #0F172A;'>{prod['referencia']}</h2>", unsafe_allow_html=True)
                st.markdown(f"<p style='color: #64748B; font-size: 14px; margin: 0 0 14px 0;'>{prod['descripcion']}</p>", unsafe_allow_html=True)

                # Grilla de especificaciones
                c_sp1, c_sp2, c_sp3, c_sp4, c_sp5, c_sp6 = st.columns(6)
                with c_sp1:
                    st.caption("MARCA")
                    st.markdown(f"**{prod['marca']}**")
                with c_sp2:
                    st.caption("DISEÑO")
                    st.markdown(f"**{prod['diseno']}**")
                with c_sp3:
                    st.caption("LÍNEA / TIPO")
                    st.markdown(f"**{prod['tipo']}**")
                with c_sp4:
                    st.caption("VÁLVULA / TT-TL")
                    st.markdown(f"**{prod['tipo_valvula']}**")
                with c_sp5:
                    st.caption("LONAS / PR")
                    st.markdown(f"**{prod['pr']}**")
                with c_sp6:
                    st.caption("APLICA IVA")
                    st.markdown(f"**{prod['aplica_iva']}**")

            # =================================================================
            # 2. RECUADRO DE FOTOS POR DISEÑO (3 FOTOS) O SFR.
            # =================================================================
            with st.container(border=True):
                st.markdown("### 📸 FOTOS OFICIALES POR DISEÑO (3 VISTAS DEL PRODUCTO)")

                if photo_data and photo_data.get('photos'):
                    photos = photo_data['photos']
                    folder_title = photo_data.get('folder_name', 'Diseño')
                    cat_title = photo_data.get('category', 'Categoría')

                    st.markdown(f"🏷️ **Diseño identificado:** `{folder_title}` | 📂 **Categoría:** `{cat_title}` | **{len(photos)} fotos oficiales**")

                    cols_f = st.columns(3)
                    labels = ["Foto 1: Vista Frontal / Labrado", "Foto 2: Banda de Rodamiento", "Foto 3: Perfil / Costado"]

                    for i, c_f in enumerate(cols_f):
                        with c_f:
                            if i < len(photos) and os.path.exists(photos[i]):
                                st.markdown(f"<div style='font-size: 12px; font-weight: 700; color: #0284C7; text-align: center; margin-bottom: 6px;'>{labels[i]}</div>", unsafe_allow_html=True)
                                st.image(photos[i], use_container_width=True)
                                
                                with open(photos[i], "rb") as file_bytes:
                                    fname = os.path.basename(photos[i])
                                    st.download_button(
                                        label=f"📥 Descargar Foto {i+1}",
                                        data=file_bytes,
                                        file_name=f"{prod['cod_cg1']}_{prod['referencia']}_Foto_{i+1}_{fname}",
                                        mime="image/jpeg",
                                        key=f"dl_photo_{i}_{prod['cod_cg1']}",
                                        use_container_width=True
                                    )
                            else:
                                render_html("""
                                <div class="sfr-box" style="padding: 16px;">
                                    <div style="font-size: 22px; font-weight: 900; color: #E11D48;">SFR.</div>
                                    <div style="font-size: 11px; color: #9F1239;">Vista no disponible</div>
                                </div>
                                """)

                else:
                    # RÓTULO DESTACADO EXACTAMENTE SFR.
                    render_html("""
                    <div class="sfr-box">
                        <div class="sfr-title">SFR.</div>
                        <div class="sfr-sub">Sin Foto Registrada para este diseño en el catálogo oficial</div>
                        <p style="color: #9F1239; font-size: 12.5px; margin: 6px 0 0 0;">
                            No se encontraron fotografías vinculadas en la carpeta de imágenes para esta marca/diseño.
                        </p>
                    </div>
                    """)

            # Fin de la ficha y fotos del producto

# =========================================================================
# TAB 2: CATÁLOGO DE FOTOS POR DISEÑO
# =========================================================================
with tab_gallery:
    with st.container(border=True):
        st.markdown("### 🖼️ Catálogo Oficial de Diseños (284 Diseños con 3 Fotos)")
        st.write("Explora libremente todas las carpetas fotográficas de diseño por categoría y marca.")

    c_g1, c_g2, c_g3 = st.columns([2, 2, 3])
    with c_g1:
        categories = ["Todas"] + sorted(list(set(d['category'] for d in all_designs)))
        sel_cat = st.selectbox("Filtrar por Categoría:", categories)
    with c_g2:
        brands = ["Todas"] + sorted(list(set(d['brand'] for d in all_designs if d['brand'])))
        sel_brand = st.selectbox("Filtrar por Marca:", brands)
    with c_g3:
        search_design = st.text_input("Buscar por nombre de diseño o medida:", placeholder="Ej: R1, 6.00-12, QH651, GL073A...")

    filtered_designs = all_designs
    if sel_cat != "Todas":
        filtered_designs = [d for d in filtered_designs if d['category'] == sel_cat]
    if sel_brand != "Todas":
        filtered_designs = [d for d in filtered_designs if d['brand'] == sel_brand]
    if search_design:
        q_d = search_design.lower().strip()
        filtered_designs = [d for d in filtered_designs if q_d in d['folder_name'].lower()]

    st.markdown(f"**Mostrando {len(filtered_designs)} diseños encontrados:**")

    page_size = 6
    total_pages = max(1, (len(filtered_designs) + page_size - 1) // page_size)
    page = st.number_input(f"Página (1 a {total_pages}):", min_value=1, max_value=total_pages, value=1, step=1)

    start_idx = (page - 1) * page_size
    page_designs = filtered_designs[start_idx : start_idx + page_size]

    for d in page_designs:
        with st.expander(f"📁 {d['category']} | {d['folder_name']} ({len(d['photos'])} fotos)", expanded=True):
            cols_d = st.columns(3)
            for j in range(3):
                with cols_d[j]:
                    if j < len(d['photos']) and os.path.exists(d['photos'][j]):
                        st.image(d['photos'][j], use_container_width=True, caption=f"Foto {j+1}")
                    else:
                        st.markdown("<div class='sfr-box' style='padding: 10px;'><div style='font-size:18px;font-weight:800;color:#E11D48;'>SFR.</div></div>", unsafe_allow_html=True)

# =========================================================================
# PIE DE PÁGINA
# =========================================================================
st.markdown("---")
render_html("""
<div style="text-align: center; color: #64748B; font-size: 12px; padding: 10px 0;">
    🚗 <b>Buscador de Inventarios con Fotos por Diseño</b> | Desarrollado para consulta comercial rápida | Sincronizado con OneDrive
</div>
""")
