import io
import base64
from html import escape
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Dashboard Pembinaan Jabatan Fungsional | BKN",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="auto",
)
BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = (
    BASE_DIR
    / "01102026_Konsep Dahsboard JF.xlsx"
)
LOGO_FILE = (
    BASE_DIR
    / "assets"
    / "logo-bkn.png"
)
NAVY, BLUE, SKY = '#15324F', '#1766A3', '#EAF3FA'
BG, WHITE, TEXT, MUTED, BORDER = '#F5F8FC', '#FFFFFF', '#203449', '#607489', '#E0E8F0'
GREEN, AMBER = '#21846B', '#D38B35'

st.markdown(
    f"""
    <style>
:root {{
    color-scheme:light;
}}
.stApp {{
    background:{BG};
     color:{TEXT};
}}
.block-container {{
    max-width:1500px;
     padding:1.5rem clamp(.85rem,2.6vw,2.8rem) 3rem;
}}
[data-testid="stSidebar"] {{
    background:{WHITE};
    border-right:1px solid {BORDER};
}}
[data-testid="stSidebarContent"] {{
    padding:1.1rem .9rem 1.8rem;
}}
[data-testid="stSidebar"] .stButton button {{
    min-height:45px;
    text-align:left;
    justify-content:flex-start;
    border-radius:11px;
    font-weight:650;
    border:1px solid {BORDER};
    transition:background .15s;
}}
[data-testid="stSidebar"] .stButton button[kind="primary"] {{
    background:{NAVY};
    border-color:{NAVY};
    color:white;
}}
[data-testid="stSidebar"] .stButton button[kind="secondary"] {{
    background:white;
    color:{NAVY};
}}
[data-testid="stSidebar"] .stButton button[kind="secondary"]:hover {{
    background:{SKY};
    border-color:#B8D2E6;
}}
.brand {{
    text-align:center;
    padding:8px 4px 16px;
}}
.sidebar-logo {{
    display: flex;
    justify-content: center;
    align-items: center;
    width: 100%;
    padding: 12px 0 4px;

}}
.sidebar-logo img {{
    display: block;
    width: 100px;
    max-width: 42%;
    height: auto;
    margin: 0 auto;
    object-fit: contain;

}}
.brand-name {{
    font-size:1.03rem;
    font-weight:800;
    color:{NAVY};
    margin-top:10px;
}}
.brand-sub {{
    font-size:.8rem;
    color:{MUTED};
    line-height:1.5;
    margin-top:4px;
}}
.eyebrow {{
    font-size:.75rem;
    letter-spacing:.09em;
    text-transform:uppercase;
    font-weight:750;
    color:{MUTED};
}}
.hero {{
    background:linear-gradient(115deg,{NAVY},#1D6597);
    border-radius:18px;
    padding:clamp(22px,3.5vw,40px);
    color:white;
    margin-bottom:22px;
    position:relative;
    overflow:hidden;
}}
.hero:after {{
    content:'';
    position:absolute;
    width:250px;
    height:250px;
    right:-100px;
    top:-120px;
    border:1px solid #ffffff30;
    border-radius:50%;
    box-shadow:0 0 0 55px #ffffff08;
    pointer-events:none;
}}
.hero .eyebrow {{
    color:#C4DCEE;
}}
.hero h1 {{
    font-size:clamp(1.45rem,2.7vw,2.4rem);
    line-height:1.2;
    letter-spacing:-.025em;
    margin:.6rem 0;
    font-weight:800;
    max-width:850px;
}}
.hero p {{
    font-size:clamp(.86rem,1.2vw,1rem);
    color:#D9E9F5;
    line-height:1.6;
    margin:0;
    max-width:750px;
}}
.section-head {{
    margin:2px 0 17px;
}}
.section-head h2 {{
    font-size:clamp(1.1rem,1.7vw,1.4rem);
    font-weight:800;
    color:{NAVY};
    margin:0 0 4px;
}}
.section-head p {{
    font-size:.87rem;
    color:{MUTED};
    margin:0;
    line-height:1.6;
}}
.subhead {{
    font-weight:750;
    font-size:1rem;
    color:{NAVY};
    margin:16px 0 12px;
}}
.kpi {{
    background:white;
    border:1px solid {BORDER};
    border-radius:14px;
    padding:clamp(15px,2vw,22px);
    min-height:132px;
    height:100%;
    box-shadow:0 3px 13px #193B5908;
}}
.kpi-label {{
    font-size:.81rem;
    font-weight:700;
    color:{MUTED};
    line-height:1.4;
}}
.kpi-value {{
    font-size:clamp(1.55rem,2.3vw,2.15rem);
    font-weight:800;
    color:{NAVY};
    line-height:1.3;
    margin:8px 0 5px;
    overflow-wrap:anywhere;
}}
.kpi-note {{
    font-size:.76rem;
    color:{MUTED};
    line-height:1.45;
}}
.info-card {{
    background:{SKY};
    border:1px solid #D5E6F3;
    border-radius:12px;
    padding:16px;
    min-height:112px;
    height:100%;
}}
.info-card .kpi-value {{
    font-size:clamp(.92rem,1.4vw,1.13rem);
    line-height:1.4;
}}
[data-testid="stVerticalBlockBorderWrapper"] {{
    background:{WHITE};
    border-color:{BORDER};
    border-radius:16px;
    box-shadow:0 4px 16px #193B5908;
}}
[data-testid="stVerticalBlockBorderWrapper"] > div {{
    min-width:0;
}}
[data-testid="stHorizontalBlock"] {{
    align-items:stretch;
}}
[data-testid="column"] {{
    min-width:0;
}}
[data-testid="stPlotlyChart"] {{
    width:100%;
    min-width:0;
}}
[data-testid="stDataFrame"] {{
    border:1px solid {BORDER};
    border-radius:10px;
    overflow:hidden;
}}
.stButton button,.stDownloadButton button {{
    border-radius:10px;
    min-height:42px;
}}
.stDownloadButton button {{
    background:{BLUE};
    color:white;
    border-color:{BLUE};
    font-weight:700;
}}
.stDownloadButton button:hover {{
    background:{NAVY};
    color:white;
    border-color:{NAVY};
}}
[data-testid="stWidgetLabel"] p {{
    font-weight:650;
    color:{TEXT};
    font-size:.86rem;
}}
.stTextInput input,.stSelectbox [data-baseweb="select"] {{
    border-radius:9px;
}}
.footer {{
    border-top:1px solid {BORDER};
    padding-top:18px;
    margin-top:32px;
    text-align:center;
    color:{MUTED};
    font-size:.78rem;
}}
@media(max-width:1100px) {{
    .block-container {{
    padding:1.1rem 1rem 2.5rem;
}}
}}
@media(max-width:768px) {{
 .block-container {{
    padding:1rem .75rem 2rem;
}}
 .hero {{
    border-radius:14px;
    margin-bottom:16px;
    padding:22px 18px;
}}
 .hero:after {{
    width:150px;
    height:150px;
    right:-90px;
    top:-65px;
}}
 .kpi {{
    min-height:115px;
    padding:15px;
}}
 .info-card {{
    min-height:96px;
}}
 [data-testid="stHorizontalBlock"] {{
    flex-wrap:wrap;
}}
 [data-testid="column"] {{
    min-width:min(100%,240px);
    flex:1 1 240px;
}}
 [data-testid="stSidebar"] {{
    max-width:min(85vw,330px);
}}
}}
@media(max-width:480px) {{
 .block-container {{
    padding:.75rem .6rem 1.5rem;
}}
 .hero {{
    padding:20px 15px;
}}
 [data-testid="column"] {{
    min-width:100%;
    flex-basis:100%;
}}
 .kpi {{
    min-height:100px;
}}
 .section-head {{
    margin-bottom:12px;
}}
}}
/* Consistent card grid independent of Streamlit column stacking. */
.kpi-grid {{
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 14px;
    margin: 4px 0 18px;

}}
.kpi-grid > article {{
    min-width: 0;
    min-height: 125px;
    overflow-wrap: anywhere;

}}
/* Sidebar logo alignment */
[data-testid="stSidebar"] [data-testid="stImage"] {{
    display: flex;
    justify-content: center;
    width: 100%;

}}
[data-testid="stSidebar"] [data-testid="stImage"] img {{
    margin-inline: auto;

}}
@media (max-width: 1000px) {{
    .kpi-grid {{
    grid-template-columns: repeat(2, minmax(0, 1fr));
}}
}}
@media (max-width: 768px) {{

    /* Leave room for Streamlit's fixed top toolbar on mobile. */
    .block-container {{
    padding-top: 5rem !important;
}}
    .kpi-grid {{
    gap: 10px;
}}
    .kpi-grid > article {{
    min-height: 116px;
     padding: 14px 12px;
}}
    .kpi-grid .kpi-label {{
    font-size: .75rem;
}}
    .kpi-grid .kpi-value {{
    font-size: clamp(1.25rem, 5vw, 1.65rem);
}}
    .kpi-grid .kpi-note {{
    font-size: .7rem;
}}
}}
@media (max-width: 390px) {{
    .kpi-grid {{
    gap: 8px;
}}
    .kpi-grid > article {{
    padding: 12px 10px;
}}
}}
@media(prefers-reduced-motion:reduce) {{
    *,*:before,*:after {{
    animation:none!important;
    transition:none!important;
}}
}}
</style>
    """,
    unsafe_allow_html=True,
)


def fmt(value):
    try:
        return f'{float(value):,.0f}'.replace(',', '.')
    except (ValueError, TypeError):
        return '0'


def section(title, description=''):
    st.html(
        '<div class="section-head">'
        f'<h2>{escape(title)}</h2>'
        f'<p>{escape(description)}</p>'
        '</div>'
    )


def subhead(title):
    st.html(
        f'<div class="subhead">{escape(title)}</div>'
    )


def cards(items, info=False):
    """Responsive card grid: four desktop, two mobile columns."""
    card_class = "info-card" if info else "kpi"
    markup = []

    for label, value, note in items:
        markup.append(
            f'<article class="{card_class}">'
            f'<div class="kpi-label">{escape(str(label))}</div>'
            f'<div class="kpi-value">{escape(str(value))}</div>'
            f'<div class="kpi-note">{escape(str(note))}</div>'
            '</article>'
        )

    st.html(
        '<div class="kpi-grid">'
        + "".join(markup)
        + "</div>"
    )


def normalize_uji(series):
    return series.fillna('').astype(str).str.strip().str.casefold().isin({'✓','✔','v','ya','sudah','yes','true'})


@st.cache_data(show_spinner=False)
def load_excel(content):
    xls = io.BytesIO(content)
    status = pd.read_excel(xls, sheet_name='DB_Status Pembinaan', header=None).iloc[2:, 1:5].copy()
    status.columns = ['Instansi Pembina','Status Pembinaan','Jumlah Pembinaan','Rencana Pembinaan']
    status = status[status['Instansi Pembina'].notna()].copy()
    for c in ['Instansi Pembina','Status Pembinaan','Rencana Pembinaan']:
        status[c] = status[c].fillna('').astype(str).str.strip()
    status['Jumlah Pembinaan'] = pd.to_numeric(status['Jumlah Pembinaan'],errors='coerce').fillna(0)
    info = pd.read_excel(xls, sheet_name='DB_Informasi Pembina', header=None).iloc[3:, :13].copy()
    info.columns = ['No','Instansi Pembina','Nomenklatur','Jumlah Pemangku','Keterampilan Terendah','Keterampilan Tertinggi','Keahlian Terendah','Keahlian Tertinggi','Status Uji Kompetensi','Frekuensi/Tahun','Permenpan Lama','Permenpan Baru','Link']
    info = info[info['Instansi Pembina'].notna()].copy()
    for c in info.select_dtypes(include='object').columns:
        info[c] = info[c].fillna('').astype(str).str.strip()
    for c in ['No','Jumlah Pemangku','Frekuensi/Tahun']:
        info[c] = pd.to_numeric(info[c],errors='coerce')
    return status, info


def plot_style(fig, height=310):
    fig.update_layout(height=height, margin=dict(l=8,r=8,t=20,b=30), paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',font=dict(color=TEXT,size=12), showlegend=False, hoverlabel=dict(bgcolor=WHITE,font_color=TEXT))
    fig.update_xaxes(showgrid=False,automargin=True)
    fig.update_yaxes(gridcolor=BORDER,zeroline=False,automargin=True)
    return fig


def bar_chart(labels, values, colors=None, horizontal=False):
    frame = pd.DataFrame({'Kategori':labels,'Jumlah':values})
    if horizontal:
        fig = px.bar(frame,x='Jumlah',y='Kategori',orientation='h',text='Jumlah')
    else:
        fig = px.bar(frame,x='Kategori',y='Jumlah',text='Jumlah')
    fig.update_traces(marker_color=colors or BLUE,textposition='outside',cliponaxis=False)
    fig = plot_style(fig)
    fig.update_layout(xaxis_title=None,yaxis_title=None)
    return fig


def donut(labels, values, colors):
    if sum(values) == 0:
        return bar_chart(labels,values,colors)
    fig = go.Figure(go.Pie(labels=labels,values=values,hole=.65,marker=dict(colors=colors,line=dict(color=WHITE,width=3)),textinfo='percent',sort=False))
    plot_style(fig)
    fig.update_layout(showlegend=True,legend=dict(orientation='h',y=-.13,x=.5,xanchor='center'))
    return fig


def chart(fig):
    st.plotly_chart(fig,use_container_width=True,config={'responsive':True,'displayModeBar':False})


def table(data, height=520):
    cfg = {
        'Jumlah Pemangku':st.column_config.NumberColumn('Jumlah Pemangku',format='%d'),
        'Jumlah Pembinaan':st.column_config.NumberColumn('Jumlah Pembinaan',format='%d'),
        'Frekuensi/Tahun':st.column_config.NumberColumn('Frekuensi/Tahun',format='%d'),
        'Link':st.column_config.LinkColumn('Peraturan',display_text='Buka'),
    }
    st.dataframe(data,use_container_width=True,hide_index=True,height=height,column_config={k:v for k,v in cfg.items() if k in data.columns})


def download(data, name):
    st.download_button('↓ Unduh data CSV',data=data.to_csv(index=False).encode('utf-8-sig'),file_name=name,mime='text/csv',use_container_width=True)


pages = ['Dashboard','Data Jabatan','Status Pembinaan']
if 'active_page' not in st.session_state or st.session_state.active_page not in pages:
    st.session_state.active_page = 'Dashboard'

with st.sidebar:
    if LOGO_FILE.exists():
        # Render gambar sebagai HTML agar posisi center
        # konsisten pada sidebar desktop dan mobile.
        logo_base64 = base64.b64encode(
            LOGO_FILE.read_bytes()
        ).decode("ascii")

        st.html(
            '<div class="sidebar-logo">'
            '<img '
            'src="data:image/png;base64,'
            + logo_base64
            + '" alt="Logo BKN">'
            '</div>'
        )
    st.html(
        '<div class="brand">'
        '<div class="brand-name">'
        'Badan Kepegawaian Negara'
        '</div>'
        '<div class="brand-sub">'
        'Sistem Pembinaan Jabatan Fungsional'
        '</div>'
        '</div>'
    )
    st.divider()
    st.caption('MENU UTAMA')
    for page, icon in zip(
        pages,
        ["▦", "▤", "◉"],
    ):
        if st.button(
            f"{icon}  {page}",
            key=f"nav_{page}",
            use_container_width=True,
            type=(
                "primary"
                if st.session_state.active_page == page
                else "secondary"
            ),
        ):
            st.session_state.active_page=page
            st.rerun()
    st.divider()
    st.caption('SUMBER DATA')
    uploaded = st.file_uploader('Unggah Excel',type=['xlsx','xls'],help='Gunakan struktur sheet yang sama dengan file sumber BKN.')
    st.caption('Format: XLSX / XLS')
    st.divider()
    st.caption('Dashboard Pembinaan JF • BKN')

page_title = (
    "Dashboard Pembinaan Jabatan Fungsional"
    if st.session_state.active_page == "Dashboard"
    else st.session_state.active_page
)

st.html(
    '<div class="hero">'
    '<div class="eyebrow">BADAN KEPEGAWAIAN NEGARA</div>'
    f'<h1>{escape(page_title)}</h1>'
    '<p>Informasi, pemetaan, dan monitoring pembinaan '
    'jabatan fungsional dalam satu tampilan yang mudah '
    'diakses.</p>'
    '</div>'
)

if uploaded is not None:
    source_bytes = uploaded.getvalue()
    source_label = uploaded.name
elif EXCEL_FILE.exists():
    source_bytes = EXCEL_FILE.read_bytes()
    source_label = EXCEL_FILE.name
else:
    st.warning(f'File bawaan **{EXCEL_FILE.name}** belum tersedia. Unggah file Excel melalui sidebar untuk mulai menggunakan dashboard.')
    st.stop()
try:
    df_status,df_info = load_excel(source_bytes)
except Exception as exc:
    st.error('File Excel tidak dapat dibaca. Periksa nama sheet dan susunan kolom sesuai template asli.')
    st.exception(exc)
    st.stop()

st.caption(f'Sumber data: {source_label}  ·  {len(df_info):,} baris jabatan  ·  {len(df_status):,} baris pembinaan')

if st.session_state.active_page == 'Dashboard':
    total_instansi = df_status['Instansi Pembina'].replace('',pd.NA).nunique()
    total_jf = df_info['Nomenklatur'].replace('',pd.NA).nunique()
    total_pemangku = df_info['Jumlah Pemangku'].fillna(0).sum()
    normalized = df_status['Status Pembinaan'].str.casefold()
    sudah = int(normalized.eq('sudah').sum())
    belum = int(normalized.eq('belum').sum())
    persen = 100*sudah/total_instansi if total_instansi else 0
    with st.container(border=True):
        section('Gambaran Umum','Ringkasan data pembinaan dan jabatan fungsional.')
        cards([('Instansi Pembina',fmt(total_instansi),'Total instansi terdaftar'),('Jabatan Fungsional',fmt(total_jf),'Nomenklatur unik'),('Jumlah Pemangku',fmt(total_pemangku),'Total pemangku jabatan'),('Pembinaan Selesai',f'{persen:.1f}% ',f'{sudah} sudah · {belum} belum')])
        subhead('Analisis Pembinaan')
        a,b,c = st.columns(3,gap='medium')
        with a:
            st.markdown('**Jumlah Instansi Berdasarkan Status**')
            chart(bar_chart(['Sudah','Belum'],[sudah,belum],[GREEN,AMBER]))
        with b:
            st.markdown('**Proporsi Status Pembinaan**')
            chart(donut(['Sudah','Belum'],[sudah,belum],[GREEN,AMBER]))
        with c:
            st.markdown('**Instansi vs Jabatan Fungsional**')
            chart(bar_chart(['Instansi','Jabatan Fungsional'],[total_instansi,total_jf]))

    st.write('')
    with st.container(border=True):
        section('Pencarian Berdasarkan Instansi','Pilih instansi untuk melihat ringkasan, status, grafik, dan daftar jabatan terkait.')
        options = ['Semua'] + sorted(x for x in df_info['Instansi Pembina'].dropna().astype(str).str.strip().unique() if x)
        selected = st.selectbox('Instansi Pembina',options,key='dashboard_instansi')
        selected_data = df_info.copy() if selected=='Semua' else df_info[df_info['Instansi Pembina'].str.casefold()==selected.casefold()].copy()
        status_rows = df_status if selected=='Semua' else df_status[df_status['Instansi Pembina'].str.casefold()==selected.casefold()]
        if selected=='Semua':
            status_label='Seluruh Instansi'
            rencana='Lihat data pembinaan'
        elif status_rows.empty:
            status_label='Belum Tercatat'
            rencana='Belum tercantum'
        else:
            status_label=', '.join(status_rows['Status Pembinaan'].dropna().unique()) or 'Tidak tersedia'
            rencana='; '.join(x for x in status_rows['Rencana Pembinaan'].unique() if x) or 'Belum tercantum'
        jumlah_pembinaan = status_rows['Jumlah Pembinaan'].fillna(0).sum()
        subhead('Informasi Instansi')
        cards([('Instansi',selected if selected!='Semua' else 'Seluruh Instansi','Instansi yang dipilih'),('Status Pembinaan',status_label,'Status tercatat'),('Jumlah Pembinaan',fmt(jumlah_pembinaan),'Total kegiatan'),('Rencana Pembinaan',rencana,'Rencana tercatat')],info=True)
        uji_sudah = int(normalize_uji(selected_data['Status Uji Kompetensi']).sum())
        uji_belum = len(selected_data)-uji_sudah
        jf_selected = selected_data['Nomenklatur'].replace('',pd.NA).nunique()
        pemangku_selected = selected_data['Jumlah Pemangku'].fillna(0).sum()
        subhead('Ringkasan Jabatan Fungsional')
        cards([('Jabatan Fungsional',fmt(jf_selected),'Nomenklatur unik'),('Jumlah Pemangku',fmt(pemangku_selected),'Total pemangku'),('Uji Kompetensi Sudah',fmt(uji_sudah),'Baris data bertanda sudah'),('Uji Kompetensi Belum',fmt(uji_belum),'Baris data lainnya')])
        subhead('Analisis Instansi Terpilih')
        x,y=st.columns(2,gap='medium')
        with x:
            st.markdown('**Jumlah JF dan Pemangku**')
            chart(bar_chart(['Jabatan Fungsional','Pemangku'],[jf_selected,pemangku_selected],horizontal=True))
        with y:
            st.markdown('**Status Uji Kompetensi**')
            chart(donut(['Sudah','Belum'],[uji_sudah,uji_belum],[GREEN,AMBER]))
        subhead('Data Jabatan Fungsional')
        query = st.text_input('Cari Nama Jabatan',placeholder='Masukkan nama jabatan...',key='dashboard_search')
        display = selected_data[selected_data['Nomenklatur'].str.contains(query,case=False,na=False,regex=False)] if query else selected_data
        st.caption(f'Menampilkan {len(display):,} data jabatan')
        table(display)
        download(display,'data_jabatan_fungsional.csv')

elif st.session_state.active_page == 'Data Jabatan':
    with st.container(border=True):
        section('Data Jabatan Fungsional','Telusuri nomenklatur dan instansi pembina. Geser tabel secara horizontal pada layar kecil.')
        a,b=st.columns([2,1],gap='medium')
        with a:
            search=st.text_input('Cari Jabatan',placeholder='Ketik nama jabatan...')
        with b:
            options=['Semua']+sorted(x for x in df_info['Instansi Pembina'].unique() if x)
            instansi=st.selectbox('Filter Instansi',options)
        data=df_info.copy()
        if search:
            data=data[data['Nomenklatur'].str.contains(search,case=False,na=False,regex=False)]
        if instansi!='Semua':
            data=data[data['Instansi Pembina']==instansi]
        st.caption(f'Menampilkan {len(data):,} dari {len(df_info):,} data jabatan')
        table(data,600)
        download(data,'data_jabatan_fungsional.csv')
else:
    with st.container(border=True):
        section('Status Pembinaan Instansi','Pantau pelaksanaan pembinaan dan cari instansi berdasarkan status.')
        a,b=st.columns(2,gap='medium')
        with a:
            status_filter=st.selectbox('Status Pembinaan',['Semua','Sudah','Belum'])
        with b:
            search=st.text_input('Cari Instansi',placeholder='Ketik nama instansi...')
        data=df_status.copy()
        if status_filter!='Semua':
            data=data[data['Status Pembinaan'].str.casefold()==status_filter.casefold()]
        if search:
            data=data[data['Instansi Pembina'].str.contains(search,case=False,na=False,regex=False)]
        st.caption(f'Menampilkan {len(data):,} dari {len(df_status):,} data instansi')
        table(data,580)
        download(data,'data_status_pembinaan.csv')

st.html('<div class="footer">Badan Kepegawaian Negara · Dashboard Pembinaan Jabatan Fungsional</div>')
