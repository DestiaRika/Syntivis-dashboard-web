import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

st.set_page_config(page_title='Raw Purchase Data Processing', layout='centered')
st.title('🔗 Raw Purchase Data Processing')
st.markdown('---')

# Fungsi untuk merge AP Invoic
# e (FIXED)
def merge_ap_invoice(df_contens, df_awal, apply_filter=True):
    # Strip spasi dari nama kolom
    df_contens.columns = df_contens.columns.str.strip()
    df_awal.columns = df_awal.columns.str.strip()
    
    # Simpan jumlah baris awal untuk metric
    original_contens_len = len(df_contens)
    
    # === FILTER SEBELUM MERGE (FIXED!) ===
    item_desc_col = None
    if apply_filter:
        # Cari kolom Item Description
        for col in df_contens.columns:
            if "item" in col.lower() and "desc" in col.lower():
                item_desc_col = col
                break
        
        if item_desc_col:
            # Filter df_contens DULU sebelum merge
            mask = df_contens[item_desc_col].astype(str).str.lower().str.contains(
                r'jagung|zak|wheat bran', 
                case=False, 
                na=False, 
                regex=True
            )
            df_contens = df_contens[mask].reset_index(drop=True)
    
    # Ambil "base name" (tanpa angka di belakang) untuk kolom dari contens
    kolom_base = set(re.sub(r"[\s._]*\d+$", "", col).strip() for col in df_contens.columns)
    
    # Tentukan kolom tambahan dari df_awal yang belum ada di df_contens
    kolom_tambahan = [
        col for col in df_awal.columns
        if re.sub(r"[\s._]*\d+$", "", col).strip() not in kolom_base and col != "Doc Number"
    ]
    
    # Merge: base = df_contens (sudah difilter), tambahkan kolom tambahan dari df_awal
    df_merged = pd.merge(
        df_contens,
        df_awal[["Doc Number"] + kolom_tambahan],
        on="Doc Number",
        how="inner"  # hanya Doc Number yang cocok
    )
    
    # === Bersihkan nama kolom duplikat berbasis nama dasar ===
    def base_name(col):
        return re.sub(r"[\s._]*\d+$", "", col).strip()
    
    base_names = pd.Series([base_name(c) for c in df_merged.columns], index=df_merged.columns)
    df_merged = df_merged.loc[:, ~base_names.duplicated()]
    
    # === Rename khusus kalau ada 'doc date 2' ===
    if "doc date 2" in df_merged.columns:
        df_merged = df_merged.rename(columns={"doc date 2": "Doc Date"})
    
    # === Hapus kolom yang tidak dipakai ===
    kolom_hapus = ["Total", "Applied Amount", "Total Before Diskon", "Customer", "Unnamed: 45", "Netto 1"]
    df_merged = df_merged.drop(columns=[c for c in kolom_hapus if c in df_merged.columns])
    
    return df_merged, original_contens_len

# Fungsi untuk merge GRPO (FIXED - sama dengan AP Invoice)
def merge_grpo(df_contens, df_awal, apply_filter=True):
    """
    Menggabungkan df_contens dengan df_awal berdasarkan Doc Number untuk GRPO
    """
    # Strip spasi dari nama kolom
    df_contens.columns = df_contens.columns.str.strip()
    df_awal.columns = df_awal.columns.str.strip()
    
    # Simpan jumlah baris awal untuk metric
    original_contens_len = len(df_contens)
    
    # === FILTER SEBELUM MERGE (FIXED!) ===
    item_desc_col = None
    if apply_filter:
        # Cari kolom Item Description
        for col in df_contens.columns:
            if "item" in col.lower() and "desc" in col.lower():
                item_desc_col = col
                break
        
        if item_desc_col:
            # Filter df_contens DULU sebelum merge
            mask = df_contens[item_desc_col].astype(str).str.lower().str.contains(
                r'jagung|zak|wheat bran', 
                case=False, 
                na=False, 
                regex=True
            )
            df_contens = df_contens[mask].reset_index(drop=True)
    
    # Ambil "base name" (tanpa angka di belakang) untuk kolom dari contens
    kolom_base = set(re.sub(r"[\s._]*\d+$", "", col).strip() for col in df_contens.columns)
    
    # Tentukan kolom tambahan dari df_awal yang belum ada di df_contens
    kolom_tambahan = [
        col for col in df_awal.columns
        if re.sub(r"[\s._]*\d+$", "", col).strip() not in kolom_base and col != "Doc Number"
    ]
    
    # Merge: base = df_contens (sudah difilter), tambahkan kolom tambahan dari df_awal
    df_merged = pd.merge(
        df_contens,
        df_awal[["Doc Number"] + kolom_tambahan],
        on="Doc Number",
        how="inner"  # hanya Doc Number yang cocok
    )
    
    # === Bersihkan nama kolom duplikat berbasis nama dasar ===
    def base_name(col):
        return re.sub(r"[\s._]*\d+$", "", col).strip()
    
    base_names = pd.Series([base_name(c) for c in df_merged.columns], index=df_merged.columns)
    df_merged = df_merged.loc[:, ~base_names.duplicated()]
    
    # === Rename khusus kalau ada 'doc date 2' ===
    if "doc date 2" in df_merged.columns:
        df_merged = df_merged.rename(columns={"doc date 2": "Doc Date"})
    
    # === Hapus kolom yang tidak dipakai ===
    kolom_hapus = ["Total", "Applied Amount", "Total Before Diskon", "Customer", "Unnamed: 45", "Netto 1"]
    df_merged = df_merged.drop(columns=[c for c in kolom_hapus if c in df_merged.columns])
    
    return df_merged, original_contens_len

# Fungsi untuk convert dataframe ke Excel
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

st.header('💰 A/R Invoice')

col1, col2 = st.columns(2)

with col1:
    file_contens_ar = st.file_uploader(
        'Upload A/R Invoice Contents File:', 
        type=['xlsx'], 
        key='contens_ar',
        help='File yang berisi data contents/detail'
    )

with col2:
    file_awal_ar = st.file_uploader(
        'Upload A/R Invoice Base File:', 
        type=['xlsx'], 
        key='awal_ar',
        help='File yang berisi data awal/header'
    )

if file_contens_ar and file_awal_ar:
    try:
        with st.spinner('Processing A/R Invoice Data...'):
            # Baca file
            df_contens_ar = pd.read_excel(file_contens_ar)
            df_awal_ar = pd.read_excel(file_awal_ar)
            
            # Simpan jumlah baris awal
            original_contens_len = len(df_contens_ar)
            original_awal_len = len(df_awal_ar)
            
            # ===== STANDARISASI FORMAT POSTING DATE =====
            def standardize_date(df, date_column='Posting Date'):
                """Convert Posting Date ke format DD-MM-YYYY"""
                if date_column in df.columns:
                    df[date_column] = pd.to_datetime(df[date_column], errors='coerce')
                    df[date_column] = df[date_column].dt.strftime('%d-%m-%Y')
                return df
            
            # Standarisasi Posting Date di kedua file
            df_contens_ar = standardize_date(df_contens_ar)
            df_awal_ar = standardize_date(df_awal_ar)
            
            st.info('✅ Format Posting Date diseragamkan: DD-MM-YYYY')
            
            # ===== BERSIHKAN DOC NUMBER =====
            df_contens_ar['Doc Number'] = df_contens_ar['Doc Number'].astype(str).str.strip()
            df_awal_ar['Doc Number'] = df_awal_ar['Doc Number'].astype(str).str.strip()
            
            # ===== MERGE DATA (LEFT JOIN - PRIORITAS CONTENTS) =====
            df_merged_ar = pd.merge(
                df_contens_ar,
                df_awal_ar,
                on='Doc Number',
                how='left',
                suffixes=('', '_base')
            )
            
            st.success(f'✅ Merge selesai! Semua {len(df_contens_ar)} baris Contents dipertahankan')
            
            # Hitung berapa yang match dengan Base
            base_columns = [col for col in df_merged_ar.columns if col.endswith('_base')]
            if base_columns:
                matched_count = df_merged_ar[base_columns[0]].notna().sum()
            else:
                matched_count = len(df_merged_ar)
            
            # ===== FILTER ITEM DESCRIPTION =====
            # Filter hanya yang mengandung: Zak, Jagung, atau Wheat
            df_before_item_filter = df_merged_ar.copy()
            
            if 'Item Description' in df_merged_ar.columns:
                # Buat pattern untuk cari kata (case insensitive)
                keywords = ['Zak', 'Jagung', 'Wheat']
                pattern = '|'.join(keywords)
                
                # Filter data yang mengandung salah satu keyword
                df_merged_ar = df_merged_ar[
                    df_merged_ar['Item Description'].str.contains(pattern, case=False, na=False)
                ].copy()
                
                filtered_item_count = len(df_before_item_filter) - len(df_merged_ar)
                st.info(f'🔍 Filter Item Description: Hanya Zak, Jagung, Wheat ({filtered_item_count} baris dihapus)')
            else:
                st.warning('⚠️ Kolom "Item Description" tidak ditemukan, skip filter item')
            
            # ===== EKSTRAK TAHUN YANG ADA DI DATA =====
            available_years = []
            if 'Posting Date' in df_merged_ar.columns and len(df_merged_ar) > 0:
                posting_dates_temp = pd.to_datetime(
                    df_merged_ar['Posting Date'], 
                    format='%d-%m-%Y', 
                    errors='coerce'
                )
                available_years = sorted(
                    posting_dates_temp.dt.year.dropna().unique().astype(int).tolist(),
                    reverse=True  # Tahun terbaru di atas
                )
            
            # ===== TAMPILKAN FILTER TAHUN DINAMIS =====
            filter_year_ar = None
            if available_years:
                st.write("---")
                col_filter1, col_filter2 = st.columns([1, 3])
                with col_filter1:
                    filter_year_ar = st.selectbox(
                        'Filter Tahun:',
                        options=['Semua Tahun'] + [str(year) for year in available_years],
                        index=0,
                        key='filter_year_ar_dynamic',
                        help='Pilih tahun berdasarkan data yang tersedia'
                    )
                with col_filter2:
                    st.info(f"📅 Tahun tersedia: {', '.join(map(str, available_years))}")
            
            # ===== FILTER BERDASARKAN TAHUN (jika dipilih) =====
            df_before_year_filter = df_merged_ar.copy()
            
            if filter_year_ar and filter_year_ar != 'Semua Tahun' and 'Posting Date' in df_merged_ar.columns:
                # Convert Posting Date dari string ke datetime untuk filter
                df_merged_ar['Posting Date_datetime'] = pd.to_datetime(
                    df_merged_ar['Posting Date'], 
                    format='%d-%m-%Y', 
                    errors='coerce'
                )
                
                # Ekstrak tahun
                df_merged_ar['Year'] = df_merged_ar['Posting Date_datetime'].dt.year
                
                # Filter hanya tahun yang dipilih
                selected_year = int(filter_year_ar)
                df_merged_ar = df_merged_ar[df_merged_ar['Year'] == selected_year].copy()
                
                # Hapus kolom temporary
                df_merged_ar = df_merged_ar.drop(columns=['Posting Date_datetime', 'Year'])
                
                filtered_year_count = len(df_before_year_filter) - len(df_merged_ar)
                st.info(f'📅 Data difilter untuk tahun {filter_year_ar} ({filtered_year_count} baris dihapus)')
            
            # ===== TAMPILKAN STATISTIK =====
            st.write("---")
            col_info1, col_info2, col_info3, col_info4 = st.columns(4)
            with col_info1:
                st.metric('📋 Baris Contents', original_contens_len)
            with col_info2:
                st.metric('📄 Baris Base', original_awal_len)
            with col_info3:
                st.metric('✅ Match dengan Base', matched_count)
            with col_info4:
                st.metric('📊 Hasil Akhir', len(df_merged_ar))
            
            # Warning jika ada data yang tidak match
            not_matched = original_contens_len - matched_count
            if not_matched > 0:
                st.warning(f'⚠️ {not_matched} baris Contents tidak memiliki data Base yang cocok (tetap dipertahankan)')
            
            # Info hasil filter
            total_filtered = original_contens_len - len(df_merged_ar)
            if total_filtered > 0:
                st.info(f'📊 Total difilter: {total_filtered} baris dari {original_contens_len} baris awal')
            
            if len(df_merged_ar) == 0:
                st.error(f'❌ Tidak ada data setelah filter!')
            else:
                # ===== PREVIEW DATA =====
                st.subheader('📊 Preview Result Data:')
                st.dataframe(df_merged_ar.head(10), use_container_width=True)
                
                # ===== DISTRIBUSI TAHUN DAN ITEM =====
                with st.expander("📈 Analisis Data"):
                    col_anal1, col_anal2 = st.columns(2)
                    
                    with col_anal1:
                        st.write("**Distribusi per Tahun:**")
                        if 'Posting Date' in df_merged_ar.columns:
                            posting_dates = pd.to_datetime(
                                df_merged_ar['Posting Date'], 
                                format='%d-%m-%Y', 
                                errors='coerce'
                            )
                            year_counts = posting_dates.dt.year.value_counts().sort_index()
                            for year, count in year_counts.items():
                                if pd.notna(year):
                                    st.write(f"- {int(year)}: {count} baris")
                    
                    with col_anal2:
                        st.write("**Distribusi Item Description:**")
                        if 'Item Description' in df_merged_ar.columns:
                            # Hitung kemunculan keyword
                            for keyword in ['Zak', 'Jagung', 'Wheat']:
                                count = df_merged_ar['Item Description'].str.contains(
                                    keyword, case=False, na=False
                                ).sum()
                                st.write(f"- {keyword}: {count} baris")
                
                # ===== NAMA FILE & DOWNLOAD =====
                # Buat nama file berdasarkan filter yang dipilih
                if filter_year_ar and filter_year_ar != 'Semua Tahun':
                    # Jika user pilih tahun spesifik, gunakan tahun itu
                    file_name = f'AR_INVOICE_{filter_year_ar}.xlsx'
                else:
                    # Jika "Semua Tahun", gunakan range tahun dari data
                    years_in_data = []
                    if 'Posting Date' in df_merged_ar.columns:
                        posting_dates = pd.to_datetime(
                            df_merged_ar['Posting Date'], 
                            format='%d-%m-%Y', 
                            errors='coerce'
                        )
                        years_in_data = sorted(posting_dates.dt.year.dropna().unique().astype(int).tolist())
                    
                    if years_in_data:
                        if len(years_in_data) == 1:
                            year_str = str(years_in_data[0])
                        else:
                            year_str = f"{years_in_data[0]}-{years_in_data[-1]}"
                        file_name = f'AR_INVOICE_{year_str}.xlsx'
                    else:
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        file_name = f'AR_INVOICE_{timestamp}.xlsx'
                
                # Button download
                excel_data = to_excel(df_merged_ar)
                
                st.download_button(
                    label=f'📥 Download A/R Invoice Data ({len(df_merged_ar)} baris)',
                    data=excel_data,
                    file_name=file_name,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    type='primary'
                )
            
    except Exception as e:
        st.error(f'❌ Error: {str(e)}')
        st.exception(e)
        st.info('💡 Pastikan:')
        st.write('- Kedua file memiliki kolom "Doc Number"')
        st.write('- File Contents memiliki kolom "Posting Date" dan "Item Description"')
        st.write('- Format file adalah .xlsx yang valid')

st.markdown('---')

# Section 2: GRPO
st.header('📦 Goods Receipt Unloading')

col3, col4 = st.columns(2)

with col3:
    file_contens_grpo = st.file_uploader(
        'Upload Goods Receipt Unloading Contents File:', 
        type=['xlsx'], 
        key='contens_grpo',
        help='File yang berisi data contents/detail'
    )

with col4:
    file_awal_grpo = st.file_uploader(
        'Upload Goods Receipt Unloading Base File:', 
        type=['xlsx'], 
        key='awal_grpo',
        help='File yang berisi data awal/header'
    )

if file_contens_grpo and file_awal_grpo:
    try:
        with st.spinner('Process Goods Receipt Unloading Data...'):
            # Baca file
            df_contens_grpo = pd.read_excel(file_contens_grpo)
            df_awal_grpo = pd.read_excel(file_awal_grpo)
            
            # Simpan jumlah baris awal
            original_awal_grpo_len = len(df_awal_grpo)
            
            # Merge data dengan logika GRPO
            df_merged_grpo, original_contens_grpo_len = merge_grpo(
                df_contens_grpo, df_awal_grpo, apply_filter=True  # Selalu filter
            )
            
            # Tampilkan info
            st.success('Success!')
            
            col_info4, col_info5, col_info6 = st.columns(3)
            with col_info4:
                st.metric('Baris Contents', original_contens_grpo_len)
            with col_info5:
                st.metric('Baris Awal', original_awal_grpo_len)
            with col_info6:
                st.metric('Baris Hasil', len(df_merged_grpo))
            
            # Preview data
            st.subheader('Preview Result Data:')
            st.dataframe(df_merged_grpo.head(10), use_container_width=True)
            
            # Ekstrak tahun dari Posting Date untuk nama file
            years_in_data_grpo = []
            if 'Posting Date' in df_merged_grpo.columns:
                # Convert ke datetime dan ekstrak tahun
                posting_dates_grpo = pd.to_datetime(df_merged_grpo['Posting Date'], errors='coerce')
                years_in_data_grpo = sorted(posting_dates_grpo.dt.year.dropna().unique().astype(int).tolist())
            
            # Buat nama file dengan tahun (format sama dengan AP Invoice)
            if years_in_data_grpo:
                if len(years_in_data_grpo) == 1:
                    year_str_grpo = str(years_in_data_grpo[0])
                else:
                    year_str_grpo = f"{years_in_data_grpo[0]}-{years_in_data_grpo[-1]}"
                file_name_grpo = f'GRPO_{year_str_grpo}.xlsx'
            else:
                # Fallback jika tidak ada tahun
                timestamp_grpo = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name_grpo = f'GRPO_{timestamp_grpo}.xlsx'
            
            # Button download
            excel_data_grpo = to_excel(df_merged_grpo)
            
            st.download_button(
                label='📥 Download Goods Receipt Unloading Data',
                data=excel_data_grpo,
                file_name=file_name_grpo,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary'
            )
            
    except Exception as e:
        st.error(f'❌ Error: {str(e)}')
        st.info('💡 Make sure both files have a "Doc Number" column')

st.markdown('---')

# Section 3: AP Down Payment
st.header('💰 A/P Down Payment')

col5, col6 = st.columns(2)

with col5:
    file_contens_apdp = st.file_uploader(
        'Upload A/P Down Payment Contents File:', 
        type=['xlsx'], 
        key='contens_apdp',
        help='File yang berisi data contents/detail'
    )

with col6:
    file_awal_apdp = st.file_uploader(
        'Upload A/P Down Payment Base File:', 
        type=['xlsx'], 
        key='awal_apdp',
        help='File yang berisi data awal/header'
    )

if file_contens_apdp and file_awal_apdp:
    try:
        with st.spinner('Process A/P Down Payment...'):
            # Baca file
            df_contens_apdp = pd.read_excel(file_contens_apdp)
            df_awal_apdp = pd.read_excel(file_awal_apdp)
            
            # Simpan jumlah baris awal
            original_awal_apdp_len = len(df_awal_apdp)
            
            # Merge data dengan logika AP Invoice (sama persis)
            df_merged_apdp, original_contens_apdp_len = merge_ap_invoice(
                df_contens_apdp, df_awal_apdp, apply_filter=True  # Selalu filter
            )
            
            # Tampilkan info
            st.success('Success!')
            
            col_info7, col_info8, col_info9 = st.columns(3)
            with col_info7:
                st.metric('Baris Contents', original_contens_apdp_len)
            with col_info8:
                st.metric('Baris Awal', original_awal_apdp_len)
            with col_info9:
                st.metric('Baris Hasil', len(df_merged_apdp))
            
            # Preview data
            st.subheader('Preview Result Data:')
            st.dataframe(df_merged_apdp.head(10), use_container_width=True)
            
            # Ekstrak tahun dari Posting Date untuk nama file
            years_in_data_apdp = []
            if 'Posting Date' in df_merged_apdp.columns:
                # Convert ke datetime dan ekstrak tahun
                posting_dates_apdp = pd.to_datetime(df_merged_apdp['Posting Date'], errors='coerce')
                years_in_data_apdp = sorted(posting_dates_apdp.dt.year.dropna().unique().astype(int).tolist())
            
            # Buat nama file dengan tahun
            if years_in_data_apdp:
                if len(years_in_data_apdp) == 1:
                    year_str_apdp = str(years_in_data_apdp[0])
                else:
                    year_str_apdp = f"{years_in_data_apdp[0]}-{years_in_data_apdp[-1]}"
                file_name_apdp = f'APDP_{year_str_apdp}.xlsx'
            else:
                # Fallback jika tidak ada tahun
                timestamp_apdp = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name_apdp = f'APDP_{timestamp_apdp}.xlsx'
            
            # Button download
            excel_data_apdp = to_excel(df_merged_apdp)
            
            st.download_button(
                label='📥 Download A/P Down Payment Data',
                data=excel_data_apdp,
                file_name=file_name_apdp,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary'
            )
            
    except Exception as e:
        st.error(f'❌ Error: {str(e)}')
        st.info('💡 Make sure both files have a "Doc Number" column')

st.markdown('---')

# Section 4: PO (Purchase Order)
st.header('📋 Purchase Order')

col7, col8 = st.columns(2)

with col7:
    file_contens_po = st.file_uploader(
        'Upload Purchase Order Contents File:', 
        type=['xlsx'], 
        key='contens_po',
        help='File yang berisi data contents/detail'
    )

with col8:
    file_awal_po = st.file_uploader(
        'Upload Purchase Order Base File:', 
        type=['xlsx'], 
        key='awal_po',
        help='File yang berisi data awal/header'
    )

if file_contens_po and file_awal_po:
    try:
        with st.spinner('Process Purchase Order...'):
            # Baca file
            df_contens_po = pd.read_excel(file_contens_po)
            df_awal_po = pd.read_excel(file_awal_po)
            
            # Simpan jumlah baris awal
            original_awal_po_len = len(df_awal_po)
            
            # Merge data dengan logika AP Invoice (sama persis)
            df_merged_po, original_contens_po_len = merge_ap_invoice(
                df_contens_po, df_awal_po, apply_filter=True  # Selalu filter
            )
            
            # Tampilkan info
            st.success('Success!')
            
            col_info10, col_info11, col_info12 = st.columns(3)
            with col_info10:
                st.metric('Baris Contents', original_contens_po_len)
            with col_info11:
                st.metric('Baris Awal', original_awal_po_len)
            with col_info12:
                st.metric('Baris Hasil', len(df_merged_po))
            
            # Preview data
            st.subheader('Preview Result Data:')
            st.dataframe(df_merged_po.head(10), use_container_width=True)
            
            # Ekstrak tahun dari Posting Date untuk nama file
            years_in_data_po = []
            if 'Posting Date' in df_merged_po.columns:
                # Convert ke datetime dan ekstrak tahun
                posting_dates_po = pd.to_datetime(df_merged_po['Posting Date'], errors='coerce')
                years_in_data_po = sorted(posting_dates_po.dt.year.dropna().unique().astype(int).tolist())
            
            # Buat nama file dengan tahun
            if years_in_data_po:
                if len(years_in_data_po) == 1:
                    year_str_po = str(years_in_data_po[0])
                else:
                    year_str_po = f"{years_in_data_po[0]}-{years_in_data_po[-1]}"
                file_name_po = f'PO_{year_str_po}.xlsx'
            else:
                # Fallback jika tidak ada tahun
                timestamp_po = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name_po = f'PO_{timestamp_po}.xlsx'
            
            # Button download
            excel_data_po = to_excel(df_merged_po)
            
            st.download_button(
                label='📥 Download Purchase Order Data',
                data=excel_data_po,
                file_name=file_name_po,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary'
            )
            
    except Exception as e:
        st.error(f'❌ Error: {str(e)}')
        st.info('💡 Make sure both files have a "Doc Number" column')

st.markdown('---')

# Section 5: Timbangan (Goods Receipt Unloading)
st.header('⚖️ Goods Receipt Unloading')

col9, col10 = st.columns(2)

with col9:
    file_contens_timbangan = st.file_uploader(
        'Upload Goods Receipt Unloading Contents File:', 
        type=['xlsx'], 
        key='contens_timbangan',
        help='File yang berisi data contents/detail'
    )

with col10:
    file_awal_timbangan = st.file_uploader(
        'Upload TGoods Receipt Unloading Awal File:', 
        type=['xlsx'], 
        key='awal_timbangan',
        help='File yang berisi data awal/header'
    )

if file_contens_timbangan and file_awal_timbangan:
    try:
        with st.spinner('Process Goods Receipt Unloading...'):
            # Baca file
            df_contens_timbangan = pd.read_excel(file_contens_timbangan)
            df_awal_timbangan = pd.read_excel(file_awal_timbangan)
            
            # Simpan jumlah baris awal
            original_awal_timbangan_len = len(df_awal_timbangan)
            
            # Merge data dengan logika AP Invoice (sama persis)
            df_merged_timbangan, original_contens_timbangan_len = merge_ap_invoice(
                df_contens_timbangan, df_awal_timbangan, apply_filter=True  # Selalu filter
            )
            
            # Tampilkan info
            st.success('Success!')
            
            col_info13, col_info14, col_info15 = st.columns(3)
            with col_info13:
                st.metric('Baris Contents', original_contens_timbangan_len)
            with col_info14:
                st.metric('Baris Awal', original_awal_timbangan_len)
            with col_info15:
                st.metric('Baris Hasil', len(df_merged_timbangan))
            
            # Preview data
            st.subheader('Preview Result Data:')
            st.dataframe(df_merged_timbangan.head(10), use_container_width=True)
            
            # Ekstrak tahun dari Out Date untuk nama file (khusus Timbangan)
            years_in_data_timbangan = []
            if 'Out Date' in df_merged_timbangan.columns:
                # Convert ke datetime dan ekstrak tahun
                out_dates = pd.to_datetime(df_merged_timbangan['Out Date'], errors='coerce')
                years_in_data_timbangan = sorted(out_dates.dt.year.dropna().unique().astype(int).tolist())
            
            # Buat nama file dengan tahun
            if years_in_data_timbangan:
                if len(years_in_data_timbangan) == 1:
                    year_str_timbangan = str(years_in_data_timbangan[0])
                else:
                    year_str_timbangan = f"{years_in_data_timbangan[0]}-{years_in_data_timbangan[-1]}"
                file_name_timbangan = f'TIMBANGAN BELI_{year_str_timbangan}.xlsx'
            else:
                # Fallback jika tidak ada tahun
                timestamp_timbangan = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name_timbangan = f'TIMBANGAN BELI_{timestamp_timbangan}.xlsx'
            
            # Button download
            excel_data_timbangan = to_excel(df_merged_timbangan)
            
            st.download_button(
                label='📥 Download Weighing Data',
                data=excel_data_timbangan,
                file_name=file_name_timbangan,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary'
            )
            
    except Exception as e:
        st.error(f'❌ Error: {str(e)}')
        st.info('💡 Make sure both files have a "Doc Number" column')

st.markdown('---')
st.caption('💡 Ensure both files contain a "Doc Number" column before merging the data.')