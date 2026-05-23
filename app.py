import streamlit as st
import pandas as pd
import openpyxl
import io
import os

# Sayfa Yapılandırması (Modern ve Sade Tasarım için Geniş Ekran Modu)
st.set_page_config(
    page_title="DataMorph // Universal Data Converter",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ----------------- DİL VE ÇEVİRİ SİSTEMİ (TRANSLATION SYSTEM) -----------------
# English is set as the default/primary language (EN), Turkish is the secondary (TR)
LANGUAGES = {
    "EN": {
        "page_title": "DataMorph // Universal Data Converter & Cleaner",
        "title": "DataMorph // Universal Data Converter",
        "subtitle": "Upload complex data files, clean them in seconds, and convert them to your desired format.",
        "uploader_label": "Upload your data file for conversion and cleaning",
        "uploader_help": "Supported formats: .csv (Comma separated), .xlsx (Excel sheet), .json (JSON data)",
        "empty_file_warning": "⚠️ The uploaded file is empty or does not contain any readable rows.",
        "metrics_title": "📊 Dataset Metrics",
        "metric_total_rows": "Total Rows",
        "metric_columns": "Columns Count",
        "metric_file_size": "File Size",
        "metric_missing": "Missing Values",
        "cleaning_title": "🧹 Cleaning Options",
        "clean_missing_label": "Remove Empty & Missing Rows",
        "clean_missing_help": "Completely removes any row that contains at least one empty (NaN/null) cell.",
        "clean_text_label": "Standardize Text Fields",
        "clean_text_help": "Converts all text (String) columns to Title Case (e.g., 'jOHN sMITH' ➔ 'John Smith'). Preserves missing values.",
        "clean_dup_label": "Clean Duplicate Rows",
        "clean_dup_help": "Removes identical duplicate rows, keeping only the first occurrence.",
        "preview_title": "🔍 Data Preview",
        "tab_processed": "📋 Processed & Preview Data (First 5)",
        "tab_original": "⏳ Original Uploaded Data (First 5)",
        "rows_deleted_info": "✨ Cleaning removed a total of **{rows}** rows. Remaining rows: **{remaining}**.",
        "no_rows_deleted_info": "No rows have been deleted yet or no cleaning rules are applied.",
        "original_data_caption": "Here are the first 5 rows of your uploaded raw data:",
        "download_title": "📥 Convert & Download",
        "target_format_label": "Select target file format:",
        "download_btn_label": "✨ Clean & Download as {format}",
        "read_error": "❌ An error occurred while uploading or reading the file. Please verify file integrity.\nDetails: {error}",
        "welcome_info": "💡 To get started, upload a data file in .csv, .xlsx, or .json format by dragging and dropping it above.",
        "feat_1_title": "🚀 Fast Conversion",
        "feat_1_desc": "Convert large data files (.csv, .xlsx, .json) into your desired format in seconds without errors.",
        "feat_2_title": "🧼 Easy Cleaning",
        "feat_2_desc": "Remove missing values, standardize text fields, and eliminate duplicate entries in a single click.",
        "feat_3_title": "💎 Premium UI",
        "feat_3_desc": "Enjoy a custom dark mode, live data metrics, and a modern glassmorphic interface that keeps visual strain low.",
        "footer_text": "DataMorph © 2026 // Advanced Universal Data Converter and Cleaner"
    },
    "TR": {
        "page_title": "DataMorph // Evrensel Veri Dönüştürücü ve Temizleyici",
        "title": "DataMorph // Evrensel Veri Dönüştürücü",
        "subtitle": "Karmaşık veri dosyalarınızı yükleyin, saniyeler içinde temizleyin ve istediğiniz formata dönüştürün.",
        "uploader_label": "Dönüştürmek ve temizlemek istediğiniz veriyi yükleyin",
        "uploader_help": "Desteklenen formatlar: .csv (Virgülle ayrılmış), .xlsx (Excel tablosu), .json (JSON verisi)",
        "empty_file_warning": "⚠️ Yüklenen dosya boş veya okunabilir herhangi bir satır içermiyor.",
        "metrics_title": "📊 Dosya Metrikleri",
        "metric_total_rows": "Toplam Satır",
        "metric_columns": "Sütun Sayısı",
        "metric_file_size": "Dosya Boyutu",
        "metric_missing": "Eksik Değer",
        "cleaning_title": "🧹 Temizleme Seçenekleri",
        "clean_missing_label": "Boş ve Eksik Satırları Sil",
        "clean_missing_help": "Herhangi bir hücresi boş (NaN) olan tüm satırları tablodan tamamen kaldırır.",
        "clean_text_label": "Metin Alanlarını Standartlaştır",
        "clean_text_help": "Metin (String) içeren tüm sütunların ilk harflerini büyük, diğerlerini küçük yapar (Title Case). Boş değerleri bozmaz.",
        "clean_dup_label": "Yinelenen (Duplicate) Satırları Temizle",
        "clean_dup_help": "Tablodaki tamamen aynı olan yinelenen satırları kaldırır, yalnızca ilkini korur.",
        "preview_title": "🔍 Veri Önizleme",
        "tab_processed": "📋 İşlenmiş & Önizleme Verisi (İlk 5)",
        "tab_original": "⏳ Orijinal Yüklenen Veri (İlk 5)",
        "rows_deleted_info": "✨ Temizleme sonucunda toplam **{rows}** satır elendi. Kalan satır: **{remaining}**.",
        "no_rows_deleted_info": "Şu anda herhangi bir satır silinmedi veya temizleme kuralları seçilmedi.",
        "original_data_caption": "Yüklediğiniz ham verinin ilk 5 satırı aşağıdadır:",
        "download_title": "📥 Dönüştür ve İndir",
        "target_format_label": "Dönüştürmek istediğiniz hedef dosya formatını seçin:",
        "download_btn_label": "✨ Temizle ve {format} Olarak İndir",
        "read_error": "❌ Dosya yüklenirken veya okunurken bir sorun oluştu. Lütfen dosya bütünlüğünü kontrol edin.\nDetay: {error}",
        "welcome_info": "💡 Başlamak için yukarıdaki alana sürükleyip bırakarak veya dosyayı seçerek .csv, .xlsx veya .json formatında bir veri dosyası yükleyin.",
        "feat_1_title": "🚀 Hızlı Dönüşüm",
        "feat_1_desc": "Büyük veri dosyalarınızı (.csv, .xlsx, .json) dilediğiniz formata saniyeler içinde hatasız bir şekilde dönüştürün.",
        "feat_2_title": "🧼 Kolay Temizlik",
        "feat_2_desc": "Tek tıkla eksik verileri silin, metinlerinizi standartlaştırın ve yinelenen kayıtları temizleyerek veri kalitenizi artırın.",
        "feat_3_title": "💎 Premium Arayüz",
        "feat_3_desc": "Göz yormayan koyu mod tasarımı, anlık güncellenen veri metrikleri ve modern arayüz ile veri analizi artık çok daha keyifli.",
        "footer_text": "DataMorph © 2026 // Gelişmiş Evrensel Veri Dönüştürücü ve Temizleyici"
    }
}

# Premium UI/UX için Gelişmiş CSS Tasarım Sistemi Enjeksiyonu
st.markdown("""
<style>
/* Modern Google Fontu (Outfit) Entegrasyonu */
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&display=swap');

/* Tüm Uygulama ve Zemin Özelleştirmeleri */
html, body, [data-testid="stAppViewContainer"] {
    font-family: 'Outfit', sans-serif;
    background-color: #0b0f19;
    color: #e2e8f0;
}

/* Üst Header Blur Efekti */
[data-testid="stHeader"] {
    background-color: rgba(11, 15, 25, 0.8) !important;
    backdrop-filter: blur(12px);
}

/* Glassmorphism Konteyner Sınıfı (Özel Kutular İçin) */
.glass-card {
    background: rgba(17, 25, 40, 0.65);
    backdrop-filter: blur(12px) saturate(180%);
    -webkit-backdrop-filter: blur(12px) saturate(180%);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 16px;
    padding: 26px;
    margin-bottom: 24px;
    box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
}

/* Gradyan Başlık Tasarımı */
.gradient-title {
    background: linear-gradient(135deg, #a78bfa 0%, #22d3ee 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-weight: 700;
    font-size: 2.6rem;
    letter-spacing: -0.03em;
    margin-bottom: 4px;
    display: inline-block;
}

.gradient-subtitle {
    color: #94a3b8;
    font-size: 1.05rem;
    margin-bottom: 30px;
}

/* Özel Metrik Panelleri */
.metric-container {
    display: flex;
    gap: 16px;
    margin-bottom: 20px;
}

.metric-card {
    flex: 1;
    background: rgba(255, 255, 255, 0.03);
    border: 1px solid rgba(255, 255, 255, 0.06);
    border-radius: 12px;
    padding: 16px;
    text-align: center;
    transition: all 0.3s ease;
}

.metric-card:hover {
    border-color: rgba(34, 211, 238, 0.4);
    background: rgba(255, 255, 255, 0.05);
}

.metric-num {
    font-size: 1.8rem;
    font-weight: 700;
    color: #22d3ee;
}

.metric-num.deleted {
    color: #f43f5e;
}

.metric-title {
    font-size: 0.85rem;
    color: #94a3b8;
    margin-top: 6px;
    font-weight: 500;
}

/* Seçenek Kutuları (Checkbox) Tasarımı */
div[data-testid="stCheckbox"] {
    background: rgba(255, 255, 255, 0.025);
    border: 1px solid rgba(255, 255, 255, 0.06);
    border-radius: 12px;
    padding: 12px 18px;
    margin-bottom: 12px;
    transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
}

div[data-testid="stCheckbox"]:hover {
    background: rgba(255, 255, 255, 0.05);
    border-color: rgba(167, 139, 250, 0.5);
    transform: translateX(4px);
}

/* Tablo Sınırlarını Yumuşatma */
div[data-testid="stDataFrame"] {
    border: 1px solid rgba(255, 255, 255, 0.08) !important;
    border-radius: 12px !important;
    overflow: hidden;
}

/* Premium "Temizle ve İndir" Butonu */
div.stButton > button, div.stDownloadButton > button {
    background: linear-gradient(135deg, #7c3aed 0%, #06b6d4 100%) !important;
    color: #ffffff !important;
    border: none !important;
    padding: 14px 28px !important;
    font-weight: 600 !important;
    font-size: 1.1rem !important;
    border-radius: 12px !important;
    box-shadow: 0 4px 20px rgba(124, 58, 237, 0.4) !important;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    width: 100% !important;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-top: 15px;
}

div.stButton > button:hover, div.stDownloadButton > button:hover {
    transform: translateY(-3px) !important;
    box-shadow: 0 8px 25px rgba(6, 182, 212, 0.6) !important;
    background: linear-gradient(135deg, #6d28d9 0%, #0891b2 100%) !important;
}

div.stButton > button:active, div.stDownloadButton > button:active {
    transform: translateY(1px) !important;
}

/* Dil seçici alanı için küçük düzenleme */
.lang-selector-box {
    margin-bottom: 20px;
}

/* Alt Bilgi Footer Alanı */
.footer {
    text-align: center;
    margin-top: 50px;
    color: #475569;
    font-size: 0.85rem;
    border-top: 1px solid rgba(255, 255, 255, 0.04);
    padding-top: 20px;
}
</style>
""", unsafe_allow_html=True)

# ----------------- EN ÜST ALAN: BAŞLIK VE DİL SEÇİCİ -----------------
# 2 Kolon: Sol kolon gradient başlık, Sağ kolon modern dil seçici dropdown
col_header_left, col_header_right = st.columns([5, 1], gap="medium")

with col_header_right:
    # Language dropdown (English is default with index 0)
    lang_choice = st.selectbox(
        "🌐 Language / Dil",
        options=["English (EN)", "Türkçe (TR)"],
        index=0,
        help="Switch application language / Uygulama dilini değiştirin"
    )
    # Get lang code
    lang_code = "EN" if "English" in lang_choice else "TR"
    # Load translation dictionary
    t = LANGUAGES[lang_code]

with col_header_left:
    st.markdown(f'<div class="gradient-title">{t["title"]}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="gradient-subtitle">{t["subtitle"]}</div>', unsafe_allow_html=True)


# ----------------- YARDIMCI FONKSİYONLAR -----------------
def get_file_size(uploaded_file):
    """Dosya boyutunu okuyup uygun birimle (KB veya MB) döndürür."""
    try:
        bytes_data = uploaded_file.getvalue()
        size_bytes = len(bytes_data)
        if size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.2f} KB"
        else:
            return f"{size_bytes / (1024 * 1024):.2f} MB"
    except Exception:
        return "N/A"

# ----------------- DOSYA YÜKLEME ALANI -----------------
st.markdown('<div class="glass-card">', unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    t["uploader_label"], 
    type=["csv", "xlsx", "json"],
    help=t["uploader_help"]
)
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file is not None:
    # Dosya adını ve uzantısını alalım
    file_name = uploaded_file.name
    file_extension = os.path.splitext(file_name)[1].lower()
    
    # Dosyayı Pandas DataFrame olarak okuma
    df_raw = None
    read_success = True
    error_msg = ""
    
    try:
        if file_extension == '.csv':
            # İlk olarak utf-8 ile okumayı dene, hata olursa latin-1 dene (Türkçe karakter desteği için)
            try:
                df_raw = pd.read_csv(uploaded_file)
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, encoding='latin-1')
        elif file_extension == '.xlsx':
            df_raw = pd.read_excel(uploaded_file, engine='openpyxl')
        elif file_extension == '.json':
            df_raw = pd.read_json(uploaded_file)
    except Exception as e:
        read_success = False
        error_msg = str(e)
        
    # Eğer dosya okuma başarılıysa ana paneli gösterelim
    if read_success and df_raw is not None:
        # Boş veri kontrolü
        if df_raw.empty:
            st.warning(t["empty_file_warning"])
        else:
            # Temel istatistikleri ve bilgileri hazırlayalım
            original_rows = len(df_raw)
            original_cols = len(df_raw.columns)
            total_missing = df_raw.isnull().sum().sum()
            file_size_str = get_file_size(uploaded_file)
            
            # Uygulama Arayüzünü 2 Kolona Bölelim
            # Sol Kolon: Temizleme Ayarları & Dosya Bilgileri
            # Sağ Kolon: Veri Önizleme & Son Dönüştürme Kontrolleri
            col_left, col_right = st.columns([1, 2], gap="large")
            
            with col_left:
                st.markdown(f'<h3 style="color: #a78bfa; margin-bottom: 15px;">{t["metrics_title"]}</h3>', unsafe_allow_html=True)
                
                # Şık Metrik Kartları
                st.markdown(f"""
                <div class="metric-container">
                    <div class="metric-card">
                        <div class="metric-num">{original_rows}</div>
                        <div class="metric-title">{t["metric_total_rows"]}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-num">{original_cols}</div>
                        <div class="metric-title">{t["metric_columns"]}</div>
                    </div>
                </div>
                <div class="metric-container">
                    <div class="metric-card">
                        <div class="metric-num">{file_size_str}</div>
                        <div class="metric-title">{t["metric_file_size"]}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-num {'deleted' if total_missing > 0 else ''}">{total_missing}</div>
                        <div class="metric-title">{t["metric_missing"]}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown('<hr style="border-color: rgba(255,255,255,0.05); margin: 25px 0;">', unsafe_allow_html=True)
                
                st.markdown(f'<h3 style="color: #a78bfa; margin-bottom: 15px;">{t["cleaning_title"]}</h3>', unsafe_allow_html=True)
                
                # Temizleme Seçenekleri Checkbox Alanı
                clean_missing = st.checkbox(
                    t["clean_missing_label"],
                    value=False,
                    help=t["clean_missing_help"]
                )
                
                clean_text = st.checkbox(
                    t["clean_text_label"],
                    value=False,
                    help=t["clean_text_help"]
                )
                
                clean_duplicates = st.checkbox(
                    t["clean_dup_label"],
                    value=False,
                    help=t["clean_dup_help"]
                )
                
            # Veriyi temizleme ayarlarına göre dönüştürelim
            df_cleaned = df_raw.copy()
            
            # Adım 1: Boş ve eksik satırları sil
            if clean_missing:
                df_cleaned = df_cleaned.dropna()
                
            # Adım 2: Metin alanlarını standartlaştır (Title Case)
            if clean_text:
                # Sadece object/string türündeki sütunları tespit edip dönüştür
                text_columns = df_cleaned.select_dtypes(include=['object', 'string']).columns
                for col in text_columns:
                    # null değerleri bozmadan dönüştürme işlemi
                    df_cleaned[col] = df_cleaned[col].apply(
                        lambda x: str(x).title() if pd.notnull(x) and str(x).strip() != '' else x
                    )
            
            # Adım 3: Yinelenen satırları temizle
            if clean_duplicates:
                df_cleaned = df_cleaned.drop_duplicates()
                
            # Temizlenmiş veri boyutları ve değişim metrikleri
            cleaned_rows = len(df_cleaned)
            rows_deleted = original_rows - cleaned_rows
            
            with col_right:
                st.markdown(f'<h3 style="color: #22d3ee; margin-bottom: 15px;">{t["preview_title"]}</h3>', unsafe_allow_html=True)
                
                # Orijinal ve Temizlenmiş durumların farkını gösteren sekmeler (Tabs)
                tab1, tab2 = st.tabs([t["tab_processed"], t["tab_original"]])
                
                with tab1:
                    if rows_deleted > 0:
                        st.info(t["rows_deleted_info"].format(rows=rows_deleted, remaining=cleaned_rows))
                    else:
                        st.caption(t["no_rows_deleted_info"])
                    
                    # İlk 5 satırı şık bir tabloda gösterelim
                    st.dataframe(df_cleaned.head(5), use_container_width=True)
                    
                with tab2:
                    st.caption(t["original_data_caption"])
                    st.dataframe(df_raw.head(5), use_container_width=True)
                
                st.markdown('<hr style="border-color: rgba(255,255,255,0.05); margin: 25px 0;">', unsafe_allow_html=True)
                
                st.markdown(f'<h3 style="color: #22d3ee; margin-bottom: 15px;">{t["download_title"]}</h3>', unsafe_allow_html=True)
                
                # Hedef format seçimi
                target_format = st.selectbox(
                    t["target_format_label"],
                    options=["CSV (.csv)", "Excel (.xlsx)", "JSON (.json)"],
                    index=0
                )
                
                # Bellek içi tampon oluşturma ve indirme butonu
                download_ready = False
                output_bytes = None
                output_mime = ""
                output_filename = ""
                
                # Format eşleştirme
                if target_format == "CSV (.csv)":
                    try:
                        csv_string = df_cleaned.to_csv(index=False)
                        output_bytes = csv_string.encode('utf-8')
                        output_mime = "text/csv"
                        base_name = os.path.splitext(file_name)[0]
                        output_filename = f"{base_name}_datamorph.csv"
                        download_ready = True
                    except Exception as e:
                        st.error(f"CSV Conversion Error: {str(e)}")
                        
                elif target_format == "Excel (.xlsx)":
                    try:
                        # Bellek içinde Excel dosyası oluşturmak için BytesIO kullanımı
                        excel_buffer = io.BytesIO()
                        # openpyxl motoru ile Excel dosyasına yazma
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            df_cleaned.to_excel(writer, index=False, sheet_name='DataMorph_Clean')
                        output_bytes = excel_buffer.getvalue()
                        output_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        base_name = os.path.splitext(file_name)[0]
                        output_filename = f"{base_name}_datamorph.xlsx"
                        download_ready = True
                    except Exception as e:
                        st.error(f"Excel Conversion Error: {str(e)}")
                        
                elif target_format == "JSON (.json)":
                    try:
                        # JSON verisini düzgün satır aralıkları çıktı alma
                        json_str = df_cleaned.to_json(orient='records', force_ascii=False, indent=4)
                        output_bytes = json_str.encode('utf-8')
                        output_mime = "application/json"
                        base_name = os.path.splitext(file_name)[0]
                        output_filename = f"{base_name}_datamorph.json"
                        download_ready = True
                    except Exception as e:
                        st.error(f"JSON Conversion Error: {str(e)}")
                
                # Dinamik İndirme Butonu Gösterimi
                if download_ready and output_bytes is not None:
                    # İndirme butonunun üzerine gelindiğinde kullanıcıyı bilgilendirecek metin
                    btn_format_name = target_format.split(' ')[0]
                    btn_label = t["download_btn_label"].format(format=btn_format_name)
                    st.download_button(
                        label=btn_label,
                        data=output_bytes,
                        file_name=output_filename,
                        mime=output_mime,
                        use_container_width=True
                    )
                    
    else:
        st.error(t["read_error"].format(error=error_msg))

else:
    # Dosya yüklenmediğinde şık bir bilgilendirme karşılama alanı gösterelim
    st.info(t["welcome_info"])
    
    # Uygulama özellikleri hakkında kısa görsel kartlar sunalım
    col_feat1, col_feat2, col_feat3 = st.columns(3)
    with col_feat1:
        st.markdown(f"""
        <div class="metric-card" style="min-height: 150px; text-align: left; padding: 20px;">
            <div style="font-size: 1.5rem; margin-bottom: 8px;">{t["feat_1_title"]}</div>
            <div style="font-size: 0.85rem; color: #94a3b8; line-height: 1.4;">{t["feat_1_desc"]}</div>
        </div>
        """, unsafe_allow_html=True)
    with col_feat2:
        st.markdown(f"""
        <div class="metric-card" style="min-height: 150px; text-align: left; padding: 20px;">
            <div style="font-size: 1.5rem; margin-bottom: 8px;">{t["feat_2_title"]}</div>
            <div style="font-size: 0.85rem; color: #94a3b8; line-height: 1.4;">{t["feat_2_desc"]}</div>
        </div>
        """, unsafe_allow_html=True)
    with col_feat3:
        st.markdown(f"""
        <div class="metric-card" style="min-height: 150px; text-align: left; padding: 20px;">
            <div style="font-size: 1.5rem; margin-bottom: 8px;">{t["feat_3_title"]}</div>
            <div style="font-size: 0.85rem; color: #94a3b8; line-height: 1.4;">{t["feat_3_desc"]}</div>
        </div>
        """, unsafe_allow_html=True)

# ----------------- UYGULAMA ALT BİLGİSİ -----------------
st.markdown(f'<div class="footer">{t["footer_text"]}</div>', unsafe_allow_html=True)
