
#========================================================
# streamlit pandas openpyxl plotly jdatetime
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import re
import jdatetime


#============= Page Config ==============================
st.set_page_config(
    page_title="داشبورد مدیریت تضمین‌ها",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @font-face {
        font-family: 'BNazanin';
        src: url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.eot');
        src: url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.eot?#iefix') format('embedded-opentype'),
             url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.woff2') format('woff2'),
             url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.woff') format('woff'),
             url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.ttf') format('truetype');
        font-weight: normal;
        font-style: normal;
    }
    
    @font-face {
        font-family: 'BNazanin';
        src: url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin-Bold.eot');
        src: url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin-Bold.eot?#iefix') format('embedded-opentype'),
             url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin-Bold.woff2') format('woff2'),
             url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin-Bold.woff') format('woff'),
             url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin-Bold.ttf') format('truetype');
        font-weight: bold;
        font-style: normal;
    }
    
    html, body, .stApp, .main {
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
    }
    
    .main {
        background-color: #f8f9fa;
    }
    
    .stMarkdown, .stText, .stTitle, .stSubheader, .stCaption, .stLabel {
        font-weight: bold !important;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
    }
    
    .main-header {
        background: linear-gradient(135deg, #1e3a6f 0%, #2a5298 100%);
        padding: 1.5rem 2rem;
        border-radius: 15px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.15);
    }
    
    .main-header h1 {
        color: white !important;
        font-size: 2.5rem !important;
        font-weight: 900 !important;
        margin: 0 !important;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
    }
    
    .main-header p {
        color: rgba(255,255,255,0.90) !important;
        font-size: 1.2rem !important;
        font-weight: bold !important;
        margin: 0 !important;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
    }
    
    .metric-card {
        background: white;
        padding: 1.2rem 1.5rem;
        border-radius: 15px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        border-right: 5px solid #2a5298;
        transition: all 0.3s ease;
        height: 100%;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.12);
    }
    
    .metric-card .icon {
        font-size: 2rem;
        margin-bottom: 0.5rem;
    }
    
    .metric-card .value {
        font-size: 1.8rem;
        font-weight: 900 !important;
        color: #1e3a6f;
        direction: ltr;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
    }
    
    .metric-card .label {
        font-size: 0.95rem;
        color: #2c3e50;
        font-weight: 700 !important;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
    }
    
    .metric-card .delta {
        font-size: 0.85rem;
        font-weight: 700 !important;
        padding: 0.2rem 0.8rem;
        border-radius: 20px;
        display: inline-block;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
    }
    
    .delta-danger {
        background: #fee2e2;
        color: #dc2626;
    }
    
    .delta-success {background: #dcfce7;
        color: #16a34a;
    }
    
    .metric-blue { border-right-color: #2a5298; }
    .metric-green { border-right-color: #16a34a; }
    .metric-orange { border-right-color: #f59e0b; }
    .metric-red { border-right-color: #dc2626; }
    .metric-purple { border-right-color: #7c3aed; }
    
    .stSidebar .stMarkdown, .stSidebar .stText, .stSidebar label {
        font-weight: bold !important;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
    }
    
    .stDataFrame, .stDataFrame table, .stDataFrame thead, .stDataFrame tbody, .stDataFrame th, .stDataFrame td {
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
        font-weight: bold !important;
    }
    
    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 2px 10px rgba(0,0,0,0.06);
    }
    
    .stButton button {
        background: linear-gradient(135deg, #1e3a6f 0%, #2a5298 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.6rem 1.5rem;
        font-weight: 700 !important;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
        transition: all 0.3s ease;
        font-size: 1rem;
    }
    
    .stButton button:hover {
        transform: scale(1.03);
        box-shadow: 0 4px 15px rgba(42, 82, 152, 0.4);
    }
    
    .warning-box {
        background: #fef3c7;
        border-right: 5px solid #f59e0b;
        padding: 1rem 1.5rem;
        border-radius: 12px;
        margin: 0.5rem 0;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
        font-weight: bold !important;
    }
    
    .success-box {
        background: #dcfce7;
        border-right: 5px solid #16a34a;
        padding: 1rem 1.5rem;
        border-radius: 12px;
        margin: 0.5rem 0;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
        font-weight: bold !important;
    }
    
    .section-title {
        font-size: 1.5rem;
        font-weight: 900 !important;
        color: #1e3a6f;
        margin: 1.5rem 0 1rem 0;
        padding-bottom: 0.5rem;
        border-bottom: 4px solid #2a5298;
        display: inline-block;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
    }
    
    .stTextInput input {
        border-radius: 10px !important;
        border: 2px solid #e5e7eb !important;
        padding: 0.6rem 1rem !important;
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
        font-weight: bold !important;
    }
    
    .stSelectbox, .stRadio, .stSelectbox label, .stRadio label {
        font-family: 'BNazanin', 'Vazir', 'Tahoma', sans-serif !important;
        font-weight: bold !important;
    }
</style>
""", unsafe_allow_html=True)

#============== Header ===================================

st.markdown("""
<div class="main-header" style="text-align: center;">
    <h1>📋 داشبورد مدیریت ضمانتنامه‌ها</h1>
    <p>نمایش و تحلیل وضعیت تضمین‌های پیمانکاران</p>
</div>
""", unsafe_allow_html=True)

#============== File Loading ==============================

file_path = "tazmin.xlsx"

@st.cache_data
def load_data():
    try:
        df = pd.read_excel(file_path)
        if df.empty:
            st.error("⚠️ فایل پیدا شد اما خالی است!")
            return None
        return df
    except FileNotFoundError:
        st.error(f"❌ فایل '{file_path}' پیدا نشد!")
        return None
    except Exception as e:
        st.error(f"❌ خطا در خواندن فایل: {e}")
        return None

df = load_data()
if df is None:
    st.stop()

#============== Persian to Gregorgian Data Conversion =====

def jalali_to_gregorian(jy, jm, jd):
    jalali_days = [31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29]
    day_of_year = sum(jalali_days[:jm-1]) + jd
    leap_years = (jy - 1) // 4
    total_days = (jy - 1) * 365 + leap_years + day_of_year - 1
    base_date = datetime(622, 3, 22)
    result_date = base_date + timedelta(days=total_days)
    return result_date.date()

def convert_persian_date(date_str):
    if pd.isna(date_str):
        return None
    try:
        if isinstance(date_str, (int, float)):
            date_str = str(int(date_str))
        
        date_str = str(date_str).strip()
        date_str = date_str.replace('-', ' ').replace('/', ' ')
        date_str = re.sub(r'\s+', ' ', date_str).strip()
        
        parts = date_str.split()
        if len(parts) >= 3:
            year = int(parts[0])
            month = int(parts[1])
            day = int(parts[2])
            
            if year < 1300 or year > 1500 or month < 1 or month > 12 or day < 1 or day > 31:
                return None
            
            # تبدیل با jdatetime
            jalali_date = jdatetime.date(year, month, day)
            gregorian_date = jalali_date.togregorian()
            return gregorian_date
        
        return None
    except Exception as e:
        return None
    
#============== Sidebar ==================================

with st.sidebar:
    st.markdown("## 🔍 فیلترها")
    st.markdown("---")
    
    # نمایش ستون‌ها
    with st.expander("📌 ساختار داده", expanded=False):
        st.write(df.columns.tolist())
    
    # تبدیل تاریخ‌ها
    if 'تاریخ سر رسید' in df.columns:
        df['تاریخ سررسید میلادی'] = df['تاریخ سر رسید'].apply(convert_persian_date)
        converted_count = df['تاریخ سررسید میلادی'].notna().sum()
        st.success(f"✅ {converted_count} تاریخ با موفقیت تبدیل شد")
    
    if 'تاریخ صدور' in df.columns:
        df['تاریخ صدور میلادی'] = df['تاریخ صدور'].apply(convert_persian_date)
    
    st.markdown("---")
    
    # فیلتر پیمانکار
    if 'نام پیمانکار' in df.columns:
        contractors = ['همه'] + sorted(df['نام پیمانکار'].dropna().unique().tolist())
        selected_contractor = st.selectbox("🏢 انتخاب پیمانکار", contractors)
        if selected_contractor != 'همه':
            df = df[df['نام پیمانکار'] == selected_contractor]
    
    # فیلتر نوع تضمین
    if 'نوع اسناد تضمینی' in df.columns:
        types = ['همه'] + sorted(df['نوع اسناد تضمینی'].dropna().unique().tolist())
        selected_type = st.selectbox("📄 نوع تضمین", types)
        if selected_type != 'همه':
            df = df[df['نوع اسناد تضمینی'] == selected_type]
    
    # فیلتر وضعیت
    status_filter = st.radio("📊 وضعیت", ["همه", "فعال ✅", "سررسید شده ⚠️"])
    
    if 'تاریخ سررسید میلادی' in df.columns:
        today = datetime.now().date()
        if status_filter == "فعال ✅":
            df = df[df['تاریخ سررسید میلادی'] >= today]
        elif status_filter == "سررسید شده ⚠️":
            df = df[df['تاریخ سررسید میلادی'] < today]
    
    st.markdown("---")
    
    # دانلود
    st.markdown("## 📥 خروجی")
    if st.button("📊 دانلود گزارش CSV", width='stretch'):
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="✅ دانلود فایل",
            data=csv,
            file_name="گزارش_تضمین‌ها.csv",
            mime="text/csv",
            width='stretch'
        )

#============== Metric Cards ==================================

st.markdown('<p class="section-title">📊 خلاصه وضعیت</p>', unsafe_allow_html=True)

col1, col2, col3, col4, col5 = st.columns(5)

total_count = len(df)

expired_count = 0
near_expiry_count = 0
if 'تاریخ سررسید میلادی' in df.columns:
    today = datetime.now().date()
    expired_count = df[df['تاریخ سررسید میلادی'] < today].shape[0]
    
    threshold = today + timedelta(days=30)
    near_expiry_count = df[
        (df['تاریخ سررسید میلادی'] >= today) & 
        (df['تاریخ سررسید میلادی'] <= threshold)
    ].shape[0]

active_count = total_count - expired_count

returned_count = 0
if 'ملاحظات' in df.columns:
    returned_count = df[df['ملاحظات'].str.contains('عودت', na=False)].shape[0]

with col1:
    st.markdown(f"""
    <div class="metric-card metric-blue">
        <div class="icon">📄</div>
        <div class="value">{total_count:,}</div>
        <div class="label">تعداد کل تضمین‌ها</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="metric-card metric-blue">
        <div class="icon">✅</div>
        <div class="value">{active_count:,}</div>
        <div class="label">تضمین‌های فعال</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    delta_class = "delta-danger" if near_expiry_count > 0 else "delta-success"
    st.markdown(f"""
    <div class="metric-card metric-orange">
        <div class="icon">🔔</div>
        <div class="value">{near_expiry_count:,}</div>
        <div class="label">سررسیدهای نزدیک (۳۰ روز)</div>
        <div class="delta {delta_class}">{'⚠️ نیاز به بررسی' if near_expiry_count > 0 else '✅ بدون مشکل'}</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    delta_class = "delta-danger" if expired_count > 0 else "delta-success"
    st.markdown(f"""
    <div class="metric-card metric-red">
        <div class="icon">⏰</div>
        <div class="value">{expired_count:,}</div>
        <div class="label">تضمین‌های سررسید شده</div>
    </div>
    """, unsafe_allow_html=True)

with col5:
    st.markdown(f"""
    <div class="metric-card metric-purple">
        <div class="icon">🔄</div>
        <div class="value">{returned_count:,}</div>
        <div class="label">عودت داده شده</div>
    </div>
    """, unsafe_allow_html=True)

#============== Charts ======================================

st.markdown('<p class="section-title">📊 تحلیل‌های بصری</p>', unsafe_allow_html=True)

# نمودار ۱: تعداد تضمین به تفکیک پیمانکار (به جای مبلغ)
if 'نام پیمانکار' in df.columns:
    contractor_count = df['نام پیمانکار'].value_counts().head(10)
    if not contractor_count.empty:
        fig1 = px.bar(
            contractor_count,
            title="🔝 ۱۰ پیمانکار برتر بر اساس تعداد تضمین",
            labels={'value': 'تعداد تضمین', 'نام پیمانکار': 'پیمانکار'},
            color=contractor_count.values,
            color_continuous_scale="Blues",
            text_auto=True
        )
        fig1.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="BNazanin, Vazir, Tahoma, sans-serif"),
            title_font=dict(size=18, color='#1e3a6f'),
            height=500,
            showlegend=False,
            margin=dict(l=40, r=40, t=60, b=40)
        )
        fig1.update_traces(marker_line_width=0, textfont_size=12)
        st.plotly_chart(fig1, width='stretch')

col_chart2, col_chart3 = st.columns(2)

with col_chart2:
    if 'نوع اسناد تضمینی' in df.columns:
        type_data = df['نوع اسناد تضمینی'].value_counts()
        if not type_data.empty:
            fig2 = px.pie(
                type_data,
                values=type_data.values,
                names=type_data.index,
                title="📌 توزیع نوع اسناد تضمینی",
                color_discrete_sequence=['#1e3a6f', '#2a5298', '#3b6fc0', '#4a8ad4', '#6aa6e0', '#8fc2ec'],
                hole=0.3
            )
            fig2.update_traces(textposition='inside', textinfo='percent+label')
            fig2.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family="BNazanin, Vazir, Tahoma, sans-serif"),
                title_font=dict(size=16, color='#1e3a6f'),
                height=400,
                showlegend=False
            )
            st.plotly_chart(fig2, width='stretch')

with col_chart3:
    if 'تاریخ صدور میلادی' in df.columns:
        df_temp = df[df['تاریخ صدور میلادی'].notna()].copy()
        if not df_temp.empty:
            try:
                df_temp['تاریخ صدور میلادی'] = pd.to_datetime(df_temp['تاریخ صدور میلادی'])
                df_temp['ماه صدور'] = df_temp['تاریخ صدور میلادی'].dt.to_period('M')
                monthly_data = df_temp.groupby('ماه صدور').size().reset_index()
                monthly_data['ماه صدور'] = monthly_data['ماه صدور'].astype(str)
                
                if len(monthly_data) > 1:
                    fig3 = px.line(
                        monthly_data,
                        x='ماه صدور',
                        y=0,
                        title="📈 روند صدور تضمین‌ها",
                        labels={'value': 'تعداد', 'month': 'ماه'},
                        markers=True,
                        color_discrete_sequence=['#2a5298']
                    )
                    fig3.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="BNazanin, Vazir, Tahoma, sans-serif"),
                        title_font=dict(size=16, color='#1e3a6f'),
                        height=400
                    )
                    fig3.update_traces(line=dict(width=3), marker=dict(size=10))
                    st.plotly_chart(fig3, width='stretch')
                else:
                    st.info("📊 داده‌های کافی برای رسم نمودار روند وجود ندارد.")
            except Exception as e:
                st.warning(f"⚠️ خطا در رسم نمودار روند: {e}")

#============== Expiry Alerts ======================================

st.markdown('<p class="section-title">⏰ هشدار سررسیدهای نزدیک</p>', unsafe_allow_html=True)

if 'تاریخ سررسید میلادی' in df.columns:
    today = datetime.now().date()
    threshold = today + timedelta(days=30)
    
    expiring_soon = df[
        (df['تاریخ سررسید میلادی'] >= today) & 
        (df['تاریخ سررسید میلادی'] <= threshold)
    ]
    
    if len(expiring_soon) > 0:
        st.markdown(f"""
        <div class="warning-box">
            ⚠️ <strong>{len(expiring_soon)} تضمین</strong> در ۳۰ روز آینده سررسید می‌شوند.
        </div>
        """, unsafe_allow_html=True)
        cols_to_show = ['نام پیمانکار', 'شرح قرارداد', 'تاریخ سر رسید']
        available_cols = [col for col in cols_to_show if col in df.columns]
        st.dataframe(expiring_soon[available_cols], width='stretch', height=200)
    else:
        st.markdown("""
        <div class="success-box">
            ✅ هیچ تضمینی در ۳۰ روز آینده سررسید نمی‌شود.
        </div>
        """, unsafe_allow_html=True)

#============== Data Table ======================================

st.markdown('<p class="section-title">📋 لیست کامل تضمین‌ها</p>', unsafe_allow_html=True)

search_term = st.text_input("🔎 جستجو در جدول (نام پیمانکار، شرح قرارداد، ...)", placeholder="متن مورد نظر را وارد کنید...")

display_cols = ['کارشناس مربوطه', 'نام پیمانکار', 'شرح قرارداد', 'نوع اسناد تضمینی', 
                'تاریخ صدور', 'تاریخ سر رسید', 'ملاحظات']

available_cols = [col for col in display_cols if col in df.columns]

if search_term:
    mask = df[available_cols].astype(str).apply(lambda x: x.str.contains(search_term, na=False)).any(axis=1)
    filtered_df = df[mask]
    st.caption(f"🔍 {len(filtered_df)} نتیجه برای جستجوی '{search_term}'")
else:
    filtered_df = df

st.dataframe(
    filtered_df[available_cols],
    width='stretch',
    height=400
)

#============== Footer =========================================
st.markdown("---")
st.caption("📋 داشبورد مدیریت ضمانتنامه‌ها | آخرین به‌روزرسانی: " + datetime.now().strftime("%Y-%m-%d %H:%M"))
