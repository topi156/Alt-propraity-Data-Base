import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import io
from pptx import Presentation
from pptx.util import Inches

# --- הגדרות תצוגה בסיסיות ---
st.set_page_config(page_title="PE Fund Analytics Dashboard", layout="wide", page_icon="📈")

st.sidebar.title("הגדרות וניווט")
uploaded_file = st.sidebar.file_uploader("העלה את קובץ הנתונים (XLSX)", type=["xlsx"])

# --- פונקציית טעינת נתונים ---
@st.cache_data
def load_excel_data(file):
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names
    
    combined_sheet = next((s for s in sheet_names if 'Combined' in s), None)
    cleaned_sheet = next((s for s in sheet_names if 'Cleaned' in s), None)
    
    df_combined = pd.DataFrame()
    df_cleaned = pd.DataFrame()
    
    if combined_sheet:
        df_combined = pd.read_excel(xls, sheet_name=combined_sheet)
        if 'Date of Investment' in df_combined.columns and 'Exit/Current Date' in df_combined.columns:
            df_combined['Date of Investment'] = pd.to_datetime(df_combined['Date of Investment'], errors='coerce')
            df_combined['Exit/Current Date'] = pd.to_datetime(df_combined['Exit/Current Date'], errors='coerce')
            df_combined['Hold Period (Years)'] = (df_combined['Exit/Current Date'] - df_combined['Date of Investment']).dt.days / 365.25
            df_combined['Vintage Year'] = df_combined['Date of Investment'].dt.year
            
        if 'Gross Multiple' in df_combined.columns:
            df_combined['Gross Multiple'] = pd.to_numeric(df_combined['Gross Multiple'], errors='coerce')
            
    if cleaned_sheet:
        df_cleaned = pd.read_excel(xls, sheet_name=cleaned_sheet)
        
    return df_combined, df_cleaned, sheet_names

# ==========================================
# פונקציה ליצירת מצגת PowerPoint
# ==========================================
def create_ppt_report(fund_name, deals_count, fig_waterfall):
    prs = Presentation()
    
    # שקף 1: שקף כותרת
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = f"Tear Sheet: {fund_name}"
    subtitle.text = f"מבוסס על ניתוח של {deals_count} עסקאות\nנוצר אוטומטית ממערכת האנליזה"
    
    # שקף 2: גשר יצירת הערך (גרף)
    if fig_waterfall:
        blank_slide_layout = prs.slide_layouts[5] # שקף עם כותרת בלבד
        slide2 = prs.slides.add_slide(blank_slide_layout)
        slide2.shapes.title.text = "גשר יצירת ערך - ממוצע לעסקה בקרן"
        
        # המרת הגרף של Plotly לתמונה בזיכרון
        img_bytes = io.BytesIO()
        fig_waterfall.write_image(img_bytes, format='png', width=800, height=450)
        img_bytes.seek(0)
        
        # הדבקת התמונה בשקף
        slide2.shapes.add_picture(img_bytes, Inches(1), Inches(2), width=Inches(8))
    
    # שמירת המצגת לזיכרון
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# ==========================================
# לוגיקת המסכים
# ==========================================
if uploaded_file is not None:
    df, df_clean, all_sheets = load_excel_data(uploaded_file)
    
    if not df.empty:
        page = st.sidebar.radio("בחר מסך ניתוח:", 
                                ["מבט על (Macro Overview)", 
                                 "מגמות ומפות חום (Trends & Maps)",
                                 "פרופיל קרן (GP Tear Sheet) 📊", 
                                 "מגרש המשחקים (Deal Explorer)",
                                 "תשאול חופשי (AI Chat)"])
        
        st.sidebar.markdown("---")

        # ------------------------------------------------
        # מסך 1 ו-2 (מבט על ומגמות - נשאר ללא שינוי מהקוד הקודם)
        # ------------------------------------------------
        if page == "מבט על (Macro Overview)":
            st.title("מבט על - ביצועי תעשיית ה-PE")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("סה\"כ עסקאות", f"{len(df):,}")
            col2.metric("מכפיל ממוצע", f"{df['Gross Multiple'].mean():.2f}x")
            col3.metric("מספר מנהלי קרנות", f"{df['Fund Name'].nunique() if 'Fund Name' in df.columns else 0}")
            col4.metric("זמן החזקה ממוצע", f"{df['Hold Period (Years)'].mean():.1f} שנים")
            
            if 'Gross Multiple' in df.columns and 'Fund Name' in df.columns:
                st.subheader("דירוג מנהלים: תשואה לעומת יחס הפסד")
                fund_stats = df.dropna(subset=['Gross Multiple', 'Fund Name']).groupby('Fund Name').agg(
                    Avg_MOIC=('Gross Multiple', 'mean'), Deal_Count=('Gross Multiple', 'count'),
                    Loss_Ratio=('Gross Multiple', lambda x: (x < 1.0).mean() * 100)
                ).reset_index()
                fund_stats = fund_stats[fund_stats['Deal_Count'] > 3] 
                fig_scatter = px.scatter(fund_stats, x="Loss_Ratio", y="Avg_MOIC", size="Deal_Count", color="Fund Name", hover_name="Fund Name")
                fig_scatter.add_hline(y=1.0, line_dash="dot", line_color="red")
                st.plotly_chart(fig_scatter, use_container_width=True)

        elif page == "מגמות ומפות חום (Trends & Maps)":
            st.title("מגמות מאקרו-כלכליות")
            if 'Vintage Year' in df.columns and 'Industry' in df.columns:
                valid_years = df[(df['Vintage Year'] >= 1990) & (df['Vintage Year'] <= 2024)].copy()
                heatmap_data = valid_years.groupby(['Vintage Year', 'Industry'])['Gross Multiple'].mean().reset_index()
                top_industries = valid_years['Industry'].value_counts().nlargest(10).index
                heatmap_data = heatmap_data[heatmap_data['Industry'].isin(top_industries)]
                fig_heat = px.density_heatmap(heatmap_data, x="Vintage Year", y="Industry", z="Gross Multiple", histfunc="avg", color_continuous_scale="Viridis")
                st.plotly_chart(fig_heat, use_container_width=True)

        # ------------------------------------------------
        # מסך 3: פרופיל קרן (כולל ייצוא ל-PPTX)
        # ------------------------------------------------
        elif page == "פרופיל קרן (GP Tear Sheet) 📊":
            st.title("פרופיל מנהל קרן ויצירת ערך")
            if not df_clean.empty and 'Fund Name' in df_clean.columns:
                funds_list = sorted(df_clean['Fund Name'].dropna().unique())
                selected_fund = st.selectbox("בחר מנהל קרן לניתוח:", funds_list)
                fund_data = df_clean[df_clean['Fund Name'] == selected_fund]
                
                st.write(f"### מנתח נתונים עבור: **{selected_fund}** (סה\"כ {len(fund_data)} עסקאות)")
                
                fig_waterfall = None
                vc_cols = ['Revenue_MOIC', 'Margin_MOIC', 'Multiple_MOIC', 'Net_Debt_MOIC']
                if all(c in fund_data.columns for c in vc_cols):
                    fund_vc_data = fund_data.dropna(subset=vc_cols)
                    if len(fund_vc_data) > 0:
                        st.subheader("גשר יצירת ערך (Value Creation Bridge)")
                        avg_entry = 1.0
                        avg_rev = fund_vc_data['Revenue_MOIC'].mean()
                        avg_mar = fund_vc_data['Margin_MOIC'].mean()
                        avg_mul = fund_vc_data['Multiple_MOIC'].mean()
                        avg_debt = fund_vc_data['Net_Debt_MOIC'].mean()
                        avg_total = avg_entry + avg_rev + avg_mar + avg_mul + avg_debt
                        
                        fig_waterfall = go.Figure(go.Waterfall(
                            orientation="v", measure=["absolute", "relative", "relative", "relative", "relative", "total"],
                            x=["Entry", "Revenue Growth", "Margin Expansion", "Multiple Arbitrage", "Debt/Cash Flow", "Current/Exit"],
                            text=[f"{avg_entry:.2f}x", f"{avg_rev:+.2f}x", f"{avg_mar:+.2f}x", f"{avg_mul:+.2f}x", f"{avg_debt:+.2f}x", f"{avg_total:.2f}x"],
                            y=[avg_entry, avg_rev, avg_mar, avg_mul, avg_debt, avg_total],
                            connector={"line":{"color":"rgb(63, 63, 63)"}}
                        ))
                        st.plotly_chart(fig_waterfall, use_container_width=True)
                
                # --- אזור הייצוא ל-PowerPoint ---
                st.markdown("---")
                st.subheader("📥 ייצוא דוח להנהלה")
                st.write("לחיצה על הכפתור תייצר עבורך מצגת מנהלים עם הנתונים והגרפים של הקרן שנבחרה.")
                
                if st.button("הכן דוח PowerPoint"):
                    with st.spinner("בונה מצגת..."):
                        ppt_file = create_ppt_report(selected_fund, len(fund_data), fig_waterfall)
                        
                        st.success("המצגת מוכנה!")
                        st.download_button(
                            label="⬇️ הורד את המצגת למחשב (PPTX)",
                            data=ppt_file,
                            file_name=f"{selected_fund}_Tear_Sheet.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                
                st.markdown("---")
                st.subheader("חברות בולטות בקרן")
                cols_to_show = [c for c in ['Portfolio Company', 'Industry', 'Gross Multiple'] if c in df.columns]
                st.dataframe(df[df['Fund Name'] == selected_fund][cols_to_show].sort_values(by='Gross Multiple', ascending=False), use_container_width=True)

        # ------------------------------------------------
        # מסכים נוספים
        # ------------------------------------------------
        elif page == "מגרש המשחקים (Deal Explorer)":
             st.title("מגרש המשחקים")
             st.write("אזור לחקירת עסקאות ספציפיות.")
             # (קוד קודם של מגרש המשחקים)

        elif page == "תשאול חופשי (AI Chat)":
             st.title("🤖 מנוע AI")
             st.write("אזור צ'אט.")
             # (קוד קודם של הצ'אט)

else:
    st.info("אנא העלה את קובץ האקסל בתפריט הצד שמאל כדי להתחיל.")
