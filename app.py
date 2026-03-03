import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. הגדרות תצוגה בסיסיות ל-Streamlit ---
st.set_page_config(page_title="PE Fund Analytics Dashboard", layout="wide")

# --- 2. תפריט צד (Sidebar) והעלאת קבצים ---
st.sidebar.title("הגדרות וניווט")
uploaded_file = st.sidebar.file_uploader("העלה את קובץ הנתונים (Excel)", type=["xlsx"])

# --- 3. פונקציה לטעינת נתוני האקסל עם זיכרון מטמון (Cache) לביצועים מהירים ---
@st.cache_data
def load_excel_data(file):
    # טעינת קובץ האקסל השלם לזיכרון
    xls = pd.ExcelFile(file)
    
    # ניסיון לשלוף את הגיליונות הספציפיים שזיהינו כמאגר המאוחד והמנוקה
    try:
        df_combined = pd.read_excel(xls, sheet_name='Combined Proprietary Database')
        df_cleaned = pd.read_excel(xls, sheet_name='Cleaned Proprietary Database')
    except ValueError:
        st.error("שגיאה: לא נמצאו גיליונות בשם 'Combined Proprietary Database' או 'Cleaned Proprietary Database'. אנא ודא שזהו הקובץ הנכון.")
        return pd.DataFrame(), pd.DataFrame(), []

    # עיבוד נתונים בסיסי לגיליון המאוחד (תיקון תאריכים וחישוב זמני החזקה)
    if 'Date of Investment' in df_combined.columns and 'Exit/Current Date' in df_combined.columns:
        df_combined['Date of Investment'] = pd.to_datetime(df_combined['Date of Investment'], errors='coerce')
        df_combined['Exit/Current Date'] = pd.to_datetime(df_combined['Exit/Current Date'], errors='coerce')
        # חישוב משך החזקה בשנים
        df_combined['Hold Period (Years)'] = (df_combined['Exit/Current Date'] - df_combined['Date of Investment']).dt.days / 365.25
    
    if 'Gross Multiple' in df_combined.columns:
        df_combined['Gross Multiple'] = pd.to_numeric(df_combined['Gross Multiple'], errors='coerce')
        
    return df_combined, df_cleaned, xls.sheet_names

# ==========================================
# 4. לוגיקת המערכת והמסכים (רק אם הועלה קובץ)
# ==========================================
if uploaded_file is not None:
    # קריאה לפונקציית הטעינה
    df, df_clean, all_sheets = load_excel_data(uploaded_file)
    
    if not df.empty:
        # תפריט ניווט במסכים
        page = st.sidebar.radio("בחר מסך ניתוח:", 
                                ["מבט על (Macro Overview)", 
                                 "פרופיל קרן (GP Tear Sheet)", 
                                 "מגרש המשחקים (Deal Explorer)"])
        
        st.sidebar.markdown("---")
        st.sidebar.success(f"בהצלחה! נטענו {len(all_sheets)} גיליונות מקובץ האקסל.")

        # ------------------------------------------------
        # מסך 1: מבט על (Macro Overview)
        # ------------------------------------------------
        if page == "מבט על (Macro Overview)":
            st.title("מבט על - ביצועי תעשיית ה-PE")
            
            # מדדי KPI
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("סה\"כ עסקאות בבסיס הנתונים", f"{len(df):,}")
            col2.metric("מכפיל (MOIC) ממוצע", f"{df['Gross Multiple'].mean():.2f}x")
            col3.metric("מספר מנהלי קרנות (GPs)", f"{df['Fund Name'].nunique()}")
            col4.metric("זמן החזקה ממוצע (שנים)", f"{df['Hold Period (Years)'].mean():.1f}")
            
            st.markdown("---")
            
            # גרף: ביצועים מול סיכון (Loss Ratio vs MOIC)
            st.subheader("דירוג מנהלים: תשואה לעומת יחס הפסד (Loss Ratio)")
            
            # אגרגציה של הנתונים ברמת הקרן
            fund_stats = df.dropna(subset=['Gross Multiple', 'Fund Name']).groupby('Fund Name').agg(
                Avg_MOIC=('Gross Multiple', 'mean'),
                Deal_Count=('Gross Multiple', 'count'),
                Loss_Ratio=('Gross Multiple', lambda x: (x < 1.0).mean() * 100)
            ).reset_index()
            fund_stats = fund_stats[fund_stats['Deal_Count'] > 5] # סינון קרנות עם מעט מדי עסקאות
            
            fig_scatter = px.scatter(fund_stats, x="Loss_Ratio", y="Avg_MOIC", size="Deal_Count", color="Fund Name",
                                     hover_name="Fund Name", size_max=40,
                                     labels={"Loss_Ratio": "יחס הפסד - % העסקאות מתחת ל-1.0x", "Avg_MOIC": "מכפיל ממוצע (MOIC)"},
                                     title="GP Benchmarking (גודל הבועה = מספר עסקאות בקרן)")
            fig_scatter.add_hline(y=1.0, line_dash="dot", line_color="red", annotation_text="Break-even (1.0x)", annotation_position="bottom right")
            fig_scatter.update_layout(showlegend=False) # הסתרת מקרא ארוך מדי
            st.plotly_chart(fig_scatter, use_container_width=True)

        # ------------------------------------------------
        # מסך 2: פרופיל קרן וגשר יצירת ערך (GP Tear Sheet)
        # ------------------------------------------------
        elif page == "פרופיל קרן (GP Tear Sheet)":
            st.title("פרופיל מנהל קרן (GP Analysis)")
            
            # בחירת קרן
            funds_list = sorted(df_clean['Fund Name'].dropna().unique())
            selected_fund = st.selectbox("בחר מנהל קרן לניתוח:", funds_list)
            
            fund_data = df_clean[df_clean['Fund Name'] == selected_fund]
            st.write(f"### מנתח נתונים עבור: **{selected_fund}** (סה\"כ {len(fund_data)} עסקאות קיימות)")
            
            # בניית Value Bridge
            vc_cols = ['Revenue_MOIC', 'Margin_MOIC', 'Multiple_MOIC', 'Net_Debt_MOIC']
            fund_vc_data = fund_data.dropna(subset=vc_cols)
            
            if len(fund_vc_data) > 0:
                st.subheader("גשר יצירת ערך (Value Creation Bridge) - ממוצע לעסקה בקרן")
                
                avg_entry = 1.0
                avg_rev = fund_vc_data['Revenue_MOIC'].mean()
                avg_mar = fund_vc_data['Margin_MOIC'].mean()
                avg_mul = fund_vc_data['Multiple_MOIC'].mean()
                avg_debt = fund_vc_data['Net_Debt_MOIC'].mean()
                avg_total = avg_entry + avg_rev + avg_mar + avg_mul + avg_debt
                
                fig_waterfall = go.Figure(go.Waterfall(
                    orientation="v",
                    measure=["absolute", "relative", "relative", "relative", "relative", "total"],
                    x=["Entry", "Revenue Growth", "Margin Expansion", "Multiple Arbitrage", "Debt/Cash Flow", "Current/Exit"],
                    text=[f"{avg_entry:.2f}x", f"{avg_rev:+.2f}x", f"{avg_mar:+.2f}x", f"{avg_mul:+.2f}x", f"{avg_debt:+.2f}x", f"{avg_total:.2f}x"],
                    y=[avg_entry, avg_rev, avg_mar, avg_mul, avg_debt, avg_total],
                    connector={"line":{"color":"rgb(63, 63, 63)"}},
                    increasing={"marker":{"color":"#2ca02c"}},
                    decreasing={"marker":{"color":"#d62728"}},
                    totals={"marker":{"color":"#1f77b4"}}
                ))
                fig_waterfall.update_layout(showlegend=False, plot_bgcolor='rgba(0,0,0,0)', title="פירוק הגורמים לתשואת הקרן")
                st.plotly_chart(fig_waterfall, use_container_width=True)
            else:
                st.warning("אין מספיק נתוני Value Bridge מלאים עבור מנהל קרן זה במסד הנתונים.")
                
            st.subheader("חברות פורטפוליו בולטות")
            display_cols = ['Portfolio Company', 'Industry', 'Date of Investment', 'Hold Period (Years)', 'Gross Multiple']
            # הצגת הנתונים כטבלה מסודרת מההחזר הגבוה לנמוך
            st.dataframe(df[df['Fund Name'] == selected_fund][display_cols].sort_values(by='Gross Multiple', ascending=False), use_container_width=True)

        # ------------------------------------------------
        # מסך 3: מגרש המשחקים (Deal Explorer)
        # ------------------------------------------------
        elif page == "מגרש המשחקים (Deal Explorer)":
            st.title("מגרש המשחקים - חקר עסקאות חופשי")
            
            # פילטרים בסיידבר
            st.sidebar.markdown("### סינון עסקאות:")
            selected_sector = st.sidebar.multiselect("סנן לפי סקטור (Industry):", options=df['Industry'].dropna().unique())
            
            # וידוא שאין ערכי NaN ששוברים את הסליידר
            min_val = float(df['Gross Multiple'].min()) if not df['Gross Multiple'].isna().all() else 0.0
            max_val = float(df['Gross Multiple'].max()) if not df['Gross Multiple'].isna().all() else 15.0
            min_moic, max_moic = st.sidebar.slider("סנן לפי טווח מכפיל (MOIC):", 0.0, min(max_val, 20.0), (0.0, 10.0))
            
            # הפעלת הפילטרים על ה-DataFrame
            filtered_df = df.copy()
            if selected_sector:
                filtered_df = filtered_df[filtered_df['Industry'].isin(selected_sector)]
            filtered_df = filtered_df[(filtered_df['Gross Multiple'] >= min_moic) & (filtered_df['Gross Multiple'] <= max_moic)]
            
            st.write(f"מציג **{len(filtered_df)}** עסקאות בהתאם לסינון שבחרת.")
            
            # גרף פיזור
            fig_deals = px.scatter(filtered_df, x="Hold Period (Years)", y="Gross Multiple", 
                                   color="Fund Type", hover_name="Portfolio Company",
                                   hover_data=["Fund Name", "Industry"],
                                   title="משך החזקה לעומת מכפיל (העבר את העכבר על הנקודות לפרטים)",
                                   labels={"Hold Period (Years)": "זמן החזקה (בשנים)", "Gross Multiple": "מכפיל (MOIC)"})
            fig_deals.add_hline(y=1.0, line_dash="dash", line_color="red", annotation_text="Break-even")
            st.plotly_chart(fig_deals, use_container_width=True)

else:
    # מצב המתנה כשאין קובץ במערכת
    st.info("👈 אנא העלה את קובץ האקסל המלא (Fund Proprietary Database.xlsx) בתפריט הצד שמאל כדי להתחיל בניתוח.")
    st.image("https://images.unsplash.com/photo-1551288049-bebda4e38f71?auto=format&fit=crop&w=800&q=80", width=600, caption="המערכת ממתינה לטעינת הנתונים...")