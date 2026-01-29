import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- הגדרות עמוד (כותרת, אייקון, פריסה) ---
st.set_page_config(page_title="מחולל אקסל מתקדם", page_icon="📊", layout="wide")

# --- פונקציות עזר ---
def init_session_state():
    """מאתחל את מבנה הנתונים בזיכרון הדפדפן אם הוא לא קיים"""
    if 'df' not in st.session_state:
        # טבלה התחלתית ריקה
        st.session_state.df = pd.DataFrame(columns=["שם פריט", "כמות", "מחיר", "תאריך"])

def convert_df_to_excel(df):
    """ממיר את הטבלה לקובץ אקסל בזיכרון (ללא שמירה בדיסק)"""
    output = io.BytesIO()
    # שימוש ב-xlsxwriter לעיצוב מתקדם
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        # גישה לאובייקט ה-workbook וה-worksheet לעיצוב
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # הגדרת עיצוב לכותרות (מודגש, רקע תכלת, גבולות)
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        
        # החלת העיצוב על השורה הראשונה
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
    processed_data = output.getvalue()
    return processed_data

# --- גוף האפליקציה ---
def main():
    init_session_state()

    st.title("📊 מחולל טבלאות אקסל לאתרי אינטרנט")
    st.markdown("כלי זה מאפשר לבנות טבלאות, לערוך אותן בזמן אמת ולהוריד אותן כאקסל מעוצב.")

    # --- סרגל צד: הגדרות קובץ ---
    with st.sidebar:
        st.header("הגדרות ייצוא")
        file_name_input = st.text_input("שם הקובץ לשמירה:", value="הטבלה_שלי")
        if not file_name_input.endswith(".xlsx"):
            file_name_input += ".xlsx"
        
        st.divider()
        st.write("### ניהול עמודות")
        new_col = st.text_input("הוסף עמודה חדשה:")
        if st.button("הוסף עמודה"):
            if new_col and new_col not in st.session_state.df.columns:
                st.session_state.df[new_col] = ""
                st.success(f"עמודה '{new_col}' נוספה!")
                st.rerun()

        if st.button("נקה את כל הטבלה", type="primary"):
            st.session_state.df = pd.DataFrame(columns=["עמודה 1"])
            st.rerun()

    # --- אזור מרכזי: עריכת הנתונים ---
    st.subheader("עריכת הנתונים (ממשק חי)")
    
    # רכיב data_editor מאפשר עריכה כמו באקסל בתוך הדפדפן
    edited_df = st.data_editor(
        st.session_state.df,
        num_rows="dynamic", # מאפשר למשתמש להוסיף שורות לבד
        use_container_width=True,
        key="editor"
    )

    # עדכון ה-State עם השינויים שהמשתמש עשה
    if not edited_df.equals(st.session_state.df):
        st.session_state.df = edited_df

    st.divider()

    # --- אזור הורדה ---
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.info(f"הקובץ יישמר בשם: **{file_name_input}**")
        
        # המרה לאקסל
        excel_data = convert_df_to_excel(edited_df)
        
        # כפתור ההורדה
        st.download_button(
            label="📥 הורד קובץ Excel מוכן",
            data=excel_data,
            file_name=file_name_input,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        # סטטיסטיקה מהירה
        st.metric("מספר שורות", edited_df.shape[0])
        st.metric("מספר עמודות", edited_df.shape[1])

if __name__ == "__main__":
    main()
