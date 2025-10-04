import streamlit as st
from fpdf import FPDF
import pandas as pd
from datetime import datetime
import os

# --- PDF Generation Function ---
def create_pdf(report_date, department, current_activities, reflections, next_activities):
    pdf = FPDF(format='A4')
    
    # Check if font files exist before adding them
    script_dir = os.path.dirname(__file__) if '__file__' in locals() else os.getcwd()
    font_path_regular = os.path.join(script_dir, 'BIZ-UDGothic-Regular.ttf')
    font_path_bold = os.path.join(script_dir, 'BIZ-UDGothic-Bold.ttf')

    # Fallback to a generic font if BIZ-UDGothic is not found
    try:
        pdf.add_font('BIZ-UDGothic', '', font_path_regular, uni=True)
        pdf.add_font('BIZ-UDGothic', 'B', font_path_bold, uni=True)
    except Exception as e:
        st.warning(f"フォントファイルの読み込みに失敗しました: {e}。PDFの日本語表示に問題がある可能性があります。Arialを代替フォントとして使用します。")
        pdf.add_font('BIZ-UDGothic', '', 'arial.ttf', uni=True)
        pdf.add_font('BIZ-UDGothic', 'B', 'arialbd.ttf', uni=True)


    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font('BIZ-UDGothic', '', 12)

    # Header
    pdf.set_y(15)
    pdf.set_x(10)
    pdf.multi_cell(0, 10, "***運営委員会にて提出をお願いします***", align='C', border=0)
    pdf.multi_cell(0, 10, "事業内容報告書", align='C', border=0)
    pdf.set_font('BIZ-UDGothic', '', 10)
    pdf.set_y(35)
    pdf.set_x(140)
    reiwa_year = report_date.year - 2018
    pdf.cell(0, 10, f"令和{reiwa_year}年 {report_date.month}月 {report_date.day}日", align='R', ln=1)

    # Department
    pdf.set_y(50)
    pdf.set_x(10)
    pdf.set_font('BIZ-UDGothic', '', 12)
    pdf.cell(0, 10, "学年", align='L', ln=1)
    pdf.set_font('BIZ-UDGothic', 'B', 14)
    pdf.set_fill_color(200, 200, 200)
    pdf.cell(pdf.get_string_width(department) + 4, 10, department, align='C', border=1, fill=True)
    pdf.ln(15)

    # Current Activity Report (仕様書3)
    pdf.set_font('BIZ-UDGothic', '', 12)
    pdf.set_fill_color(200, 200, 200)
    pdf.cell(30, 10, "日程", border=1, align='C', fill=True)
    pdf.cell(0, 10, "事業内容報告", border=1, align='C', fill=True, ln=1)
    pdf.set_fill_color(255, 255, 255)

    for _, row in current_activities.iterrows():
        date_str = str(row["日程"])
        content_str = str(row["事業内容報告"])

        pdf.set_x(10)
        start_y = pdf.get_y()
        pdf.multi_cell(30, 7, date_str, border='LR', align='C')
        end_y_date = pdf.get_y()
        
        pdf.set_xy(40, start_y)
        pdf.multi_cell(160, 7, content_str, border='R', align='L')
        
        max_y = max(end_y_date, pdf.get_y())
        pdf.line(10, max_y, 200, max_y)
        pdf.set_y(max_y)

    # Reflections and Challenges (仕様書4)
    pdf.ln(10)
    pdf.set_fill_color(200, 200, 200)
    pdf.cell(0, 10, "活動の反省と課題", border=1, align='C', fill=True, ln=1)
    pdf.set_fill_color(255, 255, 255)
    pdf.set_font('BIZ-UDGothic', '', 10)
    pdf.multi_cell(0, 7, "(次年度以降の改善材料になりますので詳細にお願いします)", border='LR', align='L', ln=1)
    
    if reflections:
        pdf.multi_cell(0, 7, reflections, border='LRB', align='L', ln=1)
    else:
        # If empty, ensure a substantial empty box is drawn.
        # Estimate ~5 lines of text height for the empty box if content is usually large.
        pdf.cell(0, 7 * 5, "", border='LRB', align='L', ln=1) 
        
    pdf.set_font('BIZ-UDGothic', '', 12)

    # Next Activities (仕様書5)
    pdf.ln(10)
    pdf.set_fill_color(200, 200, 200)
    pdf.cell(30, 10, "日程", border=1, align='C', fill=True)
    pdf.cell(0, 10, "次回運営委員会までの活動予定", border=1, align='C', fill=True, ln=1)
    pdf.set_fill_color(255, 255, 255)

    for _, row in next_activities.iterrows():
        date_str = str(row["日程"])
        content_str = str(row["次回運営委員会までの活動予定"])

        pdf.set_x(10)
        start_y = pdf.get_y()
        pdf.multi_cell(30, 7, date_str, border='LR', align='C')
        end_y_date = pdf.get_y()
        
        pdf.set_xy(40, start_y)
        pdf.multi_cell(160, 7, content_str, border='R', align='L')
        
        max_y = max(end_y_date, pdf.get_y())
        pdf.line(10, max_y, 200, max_y)
        pdf.set_y(max_y)

    pdf_output = pdf.output(dest='S').encode('latin-1')
    return pdf_output

# --- Streamlit App ---
st.set_page_config(layout="wide")
st.title("事業内容報告書 作成アプリ")

# Initialize session state for dynamic inputs if not already present
if 'num_current_activities' not in st.session_state:
    st.session_state.num_current_activities = 5 # Default to 5 rows for current activities
if 'num_next_activities' not in st.session_state:
    st.session_state.num_next_activities = 5 # Default to 5 rows for next activities

st.header("報告書入力フォーム")

with st.form("report_form"):
    st.subheader("1. 報告日")
    report_date = st.date_input("報告書の日付", datetime.now())

    st.subheader("2. 担当部署")
    departments = [
        "学年委員１年", "学年委員２年", "学年委員３年", "学年委員４年",
        "学年委員５年", "学年委員６年", "学年委員あゆみ", "広報部",
        "校外安全指導部", "教養部", "環境厚生部", "選考委員会", "育成会本部"
    ]
    department = st.selectbox("担当部署を選択してください", departments)

    st.subheader("3. 事業内容報告 (枠内上部)") # 仕様書3
    st.write("活動日程と内容をそれぞれ入力してください。")

    current_activity_data = []
    
    cols_header_current = st.columns([1, 3])
    with cols_header_current[0]:
        st.markdown("**日程**")
    with cols_header_current[1]:
        st.markdown("**事業内容報告**")

    for i in range(st.session_state.num_current_activities):
        cols = st.columns([1, 3])
        with cols[0]:
            date_input = st.text_input(f"日程 {i+1}", key=f"current_date_{i}", label_visibility="collapsed")
        with cols[1]:
            content_input = st.text_area(f"事業内容 {i+1}", key=f"current_content_{i}", height=50, label_visibility="collapsed")
        current_activity_data.append({"日程": date_input, "事業内容報告": content_input})

    if st.button("事業内容報告を追加", key="add_current_activity"):
        st.session_state.num_current_activities += 1
        st.rerun()

    current_activities_df = pd.DataFrame(current_activity_data).dropna(how='all')

    st.subheader("4. 活動の反省と課題") # 仕様書4
    # Set height for a larger text area, roughly corresponding to PDF's multi_cell height
    # A multi_cell with height 7 and 5 lines would be 35. So, use a slightly larger Streamlit height.
    reflections = st.text_area(
        "次年度以降の改善材料になりますので詳細にお願いします。",
        value=st.session_state.get('reflections', ''),
        key="reflections_input",
        height=200 # 大きめの入力欄
    )
    st.session_state.reflections = reflections

    st.subheader("5. 次回運営委員会までの活動予定 (枠内下部)") # 仕様書5
    st.write("次回までの活動日程と内容をそれぞれ入力してください。")

    next_activity_data = []

    cols_header_next = st.columns([1, 3])
    with cols_header_next[0]:
        st.markdown("**日程**")
    with cols_header_next[1]:
        st.markdown("**次回運営委員会までの活動予定**")

    for i in range(st.session_state.num_next_activities):
        cols = st.columns([1, 3])
        with cols[0]:
            date_input = st.text_input(f"日程 {i+1}", key=f"next_date_{i}", label_visibility="collapsed")
        with cols[1]:
            content_input = st.text_area(f"活動予定 {i+1}", key=f"next_content_{i}", height=50, label_visibility="collapsed")
        next_activity_data.append({"日程": date_input, "次回運営委員会までの活動予定": content_input})

    if st.button("活動予定を追加", key="add_next_activity"):
        st.session_state.num_next_activities += 1
        st.rerun()

    next_activities_df = pd.DataFrame(next_activity_data).dropna(how='all')

    submitted = st.form_submit_button("入力完了 - プレビュー表示 (6. 入力完了ボタン)") # 仕様書6

if submitted:
    st.subheader("完成版プレビュー")

    # Display a basic HTML/Markdown preview
    st.markdown(f"<div style='text-align: center; font-weight: bold;'>***運営委員会にて提出をお願いします***</div>", unsafe_allow_html=True)
    st.markdown(f"<div style='text-align: center; font-weight: bold; font-size: 24px;'>事業内容報告書</div>", unsafe_allow_html=True)
    
    reiwa_year = report_date.year - 2018
    st.markdown(f"<div style='text-align: right; font-size: 14px;'>令和{reiwa_year}年 {report_date.month}月 {report_date.day}日</div>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size: 16px; margin-top: 10px;'>学年</div>", unsafe_allow_html=True)
    st.markdown(f"<div style='background-color: #D3D3D3; padding: 5px; border: 1px solid black; display: inline-block; font-weight: bold; font-size: 18px;'>{department}</div>", unsafe_allow_html=True)
    st.markdown("<hr style='border: 0; height: 1px; background: #333; margin: 1em 0;' />")

    st.markdown("### 事業内容報告")
    st.dataframe(current_activities_df, use_container_width=True, hide_index=True)
    st.markdown("<br>")

    st.markdown("### 活動の反省と課題")
    st.markdown("(次年度以降の改善材料になりますので詳細にお願いします)")
    st.write(reflections) # 使用者の入力そのままを表示
    st.markdown("<br>")

    st.markdown("### 次回運営委員会までの活動予定")
    st.dataframe(next_activities_df, use_container_width=True, hide_index=True)
    st.markdown("<hr style='border: 0; height: 1px; background: #333; margin: 1em 0;' />")

    st.subheader("PDFダウンロード (7. ダウンロードボタン)") # 仕様書7
    pdf_output = create_pdf(report_date, department, current_activities_df, reflections, next_activities_df)
    st.download_button(
        label="PDFをダウンロード",
        data=pdf_output,
        file_name=f"事業内容報告書_{department}_{report_date.strftime('%Y%m%d')}.pdf",
        mime="application/pdf"
    )
