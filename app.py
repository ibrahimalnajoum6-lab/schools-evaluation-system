import streamlit as st
import sqlite3
import pandas as pd
import os
import json
import io
import base64
import requests
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import streamlit.components.v1 as components

# -------------------------------------------------------------
# إعداد رابط Google Apps Script للرفع المباشر
# -------------------------------------------------------------
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzbXt7ZJ1qjdnES24kGMDTYifU9MG3eKQRbH3nRu-QUz1Nk2cHvJnSVatZEd2noYARs/exec"

def upload_file_to_drive(file_bytes, filename, school_name, visit_date_obj, mime_type='application/octet-stream'):
    """رفع الملف مباشرة عبر Google Apps Script داخل: اسم المدرسة / الشهر-السنة"""
    try:
        month_folder_name = f"{visit_date_obj.month}-{visit_date_obj.year}"
        encoded_file = base64.b64encode(file_bytes).decode('utf-8')
        
        payload = {
            "fileName": filename,
            "fileBytes": encoded_file,
            "mimeType": mime_type,
            "schoolName": school_name,
            "monthYear": month_folder_name
        }
        
        response = requests.post(APPS_SCRIPT_URL, json=payload, timeout=30)
        res_data = response.json()
        
        if res_data.get("status") == "success":
            return res_data.get("url")
        else:
            st.error(f"خطأ أثناء الرفع إلى درايف: {res_data.get('message')}")
            return None
    except Exception as e:
        st.error(f"فشل الاتصال برابط الرفع: {e}")
        return None

# -------------------------------------------------------------
# إعداد الصفحة وتجاوبها المحترف مع الجوال (Mobile-First CSS)
# -------------------------------------------------------------
st.set_page_config(
    page_title="منظومة تقييم المدارس الشرعية",
    page_icon="🕌",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap');
    
    html, body, [class*="css"] {
        direction: rtl;
        text-align: right;
        font-family: 'Tajawal', sans-serif !important;
    }
    
    /* إخفاء القائمة الجانبية تماماً لمنع التداخل والتشوه على الجوال */
    [data-testid="stSidebar"] {
        display: none !important;
    }
    [data-testid="stSidebarCollapseButton"], [data-testid="collapsedControl"] {
        display: none !important;
    }
    
    /* إخفاء رأس صفحة ستريمليت وشريط الأدوات */
    header[data-testid="stHeader"] {
        display: none !important;
    }
    
    /* ضبط هوامش الصفحة لتناسب شاشة الهاتف بالكامل */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
        max-width: 100% !important;
    }
    
    /* تصميم بطاقات مريحة للعين */
    .mobile-card {
        background-color: #ffffff;
        border-radius: 14px;
        padding: 16px;
        margin-bottom: 14px;
        border: 1px solid #eef2f6;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
    }
    
    /* بطاقة النتيجة الحية العائمة */
    .score-banner {
        background: linear-gradient(135deg, #0d5c3a 0%, #178a58 100%);
        color: white;
        border-radius: 16px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 6px 18px rgba(13,92,58,0.25);
        margin-bottom: 18px;
    }
    
    /* أزرار عريضة ومريحة للّمس */
    .stButton > button {
        border-radius: 12px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        padding: 12px 18px !important;
    }
    
    /* بطاقات بنود التقييم */
    .criterion-box {
        background-color: #f8fafc;
        border-right: 4px solid #0d5c3a;
        border-radius: 8px;
        padding: 12px;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# قاعدة البيانات والتهيئة
# -------------------------------------------------------------
def init_db():
    conn = sqlite3.connect("evaluation_system.db")
    c = conn.cursor()
    
    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        password TEXT,
        full_name TEXT,
        specialization TEXT,
        role TEXT
    )''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS schools (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        gender TEXT,
        location TEXT
    )''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS evaluations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        committee_no TEXT,
        visit_date TEXT,
        academic_year TEXT,
        semester TEXT,
        school_name TEXT,
        gender_type TEXT,
        supervisor_name TEXT,
        teacher_name TEXT,
        subject TEXT,
        specialization TEXT,
        student_count INTEGER,
        grade_level TEXT,
        section TEXT,
        lesson_topic TEXT,
        job_status TEXT,
        experience TEXT,
        scores_json TEXT,
        total_score INTEGER,
        rating TEXT,
        excellence_points TEXT,
        dev_points TEXT,
        suggestions TEXT,
        media_paths TEXT,
        drive_links TEXT,
        status TEXT
    )''')
    
    c.execute("PRAGMA table_info(evaluations)")
    columns = [column[1] for column in c.fetchall()]
    if "drive_links" not in columns:
        c.execute("ALTER TABLE evaluations ADD COLUMN drive_links TEXT")

    c.execute("SELECT COUNT(*) FROM users")
    if c.fetchone()[0] == 0:
        supervisors = [
            ("admin", "admin123", "مدير التعليم الشرعي", "إدارة", "Admin"),
            ("ibrahim", "123456", "إبراهيم احمد النجوم", "لغة إنكليزية", "Supervisor"),
            ("m_shabo", "123456", "محمد مصطفى المصطفى الشعبو", "شريعة", "Supervisor"),
            ("shadi_h", "123456", "شادي أحمد حلاق", "شريعة", "Supervisor"),
            ("m_hout", "123456", "محمد حسن الحوت", "شريعة", "Supervisor"),
            ("abdullah_a", "123456", "عبد الله عارف", "شريعة", "Supervisor"),
            ("m_muhyiddin", "123456", "محمد محي الدين محمد", "شريعة", "Supervisor"),
            ("waddah_m", "123456", "وضاح محمد لمعت مخللاتي", "شريعة", "Supervisor"),
            ("amina_y", "123456", "أمينة يوسف العبد الله", "شريعة", "Supervisor"),
            ("muna_h", "123456", "منى جمعة الحسن", "شريعة", "Supervisor"),
            ("hasnaa_h", "123456", "حسناء حسن الحاج إبراهيم", "شريعة", "Supervisor"),
            ("muna_haj", "123456", "منى حسن الحاج إبراهيم", "شريعة", "Supervisor"),
            ("abdulqader_a", "123456", "عبد القادر محمد شحادة الأحمد", "لغة عربية", "Supervisor"),
            ("ruba_m", "123456", "ربى ماهر مكتبي", "لغة عربية", "Supervisor"),
            ("muthanna_h", "123456", "المثنى محمود الدياب الحماده", "رياضيات", "Supervisor"),
            ("baraa_s", "123456", "براءه محمد جهاد سيلم", "رياضيات", "Supervisor"),
            ("m_masto", "123456", "محمد حسين مسطو", "علم أحياء", "Supervisor"),
            ("rana_m", "123456", "رنا عبد الحميد معرستاوي", "علم أحياء", "Supervisor"),
            ("m_omran", "123456", "محمد أمين عبد الله عمران", "فيزياء وكيمياء", "Supervisor"),
            ("bushra_e", "123456", "بشرى محمود عيد", "فيزياء وكيمياء", "Supervisor"),
            ("hossam_r", "123456", "حسام عمر رسلان", "فلسفة", "Supervisor"),
            ("aya_z", "123456", "آية أنور زيتاني", "فلسفة", "Supervisor"),
            ("abdulrahim_d", "123456", "عبد الرحيم محمد دلو", "جغرافيا", "Supervisor"),
            ("shahla_s", "123456", "شهلا محمد شيخوني", "جغرافيا", "Supervisor"),
            ("m_rajab", "123456", "محمد مصطفى الرجب", "تاريخ", "Supervisor"),
            ("nafisa_f", "123456", "نفيسة فارس الفارس", "لغة إنكليزية", "Supervisor")
        ]
        c.executemany("INSERT INTO users (username, password, full_name, specialization, role) VALUES (?, ?, ?, ?, ?)", supervisors)

    c.execute("SELECT COUNT(*) FROM schools")
    if c.fetchone()[0] == 0:
        schools_data = [
            ("الثانوية الخسروية", "ذكور", "بجانب القلعة"),
            ("الثانوية الشرعية الأولى (العرقوب)", "ذكور", "العرقوب"),
            ("ثانوية صلاح الدين", "ذكور", "صلاح الدين"),
            ("ثانوية الصاخور الشرعية للبنين", "ذكور", "الصاخور"),
            ("ثانوية الصاخور الشرعية للبنات", "إناث", "الصاخور"),
            ("الثانوية النعمانية الشرعية للبنين", "ذكور", "جمعية الزهراء"),
            ("الثانوية النعمانية الشرعية للبنات", "إناث", "جمعية الزهراء"),
            ("ثانوية هارون الرشيد الشرعية للبنين", "ذكور", "الراموسة"),
            ("ثانوية هارون الرشيد الشرعية للبنات", "إناث", "الراموسة"),
            ("ثانوية معروف الكرخي للبنين", "ذكور", "الشيخ مقصود"),
            ("ثانوية معروف الكرخي للبنات", "إناث", "الشيخ مقصود"),
            ("ثانوية المنصورة الشرعية للبنين", "ذكور", "المنصورة"),
            ("ثانوية المنصورة الشرعية للبنات", "إناث", "المنصورة"),
            ("ثانوية الإرث النبوي للبنين", "ذكور", "باب النيرب"),
            ("ثانوية الإرث النبوي للبنات", "إناث", "المدرسة العلمية"),
            ("ثانوية الحمدانية الشرعية للبنين", "ذكور", "الحمدانية"),
            ("ثانوية ابن الجوزي الشرعية للبنين", "ذكور", "جمعية الريادة"),
            ("ثانوية ابن الجوزي الشرعية للبنات", "إناث", "جمعية الريادة"),
            ("ثانوية عز العلماء للبنين", "ذكور", "المشارقة"),
            ("ثانوية زين العابدين الشرعية للبنات", "إناث", "حلب الجديدة"),
            ("الثانوية الشرعية للبنين في النيرب", "ذكور", "النيرب"),
            ("الثانوية الشرعية للبنات في النيرب", "إناث", "النيرب")
        ]
        c.executemany("INSERT INTO schools (name, gender, location) VALUES (?, ?, ?)", schools_data)
        
    conn.commit()
    conn.close()

init_db()

DOMAINS = ["التخطيط", "تنفيذ الدرس", "إدارة الصف", "شخصية المعلم", "المادة العلمية"]

CRITERIA = [
    {"domain": "التخطيط", "id": 1, "text": "يلتزم بتوزيع المنهاج", "max": 3},
    {"domain": "التخطيط", "id": 2, "text": "يخطط لتهيئة محفزة تراعي فيها المعلومات السابقة وتمهد للدرس الجديد", "max": 3},
    {"domain": "التخطيط", "id": 3, "text": "يصوغ الأهداف صياغة صحيحة للمجالات الثلاثة (المعرفي - المهاري - الوجداني)", "max": 4},
    {"domain": "التخطيط", "id": 4, "text": "يحدد استراتيجيات التدريس والأنشطة المناسبة لها", "max": 3},
    {"domain": "التخطيط", "id": 5, "text": "يحدد الوسائط المتعددة والبدائل التعليمية المناسبة", "max": 2},
    {"domain": "التخطيط", "id": 6, "text": "يخطط لتقويم لمراحله المختلفة (المرحلي - النهائي)", "max": 4},
    {"domain": "تنفيذ الدرس", "id": 7, "text": "يستثير دافعية المتعلمين نحو التعلم", "max": 4},
    {"domain": "تنفيذ الدرس", "id": 8, "text": "يوظف طرائق التدريس والأنشطة والوسائط بصورة مناسبة لمستويات المتعلمين وأهداف الدرس", "max": 6},
    {"domain": "تنفيذ الدرس", "id": 9, "text": "يستخدم السبورة بالشكل الأمثل للمحتوى التعليمي", "max": 4},
    {"domain": "تنفيذ الدرس", "id": 10, "text": "يراعي تنوع أهداف الدرس في التقويم وتنوع الأسئلة", "max": 4},
    {"domain": "تنفيذ الدرس", "id": 11, "text": "يعزز المتعلمين في الوقت المناسب وبالطريقة المناسبة", "max": 4},
    {"domain": "تنفيذ الدرس", "id": 12, "text": "يخصص الوقت المناسب لكل مرحلة من مراحل الدرس", "max": 3},
    {"domain": "تنفيذ الدرس", "id": 13, "text": "ينفذ ختاماً مناسباً للدرس", "max": 3},
    {"domain": "إدارة الصف", "id": 14, "text": "يهيئ بيئة صفية مادية ونفسية مريحة ومناسبة للمتعلمين", "max": 3},
    {"domain": "إدارة الصف", "id": 15, "text": "يوزع الاهتمام بين المتعلمين ويتعامل معهم بطريقة عادلة", "max": 3},
    {"domain": "إدارة الصف", "id": 16, "text": "يتبع قواعد صفية تنظم تفاعل وأنشطة المتعلمين بما يحقق أهداف الدرس", "max": 3},
    {"domain": "إدارة الصف", "id": 17, "text": "يراعي الفروق الفردية (فجوات الفاقد التعليمي) ويوزع الاهتمام بين المتعلمين", "max": 4},
    {"domain": "شخصية المعلم", "id": 18, "text": "يلتزم بالسمت الشرعي والمظهر اللائق", "max": 5},
    {"domain": "شخصية المعلم", "id": 19, "text": "يستخدم اللغة السليمة ويتنوع في نبرات الصوت ويستخدم لغة الجسد المناسبة", "max": 4},
    {"domain": "المادة العلمية", "id": 20, "text": "يهتم بالحيوية والنشاط داخل الصف", "max": 3},
    {"domain": "المادة العلمية", "id": 21, "text": "يركز على الجانب القيمي للطلاب ويستثمر الأمثلة لتعزيز القيم والأخلاق الحسنة", "max": 5},
    {"domain": "المادة العلمية", "id": 22, "text": "يتصف بالاتزان الانفعالي والذكاء العاطفي ويتقبل الملاحظات اللازمة للتطوير", "max": 3},
    {"domain": "المادة العلمية", "id": 23, "text": "يربط المادة العلمية بالحياة العملية ويحقق الترابط والتكامل مع المواد الأخرى", "max": 4},
    {"domain": "المادة العلمية", "id": 24, "text": "يراعي التدرج والتسلسل المنطقي للمادة والانتقال الصحيح", "max": 4},
    {"domain": "المادة العلمية", "id": 25, "text": "يظهر تمكناً من المادة العلمية التي يقدمها", "max": 12},
]

# -------------------------------------------------------------
# دالة Excel الرسمية
# -------------------------------------------------------------
def generate_evaluation_excel_form(eval_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "استمارة التقييم"
    ws.views.sheetView[0].rightToLeft = True

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    font_bold = Font(name="Arial", size=10, bold=True)
    font_norm = Font(name="Arial", size=9)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    ws.merge_cells("A1:B1")
    ws["A1"] = "مديرية أوقاف حلب"
    ws["A1"].font = font_bold
    ws["A1"].alignment = align_right

    ws.merge_cells("C1:D1")
    ws["C1"] = "لجنة تقييم أداء المدرس"
    ws["C1"].font = Font(name="Arial", size=13, bold=True)
    ws["C1"].alignment = align_center

    ws.merge_cells("E1:F1")
    ws["E1"] = f"رقم اللجنة: {eval_data['committee_no']}"
    ws["E1"].font = font_bold
    ws["E1"].alignment = Alignment(horizontal="left", vertical="center")

    meta_rows = [
        ("التاريخ:", str(eval_data['visit_date']), "المؤسسة التعليمية الشرعية:", eval_data['school_name']),
        ("العام الدراسي:", eval_data['academic_year'], "الفصل الدراسي:", eval_data['semester']),
        ("المدرس:", eval_data['teacher_name'], "المسؤول العلمي:", eval_data['supervisor_name']),
        ("المادة:", eval_data['subject'], "الاختصاص:", eval_data['specialization']),
        ("عدد الطلاب:", eval_data['student_count'], "الصف:", eval_data['grade_level']),
        ("الشعبة:", eval_data['section'], "الموضوع:", eval_data['lesson_topic']),
        ("الخبرة في التعليم:", eval_data['experience'], "الوضع الوظيفي:", eval_data['job_status'])
    ]

    r_idx = 3
    for r in meta_rows:
        ws.cell(row=r_idx, column=1, value=r[0]).fill = fill_gray
        ws.cell(row=r_idx, column=1).font = font_bold
        ws.cell(row=r_idx, column=2, value=r[1]).font = font_norm
        ws.merge_cells(start_row=r_idx, start_column=2, end_row=r_idx, end_column=3)
        
        ws.cell(row=r_idx, column=4, value=r[2]).fill = fill_gray
        ws.cell(row=r_idx, column=4).font = font_bold
        ws.cell(row=r_idx, column=5, value=r[3]).font = font_norm
        ws.merge_cells(start_row=r_idx, start_column=5, end_row=r_idx, end_column=6)
        
        for c in range(1, 7):
            ws.cell(row=r_idx, column=c).border = border
            ws.cell(row=r_idx, column=c).alignment = align_center
        r_idx += 1

    r_idx += 1
    headers = ["م", "مجال التقييم", "بنود اللائحة", "الدرجة المستحقة", "الدرجة الفعلية", "الملاحظات"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=r_idx, column=col_idx, value=h)
        cell.font = font_bold
        cell.fill = fill_gray
        cell.alignment = align_center
        cell.border = border

    scores = json.loads(eval_data['scores_json']) if isinstance(eval_data['scores_json'], str) else eval_data['scores_json']
    start_eval_row = r_idx + 1

    for item in CRITERIA:
        r_idx += 1
        actual = scores.get(str(item['id']), item['max'])
        ws.cell(row=r_idx, column=1, value=item['id']).alignment = align_center
        ws.cell(row=r_idx, column=2, value=item['domain']).alignment = align_center
        ws.cell(row=r_idx, column=3, value=item['text']).alignment = align_right
        ws.cell(row=r_idx, column=4, value=item['max']).alignment = align_center
        ws.cell(row=r_idx, column=5, value=actual).alignment = align_center
        
        for c in range(1, 7):
            ws.cell(row=r_idx, column=c).border = border
            ws.cell(row=r_idx, column=c).font = font_norm

    ws.merge_cells(f"B{start_eval_row}:B{start_eval_row+5}")
    ws.merge_cells(f"B{start_eval_row+6}:B{start_eval_row+12}")
    ws.merge_cells(f"B{start_eval_row+13}:B{start_eval_row+16}")
    ws.merge_cells(f"B{start_eval_row+17}:B{start_eval_row+18}")
    ws.merge_cells(f"B{start_eval_row+19}:B{r_idx}")

    ws.merge_cells(f"F{start_eval_row}:F{start_eval_row+9}")
    ws[f"F{start_eval_row}"] = f"نقاط التميز:\n{eval_data.get('excellence_points', '')}"
    ws[f"F{start_eval_row}"].alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)

    ws.merge_cells(f"F{start_eval_row+10}:F{start_eval_row+18}")
    ws[f"F{start_eval_row+10}"] = f"نقاط التطوير:\n{eval_data.get('dev_points', '')}"
    ws[f"F{start_eval_row+10}"].alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)

    ws.merge_cells(f"F{start_eval_row+19}:F{r_idx}")
    ws[f"F{start_eval_row+19}"] = f"المقترحات:\n{eval_data.get('suggestions', '')}"
    ws[f"F{start_eval_row+19}"].alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)

    r_idx += 1
    ws.merge_cells(f"A{r_idx}:C{r_idx}")
    ws[f"A{r_idx}"] = f"مجموع الدرجات: {eval_data['total_score']} / 100"
    ws[f"A{r_idx}"].font = font_bold
    ws[f"A{r_idx}"].alignment = align_center
    ws[f"A{r_idx}"].fill = fill_gray

    ws.merge_cells(f"D{r_idx}:F{r_idx}")
    ws[f"D{r_idx}"] = f"التقدير النهائي: {eval_data['rating']}"
    ws[f"D{r_idx}"].font = font_bold
    ws[f"D{r_idx}"].alignment = align_center
    ws[f"D{r_idx}"].fill = fill_gray

    for c in range(1, 7):
        ws.cell(row=r_idx, column=c).border = border

    r_idx += 2
    sigs = ["المدرسة", "", "الموجّه الاختصاصي", "", "شعبة التوجيه الاختصاصي", "رئيس إدارة التعليم الشرعي"]
    for i, s in enumerate(sigs, 1):
        ws.cell(row=r_idx, column=i, value=s).font = font_bold
        ws.cell(row=r_idx, column=i).alignment = align_center

    col_widths = {1: 5, 2: 15, 3: 45, 4: 12, 5: 12, 6: 28}
    for c, w in col_widths.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer.getvalue()

# -------------------------------------------------------------
# تسجيل الدخول
# -------------------------------------------------------------
if "user" not in st.session_state:
    st.session_state.user = None

def login(username, password):
    conn = sqlite3.connect("evaluation_system.db")
    c = conn.cursor()
    c.execute("SELECT id, username, full_name, specialization, role FROM users WHERE username=? AND password=?", (username, password))
    user = c.fetchone()
    conn.close()
    return user

if st.session_state.user is None:
    st.markdown("<div style='text-align: center; margin-top: 25px;'>", unsafe_allow_html=True)
    st.markdown("## 🕌 منظومة التقييم الشرعي")
    st.caption("تسجيل دخول الموجهين والمشرفين")
    st.markdown("</div>", unsafe_allow_html=True)
    
    with st.container():
        st.markdown("<div class='mobile-card'>", unsafe_allow_html=True)
        u_name = st.text_input("اسم المستخدم", placeholder="اسم المستخدم...")
        u_pass = st.text_input("كلمة المرور", type="password", placeholder="••••••")
        
        if st.button("تسجيل الدخول", type="primary", use_container_width=True):
            user_data = login(u_name, u_pass)
            if user_data:
                st.session_state.user = {
                    "id": user_data[0], "username": user_data[1],
                    "name": user_data[2], "specialization": user_data[3], "role": user_data[4]
                }
                st.rerun()
            else:
                st.error("اسم المستخدم أو كلمة المرور غير صحيحة")
        st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

# -------------------------------------------------------------
# الشريط العلوي والتنقل
# -------------------------------------------------------------
st.markdown(f"""
<div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; background: #fff; padding: 10px 14px; border-radius: 12px; box-shadow: 0 2px 6px rgba(0,0,0,0.04);'>
    <div>
        <span style='font-size: 15px; font-weight: 800; color: #0d5c3a;'>👤 {st.session_state.user['name']}</span><br>
        <span style='font-size: 11px; color: #64748b;'>{st.session_state.user['specialization']} | {st.session_state.user['role']}</span>
    </div>
</div>
""", unsafe_allow_html=True)

menu_options = ["📝 تقييم جديد", "🔍 سجل الزيارات"]
if st.session_state.user["role"] == "Admin":
    menu_options.append("⚙️ لوحة الإدارة")

c_nav, c_out = st.columns([4, 1])
with c_nav:
    choice = st.radio("التنقل السريع", menu_options, horizontal=True, label_visibility="collapsed")
with c_out:
    if st.button("خروج", use_container_width=True):
        st.session_state.user = None
        st.rerun()

# -------------------------------------------------------------
# 1. شاشة استمارة تقييم جديدة
# -------------------------------------------------------------
if choice == "📝 تقييم جديد":
    
    for item in CRITERIA:
        k = f"q_{item['id']}"
        if k not in st.session_state:
            st.session_state[k] = item['max']

    total_live = sum(st.session_state[f"q_{item['id']}"] for item in CRITERIA)
    rating_live = "ممتاز" if total_live >= 90 else "جيد جداً" if total_live >= 80 else "جيد" if total_live >= 70 else "مقبول" if total_live >= 50 else "ضعيف"
    
    st.markdown(f"""
    <div class='score-banner'>
        <div style='font-size: 13px; opacity: 0.9;'>المجموع الكلي المباشر</div>
        <div style='font-size: 32px; font-weight: 800; margin: 4px 0;'>{total_live} <span style='font-size: 16px; font-weight: 500;'>/ 100</span></div>
        <div style='display: inline-block; background: rgba(255,255,255,0.25); padding: 3px 12px; border-radius: 20px; font-size: 13px; font-weight: 700;'>
            التقدير: {rating_live}
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.expander("📌 البيانات الأساسية للزيارة", expanded=True):
        conn = sqlite3.connect("evaluation_system.db")
        schools_df = pd.read_sql_query("SELECT name, gender FROM schools", conn)
        conn.close()
        
        gender_type = st.radio("نوع المدرسة", ["ذكور", "إناث"], horizontal=True)
        filtered_schools = schools_df[schools_df['gender'] == gender_type]['name'].tolist()
        school_name = st.selectbox("المدرسة", filtered_schools if filtered_schools else ["لا توجد مدارس"])
        
        c_t1, c_t2 = st.columns(2)
        with c_t1:
            teacher_name = st.text_input("اسم المدرس *")
            subject = st.text_input("المادة", value=st.session_state.user["specialization"])
        with c_t2:
            grade_level = st.selectbox("الصف", ["السابع", "الثامن", "التاسع", "العاشر", "الحادي عشر", "الثالث الثانوي"])
            section = st.text_input("الشعبة", value="الأولى")

        lesson_topic = st.text_input("موضوع الدرس")
        
        with st.expander("بيانات إضافية (اختياري)"):
            c_e1, c_e2 = st.columns(2)
            with c_e1:
                visit_date = st.date_input("تاريخ الزيارة", value=date.today())
                academic_year = st.selectbox("العام الدراسي", ["2026 - 2027", "2025 - 2026"])
                semester = st.selectbox("الفصل الدراسي", ["الفصل الأول", "الفصل الثاني"])
                committee_no = st.text_input("رقم اللجنة", value="1")
            with c_e2:
                specialization = st.text_input("الاختصاص")
                student_count = st.number_input("عدد الطلاب", min_value=1, value=25)
                job_status = st.selectbox("الوضع الوظيفي", ["أصيل", "وكيل", "مكلف"])
                experience = st.text_input("الخبرة", value="5 سنوات")

    st.markdown("#### 📋 بنود التقييم حسب المجال")
    tabs = st.tabs(DOMAINS)
    
    for idx, domain in enumerate(DOMAINS):
        with tabs[idx]:
            domain_items = [i for i in CRITERIA if i['domain'] == domain]
            for item in domain_items:
                st.markdown(f"""
                <div class='criterion-box'>
                    <div style='font-weight: 700; font-size: 14px; color: #1e293b;'>{item['id']}. {item['text']}</div>
                    <div style='font-size: 12px; color: #64748b;'>الدرجة القصوى المستحقة: ({item['max']})</div>
                </div>
                """, unsafe_allow_html=True)
                
                scores_options = list(range(item['max'] + 1))
                st.select_slider(
                    f"درجة بند {item['id']}",
                    options=scores_options,
                    value=st.session_state[f"q_{item['id']}"],
                    key=f"slider_{item['id']}",
                    on_change=lambda i=item['id']: st.session_state.update({f"q_{i}": st.session_state[f"slider_{i}"]}),
                    label_visibility="collapsed"
                )

    st.markdown("#### ✍️ الملاحظات والشواهد")
    with st.container():
        excellence_points = st.text_area("🌟 نقاط التميز", placeholder="اكتب نقاط القوة والتميز...")
        dev_points = st.text_area("💡 نقاط التطوير", placeholder="اكتب نقاط التحسين والتطوير...")
        suggestions = st.text_area("📌 المقترحات والتوصيات", placeholder="المقترحات والتوجيهات...")
        
        uploaded_files = st.file_uploader("📷 رفع شواهد وصور من الكاميرا / المعرض", accept_multiple_files=True)

    if st.button("💾 حفظ الاستمارة ورفع الملفات إلى جوجل درايف", type="primary", use_container_width=True):
        if not teacher_name.strip():
            st.error("⚠️ يرجى إدخال اسم المدرس أولاً.")
        else:
            final_scores = {str(item['id']): st.session_state[f"q_{item['id']}"] for item in CRITERIA}
            saved_media = []
            drive_links = []
            
            if uploaded_files:
                os.makedirs("uploads", exist_ok=True)
                for f in uploaded_files:
                    f_bytes = f.getbuffer()
                    path = os.path.join("uploads", f.name)
                    with open(path, "wb") as file_out:
                        file_out.write(f_bytes)
                    saved_media.append(path)
                    
                    with st.spinner(f"جاري رفع {f.name}..."):
                        link = upload_file_to_drive(f.getvalue(), f.name, school_name, visit_date, f.type)
                        if link:
                            drive_links.append(link)

            eval_data_dict = {
                "committee_no": committee_no, "visit_date": str(visit_date), "academic_year": academic_year,
                "semester": semester, "school_name": school_name, "gender_type": gender_type,
                "supervisor_name": st.session_state.user["name"], "teacher_name": teacher_name, "subject": subject,
                "specialization": specialization if specialization else subject, "student_count": student_count, "grade_level": grade_level,
                "section": section, "lesson_topic": lesson_topic, "job_status": job_status,
                "experience": experience, "scores_json": json.dumps(final_scores), "total_score": total_live,
                "rating": rating_live, "excellence_points": excellence_points, "dev_points": dev_points,
                "suggestions": suggestions
            }
            
            with st.spinner("جاري إنشاء ورفع تقرير الإكسل الرسمي إلى درايف..."):
                excel_bytes = generate_evaluation_excel_form(eval_data_dict)
                excel_filename = f"استمارة_{teacher_name}_{visit_date}.xlsx"
                excel_link = upload_file_to_drive(
                    excel_bytes, excel_filename, school_name, visit_date,
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                if excel_link:
                    drive_links.append(excel_link)
            
            conn = sqlite3.connect("evaluation_system.db")
            c = conn.cursor()
            c.execute('''INSERT INTO evaluations (
                committee_no, visit_date, academic_year, semester, school_name, gender_type,
                supervisor_name, teacher_name, subject, specialization, student_count,
                grade_level, section, lesson_topic, job_status, experience,
                scores_json, total_score, rating, excellence_points, dev_points, suggestions,
                media_paths, drive_links, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', (
                committee_no, str(visit_date), academic_year, semester, school_name, gender_type,
                st.session_state.user["name"], teacher_name, subject, eval_data_dict["specialization"], student_count,
                grade_level, section, lesson_topic, job_status, experience,
                json.dumps(final_scores), total_live, rating_live, excellence_points, dev_points, suggestions,
                ",".join(saved_media), ",".join(drive_links), "معتمد"
            ))
            conn.commit()
            conn.close()
            st.success(f"✅ تم حفظ الاستمارة ورفع ملفاتها إلى مجلد ({school_name} / {visit_date.month}-{visit_date.year}) في جوجل درايف بنجاح!")

# -------------------------------------------------------------
# 2. سجل الزيارات
# -------------------------------------------------------------
elif choice == "🔍 سجل الزيارات":
    st.markdown("### 🔍 سجل الزيارات والتصدير")
    
    conn = sqlite3.connect("evaluation_system.db")
    if st.session_state.user["role"] == "Admin":
        df = pd.read_sql_query("SELECT * FROM evaluations ORDER BY id DESC", conn)
    else:
        df = pd.read_sql_query("SELECT * FROM evaluations WHERE supervisor_name=? ORDER BY id DESC", conn, params=(st.session_state.user["name"],))
    conn.close()
    
    if df.empty:
        st.info("لا توجد استمارات مسجلة حتى الآن.")
    else:
        search_txt = st.text_input("🔍 بحث باسم المدرس أو المدرسة", placeholder="اكتب للبحث السريع...")
        filtered_df = df.copy()
        if search_txt:
            filtered_df = filtered_df[
                filtered_df["teacher_name"].str.contains(search_txt, na=False) |
                filtered_df["school_name"].str.contains(search_txt, na=False)
            ]

        for _, row in filtered_df.iterrows():
            with st.expander(f"📄 {row['teacher_name']} — {row['school_name']} ({row['total_score']} / 100)"):
                st.write(f"**المادة:** {row['subject']} | **الصف:** {row['grade_level']}")
                st.write(f"**التاريخ:** {row['visit_date']} | **التقدير:** {row['rating']}")
                st.write(f"**الموجه:** {row['supervisor_name']}")
                
                rec_dict = row.to_dict()
                excel_form_bytes = generate_evaluation_excel_form(rec_dict)
                st.download_button(
                    label="📊 تحميل الاستمارة (Excel)",
                    data=excel_form_bytes,
                    file_name=f"استمارة_{row['teacher_name']}_{row['visit_date']}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"dl_{row['id']}"
                )
                
                if row.get("drive_links"):
                    st.markdown("📂 **روابط درايف:**")
                    for i, link in enumerate(str(row["drive_links"]).split(','), 1):
                        if link.strip():
                            st.markdown(f"- [عرض الملف {i} على درايف]({link.strip()})")

# -------------------------------------------------------------
# 3. لوحة الإدارة
# -------------------------------------------------------------
elif choice == "⚙️ لوحة الإدارة":
    st.markdown("### ⚙️ إدارة النظام")
    admin_tab1, admin_tab2 = st.tabs(["👥 الموجهين", "🏫 المدارس"])
    
    with admin_tab1:
        conn = sqlite3.connect("evaluation_system.db")
        users_df = pd.read_sql_query("SELECT id, username, full_name, specialization, role FROM users", conn)
        conn.close()
        st.dataframe(users_df, use_container_width=True)
        
        with st.expander("➕ إضافة موجه جديد"):
            with st.form("add_user_mobile"):
                n_name = st.text_input("الاسم الكامل")
                n_user = st.text_input("اسم الدخول")
                n_pass = st.text_input("كلمة السر", type="password")
                n_spec = st.text_input("التخصص")
                n_role = st.selectbox("الدور", ["Supervisor", "Admin"])
                if st.form_submit_button("إضافة", use_container_width=True):
                    if n_name and n_user and n_pass:
                        try:
                            conn = sqlite3.connect("evaluation_system.db")
                            c = conn.cursor()
                            c.execute("INSERT INTO users (username, password, full_name, specialization, role) VALUES (?, ?, ?, ?, ?)",
                                      (n_user, n_pass, n_name, n_spec, n_role))
                            conn.commit()
                            conn.close()
                            st.success("تمت الإضافة بنجاح")
                            st.rerun()
                        except Exception as e:
                            st.error(f"خطأ: {e}")

    with admin_tab2:
        conn = sqlite3.connect("evaluation_system.db")
        schools_all = pd.read_sql_query("SELECT id, name, gender, location FROM schools", conn)
        conn.close()
        st.dataframe(schools_all, use_container_width=True)
        
        with st.expander("➕ إضافة مدرسة جديدة"):
            with st.form("add_school_mobile"):
                s_name = st.text_input("اسم المدرسة")
                s_gen = st.selectbox("النوع", ["ذكور", "إناث"])
                s_loc = st.text_input("الموقع")
                if st.form_submit_button("إضافة المدرسة", use_container_width=True):
                    if s_name:
                        conn = sqlite3.connect("evaluation_system.db")
                        c = conn.cursor()
                        c.execute("INSERT INTO schools (name, gender, location) VALUES (?, ?, ?)", (s_name, s_gen, s_loc))
                        conn.commit()
                        conn.close()
                        st.success("تمت الإضافة بنجاح")
                        st.rerun()