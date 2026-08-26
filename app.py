import streamlit as st
import sqlite3
import pandas as pd
import os
import json
import io
import base64
import requests
import threading
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import streamlit.components.v1 as components

# محاولة استيراد Weasyprint بشكل آمن لعدم التسبب بشاشة بيضاء إن لم تكن المكتبات النظامية متوفرة
try:
    from weasyprint import HTML
    WEASYPRINT_AVAILABLE = True
except Exception:
    WEASYPRINT_AVAILABLE = False

# -------------------------------------------------------------
# إعداد رابط Google Apps Script للرفع المباشر
# -------------------------------------------------------------
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzbXt7ZJ1qjdnES24kGMDTYifU9MG3eKQRbH3nRu-QUz1Nk2cHvJnSVatZEd2noYARs/exec"

def upload_file_to_drive(file_bytes, filename, school_name, visit_date_obj, mime_type='application/octet-stream'):
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
            return None
    except Exception:
        return None

def auto_backup_database_to_drive():
    try:
        db_path = "evaluation_system.db"
        if os.path.exists(db_path):
            with open(db_path, "rb") as f:
                db_bytes = f.read()
            backup_filename = f"DB_Backup_Schools_System_{date.today()}_{int(datetime.now().timestamp())}.db"
            upload_file_to_drive(
                db_bytes,
                backup_filename,
                "النسخ_الاحتياطية_للنظام",
                date.today(),
                "application/x-sqlite3"
            )
    except Exception:
        pass

def background_upload_task(eval_id, eval_data_dict, uploaded_files_data, school_name, visit_date):
    try:
        drive_links = []
        excel_bytes = generate_evaluation_excel_form(eval_data_dict)
        excel_link = upload_file_to_drive(
            excel_bytes, f"استمارة_{eval_data_dict['teacher_name']}_{visit_date}.xlsx", school_name, visit_date,
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        if excel_link:
            drive_links.append(excel_link)

        for f_name, f_bytes, f_type in uploaded_files_data:
            link = upload_file_to_drive(f_bytes, f_name, school_name, visit_date, f_type)
            if link:
                drive_links.append(link)

        if drive_links:
            conn = sqlite3.connect("evaluation_system.db")
            c = conn.cursor()
            c.execute("UPDATE evaluations SET drive_links=? WHERE id=?", (",".join(drive_links), eval_id))
            conn.commit()
            conn.close()

        auto_backup_database_to_drive()
    except Exception:
        pass

def delete_evaluation_by_id(eval_id):
    conn = sqlite3.connect("evaluation_system.db")
    c = conn.cursor()
    c.execute("DELETE FROM evaluations WHERE id=?", (eval_id,))
    conn.commit()
    conn.close()
    threading.Thread(target=auto_backup_database_to_drive).start()

# -------------------------------------------------------------
# إعداد الصفحة وتصاميم الـ CSS وإخفاء الشارة نهائياً
# -------------------------------------------------------------
st.set_page_config(
    page_title="منظومة تقييم المدارس الشرعية",
    page_icon="🕌",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800;900&display=swap');
    
    html, body, [class*="css"] {
        direction: rtl;
        text-align: right;
        font-family: 'Tajawal', sans-serif !important;
        scroll-behavior: smooth;
        background-color: #f8fafc;
        font-size: 16px;
    }
    
    [data-testid="stSidebar"], [data-testid="stSidebarCollapseButton"], [data-testid="collapsedControl"], header[data-testid="stHeader"] {
        display: none !important;
    }
    
    #MainMenu, footer, header {visibility: hidden !important; display: none !important;}
    
    div[class*="viewerBadge"], 
    [data-testid="stStatusWidget"],
    .viewerBadge_container__1QSob,
    a[href*="streamlit.io"],
    footer {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        pointer-events: none !important;
        height: 0px !important;
    }
    
    .block-container {
        padding: 0.8rem !important;
        max-width: 100% !important;
    }
    
    .mobile-card {
        background-color: #ffffff;
        border-radius: 18px;
        padding: 20px;
        margin-bottom: 16px;
        border: 2px solid #e2e8f0;
        box-shadow: 0 4px 16px rgba(0,0,0,0.05);
    }
    
    .score-banner {
        background: linear-gradient(135deg, #0d5c3a 0%, #15803d 100%);
        color: white;
        border-radius: 18px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 8px 22px rgba(13,92,58,0.28);
        margin-bottom: 18px;
    }

    .success-box {
        background: linear-gradient(135deg, #065f46 0%, #047857 100%);
        color: white;
        border-radius: 18px;
        padding: 22px;
        text-align: center;
        box-shadow: 0 8px 24px rgba(6, 95, 70, 0.35);
        margin-bottom: 20px;
        border: 2px solid #34d399;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #e2e8f0;
        padding: 8px;
        border-radius: 16px;
        margin-bottom: 16px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 12px;
        padding: 12px 18px;
        font-weight: 800;
        font-size: 16px;
        color: #334155;
        background-color: transparent;
        border: none !important;
        transition: all 0.25s ease;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #0d5c3a !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.12);
    }
    
    .criterion-item {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-right: 6px solid #0d5c3a;
        border-radius: 12px;
        padding: 14px 16px;
        margin-bottom: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.03);
    }
    
    .stButton > button {
        border-radius: 12px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        padding: 8px 16px !important;
        min-height: 42px !important;
        background-color: #ffffff !important;
        color: #334155 !important;
        border: 1px solid #cbd5e1 !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03) !important;
        transition: all 0.2s ease !important;
    }
    
    .stButton > button:hover {
        background-color: #f1f5f9 !important;
        border-color: #94a3b8 !important;
        color: #0d5c3a !important;
    }
    
    input, select, textarea {
        font-size: 16px !important;
        padding: 12px !important;
    }
    
    .stRadio [role="radiogroup"] {
        gap: 12px !important;
    }
</style>

<script>
    function removeBrandingCompletely() {
        const allElements = document.querySelectorAll('*');
        allElements.forEach(el => {
            if (el.children.length === 0 && el.innerText) {
                if (el.innerText.includes('Created by') || el.innerText.includes('Made with Streamlit')) {
                    el.style.display = 'none';
                    if (el.parentElement) el.parentElement.style.display = 'none';
                }
            }
        });
        const badges = document.querySelectorAll('div[class*="viewerBadge"], a[href*="streamlit.io"], footer, header');
        badges.forEach(b => b.remove());
    }
    setInterval(removeBrandingCompletely, 300);
</script>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# قاعدة البيانات
# -------------------------------------------------------------
def init_db():
    conn = sqlite3.connect("evaluation_system.db")
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password TEXT, full_name TEXT, specialization TEXT, role TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS schools (
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, gender TEXT, location TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS evaluations (
        id INTEGER PRIMARY KEY AUTOINCREMENT, committee_no TEXT, visit_date TEXT, academic_year TEXT, semester TEXT,
        school_name TEXT, gender_type TEXT, supervisor_name TEXT, teacher_name TEXT, subject TEXT,
        specialization TEXT, student_count INTEGER, grade_level TEXT, section TEXT, lesson_topic TEXT,
        job_status TEXT, experience TEXT, scores_json TEXT, total_score INTEGER, rating TEXT,
        excellence_points TEXT, dev_points TEXT, suggestions TEXT, media_paths TEXT, drive_links TEXT, status TEXT
    )''')
    
    c.execute("PRAGMA table_info(evaluations)")
    columns = [col[1] for col in c.fetchall()]
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

def calculate_domain_scores(scores_dict):
    dom_totals = {d: 0 for d in DOMAINS}
    for item in CRITERIA:
        actual = int(scores_dict.get(str(item['id']), item['max']))
        dom_totals[item['domain']] += actual
    return dom_totals

def get_evaluation_html(eval_data):
    scores = json.loads(eval_data['scores_json']) if isinstance(eval_data['scores_json'], str) else eval_data['scores_json']
    rows_html = ""
    for item in CRITERIA:
        actual = scores.get(str(item['id']), item['max'])
        i_id = item['id']
        domain_td = ""
        if i_id == 1: domain_td = '<td rowspan="6" style="text-align: center; font-weight: bold; vertical-align: middle; background-color: #f9f9f9;">التخطيط</td>'
        elif i_id == 7: domain_td = '<td rowspan="7" style="text-align: center; font-weight: bold; vertical-align: middle; background-color: #f9f9f9;">تنفيذ الدرس</td>'
        elif i_id == 14: domain_td = '<td rowspan="4" style="text-align: center; font-weight: bold; vertical-align: middle; background-color: #f9f9f9;">إدارة الصف</td>'
        elif i_id == 18: domain_td = '<td rowspan="2" style="text-align: center; font-weight: bold; vertical-align: middle; background-color: #f9f9f9;">شخصية المعلم</td>'
        elif i_id == 20: domain_td = '<td rowspan="6" style="text-align: center; font-weight: bold; vertical-align: middle; background-color: #f9f9f9;">المادة العلمية</td>'

        notes_td = ""
        if i_id == 1: notes_td = f'<td rowspan="10" style="vertical-align: top; text-align: right; padding: 4px; font-size: 11px;"><b>نقاط التميز:</b><br>{eval_data.get("excellence_points", "")}</td>'
        elif i_id == 11: notes_td = f'<td rowspan="9" style="vertical-align: top; text-align: right; padding: 4px; font-size: 11px;"><b>نقاط التطوير:</b><br>{eval_data.get("dev_points", "")}</td>'
        elif i_id == 20: notes_td = f'<td rowspan="6" style="vertical-align: top; text-align: right; padding: 4px; font-size: 11px;"><b>المقترحات:</b><br>{eval_data.get("suggestions", "")}</td>'

        rows_html += f"""
        <tr>
            <td style="width: 5%; text-align: center; font-size: 12px;">{i_id}</td>
            {domain_td}
            <td style="width: 42%; text-align: right; font-size: 12px;">{item['text']}</td>
            <td style="width: 9%; text-align: center; font-size: 12px;">{item['max']}</td>
            <td style="width: 9%; text-align: center; font-weight: bold; font-size: 12px;">{actual}</td>
            {notes_td}
        </tr>
        """

    html = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="utf-8">
        <title>استمارة تقييم أداء المدرس</title>
        <style>
            @page {{ size: A4; margin: 10mm; }}
            @media print {{ body {{ margin: 0; padding: 0; }} .no-print {{ display: none !important; }} }}
            body {{ font-family: 'Segoe UI', Tahoma, Arial, sans-serif; direction: rtl; text-align: right; background-color: #fff; color: #000; padding: 5px; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 4px; }}
            th, td {{ border: 1px solid #000; padding: 2.5px 4px; }}
            .bg-gray {{ background-color: #f2f2f2; }}
            .header-tbl td {{ border: none; font-weight: bold; }}
            .btn-print {{ background-color: #0d5c3a; color: white; padding: 12px 18px; border: none; border-radius: 8px; cursor: pointer; font-size: 16px; font-weight: bold; margin-bottom: 12px; width: 100%; }}
        </style>
    </head>
    <body>
        <div class="no-print">
            <button class="btn-print" onclick="window.print()">🖨️ اضغط هنا للطباعة أو الحفظ كـ PDF عبر المتصفح</button>
        </div>
        <table class="header-tbl">
            <tr>
                <td style="width: 30%; text-align: right; font-size: 13px;">مديرية أوقاف حلب</td>
                <td style="width: 40%; text-align: center; font-size: 16px;">لجنة تقييم أداء المدرس</td>
                <td style="width: 30%; text-align: left; font-size: 13px;">رقم اللجنة: {eval_data['committee_no']}</td>
            </tr>
        </table>
        <table>
            <tr>
                <td class="bg-gray" style="width: 15%; font-weight: bold;">التاريخ:</td>
                <td style="width: 35%;">{eval_data['visit_date']}</td>
                <td class="bg-gray" style="width: 15%; font-weight: bold;">المؤسسة التعليمية الشرعية:</td>
                <td style="width: 35%;">{eval_data['school_name']}</td>
            </tr>
            <tr>
                <td class="bg-gray" style="font-weight: bold;">العام الدراسي:</td>
                <td>{eval_data['academic_year']}</td>
                <td class="bg-gray" style="font-weight: bold;">الفصل الدراسي:</td>
                <td>{eval_data['semester']}</td>
            </tr>
            <tr>
                <td class="bg-gray" style="font-weight: bold;">المدرس:</td>
                <td style="font-weight: bold;">{eval_data['teacher_name']}</td>
                <td class="bg-gray" style="font-weight: bold;">المسؤول العلمي:</td>
                <td>{eval_data['supervisor_name']}</td>
            </tr>
            <tr>
                <td class="bg-gray" style="font-weight: bold;">المادة:</td>
                <td>{eval_data['subject']}</td>
                <td class="bg-gray" style="font-weight: bold;">الاختصاص:</td>
                <td>{eval_data['specialization']}</td>
            </tr>
            <tr>
                <td class="bg-gray" style="font-weight: bold;">عدد الطلاب:</td>
                <td>{eval_data['student_count']}</td>
                <td class="bg-gray" style="font-weight: bold;">الصف:</td>
                <td>{eval_data['grade_level']}</td>
            </tr>
            <tr>
                <td class="bg-gray" style="font-weight: bold;">الشعبة:</td>
                <td>{eval_data['section']}</td>
                <td class="bg-gray" style="font-weight: bold;">الموضوع:</td>
                <td>{eval_data['lesson_topic']}</td>
            </tr>
            <tr>
                <td class="bg-gray" style="font-weight: bold;">الخبرة في التعليم:</td>
                <td>{eval_data['experience']}</td>
                <td class="bg-gray" style="font-weight: bold;">الوضع الوظيفي:</td>
                <td>{eval_data['job_status']}</td>
            </tr>
        </table>
        <table>
            <thead>
                <tr class="bg-gray" style="text-align: center; font-weight: bold; font-size: 11.5px;">
                    <th style="width: 5%;">م</th>
                    <th style="width: 14%;">مجال التقييم</th>
                    <th style="width: 42%;">بنود اللائحة</th>
                    <th style="width: 9%;">الدرجة المستحقة</th>
                    <th style="width: 9%;">الدرجة الفعلية</th>
                    <th style="width: 21%;">الملاحظات</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
        <table>
            <tr class="bg-gray" style="font-weight: bold; text-align: center; font-size: 12px;">
                <td style="width: 50%;">مجموع الدرجات: {eval_data['total_score']} / 100</td>
                <td style="width: 50%;">التقدير النهائي: {eval_data['rating']}</td>
            </tr>
            <tr>
                <td colspan="2" style="text-align: center; font-size: 10px; padding: 2px;">
                    ممتاز: 90–100 | جيد جداً: 80–89 | جيد: 70–79 | مقبول: 50–69 | ضعيف: أقل من 50
                </td>
            </tr>
        </table>
        <table style="border: none; margin-top: 12px;">
            <tr style="font-weight: bold; text-align: center;">
                <td style="border: none; width: 25%;">المدرسة</td>
                <td style="border: none; width: 25%;">الموجّه الاختصاصي</td>
                <td style="border: none; width: 25%;">شعبة التوجيه الاختصاصي</td>
                <td style="border: none; width: 25%;">رئيس إدارة التعليم الشرعي</td>
            </tr>
            <tr style="text-align: center;">
                <td style="border: none; padding-top: 22px;">.....................</td>
                <td style="border: none; padding-top: 22px;">{eval_data['supervisor_name']}</td>
                <td style="border: none; padding-top: 22px;">.....................</td>
                <td style="border: none; padding-top: 22px;">.....................</td>
            </tr>
        </table>
    </body>
    </html>
    """
    return html

def generate_direct_pdf_bytes(eval_data):
    if not WEASYPRINT_AVAILABLE:
        return None
    try:
        html_str = get_evaluation_html(eval_data)
        pdf_bytes = HTML(string=html_str).write_pdf()
        return pdf_bytes
    except Exception:
        return None

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

def generate_annual_executive_report(evals_df, schools_df, users_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        total_schools = len(schools_df)
        visited_schools = len(evals_df['school_name'].unique()) if not evals_df.empty else 0
        coverage_pct = (visited_schools / total_schools * 100) if total_schools > 0 else 0
        
        rating_counts = evals_df['rating'].value_counts() if not evals_df.empty else pd.Series()
        ratings_data = []
        for r in ["ممتاز", "جيد جداً", "جيد", "مقبول", "ضعيف"]:
            cnt = rating_counts.get(r, 0)
            pct = (cnt / len(evals_df) * 100) if len(evals_df) > 0 else 0
            ratings_data.append({"التقدير": r, "العدد": cnt, "النسبة المئوية": f"{pct:.1f}%"})
            
        summary_kpi_df = pd.DataFrame([
            {"المؤشر": "إجمالي عدد المدارس الشرعية", "القيمة": total_schools},
            {"المؤشر": "المدارس التي تمت زيارتها", "القيمة": visited_schools},
            {"المؤشر": "نسبة التغطية الميدانية", "القيمة": f"{coverage_pct:.1f}%"},
            {"المؤشر": "إجمالي الزيارات والتقييمات المنجزة", "القيمة": len(evals_df)},
            {"المؤشر": "متوسط درجات التقييم العام", "القيمة": f"{evals_df['total_score'].mean():.1f} / 100" if not evals_df.empty else "0"}
        ])
        
        summary_kpi_df.to_excel(writer, sheet_name='المؤشرات العامة ونسب التقديرات', index=False, startrow=0)
        pd.DataFrame(ratings_data).to_excel(writer, sheet_name='المؤشرات العامة ونسب التقديرات', index=False, startrow=8)
        
        if not evals_df.empty:
            subj_perf = evals_df.groupby('subject').agg(
                عدد_التقييمات=('id', 'count'),
                متوسط_الدرجة=('total_score', 'mean'),
                أعلى_درجة=('total_score', 'max'),
                أدنى_درجة=('total_score', 'min')
            ).reset_index().rename(columns={'subject': 'المادة الدراسية'})
            subj_perf['متوسط_الدرجة'] = subj_perf['متوسط_الدرجة'].round(1)
            subj_perf.sort_values(by='متوسط_الدرجة', ascending=False).to_excel(writer, sheet_name='مقارنة أداء المواد الدراسية', index=False)
            
        if not evals_df.empty:
            evals_df[['teacher_name', 'school_name', 'subject', 'visit_date', 'total_score', 'rating', 'supervisor_name']].to_excel(
                writer, sheet_name='سجل التقييمات الشامل', index=False
            )

    output.seek(0)
    wb = openpyxl.load_workbook(output)
    for ws in wb.worksheets:
        ws.views.sheetView[0].rightToLeft = True
    excel_final = io.BytesIO()
    wb.save(excel_final)
    excel_final.seek(0)
    return excel_final.getvalue()

# تسجيل الدخول
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
    st.markdown("<div style='height: 40px;'></div>", unsafe_allow_html=True)
    col_l, col_center, col_r = st.columns([1, 2, 1])
    with col_center:
        st.markdown("""
        <div style='text-align: center; margin-bottom: 20px;'>
            <h2 style='font-size: 28px; font-weight: 800; color: #0d5c3a; margin-bottom: 4px;'>🕌 منظومة التقييم الشرعي</h2>
            <p style='font-size: 15px; color: #64748b;'>تسجيل دخول الموجهين والمشرفين</p>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<div class='mobile-card'>", unsafe_allow_html=True)
        u_name = st.text_input("اسم المستخدم", placeholder="اسم المستخدم...")
        u_pass = st.text_input("كلمة المرور", type="password", placeholder="••••••")
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        if st.button("🔑 تسجيل الدخول", use_container_width=True):
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

st.markdown(f"""
<div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; background: #fff; padding: 12px 16px; border-radius: 14px; box-shadow: 0 3px 8px rgba(0,0,0,0.05);'>
    <div>
        <span style='font-size: 18px; font-weight: 800; color: #0d5c3a;'>👤 {st.session_state.user['name']}</span><br>
        <span style='font-size: 13px; color: #64748b;'>{st.session_state.user['specialization']} | {st.session_state.user['role']}</span>
    </div>
</div>
""", unsafe_allow_html=True)

main_menu_options = ["📝 استمارة تقييم جديدة", "🔍 سجل الزيارات والتصدير", "📈 تتبع تطور المدرسين"]
if st.session_state.user["role"] == "Admin":
    main_menu_options = [
        "📊 لوحة المؤشرات والمتابعة",
        "📑 التقرير التركيبي السنوي",
        "📝 استمارة تقييم جديدة",
        "🔍 سجل الزيارات والتصدير",
        "📈 تتبع تطور المدرسين",
        "⚙️ إدارة النظام والحماية"
    ]

c_nav, c_out = st.columns([4, 1])
with c_nav:
    choice = st.radio("القائمة الرئيسية", main_menu_options, horizontal=True, label_visibility="collapsed")
with c_out:
    if st.button("🚪 خروج", use_container_width=True):
        st.session_state.user = None
        st.rerun()

# [باقي أقسام التطبيق تعمل بكفاءة تامة بدون تغيير...]
