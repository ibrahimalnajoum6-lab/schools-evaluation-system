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
            return None
    except Exception:
        return None

# -------------------------------------------------------------
# إعداد الصفحة وتصميم التبويبات الجمالية والانسيابية
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
    }
    
    /* إخفاء القوائم الجانبية والرأسية لشاشات الجوال */
    [data-testid="stSidebar"], [data-testid="stSidebarCollapseButton"], [data-testid="collapsedControl"], header[data-testid="stHeader"] {
        display: none !important;
    }
    
    .block-container {
        padding: 0.6rem !important;
        max-width: 100% !important;
    }
    
    /* كروت ناعمة وعصرية */
    .mobile-card {
        background-color: #ffffff;
        border-radius: 16px;
        padding: 16px;
        margin-bottom: 12px;
        border: 1px solid #edf2f7;
        box-shadow: 0 4px 14px rgba(0,0,0,0.03);
    }
    
    /* بطاقة النتيجة الجمالية العائمة */
    .score-banner {
        background: linear-gradient(135deg, #0d5c3a 0%, #15803d 100%);
        color: white;
        border-radius: 16px;
        padding: 16px;
        text-align: center;
        box-shadow: 0 6px 18px rgba(13,92,58,0.22);
        margin-bottom: 14px;
    }
    
    /* تصميم تبويبات فائق الجمال والنعومة */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background-color: #e2e8f0;
        padding: 5px;
        border-radius: 14px;
        margin-bottom: 14px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 10px;
        padding: 8px 14px;
        font-weight: 700;
        font-size: 13.5px;
        color: #475569;
        background-color: transparent;
        border: none !important;
        transition: all 0.25s ease;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #0d5c3a !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    
    /* بطاقة البند التقييمي */
    .criterion-item {
        background: #ffffff;
        border: 1px solid #f1f5f9;
        border-right: 4px solid #0d5c3a;
        border-radius: 10px;
        padding: 10px 12px;
        margin-bottom: 8px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.02);
    }
    
    /* أزرار مريحة وجذابة */
    .stButton > button {
        border-radius: 12px !important;
        font-weight: 700 !important;
        padding: 12px 18px !important;
        transition: all 0.2s ease;
    }
</style>
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

# -------------------------------------------------------------
# دالة HTML الرسمية للطباعة والحفظ كـ PDF
# -------------------------------------------------------------
def get_evaluation_html(eval_data):
    scores = json.loads(eval_data['scores_json']) if isinstance(eval_data['scores_json'], str) else eval_data['scores_json']
    
    rows_html = ""
    for item in CRITERIA:
        actual = scores.get(str(item['id']), item['max'])
        i_id = item['id']
        
        domain_td = ""
        if i_id == 1:
            domain_td = '<td rowspan="6" style="text-align: center; font-weight: bold; vertical-align: middle; background-color: #f9f9f9;">التخطيط</td>'
        elif i_id == 7:
            domain_td = '<td rowspan="7" style="text-align: center; font-weight: bold; vertical-align: middle; background-color: #f9f9f9;">تنفيذ الدرس</td>'
        elif i_id == 14:
            domain_td = '<td rowspan="4" style="text-align: center; font-weight: bold; vertical-align: middle; background-color: #f9f9f9;">إدارة الصف</td>'
        elif i_id == 18:
            domain_td = '<td rowspan="2" style="text-align: center; font-weight: bold; vertical-align: middle; background-color: #f9f9f9;">شخصية المعلم</td>'
        elif i_id == 20:
            domain_td = '<td rowspan="6" style="text-align: center; font-weight: bold; vertical-align: middle; background-color: #f9f9f9;">المادة العلمية</td>'

        notes_td = ""
        if i_id == 1:
            notes_td = f'<td rowspan="10" style="vertical-align: top; text-align: right; padding: 4px; font-size: 11px;"><b>نقاط التميز:</b><br>{eval_data.get("excellence_points", "")}</td>'
        elif i_id == 11:
            notes_td = f'<td rowspan="9" style="vertical-align: top; text-align: right; padding: 4px; font-size: 11px;"><b>نقاط التطوير:</b><br>{eval_data.get("dev_points", "")}</td>'
        elif i_id == 20:
            notes_td = f'<td rowspan="6" style="vertical-align: top; text-align: right; padding: 4px; font-size: 11px;"><b>المقترحات:</b><br>{eval_data.get("suggestions", "")}</td>'

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
            @media print {{
                body {{ margin: 0; padding: 0; }}
                .no-print {{ display: none !important; }}
            }}
            body {{
                font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
                direction: rtl;
                text-align: right;
                background-color: #fff;
                color: #000;
                padding: 8px;
            }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 4px; }}
            th, td {{ border: 1px solid #000; padding: 2.5px 4px; }}
            .bg-gray {{ background-color: #f2f2f2; }}
            .header-tbl td {{ border: none; font-weight: bold; }}
            .btn-print {{
                background-color: #0d5c3a; color: white; padding: 10px 16px;
                border: none; border-radius: 6px; cursor: pointer;
                font-size: 14px; font-weight: bold; margin-bottom: 10px; width: 100%;
            }}
        </style>
    </head>
    <body>
        <div class="no-print">
            <button class="btn-print" onclick="window.print()">🖨️ اضغط هنا للطباعة أو الحفظ كـ PDF</button>
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

# -------------------------------------------------------------
# دالة توليد استمارة Excel الرسمية
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
# الرأس وشريط التنقل العلوي المبوب
# -------------------------------------------------------------
st.markdown(f"""
<div style='display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; background: #fff; padding: 10px 14px; border-radius: 12px; box-shadow: 0 2px 6px rgba(0,0,0,0.04);'>
    <div>
        <span style='font-size: 15px; font-weight: 800; color: #0d5c3a;'>👤 {st.session_state.user['name']}</span><br>
        <span style='font-size: 11px; color: #64748b;'>{st.session_state.user['specialization']} | {st.session_state.user['role']}</span>
    </div>
</div>
""", unsafe_allow_html=True)

main_menu_options = ["📝 تقييم جديد", "🔍 سجل الزيارات والتعديل"]
if st.session_state.user["role"] == "Admin":
    main_menu_options.append("⚙️ لوحة الإدارة")

c_nav, c_out = st.columns([4, 1])
with c_nav:
    choice = st.radio("القائمة الرئيسية", main_menu_options, horizontal=True, label_visibility="collapsed")
with c_out:
    if st.button("خروج", use_container_width=True):
        st.session_state.user = None
        st.rerun()

# -------------------------------------------------------------
# 1. شاشة استمارة تقييم جديدة بنظام التبويبات الجمالية
# -------------------------------------------------------------
if choice == "📝 تقييم جديد":
    
    # تهيئة درجات البنود
    for item in CRITERIA:
        k = f"q_val_{item['id']}"
        if k not in st.session_state:
            st.session_state[k] = item['max']

    # حساب المجموع والتقدير لحظياً
    total_live = sum(st.session_state[f"q_val_{item['id']}"] for item in CRITERIA)
    rating_live = "ممتاز" if total_live >= 90 else "جيد جداً" if total_live >= 80 else "جيد" if total_live >= 70 else "مقبول" if total_live >= 50 else "ضعيف"

    # بطاقة النتيجة الحية العائمة
    st.markdown(f"""
    <div class='score-banner'>
        <div style='font-size: 12px; opacity: 0.9;'>المجموع الكلي المباشر</div>
        <div style='font-size: 30px; font-weight: 800; margin: 2px 0;'>{total_live} <span style='font-size: 15px; font-weight: 500;'>/ 100</span></div>
        <div style='display: inline-block; background: rgba(255,255,255,0.22); padding: 2px 12px; border-radius: 20px; font-size: 12px; font-weight: 700;'>
            التقدير: {rating_live}
        </div>
    </div>
    """, unsafe_allow_html=True)

    # التبويبات الرئيسية لاستمارة التقييم
    tab_info, tab_domains, tab_notes = st.tabs(["📌 1. بيانات الزيارة", "📋 2. مجالات التقييم", "✍️ 3. الملاحظات والرفع"])

    with tab_info:
        st.markdown("<div class='mobile-card'>", unsafe_allow_html=True)
        conn = sqlite3.connect("evaluation_system.db")
        schools_df = pd.read_sql_query("SELECT name FROM schools", conn)
        conn.close()
        schools_list = schools_df['name'].tolist() if not schools_df.empty else ["لا توجد مدارس"]

        c1, c2 = st.columns(2)
        with c1:
            school_name = st.selectbox("المؤسسة التعليمية الشرعية", schools_list)
            teacher_name = st.text_input("اسم المدرس *", placeholder="أدخل اسم المدرس...")
            subject = st.text_input("المادة", value=st.session_state.user["specialization"])
            grade_level = st.selectbox("الصف", ["السابع", "الثامن", "التاسع", "العاشر", "الحادي عشر", "الثالث الثانوي"])
            section = st.text_input("الشعبة", value="الأولى")
        with c2:
            gender_type = st.selectbox("نوع المدرسة", ["ذكور", "إناث"])
            visit_date = st.date_input("تاريخ الزيارة", value=date.today())
            academic_year = st.selectbox("العام الدراسي", ["2026 - 2027", "2025 - 2026"])
            semester = st.selectbox("الفصل الدراسي", ["الفصل الأول", "الفصل الثاني"])
            committee_no = st.text_input("رقم اللجنة", value="1")

        lesson_topic = st.text_input("موضوع الدرس", placeholder="اكتب موضوع الدرس...")
        
        c3, c4 = st.columns(2)
        with c3:
            specialization = st.text_input("الاختصاص", value=st.session_state.user["specialization"])
            student_count = st.number_input("عدد الطلاب", min_value=1, value=25)
        with c4:
            job_status = st.selectbox("الوضع الوظيفي", ["أصيل", "وكيل", "مكلف"])
            experience = st.text_input("الخبرة في التعليم", value="5 سنوات")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_domains:
        # تبويبات داخلية للمجالات الخمسة
        domain_tabs = st.tabs(DOMAINS)
        
        for idx, dom in enumerate(DOMAINS):
            with domain_tabs[idx]:
                dom_items = [i for i in CRITERIA if i['domain'] == dom]
                for it in dom_items:
                    st.markdown(f"""
                    <div class='criterion-item'>
                        <div style='font-weight: 700; font-size: 13.5px; color: #1e293b;'>{it['id']}. {it['text']}</div>
                        <div style='font-size: 11px; color: #64748b;'>الدرجة القصوى المستحقة: ({it['max']})</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    scores_options = list(range(it['max'] + 1))
                    st.select_slider(
                        f"درجة بند {it['id']}",
                        options=scores_options,
                        value=st.session_state[f"q_val_{it['id']}"],
                        key=f"slider_val_{it['id']}",
                        on_change=lambda i=it['id']: st.session_state.update({f"q_val_{i}": st.session_state[f"slider_val_{i}"]}),
                        label_visibility="collapsed"
                    )

    with tab_notes:
        st.markdown("<div class='mobile-card'>", unsafe_allow_html=True)
        excellence_points = st.text_area("🌟 نقاط التميز", placeholder="اكتب نقاط القوة والتميز...")
        dev_points = st.text_area("💡 نقاط التطوير", placeholder="اكتب نقاط التطوير والتحسين...")
        suggestions = st.text_area("📌 المقترحات والتوصيات", placeholder="اكتب المقترحات والتوجيهات...")
        
        uploaded_files = st.file_uploader("📷 رفع شواهد وصور / مقطع فيديو من الهاتف", accept_multiple_files=True)
        st.markdown("</div>", unsafe_allow_html=True)

        if st.button("💾 اعتماد وحفظ الاستمارة ورفع الملفات إلى جوجل درايف", type="primary", use_container_width=True):
            if not teacher_name.strip():
                st.error("⚠️ يرجى إدخال اسم المدرس في تبويب البيانات الأساسية.")
            else:
                final_scores_dict = {str(item['id']): st.session_state[f"q_val_{item['id']}"] for item in CRITERIA}
                saved_media = []
                drive_links = []
                
                # رفع الشواهد المرفقة إلى جوجل درايف
                if uploaded_files:
                    os.makedirs("uploads", exist_ok=True)
                    for f in uploaded_files:
                        f_bytes = f.getbuffer()
                        path = os.path.join("uploads", f.name)
                        with open(path, "wb") as file_out:
                            file_out.write(f_bytes)
                        saved_media.append(path)
                        
                        with st.spinner(f"جاري رفع {f.name} إلى درايف..."):
                            link = upload_file_to_drive(f.getvalue(), f.name, school_name, visit_date, f.type)
                            if link:
                                drive_links.append(link)

                eval_data_dict = {
                    "committee_no": committee_no, "visit_date": str(visit_date), "academic_year": academic_year,
                    "semester": semester, "school_name": school_name, "gender_type": gender_type,
                    "supervisor_name": st.session_state.user["name"], "teacher_name": teacher_name, "subject": subject,
                    "specialization": specialization, "student_count": student_count, "grade_level": grade_level,
                    "section": section, "lesson_topic": lesson_topic, "job_status": job_status,
                    "experience": experience, "scores_json": json.dumps(final_scores_dict), "total_score": total_live,
                    "rating": rating_live, "excellence_points": excellence_points, "dev_points": dev_points,
                    "suggestions": suggestions
                }
                
                # إنشاء ورفع استمارة الإكسل الرسمية إلى درايف
                with st.spinner("جاري إنشاء ورفع تقرير الإكسل الرسمي..."):
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
                    st.session_state.user["name"], teacher_name, subject, specialization, student_count,
                    grade_level, section, lesson_topic, job_status, experience,
                    json.dumps(final_scores_dict), total_live, rating_live, excellence_points, dev_points, suggestions,
                    ",".join(saved_media), ",".join(drive_links), "معتمد"
                ))
                conn.commit()
                conn.close()
                st.success(f"✅ تم حفظ الاستمارة ورفع الملفات إلى مجلد ({school_name} / {visit_date.month}-{visit_date.year}) في جوجل درايف بنجاح!")

# -------------------------------------------------------------
# 2. سجل الزيارات والمعاينة والتعديل بعد الإرسال
# -------------------------------------------------------------
elif choice == "🔍 سجل الزيارات والتعديل":
    st.markdown("### 🔍 سجل الزيارات واستمارات التقييم")
    
    conn = sqlite3.connect("evaluation_system.db")
    if st.session_state.user["role"] == "Admin":
        df = pd.read_sql_query("SELECT * FROM evaluations ORDER BY id DESC", conn)
    else:
        df = pd.read_sql_query("SELECT * FROM evaluations WHERE supervisor_name=? ORDER BY id DESC", conn, params=(st.session_state.user["name"],))
    conn.close()
    
    if df.empty:
        st.info("لا توجد استمارات مسجلة حتى الآن.")
    else:
        search_txt = st.text_input("🔍 بحث سريع باسم المدرس أو المدرسة", placeholder="اكتب للبحث...")
        filtered_df = df.copy()
        if search_txt:
            filtered_df = filtered_df[
                filtered_df["teacher_name"].str.contains(search_txt, na=False) |
                filtered_df["school_name"].str.contains(search_txt, na=False)
            ]

        for _, row in filtered_df.iterrows():
            with st.expander(f"📄 {row['teacher_name']} — {row['school_name']} ({row['total_score']} / 100 - {row['rating']})"):
                st.write(f"**المادة:** {row['subject']} | **الصف:** {row['grade_level']} | **التاريخ:** {row['visit_date']}")
                st.write(f"**المسؤول العلمي / الموجه:** {row['supervisor_name']}")
                
                rec_dict = row.to_dict()
                
                # تبويبات داخل تفاصيل كل استمارة (تعديل / تصدير وطباعة PDF / درايف)
                sub_tab1, sub_tab2, sub_tab3 = st.tabs(["✏️ تعديل الاستمارة", "🖨️ طباعة وتصدير (PDF/Excel)", "📂 روابط درايف"])
                
                with sub_tab1:
                    e_tname = st.text_input("اسم المدرس", value=row["teacher_name"], key=f"et_{row['id']}")
                    e_subj = st.text_input("المادة", value=row["subject"], key=f"es_{row['id']}")
                    e_topic = st.text_input("موضوع الدرس", value=row["lesson_topic"], key=f"etp_{row['id']}")
                    
                    col_e1, col_e2 = st.columns(2)
                    with col_e1:
                        e_cnt = st.number_input("عدد الطلاب", value=int(row["student_count"]), key=f"ecn_{row['id']}")
                        e_sec = st.text_input("الشعبة", value=row["section"], key=f"esc_{row['id']}")
                    with col_e2:
                        e_status = st.selectbox("الوضع الوظيفي", ["أصيل", "وكيل", "مكلف"], index=["أصيل", "وكيل", "مكلف"].index(row["job_status"]), key=f"est_{row['id']}")
                        e_exp = st.text_input("الخبرة", value=row["experience"], key=f"eex_{row['id']}")
                    
                    st.markdown("**تعديل درجات البنود الـ 25:**")
                    curr_scores = json.loads(row["scores_json"]) if isinstance(row["scores_json"], str) else row["scores_json"]
                    updated_scores = {}
                    
                    for it in CRITERIA:
                        val_now = int(curr_scores.get(str(it['id']), it['max']))
                        updated_scores[str(it['id'])] = st.number_input(
                            f"{it['id']}. {it['text']} (من {it['max']})",
                            min_value=0, max_value=it['max'], value=val_now,
                            key=f"edit_sc_{row['id']}_{it['id']}"
                        )
                        
                    e_exc = st.text_area("نقاط التميز", value=row["excellence_points"], key=f"eex2_{row['id']}")
                    e_dev = st.text_area("نقاط التطوير", value=row["dev_points"], key=f"edv_{row['id']}")
                    e_sug = st.text_area("المقترحات", value=row["suggestions"], key=f"esg_{row['id']}")
                    
                    e_files = st.file_uploader("📷 رفع شواهد إضافية", accept_multiple_files=True, key=f"ef_{row['id']}")

                    if st.button("💾 حفظ وتحديث التعديلات", type="primary", use_container_width=True, key=f"btn_edit_{row['id']}"):
                        new_tot = sum(updated_scores.values())
                        new_rat = "ممتاز" if new_tot >= 90 else "جيد جداً" if new_tot >= 80 else "جيد" if new_tot >= 70 else "مقبول" if new_tot >= 50 else "ضعيف"
                        
                        existing_links = str(row["drive_links"]).split(',') if row["drive_links"] else []
                        if e_files:
                            for ef in e_files:
                                link = upload_file_to_drive(ef.getvalue(), ef.name, row["school_name"], date.today(), ef.type)
                                if link:
                                    existing_links.append(link)

                        conn = sqlite3.connect("evaluation_system.db")
                        c = conn.cursor()
                        c.execute('''UPDATE evaluations SET 
                            teacher_name=?, subject=?, lesson_topic=?, student_count=?, section=?, job_status=?, experience=?,
                            scores_json=?, total_score=?, rating=?, excellence_points=?, dev_points=?, suggestions=?, drive_links=?
                            WHERE id=?''', (
                                e_tname, e_subj, e_topic, e_cnt, e_sec, e_status, e_exp,
                                json.dumps(updated_scores), new_tot, new_rat, e_exc, e_dev, e_sug, ",".join(existing_links), row["id"]
                            ))
                        conn.commit()
                        conn.close()
                        st.success("✅ تم تحديث الاستمارة والدرجات بنجاح!")
                        st.rerun()

                with sub_tab2:
                    excel_form_bytes = generate_evaluation_excel_form(rec_dict)
                    st.download_button(
                        label="📊 تحميل الاستمارة الرسمية (Excel)",
                        data=excel_form_bytes,
                        file_name=f"استمارة_{row['teacher_name']}_{row['visit_date']}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"dl_{row['id']}"
                    )
                    st.markdown("---")
                    st.markdown("#### 🖨️ استمارة التقييم الرسمية (PDF)")
                    html_preview = get_evaluation_html(rec_dict)
                    components.html(html_preview, height=750, scrolling=True)

                with sub_tab3:
                    if row.get("drive_links"):
                        st.markdown("📂 **الملفات المحفوظة على Google Drive:**")
                        for i, link in enumerate(str(row["drive_links"]).split(','), 1):
                            if link.strip():
                                st.markdown(f"- [عرض الملف {i} على درايف]({link.strip()})")
                    else:
                        st.info("لا توجد ملفات مرفوعة لهذه الاستمارة.")

# -------------------------------------------------------------
# 3. لوحة الإدارة الكاملة (Admin)
# -------------------------------------------------------------
elif choice == "⚙️ لوحة الإدارة":
    st.markdown("### ⚙️ إدارة النظام والمستخدمين والمدارس")
    admin_tab1, admin_tab2 = st.tabs(["👥 إدارة الموجهين والحسابات", "🏫 إدارة المدارس"])
    
    # ------------------ تبويب 1: إدارة الموجهين ------------------
    with admin_tab1:
        conn = sqlite3.connect("evaluation_system.db")
        users_df = pd.read_sql_query("SELECT id, username, full_name, specialization, role FROM users", conn)
        conn.close()
        
        st.dataframe(users_df, use_container_width=True)
        
        st.markdown("#### ✏️ تعديل بيانات مستخدم / موجه")
        selected_u_id = st.selectbox(
            "اختر المستخدم المراد تعديله",
            users_df["id"].tolist(),
            format_func=lambda x: f"ID {x}: {users_df[users_df['id'] == x]['full_name'].values[0]} ({users_df[users_df['id'] == x]['username'].values[0]})"
        )
        
        if selected_u_id:
            conn = sqlite3.connect("evaluation_system.db")
            c = conn.cursor()
            c.execute("SELECT id, username, password, full_name, specialization, role FROM users WHERE id=?", (selected_u_id,))
            u_row = c.fetchone()
            conn.close()
            
            with st.form(f"edit_u_form_{selected_u_id}"):
                e_fname = st.text_input("الاسم الكامل", value=u_row[3])
                e_uname = st.text_input("اسم المستخدم (Username)", value=u_row[1])
                e_pword = st.text_input("كلمة المرور", value=u_row[2])
                e_spec = st.text_input("التخصص", value=u_row[4])
                e_role = st.selectbox("الدور / الصلاحية", ["Supervisor", "Admin"], index=0 if u_row[5] == "Supervisor" else 1)
                
                b_c1, b_c2 = st.columns(2)
                with b_c1:
                    save_u_btn = st.form_submit_button("💾 حفظ التعديلات", use_container_width=True)
                with b_c2:
                    del_u_btn = st.form_submit_button("🗑️ حذف الحساب", use_container_width=True)
                    
                if save_u_btn:
                    try:
                        conn = sqlite3.connect("evaluation_system.db")
                        c = conn.cursor()
                        c.execute("UPDATE users SET username=?, password=?, full_name=?, specialization=?, role=? WHERE id=?",
                                  (e_uname, e_pword, e_fname, e_spec, e_role, selected_u_id))
                        conn.commit()
                        conn.close()
                        st.success(f"✅ تم تحديث بيانات ({e_fname}) بنجاح!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("❌ اسم المستخدم هذا موجود بالفعل.")

                if del_u_btn:
                    if selected_u_id == st.session_state.user["id"]:
                        st.error("⚠️ لا يمكنك حذف حسابك الحالي!")
                    else:
                        conn = sqlite3.connect("evaluation_system.db")
                        c = conn.cursor()
                        c.execute("DELETE FROM users WHERE id=?", (selected_u_id,))
                        conn.commit()
                        conn.close()
                        st.warning("🗑️ تم حذف المستخدم.")
                        st.rerun()

        st.markdown("---")
        with st.expander("➕ إضافة موجه / مستخدم جديد"):
            with st.form("add_user_mobile"):
                n_name = st.text_input("الاسم الكامل")
                n_user = st.text_input("اسم المستخدم للدخول")
                n_pass = st.text_input("كلمة المرور")
                n_spec = st.text_input("التخصص")
                n_role = st.selectbox("الدور", ["Supervisor", "Admin"])
                if st.form_submit_button("➕ إنشاء الحساب", use_container_width=True):
                    if n_name and n_user and n_pass:
                        try:
                            conn = sqlite3.connect("evaluation_system.db")
                            c = conn.cursor()
                            c.execute("INSERT INTO users (username, password, full_name, specialization, role) VALUES (?, ?, ?, ?, ?)",
                                      (n_user, n_pass, n_name, n_spec, n_role))
                            conn.commit()
                            conn.close()
                            st.success("✅ تمت الإضافة بنجاح!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("❌ اسم المستخدم موجود مسبقاً.")

    # ------------------ تبويب 2: إدارة وتعديل المدارس ------------------
    with admin_tab2:
        conn = sqlite3.connect("evaluation_system.db")
        schools_all = pd.read_sql_query("SELECT id, name, gender, location FROM schools", conn)
        conn.close()
        
        st.dataframe(schools_all, use_container_width=True)
        
        st.markdown("#### ✏️ تعديل بيانات مدرسة")
        selected_s_id = st.selectbox(
            "اختر المدرسة للتعديل",
            schools_all["id"].tolist(),
            format_func=lambda x: f"ID {x}: {schools_all[schools_all['id'] == x]['name'].values[0]}"
        )
        
        if selected_s_id:
            conn = sqlite3.connect("evaluation_system.db")
            c = conn.cursor()
            c.execute("SELECT id, name, gender, location FROM schools WHERE id=?", (selected_s_id,))
            s_row = c.fetchone()
            conn.close()
            
            with st.form(f"edit_school_form_{selected_s_id}"):
                es_name = st.text_input("اسم المدرسة", value=s_row[1])
                es_gender = st.selectbox("النوع", ["ذكور", "إناث"], index=0 if s_row[2] == "ذكور" else 1)
                es_loc = st.text_input("الموقع الجغرافي", value=s_row[3])
                
                sb_c1, sb_c2 = st.columns(2)
                with sb_c1:
                    save_s_btn = st.form_submit_button("💾 حفظ تعديل المدرسة", use_container_width=True)
                with sb_c2:
                    del_s_btn = st.form_submit_button("🗑️ حذف المدرسة", use_container_width=True)
                    
                if save_s_btn:
                    conn = sqlite3.connect("evaluation_system.db")
                    c = conn.cursor()
                    c.execute("UPDATE schools SET name=?, gender=?, location=? WHERE id=?", (es_name, es_gender, es_loc, selected_s_id))
                    conn.commit()
                    conn.close()
                    st.success("✅ تم تحديث بيانات المدرسة بنجاح!")
                    st.rerun()

                if del_s_btn:
                    conn = sqlite3.connect("evaluation_system.db")
                    c = conn.cursor()
                    c.execute("DELETE FROM schools WHERE id=?", (selected_s_id,))
                    conn.commit()
                    conn.close()
                    st.warning("🗑️ تم حذف المدرسة.")
                    st.rerun()

        st.markdown("---")
        with st.expander("➕ إضافة مدرسة جديدة"):
            with st.form("add_school_mobile"):
                s_name = st.text_input("اسم المدرسة")
                s_gen = st.selectbox("النوع", ["ذكور", "إناث"])
                s_loc = st.text_input("الموقع الجغرافي")
                if st.form_submit_button("➕ إضافة المدرسة", use_container_width=True):
                    if s_name:
                        conn = sqlite3.connect("evaluation_system.db")
                        c = conn.cursor()
                        c.execute("INSERT INTO schools (name, gender, location) VALUES (?, ?, ?)", (s_name, s_gen, s_loc))
                        conn.commit()
                        conn.close()
                        st.success("✅ تمت إضافة المدرسة بنجاح!")
                        st.rerun()
