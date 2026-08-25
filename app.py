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
        
        response = requests.post(APPS_SCRIPT_URL, json=payload)
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
# إعداد الصفحة وتجاوبها مع الجوال
# -------------------------------------------------------------
st.set_page_config(
    page_title="منظومة تقييم المدارس الشرعية",
    page_icon="🕌",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    html, body, [class*="css"] {
        direction: rtl;
        text-align: right;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
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
# دالة HTML مطابقة للاستمارة الرسمية
# -------------------------------------------------------------
def get_evaluation_html(eval_data):
    scores = json.loads(eval_data['scores_json']) if isinstance(eval_data['scores_json'], str) else eval_data['scores_json']
    
    rows_html = ""
    for item in CRITERIA:
        actual = scores.get(str(item['id']), item['max'])
        i_id = item['id']
        
        domain_td = ""
        if i_id == 1:
            domain_td = '<td rowspan="6" style="text-align: center; font-weight: bold; vertical-align: middle;">التخطيط</td>'
        elif i_id == 7:
            domain_td = '<td rowspan="7" style="text-align: center; font-weight: bold; vertical-align: middle;">تنفيذ الدرس</td>'
        elif i_id == 14:
            domain_td = '<td rowspan="4" style="text-align: center; font-weight: bold; vertical-align: middle;">إدارة الصف</td>'
        elif i_id == 18:
            domain_td = '<td rowspan="2" style="text-align: center; font-weight: bold; vertical-align: middle;">شخصية المعلم</td>'
        elif i_id == 20:
            domain_td = '<td rowspan="6" style="text-align: center; font-weight: bold; vertical-align: middle;">المادة العلمية</td>'

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
                body {{
                    margin: 0;
                    padding: 0;
                }}
                .no-print {{
                    display: none !important;
                }}
            }}
            body {{
                font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
                direction: rtl;
                text-align: right;
                background-color: #fff;
                color: #000;
                padding: 10px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 4px;
            }}
            th, td {{
                border: 1px solid #000;
                padding: 2.5px 4px;
            }}
            .bg-gray {{
                background-color: #f2f2f2;
            }}
            .header-tbl td {{
                border: none;
                font-weight: bold;
            }}
            .btn-print {{
                background-color: #0066cc;
                color: white;
                padding: 8px 16px;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 13px;
                font-weight: bold;
                margin-bottom: 10px;
            }}
        </style>
    </head>
    <body>
        <div class="no-print" style="text-align: left;">
            <button class="btn-print" onclick="window.print()">🖨️ طباعة الاستمارة / حفظ كـ PDF</button>
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
    st.title("🕌 منظومة تقييم أداء المدرسين - مديرية التعليم الشرعي")
    st.subheader("تسجيل الدخول")
    col1, col2 = st.columns([1, 1])
    with col1:
        u_name = st.text_input("اسم المستخدم")
        u_pass = st.text_input("كلمة المرور", type="password")
        if st.button("دخول", use_container_width=True):
            user_data = login(u_name, u_pass)
            if user_data:
                st.session_state.user = {
                    "id": user_data[0], "username": user_data[1],
                    "name": user_data[2], "specialization": user_data[3], "role": user_data[4]
                }
                st.rerun()
            else:
                st.error("اسم المستخدم أو كلمة المرور غير صحيحة")
    st.stop()

# -------------------------------------------------------------
# القائمة الجانبية
# -------------------------------------------------------------
st.sidebar.title(f"👤 {st.session_state.user['name']}")
st.sidebar.caption(f"الصفة: {st.session_state.user['role']} | التخصص: {st.session_state.user['specialization']}")

menu_options = ["استمارة تقييم جديدة", "سجل الزيارات والبحث والتعديل والتصدير"]
if st.session_state.user["role"] == "Admin":
    menu_options.append("لوحة التحكم والإدارة")

choice = st.sidebar.radio("القائمة الرئيسية", menu_options)

if st.sidebar.button("تسجيل الخروج"):
    st.session_state.user = None
    st.rerun()

# -------------------------------------------------------------
# 1. شاشة استمارة تقييم جديدة
# -------------------------------------------------------------
if choice == "استمارة تقييم جديدة":
    st.title("📋 استمارة تقييم أداء المدرس")
    
    with st.expander("📌 البيانات الأساسية للزيارة", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            gender_type = st.radio("نوع المدرسة", ["ذكور", "إناث"], horizontal=True)
            conn = sqlite3.connect("evaluation_system.db")
            schools_df = pd.read_sql_query("SELECT name FROM schools WHERE gender=?", conn, params=(gender_type,))
            conn.close()
            
            schools_list = schools_df["name"].tolist() if not schools_df.empty else ["لا توجد مدارس مسجلة"]
            school_name = st.selectbox("المؤسسة التعليمية الشرعية", schools_list)
            committee_no = st.text_input("رقم اللجنة", value="1")
            
        with c2:
            visit_date = st.date_input("تاريخ الزيارة", value=date.today())
            academic_year = st.selectbox("العام الدراسي", ["2026 - 2027", "2025 - 2026"])
            semester = st.selectbox("الفصل الدراسي", ["الفصل الأول", "الفصل الثاني"])
            
        with c3:
            if st.session_state.user["role"] == "Admin":
                conn = sqlite3.connect("evaluation_system.db")
                sups = pd.read_sql_query("SELECT full_name FROM users WHERE role='Supervisor'", conn)["full_name"].tolist()
                conn.close()
                sups_list = sups if sups else [st.session_state.user["name"]]
                supervisor_name = st.selectbox("المسؤول العلمي / الموجه", sups_list)
            else:
                supervisor_name = st.text_input("المسؤول العلمي / الموجه", value=st.session_state.user["name"], disabled=True)
            
            teacher_name = st.text_input("اسم المدرس")
            subject = st.text_input("المادة", value=st.session_state.user["specialization"])

        c4, c5, c6 = st.columns(3)
        with c4:
            specialization = st.text_input("الاختصاص")
            grade_level = st.selectbox("الصف", ["السابع", "الثامن", "التاسع", "العاشر", "الحادي عشر", "الثالث الثانوي"])
        with c5:
            job_status = st.selectbox("الوضع الوظيفي", ["أصيل", "وكيل", "مكلف"])
            section = st.text_input("الشعبة", value="الأولى")
        with c6:
            experience = st.text_input("الخبرة في التعليم", value="5 سنوات")
            student_count = st.number_input("عدد الطلاب", min_value=1, value=25)
            
        lesson_topic = st.text_input("موضوع الدرس")

    st.subheader("📊 بنود التقييم وتوزيع الدرجات")
    scores = {}
    current_domain = ""
    for item in CRITERIA:
        if item["domain"] != current_domain:
            current_domain = item["domain"]
            st.markdown(f"#### 🔹 مجال: {current_domain}")
            
        col_text, col_score = st.columns([4, 1])
        with col_text:
            st.write(f"**{item['id']}. {item['text']}**")
        with col_score:
            scores[str(item['id'])] = st.number_input(
                f"الدرجة (من {item['max']})", 0, item['max'], item['max'], key=f"new_q_{item['id']}"
            )

    total_score = sum(scores.values())
    rating = "ممتاز" if total_score >= 90 else "جيد جداً" if total_score >= 80 else "جيد" if total_score >= 70 else "مقبول" if total_score >= 50 else "ضعيف"
    st.info(f"🎯 المجموع الكلي: **{total_score} / 100** | التقدير النهائي: **{rating}**")

    excellence_points = st.text_area("نقاط التميز")
    dev_points = st.text_area("نقاط التطوير")
    suggestions = st.text_area("المقترحات والتوصيات")
    uploaded_files = st.file_uploader("رفع شواهد وصور / مقطع فيديو", accept_multiple_files=True)

    if st.button("💾 حفظ الاستمارة ورفع الملفات إلى جوجل درايف", type="primary", use_container_width=True):
        if not teacher_name:
            st.error("يرجى إدخال اسم المدرس.")
        else:
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
                    
                    with st.spinner(f"جاري رفع الملف {f.name} إلى درايف..."):
                        link = upload_file_to_drive(f.getvalue(), f.name, school_name, visit_date, f.type)
                        if link:
                            drive_links.append(link)

            eval_data_dict = {
                "committee_no": committee_no, "visit_date": str(visit_date), "academic_year": academic_year,
                "semester": semester, "school_name": school_name, "gender_type": gender_type,
                "supervisor_name": supervisor_name, "teacher_name": teacher_name, "subject": subject,
                "specialization": specialization, "student_count": student_count, "grade_level": grade_level,
                "section": section, "lesson_topic": lesson_topic, "job_status": job_status,
                "experience": experience, "scores_json": json.dumps(scores), "total_score": total_score,
                "rating": rating, "excellence_points": excellence_points, "dev_points": dev_points,
                "suggestions": suggestions
            }
            
            with st.spinner("جاري إنشاء ورفع استمارة الإكسل الرسمية إلى درايف..."):
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
                supervisor_name, teacher_name, subject, specialization, student_count,
                grade_level, section, lesson_topic, job_status, experience,
                json.dumps(scores), total_score, rating, excellence_points, dev_points, suggestions,
                ",".join(saved_media), ",".join(drive_links), "معتمد"
            ))
            conn.commit()
            conn.close()
            st.success(f"✅ تم حفظ الاستمارة ورفع الملفات إلى مجلد ({school_name} / {visit_date.month}-{visit_date.year}) في جوجل درايف بنجاح!")

# -------------------------------------------------------------
# 2. شاشة السجل والبحث والتعديل وتصدير PDF / Excel
# -------------------------------------------------------------
elif choice == "سجل الزيارات والبحث والتعديل والتصدير":
    st.title("🔍 سجل الزيارات والتعديل وتصدير التقارير الرسمية")
    
    conn = sqlite3.connect("evaluation_system.db")
    if st.session_state.user["role"] == "Admin":
        df = pd.read_sql_query("SELECT * FROM evaluations", conn)
    else:
        df = pd.read_sql_query("SELECT * FROM evaluations WHERE supervisor_name=?", conn, params=(st.session_state.user["name"],))
    conn.close()
    
    if df.empty:
        st.info("لا توجد استمارات مسجلة بعد.")
    else:
        f1, f2, f3 = st.columns(3)
        with f1:
            search_teacher = st.text_input("بحث باسم المدرس")
        with f2:
            search_school = st.selectbox("تصفية بالمدرسة", ["الكل"] + df["school_name"].unique().tolist())
        with f3:
            search_rating = st.selectbox("تصفية بالتقدير", ["الكل", "ممتاز", "جيد جداً", "جيد", "مقبول", "ضعيف"])
            
        filtered_df = df.copy()
        if search_teacher:
            filtered_df = filtered_df[filtered_df["teacher_name"].str.contains(search_teacher, na=False)]
        if search_school != "الكل":
            filtered_df = filtered_df[filtered_df["school_name"] == search_school]
        if search_rating != "الكل":
            filtered_df = filtered_df[filtered_df["rating"] == search_rating]

        st.dataframe(filtered_df[["id", "teacher_name", "school_name", "subject", "visit_date", "total_score", "rating", "supervisor_name"]], use_container_width=True)

        out_excel = io.BytesIO()
        with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, index=False, sheet_name='سجل التقييمات')
        st.download_button(
            label="📥 تحميل السجل المجمع لجميع المدرسين (Excel)",
            data=out_excel.getvalue(),
            file_name=f"سجل_تقييمات_المدارس_الشرعية_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.markdown("---")
        st.subheader("📑 استعراض / تعديل استمارة واستخراج النماذج الرسمية")
        selected_id = st.selectbox("اختر رقم الاستمارة (ID)", filtered_df["id"].tolist())
        
        if selected_id:
            record = filtered_df[filtered_df["id"] == selected_id].iloc[0]
            rec_dict = record.to_dict()
            
            excel_form_bytes = generate_evaluation_excel_form(rec_dict)
            st.download_button(
                label=f"📊 تحميل استمارة ({record['teacher_name']}) بصيغة Excel رسمية",
                data=excel_form_bytes,
                file_name=f"استمارة_{record['teacher_name']}_{record['visit_date']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
            if "drive_links" in record and record["drive_links"]:
                st.markdown("📂 **الملفات المحفوظة على Google Drive:**")
                for i, link in enumerate(str(record["drive_links"]).split(','), 1):
                    if link.strip():
                        st.markdown(f"- [رابط الملف {i} على درايف]({link.strip()})")

            st.markdown("#### 📄 معاينة وطباعة الاستمارة الرسمية (PDF)")
            html_doc = get_evaluation_html(rec_dict)
            components.html(html_doc, height=950, scrolling=True)
            
            st.markdown("#### ✏️ تعديل بيانات الاستمارة والدرجات بعد الإرسال")
            with st.expander("فتح نموذج التعديل", expanded=False):
                with st.form("edit_full_form"):
                    e_c1, e_c2, e_c3 = st.columns(3)
                    with e_c1:
                        new_teacher = st.text_input("اسم المدرس", value=record["teacher_name"])
                        new_subject = st.text_input("المادة", value=record["subject"])
                    with e_c2:
                        new_topic = st.text_input("الموضوع", value=record["lesson_topic"])
                        new_count = st.number_input("عدد الطلاب", value=int(record["student_count"]))
                    with e_c3:
                        new_status = st.selectbox("الوضع الوظيفي", ["أصيل", "وكيل", "مكلف"], index=["أصيل", "وكيل", "مكلف"].index(record["job_status"]))
                    
                    st.write("**تعديل درجات البنود الـ 25:**")
                    current_saved_scores = json.loads(record["scores_json"]) if isinstance(record["scores_json"], str) else record["scores_json"]
                    updated_scores = {}
                    
                    for item in CRITERIA:
                        c_val = int(current_saved_scores.get(str(item['id']), item['max']))
                        updated_scores[str(item['id'])] = st.number_input(
                            f"{item['id']}. {item['text']} (من {item['max']})", 0, item['max'], c_val, key=f"edit_q_{item['id']}"
                        )
                    
                    new_total = sum(updated_scores.values())
                    new_rating = "ممتاز" if new_total >= 90 else "جيد جداً" if new_total >= 80 else "جيد" if new_total >= 70 else "مقبول" if new_total >= 50 else "ضعيف"
                    
                    new_excellence = st.text_area("نقاط التميز", value=record["excellence_points"])
                    new_dev = st.text_area("نقاط التطوير", value=record["dev_points"])
                    new_suggestions = st.text_area("المقترحات والتوصيات", value=record["suggestions"])
                    
                    if st.form_submit_button("💾 تحديث وحفظ التعديلات في قاعدة البيانات", use_container_width=True):
                        conn = sqlite3.connect("evaluation_system.db")
                        c = conn.cursor()
                        c.execute('''UPDATE evaluations SET 
                            teacher_name=?, subject=?, lesson_topic=?, student_count=?, job_status=?,
                            scores_json=?, total_score=?, rating=?, excellence_points=?, dev_points=?, suggestions=?
                            WHERE id=?''', (
                                new_teacher, new_subject, new_topic, new_count, new_status,
                                json.dumps(updated_scores), new_total, new_rating,
                                new_excellence, new_dev, new_suggestions, selected_id
                            ))
                        conn.commit()
                        conn.close()
                        st.success("✅ تم تحديث الاستمارة والدرجات بنجاح!")
                        st.rerun()

# -------------------------------------------------------------
# 3. لوحة الإدارة الموسعة (تعديل المدارس، الموجهين، الحسابات)
# -------------------------------------------------------------
elif choice == "لوحة التحكم والإدارة":
    st.title("⚙️ لوحة إدارة النظام والمستخدمين والمدارس")
    
    tab1, tab2 = st.tabs(["👥 إدارة الموجهين والمستخدمين", "🏫 إدارة المدارس والمؤسسات"])
    
    # ------------------ تبويب 1: إدارة الموجهين والمستخدمين ------------------
    with tab1:
        st.subheader("📋 قائمة المستخدمين والموجهين المسجلين")
        conn = sqlite3.connect("evaluation_system.db")
        users_df = pd.read_sql_query("SELECT id, username, full_name, specialization, role FROM users", conn)
        conn.close()
        
        st.dataframe(users_df, use_container_width=True)
        
        col_edit_u, col_add_u = st.columns(2)
        
        # قسم تعديل بيانات موجه أو كلمة سره أو اسم المستخدم
        with col_edit_u:
            st.markdown("### ✏️ تعديل بيانات مستخدم / موجه")
            selected_user_id = st.selectbox("اختر المستخدم للتعديل", users_df["id"].tolist(), format_func=lambda x: f"ID: {x} - {users_df[users_df['id'] == x]['full_name'].values[0]} ({users_df[users_df['id'] == x]['username'].values[0]})")
            
            if selected_user_id:
                conn = sqlite3.connect("evaluation_system.db")
                c = conn.cursor()
                c.execute("SELECT id, username, password, full_name, specialization, role FROM users WHERE id=?", (selected_user_id,))
                user_row = c.fetchone()
                conn.close()
                
                with st.form("form_edit_user"):
                    edit_full_name = st.text_input("الاسم الكامل للموجه", value=user_row[3])
                    edit_username = st.text_input("اسم المستخدم (Username)", value=user_row[1])
                    edit_password = st.text_input("كلمة المرور الجديدة", value=user_row[2], type="password")
                    edit_spec = st.text_input("التخصص / المادة", value=user_row[4])
                    edit_role = st.selectbox("الصلاحية / الدور", ["Supervisor", "Admin"], index=0 if user_row[5] == "Supervisor" else 1)
                    
                    c_btn1, c_btn2 = st.columns(2)
                    with c_btn1:
                        submit_edit_u = st.form_submit_button("💾 حفظ التعديلات", use_container_width=True)
                    with c_btn2:
                        delete_u = st.form_submit_button("🗑️ حذف المستخدم", use_container_width=True)
                        
                    if submit_edit_u:
                        try:
                            conn = sqlite3.connect("evaluation_system.db")
                            c = conn.cursor()
                            c.execute('''UPDATE users SET username=?, password=?, full_name=?, specialization=?, role=? WHERE id=?''',
                                      (edit_username, edit_password, edit_full_name, edit_spec, edit_role, selected_user_id))
                            conn.commit()
                            conn.close()
                            st.success(f"✅ تم تحديث بيانات المستخدم ({edit_full_name}) بنجاح!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("❌ اسم المستخدم هذا مستخدم مسبقاً، يرجى اختيار اسم آخر.")

                    if delete_u:
                        if selected_user_id == st.session_state.user["id"]:
                            st.error("⚠️ لا يمكنك حذف الحساب الذي تستخدمه حالياً لتسجيل الدخول!")
                        else:
                            conn = sqlite3.connect("evaluation_system.db")
                            c = conn.cursor()
                            c.execute("DELETE FROM users WHERE id=?", (selected_user_id,))
                            conn.commit()
                            conn.close()
                            st.warning("🗑️ تم حذف المستخدم بنجاح!")
                            st.rerun()

        # قسم إضافة موجه جديد
        with col_add_u:
            st.markdown("### ➕ إضافة موجه / مستخدم جديد")
            with st.form("form_add_user"):
                new_u_fullname = st.text_input("الاسم الكامل")
                new_u_name = st.text_input("اسم المستخدم (للدخول)")
                new_u_pass = st.text_input("كلمة المرور", type="password")
                new_u_spec = st.text_input("التخصص")
                new_u_role = st.selectbox("الدور", ["Supervisor", "Admin"])
                
                if st.form_submit_button("➕ إنشاء الحساب", use_container_width=True):
                    if not new_u_name or not new_u_pass or not new_u_fullname:
                        st.error("يرجى ملء جميع الحقول المطلوبة.")
                    else:
                        try:
                            conn = sqlite3.connect("evaluation_system.db")
                            c = conn.cursor()
                            c.execute("INSERT INTO users (username, password, full_name, specialization, role) VALUES (?, ?, ?, ?, ?)",
                                      (new_u_name, new_u_pass, new_u_fullname, new_u_spec, new_u_role))
                            conn.commit()
                            conn.close()
                            st.success(f"✅ تم إنشاء حساب الموجه ({new_u_fullname}) بنجاح!")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("❌ اسم المستخدم هذا موجود بالفعل!")

    # ------------------ تبويب 2: إدارة وتعديل المدارس ------------------
    with tab2:
        st.subheader("🏫 قائمة المدارس والمؤسسات التعليمية")
        conn = sqlite3.connect("evaluation_system.db")
        schools_all = pd.read_sql_query("SELECT id, name, gender, location FROM schools", conn)
        conn.close()
        
        st.dataframe(schools_all, use_container_width=True)
        
        col_edit_s, col_add_s = st.columns(2)
        
        # قسم تعديل المدرسة
        with col_edit_s:
            st.markdown("### ✏️ تعديل بيانات مدرسة")
            selected_school_id = st.selectbox("اختر المدرسة للتعديل", schools_all["id"].tolist(), format_func=lambda x: f"ID: {x} - {schools_all[schools_all['id'] == x]['name'].values[0]}")
            
            if selected_school_id:
                conn = sqlite3.connect("evaluation_system.db")
                c = conn.cursor()
                c.execute("SELECT id, name, gender, location FROM schools WHERE id=?", (selected_school_id,))
                school_row = c.fetchone()
                conn.close()
                
                with st.form("form_edit_school"):
                    edit_s_name = st.text_input("اسم المدرسة / المؤسسة", value=school_row[1])
                    edit_s_gender = st.selectbox("نوع المدرسة", ["ذكور", "إناث"], index=0 if school_row[2] == "ذكور" else 1)
                    edit_s_location = st.text_input("الموقع الجغرافي / المنطقة", value=school_row[3])
                    
                    sc_btn1, sc_btn2 = st.columns(2)
                    with sc_btn1:
                        submit_edit_s = st.form_submit_button("💾 حفظ تعديل المدرسة", use_container_width=True)
                    with sc_btn2:
                        delete_s = st.form_submit_button("🗑️ حذف المدرسة", use_container_width=True)
                        
                    if submit_edit_s:
                        conn = sqlite3.connect("evaluation_system.db")
                        c = conn.cursor()
                        c.execute("UPDATE schools SET name=?, gender=?, location=? WHERE id=?",
                                  (edit_s_name, edit_s_gender, edit_s_location, selected_school_id))
                        conn.commit()
                        conn.close()
                        st.success("✅ تم تحديث بيانات المدرسة بنجاح!")
                        st.rerun()
                        
                    if delete_s:
                        conn = sqlite3.connect("evaluation_system.db")
                        c = conn.cursor()
                        c.execute("DELETE FROM schools WHERE id=?", (selected_school_id,))
                        conn.commit()
                        conn.close()
                        st.warning("🗑️ تم حذف المدرسة بنجاح!")
                        st.rerun()

        # قسم إضافة مدرسة جديدة
        with col_add_s:
            st.markdown("### ➕ إضافة مدرسة جديدة")
            with st.form("form_add_school"):
                new_s_name = st.text_input("اسم المدرسة الجديد")
                new_s_gender = st.selectbox("النوع", ["ذكور", "إناث"])
                new_s_location = st.text_input("الموقع / العنوان")
                
                if st.form_submit_button("➕ إضافة المدرسة للقائمة", use_container_width=True):
                    if not new_s_name:
                        st.error("يرجى إدخال اسم المدرسة.")
                    else:
                        conn = sqlite3.connect("evaluation_system.db")
                        c = conn.cursor()
                        c.execute("INSERT INTO schools (name, gender, location) VALUES (?, ?, ?)",
                                  (new_s_name, new_s_gender, new_s_location))
                        conn.commit()
                        conn.close()
                        st.success(f"✅ تمت إضافة مدرسة ({new_s_name}) بنجاح!")
                        st.rerun()